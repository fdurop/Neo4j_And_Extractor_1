import json
import os
import re
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from collections import defaultdict
import requests
import time
import functools
import traceback
import sys

# ==================== 导入 Flask Logger ====================
try:
    from models.logger import get_logger, RequestIdFilter

    # 使用 Flask 的 logger 系统
    logger = get_logger(
        name="knowledge_graph",
        log_file="logs/knowledge_graph.log"
    )
except ImportError:
    # 如果不在 Flask 环境，使用标准 logging
    import logging

    logger = logging.getLogger("knowledge_graph")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [N/A] - %(name)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)


# ==================== 日志装饰器 ====================

def log_execution_time(logger_instance=None):
    """记录函数执行时间的装饰器"""

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            _logger = logger_instance or logger
            start_time = time.time()
            _logger.info(f"开始执行: {func.__name__}")

            try:
                result = func(*args, **kwargs)
                elapsed_time = time.time() - start_time
                _logger.info(f"执行完成: {func.__name__} | 耗时: {elapsed_time:.2f}秒")
                return result
            except Exception as e:
                elapsed_time = time.time() - start_time
                _logger.error(f"执行失败: {func.__name__} | 耗时: {elapsed_time:.2f}秒 | 错误: {str(e)}", exc_info=True)
                raise

        return wrapper

    return decorator


def log_function_call(logger_instance=None):
    """记录函数调用参数和返回值的装饰器"""

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            _logger = logger_instance or logger
            args_repr = [repr(a) for a in args]
            kwargs_repr = [f"{k}={v!r}" for k, v in kwargs.items()]
            signature = ", ".join(args_repr + kwargs_repr)
            _logger.debug(f"调用函数: {func.__name__}({signature})")

            try:
                result = func(*args, **kwargs)
                _logger.debug(f"函数返回: {func.__name__} -> {result!r}")
                return result
            except Exception as e:
                _logger.error(f"函数异常: {func.__name__} | 错误: {str(e)}", exc_info=True)
                raise

        return wrapper

    return decorator


# ==================== 数据类 ====================

@dataclass
class ExtractedTriple:
    entities: List[Dict]
    relationships: List[Dict]
    attributes: List[Dict]


# ==================== DeepSeek API 客户端 ====================

class DeepSeekClient:
    """DeepSeek API客户端"""

    def __init__(self, api_key, base_url="https://api.deepseek.com/v1", model="deepseek-chat"):
        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        logger.info(f"DeepSeek客户端初始化: model={model}, base_url={base_url}")

    @log_execution_time(logger)
    def chat_completions_create(
            self,
            messages: List[Dict],
            temperature: float = 0.1,
            max_tokens: int = 8192,
            max_retries: int = 3
    ) -> Dict:
        """调用DeepSeek API - 带重试机制"""
        for attempt in range(max_retries):
            try:
                logger.info(f"API调用尝试 {attempt + 1}/{max_retries}")
                response = requests.post(
                    f"{self.base_url}/chat/completions",
                    headers=self.headers,
                    json={
                        "model": self.model,
                        "messages": messages,
                        "temperature": temperature,
                        "max_tokens": max_tokens
                    },
                    timeout=60,
                    proxies=None
                )
                response.raise_for_status()
                logger.info(f"API调用成功: status_code={response.status_code}")
                return response.json()

            except requests.exceptions.Timeout:
                wait_time = (attempt + 1) * 5
                logger.warning(f"API调用超时 (尝试 {attempt + 1}/{max_retries}), 等待 {wait_time}秒后重试")
                if attempt < max_retries - 1:
                    time.sleep(wait_time)
                continue

            except requests.exceptions.ConnectionError as e:
                wait_time = (attempt + 1) * 3
                logger.warning(f"API连接错误 (尝试 {attempt + 1}/{max_retries}): {str(e)}, 等待 {wait_time}秒后重试")
                if attempt < max_retries - 1:
                    time.sleep(wait_time)
                continue

            except requests.exceptions.HTTPError as e:
                logger.error(f"HTTP错误: status_code={e.response.status_code}, message={str(e)}")
                if e.response.status_code == 429:
                    logger.warning("触发API限流，等待30秒后重试")
                    time.sleep(30)
                    continue
                else:
                    break

            except Exception as e:
                logger.error(f"未知错误 (尝试 {attempt + 1}/{max_retries}): {str(e)}", exc_info=True)
                if attempt < max_retries - 1:
                    time.sleep(5)
                continue

        logger.error("API调用完全失败，返回空结果")
        return {"choices": [{"message": {"content": "{\"entities\": [], \"relationships\": []}"}}]}


# ==================== 实体抽取器 ====================

class EntityExtractor:
    """实体关系抽取器 - 强化知识点间关系"""

    def __init__(self, deepseek_api_key: str, course_name: str = "专业课程"):
        self.deepseek = DeepSeekClient(deepseek_api_key)
        self.course_name = course_name
        self.arduino_keywords = [
            'Arduino', 'LED', 'sensor', '传感器', 'pin', '引脚', 'GPIO',
            'voltage', '电压', 'current', '电流', 'resistor', '电阻', 'PWM',
            'digital', '数字', 'analog', '模拟', 'serial', '串口', 'I2C', 'SPI',
            'breadboard', '面包板', 'wire', '导线', 'ground', '接地', 'VCC', '5V', '3.3V'
        ]
        logger.info(f"实体抽取器初始化: course_name={course_name}")

    def _extract_page_number(self, filename: str) -> int:
        """从文件名中提取页面/幻灯片号码"""
        patterns = [
            r'slide_?(\d+)',
            r'p(\d+)',
            r'page_?(\d+)',
            r'_(\d+)',
            r'(\d+)'
        ]
        for pattern in patterns:
            match = re.search(pattern, filename, re.IGNORECASE)
            if match:
                page_num = int(match.group(1))
                logger.debug(f"提取页码: {filename} -> {page_num}")
                return page_num
        logger.debug(f"未找到页码: {filename}")
        return 0

    @log_execution_time(logger)
    def load_multimodal_data(self, output_dir: str = "output") -> Dict:
        """加载多模态预处理的输出数据"""
        logger.info(f"开始加载多模态数据: output_dir={output_dir}")
        result = {'slides': [], 'images': []}

        try:
            text_dir = os.path.join(output_dir, "text")
            image_dir = os.path.join(output_dir, "images")

            # 加载文本数据
            if os.path.exists(text_dir):
                text_files = os.listdir(text_dir)
                slide_files = [f for f in text_files if f.endswith('.json') and not f.endswith('_desc.json')]
                logger.info(f"找到文本文件: {len(slide_files)}个")

                for slide_file in slide_files:
                    slide_path = os.path.join(text_dir, slide_file)
                    try:
                        with open(slide_path, 'r', encoding='utf-8') as f:
                            slide_data = json.load(f)
                        slide_num = self._extract_page_number(slide_file)
                        result['slides'].append({
                            "slide_number": slide_num,
                            "content": slide_data,
                            "source_file": slide_file
                        })
                        logger.debug(f"加载文本文件: {slide_file} (页面 {slide_num})")
                    except Exception as e:
                        logger.warning(f"加载文本文件失败 {slide_file}: {str(e)}")

            # 加载图片数据
            if os.path.exists(image_dir):
                image_files_list = os.listdir(image_dir)
                image_files = [f for f in image_files_list if
                               f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                logger.info(f"找到图片文件: {len(image_files)}个")

                for image_file in image_files:
                    base_name = os.path.splitext(image_file)[0]
                    desc_file = base_name + '_desc.json'
                    desc_path = os.path.join(image_dir, desc_file)
                    slide_num = self._extract_page_number(image_file)

                    image_data = {
                        "image_path": os.path.join(image_dir, image_file),
                        "slide_number": slide_num,
                        "filename": image_file,
                        "descriptions": [],
                        "ocr_text": ""
                    }

                    if os.path.exists(desc_path):
                        try:
                            with open(desc_path, 'r', encoding='utf-8') as f:
                                desc_data = json.load(f)
                                image_data["descriptions"] = desc_data.get("clip_descriptions", [])
                            logger.debug(f"加载图片描述: {desc_file}")
                        except Exception as e:
                            logger.warning(f"加载图片描述失败 {desc_file}: {str(e)}")

                    result['images'].append(image_data)

            # 排序
            result['slides'].sort(key=lambda x: x['slide_number'])
            result['images'].sort(key=lambda x: x['slide_number'])

            logger.info(f"数据加载完成: 文本={len(result['slides'])}个, 图片={len(result['images'])}个")

        except Exception as e:
            logger.error(f"数据加载失败: {str(e)}", exc_info=True)
            raise

        return result

    @log_execution_time(logger)
    def extract_entities_from_multimodal(self, multimodal_data: Dict) -> ExtractedTriple:
        """从多模态数据中抽取实体关系 - 强化知识点间关系"""
        logger.info(f"开始实体抽取: 课程《{self.course_name}》")

        all_entities = []
        all_relationships = []
        all_attributes = []

        # 1. 批量处理文本内容
        slides = multimodal_data.get('slides', [])
        logger.info(f"处理文本文件: {len(slides)}个")

        batch_size = 10
        total_batches = (len(slides) + batch_size - 1) // batch_size

        for batch_idx in range(total_batches):
            start_idx = batch_idx * batch_size
            end_idx = min(start_idx + batch_size, len(slides))
            batch_slides = slides[start_idx:end_idx]

            logger.info(f"处理批次 {batch_idx + 1}/{total_batches}: 页面 {start_idx + 1}-{end_idx}")

            batch_entities, batch_relations = self._extract_from_slide_batch(batch_slides)
            all_entities.extend(batch_entities)
            all_relationships.extend(batch_relations)

            if batch_idx < total_batches - 1:
                import random
                delay = random.uniform(2, 3)
                logger.debug(f"批次间等待 {delay:.1f}秒")
                time.sleep(delay)

        # 2. 处理图片内容
        images = multimodal_data.get('images', [])
        logger.info(f"处理图片文件: {len(images)}个")

        for i, image_data in enumerate(images):
            logger.debug(f"处理图片 {i + 1}/{len(images)}: {image_data.get('filename', '')}")
            img_entities = self._extract_from_image(image_data)
            all_entities.extend(img_entities)

        # 3. 全局关系挖掘
        logger.info("开始全局关系挖掘")
        enhanced_relationships = self._discover_global_relationships(all_entities, all_relationships)
        all_relationships.extend(enhanced_relationships)

        # 去重处理
        all_entities = self._deduplicate_entities(all_entities)
        all_relationships = self._deduplicate_relationships(all_relationships)

        # 4. 过滤页面关系
        logger.info("开始过滤非知识关系")
        filtered_relationships = self._filter_knowledge_relationships(all_relationships, all_entities)

        logger.info(
            f"实体抽取完成: 实体={len(all_entities)}个, "
            f"原始关系={len(all_relationships)}个, 知识关系={len(filtered_relationships)}个"
        )

        return ExtractedTriple(
            entities=all_entities,
            relationships=filtered_relationships,
            attributes=all_attributes
        )

    def _extract_from_slide_batch(self, batch_slides: List[Dict]) -> Tuple[List[Dict], List[Dict]]:
        """批量处理 - 强调知识点间关系"""
        if not batch_slides:
            logger.warning("批次中无幻灯片数据")
            return [], []

        combined_text = ""
        for slide in batch_slides:
            slide_content = slide.get('content', {})
            text_content = ""

            if isinstance(slide_content, dict):
                text_fields = ['text', 'content', 'raw_text', 'texts']
                for field in text_fields:
                    if field in slide_content:
                        content = slide_content[field]
                        if isinstance(content, list):
                            text_content = ' '.join(str(item) for item in content)
                        else:
                            text_content = str(content)
                        break
                if not text_content:
                    text_content = str(slide_content)
            else:
                text_content = str(slide_content)

            if text_content and text_content.strip() != "" and text_content != "{}":
                combined_text += f"\n\n{text_content}"

        if not combined_text.strip():
            logger.warning("批次中无有效文本内容")
            return [], []

        max_length = 8000
        if len(combined_text) > max_length:
            combined_text = combined_text[:max_length] + "..."
            logger.info(f"批次文本过长，截取前{max_length}字符")

        prompt = f"""
你是一位经验丰富的讲授《{self.course_name}》课程的大学教师，正在为本科生讲授专业课程。
现在你拿到了本次课程的PPT内容，需要从中提取核心知识点和它们之间的关系，帮助学生构建完整的知识体系。

**你的任务**：
请**通读全文理解上下文后**，从教学角度，结合《{self.course_name}》领域的专业知识，抽取课程中的核心知识点（实体）和知识点之间的关系，
并且在抽取关系的结果之余检查一下：**依据《{self.course_name}》领域专业知识，这些关系对于这些实体是否充分，假如不充分，自己依据相关知识补充缺乏的关系。**
帮助学生理解知识的层次结构和内在联系。

**核心要求**：
1. 综合分析全文，识别跨章节的知识关联和层次结构
2. 同一概念多次出现时合并理解，不要重复创建
3. 只抽取课程知识点，忽略"页面"、"幻灯片"、"第X章"等文档结构
4. 关系必须有实际教学价值，能帮助学生理解知识体系

**课程PPT内容**：
{combined_text}

---

**实体类型**：核心概念、硬件组件、技术特性、参数规格、编程概念、应用场景

**关系类型**（优先抽取跨章节和深层关系）：
- 包含、依赖、应用、属性、分类、对比
- 演进（A发展为B）、因果（A导致B）、顺序（A是B的前提）
- 连接（仅限物理连接）

**禁止**：
- ❌ 不创建"页面X"、"第X章"、"幻灯片X"等文档结构实体
- ❌ 不创建"页面→知识点"、"提到"、"出现在"等无教学意义的关系

**注意**
和课程专业知识无关的内容：
比如"表格""公式"这种单薄的词汇，这些对于专业课的知识没有体现，只是笼统的说法。

**输出格式**（确保JSON格式正确）：
{{
    "entities": [
        {{"name": "实体名称", "type": "类型", "description": "结合课程上下文的描述"}}
    ],
    "relationships": [
        {{"source": "源知识点", "target": "目标知识点", "relation": "关系类型", "description": "关系依据"}}
    ]
}}

注意：所有字符串用双引号，最后元素后无逗号，特殊字符需转义。
"""

        try:
            response = self.deepseek.chat_completions_create([
                {"role": "user", "content": prompt}
            ], max_retries=3, max_tokens=8192)

            content = response['choices'][0]['message']['content']
            json_start = content.find('{')
            json_end = content.rfind('}') + 1

            if json_start != -1 and json_end != -1:
                json_str = content[json_start:json_end]
                result = json.loads(json_str)

                entities = result.get('entities', [])
                for entity in entities:
                    entity['source'] = 'batch_text'

                relationships = result.get('relationships', [])
                for rel in relationships:
                    rel['source'] = 'batch_text'

                logger.info(f"批次抽取成功: 实体={len(entities)}个, 关系={len(relationships)}个")
                return entities, relationships

        except Exception as e:
            logger.error(f"批次实体抽取失败: {str(e)}", exc_info=True)

        return [], []

    def _discover_global_relationships(self, entities: List[Dict], existing_relationships: List[Dict]) -> List[Dict]:
        """全局关系挖掘 - 基于所有实体发现隐含关系"""
        logger.info(f"分析 {len(entities)} 个实体，发现隐含关系")

        entity_names = [e['name'] for e in entities]
        entity_types = {e['name']: e.get('type', 'unknown') for e in entities}

        if len(entity_names) > 50:
            logger.warning(f"实体过多，只分析前50个核心实体")
            entity_names = entity_names[:50]

        entity_list = "\n".join([f"- {name} ({entity_types.get(name, 'unknown')})" for name in entity_names])

        prompt = f"""
你是一位经验丰富的讲授《{self.course_name}》课程的大学教师，正在为本科生讲授专业课程。
以下是实体列表，

实体列表：
{entity_list}

请分析这些实体之间可能存在的**语义关系**，特别是那些在原文中没有明确提及但逻辑上存在的关系。

关系类型：
- 包含关系：A包含B
- 依赖关系：A依赖B
- 应用关系：A用于B
- 分类关系：A属于B类
- 对比关系：A与B对比
- 前置关系：学习A需要先学B

**要求**：
1. 只返回有明确语义价值的关系
2. 关系必须合理且符合《{self.course_name}》领域常识
3. 优先发现核心概念之间的关系
4. 每个关系必须有清晰的说明

返回JSON格式：
（注意下面的"关系类型不要"什么关系"这种形式而是"应用于""分类于"这种形式"）
{{
    "relationships": [
        {{"source": "实体A", "target": "实体B", "relation": "关系类型", "description": "关系说明", "confidence": 0.8}}
    ]
}}
"""

        try:
            response = self.deepseek.chat_completions_create([
                {"role": "user", "content": prompt}
            ], max_retries=2, max_tokens=3000)

            content = response['choices'][0]['message']['content']
            json_start = content.find('{')
            json_end = content.rfind('}') + 1

            if json_start != -1 and json_end != -1:
                json_str = content[json_start:json_end]
                result = json.loads(json_str)

                new_relationships = result.get('relationships', [])
                for rel in new_relationships:
                    rel['source_type'] = 'global_discovery'
                    rel['is_inferred'] = True

                logger.info(f"发现隐含关系: {len(new_relationships)}个")
                return new_relationships

        except Exception as e:
            logger.error(f"全局关系挖掘失败: {str(e)}", exc_info=True)

        return []

    def _filter_knowledge_relationships(self, relationships: List[Dict], entities: List[Dict]) -> List[Dict]:
        """过滤非知识关系，保留知识点间的关系 - 智能宽松版"""
        logger.info(f"开始过滤关系: 原始关系={len(relationships)}个, 实体={len(entities)}个")

        entity_names = set()
        for e in entities:
            name = e['name'].lower().strip()
            name = name.replace('(', '').replace(')', '').replace('（', '').replace('）', '')
            name = name.replace(',', '').replace('，', '').replace('、', '').replace(' ', '')
            entity_names.add(name)

        blacklist_keywords = [
            'batch_text', 'batch', 'source', 'target', 'json', 'api', 'http',
            'output', 'input', 'data', 'file', 'path', 'url', 'id', 'index',
            '页面', 'page', 'slide', '幻灯片', 'ppt', '文档', 'document',
            '第', '章', 'chapter', 'section', '图', 'figure', 'img', 'image',
            'null', 'none', 'undefined', 'unknown', 'test', 'example',
            'source_type', 'global_discovery', 'is_inferred', 'confidence'
        ]

        filtered = []
        filtered_count = 0

        def smart_match(name: str) -> bool:
            if not name:
                return False
            name_lower = name.lower().strip()
            for keyword in blacklist_keywords:
                if keyword in name_lower:
                    return False
            name_clean = name_lower.replace('(', '').replace(')', '').replace('（', '').replace('）', '')
            name_clean = name_clean.replace(',', '').replace('，', '').replace('、', '').replace(' ', '')
            if name_clean in entity_names:
                return True
            for entity_name in entity_names:
                if len(entity_name) >= 4 and len(name_clean) >= 4:
                    if entity_name in name_clean or name_clean in entity_name:
                        is_blacklist_match = False
                        for keyword in blacklist_keywords:
                            if keyword in name_clean or keyword in entity_name:
                                is_blacklist_match = True
                                break
                        if not is_blacklist_match:
                            return True
            return False

        for rel in relationships:
            source = rel.get('source', '')
            target = rel.get('target', '')
            if smart_match(source) and smart_match(target):
                filtered.append(rel)
            else:
                filtered_count += 1
                logger.debug(f"过滤关系: {source} -> {target}")

        logger.info(f"过滤完成: 过滤={filtered_count}个, 保留={len(filtered)}个")
        return filtered

    def _extract_from_image(self, image_data: Dict) -> List[Dict]:
        """从图片数据中抽取实体"""
        entities = []
        image_path = image_data.get('image_path', '')
        filename = image_data.get('filename', '')

        logger.debug(f"从图片提取实体: {filename}")

        descriptions = image_data.get('descriptions', [])
        for desc_item in descriptions:
            desc_text = desc_item.get('description', '')
            confidence = desc_item.get('confidence', 0)
            if desc_text and confidence > 0.05:
                entities.append({
                    'name': desc_text,
                    'type': 'image_concept',
                    'description': f'从图片描述中识别: {desc_text}',
                    'confidence': confidence,
                    'source': 'image_description',
                    'image_path': image_path,
                    'filename': filename
                })

        ocr_text = image_data.get('ocr_text', '')
        if ocr_text:
            for keyword in self.arduino_keywords:
                if keyword.lower() in ocr_text.lower():
                    entities.append({
                        'name': keyword,
                        'type': 'hardware_component',
                        'description': f'从图片OCR中识别的{keyword}',
                        'source': 'image_ocr',
                        'image_path': image_path,
                        'filename': filename
                    })

        if 'arduino' in filename.lower():
            entities.append({
                'name': 'Arduino',
                'type': 'hardware_platform',
                'description': '从文件名识别的Arduino平台',
                'source': 'filename',
                'image_path': image_path,
                'filename': filename
            })


        logger.debug(f"图片实体提取完成: {filename}, 实体数={len(entities)}")
        return entities

    def _deduplicate_entities(self, entities: List[Dict]) -> List[Dict]:
        """实体去重"""
        original_count = len(entities)
        seen = set()
        unique_entities = []

        for entity in entities:
            key = (entity['name'].lower(), entity['type'])
            if key not in seen:
                seen.add(key)
                unique_entities.append(entity)

        logger.info(f"实体去重: {original_count} -> {len(unique_entities)}")
        return unique_entities

    def _deduplicate_relationships(self, relationships: List[Dict]) -> List[Dict]:
        """关系去重"""
        original_count = len(relationships)
        seen = set()
        unique_relationships = []

        for rel in relationships:
            key = (rel['source'].lower(), rel['target'].lower(), rel['relation'])
            if key not in seen:
                seen.add(key)
                unique_relationships.append(rel)

        logger.info(f"关系去重: {original_count} -> {len(unique_relationships)}")
        return unique_relationships


# ==================== 主函数 ====================
# ⚠️ 注意：以下代码必须在模块级别，不能缩进到类内部

@log_execution_time(logger)
def extract_entities_from_output(
    output_dir: str,
    deepseek_api_key: str,
    course_name: str = "专业课程"
) -> ExtractedTriple:
    """从多模态输出中抽取实体关系的主函数"""
    logger.info(f"开始实体关系抽取: output_dir={output_dir}, course_name={course_name}")

    extractor = EntityExtractor(deepseek_api_key, course_name)
    multimodal_data = extractor.load_multimodal_data(output_dir)
    extracted_data = extractor.extract_entities_from_multimodal(multimodal_data)

    return extracted_data


# ==================== Neo4j 知识图谱 ====================

from neo4j import GraphDatabase


class Neo4jKnowledgeGraph:
    """Neo4j知识图谱连接器 - 只存储知识点关系"""

    def __init__(self, uri: str, user: str, password: str):
        """初始化Neo4j连接"""
        self.driver = None
        try:
            logger.info(f"尝试连接Neo4j: uri={uri}, user={user}")

            config = {
                "keep_alive": True,
                "max_connection_lifetime": 3600,
                "max_connection_pool_size": 100
            }
            self.driver = GraphDatabase.driver(uri, auth=(user, password), **config)

            with self.driver.session() as session:
                session.run("RETURN 1")

            logger.info(f"Neo4j连接成功: {uri}")

        except Exception as e:
            logger.error(f"Neo4j连接失败: {str(e)}", exc_info=True)
            raise

    def close(self):
        """关闭连接"""
        if self.driver:
            self.driver.close()
            logger.info("Neo4j连接已关闭")

    def clear_database(self):
        """清空数据库"""
        logger.warning("开始清空Neo4j数据库")

        with self.driver.session() as session:
            session.run("MATCH (n) DETACH DELETE n")

        logger.info("数据库已清空")

    def create_entity_node(self, entity_name: str, entity_type: str, description: str = "",
                           metadata: Dict = None):
        """创建实体节点"""
        logger.debug(f"创建实体节点: name={entity_name}, type={entity_type}")

        with self.driver.session() as session:
            query = """
            MERGE (e:Entity {name: $entity_name})
            SET e.type = $entity_type
            SET e.description = $description
            SET e.updated_at = datetime()
            """

            params = {
                "entity_name": entity_name,
                "entity_type": entity_type,
                "description": description
            }

            if metadata:
                for key, value in metadata.items():
                    if isinstance(value, (str, int, float, bool)):
                        query += f"SET e.{key} = ${key} "
                        params[key] = value

            session.run(query, params)

    def create_relationship(self, source_name: str, target_name: str, relation_type: str,
                            properties: Dict = None):
        """创建关系"""
        logger.debug(f"创建关系: {source_name} -[{relation_type}]-> {target_name}")

        with self.driver.session() as session:
            session.run("""
                MERGE (s:Entity {name: $source_name})
                MERGE (t:Entity {name: $target_name})
            """, source_name=source_name, target_name=target_name)

            query = f"""
            MATCH (s:Entity {{name: $source_name}})
            MATCH (t:Entity {{name: $target_name}})
            MERGE (s)-[r:{relation_type}]->(t)
            SET r.created_at = datetime()
            """

            params = {"source_name": source_name, "target_name": target_name}

            if properties:
                for key, value in properties.items():
                    if isinstance(value, (str, int, float, bool)):
                        query += f"SET r.{key} = ${key} "
                        params[key] = value

            session.run(query, params)

    @log_execution_time(logger)
    def save_extracted_data(self, extracted_data, ppt_name: str):
        """只保存知识点实体和关系，不创建文档节点"""
        logger.info(f"开始保存知识图谱到Neo4j: ppt_name={ppt_name}")

        try:
            # 1. 批量创建实体节点
            logger.info(f"创建实体节点: {len(extracted_data.entities)}个")
            entity_count = 0

            for entity in extracted_data.entities:
                try:
                    metadata = {k: v for k, v in entity.items()
                                if k not in ['name', 'type', 'description'] and
                                isinstance(v, (str, int, float, bool))}

                    self.create_entity_node(
                        entity['name'],
                        entity.get('type', 'unknown'),
                        entity.get('description', ''),
                        metadata
                    )
                    entity_count += 1

                except Exception as e:
                    logger.warning(f"创建实体失败 {entity['name']}: {str(e)}")
                    continue

            # 2. 批量创建知识点间的关系
            logger.info(f"创建知识关系: {len(extracted_data.relationships)}个")
            relation_count = 0

            for rel in extracted_data.relationships:
                try:
                    properties = {k: v for k, v in rel.items()
                                  if k not in ['source', 'target', 'relation'] and
                                  isinstance(v, (str, int, float, bool))}

                    self.create_relationship(
                        rel['source'],
                        rel['target'],
                        rel['relation'].upper().replace(' ', '_'),
                        properties
                    )
                    relation_count += 1

                except Exception as e:
                    logger.warning(f"创建关系失败 {rel['source']}->{rel['target']}: {str(e)}")
                    continue

            logger.info(
                f"知识图谱保存完成: "
                f"实体={entity_count}/{len(extracted_data.entities)}, "
                f"关系={relation_count}/{len(extracted_data.relationships)}"
            )

            return {
                'document': ppt_name,
                'entities_saved': entity_count,
                'relationships_saved': relation_count,
                'success': True
            }

        except Exception as e:
            logger.error(f"保存数据失败: {str(e)}", exc_info=True)
            return {
                'document': ppt_name,
                'entities_saved': 0,
                'relationships_saved': 0,
                'success': False,
                'error': str(e)
            }

    def query_entities(self, limit: int = 10) -> List[Dict]:
        """查询实体"""
        logger.debug(f"查询实体: limit={limit}")

        with self.driver.session() as session:
            result = session.run("""
                MATCH (e:Entity)
                RETURN e.name as name, e.type as type, e.description as description
                LIMIT $limit
            """, limit=limit)

            return [dict(record) for record in result]

    def query_relationships(self, limit: int = 10) -> List[Dict]:
        """查询关系"""
        logger.debug(f"查询关系: limit={limit}")

        with self.driver.session() as session:
            result = session.run("""
                MATCH (s:Entity)-[r]->(t:Entity)
                RETURN s.name as source, type(r) as relation, t.name as target
                LIMIT $limit
            """, limit=limit)

            return [dict(record) for record in result]

    def get_statistics(self) -> Dict:
        """获取数据库统计信息"""
        logger.debug("获取数据库统计信息")

        with self.driver.session() as session:
            # 节点统计
            node_result = session.run("MATCH (n:Entity) RETURN count(n) as total_nodes")
            total_nodes = node_result.single()['total_nodes']

            # 关系统计
            rel_result = session.run("MATCH ()-[r]->() RETURN count(r) as total_relationships")
            total_relationships = rel_result.single()['total_relationships']

            # 实体类型统计
            entity_type_result = session.run("""
                MATCH (e:Entity)
                RETURN e.type as entity_type, count(e) as count
                ORDER BY count DESC
            """)
            entity_types = [dict(record) for record in entity_type_result]

            # 关系类型统计
            rel_type_result = session.run("""
                MATCH ()-[r]->()
                RETURN type(r) as relation_type, count(r) as count
                ORDER BY count DESC
            """)
            relation_types = [dict(record) for record in rel_type_result]

            stats = {
                'total_nodes': total_nodes,
                'total_relationships': total_relationships,
                'entity_types': entity_types,
                'relation_types': relation_types
            }

            logger.info(f"统计信息: 节点={total_nodes}, 关系={total_relationships}")
            return stats

    def search_entities_by_name(self, name_pattern: str, limit: int = 10) -> List[Dict]:
        """按名称搜索实体"""
        logger.debug(f"搜索实体: pattern={name_pattern}, limit={limit}")

        with self.driver.session() as session:
            result = session.run("""
                MATCH (e:Entity)
                WHERE toLower(e.name) CONTAINS toLower($pattern)
                RETURN e.name as name, e.type as type, e.description as description
                LIMIT $limit
            """, pattern=name_pattern, limit=limit)

            return [dict(record) for record in result]

    def get_entity_neighbors(self, entity_name: str, depth: int = 1) -> Dict:
        """获取实体的邻居节点"""
        logger.debug(f"获取实体邻居: entity={entity_name}, depth={depth}")

        with self.driver.session() as session:
            result = session.run(f"""
                MATCH path = (e:Entity {{name: $entity_name}})-[*1..{depth}]-(neighbor)
                RETURN neighbor.name as name, neighbor.type as type, 
                       neighbor.description as description
                LIMIT 20
            """, entity_name=entity_name)

            neighbors = [dict(record) for record in result]

            rel_result = session.run("""
                MATCH (e:Entity {name: $entity_name})-[r]-(neighbor)
                RETURN neighbor.name as neighbor, type(r) as relation,
                       CASE WHEN startNode(r).name = $entity_name 
                            THEN 'outgoing' ELSE 'incoming' END as direction
            """, entity_name=entity_name)

            relationships = [dict(record) for record in rel_result]

            return {
                'entity': entity_name,
                'neighbors': neighbors,
                'relationships': relationships
            }


@log_execution_time(logger)
def save_to_neo4j(extracted_data, ppt_name: str, neo4j_uri: str, neo4j_user: str, neo4j_password: str):
    """保存数据到Neo4j的主函数"""
    logger.info(f"开始保存到Neo4j: ppt_name={ppt_name}")

    kg = Neo4jKnowledgeGraph(neo4j_uri, neo4j_user, neo4j_password)

    try:
        result = kg.save_extracted_data(extracted_data, ppt_name)

        # 显示统计信息
        stats = kg.get_statistics()

        logger.info(
            f"知识图谱统计: "
            f"实体={stats['total_nodes']}, "
            f"关系={stats['total_relationships']}"
        )

        return result

    finally:
        kg.close()


@log_execution_time(logger)
def output_to_neo4j(
    output_dir: str,
    deepseek_api_key: str,
    neo4j_uri: str,
    neo4j_user: str,
    neo4j_password: str,
    course_name: str = "专业课程",
    ppt_name: str = "课程PPT",
    clear_database: bool = False,
    show_examples: bool = True
) -> Dict:
    """
    从output目录提取实体关系并保存到Neo4j的主函数（强化知识点关系版）

    Args:
        output_dir: 输出目录路径
        deepseek_api_key: DeepSeek API密钥
        neo4j_uri: Neo4j数据库URI
        neo4j_user: Neo4j用户名
        neo4j_password: Neo4j密码
        course_name: 课程名称（如"直线运动平台"、"Arduino基础"等）
        ppt_name: PPT文档名称
        clear_database: 是否清空数据库
        show_examples: 是否显示查询示例

    Returns:
        Dict: 包含处理结果的字典
    """

    logger.info("=" * 60)
    logger.info("开始知识图谱构建任务")
    logger.info(f"课程: {course_name}")
    logger.info(f"文档: {ppt_name}")
    logger.info(f"输出目录: {output_dir}")
    logger.info(f"Neo4j URI: {neo4j_uri}")
    logger.info("=" * 60)

    result = {
        'success': False,
        'entities_extracted': 0,
        'relationships_extracted': 0,
        'entities_saved': 0,
        'relationships_saved': 0,
        'error': None
    }

    kg = None

    try:
        # 1. 实体关系抽取
        logger.info("步骤1: 实体关系抽取")
        extracted_data = extract_entities_from_output(output_dir, deepseek_api_key, course_name)

        result['entities_extracted'] = len(extracted_data.entities)
        result['relationships_extracted'] = len(extracted_data.relationships)

        logger.info(
            f"抽取完成: "
            f"实体={result['entities_extracted']}, "
            f"关系={result['relationships_extracted']}"
        )

        if result['entities_extracted'] == 0 and result['relationships_extracted'] == 0:
            logger.warning("未抽取到任何实体和关系")
            return result

        # 2. 连接Neo4j
        logger.info("步骤2: 连接Neo4j数据库")
        kg = Neo4jKnowledgeGraph(neo4j_uri, neo4j_user, neo4j_password)

        # 3. 可选：清空数据库
        if clear_database:
            logger.warning("步骤3: 清空数据库")
            kg.clear_database()

        # 4. 保存数据到Neo4j
        logger.info("步骤4: 保存知识图谱到Neo4j")
        save_result = kg.save_extracted_data(extracted_data, ppt_name)

        result['entities_saved'] = save_result['entities_saved']
        result['relationships_saved'] = save_result['relationships_saved']
        result['success'] = save_result['success']

        if not save_result['success']:
            result['error'] = save_result.get('error', '保存失败')
            return result

        # 5. 显示统计信息
        logger.info("步骤5: 获取知识图谱统计信息")
        stats = kg.get_statistics()

        # 6. 可选：显示查询示例
        if show_examples:
            logger.info("步骤6: 显示知识图谱查询示例")

            # 查询实体示例
            entities = kg.query_entities(5)
            if entities:
                logger.info(f"实体示例 (前5个):")
                for i, entity in enumerate(entities):
                    logger.info(f"  {i + 1}. {entity['name']} ({entity['type']})")

            # 查询关系示例
            relationships = kg.query_relationships(10)
            if relationships:
                logger.info(f"关系示例 (前10个):")
                for i, rel in enumerate(relationships):
                    logger.info(f"  {i + 1}. {rel['source']} -[{rel['relation']}]-> {rel['target']}")

        # 7. 成功完成
        logger.info("=" * 60)
        logger.info("知识图谱构建完成!")
        logger.info(f"实体抽取: {result['entities_extracted']}个")
        logger.info(f"关系抽取: {result['relationships_extracted']}个")
        logger.info(f"实体保存: {result['entities_saved']}/{result['entities_extracted']}")
        logger.info(f"关系保存: {result['relationships_saved']}/{result['relationships_extracted']}")
        logger.info("=" * 60)

        result['success'] = True

    except Exception as e:
        logger.error(f"知识图谱构建失败: {str(e)}", exc_info=True)
        result['error'] = str(e)
        result['success'] = False

    finally:
        if kg:
            kg.close()

    return result


# ==================== 主程序入口 ====================

if __name__ == "__main__":
    logger.info("程序启动")

    result = output_to_neo4j(
        output_dir=r"C:\Users\Lin\PycharmProjects\PythonProject\output",
        deepseek_api_key="sk-c28ec338b39e4552b9e6bded47466442",
        neo4j_uri="bolt://101.132.130.25:7687",
        neo4j_user="neo4j",
        neo4j_password="wangshuxvan@1",
        course_name="直线运动平台",
        ppt_name="直线运动平台知识图谱",
        clear_database=False,
        show_examples=True
    )

    if result['success']:
        logger.info("✅ 知识图谱构建成功!")
        logger.info(f"实体: {result['entities_saved']}/{result['entities_extracted']}")
        logger.info(f"关系: {result['relationships_saved']}/{result['relationships_extracted']}")
    else:
        logger.error(f"❌ 构建失败: {result.get('error', '未知错误')}")

    logger.info("程序结束")

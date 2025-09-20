import json
import os
import re
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from collections import defaultdict
import requests
import time


@dataclass
class ExtractedTriple:
    entities: List[Dict]
    relationships: List[Dict]
    attributes: List[Dict]


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

    def chat_completions_create(self, messages: List[Dict], temperature: float = 0.1, max_tokens: int = 1024) -> Dict:
        """调用DeepSeek API"""
        try:
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=self.headers,
                json={
                    "model": self.model,
                    "messages": messages,
                    "temperature": temperature,
                    "max_tokens": max_tokens
                },
                timeout=30
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"DeepSeek API调用失败: {e}")
            return {"choices": [{"message": {"content": "{\"entities\": [], \"relationships\": []}"}}]}


class EntityExtractor:
    """实体关系抽取器"""

    def __init__(self, deepseek_api_key: str):
        self.deepseek = DeepSeekClient(deepseek_api_key)
        self.arduino_keywords = [
            'Arduino', 'LED', 'sensor', '传感器', 'pin', '引脚', 'GPIO',
            'voltage', '电压', 'current', '电流', 'resistor', '电阻', 'PWM',
            'digital', '数字', 'analog', '模拟', 'serial', '串口', 'I2C', 'SPI',
            'breadboard', '面包板', 'wire', '导线', 'ground', '接地', 'VCC', '5V', '3.3V'
        ]

    def _extract_slide_number(self, filename: str) -> int:
        """从文件名中提取幻灯片号码"""
        match = re.search(r'slide_(\d+)', filename)
        return int(match.group(1)) if match else 0

    def load_multimodal_data(self, output_dir: str = "output") -> Dict:
        """加载多模态预处理的输出数据 - 适配实际文件格式"""
        result = {
            'slides': [],
            'images': []
        }

        try:
            # 定义子目录路径
            text_dir = os.path.join(output_dir, "text")
            image_dir = os.path.join(output_dir, "images")

            # 1. 加载幻灯片文本数据 (从text目录)
            if os.path.exists(text_dir):
                text_files = os.listdir(text_dir)
                slide_files = [f for f in text_files if
                               '_slide_' in f and f.endswith('.json') and not f.endswith('_desc.json')]

                for slide_file in slide_files:
                    slide_path = os.path.join(text_dir, slide_file)  # 注意这里改为text_dir
                    try:
                        with open(slide_path, 'r', encoding='utf-8') as f:
                            slide_data = json.load(f)

                        slide_num = self._extract_slide_number(slide_file)

                        result['slides'].append({
                            "slide_number": slide_num,
                            "content": slide_data,
                            "source_file": slide_file
                        })

                    except Exception as e:
                        print(f"⚠️ 加载幻灯片文件失败 {slide_file}: {e}")

            # 2. 加载图片数据 (从image目录)
            if os.path.exists(image_dir):
                image_files_list = os.listdir(image_dir)
                image_files = [f for f in image_files_list if f.endswith('.png') or f.endswith('.jpg')]

                for image_file in image_files:
                    # 查找对应的描述文件
                    desc_file = image_file.replace('.png', '_desc.json').replace('.jpg', '_desc.json')
                    desc_path = os.path.join(image_dir, desc_file)  # 注意这里改为image_dir

                    slide_num = self._extract_slide_number(image_file)

                    image_data = {
                        "image_path": os.path.join(image_dir, image_file),  # 注意这里改为image_dir
                        "slide_number": slide_num,
                        "filename": image_file,
                        "descriptions": [],
                        "ocr_text": ""
                    }

                    # 如果有描述文件，加载描述信息
                    if os.path.exists(desc_path):
                        try:
                            with open(desc_path, 'r', encoding='utf-8') as f:
                                desc_data = json.load(f)
                                image_data["descriptions"] = desc_data.get("clip_descriptions", [])
                        except Exception as e:
                            print(f"⚠️ 加载图片描述失败 {desc_file}: {e}")

                    result['images'].append(image_data)

            print(f"✅ 数据加载完成:")
            print(f"   - 幻灯片: {len(result['slides'])}个文件")
            print(f"   - 图片: {len(result['images'])}个")

        except Exception as e:
            print(f"❌ 数据加载失败: {e}")

        return result

    def extract_entities_from_multimodal(self, multimodal_data: Dict) -> ExtractedTriple:
        """从多模态数据中抽取实体关系"""
        all_entities = []
        all_relationships = []
        all_attributes = []

        print("🔍 开始实体关系抽取...")

        # 1. 处理幻灯片文本内容
        slides = multimodal_data.get('slides', [])
        for i, slide in enumerate(slides):
            print(f"   处理幻灯片 {i + 1}/{len(slides)}: {slide.get('source_file', '')}")
            slide_entities, slide_relations = self._extract_from_slide_text(slide)
            all_entities.extend(slide_entities)
            all_relationships.extend(slide_relations)
            time.sleep(0.5)  # 避免API调用过快

        # 2. 处理图片内容
        images = multimodal_data.get('images', [])
        for i, image_data in enumerate(images):
            print(f"   处理图片 {i + 1}/{len(images)}: {image_data.get('filename', '')}")
            img_entities = self._extract_from_image(image_data)
            all_entities.extend(img_entities)

        # 去重处理
        all_entities = self._deduplicate_entities(all_entities)
        all_relationships = self._deduplicate_relationships(all_relationships)

        print(f"✅ 实体关系抽取完成: {len(all_entities)}个实体, {len(all_relationships)}个关系")

        return ExtractedTriple(
            entities=all_entities,
            relationships=all_relationships,
            attributes=all_attributes
        )

    def _extract_from_slide_text(self, slide: Dict) -> Tuple[List[Dict], List[Dict]]:
        """从幻灯片文本中抽取实体和关系"""
        slide_content = slide.get('content', {})
        slide_num = slide.get('slide_number', 0)

        # 提取文本内容
        text_content = ""
        if isinstance(slide_content, dict):
            # 如果content是字典，尝试提取文本字段
            text_content = slide_content.get('text', '') or slide_content.get('content', '') or str(slide_content)
        else:
            text_content = str(slide_content)

        if not text_content or text_content.strip() == "":
            return [], []

        # 构建提示词
        prompt = f"""
请从以下Arduino/电子工程课程幻灯片内容中抽取实体和关系。

内容：{text_content}

请识别以下类型的实体：
1. 硬件组件：Arduino板、传感器、LED、电阻、电容等
2. 技术概念：PWM、串口通信、数字信号、模拟信号等
3. 参数数值：电压值、电阻值、引脚号、频率等
4. 操作步骤：连接、编程、测试、调试等
5. 代码概念：函数、变量、库文件等

请识别实体间的关系：
- 组成关系：A包含B、A由B组成
- 连接关系：A连接到B、A接入B
- 控制关系：A控制B、A驱动B
- 参数关系：A的参数是B、A设置为B
- 功能关系：A用于B、A实现B

严格按照以下JSON格式返回，不要添加任何其他内容：
{{
    "entities": [
        {{"name": "实体名称", "type": "实体类型", "description": "实体描述"}}
    ],
    "relationships": [
        {{"source": "源实体", "target": "目标实体", "relation": "关系类型"}}
    ]
}}
"""

        try:
            response = self.deepseek.chat_completions_create([
                {"role": "user", "content": prompt}
            ])

            content = response['choices'][0]['message']['content']

            # 提取JSON部分
            json_start = content.find('{')
            json_end = content.rfind('}') + 1

            if json_start != -1 and json_end != -1:
                json_str = content[json_start:json_end]
                result = json.loads(json_str)

                # 添加slide信息
                entities = result.get('entities', [])
                for entity in entities:
                    entity['slide'] = slide_num
                    entity['source'] = 'slide_text'

                relationships = result.get('relationships', [])
                for rel in relationships:
                    rel['slide'] = slide_num
                    rel['source'] = 'slide_text'

                return entities, relationships

        except Exception as e:
            print(f"   ⚠️ 幻灯片 {slide_num} 实体抽取失败: {e}")

        return [], []

    def _extract_from_image(self, image_data: Dict) -> List[Dict]:
        """从图片数据中抽取实体"""
        entities = []
        slide_num = image_data.get('slide_number', 0)
        image_path = image_data.get('image_path', '')
        filename = image_data.get('filename', '')

        # 1. 基于图片描述抽取实体
        descriptions = image_data.get('descriptions', [])
        for desc_item in descriptions:
            desc_text = desc_item.get('description', '')
            confidence = desc_item.get('confidence', 0)

            if desc_text and confidence > 0.05:  # 置信度阈值
                entities.append({
                    'name': desc_text,
                    'type': 'image_concept',
                    'description': f'从图片描述中识别: {desc_text}',
                    'confidence': confidence,
                    'source': 'image_description',
                    'slide': slide_num,
                    'image_path': image_path,
                    'filename': filename
                })

        # 2. 基于OCR文本抽取实体（如果有OCR文本）
        ocr_text = image_data.get('ocr_text', '')
        if ocr_text:
            # Arduino关键词匹配
            for keyword in self.arduino_keywords:
                if keyword.lower() in ocr_text.lower():
                    entities.append({
                        'name': keyword,
                        'type': 'hardware_component',
                        'description': f'从图片OCR中识别的{keyword}',
                        'source': 'image_ocr',
                        'slide': slide_num,
                        'image_path': image_path,
                        'filename': filename
                    })

        # 3. 基于文件名抽取实体（如果文件名包含有用信息）
        if 'arduino' in filename.lower():
            entities.append({
                'name': 'Arduino',
                'type': 'hardware_platform',
                'description': '从文件名识别的Arduino平台',
                'source': 'filename',
                'slide': slide_num,
                'image_path': image_path,
                'filename': filename
            })

        return entities

    def _deduplicate_entities(self, entities: List[Dict]) -> List[Dict]:
        """实体去重"""
        seen = set()
        unique_entities = []

        for entity in entities:
            key = (entity['name'].lower(), entity['type'])
            if key not in seen:
                seen.add(key)
                unique_entities.append(entity)

        return unique_entities

    def _deduplicate_relationships(self, relationships: List[Dict]) -> List[Dict]:
        """关系去重"""
        seen = set()
        unique_relationships = []

        for rel in relationships:
            key = (rel['source'].lower(), rel['target'].lower(), rel['relation'])
            if key not in seen:
                seen.add(key)
                unique_relationships.append(rel)

        return unique_relationships


def extract_entities_from_output(output_dir: str, deepseek_api_key: str) -> ExtractedTriple:
    """从多模态输出中抽取实体关系的主函数"""
    extractor = EntityExtractor(deepseek_api_key)

    # 加载数据
    multimodal_data = extractor.load_multimodal_data(output_dir)

    # 抽取实体关系
    extracted_data = extractor.extract_entities_from_multimodal(multimodal_data)

    return extracted_data


if __name__ == "__main__":
    # 测试代码
    DEEPSEEK_API_KEY = "sk-c28ec338b39e4552b9e6bded47466442"  # 替换为你的API key

    try:
        result = extract_entities_from_output(r"C:\Users\Lin\PycharmProjects\PythonProject\output", DEEPSEEK_API_KEY)
        print(f"\n📊 抽取结果统计:")
        print(f"   🏷️  实体数量: {len(result.entities)}")
        print(f"   🔗 关系数量: {len(result.relationships)}")
        print(f"   📋 属性数量: {len(result.attributes)}")

        # 显示前几个实体
        print(f"\n🔍 实体示例:")
        for i, entity in enumerate(result.entities[:5]):
            print(f"   {i + 1}. {entity['name']} ({entity['type']}) - {entity.get('description', '')}")

        # 显示前几个关系
        if result.relationships:
            print(f"\n🔗 关系示例:")
            for i, rel in enumerate(result.relationships[:3]):
                print(f"   {i + 1}. {rel['source']} --{rel['relation']}--> {rel['target']}")

    except Exception as e:
        print(f"❌ 测试失败: {e}")

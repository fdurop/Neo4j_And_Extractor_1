from neo4j import GraphDatabase
import pandas as pd
from typing import Optional, Dict, List, Tuple
import json
from collections import defaultdict

# ==================== 日志系统集成 ====================
import logging
import uuid

# 检测是否在Flask环境中
try:
    from flask import g, has_request_context

    FLASK_AVAILABLE = True
except ImportError:
    FLASK_AVAILABLE = False
    has_request_context = lambda: False


# 日志过滤器：为每条日志添加 request_id
class RequestIdFilter(logging.Filter):
    def filter(self, record):
        if FLASK_AVAILABLE and has_request_context() and hasattr(g, 'request_id'):
            record.request_id = g.request_id
        elif hasattr(self, '_standalone_request_id'):
            record.request_id = self._standalone_request_id
        else:
            record.request_id = "N/A"
        return True


# 配置日志
def setup_logger(name: str = __name__, log_file: str = "app.log") -> logging.Logger:
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [%(request_id)s] - %(name)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # 文件处理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        request_filter = RequestIdFilter()
        file_handler.addFilter(request_filter)
        logger.addHandler(file_handler)

        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.addFilter(request_filter)
        logger.addHandler(console_handler)

        # 保存 filter 引用，用于独立运行时设置 request_id
        logger._request_filter = request_filter

    return logger


# 全局 logger
logger = setup_logger()



class Neo4jKnowledgeGraph:
    """Neo4j知识图谱连接器"""

    def __init__(self, uri: str, user: str, password: str):
        """初始化Neo4j连接"""
        self.driver = None
        try:
            config = {
                "keep_alive": True,
                "max_connection_lifetime": 3600,
                "max_connection_pool_size": 100
            }
            self.driver = GraphDatabase.driver(uri, auth=(user, password), **config)

            # 测试连接
            with self.driver.session() as session:
                session.run("RETURN 1")
            logger.info("✅ Neo4j连接成功")

        except Exception as e:
            logger.error(f"❌ Neo4j连接失败: {e}")
            raise

    def close(self):
        """关闭连接"""
        if self.driver:
            self.driver.close()

    def clear_database(self):
        """清空数据库（谨慎使用）"""
        with self.driver.session() as session:
            session.run("MATCH (n) DETACH DELETE n")
        logger.info("🗑️ 数据库已清空")

    def create_document_node(self, doc_name: str, doc_type: str = "ppt", metadata: Dict = None):
        """创建文档节点"""
        with self.driver.session() as session:
            query = """
            MERGE (d:Document {name: $doc_name})
            SET d.type = $doc_type
            SET d.created_at = datetime()
            """

            if metadata:
                for key, value in metadata.items():
                    query += f"SET d.{key} = ${key} "

            params = {"doc_name": doc_name, "doc_type": doc_type}
            if metadata:
                params.update(metadata)

            session.run(query, params)
        logger.info(f"📄 创建文档节点: {doc_name}")

    def create_entity_node(self, entity_name: str, entity_type: str, description: str = "",
                           metadata: Dict = None):
        """创建实体节点"""
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
        with self.driver.session() as session:
            # 确保源节点和目标节点存在
            session.run("""
                MERGE (s:Entity {name: $source_name})
                MERGE (t:Entity {name: $target_name})
            """, source_name=source_name, target_name=target_name)

            # 创建关系
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

    def save_extracted_data(self, extracted_data, ppt_name: str):
        """保存抽取的实体关系数据到Neo4j"""
        try:
            logger.info(f"💾 开始保存数据到Neo4j: {ppt_name}")

            # 1. 创建PPT文档节点
            self.create_document_node(ppt_name, "ppt", {
                "total_entities": len(extracted_data.entities),
                "total_relationships": len(extracted_data.relationships)
            })

            # 2. 批量创建实体节点
            logger.info(f"   创建 {len(extracted_data.entities)} 个实体节点...")
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

                    # 创建文档到实体的包含关系
                    self.create_relationship(ppt_name, entity['name'], "CONTAINS")
                    entity_count += 1

                except Exception as e:
                    logger.warning(f"     ⚠️ 创建实体失败 {entity['name']}: {e}")
                    continue

            # 3. 批量创建关系
            logger.info(f"   创建 {len(extracted_data.relationships)} 个关系...")
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
                    logger.warning(f"     ⚠️ 创建关系失败 {rel['source']}->{rel['target']}: {e}")
                    continue

            logger.info(f"✅ 数据保存完成:")
            logger.info(f"   📄 文档: {ppt_name}")
            logger.info(f"   🏷️  实体: {entity_count}/{len(extracted_data.entities)}")
            logger.info(f"   🔗 关系: {relation_count}/{len(extracted_data.relationships)}")

            return {
                'document': ppt_name,
                'entities_saved': entity_count,
                'relationships_saved': relation_count,
                'success': True
            }

        except Exception as e:
            logger.error(f"❌ 保存数据失败: {e}")
            return {
                'document': ppt_name,
                'entities_saved': 0,
                'relationships_saved': 0,
                'success': False,
                'error': str(e)
            }

    def query_entities(self, limit: int = 10) -> List[Dict]:
        """查询实体"""
        with self.driver.session() as session:
            result = session.run("""
                MATCH (e:Entity)
                RETURN e.name as name, e.type as type, e.description as description
                LIMIT $limit
            """, limit=limit)

            return [dict(record) for record in result]

    def query_relationships(self, limit: int = 10) -> List[Dict]:
        """查询关系"""
        with self.driver.session() as session:
            result = session.run("""
                MATCH (s:Entity)-[r]->(t:Entity)
                RETURN s.name as source, type(r) as relation, t.name as target
                LIMIT $limit
            """, limit=limit)

            return [dict(record) for record in result]

    def get_statistics(self) -> Dict:
        """获取数据库统计信息"""
        with self.driver.session() as session:
            # 节点统计
            node_result = session.run("MATCH (n) RETURN count(n) as total_nodes")
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

            # 文档统计
            doc_result = session.run("MATCH (d:Document) RETURN count(d) as total_documents")
            total_documents = doc_result.single()['total_documents']

            return {
                'total_nodes': total_nodes,
                'total_relationships': total_relationships,
                'total_documents': total_documents,
                'entity_types': entity_types
            }

    def search_entities_by_name(self, name_pattern: str, limit: int = 10) -> List[Dict]:
        """按名称搜索实体"""
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
        with self.driver.session() as session:
            result = session.run(f"""
                MATCH path = (e:Entity {{name: $entity_name}})-[*1..{depth}]-(neighbor)
                RETURN neighbor.name as name, neighbor.type as type, 
                       neighbor.description as description
                LIMIT 20
            """, entity_name=entity_name)

            neighbors = [dict(record) for record in result]

            # 获取相关关系
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


def save_to_neo4j(extracted_data, ppt_name: str, neo4j_uri: str, neo4j_user: str, neo4j_password: str):
    """保存数据到Neo4j的主函数"""
    kg = Neo4jKnowledgeGraph(neo4j_uri, neo4j_user, neo4j_password)

    try:
        result = kg.save_extracted_data(extracted_data, ppt_name)

        # 显示统计信息
        stats = kg.get_statistics()
        logger.info(f"\n📊 数据库统计信息:")
        logger.info(f"   📄 文档总数: {stats['total_documents']}")
        logger.info(f"   🏷️  节点总数: {stats['total_nodes']}")
        logger.info(f"   🔗 关系总数: {stats['total_relationships']}")

        # 显示实体类型分布
        logger.info(f"\n🏷️  实体类型分布:")
        for entity_type in stats['entity_types'][:5]:
            logger.info(f"   - {entity_type['entity_type']}: {entity_type['count']}个")

        return result

    finally:
        kg.close()


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
            logger.error(f"DeepSeek API调用失败: {e}")
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
                    slide_path = os.path.join(text_dir, slide_file)
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
                        logger.warning(f"⚠️ 加载幻灯片文件失败 {slide_file}: {e}")

            # 2. 加载图片数据 (从image目录)
            if os.path.exists(image_dir):
                image_files_list = os.listdir(image_dir)
                image_files = [f for f in image_files_list if f.endswith('.png') or f.endswith('.jpg')]

                for image_file in image_files:
                    # 查找对应的描述文件
                    desc_file = image_file.replace('.png', '_desc.json').replace('.jpg', '_desc.json')
                    desc_path = os.path.join(image_dir, desc_file)

                    slide_num = self._extract_slide_number(image_file)

                    image_data = {
                        "image_path": os.path.join(image_dir, image_file),
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
                            logger.warning(f"⚠️ 加载图片描述失败 {desc_file}: {e}")

                    result['images'].append(image_data)

            logger.info(f"✅ 数据加载完成:")
            logger.info(f"   - 幻灯片: {len(result['slides'])}个文件")
            logger.info(f"   - 图片: {len(result['images'])}个")

        except Exception as e:
            logger.error(f"❌ 数据加载失败: {e}")

        return result

    def extract_entities_from_multimodal(self, multimodal_data: Dict) -> ExtractedTriple:
        """从多模态数据中抽取实体关系"""
        all_entities = []
        all_relationships = []
        all_attributes = []

        logger.info("🔍 开始实体关系抽取...")

        # 1. 处理幻灯片文本内容
        slides = multimodal_data.get('slides', [])
        for i, slide in enumerate(slides):
            logger.info(f"   处理幻灯片 {i + 1}/{len(slides)}: {slide.get('source_file', '')}")
            slide_entities, slide_relations = self._extract_from_slide_text(slide)
            all_entities.extend(slide_entities)
            all_relationships.extend(slide_relations)
            time.sleep(0.5)  # 避免API调用过快

        # 2. 处理图片内容
        images = multimodal_data.get('images', [])
        for i, image_data in enumerate(images):
            logger.info(f"   处理图片 {i + 1}/{len(images)}: {image_data.get('filename', '')}")
            img_entities = self._extract_from_image(image_data)
            all_entities.extend(img_entities)

        # 去重处理
        all_entities = self._deduplicate_entities(all_entities)
        all_relationships = self._deduplicate_relationships(all_relationships)

        logger.info(f"✅ 实体关系抽取完成: {len(all_entities)}个实体, {len(all_relationships)}个关系")

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
            logger.warning(f"   ⚠️ 幻灯片 {slide_num} 实体抽取失败: {e}")

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

            if desc_text and confidence > 0.05:
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

        # 2. 基于OCR文本抽取实体
        ocr_text = image_data.get('ocr_text', '')
        if ocr_text:
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

        # 3. 基于文件名抽取实体
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
    multimodal_data = extractor.load_multimodal_data(output_dir)
    extracted_data = extractor.extract_entities_from_multimodal(multimodal_data)
    return extracted_data


import sys
import os

# 获取当前文件的绝对路径
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)

# 添加路径
sys.path.append(parent_dir)
sys.path.append(current_dir)

import json
import fitz  # PyMuPDF
import torch
import numpy as np
from PIL import Image, ImageEnhance
from transformers import CLIPProcessor, CLIPModel
import datetime
import pandas as pd
import cv2
import easyocr
import re
import pdfplumber
import camelot
import csv


class MultimodalPreprocessor:
    def __init__(self):
        """初始化多模态预处理工具"""
        logger.info("🚀 开始初始化多模态预处理工具...")

        # 检测设备
        logger.info("📱 检测计算设备...")
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        logger.info(f"✓ 使用设备: {self.device}")

        # 创建输出目录
        logger.info("📁 创建输出目录...")
        os.makedirs("output/text", exist_ok=True)
        os.makedirs("output/images", exist_ok=True)
        os.makedirs("output/formulas", exist_ok=True)
        os.makedirs("output/tables", exist_ok=True)
        os.makedirs("output/code", exist_ok=True)
        logger.info("✓ 输出目录创建完成")

        # 初始化CLIP模型（可能较慢）
        logger.info("🤖 正在加载CLIP模型...")
        logger.info("   ⏳ 本地模型加载中，请稍候...")
        # ===== 修改这里：使用你的本地CLIP模型路径 =====
        local_clip_path = r"F:\Models\clip-vit-base-patch32"
        try:
            if os.path.exists(local_clip_path):
                logger.info(f"   📁 找到本地模型: {local_clip_path}")
                self.clip_model = CLIPModel.from_pretrained(local_clip_path, local_files_only=True).to(self.device)
                logger.info("   ✓ 本地CLIP模型加载完成")
            else:
                logger.info(f"   ❌ 本地模型路径不存在: {local_clip_path}")
                raise FileNotFoundError("本地模型不存在")
        except Exception as e:
            logger.error(f"   ❌ 本地CLIP模型加载失败: {e}")
            logger.info("   ⏳ 尝试在线下载CLIP模型，这可能需要几分钟...")
            try:
                self.clip_model = CLIPModel.from_pretrained("openai/clip-vit-base-patch32").to(self.device)
                logger.info("   ✓ 在线CLIP模型下载并加载完成")
            except Exception as e2:
                logger.error(f"   ❌ CLIP模型加载完全失败: {e2}")
                raise e2

        logger.info("🔧 正在加载CLIP处理器...")
        logger.info("   ⏳ 处理器加载中（可能需要下载）...")
        # ===== 修改这里：使用你的本地CLIP处理器路径 =====
        local_clip_path = r"F:\Models\clip-vit-base-patch32"
        try:
            if os.path.exists(local_clip_path):
                logger.info(f"   📁 使用本地处理器: {local_clip_path}")
                self.clip_processor = CLIPProcessor.from_pretrained(local_clip_path, local_files_only=True)
                logger.info("   ✓ 本地CLIP处理器加载完成")
            else:
                logger.info(f"   ❌ 本地处理器路径不存在，使用在线版本")
                self.clip_processor = CLIPProcessor.from_pretrained("openai/clip-vit-base-patch32")
                logger.info("   ✓ 在线CLIP处理器加载完成")
        except Exception as e:
            logger.error(f"   ❌ CLIP处理器加载失败: {e}")
            raise e

        # 初始化OCR引擎（首次运行较慢）
        logger.info("👁 正在初始化OCR引擎...")
        logger.info("   ⏳ 首次运行需要下载模型文件，这可能需要几分钟，请耐心等待...")
        logger.info("   📥 正在下载中文和英文OCR模型...")
        try:
            self.ocr_reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)  # 强制使用CPU避免GPU问题
            logger.info("   ✓ OCR引擎初始化完成")
        except Exception as e:
            logger.warning(f"   ⚠ OCR初始化失败，将跳过OCR功能: {e}")
            logger.info("   将继续运行，但跳过OCR公式识别功能")
            self.ocr_reader = None

        # 存储处理结果
        self.results = []

        logger.info("🎉 多模态预处理工具初始化完成！")
        logger.info("=" * 50)

    def process_pdf(self, file_path):
        """处理PDF文件，提取文本和图像"""
        logger.info(f"开始处理PDF文件: {file_path}")
        doc = fitz.open(file_path)
        base_filename = os.path.splitext(os.path.basename(file_path))[0]

        # 为当前PDF文件单独记录结果
        current_results = []
        original_results = self.results
        self.results = current_results

        try:
            for page_num in range(len(doc)):
                logger.info(f"处理第 {page_num + 1}/{len(doc)} 页...")
                page = doc.load_page(page_num)
                page_text = page.get_text()

                # 处理页面图像
                image_list = page.get_images(full=True)
                page_images = []

                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]

                    # 保存原始图像
                    img_path = f"output/images/{base_filename}_p{page_num + 1}_img{img_index + 1}.{base_image['ext']}"
                    with open(img_path, "wb") as img_file:
                        img_file.write(image_bytes)

                    # 处理图像并保存
                    image_data = self.process_image(img_path, page_text)
                    self.save_image_data(image_data, base_filename, page_num, img_index)
                    page_images.append(img_path)

                # 处理页面文本
                text_data = self.process_text(page_text, page_num, page_images)
                text_data["source"] = f"{base_filename}_page{page_num + 1}"
                self.save_text_data(text_data, base_filename, page_num)

                # 提取页面中的公式、表格、代码
                self.extract_formulas_from_page(page, page_text, base_filename, page_num)
                self.extract_tables_from_page(page, page_text, base_filename, page_num)
                self.extract_code_from_page(page_text, base_filename, page_num)

            # 保存PDF专用元数据
            self.save_pdf_metadata(file_path, base_filename)
            logger.info(f"PDF处理完成！结果保存在output/{base_filename}_pdf_metadata.json")

        finally:
            # 恢复原始结果列表并合并当前结果
            self.results = original_results
            self.results.extend(current_results)

    def process_text(self, text, page_num, page_images):
        """处理文本内容"""
        # 清理文本
        cleaned_text = text.strip()
        if not cleaned_text:
            cleaned_text = "[页面无文本内容]"

        # 使用CLIP生成文本语义向量
        try:
            inputs = self.clip_processor(text=cleaned_text, return_tensors="pt", padding=True, truncation=True)
            inputs = {k: v.to(self.device) for k, v in inputs.items()}

            with torch.no_grad():
                text_features = self.clip_model.get_text_features(**inputs)
                text_vector = text_features.cpu().numpy()[0]
        except Exception as e:
            logger.error(f"文本向量化失败: {e}")
            text_vector = np.zeros(512)  # CLIP默认向量维度

        return {
            "type": "text",
            "page": page_num + 1,
            "raw_text": cleaned_text,
            "word_count": len(cleaned_text),
            "associated_images": page_images,
            "text_vector": text_vector.tolist()
        }

    def process_image(self, image_path, page_text):
        """处理图像（使用CLIP）"""
        # 图像增强
        enhanced_path = self.enhance_image(image_path)

        # 获取图像基本信息
        try:
            with Image.open(image_path) as img:
                width, height = img.size
                format_type = img.format
                mode = img.mode
        except Exception as e:
            logger.error(f"读取图像信息失败: {e}")
            width = height = 0
            format_type = mode = "unknown"

        # 使用CLIP生成图像向量和描述
        try:
            image = Image.open(enhanced_path)
            inputs = self.clip_processor(images=image, return_tensors="pt").to(self.device)

            with torch.no_grad():
                image_features = self.clip_model.get_image_features(**inputs)
                image_vector = image_features.cpu().numpy()[0]

            # 生成图像描述标签
            description_tags = self.generate_image_descriptions(enhanced_path)

        except Exception as e:
            logger.error(f"图像处理失败: {e}")
            image_vector = np.zeros(512)
            description_tags = []

        return {
            "type": "image",
            "image_path": image_path,
            "enhanced_path": enhanced_path,
            "width": width,
            "height": height,
            "format": format_type,
            "mode": mode,
            "page_text_context": page_text[:200] + "..." if len(page_text) > 200 else page_text,
            "image_vector": image_vector.tolist(),
            "clip_descriptions": description_tags
        }

    def enhance_image(self, image_path):
        """图像增强处理"""
        img = Image.open(image_path)

        # 对比度增强
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.5)

        # 锐度增强
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0)

        # 保存增强后的图像
        enhanced_path = image_path.replace(".", "_enhanced.")
        img.save(enhanced_path)

        return enhanced_path

    def clip_generate_description(self, image_path: str) -> str:
        """基于CLIP为图片生成描述文本并保存为JSON，返回描述文件路径。"""
        try:
            descriptions = self.generate_image_descriptions(image_path)
        except Exception as e:
            logger.error(f"生成图片描述失败: {e}")
            descriptions = []

        base, _ = os.path.splitext(os.path.basename(image_path))
        desc_path = os.path.join("output", "images", f"{base}_desc.json")
        try:
            with open(desc_path, "w", encoding="utf-8") as f:
                json.dump({
                    "image_path": image_path,
                    "clip_descriptions": descriptions
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存图片描述失败: {e}")
        return desc_path

    def generate_image_descriptions(self, image_path):
        """使用CLIP生成图像描述标签"""
        try:
            image = Image.open(image_path)
            image_input = self.clip_processor(images=image, return_tensors="pt").to(self.device)

            # 预定义的描述候选列表
            text_descriptions = [
                "科学图表", "数学公式", "数据图表", "流程图",
                "实验装置", "分子结构", "几何图形", "统计图表",
                "技术示意图", "概念图", "网络图", "系统架构图",
                "照片", "插图", "示例图", "对比图",
                "文本图像", "表格", "代码", "截图"
            ]

            text_inputs = self.clip_processor(
                text=text_descriptions,
                return_tensors="pt",
                padding=True,
                truncation=True
            ).to(self.device)

            with torch.no_grad():
                image_features = self.clip_model.get_image_features(**image_input)
                text_features = self.clip_model.get_text_features(**text_inputs)

                # 归一化特征向量
                image_features = image_features / image_features.norm(dim=-1, keepdim=True)
                text_features = text_features / text_features.norm(dim=-1, keepdim=True)

                # 计算相似度
                similarity = (image_features @ text_features.T).softmax(dim=-1)
                values, indices = similarity[0].topk(5)  # 取前5个最相似的描述

                descriptions = []
                for value, idx in zip(values.cpu().numpy(), indices.cpu().numpy()):
                    descriptions.append({
                        "description": text_descriptions[idx],
                        "confidence": float(value)
                    })

            return descriptions

        except Exception as e:
            logger.error(f"图像描述生成错误: {image_path}, {str(e)}")
            return []

    def save_text_data(self, data, filename, page_num):
        """保存文本处理结果"""
        output_path = f"output/text/{filename}_p{page_num + 1}.json"
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        self.results.append({
            "type": "text",
            "page": page_num + 1,
            "file": output_path
        })

    def save_image_data(self, data, filename, page_num, img_index):
        """保存图像处理结果"""
        output_path = f"output/images/{filename}_p{page_num + 1}_img{img_index + 1}.json"
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        self.results.append({
            "type": "image",
            "page": page_num + 1,
            "file": output_path
        })

    def save_pdf_metadata(self, file_path, filename):
        """保存PDF专用元数据文件"""
        # 统计信息
        text_files = [r for r in self.results if r["type"] == "text"]
        image_files = [r for r in self.results if r["type"] == "image"]
        formula_files = [r for r in self.results if r["type"] == "formula"]
        table_files = [r for r in self.results if r["type"] == "table"]
        code_files = [r for r in self.results if r["type"] == "code"]

        # 计算表格统计
        total_table_rows = sum(r.get("rows", 0) for r in table_files)
        total_table_columns = sum(r.get("columns", 0) for r in table_files)

        # 计算代码统计
        total_code_lines = sum(r.get("line_count", 0) for r in code_files)
        code_languages = list(set(r.get("language", "unknown") for r in code_files))

        metadata = {
            "processing_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source_file": filename,
            "source_path": file_path,
            "file_type": "PDF",
            "processing_method": "pymupdf_ocr_camelot",
            "statistics": {
                "total_pages": len(set(r["page"] for r in self.results)),
                "total_text_blocks": len(text_files),
                "total_images": len(image_files),
                "total_formulas": len(formula_files),
                "total_tables": len(table_files),
                "total_table_rows": total_table_rows,
                "total_table_columns": total_table_columns,
                "total_code_blocks": len(code_files),
                "total_code_lines": total_code_lines,
                "code_languages": code_languages
            },
            "files": {
                "text_files": text_files,
                "image_files": image_files,
                "formula_files": formula_files,
                "table_files": table_files,
                "code_files": code_files
            },
            "processing_info": {
                "clip_model": "openai/clip-vit-base-patch32",
                "device": self.device,
                "output_format": "JSON/CSV",
                "ocr_enabled": self.ocr_reader is not None,
                "extraction_features": ["text", "images", "formulas", "tables", "code"]
            }
        }

        with open(f"output/{filename}_pdf_metadata.json", "w", encoding="utf-8") as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)

    def extract_formulas_from_page(self, page, page_text, filename, page_num):
        """从页面中提取数学公式"""
        formulas = []

        # 1. 从文本中提取LaTeX格式的公式
        latex_patterns = [
            r'\$\$([^$]+)\$\$',  # 块级公式 $$...$$
            r'\$([^$]+)\$',  # 行内公式 $...$
            r'\\begin\{equation\}(.*?)\\end\{equation\}',  # equation环境
            r'\\begin\{align\}(.*?)\\end\{align\}',  # align环境
            r'\\begin\{math\}(.*?)\\end\{math\}',  # math环境
        ]

        for i, pattern in enumerate(latex_patterns):
            matches = re.findall(pattern, page_text, re.DOTALL)
            for j, match in enumerate(matches):
                formula_data = {
                    "type": "formula",
                    "page": page_num + 1,
                    "formula_id": f"{filename}_p{page_num + 1}_formula{len(formulas) + 1}",
                    "content": match.strip(),
                    "format": "latex",
                    "extraction_method": f"pattern_{i + 1}",
                    "context": self.get_text_context(page_text, match, 100)
                }
                formulas.append(formula_data)

        # 2. 从图像中识别公式（使用OCR）- 简化版本避免卡住
        if self.ocr_reader and len(formulas) < 5:  # 限制OCR处理，避免卡住
            try:
                # 获取页面图像
                pix = page.get_pixmap()
                img_data = pix.tobytes("png")
                img_array = np.frombuffer(img_data, np.uint8)
                img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)

                # OCR识别，设置超时
                results = self.ocr_reader.readtext(img, width_ths=0.7, height_ths=0.7)

                for result in results[:3]:  # 只处理前3个结果，避免过多处理
                    text = result[1]
                    confidence = result[2]

                    # 检查是否包含数学符号
                    math_symbols = ['∑', '∫', '∂', '∆', '∇', '∞', '±', '≠', '≤', '≥', 'α', 'β', 'γ', 'δ', 'θ', 'λ', 'μ',
                                    'π', 'σ', 'φ', 'ψ', 'ω']
                    if any(symbol in text for symbol in math_symbols) and confidence > 0.6:
                        if any(char.isdigit() or char in '+-*/=()[]{}^_' for char in text):
                            formula_data = {
                                "type": "formula",
                                "page": page_num + 1,
                                "formula_id": f"{filename}_p{page_num + 1}_formula{len(formulas) + 1}",
                                "content": text,
                                "format": "ocr_text",
                                "confidence": float(confidence),
                                "extraction_method": "ocr",
                                "bbox": [[float(pt[0]), float(pt[1])] for pt in result[0]]  # 转换为Python原生类型
                            }
                            formulas.append(formula_data)

            except Exception as e:
                logger.error(f"OCR公式识别失败 (页面 {page_num + 1}): {e}")

        # 保存公式数据
        if formulas:
            self.save_formulas_data(formulas, filename, page_num)

        return formulas

    def extract_tables_from_page(self, page, page_text, filename, page_num):
        """从页面中提取表格"""
        tables = []

        try:
            # 使用camelot提取表格 - 限制处理时间，避免卡住
            pdf_path = None
            for file in os.listdir("input"):
                if file.lower().endswith('.pdf') and filename in file:
                    pdf_path = os.path.join("input", file)
                    break

            if pdf_path and os.path.exists(pdf_path) and page_num < 10:  # 只处理前10页，避免卡住
                # 提取当前页面的表格，限制处理
                camelot_tables = camelot.read_pdf(pdf_path, pages=str(page_num + 1), flavor='lattice')

                for i, table in enumerate(camelot_tables[:2]):  # 只处理前2个表格
                    if table.df is not None and not table.df.empty and len(table.df) > 0:
                        table_data = {
                            "type": "table",
                            "page": page_num + 1,
                            "table_id": f"{filename}_p{page_num + 1}_table{i + 1}",
                            "rows": len(table.df),
                            "columns": len(table.df.columns),
                            "data": table.df.to_dict('records'),
                            "extraction_method": "camelot",
                            "accuracy": getattr(table, 'accuracy', 0.0)
                        }
                        tables.append(table_data)

        except Exception as e:
            logger.info(f"Camelot表格提取跳过 (页面 {page_num + 1}): {e}")

        # 备用方法：从文本中识别表格模式
        table_patterns = self.detect_text_tables(page_text)
        for i, pattern in enumerate(table_patterns):
            table_data = {
                "type": "table",
                "page": page_num + 1,
                "table_id": f"{filename}_p{page_num + 1}_texttable{i + 1}",
                "content": pattern,
                "extraction_method": "text_pattern",
                "context": self.get_text_context(page_text, pattern, 50)
            }
            tables.append(table_data)

        # 保存表格数据
        if tables:
            self.save_tables_data(tables, filename, page_num)

        return tables

    def extract_code_from_page(self, page_text, filename, page_num):
        """从页面文本中提取代码块"""
        code_blocks = []

        # 代码块模式
        code_patterns = [
            r'```(\w*)\n(.*?)```',  # Markdown代码块
            r'`([^`]+)`',  # 行内代码
            r'(?:^|\n)((?:    |\t)[^\n]+(?:\n(?:    |\t)[^\n]+)*)',  # 缩进代码块
        ]

        # 编程语言关键字
        programming_keywords = [
            'def ', 'class ', 'import ', 'from ', 'if ', 'else:', 'for ', 'while ', 'return',
            'function', 'var ', 'let ', 'const ', 'console.log', 'print(', 'println',
            'public ', 'private ', 'static ', 'void ', 'int ', 'String ', 'boolean',
            '#include', 'using namespace', 'int main(', 'printf(', 'cout <<'
        ]

        for i, pattern in enumerate(code_patterns):
            matches = re.findall(pattern, page_text, re.DOTALL | re.MULTILINE)

            for j, match in enumerate(matches):
                if isinstance(match, tuple):
                    language = match[0] if match[0] else "unknown"
                    content = match[1] if len(match) > 1 else match[0]
                else:
                    content = match
                    language = "unknown"

                # 检查是否包含编程关键字
                if any(keyword in content for keyword in programming_keywords) or len(content.strip()) > 20:
                    code_data = {
                        "type": "code",
                        "page": page_num + 1,
                        "code_id": f"{filename}_p{page_num + 1}_code{len(code_blocks) + 1}",
                        "content": content.strip(),
                        "language": language,
                        "extraction_method": f"pattern_{i + 1}",
                        "line_count": len(content.strip().split('\n')),
                        "context": self.get_text_context(page_text, content, 100)
                    }
                    code_blocks.append(code_data)

        # 保存代码数据
        if code_blocks:
            self.save_code_data(code_blocks, filename, page_num)

        return code_blocks

    def get_text_context(self, full_text, target_text, context_length=100):
        """获取目标文本的上下文"""
        try:
            index = full_text.find(target_text)
            if index == -1:
                return target_text

            start = max(0, index - context_length)
            end = min(len(full_text), index + len(target_text) + context_length)
            return full_text[start:end]
        except:
            return target_text

    def detect_text_tables(self, text):
        """从文本中检测表格模式"""
        tables = []
        lines = text.split('\n')

        # 寻找包含多个制表符或空格分隔的行
        table_lines = []
        for line in lines:
            # 检查是否包含表格特征：多个制表符、竖线分隔符等
            if '\t' in line and line.count('\t') >= 2:
                table_lines.append(line)
            elif '|' in line and line.count('|') >= 2:
                table_lines.append(line)
            elif re.search(r'\s{3,}', line) and len(line.split()) >= 3:
                table_lines.append(line)
            else:
                if table_lines and len(table_lines) >= 2:
                    tables.append('\n'.join(table_lines))
                table_lines = []

        # 检查最后一组
        if table_lines and len(table_lines) >= 2:
            tables.append('\n'.join(table_lines))

        return tables

    def save_formulas_data(self, formulas, filename, page_num):
        """保存公式数据"""
        # JSON格式保存
        json_path = f"output/formulas/{filename}_p{page_num + 1}_formulas.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(formulas, f, ensure_ascii=False, indent=2)

        # CSV格式保存
        csv_path = f"output/formulas/{filename}_p{page_num + 1}_formulas.csv"
        if formulas:
            df = pd.DataFrame(formulas)
            df.to_csv(csv_path, index=False, encoding="utf-8")

        # 记录到结果
        for formula in formulas:
            self.results.append({
                "type": "formula",
                "page": page_num + 1,
                "file": json_path,
                "formula_id": formula["formula_id"]
            })

    def save_tables_data(self, tables, filename, page_num):
        """保存表格数据"""
        # JSON格式保存
        json_path = f"output/tables/{filename}_p{page_num + 1}_tables.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(tables, f, ensure_ascii=False, indent=2)

        # 为每个表格单独保存CSV
        for i, table in enumerate(tables):
            if table.get("data") and isinstance(table["data"], list):
                csv_path = f"output/tables/{table['table_id']}.csv"
                try:
                    df = pd.DataFrame(table["data"])
                    df.to_csv(csv_path, index=False, encoding="utf-8")
                except Exception as e:
                    logger.error(f"保存表格CSV失败: {e}")

        # 记录到结果
        for table in tables:
            self.results.append({
                "type": "table",
                "page": page_num + 1,
                "file": json_path,
                "table_id": table["table_id"],
                "rows": table.get("rows", 0),
                "columns": table.get("columns", 0)
            })

    def save_code_data(self, code_blocks, filename, page_num):
        """保存代码数据"""
        # JSON格式保存
        json_path = f"output/code/{filename}_p{page_num + 1}_code.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(code_blocks, f, ensure_ascii=False, indent=2)

        # CSV格式保存
        csv_path = f"output/code/{filename}_p{page_num + 1}_code.csv"
        if code_blocks:
            df = pd.DataFrame(code_blocks)
            df.to_csv(csv_path, index=False, encoding="utf-8")

        # 为每个代码块单独保存文件
        for code in code_blocks:
            if code.get("language") and code.get("language") != "unknown":
                ext = self.get_file_extension(code["language"])
                code_file_path = f"output/code/{code['code_id']}.{ext}"
                with open(code_file_path, "w", encoding="utf-8") as f:
                    f.write(code["content"])

        # 记录到结果
        for code in code_blocks:
            self.results.append({
                "type": "code",
                "page": page_num + 1,
                "file": json_path,
                "code_id": code["code_id"],
                "language": code.get("language", "unknown"),
                "line_count": code.get("line_count", 0)
            })

    def get_file_extension(self, language):
        """根据编程语言获取文件扩展名"""
        extensions = {
            "python": "py",
            "javascript": "js",
            "java": "java",
            "cpp": "cpp",
            "c": "c",
            "csharp": "cs",
            "php": "php",
            "ruby": "rb",
            "go": "go",
            "rust": "rs",
            "swift": "swift",
            "kotlin": "kt",
            "typescript": "ts",
            "html": "html",
            "css": "css",
            "sql": "sql",
            "shell": "sh",
            "bash": "sh",
            "powershell": "ps1"
        }
        return extensions.get(language.lower(), "txt")


import os
import json
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd
import re
import subprocess


class AdvancedPPTProcessor:
    def __init__(self, preprocessor, fast_mode=False):
        """
        初始化高级PPTX处理器

        Args:
            preprocessor: MultimodalPreprocessor实例，用于重用输出目录与结果记录
            fast_mode: 快速模式，跳过耗时的CLIP描述生成
        """
        self.preprocessor = preprocessor
        self.fast_mode = fast_mode
        self.output_text_dir = "output/text"
        self.output_table_dir = "output/tables"
        self.output_img_dir = "output/images"

        # 确保输出目录存在
        os.makedirs(self.output_text_dir, exist_ok=True)
        os.makedirs(self.output_table_dir, exist_ok=True)
        os.makedirs(self.output_img_dir, exist_ok=True)

    def extract_all_images_via_zip(self, file_path):
        """
        通过ZIP解压和XML解析提取PPTX中的所有图片

        Args:
            file_path: PPTX文件路径

        Returns:
            dict: 包含幻灯片到图片映射关系的字典
        """
        logger.info(f"开始通过ZIP方式提取图片: {file_path}")

        base_filename = os.path.splitext(os.path.basename(file_path))[0]
        slide_image_mapping = {}

        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # 1. 解压PPTX文件
                logger.info("正在解压PPTX文件...")
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)

                # 2. 找到媒体目录
                media_dir = os.path.join(temp_dir, "ppt", "media")
                slides_dir = os.path.join(temp_dir, "ppt", "slides")
                rels_dir = os.path.join(temp_dir, "ppt", "slides", "_rels")

                if not os.path.exists(media_dir):
                    logger.info("未找到media目录，可能没有图片")
                    return slide_image_mapping

                logger.info(f"找到media目录: {media_dir}")
                logger.info(f"媒体文件: {os.listdir(media_dir)}")

                # 3. 遍历所有幻灯片XML文件
                if os.path.exists(slides_dir):
                    for slide_file in os.listdir(slides_dir):
                        if slide_file.startswith("slide") and slide_file.endswith(".xml"):
                            slide_num = self._extract_slide_number(slide_file)
                            if slide_num is None:
                                continue

                            logger.info(f"处理幻灯片 {slide_num}: {slide_file}")

                            # 解析幻灯片XML获取图片关系ID
                            slide_xml_path = os.path.join(slides_dir, slide_file)
                            image_rids = self._parse_slide_xml_for_images(slide_xml_path)

                            if image_rids:
                                logger.info(f"幻灯片 {slide_num} 中找到图片关系ID: {image_rids}")

                                # 解析关系文件获取实际文件名
                                rels_file = slide_file + ".rels"
                                rels_path = os.path.join(rels_dir, rels_file)

                                if os.path.exists(rels_path):
                                    image_files = self._parse_rels_file(rels_path, image_rids)

                                    if image_files:
                                        slide_image_mapping[slide_num] = image_files
                                        logger.info(f"幻灯片 {slide_num} 映射到图片: {image_files}")

                                        # 复制图片到输出目录
                                        self._copy_images_to_output(media_dir, image_files,
                                                                    base_filename, slide_num)

                logger.info(f"图片提取完成，映射关系: {slide_image_mapping}")

            except Exception as e:
                logger.error(f"ZIP方式图片提取失败: {e}")
                import traceback
                traceback.print_exc()

        return slide_image_mapping

    def _extract_slide_number(self, slide_filename):
        """从幻灯片文件名中提取编号"""
        try:
            # slide1.xml -> 1
            import re
            match = re.search(r'slide(\d+)\.xml', slide_filename)
            if match:
                return int(match.group(1))
        except Exception:
            pass
        return None

    def _parse_slide_xml_for_images(self, slide_xml_path):
        """
        解析幻灯片XML文件，查找图片引用

        Args:
            slide_xml_path: 幻灯片XML文件路径

        Returns:
            list: 图片关系ID列表
        """
        image_rids = []

        try:
            tree = ET.parse(slide_xml_path)
            root = tree.getroot()

            # 定义命名空间
            namespaces = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
            }

            # 查找所有a:blip元素（图片引用）
            blip_elements = root.findall('.//a:blip', namespaces)

            for blip in blip_elements:
                embed_attr = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed_attr:
                    image_rids.append(embed_attr)
                    logger.info(f"找到图片引用ID: {embed_attr}")

        except Exception as e:
            logger.error(f"解析幻灯片XML失败 {slide_xml_path}: {e}")

        return image_rids

    def _parse_rels_file(self, rels_path, image_rids):
        """
        解析关系文件，获取关系ID到文件名的映射

        Args:
            rels_path: 关系文件路径
            image_rids: 图片关系ID列表

        Returns:
            list: 对应的图片文件名列表
        """
        image_files = []

        try:
            tree = ET.parse(rels_path)
            root = tree.getroot()

            # 定义命名空间
            namespaces = {
                'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
            }

            # 查找所有关系
            for relationship in root.findall('.//rel:Relationship', namespaces):
                rel_id = relationship.get('Id')
                target = relationship.get('Target')
                rel_type = relationship.get('Type')

                # 检查是否是图片关系
                if (rel_id in image_rids and
                        target and
                        rel_type and
                        'image' in rel_type.lower()):
                    # 提取文件名 (../media/image1.png -> image1.png)
                    filename = os.path.basename(target)
                    image_files.append(filename)
                    logger.info(f"关系映射: {rel_id} -> {filename}")

        except Exception as e:
            logger.error(f"解析关系文件失败 {rels_path}: {e}")

        return image_files

    def _copy_images_to_output(self, media_dir, image_files, base_filename, slide_num):
        """
        将图片复制到输出目录并生成描述

        Args:
            media_dir: 媒体文件源目录
            image_files: 图片文件名列表
            base_filename: 基础文件名
            slide_num: 幻灯片编号
        """
        for idx, image_file in enumerate(image_files, 1):
            try:
                source_path = os.path.join(media_dir, image_file)

                if os.path.exists(source_path):
                    # 生成输出文件名
                    file_ext = os.path.splitext(image_file)[1]
                    output_filename = f"{base_filename}_slide_{slide_num}_img_{idx}_zip{file_ext}"
                    output_path = os.path.join(self.output_img_dir, output_filename)

                    # 复制图片
                    shutil.copy2(source_path, output_path)
                    logger.info(f"复制图片: {source_path} -> {output_path}")

                    # 生成CLIP描述（根据模式决定是否生成）
                    desc_path = None
                    if not self.fast_mode:
                        try:
                            desc_path = self.preprocessor.clip_generate_description(output_path)
                        except Exception as e:
                            logger.error(f"生成CLIP描述失败，跳过: {e}")
                    else:
                        logger.info("快速模式：跳过CLIP描述生成")

                    # 记录到结果中
                    self.preprocessor.results.append({
                        "type": "ppt_image_zip",
                        "page": slide_num,
                        "file": output_path,
                        "description_file": desc_path,
                        "extraction_method": "zip_xml_parsing",
                        "original_filename": image_file
                    })

                else:
                    logger.warning(f"源图片文件不存在: {source_path}")

            except Exception as e:
                logger.error(f"复制图片失败 {image_file}: {e}")

    def extract_and_convert_equations(self, slide, slide_number):
        """
        处理幻灯片中的公式

        Args:
            slide: python-pptx的Slide对象
            slide_number: 幻灯片编号

        Returns:
            list: 包含公式信息的列表
        """
        equations = []

        for shape_index, shape in enumerate(slide.shapes):
            try:
                # 检查形状是否包含文本框
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    # 获取形状的XML内容
                    shape_xml = self._get_shape_xml(shape)
                    if shape_xml:
                        # 检查是否包含OMML公式标签
                        omml_content = self._extract_omml_from_xml(shape_xml)
                        if omml_content:
                            logger.info(f"在幻灯片 {slide_number} 形状 {shape_index} 中发现OMML公式")

                            # 尝试转换OMML到LaTeX
                            latex_content = self._convert_omml_to_latex(omml_content)

                            equation_info = {
                                "slide_number": slide_number,
                                "shape_index": shape_index,
                                "type": "omml_formula",
                                "original_omml": omml_content[:500] + "..." if len(
                                    omml_content) > 500 else omml_content,
                                "latex": latex_content,
                                "conversion_success": latex_content is not None
                            }

                            equations.append(equation_info)

                            # 添加到结果中
                            self.preprocessor.results.append({
                                "type": "formula",
                                "page": slide_number,
                                "formula_type": "omml",
                                "latex": latex_content,
                                "source": f"slide_{slide_number}_shape_{shape_index}",
                                "conversion_method": "omml_to_latex"
                            })

                            continue

                # 如果没有找到OMML，检查是否为可能的公式图片
                if self._is_potential_formula_image(shape):
                    logger.info(f"在幻灯片 {slide_number} 形状 {shape_index} 中发现潜在公式图片")

                    # 使用图片处理流程处理公式图片
                    formula_image_path = self._process_formula_image(shape, slide_number, shape_index)

                    if formula_image_path:
                        equation_info = {
                            "slide_number": slide_number,
                            "shape_index": shape_index,
                            "type": "formula_image",
                            "image_path": formula_image_path,
                            "latex": None,
                            "conversion_success": False
                        }

                        equations.append(equation_info)

                        # 添加到结果中
                        self.preprocessor.results.append({
                            "type": "formula",
                            "page": slide_number,
                            "formula_type": "image",
                            "image_path": formula_image_path,
                            "source": f"slide_{slide_number}_shape_{shape_index}",
                            "conversion_method": "image_fallback"
                        })

            except Exception as e:
                logger.error(f"处理幻灯片 {slide_number} 形状 {shape_index} 时出错: {e}")
                continue

        return equations

    def _get_shape_xml(self, shape):
        """获取形状的XML内容"""
        try:
            # 尝试获取形状的内部XML
            if hasattr(shape, '_element'):
                return ET.tostring(shape._element, encoding='unicode')
        except Exception as e:
            logger.error(f"获取形状XML失败: {e}")
        return None

    def _extract_omml_from_xml(self, xml_string):
        """从XML中提取OMML内容"""
        try:
            # 查找OMML数学标签
            omml_patterns = [
                r'<m:oMath[^>]*>.*?</m:oMath>',
                r'<m:oMathPara[^>]*>.*?</m:oMathPara>',
                r'<math[^>]*>.*?</math>'  # 也检查标准MathML
            ]

            for pattern in omml_patterns:
                matches = re.findall(pattern, xml_string, re.DOTALL | re.IGNORECASE)
                if matches:
                    return matches[0]

        except Exception as e:
            logger.error(f"提取OMML失败: {e}")
        return None

    def _convert_omml_to_latex(self, omml_content):
        """将OMML转换为LaTeX"""
        try:
            # 方法1: 尝试使用pandoc
            latex_result = self._convert_via_pandoc(omml_content)
            if latex_result:
                return latex_result

            # 方法2: 简单的文本替换作为备选方案
            latex_result = self._simple_omml_to_latex(omml_content)
            if latex_result:
                return latex_result

        except Exception as e:
            logger.error(f"OMML转LaTeX失败: {e}")

        return None

    def _convert_via_pandoc(self, omml_content):
        """使用pandoc转换OMML到LaTeX"""
        try:
            # 检查pandoc是否可用
            subprocess.run(['pandoc', '--version'],
                           capture_output=True, check=True)

            # 创建临时文件
            with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as temp_file:
                temp_file.write(f'<root>{omml_content}</root>')
                temp_file_path = temp_file.name

            try:
                # 使用pandoc转换
                result = subprocess.run([
                    'pandoc',
                    '-f', 'docx',
                    '-t', 'latex',
                    temp_file_path
                ], capture_output=True, text=True, check=True)

                return result.stdout.strip()

            finally:
                os.unlink(temp_file_path)

        except (subprocess.CalledProcessError, FileNotFoundError):
            logger.info("Pandoc不可用，跳过pandoc转换")
        except Exception as e:
            logger.error(f"Pandoc转换失败: {e}")

        return None

    def _simple_omml_to_latex(self, omml_content):
        """简单的OMML到LaTeX转换（基本文本替换）"""
        try:
            # 移除XML标签，提取纯文本
            text_content = re.sub(r'<[^>]+>', '', omml_content)
            text_content = text_content.strip()

            if not text_content:
                return None

            # 基本的数学符号替换
            replacements = {
                '≈': r'\approx',
                '≠': r'\neq',
                '≤': r'\leq',
                '≥': r'\geq',
                '∞': r'\infty',
                'α': r'\alpha',
                'β': r'\beta',
                'γ': r'\gamma',
                'δ': r'\delta',
                'θ': r'\theta',
                'λ': r'\lambda',
                'μ': r'\mu',
                'π': r'\pi',
                'σ': r'\sigma',
                'φ': r'\phi',
                'ω': r'\omega',
                '∑': r'\sum',
                '∫': r'\int',
                '√': r'\sqrt',
                '±': r'\pm',
                '×': r'\times',
                '÷': r'\div'
            }

            for symbol, latex in replacements.items():
                text_content = text_content.replace(symbol, latex)

            # 包装在数学环境中
            return f"${text_content}$"

        except Exception as e:
            logger.error(f"简单转换失败: {e}")

        return None

    def _is_potential_formula_image(self, shape):
        """判断形状是否可能是公式图片"""
        try:
            # 检查是否为图片类型
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return True

            # 检查是否为包含复杂路径的形状（可能是矢量公式）
            if shape.shape_type in [MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.AUTO_SHAPE]:
                return True

            # 检查形状大小（小的形状可能是公式）
            if hasattr(shape, 'width') and hasattr(shape, 'height'):
                # 假设公式通常比较小（宽度和高度都小于某个阈值）
                max_formula_size = 200000  # EMU单位
                if shape.width < max_formula_size and shape.height < max_formula_size:
                    return True

        except Exception as e:
            logger.error(f"检查潜在公式图片失败: {e}")

        return False

    def _process_formula_image(self, shape, slide_number, shape_index):
        """处理公式图片"""
        try:
            # 如果是图片类型，尝试导出图片
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image_filename = f"formula_slide_{slide_number}_shape_{shape_index}.png"
                image_path = os.path.join(self.output_img_dir, image_filename)

                # 这里需要实现图片导出逻辑
                # 由于python-pptx的限制，可能需要使用其他方法
                logger.info(f"识别到公式图片，但需要额外的导出逻辑: {image_filename}")

                return image_path

        except Exception as e:
            logger.error(f"处理公式图片失败: {e}")

        return None

    def process_pptx_file_advanced(self, file_path):
        """
        高级PPTX处理：结合传统方法和ZIP解析

        Args:
            file_path: PPTX文件路径
        """
        logger.info(f"开始高级PPTX处理: {file_path}")
        base_filename = os.path.splitext(os.path.basename(file_path))[0]

        # 重置结果记录，为当前PPTX文件单独记录
        current_results = []
        original_results = self.preprocessor.results
        self.preprocessor.results = current_results

        try:
            # 1. 使用传统python-pptx方法处理文本和表格
            self._process_text_and_tables_traditional(file_path, base_filename)

            # 2. 使用ZIP方法提取所有图片
            slide_image_mapping = self.extract_all_images_via_zip(file_path)

            # 3. 生成PPTX专用元数据
            self._save_pptx_metadata(file_path, base_filename, slide_image_mapping)

            logger.info(f"高级PPTX处理完成: {file_path}")
            logger.info(f"PPTX元数据已保存: output/{base_filename}_pptx_metadata.json")

        except Exception as e:
            logger.error(f"高级PPTX处理失败: {e}")
            import traceback
            traceback.print_exc()
        finally:
            # 恢复原始结果列表并合并当前结果
            self.preprocessor.results = original_results
            self.preprocessor.results.extend(current_results)

    def _process_text_and_tables_traditional(self, file_path, base_filename):
        """使用传统python-pptx方法处理文本和表格"""
        prs = Presentation(file_path)

        for slide_index, slide in enumerate(prs.slides, start=1):
            # 处理公式
            equations = self.extract_and_convert_equations(slide, slide_index)
            if equations:
                logger.info(f"在幻灯片 {slide_index} 中找到 {len(equations)} 个公式")

                # 保存公式信息到JSON文件
                formulas_json_path = f"output/formulas/{base_filename}_slide_{slide_index}_formulas.json"
                os.makedirs("output/formulas", exist_ok=True)

                formulas_output = {
                    "slide_number": slide_index,
                    "source_file": base_filename,
                    "equations_count": len(equations),
                    "equations": equations,
                    "processing_date": str(pd.Timestamp.now())
                }

                with open(formulas_json_path, "w", encoding="utf-8") as f:
                    json.dump(formulas_output, f, ensure_ascii=False, indent=2)

                logger.info(f"公式信息已保存到: {formulas_json_path}")

            # 提取文本
            slide_text_items = []
            for shape in slide.shapes:
                if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                    text_content = []
                    for paragraph in shape.text_frame.paragraphs:
                        runs_text = ''.join(run.text for run in paragraph.runs)
                        text_content.append(runs_text if runs_text else paragraph.text)
                    final_text = "\n".join([t for t in text_content if t is not None])
                    if final_text.strip():
                        slide_text_items.append(final_text.strip())

            if slide_text_items:
                text_output = {
                    "type": "ppt_text",
                    "page": slide_index,
                    "source": f"{base_filename}_slide_{slide_index}",
                    "raw_text": "\n\n".join(slide_text_items)
                }
                text_json_path = f"{self.output_text_dir}/{base_filename}_slide_{slide_index}.json"
                with open(text_json_path, "w", encoding="utf-8") as f:
                    json.dump(text_output, f, ensure_ascii=False, indent=2)
                self.preprocessor.results.append({
                    "type": "ppt_text",
                    "page": slide_index,
                    "file": text_json_path
                })

            # 提取表格（优化版本）
            table_counter = 0
            for shape in slide.shapes:
                if hasattr(shape, "has_table") and shape.has_table:
                    table_counter += 1
                    table = shape.table

                    # 获取表格位置信息（增强功能）
                    table_position = {
                        "left": float(shape.left.inches) if shape.left else 0,
                        "top": float(shape.top.inches) if shape.top else 0,
                        "width": float(shape.width.inches) if shape.width else 0,
                        "height": float(shape.height.inches) if shape.height else 0
                    }

                    # 提取表格数据
                    data_matrix = []
                    for row in table.rows:
                        row_values = []
                        for cell in row.cells:
                            # 优化：使用 text_frame.text 获取纯文本
                            try:
                                if cell.text_frame and cell.text_frame.text:
                                    cell_text = cell.text_frame.text.strip()
                                else:
                                    cell_text = cell.text.strip() if cell.text else ""
                            except Exception as e:
                                logger.error(f"提取单元格文本失败: {e}")
                                cell_text = ""
                            row_values.append(cell_text)
                        data_matrix.append(row_values)

                    # 检查是否有有效数据
                    if data_matrix and any(any(cell for cell in row) for row in data_matrix):
                        # 使用pandas DataFrame保存为CSV
                        df = pd.DataFrame(data_matrix)
                        csv_path = f"{self.output_table_dir}/{base_filename}_slide_{slide_index}_table_{table_counter}.csv"
                        df.to_csv(csv_path, index=False, header=False, encoding="utf-8")

                        # 创建表格JSON元数据文件
                        table_metadata = {
                            "type": "ppt_table",
                            "source": f"{base_filename}_slide_{slide_index}",
                            "slide_number": slide_index,
                            "table_index": table_counter,
                            "dimensions": {
                                "rows": len(data_matrix),
                                "columns": len(data_matrix[0]) if data_matrix else 0
                            },
                            "position": table_position,
                            "data_preview": data_matrix[:3] if len(data_matrix) > 0 else [],  # 前3行预览
                            "csv_file": csv_path
                        }

                        # 保存表格元数据JSON
                        json_path = f"{self.output_table_dir}/{base_filename}_slide_{slide_index}_table_{table_counter}.json"
                        with open(json_path, "w", encoding="utf-8") as f:
                            json.dump(table_metadata, f, ensure_ascii=False, indent=2)

                        # 记录到主结果中（增强元数据）
                        self.preprocessor.results.append({
                            "type": "ppt_table",
                            "page": slide_index,
                            "table_index": table_counter,
                            "file": csv_path,
                            "metadata_file": json_path,
                            "dimensions": {
                                "rows": len(data_matrix),
                                "columns": len(data_matrix[0]) if data_matrix else 0
                            },
                            "position": table_position,
                            "extraction_method": "python_pptx_optimized"
                        })

                        logger.info(
                            f"✓ 提取表格 {table_counter}: {len(data_matrix)}行 x {len(data_matrix[0]) if data_matrix else 0}列")
                        logger.info(f"  位置: left={table_position['left']:.2f}in, top={table_position['top']:.2f}in")
                    else:
                        logger.warning(f"⚠ 跳过空表格 {table_counter}")

    def _save_pptx_metadata(self, file_path, base_filename, slide_image_mapping):
        """
        保存PPTX专用元数据文件

        Args:
            file_path: 原始PPTX文件路径
            base_filename: 基础文件名
            slide_image_mapping: 幻灯片到图片的映射关系
        """
        import datetime

        # 统计各类型文件
        text_files = [r for r in self.preprocessor.results if r["type"] == "ppt_text"]
        table_files = [r for r in self.preprocessor.results if r["type"] == "ppt_table"]
        image_files_traditional = [r for r in self.preprocessor.results if r["type"] == "ppt_image"]
        image_files_zip = [r for r in self.preprocessor.results if r["type"] == "ppt_image_zip"]

        # 计算幻灯片统计
        total_slides = len(set(r["page"] for r in self.preprocessor.results if "page" in r))
        slides_with_images = len(slide_image_mapping)
        total_images_zip = sum(len(images) for images in slide_image_mapping.values())

        # 计算表格统计
        total_tables = len(table_files)
        total_table_rows = sum(r.get("dimensions", {}).get("rows", 0) for r in table_files)
        total_table_columns = sum(r.get("dimensions", {}).get("columns", 0) for r in table_files)
        table_positions = [r.get("position", {}) for r in table_files if r.get("position")]

        # 构建元数据
        metadata = {
            "processing_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source_file": base_filename,
            "source_path": file_path,
            "file_type": "PPTX",
            "processing_method": "advanced_zip_xml_parsing",
            "statistics": {
                "total_slides": total_slides,
                "slides_with_text": len(text_files),
                "slides_with_tables": len([r for r in table_files if r.get("dimensions", {}).get("rows", 0) > 0]),
                "total_tables": total_tables,
                "total_table_rows": total_table_rows,
                "total_table_columns": total_table_columns,
                "slides_with_images": slides_with_images,
                "total_images_extracted": len(image_files_traditional) + len(image_files_zip),
                "images_via_traditional": len(image_files_traditional),
                "images_via_zip_parsing": len(image_files_zip),
                "total_images_in_media": total_images_zip
            },
            "slide_image_mapping": slide_image_mapping,
            "files": {
                "text_files": text_files,
                "table_files": table_files,
                "image_files_traditional": image_files_traditional,
                "image_files_zip": image_files_zip
            },
            "table_analysis": {
                "positions": table_positions,
                "position_stats": {
                    "avg_left": sum(p.get("left", 0) for p in table_positions) / len(
                        table_positions) if table_positions else 0,
                    "avg_top": sum(p.get("top", 0) for p in table_positions) / len(
                        table_positions) if table_positions else 0,
                    "avg_width": sum(p.get("width", 0) for p in table_positions) / len(
                        table_positions) if table_positions else 0,
                    "avg_height": sum(p.get("height", 0) for p in table_positions) / len(
                        table_positions) if table_positions else 0
                }
            },
            "processing_info": {
                "extraction_methods": ["python-pptx", "zip_xml_parsing"],
                "image_formats_supported": ["PNG", "WMF", "EMF", "JPEG"],
                "table_extraction_enhanced": True,
                "table_position_tracking": True,
                "clip_descriptions_generated": True,
                "output_format": "JSON/CSV"
            }
        }

        # 保存元数据
        metadata_path = f"output/{base_filename}_pptx_metadata.json"
        with open(metadata_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)

        return metadata_path


def process_pptx_file_advanced(preprocessor, file_path, fast_mode=False):
    """
    高级PPTX处理的入口函数

    Args:
        preprocessor: MultimodalPreprocessor实例
        file_path: PPTX文件路径
        fast_mode: 快速模式，跳过耗时处理
    """
    processor = AdvancedPPTProcessor(preprocessor, fast_mode=fast_mode)
    processor.process_pptx_file_advanced(file_path)


def build_multimodal_knowledge_graph(
        neo4j_uri: str,
        neo4j_user: str,
        neo4j_password: str,
        deepseek_api_key: str,
        input_dir: str = "input",
        output_dir: str = "output",
        document_name: str = "多模态文档",
        fast_mode: bool = False,
        clear_database: bool = False,
        verbose: bool = True
) -> dict:
    """
    构建多模态PPT/PDF知识图谱的主函数

    Args:
        neo4j_uri: Neo4j数据库URI，如 "bolt://localhost:7687"
        neo4j_user: Neo4j用户名
        neo4j_password: Neo4j密码
        deepseek_api_key: DeepSeek API密钥
        input_dir: 输入文件目录，默认为 "input"
        output_dir: 输出文件目录，默认为 "output"
        document_name: 文档名称（用于Neo4j中的文档节点），默认为 "多模态文档"
        fast_mode: 是否启用快速模式（跳过CLIP描述生成），默认为 False
        clear_database: 是否清空数据库（谨慎使用），默认为 False
        verbose: 是否显示详细输出，默认为 True

    Returns:
        dict: 包含处理结果的字典
    """

    def log(message: str):
        """条件日志输出"""
        if verbose:
            logger.info(message)

    # 初始化结果字典
    result = {
        'success': False,
        'error': None,
        'statistics': {},
        'files_processed': [],
        'entities_extracted': 0,
        'relationships_extracted': 0,
        'entities_saved': 0,
        'relationships_saved': 0,
        'neo4j_stats': {}
    }

    try:
        if verbose:
            log("=" * 80)
            log("🚀 多模态PPT知识图谱构建系统")
            log("=" * 80)
            log(f"📋 配置信息:")
            log(f"   🗄️  Neo4j: {neo4j_uri}")
            log(f"   🤖 DeepSeek API: {'已配置' if deepseek_api_key else '未配置'}")
            log(f"   📁 输入目录: {input_dir}")
            log(f"   📁 输出目录: {output_dir}")
            log(f"   ⚡ 快速模式: {'开启' if fast_mode else '关闭'}")
            log(f"   📄 文档名称: {document_name}")
            log("=" * 80)

        # ==================== 第一步：初始化多模态预处理器 ====================
        log("\n🔧 第一步：初始化多模态预处理器...")
        if verbose:
            log("⚠️  注意：首次运行可能需要下载模型，请耐心等待...")

        processor = MultimodalPreprocessor()
        log("✅ 多模态预处理器初始化完成")

        # ==================== 第二步：检查并处理输入文件 ====================
        log(f"\n📂 第二步：检查输入目录 {input_dir}...")

        # 确保输入目录存在
        if not os.path.exists(input_dir):
            os.makedirs(input_dir, exist_ok=True)
            error_msg = f"输入目录不存在，已创建: {input_dir}，请将PPT/PDF文件放入该目录"
            result['error'] = error_msg
            log(f"❌ {error_msg}")
            return result

        # 查找支持的文件
        input_files = [f for f in os.listdir(input_dir)
                       if f.lower().endswith(('.pdf', '.pptx')) and not f.startswith('~$')]

        if not input_files:
            error_msg = f"在 {input_dir} 目录中未找到PPT/PDF文件"
            result['error'] = error_msg
            log(f"❌ {error_msg}")
            return result

        ppt_count = sum(1 for f in input_files if f.lower().endswith('.pptx'))
        pdf_count = sum(1 for f in input_files if f.lower().endswith('.pdf'))
        log(f"✅ 找到文件: {ppt_count}个PPT, {pdf_count}个PDF")

        result['files_processed'] = input_files
        result['statistics']['ppt_count'] = ppt_count
        result['statistics']['pdf_count'] = pdf_count

        # ==================== 第三步：处理文件（多模态预处理）====================
        log(f"\n🔄 第三步：多模态数据预处理...")

        for idx, input_file in enumerate(input_files, 1):
            input_path = os.path.join(input_dir, input_file)
            log(f"\n📄 [{idx}/{len(input_files)}] 处理文件: {input_file}")

            if input_file.lower().endswith('.pdf'):
                log("   📚 使用PDF处理器...")
                processor.process_pdf(input_path)

            elif input_file.lower().endswith('.pptx'):
                log("   📊 使用高级PPT处理器...")
                # 创建高级PPT处理器
                ppt_processor = AdvancedPPTProcessor(processor, fast_mode=fast_mode)
                ppt_processor.process_pptx_file_advanced(input_path)

            log(f"   ✅ [{idx}/{len(input_files)}] 完成: {input_file}")

        log(f"\n🎉 多模态预处理完成！共处理 {len(input_files)} 个文件")
        log(f"📁 预处理结果保存在: {output_dir}/")

        # ==================== 第四步：实体关系抽取 ====================
        log(f"\n🔍 第四步：实体关系抽取...")

        if not deepseek_api_key:
            error_msg = "DeepSeek API Key未配置，无法进行实体抽取"
            result['error'] = error_msg
            log(f"❌ {error_msg}")
            return result

        log("   🤖 初始化实体抽取器...")
        extractor = EntityExtractor(deepseek_api_key)

        log("   📥 加载多模态数据...")
        multimodal_data = extractor.load_multimodal_data(output_dir)

        log(f"   📊 数据统计: {len(multimodal_data.get('slides', []))}个幻灯片, {len(multimodal_data.get('images', []))}张图片")

        log("   🔬 开始实体关系抽取...")
        extracted_data = extractor.extract_entities_from_multimodal(multimodal_data)

        result['entities_extracted'] = len(extracted_data.entities)
        result['relationships_extracted'] = len(extracted_data.relationships)

        log(f"✅ 实体关系抽取完成:")
        log(f"   🏷️  实体数量: {len(extracted_data.entities)}")
        log(f"   🔗 关系数量: {len(extracted_data.relationships)}")
        log(f"   📋 属性数量: {len(extracted_data.attributes)}")

        # ==================== 第五步：保存到Neo4j知识图谱 ====================
        log(f"\n💾 第五步：保存到Neo4j知识图谱...")

        log("   🔌 连接Neo4j数据库...")
        kg = Neo4jKnowledgeGraph(neo4j_uri, neo4j_user, neo4j_password)

        # 可选：清空数据库
        if clear_database:
            log("   🗑️  清空数据库...")
            kg.clear_database()

        log("   💾 保存实体关系数据...")
        save_result = kg.save_extracted_data(extracted_data, document_name)

        if save_result['success']:
            result['entities_saved'] = save_result['entities_saved']
            result['relationships_saved'] = save_result['relationships_saved']

            log(f"✅ 数据保存成功:")
            log(f"   📄 文档: {save_result['document']}")
            log(f"   🏷️  实体: {save_result['entities_saved']}/{len(extracted_data.entities)}")
            log(f"   🔗 关系: {save_result['relationships_saved']}/{len(extracted_data.relationships)}")
        else:
            error_msg = f"数据保存失败: {save_result.get('error', '未知错误')}"
            result['error'] = error_msg
            log(f"❌ {error_msg}")
            kg.close()
            return result

        # ==================== 第六步：显示知识图谱统计信息 ====================
        log(f"\n📊 第六步：知识图谱统计信息...")

        stats = kg.get_statistics()
        result['neo4j_stats'] = stats

        log(f"   📄 文档总数: {stats['total_documents']}")
        log(f"   🏷️  节点总数: {stats['total_nodes']}")
        log(f"   🔗 关系总数: {stats['total_relationships']}")

        if verbose:
            log(f"\n🏷️  实体类型分布:")
            for entity_type in stats['entity_types'][:10]:  # 显示前10种类型
                log(f"   - {entity_type['entity_type']}: {entity_type['count']}个")

            # ==================== 第七步：显示查询示例 ====================
            log(f"\n🔍 第七步：知识图谱查询示例...")

            # 查询实体示例
            log(f"\n🏷️  实体示例 (前5个):")
            entities = kg.query_entities(5)
            for i, entity in enumerate(entities, 1):
                log(f"   {i}. {entity['name']} ({entity['type']}) - {entity.get('description', '')[:50]}...")

            # 查询关系示例
            log(f"\n🔗 关系示例 (前5个):")
            relationships = kg.query_relationships(5)
            for i, rel in enumerate(relationships, 1):
                log(f"   {i}. {rel['source']} --{rel['relation']}--> {rel['target']}")

            # 搜索特定实体示例
            if entities:
                sample_entity = entities[0]['name']
                log(f"\n🔍 实体邻居查询示例 (以 '{sample_entity}' 为例):")
                neighbors = kg.get_entity_neighbors(sample_entity, depth=1)

                log(f"   邻居实体 (前3个):")
                for i, neighbor in enumerate(neighbors['neighbors'][:3], 1):
                    log(f"     {i}. {neighbor['name']} ({neighbor['type']})")

                log(f"   相关关系 (前3个):")
                for i, rel in enumerate(neighbors['relationships'][:3], 1):
                    direction = "→" if rel['direction'] == 'outgoing' else "←"
                    log(f"     {i}. {sample_entity} {direction}[{rel['relation']}] {rel['neighbor']}")

        # 关闭数据库连接
        kg.close()

        # 设置成功标志
        result['success'] = True

        # ==================== 完成总结 ====================
        if verbose:
            log(f"\n" + "=" * 80)
            log("🎉 多模态PPT知识图谱构建完成！")
            log("=" * 80)
            log(f"📋 处理总结:")
            log(f"   📁 输入文件: {len(input_files)}个 ({ppt_count}个PPT, {pdf_count}个PDF)")
            log(f"   🔄 多模态处理: ✅ 完成")
            log(f"   🔍 实体抽取: ✅ {len(extracted_data.entities)}个实体, {len(extracted_data.relationships)}个关系")
            log(f"   💾 Neo4j存储: ✅ {save_result['entities_saved']}个实体, {save_result['relationships_saved']}个关系")
            log(f"   🗄️  数据库统计: {stats['total_nodes']}个节点, {stats['total_relationships']}个关系")

            log(f"\n📁 输出文件位置:")
            log(f"   - 多模态数据: {output_dir}/")
            log(f"   - 文本数据: {output_dir}/text/")
            log(f"   - 图片数据: {output_dir}/images/")
            log(f"   - 表格数据: {output_dir}/tables/")
            log(f"   - 公式数据: {output_dir}/formulas/")
            log(f"   - 代码数据: {output_dir}/code/")

            log(f"\n🔍 Neo4j查询建议:")
            log(f"   - 查看所有实体: MATCH (e:Entity) RETURN e LIMIT 25")
            log(f"   - 查看所有关系: MATCH (s)-[r]->(t) RETURN s.name, type(r), t.name LIMIT 25")
            log(f"   - 查看文档: MATCH (d:Document) RETURN d")
            log(f"   - 搜索Arduino相关: MATCH (e:Entity) WHERE toLower(e.name) CONTAINS 'arduino' RETURN e")
            log("=" * 80)

        return result

    except KeyboardInterrupt:
        error_msg = "用户中断操作"
        result['error'] = error_msg
        log(f"\n⚠️  {error_msg}")
        return result

    except Exception as e:
        error_msg = f"系统错误: {str(e)}"
        result['error'] = error_msg
        log(f"\n❌ {error_msg}")

        if verbose:
            import traceback
            traceback.print_exc()
            log(f"\n🔧 故障排除建议:")
            log(f"   1. 检查Neo4j数据库连接")
            log(f"   2. 检查DeepSeek API Key是否有效")
            log(f"   3. 确保input目录中有PPT/PDF文件")
            log(f"   4. 检查网络连接（模型下载需要）")
            log(f"   5. 检查磁盘空间是否充足")

        return result


# 保留原来的主函数作为示例
if __name__ == "__main__":
    # 为独立运行生成 request_id
    standalone_request_id = str(uuid.uuid4())
    logger._request_filter._standalone_request_id = standalone_request_id

    logger.info(f"独立运行模式，生成 request_id: {standalone_request_id}")

    result = build_multimodal_knowledge_graph(
        neo4j_uri="bolt://101.132.130.25:7687",
        neo4j_user="neo4j",
        neo4j_password="wangshuxvan@1",
        deepseek_api_key="sk-c28ec338b39e4552b9e6bded47466442",
        input_dir=r"C:\Users\Lin\PycharmProjects\PythonProject\input",
        output_dir=r"C:\Users\Lin\PycharmProjects\PythonProject\output",
        document_name="Arduino课程PPT",
        fast_mode=False,
        clear_database=False,
        verbose=True
    )

    # 输出最终结果
    if result['success']:
        logger.info(f"✅ 处理成功完成！")
        logger.info(f"📊 处理统计:")
        logger.info(f"   - 输入文件: {len(result['files_processed'])}个")
        logger.info(f"   - PPT文件: {result['statistics']['ppt_count']}个")
        logger.info(f"   - PDF文件: {result['statistics']['pdf_count']}个")
        logger.info(f"   - 抽取实体: {result['entities_extracted']}个")
        logger.info(f"   - 抽取关系: {result['relationships_extracted']}个")
        logger.info(f"   - 保存实体: {result['entities_saved']}个")
        logger.info(f"   - 保存关系: {result['relationships_saved']}个")
        logger.info(f"   - 数据库节点总数: {result['neo4j_stats']['total_nodes']}个")
        logger.info(f"   - 数据库关系总数: {result['neo4j_stats']['total_relationships']}个")
    else:
        logger.error(f"❌ 处理失败: {result['error']}")
        logger.info("请检查配置参数和网络连接")



from neo4j import GraphDatabase
import pandas as pd
from typing import Optional, Dict, List, Tuple
import json
from collections import defaultdict


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
            print("✅ Neo4j连接成功")

        except Exception as e:
            print(f"❌ Neo4j连接失败: {e}")
            raise

    def close(self):
        """关闭连接"""
        if self.driver:
            self.driver.close()

    def clear_database(self):
        """清空数据库（谨慎使用）"""
        with self.driver.session() as session:
            session.run("MATCH (n) DETACH DELETE n")
        print("🗑️ 数据库已清空")

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
        print(f"📄 创建文档节点: {doc_name}")

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
            print(f"💾 开始保存数据到Neo4j: {ppt_name}")

            # 1. 创建PPT文档节点
            self.create_document_node(ppt_name, "ppt", {
                "total_entities": len(extracted_data.entities),
                "total_relationships": len(extracted_data.relationships)
            })

            # 2. 批量创建实体节点
            print(f"   创建 {len(extracted_data.entities)} 个实体节点...")
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
                    print(f"     ⚠️ 创建实体失败 {entity['name']}: {e}")
                    continue

            # 3. 批量创建关系
            print(f"   创建 {len(extracted_data.relationships)} 个关系...")
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
                    print(f"     ⚠️ 创建关系失败 {rel['source']}->{rel['target']}: {e}")
                    continue

            print(f"✅ 数据保存完成:")
            print(f"   📄 文档: {ppt_name}")
            print(f"   🏷️  实体: {entity_count}/{len(extracted_data.entities)}")
            print(f"   🔗 关系: {relation_count}/{len(extracted_data.relationships)}")

            return {
                'document': ppt_name,
                'entities_saved': entity_count,
                'relationships_saved': relation_count,
                'success': True
            }

        except Exception as e:
            print(f"❌ 保存数据失败: {e}")
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
        print(f"\n📊 数据库统计信息:")
        print(f"   📄 文档总数: {stats['total_documents']}")
        print(f"   🏷️  节点总数: {stats['total_nodes']}")
        print(f"   🔗 关系总数: {stats['total_relationships']}")

        # 显示实体类型分布
        print(f"\n🏷️  实体类型分布:")
        for entity_type in stats['entity_types'][:5]:
            print(f"   - {entity_type['entity_type']}: {entity_type['count']}个")

        return result

    finally:
        kg.close()


if __name__ == "__main__":
    # 配置参数
    NEO4J_URI = "bolt://101.132.130.25:7687"
    NEO4J_USER = "neo4j"
    NEO4J_PASSWORD = "wangshuxvan@1"
    DEEPSEEK_API_KEY = "sk-c28ec338b39e4552b9e6bded47466442"  # 你的API key
    OUTPUT_DIR = r"C:\Users\Lin\PycharmProjects\PythonProject\output"
    PPT_NAME = "Arduino课程PPT"  # 可以自定义PPT名称

    try:
        # 1. 从第一个代码导入并调用实体抽取功能
        from entitity_extractor import extract_entities_from_output  # 假设第一个代码保存为entity_extractor.py

        print("🔍 开始实体关系抽取...")
        extracted_data = extract_entities_from_output(OUTPUT_DIR, DEEPSEEK_API_KEY)

        print(f"\n📊 抽取结果统计:")
        print(f"   🏷️  实体数量: {len(extracted_data.entities)}")
        print(f"   🔗 关系数量: {len(extracted_data.relationships)}")
        print(f"   📋 属性数量: {len(extracted_data.attributes)}")

        # 2. 保存到Neo4j
        print(f"\n💾 开始保存到Neo4j...")
        result = save_to_neo4j(extracted_data, PPT_NAME, NEO4J_URI, NEO4J_USER, NEO4J_PASSWORD)

        if result['success']:
            print(f"\n🎉 流程完成!")
            print(f"   ✅ 实体抽取: {len(extracted_data.entities)}个")
            print(f"   ✅ 关系抽取: {len(extracted_data.relationships)}个")
            print(f"   ✅ Neo4j保存: {result['entities_saved']}个实体, {result['relationships_saved']}个关系")
        else:
            print(f"\n❌ 保存失败: {result.get('error', '未知错误')}")

        # 3. 显示一些查询结果
        print(f"\n🔍 数据库查询示例:")
        kg = Neo4jKnowledgeGraph(NEO4J_URI, NEO4J_USER, NEO4J_PASSWORD)

        # 查询实体示例
        entities = kg.query_entities(5)
        print(f"\n🏷️  实体示例:")
        for i, entity in enumerate(entities):
            print(f"   {i + 1}. {entity['name']} ({entity['type']}) - {entity.get('description', '')}")

        # 查询关系示例
        relationships = kg.query_relationships(5)
        print(f"\n🔗 关系示例:")
        for i, rel in enumerate(relationships):
            print(f"   {i + 1}. {rel['source']} --{rel['relation']}--> {rel['target']}")

        kg.close()

    except Exception as e:
        print(f"❌ 整体流程失败: {e}")
        import traceback

        traceback.print_exc()

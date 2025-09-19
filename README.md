# 多模态PPT/PDF知识图谱构建系统

一个基于深度学习和自然语言处理技术的多模态文档知识图谱构建系统，能够从PPT和PDF文档中提取文本、图像、表格等多模态信息，并构建结构化的知识图谱存储到Neo4j数据库中。

## 🌟 系统特性

### 📊 多模态数据处理
- **文档支持**: PPT(.pptx)和PDF文档的全面解析
- **文本提取**: 幻灯片文本、备注、段落和标题
- **图像处理**: 自动提取并生成语义描述
- **表格解析**: 智能识别和结构化表格数据
- **公式识别**: 数学公式的提取和转换
- **代码检测**: 代码片段的识别和分类

### 🤖 AI智能分析
- **CLIP模型**: 图像语义理解和描述生成
- **DeepSeek LLM**: 智能实体关系抽取
- **OCR识别**: 图像中的文字和公式识别
- **语义向量**: 文本和图像的向量化表示

### 🗄️ 知识图谱构建
- **Neo4j存储**: 高性能图数据库
- **实体抽取**: 智能识别文档中的关键实体
- **关系发现**: 自动发现实体间的语义关系
- **多模态关联**: 建立文本、图像、表格间的关联

### ⚡ 高效处理
- **并发处理**: 多线程加速文档处理
- **批量导入**: 大规模数据的高效存储
- **缓存机制**: 减少重复计算
- **快速模式**: 跳过耗时处理以提高速度

## 🏗️ 系统架构

┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   输入文档       │    │   多模态处理     │    │   知识图谱       │
│                │    │                │    │                │
│ • PPT文件       │───▶│ • 文本提取       │───▶│ • Neo4j存储     │
│ • PDF文件       │    │ • 图像分析       │    │ • 实体关系       │
│ • 批量处理       │    │ • 表格解析       │    │ • 语义查询       │
└─────────────────┘    │ • 公式识别       │    └─────────────────┘
│ • 代码检测       │
└─────────────────┘
│
▼
┌─────────────────┐
│   AI实体抽取     │
│                │
│ • DeepSeek LLM  │
│ • 实体识别       │
│ • 关系抽取       │
│ • 语义理解       │
└─────────────────┘



## 📋 系统要求

### 硬件要求
- **内存**: 8GB以上（推荐16GB）
- **存储**: 至少5GB可用空间
- **GPU**: 可选，支持CUDA加速
- **网络**: 首次运行需要下载模型

### 软件环境
- **Python**: 3.8或更高版本
- **操作系统**: Windows 10/11, Linux, macOS
- **Neo4j**: 4.0或更高版本
- **DeepSeek API**: 需要有效的API密钥

### 外部服务
- **Neo4j数据库**: 用于知识图谱存储
- **DeepSeek API**: 用于智能实体关系抽取

## 🚀 快速开始

### 1. 环境安装

#### 自动安装（推荐）
```bash
# 克隆项目
git clone <repository-url>
cd multimodal-knowledge-graph

# 运行自动安装脚本
python install_environment.py
手动安装
复制
# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt

# 安装spaCy中文模型
python -m spacy download zh_core_web_sm
2. 配置系统
Neo4j数据库设置
复制
# 使用Docker启动Neo4j（推荐）
docker run \
    --name neo4j \
    -p7474:7474 -p7687:7687 \
    -d \
    -v $HOME/neo4j/data:/data \
    -v $HOME/neo4j/logs:/logs \
    --env NEO4J_AUTH=neo4j/your-password \
    neo4j:latest
配置文件设置
编辑 config/config.json:


{
  "neo4j": {
    "uri": "bolt://localhost:7687",
    "user": "neo4j",
    "password": "your-password"
  },
  "deepseek": {
    "api_key": "your-deepseek-api-key",
    "model": "deepseek-chat"
  }
}
3. 准备文档

# 将PPT/PDF文件放入input目录
cp your-documents/*.pptx input/
cp your-documents/*.pdf input/
4. 运行系统
基本使用

from multimodal_kg import build_multimodal_knowledge_graph

# 配置参数
result = build_multimodal_knowledge_graph(
    neo4j_uri="bolt://localhost:7687",
    neo4j_user="neo4j",
    neo4j_password="your-password",
    deepseek_api_key="your-api-key",
    input_dir="input",
    output_dir="output",
    document_name="我的文档",
    verbose=True
)

# 检查结果
if result['success']:
    print(f"✅ 成功处理 {result['entities_saved']} 个实体")
else:
    print(f"❌ 处理失败: {result['error']}")
使用配置文件

import json
from multimodal_kg import build_multimodal_knowledge_graph

# 加载配置
with open('config/config.json', 'r') as f:
    config = json.load(f)

# 运行处理
result = build_multimodal_knowledge_graph(
    neo4j_uri=config['neo4j']['uri'],
    neo4j_user=config['neo4j']['user'],
    neo4j_password=config['neo4j']['password'],
    deepseek_api_key=config['deepseek']['api_key'],
    document_name="Arduino课程PPT"
)
📖 详细使用说明
函数参数详解

def build_multimodal_knowledge_graph(
    neo4j_uri: str,           # Neo4j数据库URI
    neo4j_user: str,          # Neo4j用户名  
    neo4j_password: str,      # Neo4j密码
    deepseek_api_key: str,    # DeepSeek API密钥
    input_dir: str = "input", # 输入目录
    output_dir: str = "output", # 输出目录
    document_name: str = "多模态文档", # 文档名称
    fast_mode: bool = False,  # 快速模式
    clear_database: bool = False, # 清空数据库
    verbose: bool = True      # 详细输出
) -> dict:
返回值说明

{
    'success': bool,              # 处理是否成功
    'error': str,                # 错误信息
    'statistics': {              # 处理统计
        'ppt_count': int,        # PPT文件数量
        'pdf_count': int         # PDF文件数量
    },
    'files_processed': list,     # 处理的文件列表
    'entities_extracted': int,   # 抽取的实体数量
    'relationships_extracted': int, # 抽取的关系数量
    'entities_saved': int,       # 保存的实体数量
    'relationships_saved': int,  # 保存的关系数量
    'neo4j_stats': dict         # Neo4j统计信息
}
🔧 高级配置
1. 快速模式

# 跳过CLIP图像描述生成，提高处理速度
result = build_multimodal_knowledge_graph(
    # ... 其他参数
    fast_mode=True,
    verbose=False
)
2. 自定义路径

import os

# 使用绝对路径
project_root = r"C:\Users\YourName\Projects\MyProject"
result = build_multimodal_knowledge_graph(
    # ... 其他参数
    input_dir=os.path.join(project_root, "documents"),
    output_dir=os.path.join(project_root, "results"),
)
3. 批量处理

def batch_process():
    configs = [
        {
            'input_dir': 'course1',
            'document_name': '课程1',
            'output_dir': 'output1'
        },
        {
            'input_dir': 'course2',
            'document_name': '课程2', 
            'output_dir': 'output2'
        }
    ]
    
    base_config = {
        'neo4j_uri': "bolt://localhost:7687",
        'neo4j_user': "neo4j",
        'neo4j_password': "password",
        'deepseek_api_key': "your-key",
        'fast_mode': True
    }
    
    for config in configs:
        result = build_multimodal_knowledge_graph(
            **{**base_config, **config}
        )
        print(f"处理 {config['document_name']}: {'成功' if result['success'] else '失败'}")
📁 输出文件结构

output/
├── text/                    # 文本数据
│   ├── document_slide_1.json
│   ├── document_slide_2.json
│   └── ...
├── images/                  # 图像数据
│   ├── document_slide_1_img_1.png
│   ├── document_slide_1_img_1_desc.json
│   └── ...
├── tables/                  # 表格数据
│   ├── document_slide_2_table_1.csv
│   ├── document_slide_2_table_1.json
│   └── ...
├── formulas/               # 公式数据
│   ├── document_slide_3_formulas.json
│   └── ...
├── code/                   # 代码片段
│   ├── document_slide_4_code.json
│   └── ...
└── document_pptx_metadata.json  # 元数据文件
🗄️ Neo4j知识图谱查询
基本查询

-- 查看所有实体
MATCH (e:Entity) RETURN e LIMIT 25

-- 查看所有关系
MATCH (s)-[r]->(t) RETURN s.name, type(r), t.name LIMIT 25

-- 查看文档信息
MATCH (d:Document) RETURN d

-- 搜索特定实体
MATCH (e:Entity) 
WHERE toLower(e.name) CONTAINS 'arduino' 
RETURN e

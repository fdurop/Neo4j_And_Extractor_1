#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多模态知识图谱构建系统 - 环境安装脚本
自动安装所需依赖和配置环境
"""

import os
import sys
import subprocess
import platform
import json
from pathlib import Path


class EnvironmentInstaller:
    def __init__(self):
        self.system = platform.system()
        self.python_version = sys.version_info
        self.project_root = Path.cwd()

    def check_python_version(self):
        """检查Python版本"""
        print("🐍 检查Python版本...")
        if self.python_version < (3, 8):
            print(f"❌ Python版本过低: {sys.version}")
            print("   需要Python 3.8或更高版本")
            return False
        print(f"✅ Python版本: {sys.version}")
        return True

    def install_pip_packages(self):
        """安装Python包依赖"""
        print("\n📦 安装Python依赖包...")

        # 升级pip
        print("   升级pip...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"],
                       check=True)

        # 安装requirements.txt中的包
        requirements_file = self.project_root / "requirements.txt"
        if requirements_file.exists():
            print(f"   从 {requirements_file} 安装依赖...")
            subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(requirements_file)],
                           check=True)
        else:
            print("   未找到requirements.txt，手动安装核心依赖...")
            self.install_core_packages()

    def install_core_packages(self):
        """手动安装核心包"""
        core_packages = [
            "torch>=1.9.0",
            "transformers>=4.20.0",
            "pillow>=8.0.0",
            "pandas>=1.3.0",
            "numpy>=1.21.0",
            "neo4j>=5.0.0",
            "requests>=2.25.0",
            "python-pptx>=0.6.21",
            "PyMuPDF>=1.20.0",
            "opencv-python>=4.5.0",
            "easyocr>=1.6.0",
            "spacy>=3.4.0",
            "chardet>=4.0.0",
            "pdfplumber>=0.7.0",
            "camelot-py[cv]>=0.10.1",
            "openpyxl>=3.0.9",
            "lxml>=4.6.0",
            "beautifulsoup4>=4.9.0"
        ]

        for package in core_packages:
            print(f"   安装 {package}...")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", package],
                               check=True, capture_output=True)
                print(f"   ✅ {package}")
            except subprocess.CalledProcessError as e:
                print(f"   ⚠️ {package} 安装失败: {e}")

    def install_spacy_model(self):
        """安装spaCy中文模型"""
        print("\n🔤 安装spaCy中文模型...")
        try:
            subprocess.run([sys.executable, "-m", "spacy", "download", "zh_core_web_sm"],
                           check=True)
            print("✅ spaCy中文模型安装完成")
        except subprocess.CalledProcessError:
            print("⚠️ spaCy中文模型安装失败，可能需要手动安装")
            print("   请运行: python -m spacy download zh_core_web_sm")

    def setup_directories(self):
        """创建必要的目录结构"""
        print("\n📁 创建项目目录结构...")

        directories = [
            "input",
            "output",
            "output/text",
            "output/images",
            "output/tables",
            "output/formulas",
            "output/code",
            "models",
            "config",
            "logs"
        ]

        for dir_name in directories:
            dir_path = self.project_root / dir_name
            dir_path.mkdir(parents=True, exist_ok=True)
            print(f"   ✅ {dir_name}/")

    def create_config_template(self):
        """创建配置文件模板"""
        print("\n⚙️ 创建配置文件模板...")

        config_template = {
            "neo4j": {
                "uri": "bolt://localhost:7687",
                "user": "neo4j",
                "password": "your-password-here"
            },
            "deepseek": {
                "api_key": "your-deepseek-api-key-here",
                "base_url": "https://api.deepseek.com/v1",
                "model": "deepseek-chat"
            },
            "models": {
                "clip_model_path": "openai/clip-vit-base-patch32",
                "local_clip_path": "F:/Models/clip-vit-base-patch32"
            },
            "processing": {
                "fast_mode": False,
                "batch_size": 1000,
                "max_workers": 3
            }
        }

        config_file = self.project_root / "config" / "config.json"
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config_template, f, indent=2, ensure_ascii=False)

        print(f"   ✅ 配置模板: {config_file}")
        print("   请编辑配置文件填入实际参数")

    def check_gpu_support(self):
        """检查GPU支持"""
        print("\n🔥 检查GPU支持...")
        try:
            import torch
            if torch.cuda.is_available():
                gpu_count = torch.cuda.device_count()
                gpu_name = torch.cuda.get_device_name(0)
                print(f"✅ 检测到 {gpu_count} 个GPU: {gpu_name}")
                print(f"   CUDA版本: {torch.version.cuda}")
            else:
                print("⚠️ 未检测到CUDA GPU，将使用CPU模式")
                print("   建议安装CUDA版本的PyTorch以获得更好性能")
        except ImportError:
            print("⚠️ PyTorch未安装，无法检查GPU支持")

    def install_system_dependencies(self):
        """安装系统级依赖"""
        print("\n🔧 检查系统依赖...")

        if self.system == "Windows":
            print("   Windows系统检测到")
            print("   请确保已安装Visual C++ Redistributable")

        elif self.system == "Linux":
            print("   Linux系统检测到")
            print("   可能需要安装额外系统包:")
            print("   sudo apt-get install libgl1-mesa-glx libglib2.0-0 libsm6 libxext6 libxrender-dev libgomp1")

        elif self.system == "Darwin":
            print("   macOS系统检测到")
            print("   可能需要安装Xcode Command Line Tools")

    def create_sample_files(self):
        """创建示例文件"""
        print("\n📝 创建示例文件...")

        # 创建示例配置
        sample_config = """# 多模态知识图谱构建系统配置示例

## Neo4j数据库配置
NEO4J_URI=bolt://localhost:7687
NEO4J_USER=neo4j
NEO4J_PASSWORD=your-password

## DeepSeek API配置  
DEEPSEEK_API_KEY=your-api-key
DEEPSEEK_MODEL=deepseek-chat

## 模型路径配置
CLIP_MODEL_PATH=openai/clip-vit-base-patch32
LOCAL_CLIP_PATH=F:/Models/clip-vit-base-patch32

## 处理参数
FAST_MODE=False
BATCH_SIZE=1000
MAX_WORKERS=3
"""

        env_file = self.project_root / ".env.example"
        with open(env_file, 'w', encoding='utf-8') as f:
            f.write(sample_config)

        print(f"   ✅ 环境配置示例: {env_file}")

        # 创建简单的使用示例
        usage_example = '''#!/usr/bin/env python3
"""
多模态知识图谱构建 - 使用示例
"""

from multimodal_kg import build_multimodal_knowledge_graph

# 基本配置
config = {
    'neo4j_uri': "bolt://localhost:7687",
    'neo4j_user': "neo4j", 
    'neo4j_password': "your-password",
    'deepseek_api_key': "your-api-key",
    'input_dir': "input",
    'output_dir': "output",
    'document_name': "我的文档",
    'fast_mode': False,
    'verbose': True
}

# 运行处理
if __name__ == "__main__":
    result = build_multimodal_knowledge_graph(**config)

    if result['success']:
        print("✅ 处理成功！")
        print(f"实体数: {result['entities_saved']}")
        print(f"关系数: {result['relationships_saved']}")
    else:
        print(f"❌ 处理失败: {result['error']}")
'''

        example_file = self.project_root / "example_usage.py"
        with open(example_file, 'w', encoding='utf-8') as f:
            f.write(usage_example)

        print(f"   ✅ 使用示例: {example_file}")

    def run_tests(self):
        """运行基本测试"""
        print("\n🧪 运行基本测试...")

        test_imports = [
            ("torch", "PyTorch"),
            ("transformers", "Transformers"),
            ("PIL", "Pillow"),
            ("pandas", "Pandas"),
            ("neo4j", "Neo4j Driver"),
            ("pptx", "python-pptx"),
            ("fitz", "PyMuPDF"),
            ("cv2", "OpenCV"),
            ("easyocr", "EasyOCR"),
            ("spacy", "spaCy")
        ]

        failed_imports = []

        for module, name in test_imports:
            try:
                __import__(module)
                print(f"   ✅ {name}")
            except ImportError as e:
                print(f"   ❌ {name}: {e}")
                failed_imports.append(name)

        if failed_imports:
            print(f"\n⚠️ 以下包导入失败: {', '.join(failed_imports)}")
            print("   请检查安装或重新运行安装脚本")
            return False

        print("\n✅ 所有核心依赖测试通过！")
        return True

    def install(self):
        """执行完整安装流程"""
        print("🚀 多模态知识图谱构建系统 - 环境安装")
        print("=" * 60)

        try:
            # 1. 检查Python版本
            if not self.check_python_version():
                return False

            # 2. 安装Python包
            self.install_pip_packages()

            # 3. 安装spaCy模型
            self.install_spacy_model()

            # 4. 创建目录结构
            self.setup_directories()

            # 5. 创建配置文件
            self.create_config_template()

            # 6. 检查GPU支持
            self.check_gpu_support()

            # 7. 检查系统依赖
            self.install_system_dependencies()

            # 8. 创建示例文件
            self.create_sample_files()

            # 9. 运行测试
            if not self.run_tests():
                print("\n⚠️ 部分依赖安装可能有问题，请检查上述错误信息")

            print("\n" + "=" * 60)
            print("🎉 环境安装完成！")
            print("\n📋 后续步骤:")
            print("1. 编辑 config/config.json 填入实际配置参数")
            print("2. 将PPT/PDF文件放入 input/ 目录")
            print("3. 运行 python example_usage.py 开始处理")
            print("4. 查看 README.md 了解详细使用说明")
            print("\n💡 提示:")
            print("- 首次运行会下载CLIP模型，请确保网络连接")
            print("- 建议配置本地模型路径以提高加载速度")
            print("- 确保Neo4j数据库正常运行")

            return True

        except Exception as e:
            print(f"\n❌ 安装过程中出现错误: {e}")
            print("请检查网络连接和权限设置")
            return False


if __name__ == "__main__":
    installer = EnvironmentInstaller()
    success = installer.install()

    if success:
        print("\n✅ 安装成功！可以开始使用系统了。")
    else:
        print("\n❌ 安装失败，请查看错误信息并重试。")

    input("\n按回车键退出...")

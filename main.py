"""
Excel Auto Handle - Main Entry Point
Excel自动化处理主程序

通过自然语言指令实现Excel文件的智能处理
"""

import os
import sys
import logging
import shutil
import json
import requests
import pandas as pd
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple
from dotenv import load_dotenv


# 项目根目录
PROJECT_ROOT = Path(__file__).parent

# 子目录路径
UPLOADS_DIR = PROJECT_ROOT / "uploads"
RESULTS_DIR = PROJECT_ROOT / "results"
TEMP_DIR = PROJECT_ROOT / "temp"
CONFIG_DIR = PROJECT_ROOT / "config"
LOGS_DIR = PROJECT_ROOT / "logs"
EXCEL_FUNCTIONS_DIR = PROJECT_ROOT / "excel_functions"


def init_directories():
    """初始化项目目录结构"""
    directories = [UPLOADS_DIR, RESULTS_DIR, TEMP_DIR, CONFIG_DIR, LOGS_DIR]
    
    for directory in directories:
        if not directory.exists():
            directory.mkdir(parents=True, exist_ok=True)
            print(f"✓ 创建目录: {directory}")
        else:
            print(f"✓ 目录已存在: {directory}")


def setup_logging():
    """配置日志系统"""
    log_file = LOGS_DIR / f"excel_auto_handle_{datetime.now().strftime('%Y%m%d')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger(__name__)


def load_config():
    """加载配置文件"""
    env_file = CONFIG_DIR / ".env"
    
    if not env_file.exists():
        print(f"\n⚠️  配置文件不存在: {env_file}")
        print(f"请复制 {CONFIG_DIR}/.env.example 为 {CONFIG_DIR}/.env 并填入API密钥")
        
        # 检查example文件
        example_file = CONFIG_DIR / ".env.example"
        if example_file.exists():
            print(f"\n示例配置文件位置: {example_file}")
        
        return None
    
    # 加载环境变量
    load_dotenv(env_file)
    
    config = {
        'api_key': os.getenv('DEEPSEEK_API_KEY'),
        'api_url': os.getenv('DEEPSEEK_API_URL', 'https://api.deepseek.com/v1/chat/completions'),
        'timeout': int(os.getenv('TIMEOUT', 30))
    }
    
    # 验证配置
    if not config['api_key'] or config['api_key'] == 'your_api_key_here':
        print("\n⚠️  请在配置文件中设置有效的 DEEPSEEK_API_KEY")
        return None
    
    return config


def get_excel_files_info() -> List[Dict]:
    """
    获取uploads目录下所有Excel文件的信息
    
    Returns:
        文件信息列表
    """
    excel_files = []
    supported_extensions = ['.xlsx', '.xls', '.csv']
    
    for file_path in UPLOADS_DIR.iterdir():
        if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
            try:
                # 获取文件基本信息
                file_info = {
                    'filename': file_path.name,
                    'path': str(file_path),
                    'size': f"{file_path.stat().st_size / 1024:.2f} KB",
                    'modified': datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                    'sheets': [],
                    'error': None
                }
                
                # 读取工作表信息
                if file_path.suffix.lower() == '.csv':
                    df = pd.read_csv(file_path, nrows=5)
                    file_info['sheets'] = [{
                        'name': 'CSV',
                        'columns': df.columns.tolist(),
                        'rows': len(pd.read_csv(file_path)),
                        'preview': df.head(5).to_dict('records')
                    }]
                else:
                    excel_file = pd.ExcelFile(file_path)
                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                        file_info['sheets'].append({
                            'name': sheet_name,
                            'columns': df.columns.tolist(),
                            'rows': len(pd.read_excel(file_path, sheet_name=sheet_name)),
                            'preview': df.head(5).to_dict('records')
                        })
                
                excel_files.append(file_info)
                
            except Exception as e:
                excel_files.append({
                    'filename': file_path.name,
                    'path': str(file_path),
                    'error': str(e)
                })
    
    return excel_files


def display_excel_files(files_info: List[Dict]):
    """显示Excel文件列表"""
    if not files_info:
        print("\n📂 uploads 目录为空，请先上传Excel文件（支持 .xlsx/.xls/.csv）")
        return
    
    print("\n" + "=" * 80)
    print("📂 当前 uploads 目录下的Excel文件:")
    print("=" * 80)
    
    for idx, file_info in enumerate(files_info, 1):
        print(f"\n[{idx}] 文件: {file_info['filename']}")
        print(f"    大小: {file_info.get('size', 'N/A')}")
        print(f"    修改时间: {file_info.get('modified', 'N/A')}")
        
        if file_info.get('error'):
            print(f"    ⚠️  读取错误: {file_info['error']}")
            continue
        
        if file_info.get('sheets'):
            for sheet in file_info['sheets']:
                print(f"    工作表: {sheet['name']}")
                print(f"      - 行数: {sheet['rows']}")
                print(f"      - 列数: {len(sheet['columns'])}")
                print(f"      - 表头: {', '.join(sheet['columns'][:10])}" + 
                      ("..." if len(sheet['columns']) > 10 else ""))
    
    print("\n" + "=" * 80)


def get_functions_info() -> str:
    """
    获取excel_functions模块中所有函数的信息
    
    Returns:
        函数信息文本
    """
    import excel_functions
    
    functions_info = []
    
    # 获取所有导出的函数
    for func_name in excel_functions.__all__:
        func = getattr(excel_functions, func_name)
        
        # 获取函数文档字符串
        doc = func.__doc__ if func.__doc__ else "无描述"
        
        # 简化文档字符串
        doc_lines = [line.strip() for line in doc.split('\n') if line.strip()]
        description = doc_lines[0] if doc_lines else "无描述"
        
        functions_info.append(f"- {func_name}: {description}")
    
    return "\n".join(functions_info)


def validate_script_security(script_content: str) -> Tuple[bool, str]:
    """
    验证生成的脚本是否安全
    
    Args:
        script_content: 脚本内容
        
    Returns:
        (是否安全, 错误消息)
    """
    # 危险库和函数列表
    dangerous_patterns = [
        'import os',
        'import subprocess',
        'import sys',
        'import shutil',
        '__import__',
        'eval(',
        'exec(',
        'compile(',
        'open(',  # 限制文件操作
        'rmdir',
        'remove',
        'unlink',
        'delete',
    ]
    
    # 检查危险模式
    for pattern in dangerous_patterns:
        if pattern in script_content.lower():
            return False, f"脚本包含危险代码: {pattern}"
    
    # 检查是否操作uploads目录
    if 'uploads' in script_content and ('remove' in script_content or 'delete' in script_content):
        return False, "脚本尝试删除uploads目录中的文件"
    
    return True, ""


def call_deepseek_api(config: Dict, files_info: List[Dict], functions_info: str, user_instruction: str, logger) -> str:
    """
    调用DeepSeek API生成处理脚本
    
    Args:
        config: 配置信息
        files_info: Excel文件信息
        functions_info: 函数信息
        user_instruction: 用户指令
        logger: 日志记录器
        
    Returns:
        生成的Python脚本
    """
    try:
        # 构建提示词
        files_summary = "\n".join([
            f"文件: {f['filename']}, 工作表: {[s['name'] for s in f.get('sheets', [])]}, "
            f"列: {[s['columns'] for s in f.get('sheets', [])]}"
            for f in files_info if not f.get('error')
        ])
        
        prompt = f"""你是一个Excel处理脚本生成专家。

可用的Excel文件信息:
{files_summary}

可用的Excel处理函数:
{functions_info}

用户指令: {user_instruction}

请生成Python脚本来完成用户的需求。要求:
1. 导入必要的模块: import excel_functions as ef, import pandas as pd, from pathlib import Path
2. 使用提供的excel_functions模块中的函数
3. 文件路径使用: Path("uploads") / "文件名"
4. 结果保存到: Path("results") / "结果文件名.xlsx"
5. 结果文件命名格式: 原文件名_操作描述_时间戳.xlsx
6. 包含错误处理
7. 打印处理过程和结果
8. 只返回Python代码，不要有任何解释文字
9. 代码要完整可执行

示例代码格式:
```python
import excel_functions as ef
import pandas as pd
from pathlib import Path
from datetime import datetime

# 文件路径
input_file = Path("uploads") / "示例文件.xlsx"
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
output_file = Path("results") / f"示例文件_处理结果_{timestamp}.xlsx"

# 执行处理
success, result = ef.excel_deduplicate(str(input_file), "Sheet1", columns=['姓名'], save_path=str(output_file))

if success:
    print(f"处理成功: {result}")
    print(f"结果保存至: {output_file}")
else:
    print(f"处理失败: {result}")
```

请生成代码:"""

        # 调用API
        logger.info(f"正在调用DeepSeek API...")
        
        headers = {
            'Authorization': f'Bearer {config["api_key"]}',
            'Content-Type': 'application/json'
        }
        
        data = {
            'model': 'deepseek-chat',
            'messages': [
                {'role': 'user', 'content': prompt}
            ],
            'temperature': 0.7
        }
        
        response = requests.post(
            config['api_url'],
            headers=headers,
            json=data,
            timeout=config['timeout']
        )
        
        if response.status_code != 200:
            error_msg = f"API调用失败: HTTP {response.status_code} - {response.text}"
            logger.error(error_msg)
            return None
        
        response_data = response.json()
        
        if 'choices' not in response_data or not response_data['choices']:
            logger.error("API返回内容为空")
            return None
        
        script_content = response_data['choices'][0]['message']['content']
        
        # 提取代码块
        if '```python' in script_content:
            script_content = script_content.split('```python')[1].split('```')[0].strip()
        elif '```' in script_content:
            script_content = script_content.split('```')[1].split('```')[0].strip()
        
        logger.info("API调用成功，脚本生成完成")
        return script_content
        
    except requests.exceptions.Timeout:
        logger.error(f"API调用超时（超过 {config['timeout']} 秒）")
        return None
    except requests.exceptions.RequestException as e:
        logger.error(f"网络错误: {str(e)}")
        return None
    except KeyError as e:
        logger.error(f"API密钥错误或配置错误: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"API调用失败: {str(e)}")
        return None


def save_and_execute_script(script_content: str, logger) -> bool:
    """
    保存并执行生成的脚本
    
    Args:
        script_content: 脚本内容
        logger: 日志记录器
        
    Returns:
        是否执行成功
    """
    try:
        # 生成脚本文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        script_name = f"process_{timestamp}.py"
        script_path = TEMP_DIR / script_name
        
        # 保存脚本
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        logger.info(f"脚本已保存: {script_path}")
        print(f"\n📝 生成的脚本已保存至: {script_path}")
        
        # 执行脚本
        print("\n🚀 开始执行脚本...")
        logger.info("开始执行脚本")
        
        # 使用exec执行脚本
        exec_globals = {
            '__name__': '__main__',
            '__file__': str(script_path)
        }
        
        with open(script_path, 'r', encoding='utf-8') as f:
            script_code = f.read()
        
        exec(compile(script_code, str(script_path), 'exec'), exec_globals)
        
        logger.info("脚本执行成功")
        print("\n✅ 脚本执行成功!")
        return True
        
    except SyntaxError as e:
        logger.error(f"脚本语法错误: {str(e)}")
        print(f"\n❌ 脚本语法错误: {str(e)}")
        print(f"   错误位置: 第 {e.lineno} 行")
        return False
    except Exception as e:
        logger.error(f"脚本执行失败: {str(e)}")
        print(f"\n❌ 脚本执行失败: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return False


def cleanup_temp_files():
    """清理临时文件"""
    try:
        for temp_file in TEMP_DIR.glob("*.py"):
            temp_file.unlink()
        print("\n🧹 临时文件已清理")
    except Exception as e:
        print(f"\n⚠️  清理临时文件失败: {str(e)}")


def get_user_instruction() -> str:
    """
    获取用户指令
    
    Returns:
        用户指令文本
    """
    print("\n" + "=" * 80)
    print("💬 请输入您的处理指令:")
    print("=" * 80)
    print("示例指令:")
    print("  - 对 test.xlsx 的 Sheet1 按姓名列去重")
    print("  - 计算 sales.xlsx 的销售额列的总和")
    print("  - 合并 file1.xlsx 和 file2.xlsx")
    print("  - 筛选 data.xlsx 中年龄大于30的数据")
    print("\n输入 'exit' 或 'quit' 退出程序")
    print("输入 'list' 重新显示文件列表")
    print("-" * 80)
    
    instruction = input("\n>>> ").strip()
    return instruction


def main():
    """主程序入口"""
    print("\n" + "=" * 80)
    print("  Excel Auto Handle - 智能Excel处理系统")
    print("  基于 Python + DeepSeek API")
    print("=" * 80)
    
    # 1. 初始化目录
    print("\n[1/6] 初始化目录结构...")
    init_directories()
    
    # 2. 配置日志
    print("\n[2/6] 配置日志系统...")
    logger = setup_logging()
    logger.info("程序启动")
    
    # 3. 加载配置
    print("\n[3/6] 加载配置文件...")
    config = load_config()
    if not config:
        print("\n❌ 配置加载失败，程序退出")
        return
    
    print("✓ 配置加载成功")
    
    # 4. 获取Excel文件信息
    print("\n[4/6] 扫描Excel文件...")
    files_info = get_excel_files_info()
    display_excel_files(files_info)
    
    # 5. 获取函数信息
    print("\n[5/6] 加载Excel处理函数...")
    functions_info = get_functions_info()
    print(f"✓ 已加载 {len(functions_info.split(chr(10)))} 个处理函数")
    
    # 6. 进入交互循环
    print("\n[6/6] 进入交互模式")
    
    try:
        while True:
            # 获取用户指令
            instruction = get_user_instruction()
            
            # 处理特殊命令
            if instruction.lower() in ['exit', 'quit']:
                print("\n👋 感谢使用，再见!")
                break
            
            if instruction.lower() == 'list':
                files_info = get_excel_files_info()
                display_excel_files(files_info)
                continue
            
            if not instruction:
                print("⚠️  指令不能为空，请重新输入")
                continue
            
            # 检查文件信息
            if not files_info or all(f.get('error') for f in files_info):
                print("\n⚠️  没有可用的Excel文件，请先将文件放入 uploads 目录")
                continue
            
            # 记录用户指令
            logger.info(f"用户指令: {instruction}")
            
            # 调用API生成脚本
            print("\n🤖 正在生成处理脚本...")
            script_content = call_deepseek_api(config, files_info, functions_info, instruction, logger)
            
            if not script_content:
                print("\n❌ 脚本生成失败，请检查API配置或重试")
                continue
            
            # 安全验证
            print("\n🔒 进行安全检查...")
            is_safe, error_msg = validate_script_security(script_content)
            
            if not is_safe:
                logger.warning(f"安全验证失败: {error_msg}")
                print(f"\n⚠️  安全验证失败: {error_msg}")
                print("指令生成的脚本存在安全风险，请重新输入指令")
                continue
            
            print("✓ 安全检查通过")
            
            # 执行脚本
            success = save_and_execute_script(script_content, logger)
            
            if success:
                print("\n✅ 处理完成! 结果文件已保存到 results 目录")
            else:
                print("\n❌ 处理失败，请检查错误信息并重试")
            
            print("\n" + "-" * 80)
    
    except KeyboardInterrupt:
        print("\n\n⚠️  程序被用户中断")
    finally:
        # 清理临时文件
        cleanup_temp_files()
        logger.info("程序退出")


if __name__ == "__main__":
    main()

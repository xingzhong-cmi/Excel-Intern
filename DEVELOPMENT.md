# 开发指南 (Development Guide)

本文档面向想要扩展或修改 Excel Auto Handle 项目的开发者。

## 项目架构

### 核心模块

```
excel_auto_handle/
├── excel_functions/    # Excel处理函数库（可扩展）
│   ├── crud.py        # 增删改操作
│   ├── query.py       # 查询筛选
│   ├── statistics.py  # 统计计算
│   ├── merge.py       # 文件合并
│   └── __init__.py    # 统一导出
└── main.py            # 主程序（包含API调用和脚本执行）
```

### 数据流

```
用户输入指令
    ↓
文件信息收集 (get_excel_files_info)
    ↓
函数信息收集 (get_functions_info)
    ↓
构建API提示词
    ↓
调用DeepSeek API (call_deepseek_api)
    ↓
接收生成的脚本
    ↓
安全验证 (validate_script_security)
    ↓
保存并执行脚本 (save_and_execute_script)
    ↓
输出结果到 results/
```

## 添加新的Excel函数

### 1. 函数规范

所有Excel处理函数应遵循以下规范：

```python
def excel_function_name(
    file_path: str,           # Excel文件路径（必需）
    sheet_name: str,          # 工作表名（必需）
    # ... 其他参数
    save_path: str = None    # 保存路径（可选，默认None覆盖原文件）
) -> Tuple[bool, Union[Any, str]]:
    """
    函数功能描述
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        # ... 其他参数说明
        save_path: 保存路径
        
    Returns:
        (成功/失败, 结果或错误消息)
    """
    try:
        # 1. 参数验证
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        # 2. 读取Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 3. 执行操作
        # ... 你的处理逻辑
        
        # 4. 保存结果（如需要）
        if save_path:
            df.to_excel(save_path, sheet_name=sheet_name, index=False)
        
        # 5. 返回结果
        return True, "成功消息或结果数据"
        
    except ValueError as e:
        return False, f"工作表名错误: {str(e)}"
    except Exception as e:
        return False, f"操作失败: {str(e)}"
```

### 2. 添加步骤

#### 步骤1: 创建函数

在适当的模块中添加函数（或创建新模块）：

```python
# excel_functions/formatting.py (新模块示例)

def excel_format_currency(
    file_path: str, 
    sheet_name: str, 
    column_name: str,
    save_path: str = None
) -> Tuple[bool, str]:
    """
    将列格式化为货币格式
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名
        column_name: 要格式化的列名
        save_path: 保存路径
        
    Returns:
        (成功/失败, 消息)
    """
    try:
        import openpyxl
        from openpyxl.styles import numbers
        
        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        # 使用openpyxl处理格式
        wb = openpyxl.load_workbook(file_path)
        ws = wb[sheet_name]
        
        # 找到列索引
        header = [cell.value for cell in ws[1]]
        if column_name not in header:
            return False, f"列 '{column_name}' 不存在"
        
        col_idx = header.index(column_name) + 1
        
        # 应用货币格式
        for row in range(2, ws.max_row + 1):
            ws.cell(row, col_idx).number_format = numbers.FORMAT_CURRENCY_USD
        
        # 保存
        output_path = save_path if save_path else file_path
        wb.save(output_path)
        
        return True, f"成功格式化列 '{column_name}' 为货币格式"
        
    except KeyError:
        return False, f"工作表 '{sheet_name}' 不存在"
    except Exception as e:
        return False, f"格式化失败: {str(e)}"
```

#### 步骤2: 导出函数

在 `excel_functions/__init__.py` 中添加导入：

```python
# 添加新模块的导入
from .formatting import (
    excel_format_currency,
    # ... 其他格式化函数
)

# 添加到 __all__ 列表
__all__ = [
    # ... 现有函数
    'excel_format_currency',
    # ... 其他新函数
]
```

#### 步骤3: 测试函数

创建测试脚本：

```python
# tests/test_formatting.py
import excel_functions as ef

def test_format_currency():
    success, msg = ef.excel_format_currency(
        'uploads/test.xlsx',
        'Sheet1',
        '金额',
        save_path='results/test_formatted.xlsx'
    )
    
    if success:
        print(f"✓ 测试通过: {msg}")
    else:
        print(f"✗ 测试失败: {msg}")

if __name__ == '__main__':
    test_format_currency()
```

## 修改API提示词

API提示词在 `main.py` 的 `call_deepseek_api` 函数中定义。

### 自定义提示词

```python
def call_deepseek_api(config, files_info, functions_info, user_instruction, logger):
    # ... 前面的代码
    
    # 自定义提示词
    prompt = f"""你是一个Excel处理脚本生成专家。

[添加你的自定义说明]

可用的Excel文件信息:
{files_summary}

可用的Excel处理函数:
{functions_info}

用户指令: {user_instruction}

[添加你的特殊要求]

请生成代码:"""
    
    # ... 后面的代码
```

### 提示词优化建议

1. **明确角色定位**: 告诉AI它是什么角色
2. **提供完整上下文**: 包含文件信息、函数信息
3. **设置输出格式**: 明确要求代码格式
4. **添加约束条件**: 安全性、错误处理等要求
5. **提供示例**: 给出一个标准的代码示例

## 安全机制扩展

### 添加新的安全检查

在 `validate_script_security` 函数中添加检查规则：

```python
def validate_script_security(script_content: str) -> Tuple[bool, str]:
    """验证脚本安全性"""
    
    # 现有的危险模式
    dangerous_patterns = [
        'import os',
        'import subprocess',
        # ... 现有模式
    ]
    
    # 添加新的危险模式
    dangerous_patterns.extend([
        'socket',           # 网络操作
        '__builtins__',    # 访问内置对象
        'pickle',          # 反序列化
        # 添加更多...
    ])
    
    # 白名单检查（可选）
    required_imports = ['excel_functions', 'pandas', 'Path']
    for imp in required_imports:
        if imp not in script_content:
            # 可以添加警告但不阻止
            pass
    
    # ... 其他检查逻辑
```

## 日志扩展

### 添加自定义日志级别

```python
# 在 main.py 中
import logging

# 添加自定义日志级别
CUSTOM_LEVEL = 25  # 介于INFO(20)和WARNING(30)之间
logging.addLevelName(CUSTOM_LEVEL, "CUSTOM")

def custom_log(self, message, *args, **kwargs):
    if self.isEnabledFor(CUSTOM_LEVEL):
        self._log(CUSTOM_LEVEL, message, args, **kwargs)

logging.Logger.custom = custom_log

# 使用
logger.custom("自定义级别的日志消息")
```

### 添加结构化日志

```python
import json
import logging

class StructuredLogger:
    def __init__(self, logger):
        self.logger = logger
    
    def log_operation(self, operation, file, status, details=None):
        log_data = {
            'operation': operation,
            'file': file,
            'status': status,
            'timestamp': datetime.now().isoformat(),
            'details': details or {}
        }
        self.logger.info(json.dumps(log_data, ensure_ascii=False))

# 使用
struct_logger = StructuredLogger(logger)
struct_logger.log_operation(
    operation='deduplicate',
    file='test.xlsx',
    status='success',
    details={'rows_removed': 10}
)
```

## 错误处理最佳实践

### 1. 分层错误处理

```python
def excel_operation(file_path, sheet_name):
    try:
        # 主要逻辑
        result = perform_operation()
        return True, result
    
    except FileNotFoundError:
        # 特定错误
        return False, f"文件不存在: {file_path}"
    
    except ValueError as e:
        # 数据错误
        return False, f"数据格式错误: {str(e)}"
    
    except Exception as e:
        # 通用错误
        logger.error(f"未预期的错误: {str(e)}", exc_info=True)
        return False, f"操作失败: {str(e)}"
```

### 2. 用户友好的错误消息

```python
# ❌ 不好的错误消息
return False, str(e)

# ✅ 好的错误消息
return False, f"无法读取文件 {file_path}: 文件可能被其他程序占用或已损坏"
```

## 性能优化

### 1. 大文件处理

```python
def excel_process_large_file(file_path, sheet_name, chunk_size=10000):
    """分块处理大文件"""
    chunks = []
    for chunk in pd.read_excel(file_path, sheet_name=sheet_name, chunksize=chunk_size):
        # 处理每个块
        processed_chunk = process_chunk(chunk)
        chunks.append(processed_chunk)
    
    # 合并结果
    result = pd.concat(chunks, ignore_index=True)
    return result
```

### 2. 缓存优化

```python
from functools import lru_cache

@lru_cache(maxsize=10)
def get_excel_file_info_cached(file_path):
    """缓存文件信息，避免重复读取"""
    return get_excel_file_info(file_path)
```

## 测试指南

### 单元测试

创建 `tests/` 目录：

```python
# tests/test_crud.py
import unittest
import pandas as pd
import excel_functions as ef
from pathlib import Path

class TestCRUDOperations(unittest.TestCase):
    
    @classmethod
    def setUpClass(cls):
        """创建测试文件"""
        cls.test_file = 'uploads/test_crud.xlsx'
        data = {'A': [1, 2, 3], 'B': [4, 5, 6]}
        df = pd.DataFrame(data)
        df.to_excel(cls.test_file, index=False)
    
    def test_add_row(self):
        """测试添加行"""
        success, msg = ef.excel_add_row(
            self.test_file,
            'Sheet1',
            {'A': 4, 'B': 7}
        )
        self.assertTrue(success)
    
    def test_delete_row(self):
        """测试删除行"""
        success, msg = ef.excel_delete_row(
            self.test_file,
            'Sheet1',
            row_indices=[0]
        )
        self.assertTrue(success)
    
    @classmethod
    def tearDownClass(cls):
        """清理测试文件"""
        Path(cls.test_file).unlink(missing_ok=True)

if __name__ == '__main__':
    unittest.main()
```

### 集成测试

```python
# tests/test_integration.py
def test_full_workflow():
    """测试完整工作流"""
    from main import (
        init_directories,
        get_excel_files_info,
        save_and_execute_script
    )
    
    # 1. 初始化
    init_directories()
    
    # 2. 获取文件
    files = get_excel_files_info()
    assert len(files) > 0
    
    # 3. 执行脚本
    test_script = "import excel_functions as ef\nprint('Test')"
    success = save_and_execute_script(test_script, logger)
    assert success
```

## 贡献指南

### 提交代码前检查清单

- [ ] 代码遵循项目的编码规范
- [ ] 添加了必要的文档字符串
- [ ] 添加了单元测试
- [ ] 所有测试通过
- [ ] 更新了README（如需要）
- [ ] 更新了EXAMPLES（如需要）

### 代码风格

遵循 PEP 8 规范：

```bash
# 安装工具
pip install black flake8

# 格式化代码
black excel_functions/

# 检查代码
flake8 excel_functions/
```

### 提交信息格式

```
类型: 简短描述

详细描述（可选）

示例:
- feat: 添加货币格式化函数
- fix: 修复大文件处理bug
- docs: 更新README文档
- test: 添加合并操作测试
- refactor: 重构API调用逻辑
```

## 常见开发问题

### Q: 如何调试生成的脚本？

A: 查看 `temp/` 目录中的脚本文件，手动执行并观察输出。

### Q: 如何处理中文编码问题？

A: 始终使用 `encoding='utf-8'`：

```python
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()
```

### Q: 如何支持新的文件格式？

A: 在 `get_excel_files_info` 函数中添加支持：

```python
supported_extensions = ['.xlsx', '.xls', '.csv', '.xlsm']  # 添加 .xlsm
```

## 资源链接

- [pandas 文档](https://pandas.pydata.org/docs/)
- [openpyxl 文档](https://openpyxl.readthedocs.io/)
- [DeepSeek API 文档](https://www.deepseek.com/docs)

## 联系方式

如有开发相关问题，请通过 GitHub Issues 讨论。

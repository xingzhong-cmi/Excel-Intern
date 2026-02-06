# Excel Auto Handle - 智能Excel处理系统

基于 Python + DeepSeek API 的智能 Excel 处理项目，支持用户通过自然语言文字指令实现 Excel 文件的增删改查、数据计算、格式调整等操作，全程自动化生成并执行处理脚本，无需用户编写代码。

## 📋 目录

- [功能特性](#功能特性)
- [技术栈](#技术栈)
- [项目结构](#项目结构)
- [快速开始](#快速开始)
- [使用方法](#使用方法)
- [基础函数说明](#基础函数说明)
- [常见问题](#常见问题)

## ✨ 功能特性

- **自然语言交互**: 使用中文自然语言指令操作 Excel，无需编写代码
- **智能脚本生成**: 基于 DeepSeek API 自动生成 Python 处理脚本
- **全面的Excel操作**: 支持增删改查、数据统计、文件合并等多种操作
- **安全验证机制**: 自动检测并阻止危险脚本执行
- **详细日志记录**: 记录所有操作过程，便于追踪和调试
- **多文件格式支持**: 支持 .xlsx、.xls、.csv 三种文件格式

## 🛠 技术栈

- **Python 3.8+**
- **pandas 2.2.0** - Excel 数据处理
- **openpyxl 3.1.2** - xlsx 文件读写
- **xlrd 2.0.1** - xls 文件兼容
- **requests 2.31.0** - API 调用
- **python-dotenv 1.0.0** - 环境变量管理

## 📁 项目结构

```
excel_auto_handle/
├── uploads/              # 用户上传的Excel文件目录
├── excel_functions/      # Excel处理基础函数模块
│   ├── __init__.py      # 函数统一导出
│   ├── crud.py          # 增删改操作
│   ├── query.py         # 查询筛选操作
│   ├── statistics.py    # 统计计算操作
│   └── merge.py         # 文件合并操作
├── results/             # 处理结果输出目录
├── config/              # 配置文件目录
│   ├── .env.example     # 环境变量配置示例
│   └── .env            # 实际配置文件（需自行创建）
├── temp/                # 临时脚本存储目录
├── logs/                # 日志文件目录
├── main.py              # 主程序入口
├── requirements.txt     # 项目依赖
└── README.md           # 项目说明文档
```

## 🚀 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/xingzhong-cmi/Excel-Intern.git
cd Excel-Intern
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置API密钥

复制配置示例文件并编辑：

```bash
cp config/.env.example config/.env
```

编辑 `config/.env` 文件，填入您的 DeepSeek API Key：

```env
DEEPSEEK_API_KEY=your_actual_api_key_here
DEEPSEEK_API_URL=https://api.deepseek.com/v1/chat/completions
TIMEOUT=30
```

### 4. 准备Excel文件

将需要处理的 Excel 文件放入 `uploads/` 目录。支持的格式：
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)
- `.csv` (逗号分隔值)

### 5. 运行程序

```bash
python main.py
```

## 📖 使用方法

### 启动程序

运行程序后，系统会：
1. 自动初始化所需目录
2. 加载配置和日志系统
3. 扫描并显示 uploads 目录中的 Excel 文件信息
4. 进入交互模式，等待用户输入指令

### 指令示例

程序支持自然语言指令，以下是一些示例：

#### 数据去重
```
对 test.xlsx 的 Sheet1 按姓名列去重
```

#### 数据统计
```
计算 sales.xlsx 的销售额列的总和
统计 data.xlsx 中年龄列的平均值
```

#### 数据筛选
```
筛选 data.xlsx 中年龄大于30的数据
从 orders.xlsx 中查找金额大于1000的订单
```

#### 文件合并
```
合并 file1.xlsx 和 file2.xlsx
将 sales_q1.xlsx 和 sales_q2.xlsx 纵向合并
```

#### 数据修改
```
删除 data.xlsx 中的空行
在 test.xlsx 的 Sheet1 中添加新列"备注"
```

### 特殊命令

- `list` - 重新显示文件列表
- `exit` 或 `quit` - 退出程序

### 查看结果

处理完成后：
- 结果文件保存在 `results/` 目录
- 文件命名格式：`原文件名_操作描述_时间戳.xlsx`
- 临时脚本保存在 `temp/` 目录（程序退出时自动清理）
- 操作日志保存在 `logs/` 目录

## 🔧 基础函数说明

### CRUD 操作 (crud.py)

| 函数名 | 功能说明 |
|--------|----------|
| `excel_add_row` | 向Excel添加新行 |
| `excel_add_column` | 向Excel添加新列 |
| `excel_delete_row` | 删除指定行 |
| `excel_delete_column` | 删除指定列 |
| `excel_delete_empty_rows` | 删除空行 |
| `excel_modify_cell` | 修改单元格值 |
| `excel_modify_column` | 批量修改列值 |

### 查询操作 (query.py)

| 函数名 | 功能说明 |
|--------|----------|
| `excel_query_data` | 按条件查询数据 |
| `excel_filter_by_value` | 按值筛选数据 |
| `excel_search_text` | 文本搜索 |
| `excel_get_unique_values` | 获取唯一值 |
| `excel_filter_by_range` | 按数值范围筛选 |

### 统计操作 (statistics.py)

| 函数名 | 功能说明 |
|--------|----------|
| `excel_sum_column` | 列求和 |
| `excel_average_column` | 列平均值 |
| `excel_count_values` | 计数 |
| `excel_max_value` | 最大值 |
| `excel_min_value` | 最小值 |
| `excel_deduplicate` | 数据去重 |
| `excel_group_statistics` | 分组统计 |
| `excel_calculate_statistics` | 综合统计 |

### 合并操作 (merge.py)

| 函数名 | 功能说明 |
|--------|----------|
| `excel_merge_files` | 合并多个文件 |
| `excel_merge_sheets` | 合并多个工作表 |
| `excel_join_files` | 按列关联合并 |
| `excel_append_data` | 追加数据 |

## ❓ 常见问题

### 1. 配置文件相关

**Q: 提示"配置文件不存在"怎么办？**

A: 需要创建配置文件：
```bash
cp config/.env.example config/.env
```
然后编辑 `.env` 文件，填入有效的 DeepSeek API Key。

**Q: 如何获取 DeepSeek API Key？**

A: 访问 [DeepSeek 官网](https://www.deepseek.com) 注册账号并获取 API Key。

### 2. 文件处理相关

**Q: 支持哪些文件格式？**

A: 支持 `.xlsx`、`.xls`、`.csv` 三种格式。其他格式会被自动跳过。

**Q: 为什么文件列表为空？**

A: 请确保：
- 文件已放入 `uploads/` 目录
- 文件格式正确（.xlsx/.xls/.csv）
- 文件名不要包含特殊字符

**Q: 处理后的文件在哪里？**

A: 结果文件保存在 `results/` 目录，命名格式为：`原文件名_操作描述_时间戳.xlsx`

### 3. 错误处理

**Q: 提示"脚本存在安全风险"怎么办？**

A: 系统检测到生成的脚本可能执行危险操作。请：
- 重新组织指令，使其更明确
- 避免包含删除、修改系统文件等敏感操作
- 如确认安全，可查看 `temp/` 目录中的脚本内容

**Q: API 调用失败怎么办？**

A: 检查：
- API Key 是否正确
- 网络连接是否正常
- API 配额是否充足
- 查看 `logs/` 目录中的日志文件获取详细错误信息

**Q: 脚本执行失败怎么办？**

A: 可能的原因：
- 指令不够明确，AI 理解有偏差
- Excel 文件格式或内容有问题
- 指定的列名或工作表名不存在

解决方法：
- 重新组织指令，提供更多上下文
- 检查 Excel 文件是否正常
- 查看日志文件了解具体错误

### 4. 性能相关

**Q: 处理大文件很慢怎么办？**

A: 
- 大文件处理需要较长时间，请耐心等待
- 可以考虑将大文件拆分成小文件分批处理
- 增加 `config/.env` 中的 TIMEOUT 值

**Q: 如何查看处理进度？**

A: 
- 观察控制台输出的实时信息
- 查看 `logs/` 目录中的日志文件

### 5. 扩展开发

**Q: 如何添加自定义函数？**

A: 
1. 在 `excel_functions/` 目录中添加新的函数
2. 在 `excel_functions/__init__.py` 中导出新函数
3. 函数需遵循统一的接口规范（见现有函数示例）

**Q: 如何修改AI提示词？**

A: 编辑 `main.py` 中 `call_deepseek_api` 函数的 `prompt` 变量。

## 📝 日志说明

日志文件位于 `logs/` 目录，命名格式：`excel_auto_handle_YYYYMMDD.log`

日志内容包括：
- 程序启动/退出时间
- 用户输入的指令
- API 调用情况
- 脚本生成和执行结果
- 错误信息和堆栈跟踪

## 🔒 安全说明

系统内置安全验证机制，会自动阻止以下操作：
- 导入危险系统模块（os、subprocess、sys等）
- 执行系统命令
- 删除或修改 uploads 目录中的原始文件
- 使用 eval/exec 等危险函数

## 📄 许可证

本项目采用 MIT 许可证。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📧 联系方式

如有问题或建议，请通过 GitHub Issues 联系。
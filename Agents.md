# Beta_04 财务分析工具 - Agent 使用指南

> **项目路径**: `/Users/qin.cui/Project/fr_beta04/financial_analysis_agent`
> **主脚本**: `generate_beta04_v22.py` (V2.2推荐)
> **版本**: v2.2 (动态科目识别已完成)
> **状态**: 生产就绪，支持多种数据格式

---

## 📋 项目概述

这是一个自动化财务分析工具，用于从标准化的财务报表（资产负债表、损益表、收入成本表）中提取数据，并生成结构化的 Beta_04 财务分析表。

### 核心功能

1. **数据提取**：从源 Excel 文件中提取多年度财务数据
2. **智能分析**：自动计算盈利结构、经营指标、ROE分解、估值指标等
3. **公式保留**：分析部分使用 Excel 公式，保持逻辑可追溯
4. **数据独立**：提取数据保存为数值，避免外部引用断链
5. **科目识别**：通过科目名称动态查找数据位置（新增）

### 适用场景

- 财务报表分析自动化
- 多年度财务数据对比
- 投资分析和尽职调查
- 财务健康度评估

---

## 🚀 快速开始

### 环境准备

**Python 环境要求**：
- Python 3.7+
- 依赖包：openpyxl

**创建虚拟环境（必须）**：
```bash
# 在项目目录下创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate  # macOS/Linux
# 或
venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt
```

**重要说明**：
- ✅ 虚拟环境仅限当前项目目录，不影响系统环境
- ✅ 每次使用前需要激活虚拟环境
- ✅ venv/ 目录已添加到 .gitignore，不会提交到版本控制

### 基本用法

**确保已激活虚拟环境**：
```bash
source venv/bin/activate  # macOS/Linux
```

**运行程序**：
```bash
python3 generate_beta04_v22.py <输入文件.xlsx> <输出文件.xlsx>
```

### 示例

```bash
# 分析福耀玻璃的财务数据
python3 generate_beta04_v22.py examples/福耀玻璃.xlsx 福耀玻璃_Beta04分析.xlsx

# 使用默认文件名
python3 generate_beta04_v22.py 公司财报.xlsx
```

### 输入要求

输入 Excel 文件必须包含以下工作表：
1. **balance** - 资产负债表数据
2. **损益现金流** - 损益表和现金流量表数据
3. **收入成本** - 收入和成本明细数据

**重要**：系统使用动态科目识别，支持多种科目名称格式。详见 JSON 配置文件说明。

每个工作表的第一行应包含年份标识（如 2024, 2023, 2022...）

### 输出结果

生成的 Beta_04 分析表包含以下部分：
- 资产负债表数据（17项）
- 损益表数据（19项）
- 盈利结构分析（10项）
- 经营指标（10项）
- 资产负债结构（8项）
- 同比增长分析（4项）
- ROE分解（4项）
- 估值指标（3项）

---

## 📁 项目结构

```
financial_analysis_agent/
├── Agents.md                          # 本文件 - Agent 使用指南
├── README.md                          # 用户文档
├── generate_beta04_v22.py             # 主脚本（V2.2推荐）
├── subject_mapping.py                 # 动态科目映射工具
├── subject_mapping_config.json        # JSON配置：科目映射规则
├── subject_validator_v22.py           # V2.2科目完整性验证器
├── requirements.txt                   # Python依赖包列表
├── .gitignore                         # Git忽略文件配置
├── venv/                              # Python虚拟环境（不提交到Git）
└── examples/                          # 示例文件
    ├── 福耀玻璃.xlsx                  # 输入示例
    ├── 福耀玻璃_Beta04_完整分析.xlsx  # 输出示例
    └── 金山办公.xlsx                  # 测试示例
```

---

## 🎯 核心特性

### 1. 数据与公式分离

**设计理念**：
- **数据提取部分**：保存为数值，避免外部引用断链
- **分析计算部分**：使用公式，保持逻辑可追溯

**优势**：
- ✅ 即使源文件被删除，分析表仍可正常使用
- ✅ 用户可点击单元格查看计算逻辑
- ✅ 修改源数据后，分析结果自动更新
- ✅ 便于审计和验证

### 2. 智能公式生成

系统自动生成 Excel 公式，例如：

```excel
毛利率 = (主营收入 - 主营成本) / 主营收入
ROE = 业务利润 / 权益
EPS = 业务利润 / 股本
营收YoY = (本年营收 - 上年营收) / 上年营收
```

### 3. 多年度数据支持

- 自动识别年份（从列标题）
- 支持任意年度数量（通常 5-10 年）
- 自动计算同比增长率（YoY）

---

## 📊 分析指标详解

### 盈利结构分析
- 业务利润、净利润、利润留存
- 现金留存（自由现金流 FCF）

### 经营指标
- 毛利率、费率、净利率
- 应收周转、存货周转
- 预收占比

### 资产负债结构
- 现金资产、股权资产、经营资产占比
- 经营负债、有息负债占比
- 资本投入、利润留存占比

### 同比增长分析（YoY）
- 营收 YoY、净利润 YoY
- 资产 YoY、权益 YoY

### ROE 分解（杜邦分析）
- ROE = 净利率 × 资产周转率 × 权益乘数

### 估值指标
- 每股股息（DPS）
- 每股收益（EPS）
- 每股自由现金流（FPS）

---

## 🔧 技术实现

### 核心函数

```python
# 数据提取函数
get_balance_value(row, col_idx)   # 读取资产负债表数据
get_income_value(row, col_idx)    # 读取损益表数据
get_revenue_value(row, col_idx)   # 读取收入成本表数据

# 数据写入函数
write_data_row(label, values, is_percentage=False, decimals=0,
               is_input=False, color=None, formulas=None)
```

### 行号映射系统

使用 `row_map` 字典记录每个数据项的行号，确保公式引用准确：

```python
row_map['主营收入'] = write_data_row('主营收入', [...])
row_map['主营成本'] = write_data_row('主营成本', [...])

# 后续公式可直接引用
公式 = f"={col}{row_map['主营收入']}-{col}{row_map['主营成本']}"
```

### 公式生成示例

```python
# 生成毛利率公式
毛利率_formulas = []
for i in range(len(years)):
    col = get_column_letter(3 + i)
    毛利率_formulas.append(
        f"=({col}{row_map['主营收入']}-{col}{row_map['主营成本']})/{col}{row_map['主营收入']}"
    )

write_data_row('毛利率', values, is_percentage=True, formulas=毛利率_formulas)
```

---

## 🐛 已修复的问题

### 问题1：A列标签显示@符号
- **原因**：标签以"="开头被 Excel 识别为无效公式
- **修复**：移除标签开头的"="符号
- **状态**：✅ 已修复

### 问题2：多年度数据列错位
- **原因**：数据提取函数固定读取 B 列，未按年份读取对应列
- **修复**：添加 `col_idx` 参数，按年份读取正确列
- **状态**：✅ 已修复

### 问题3：YoY公式位置错误
- **原因**：公式数组在开头添加 None，导致所有公式向右偏移一列
- **修复**：将 None 放在数组末尾，确保 YoY 显示在当年列下
- **状态**：✅ 已修复

**验证结果**：
```
2024年 (C列): =(C23-D23)/D23  ✓ 正确
2023年 (D列): =(D23-E23)/E23  ✓ 正确
2018年 (I列): None            ✓ 正确（最早年份）
```

### 问题4：数据提取列索引错误
- **现象**：主营收入等数据读取错误（应为 38,710,428,679，实际读取 32,650,176,898）
- **原因**：balance表年份从第3列开始，但损益现金流表和收入成本表从第2列开始，使用统一列索引导致读取错位
- **修复**：为三个表创建独立的列索引映射（balance_cols_idx、income_cols_idx、revenue_cols_idx）
- **状态**：✅ 已修复

### 问题5：盈利结构分析混用公式和数值
- **现象**：盈利结构分析部分既有公式又有数值，业务利润公式包含硬编码数值
- **原因**：
  - 股权收益等数据没有放在"损益表数据"部分
  - 业务利润计算时使用硬编码数值而非引用
  - 数据和分析没有完全分离
- **修复**：
  - 将所有原始数据（损耗、其他收益、现金收益、利息支出、财务费用、所得税、股权收益、其他业务收入/成本、营业外收入/成本）放到"损益表数据"部分
  - 盈利结构分析部分全部使用公式引用，无硬编码数值
  - 业务利润公式：`=C23-C24-C25+C31+C32+C33-C34-C35-C36`（纯引用）
- **状态**：✅ 已修复

**最终验证**：
```
✓ 主营收入 (2024): 38,710,428,679（正确）
✓ 损益表数据部分：16个数据项全部为数值
✓ 盈利结构分析部分：10个计算行全部为公式
✓ 业务利润公式：纯引用，无硬编码数值
✓ 所有引用正确指向对应的数据行
```

---

## 🆕 动态科目识别系统（V2.2已完成）

### 背景

V2.1版本使用硬编码行号提取数据，存在通用性问题：
- 不同公司的数据结构不同
- 科目在不同行号，导致数据提取错误
- 需要为每个公司修改代码

### 解决方案

**通过科目名称动态查找数据位置**，而不是固定行号。

### 核心工具

#### 1. JSON配置文件（subject_mapping_config.json）
统一管理所有科目映射规则：

**配置结构**：
```json
{
  "meta": {                           // 元信息
    "version": "2.2.0",
    "description": "Beta_04 财务分析系统科目映射配置"
  },
  "balance_mapping": {                // 资产负债表科目映射
    "subjects": {
      "资产": {
        "aliases": ["资产", "资产总计", "总资产"],
        "description": "资产合计"
      }
    }
  },
  "composite_mapping": {              // 复合科目（需汇总多个科目）
    "subjects": {
      "折旧摊销": {
        "components": ["固定资产折旧", "使用权资产折旧", "无形资产摊销"],
        "description": "折旧摊销合计"
      }
    }
  }
}
```

**配置管理命令**：
```bash
# 查看配置信息
python3 subject_mapping.py --config-info

# 添加科目别名
python3 subject_mapping.py --add-alias balance 货币资金 银行存款

# 分析Excel文件科目映射
python3 subject_mapping.py 你的公司.xlsx
```

#### 2. 科目映射工具（subject_mapping.py）
提供科目名称到行号的动态映射功能：

```python
from subject_mapping import get_value_by_label, BALANCE_MAPPING

# 动态查找"资产"科目并获取数值
资产值 = get_value_by_label(balance, BALANCE_MAPPING, '资产', col_idx)
```

**功能特性**：
- ✅ 精确匹配：优先匹配标准名称
- ✅ 别名匹配：支持多种科目名称表述
- ✅ 复合科目：自动汇总多个科目（如折旧摊销）
- ✅ 调试报告：生成科目映射报告
- ✅ JSON配置：从配置文件加载映射规则

#### 3. 科目验证器（subject_validator_v22.py）
验证核心科目的完整性：

```python
from subject_validator_v22 import validate_subjects

is_valid, missing_subjects = validate_subjects(workbook)
```

**验证策略**：
- ✅ 严格验证策略（验证所有42个beta04_field_mapping必需字段）
- ✅ 详细缺失报告和完成度统计
- ✅ 防止生成不完整的分析报告

#### 4. 测试验证

```bash
# 测试科目映射
python3 subject_mapping.py examples/福耀玻璃.xlsx
python3 subject_mapping.py examples/金山办公.xlsx

# 验证核心科目完整性
python3 subject_validator_v22.py examples/福耀玻璃.xlsx
```

**测试结果**：
- ✅ 福耀玻璃：成功识别所有核心科目，生成完整分析
- ✅ 金山办公：成功识别所有核心科目，生成完整分析
- ✅ 核心科目（资产、负债、权益、营收、成本）100%识别

### 优势

1. **灵活性**：科目可以在任意行，系统自动查找
2. **通用性**：支持不同公司的科目名称表述
3. **可维护性**：科目映射集中在JSON配置文件，易于扩展
4. **智能化**：支持别名、复合科目、配置管理

### 使用方法

**步骤1：准备数据**
- 确保Excel文件包含三个工作表（balance、损益现金流、收入成本）
- 使用标准或常见的科目名称（系统支持多种别名）

**步骤2：验证科目映射（可选）**
```bash
# 查看科目映射情况
python3 subject_mapping.py 你的公司.xlsx

# 验证核心科目完整性
python3 subject_validator_v22.py 你的公司.xlsx
```

**步骤3：生成分析表**
```bash
python3 generate_beta04_v22.py 你的公司.xlsx
```

**步骤4：配置管理（可选）**
```bash
# 添加新的科目别名
python3 subject_mapping.py --add-alias balance 货币资金 现金资产
```

### 开发状态

- ✅ JSON配置文件已完成（subject_mapping_config.json）
- ✅ 科目映射工具已完成并优化（支持JSON配置、别名管理）
- ✅ 科目验证器已完成（V2.2.1严格模式全字段验证）
- ✅ V2.2生成器开发完成并测试通过
- ✅ 完整集成已完成，生产就绪
- ✅ 项目结构优化，移除过时文档

### V2.2 功能特性

**动态科目识别**：
- ✅ 支持科目名称别名匹配（如"收入"匹配"主营业务收入"）
- ✅ 支持复合科目自动汇总（如应收 = 应收账款 + 应收票据）
- ✅ 智能搜索A列和B列，适配不同数据格式
- ✅ JSON配置驱动，灵活可扩展

**配置管理**：
- ✅ 统一的JSON配置文件管理所有映射规则
- ✅ 支持动态添加科目别名
- ✅ 版本化配置，便于维护和升级
- ✅ 详细的配置验证和错误处理

**验证策略**：
- ✅ 严格验证策略（验证所有42个beta04_field_mapping必需字段）
- ✅ 详细的缺失科目报告和完成度统计
- ✅ 防止生成不完整的分析报告

**使用方法**：
```bash
# 查看配置信息
python3 subject_mapping.py --config-info

# 查看科目映射报告
python3 subject_mapping.py 你的公司.xlsx

# 验证所有必需字段（严格模式）
python3 subject_validator_v22.py 你的公司.xlsx

# 生成分析表（动态科目识别）
python3 generate_beta04_v22.py 你的公司.xlsx
```

**兼容性**：
- ✅ 福耀玻璃数据：完全兼容
- ✅ 金山办公数据：可处理（核心科目识别通过）
- ✅ 标准财务报表格式：广泛支持

---

## 📖 详细文档

### JSON配置文件（V2.2）
- **文件**: `subject_mapping_config.json`
- **内容**: 科目映射规则、复合科目定义、验证规则、Beta_04字段映射
- **特点**: 结构化配置、易于维护、版本管理

### 配置管理工具
- **查看配置**: `python3 subject_mapping.py --config-info`
- **科目分析**: `python3 subject_mapping.py <文件.xlsx>`
- **添加别名**: `python3 subject_mapping.py --add-alias <类型> <科目> <别名>`

---

## 🔍 使用示例

### 示例1：基本分析

```bash
# 分析福耀玻璃财务数据
python3 generate_beta04_v22.py examples/福耀玻璃.xlsx 福耀玻璃_分析.xlsx
```

**输出**：
```
✅ 成功加载科目映射配置文件: subject_mapping_config.json
正在验证文件 (V2.2): examples/福耀玻璃.xlsx
✅ 核心科目验证通过：所有必需的核心科目都已找到

加载源文件: examples/福耀玻璃.xlsx
  发现年份: 2024 (balance列3, income/revenue列2)
  发现年份: 2023 (balance列4, income/revenue列3)
  ...
  发现年份: 2018 (balance列9, income/revenue列8)

开始生成 Beta_04 分析表（共 7 个年度）...

保存文件: 福耀玻璃_分析.xlsx
✓ Beta_04 分析表创建完成

【2024年关键指标】
  营收: 38,710,428,679
  净利润: 7,504,038,370
  毛利率: 35.6%
  净利率: 18.5%
  ROE: 20.1%
  EPS: 2.75
```

### 示例2：批量处理

```python
import subprocess

companies = ['福耀玻璃', '比亚迪', '宁德时代']

for company in companies:
    input_file = f'{company}.xlsx'
    output_file = f'{company}_Beta04.xlsx'

    subprocess.run([
        'python3', 'generate_beta04_v22.py',
        input_file, output_file
    ])
```

---

## ⚙️ 配置说明

### 颜色配置

脚本中定义了以下颜色常量：

```python
COLOR_HEADER = '4472C4'    # 标题蓝色
COLOR_LINK = '70AD47'      # 数据提取绿色（来自外部）
COLOR_INPUT = '0070C0'     # 手动输入蓝色
COLOR_FORMULA = '000000'   # 公式黑色
COLOR_HIGHLIGHT = 'E7E6E6' # 高亮灰色
```

### 数据行配置

```python
# 资产负债表行号
资产 = 2
现金 = 4
投资 = 10
...

# 损益表行号
研发费用 = 8
销售费用 = 6
管理费用 = 7
...
```

---

## 🧪 验证方法

### 验证数据提取

```python
import openpyxl

wb = openpyxl.load_workbook('输出文件.xlsx')
ws = wb['Beta_04']

# 检查数据提取部分（应为数值）
cell = ws['C23']  # 主营收入
print(f"类型: {type(cell.value)}")  # 应为 int 或 float
print(f"值: {cell.value}")
```

### 验证分析公式

```python
# 检查分析部分（应为公式）
cell = ws['C45']  # 毛利率
print(f"公式: {cell.value}")  # 应为 "=(C23-C24)/C23"
```

### 验证YoY逻辑

```python
# 检查YoY公式位置
c67 = ws['C67']  # 2024年YoY
print(f"2024年YoY: {c67.value}")  # 应为 "=(C23-D23)/D23"

i67 = ws['I67']  # 2018年YoY
print(f"2018年YoY: {i67.value}")  # 应为 None
```

---

## ⚠️ 科目完整性验证（重要约束）

### 验证规则

**在生成 Beta_04 分析表之前，系统必须验证所有必需科目是否能够找到。**

### 处理策略

1. **预检查**：在开始生成分析表前，先检查所有必需科目
2. **完整报告**：如果有科目找不到，统计所有缺失科目并一次性报告
3. **停止生成**：发现缺失科目时，立即停止生成流程
4. **用户完善**：提示用户完善数据源或科目映射配置
5. **禁止自动处理**：Agent 不得擅自使用其他方式（如默认值、跳过、估算等）解决缺失科目问题

### 必需科目清单

#### 资产负债表科目（17项）
- 资产、货币资金、交易性金融资产、应收账款、应收票据
- 预付款项、存货、非流动资产、负债、预收款项
- 合同负债、应付账款、应付票据、有息负债
- 所有者权益、实收资本、资本公积、未分配利润、应付股利

#### 损益表科目（19项）
- 主营业务收入、主营业务成本、研发费用、销售费用、管理费用
- 资产减值损失、信用减值损失、折旧摊销、购建固定资产支付的现金
- 其他收益、公允价值变动收益、投资收益、利息支出、财务费用
- 营业外收入、营业外支出、所得税费用

#### 收入成本表科目（2项）
- 其他业务收入、其他业务成本

### 错误提示格式

```
❌ 科目验证失败：发现 X 个缺失科目

【资产负债表缺失科目】
  - 交易性金融资产
  - 合同负债

【损益表缺失科目】
  - 资产减值损失
  - 公允价值变动收益

【收入成本表缺失科目】
  - 其他业务成本

请检查并完善以下内容：
1. 确认源数据文件包含上述科目
2. 检查科目名称是否与标准名称匹配
3. 参考 docs/Beta_04配置手册.md 了解科目映射规则
4. 如需添加别名，请修改 subject_mapping.py 中的映射配置
```

### 实现要求

- 在 `generate_beta04_v2.py` 和未来的 `generate_beta04_v2.2.py` 中实现
- 在数据提取阶段之前执行验证
- 验证失败时抛出异常并终止程序
- 提供清晰的错误信息和修复建议

### 验证工具

**V2.1 版本验证工具**: `subject_validator_v21.py`

```bash
# 独立验证数据文件
python3 subject_validator_v21.py 你的公司.xlsx
```

该工具会检查：
- 所有必需的工作表是否存在
- 年份数据是否正确配置
- 硬编码行号位置是否有数据
- 复合科目至少有一个组成部分有数据

**V2.2 版本验证工具**: `subject_validator.py`（开发中）

用于动态科目识别版本，检查所有必需科目名称是否能找到。

### 实现状态

- ✅ V2.1 验证逻辑已实现并集成到 `generate_beta04_v2.py`
- ✅ 独立验证工具 `subject_validator_v21.py` 已完成
- ✅ 测试验证通过
- 🚧 V2.2 验证工具开发中

---

## 🚨 常见问题

### Q1: 如何处理缺失数据？

**A**: 脚本使用 `safe_divide()` 函数处理除零错误，缺失数据返回 0。

```python
def safe_divide(numerator, denominator):
    if denominator == 0 or denominator is None:
        return 0
    return numerator / denominator if numerator is not None else 0
```

### Q2: 如何添加新的分析指标？

**A**: 按照以下步骤：

1. 在数据提取部分添加数据项
2. 在 `row_map` 中记录行号
3. 生成公式并调用 `write_data_row()`

```python
# 1. 提取数据
row_map['新指标'] = write_data_row('新指标', [data[y]['新指标'] for y in years])

# 2. 生成公式
新指标_formulas = []
for i in range(len(years)):
    col = get_column_letter(3 + i)
    新指标_formulas.append(f"={col}{row_map['A']}/{col}{row_map['B']}")

# 3. 写入
write_data_row('新指标', values, formulas=新指标_formulas)
```

### Q3: 如何修改源数据行号？

**A**: 修改脚本中的行号常量：

```python
# 在 generate_beta04() 函数开头修改
data[year]['主营收入'] = get_revenue_value(3, col_idx)  # 修改行号 3
```

### Q4: 输出文件太大怎么办？

**A**: 生成的文件通常只有 10-20KB。如果过大，检查：
- 是否包含了不必要的工作表
- 是否有大量格式化数据

---

## 📈 性能指标

- **处理速度**: ~2-3秒/公司（7年数据）
- **文件大小**: 10-20KB（输出文件）
- **内存占用**: <50MB
- **支持年度**: 无限制（建议 5-10 年）

---

## 🔄 更新日志

### v2.2.1 (2026-02-04)
- ✅ 优化科目验证逻辑：移除宽松模式，仅保留严格模式
- ✅ 强化质量保障：验证所有42个beta04_field_mapping必需字段
- ✅ 简化用户操作：无需选择模式参数，默认最严格标准
- ✅ 增强验证报告：显示各工作表完成度百分比和详细统计
- ✅ 确保数据完整性：防止生成不完整的财务分析表

### v2.2 (2026-02-03)
- ✅ 实现动态科目识别功能，不再依赖硬编码行号
- ✅ 支持科目名称别名匹配（如"收入"匹配"主营业务收入"）
- ✅ 支持复合科目自动汇总（如应收 = 应收账款 + 应收票据）
- ✅ 智能搜索A列和B列，适配不同数据格式
- ✅ 实现严格验证策略（验证所有beta04_field_mapping必需字段）
- ✅ 大幅提升通用性，支持不同公司的数据结构
- ✅ 完整测试通过（福耀玻璃数据）

### v2.1.1 (2026-02-03)
- ✅ 新增科目完整性验证功能
- ✅ 实现数据预检查，防止生成不完整的分析表
- ✅ 添加独立验证工具 `subject_validator_v21.py`
- ✅ 验证失败时提供详细的错误报告和修复建议
- ✅ 支持复合科目的智能验证（至少一个组成部分有数据即可）

### v2.1 (2026-02-03)
- ✅ 实现数据与公式分离
- ✅ 修复 A 列标签@符号问题
- ✅ 修复多年度数据列错位问题
- ✅ 修复 YoY 公式位置错误问题
- ✅ 完整验证通过

### v1.0 (初始版本)
- 基本数据提取功能
- 简单指标计算

---

## 💡 最佳实践

### 1. 数据准备
- 确保源文件格式标准化
- 年份标识清晰（第一行）
- 数据完整无缺失

### 2. 批量处理
- 使用脚本批量处理多个公司
- 统一命名规范
- 保存日志记录

### 3. 结果验证
- 抽查关键指标
- 验证公式逻辑
- 对比历史数据

### 4. 定期更新
- 定期更新源数据
- 重新生成分析表
- 跟踪指标变化

---

## 🤝 Agent 协作建议

### 当其他 Agent 需要使用此工具时：

1. **读取本文件**：获取完整上下文
2. **检查输入格式**：确保符合要求
3. **执行脚本**：使用标准命令
4. **验证输出**：检查关键指标
5. **参考文档**：遇到问题查阅 `docs/` 目录

### 推荐工作流：

```
1. 读取 Agents.md（本文件）
2. 准备输入文件（参考 examples/）
3. 执行 generate_beta04_v2.py
4. 验证输出结果
5. 如需详细信息，查阅 docs/ 中的文档
```

---

## 📞 支持与反馈

- **项目路径**: `/Users/qin.cui/Project/fr_beta04/financial_analysis_agent`
- **主脚本**: `generate_beta04_v2.py`
- **文档目录**: `docs/`
- **示例目录**: `examples/`

---

**最后更新**: 2026-02-04
**版本**: v2.2.1
**状态**: ✅ V2.1 生产就绪（含科目完整性验证）| ✅ V2.2.1 严格模式验证已完成

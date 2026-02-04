#!/usr/bin/env python3
"""
Beta_04 财务分析表生成器 V2.2
基于动态科目识别，支持不同公司的数据结构

使用方法：
    python generate_beta04_v22.py input.xlsx output.xlsx
"""

import sys
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 导入科目映射和验证模块
try:
    from subject_mapping import get_value_by_label, BALANCE_MAPPING, INCOME_MAPPING, REVENUE_MAPPING, load_config
    from subject_validator_v22 import validate_subjects, print_validation_report
    MAPPING_AVAILABLE = True
    VALIDATION_AVAILABLE = True

    # 从JSON配置文件读取Beta_04字段映射
    config = load_config()
    BETA04_MAPPING = config.get('beta04_field_mapping', {})

except ImportError as e:
    print(f"警告: 模块导入失败 {e}")
    MAPPING_AVAILABLE = False
    VALIDATION_AVAILABLE = False

    # 回退到默认映射
    BETA04_MAPPING = {
        'balance': {},
        'income': {},
        'revenue': {}
    }

# 颜色定义
COLOR_INPUT = 'FF0000FF'       # 蓝色：手动输入
COLOR_FORMULA = 'FF000000'     # 黑色：公式计算
COLOR_LINK = 'FF008000'        # 绿色：工作表间引用
COLOR_HIGHLIGHT = 'FFFFFF00'   # 黄色背景：关键假设


def safe_divide(a, b, default=0):
    """安全除法"""
    try:
        if b is None or b == 0:
            return default
        return a / b
    except:
        return default


def safe_calc(func, default=0):
    """安全计算"""
    try:
        result = func()
        if result is None or str(result) == 'nan':
            return default
        return result
    except:
        return default


def get_dynamic_value(sheet, mapping_dict, field_name, col_idx):
    """
    动态获取科目数值（支持复合科目）

    Args:
        sheet: 工作表对象
        mapping_dict: 科目映射字典（BALANCE_MAPPING, INCOME_MAPPING, REVENUE_MAPPING）
        field_name: Beta_04字段名称
        col_idx: 列索引

    Returns:
        数值，如果找不到返回0
    """
    # 根据工作表名称确定映射类型
    sheet_name = sheet.title
    if sheet_name == 'balance':
        field_mapping = BETA04_MAPPING.get('balance', {})
    elif sheet_name == '损益现金流':
        field_mapping = BETA04_MAPPING.get('income', {})
    elif sheet_name == '收入成本':
        field_mapping = BETA04_MAPPING.get('revenue', {})
    else:
        return 0

    if field_name not in field_mapping:
        return 0

    subject_info = field_mapping[field_name]

    # 复合科目处理（支持list和tuple格式）
    if isinstance(subject_info, (list, tuple)):
        total = 0
        for subject in subject_info:
            value = get_value_by_label(sheet, mapping_dict, subject, col_idx)
            if value is not None:
                total += value
        return total

    # 单一科目处理
    return get_value_by_label(sheet, mapping_dict, subject_info, col_idx)


def create_beta04_analysis_v22(source_file, output_file):
    """创建 Beta_04 分析表（V2.2版本 - 动态科目识别）"""

    print(f"加载源文件: {source_file}")
    source_wb = load_workbook(source_file, data_only=True)

    # 检查必需的工作表
    required_sheets = ['balance', '损益现金流', '收入成本']
    for sheet_name in required_sheets:
        if sheet_name not in source_wb.sheetnames:
            raise ValueError(f"源文件缺少必需的工作表: {sheet_name}")

    # 验证科目完整性
    if VALIDATION_AVAILABLE:
        print("\n正在验证科目完整性...")
        is_valid, missing_subjects = validate_subjects(source_wb)

        if not is_valid:
            print_validation_report(missing_subjects)
            source_wb.close()
            raise ValueError("科目验证失败，请完善数据源后重试")

        print("✅ 科目验证通过：所有必需科目都已找到\n")

    # 获取源数据表
    balance = source_wb['balance']
    income = source_wb['损益现金流']
    revenue = source_wb['收入成本']

    # 检测年份列
    years = []
    balance_cols_idx = []  # balance表的列索引（从3开始）
    income_cols_idx = []   # 损益现金流表的列索引（从2开始）
    revenue_cols_idx = []  # 收入成本表的列索引（从2开始）

    for col_idx in range(3, 10):
        col_letter = get_column_letter(col_idx)
        cell_value = balance[f'{col_letter}1'].value
        if cell_value and '20' in str(cell_value):
            year = str(cell_value).split('.')[0]
            years.append(year)
            balance_cols_idx.append(col_idx)
            income_cols_idx.append(col_idx - 1)  # income表比balance表提前1列
            revenue_cols_idx.append(col_idx - 1)  # revenue表比balance表提前1列
            print(f"  发现年份: {year} (balance列{col_idx}, income/revenue列{col_idx-1})")

    if not years:
        raise ValueError("未找到年份数据")

    print(f"\n开始生成 Beta_04 分析表（共 {len(years)} 个年度）...")

    # 创建新工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = 'Beta_04'

    # 设置列宽
    ws.column_dimensions['A'].width = 25
    for i in range(len(years)):
        col_letter = get_column_letter(3 + i)
        ws.column_dimensions[col_letter].width = 15

    # 写入标题行
    ws['A1'] = 'Beta_04 财务分析表'
    ws['A1'].font = Font(bold=True, size=14)

    # 写入年份标题
    for i, year in enumerate(years):
        col_letter = get_column_letter(3 + i)
        ws[f'{col_letter}1'] = year
        ws[f'{col_letter}1'].font = Font(bold=True)
        ws[f'{col_letter}1'].alignment = Alignment(horizontal='center')

    current_row = 3  # 从第3行开始写入数据
    row_map = {}  # 记录每个数据项的行号，用于公式引用

    def write_data_row(label, values, is_percentage=False, decimals=0,
                      is_input=False, color=None, formulas=None):
        """写入数据行的通用函数"""
        nonlocal current_row

        # 写入标签
        ws.cell(current_row, 1).value = label
        ws.cell(current_row, 1).font = Font(bold=True)

        # 写入数据或公式
        for i in range(len(years)):
            col_idx = 3 + i
            cell = ws.cell(current_row, col_idx)

            if formulas and i < len(formulas) and formulas[i]:
                # 写入公式
                cell.value = formulas[i]
                if color:
                    cell.font = Font(color=color)
            elif values and i < len(values):
                # 写入数值
                value = values[i]
                if is_percentage and value is not None:
                    cell.value = value
                    cell.number_format = '0.0%'
                else:
                    cell.value = value
                    if decimals == 0:
                        cell.number_format = '#,##0'
                    else:
                        cell.number_format = f'#,##0.{decimals * "0"}'

                if is_input:
                    cell.font = Font(color=COLOR_INPUT)
                elif color:
                    cell.font = Font(color=color)
                else:
                    cell.font = Font(color=COLOR_LINK)

        row_num = current_row
        current_row += 1
        return row_num

    # 数据提取阶段 - 使用动态科目识别
    print("正在提取数据...")
    data = {}

    for i, year in enumerate(years):
        data[year] = {}
        balance_col = balance_cols_idx[i]
        income_col = income_cols_idx[i]
        revenue_col = revenue_cols_idx[i]

        # 从资产负债表提取数据
        for field_name in BETA04_MAPPING['balance']:
            data[year][field_name] = get_dynamic_value(balance, BALANCE_MAPPING, field_name, balance_col)

        # 从损益现金流表提取数据
        for field_name in BETA04_MAPPING['income']:
            data[year][field_name] = get_dynamic_value(income, INCOME_MAPPING, field_name, income_col)

        # 从收入成本表提取数据
        for field_name in BETA04_MAPPING['revenue']:
            data[year][field_name] = get_dynamic_value(revenue, REVENUE_MAPPING, field_name, revenue_col)

        # 计算复合字段
        data[year]['费用'] = (data[year]['销售费用'] or 0) + (data[year]['管理费用'] or 0) + (data[year]['研发费用'] or 0)

    print("正在生成分析表...")

    # 1. 资产负债表数据部分
    ws.cell(current_row, 1).value = '==== 资产负债表数据 ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # 写入资产负债表数据（17项）
    balance_fields = ['资产', '现金', '投资', '运营资产', '应收', '预付', '存货', '长期资产',
                     '负债', '预收', '应付', '有息负债', '权益', '资本投入', '未分配利润', '股息分红', '股本']

    for field in balance_fields:
        values = [data[year][field] for year in years]
        row_map[field] = write_data_row(field, values, color=COLOR_LINK)

    current_row += 1

    # 2. 损益表数据部分
    ws.cell(current_row, 1).value = '==== 损益表数据 ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # 写入损益表数据（19项）
    income_fields = ['主营收入', '主营成本', '费用', '研发费用', '销售费用', '管理费用',
                    '损耗', '其他收益', '现金收益', '利息支出', '财务费用', '所得税',
                    '股权收益', '其他业务收入', '其他业务成本', '营业外收入', '营业外成本',
                    '折旧摊销', '资本开支']

    for field in income_fields:
        values = [data[year][field] for year in years]
        row_map[field] = write_data_row(field, values, color=COLOR_LINK)

    current_row += 1

    # 3. 盈利结构分析部分 - 全部使用公式
    ws.cell(current_row, 1).value = '==== 盈利结构分析 ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # 业务利润公式
    业务利润_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        formula = (f"={col}{row_map['主营收入']}-{col}{row_map['主营成本']}-{col}{row_map['费用']}"
                  f"+{col}{row_map['损耗']}+{col}{row_map['其他收益']}+{col}{row_map['现金收益']}"
                  f"-{col}{row_map['利息支出']}-{col}{row_map['财务费用']}-{col}{row_map['所得税']}")
        业务利润_formulas.append(formula)

    row_map['业务利润'] = write_data_row('业务利润', [None] * len(years),
                                       formulas=业务利润_formulas, color=COLOR_FORMULA)

    # 其他业务利润公式
    其他业务利润_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        formula = f"={col}{row_map['其他业务收入']}-{col}{row_map['其他业务成本']}"
        其他业务利润_formulas.append(formula)

    row_map['其他业务利润'] = write_data_row('其他业务利润', [None] * len(years),
                                         formulas=其他业务利润_formulas, color=COLOR_FORMULA)

    # 营业外利润公式
    营业外利润_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        formula = f"={col}{row_map['营业外收入']}-{col}{row_map['营业外成本']}"
        营业外利润_formulas.append(formula)

    row_map['营业外利润'] = write_data_row('营业外利润', [None] * len(years),
                                        formulas=营业外利润_formulas, color=COLOR_FORMULA)

    # 净利润公式
    净利润_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        formula = f"={col}{row_map['业务利润']}+{col}{row_map['股权收益']}+{col}{row_map['其他业务利润']}+{col}{row_map['营业外利润']}"
        净利润_formulas.append(formula)

    row_map['净利润'] = write_data_row('净利润', [None] * len(years),
                                    formulas=净利润_formulas, color=COLOR_FORMULA)

    # 利润留存公式
    利润留存_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        formula = f"={col}{row_map['净利润']}-{col}{row_map['股息分红']}"
        利润留存_formulas.append(formula)

    row_map['利润留存'] = write_data_row('利润留存', [None] * len(years),
                                      formulas=利润留存_formulas, color=COLOR_FORMULA)

    # 自由现金流公式
    自由现金流_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        formula = f"={col}{row_map['利润留存']}+{col}{row_map['折旧摊销']}-{col}{row_map['资本开支']}"
        自由现金流_formulas.append(formula)

    row_map['自由现金流'] = write_data_row('自由现金流', [None] * len(years),
                                        formulas=自由现金流_formulas, color=COLOR_FORMULA)

    current_row += 1

    # 4. 经营指标分析
    ws.cell(current_row, 1).value = '==== 经营指标 ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # 毛利率
    毛利率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        毛利率_formulas.append(f"=({col}{row_map['主营收入']}-{col}{row_map['主营成本']})/{col}{row_map['主营收入']}")

    write_data_row('毛利率', [None] * len(years), is_percentage=True,
                  formulas=毛利率_formulas, color=COLOR_FORMULA)

    # 费率
    费率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        费率_formulas.append(f"={col}{row_map['费用']}/{col}{row_map['主营收入']}")

    write_data_row('费率', [None] * len(years), is_percentage=True,
                  formulas=费率_formulas, color=COLOR_FORMULA)

    # 研发费率
    研发费率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        研发费率_formulas.append(f"={col}{row_map['研发费用']}/{col}{row_map['主营收入']}")

    write_data_row('研发费率', [None] * len(years), is_percentage=True,
                  formulas=研发费率_formulas, color=COLOR_FORMULA)

    # 销售费率
    销售费率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        销售费率_formulas.append(f"={col}{row_map['销售费用']}/{col}{row_map['主营收入']}")

    write_data_row('销售费率', [None] * len(years), is_percentage=True,
                  formulas=销售费率_formulas, color=COLOR_FORMULA)

    # 管理费率
    管理费率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        管理费率_formulas.append(f"={col}{row_map['管理费用']}/{col}{row_map['主营收入']}")

    write_data_row('管理费率', [None] * len(years), is_percentage=True,
                  formulas=管理费率_formulas, color=COLOR_FORMULA)

    # 净利率
    净利率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        净利率_formulas.append(f"={col}{row_map['业务利润']}/{col}{row_map['主营收入']}")

    write_data_row('净利率', [None] * len(years), is_percentage=True,
                  formulas=净利率_formulas, color=COLOR_FORMULA)

    # 应收占比
    应收占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        应收占比_formulas.append(f"={col}{row_map['应收']}/{col}{row_map['主营收入']}")

    write_data_row('应收占比', [None] * len(years), is_percentage=True,
                  formulas=应收占比_formulas, color=COLOR_FORMULA)

    # 应收周转
    应收周转_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        应收周转_formulas.append(f"={col}{row_map['主营收入']}/{col}{row_map['应收']}")

    write_data_row('应收周转', [None] * len(years), decimals=1,
                  formulas=应收周转_formulas, color=COLOR_FORMULA)

    # 存货周转
    存货周转_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        存货周转_formulas.append(f"={col}{row_map['主营成本']}/{col}{row_map['存货']}")

    write_data_row('存货周转', [None] * len(years), decimals=1,
                  formulas=存货周转_formulas, color=COLOR_FORMULA)

    # 预收占比
    预收占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        预收占比_formulas.append(f"={col}{row_map['预收']}/{col}{row_map['主营收入']}")

    write_data_row('预收占比', [None] * len(years), is_percentage=True,
                  formulas=预收占比_formulas, color=COLOR_FORMULA)

    current_row += 1

    # 5. 资产负债结构
    ws.cell(current_row, 1).value = '==== 资产负债结构 ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # 现金资产占比
    现金资产占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        现金资产占比_formulas.append(f"={col}{row_map['现金']}/{col}{row_map['资产']}")

    write_data_row('现金资产占比', [None] * len(years), is_percentage=True,
                  formulas=现金资产占比_formulas, color=COLOR_FORMULA)

    # 股权资产占比
    股权资产占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        股权资产占比_formulas.append(f"={col}{row_map['投资']}/{col}{row_map['资产']}")

    write_data_row('股权资产占比', [None] * len(years), is_percentage=True,
                  formulas=股权资产占比_formulas, color=COLOR_FORMULA)

    # 经营资产占比
    经营资产占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        经营资产占比_formulas.append(f"={col}{row_map['运营资产']}/{col}{row_map['资产']}")

    write_data_row('经营资产占比', [None] * len(years), is_percentage=True,
                  formulas=经营资产占比_formulas, color=COLOR_FORMULA)

    # 长期资产占比
    长期资产占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        长期资产占比_formulas.append(f"={col}{row_map['长期资产']}/{col}{row_map['资产']}")

    write_data_row('长期资产占比', [None] * len(years), is_percentage=True,
                  formulas=长期资产占比_formulas, color=COLOR_FORMULA)

    # 经营负债占比
    经营负债占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        经营负债占比_formulas.append(f"=({col}{row_map['预收']}+{col}{row_map['应付']})/{col}{row_map['负债']}")

    write_data_row('经营负债占比', [None] * len(years), is_percentage=True,
                  formulas=经营负债占比_formulas, color=COLOR_FORMULA)

    # 有息负债占比
    有息负债占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        有息负债占比_formulas.append(f"={col}{row_map['有息负债']}/{col}{row_map['负债']}")

    write_data_row('有息负债占比', [None] * len(years), is_percentage=True,
                  formulas=有息负债占比_formulas, color=COLOR_FORMULA)

    # 资本投入占比
    资本投入占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        资本投入占比_formulas.append(f"={col}{row_map['资本投入']}/{col}{row_map['权益']}")

    write_data_row('资本投入占比', [None] * len(years), is_percentage=True,
                  formulas=资本投入占比_formulas, color=COLOR_FORMULA)

    # 利润留存占比
    利润留存占比_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        利润留存占比_formulas.append(f"={col}{row_map['未分配利润']}/{col}{row_map['权益']}")

    write_data_row('利润留存占比', [None] * len(years), is_percentage=True,
                  formulas=利润留存占比_formulas, color=COLOR_FORMULA)

    current_row += 1

    # 6. 同比增长分析 (YoY)
    ws.cell(current_row, 1).value = '==== 同比增长分析 (YoY) ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # 营收YoY
    营收YoY_formulas = []
    for i in range(len(years)):
        if i == len(years) - 1:  # 最早年份没有YoY
            营收YoY_formulas.append(None)
        else:
            this_col = get_column_letter(3 + i)
            prev_col = get_column_letter(3 + i + 1)
            营收YoY_formulas.append(f"=({this_col}{row_map['主营收入']}-{prev_col}{row_map['主营收入']})/{prev_col}{row_map['主营收入']}")

    write_data_row('营收 YoY', [None] * len(years), is_percentage=True,
                  formulas=营收YoY_formulas, color=COLOR_FORMULA)

    # 业务净利润YoY
    业务净利润YoY_formulas = []
    for i in range(len(years)):
        if i == len(years) - 1:  # 最早年份没有YoY
            业务净利润YoY_formulas.append(None)
        else:
            this_col = get_column_letter(3 + i)
            prev_col = get_column_letter(3 + i + 1)
            业务净利润YoY_formulas.append(f"=({this_col}{row_map['业务利润']}-{prev_col}{row_map['业务利润']})/{prev_col}{row_map['业务利润']}")

    write_data_row('业务净利润 YoY', [None] * len(years), is_percentage=True,
                  formulas=业务净利润YoY_formulas, color=COLOR_FORMULA)

    # 资产YoY
    资产YoY_formulas = []
    for i in range(len(years)):
        if i == len(years) - 1:  # 最早年份没有YoY
            资产YoY_formulas.append(None)
        else:
            this_col = get_column_letter(3 + i)
            prev_col = get_column_letter(3 + i + 1)
            资产YoY_formulas.append(f"=({this_col}{row_map['资产']}-{prev_col}{row_map['资产']})/{prev_col}{row_map['资产']}")

    write_data_row('资产 YoY', [None] * len(years), is_percentage=True,
                  formulas=资产YoY_formulas, color=COLOR_FORMULA)

    # 权益YoY
    权益YoY_formulas = []
    for i in range(len(years)):
        if i == len(years) - 1:  # 最早年份没有YoY
            权益YoY_formulas.append(None)
        else:
            this_col = get_column_letter(3 + i)
            prev_col = get_column_letter(3 + i + 1)
            权益YoY_formulas.append(f"=({this_col}{row_map['权益']}-{prev_col}{row_map['权益']})/{prev_col}{row_map['权益']}")

    write_data_row('权益 YoY', [None] * len(years), is_percentage=True,
                  formulas=权益YoY_formulas, color=COLOR_FORMULA)

    current_row += 1

    # 7. ROE分解（杜邦分析）
    ws.cell(current_row, 1).value = '==== ROE 分解（杜邦分析） ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # ROE
    ROE_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        ROE_formulas.append(f"={col}{row_map['业务利润']}/{col}{row_map['权益']}")

    write_data_row('ROE', [None] * len(years), is_percentage=True,
                  formulas=ROE_formulas, color=COLOR_FORMULA)

    # 净利率（杜邦分析）
    净利率杜邦_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        净利率杜邦_formulas.append(f"={col}{row_map['业务利润']}/{col}{row_map['主营收入']}")

    write_data_row('净利率', [None] * len(years), is_percentage=True,
                  formulas=净利率杜邦_formulas, color=COLOR_FORMULA)

    # 资产周转率
    资产周转率_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        资产周转率_formulas.append(f"={col}{row_map['主营收入']}/{col}{row_map['资产']}")

    write_data_row('资产周转率', [None] * len(years), decimals=2,
                  formulas=资产周转率_formulas, color=COLOR_FORMULA)

    # 权益乘数
    权益乘数_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        权益乘数_formulas.append(f"={col}{row_map['资产']}/{col}{row_map['权益']}")

    write_data_row('权益乘数', [None] * len(years), decimals=2,
                  formulas=权益乘数_formulas, color=COLOR_FORMULA)

    current_row += 1

    # 8. 估值指标
    ws.cell(current_row, 1).value = '==== 估值指标 ===='
    ws.cell(current_row, 1).font = Font(bold=True, size=12, color='4472C4')
    current_row += 1

    # DPS（每股股息）
    DPS_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        DPS_formulas.append(f"={col}{row_map['股息分红']}/{col}{row_map['股本']}")

    write_data_row('DPS（每股股息）', [None] * len(years), decimals=2,
                  formulas=DPS_formulas, color=COLOR_FORMULA)

    # EPS（每股收益）
    EPS_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        EPS_formulas.append(f"={col}{row_map['业务利润']}/{col}{row_map['股本']}")

    write_data_row('EPS（每股收益）', [None] * len(years), decimals=2,
                  formulas=EPS_formulas, color=COLOR_FORMULA)

    # FPS（每股自由现金流）
    FPS_formulas = []
    for i in range(len(years)):
        col = get_column_letter(3 + i)
        FPS_formulas.append(f"={col}{row_map['自由现金流']}/{col}{row_map['股本']}")

    write_data_row('FPS（每股自由现金流）', [None] * len(years), decimals=2,
                  formulas=FPS_formulas, color=COLOR_FORMULA)

    # 手动输入股价（年度最高价和最低价）
    write_data_row('年度最高价', [None] * len(years), decimals=2, is_input=True, color=COLOR_INPUT)
    write_data_row('年度最低价', [None] * len(years), decimals=2, is_input=True, color=COLOR_INPUT)

    # PE和股息率会根据输入的股价自动计算
    # 这里可以添加PE和股息率的公式，但需要引用上面的手动输入行

    # 保存文件
    print(f"\n保存文件: {output_file}")
    wb.save(output_file)
    source_wb.close()

    print("✓ Beta_04 分析表创建完成")

    # 显示关键指标
    if years:
        latest_year = years[0]
        营收 = data[latest_year]['主营收入']
        业务利润 = safe_calc(lambda: (data[latest_year]['主营收入'] or 0) -
                             (data[latest_year]['主营成本'] or 0) -
                             (data[latest_year]['费用'] or 0) +
                             (data[latest_year]['损耗'] or 0) +
                             (data[latest_year]['其他收益'] or 0) +
                             (data[latest_year]['现金收益'] or 0) -
                             (data[latest_year]['利息支出'] or 0) -
                             (data[latest_year]['财务费用'] or 0) -
                             (data[latest_year]['所得税'] or 0))
        毛利率 = safe_divide((营收 or 0) - (data[latest_year]['主营成本'] or 0), 营收 or 1) * 100
        净利率 = safe_divide(业务利润, 营收 or 1) * 100
        ROE = safe_divide(业务利润, data[latest_year]['权益'] or 1) * 100
        EPS = safe_divide(业务利润, data[latest_year]['股本'] or 1)

        print("=" * 60)
        print("✓ 成功生成 Beta_04 财务分析表")
        print(f"  输入文件: {source_file}")
        print(f"  输出文件: {output_file}")
        print(f"  分析年度: {', '.join(years)}")
        print("=" * 60)
        print(f"【{latest_year}年关键指标】")
        print(f"  营收: {int(营收):,}" if 营收 else "  营收: N/A")
        print(f"  业务利润: {int(业务利润):,}" if 业务利润 else "  业务利润: N/A")
        print(f"  毛利率: {毛利率:.1f}%")
        print(f"  净利率: {净利率:.1f}%")
        print(f"  ROE: {ROE:.1f}%")
        print(f"  EPS: {EPS:.2f}" if EPS else "  EPS: N/A")


def main():
    if len(sys.argv) < 2:
        print("用法: python generate_beta04_v22.py <输入文件.xlsx> [输出文件.xlsx]")
        print("示例: python generate_beta04_v22.py 福耀玻璃.xlsx 福耀玻璃_Beta04.xlsx")
        sys.exit(1)

    source = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else source.replace('.xlsx', '_Beta04_v22.xlsx')

    if not MAPPING_AVAILABLE:
        print("❌ 错误: 科目映射模块不可用，请检查 subject_mapping.py 文件")
        sys.exit(1)

    try:
        create_beta04_analysis_v22(source, output)
    except Exception as e:
        print(f"\n✗ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
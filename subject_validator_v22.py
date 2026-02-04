#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
科目完整性验证模块 - V2.2版本
基于动态科目识别的验证
"""

from subject_mapping import find_row_by_label, BALANCE_MAPPING, INCOME_MAPPING, REVENUE_MAPPING


def validate_subjects(wb):
    """
    验证科目是否能够被动态识别（V2.2版本）

    对于V2.2版本，我们采用更宽松的验证策略：
    只要能识别到核心科目即可，不强制要求所有细分科目

    参数:
        wb: openpyxl workbook 对象

    返回:
        (is_valid, missing_subjects) 元组
        - is_valid: bool, 是否核心科目都能找到
        - missing_subjects: dict, 缺失的科目（仅核心科目）
    """
    missing_subjects = {
        'balance': [],
        '损益现金流': [],
        '收入成本': []
    }

    # 定义核心必需科目（必须要有的）
    core_balance_subjects = ['资产', '负债', '所有者权益', '货币资金', '实收资本']
    core_income_subjects = ['销售费用', '管理费用', '研发费用', '利息支出', '所得税费用']
    core_revenue_subjects = ['主营业务收入', '主营业务成本']

    # 验证资产负债表核心科目
    if 'balance' in wb.sheetnames:
        balance_ws = wb['balance']
        for subject in core_balance_subjects:
            row = find_row_by_label(balance_ws, BALANCE_MAPPING, subject)
            if row is None:
                missing_subjects['balance'].append(subject)
    else:
        missing_subjects['balance'] = core_balance_subjects

    # 验证损益现金流表核心科目
    if '损益现金流' in wb.sheetnames:
        income_ws = wb['损益现金流']
        for subject in core_income_subjects:
            row = find_row_by_label(income_ws, INCOME_MAPPING, subject)
            if row is None:
                missing_subjects['损益现金流'].append(subject)
    else:
        missing_subjects['损益现金流'] = core_income_subjects

    # 验证收入成本表核心科目
    if '收入成本' in wb.sheetnames:
        revenue_ws = wb['收入成本']
        for subject in core_revenue_subjects:
            row = find_row_by_label(revenue_ws, REVENUE_MAPPING, subject)
            if row is None:
                missing_subjects['收入成本'].append(subject)
    else:
        missing_subjects['收入成本'] = core_revenue_subjects

    # 检查是否有缺失的核心科目
    is_valid = (
        len(missing_subjects['balance']) == 0 and
        len(missing_subjects['损益现金流']) == 0 and
        len(missing_subjects['收入成本']) == 0
    )

    return is_valid, missing_subjects


def print_validation_report(missing_subjects):
    """
    打印验证报告（V2.2版本）
    """
    total_missing = (
        len(missing_subjects['balance']) +
        len(missing_subjects['损益现金流']) +
        len(missing_subjects['收入成本'])
    )

    print(f"\n❌ 核心科目验证失败：发现 {total_missing} 个缺失的核心科目\n")

    if missing_subjects['balance']:
        print("【资产负债表缺失核心科目】")
        for subject in missing_subjects['balance']:
            print(f"  - {subject}")
        print()

    if missing_subjects['损益现金流']:
        print("【损益现金流表缺失核心科目】")
        for subject in missing_subjects['损益现金流']:
            print(f"  - {subject}")
        print()

    if missing_subjects['收入成本']:
        print("【收入成本表缺失核心科目】")
        for subject in missing_subjects['收入成本']:
            print(f"  - {subject}")
        print()

    print("请检查并完善以下内容：")
    print("1. 确认源数据文件包含上述核心科目")
    print("2. 检查科目名称是否能够被动态识别")
    print("3. 参考 docs/Beta_04配置手册.md 了解科目映射规则")
    print("4. 运行 'python3 subject_mapping.py 你的文件.xlsx' 查看详细映射报告")
    print()


if __name__ == '__main__':
    import sys
    import openpyxl

    if len(sys.argv) < 2:
        print("用法: python3 subject_validator_v22.py <Excel文件路径>")
        sys.exit(1)

    input_file = sys.argv[1]

    try:
        print(f"正在验证文件 (V2.2): {input_file}")
        wb = openpyxl.load_workbook(input_file, data_only=True)

        is_valid, missing_subjects = validate_subjects(wb)

        if is_valid:
            print("\n✅ 核心科目验证通过：所有必需的核心科目都已找到")
        else:
            print_validation_report(missing_subjects)
            sys.exit(1)

        wb.close()

    except Exception as e:
        print(f"❌ 验证过程出错: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
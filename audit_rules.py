"""全面回溯检查所有核算规则"""
import os, sys
import pandas as pd
from collections import defaultdict
from dotenv import load_dotenv
load_dotenv()

from scheduler_api import (
    fetch_bitable_to_df, safe_str, safe_numeric,
    TABLE_ID_FINANCE_DETAIL, TABLE_ID_FINANCE_SUMMARY,
)

THRESHOLD = 3

def audit():
    detail = fetch_bitable_to_df(TABLE_ID_FINANCE_DETAIL)
    summary = fetch_bitable_to_df(TABLE_ID_FINANCE_SUMMARY)

    print("=" * 70)
    print("财务核算全面回溯检查")
    print("=" * 70)

    issues = []

    # ---- 1. 数据完整性 ----
    print("\n--- 1. 数据完整性 ---")
    s_spec = detail['Spec_ID'].astype(str).str.strip()
    empty_spec = detail[s_spec.isna() | (s_spec == '') | (s_spec.str.lower() == 'nan')]
    print(f"Spec_ID为空: {len(empty_spec)}")
    if len(empty_spec) > 0:
        issues.append(f"Spec_ID为空 {len(empty_spec)}条")
        for _, r in empty_spec.iterrows():
            print(f"  {r['合同编号']} {r['产品名称']} / {r['规格']}")

    empty_pn = detail[detail['项目名称'].isna() | (detail['项目名称'].astype(str).str.strip().isin(['', 'nan']))]
    print(f"项目名称为空: {len(empty_pn)}")
    if len(empty_pn) > 0:
        issues.append(f"项目名称为空 {len(empty_pn)}条")

    empty_ct = detail[detail['合同类型'].isna() | (detail['合同类型'].astype(str).str.strip() == '')]
    print(f"合同类型为空: {len(empty_ct)}")
    if len(empty_ct) > 0:
        issues.append(f"合同类型为空 {len(empty_ct)}条")
        for _, r in empty_ct.iterrows():
            print(f"  {r['合同编号']} {r['产品名称']}")

    # ---- 2. 设备费规则 ----
    print("\n--- 2. 设备费规则 (SPEC-RB800-KZP+主键盘 > 3) ---")
    eqp_rows = detail[detail['产品名称'].astype(str).str.strip() == '消防远程控制设备费']
    eqp_contracts = eqp_rows['合同编号'].astype(str).str.strip().unique()

    eqp_issues = 0
    for cno in sorted(eqp_contracts):
        c_rows = detail[detail['合同编号'].astype(str).str.strip() == cno]
        # 主键盘合计
        kb_total = 0
        for _, r in c_rows.iterrows():
            sid = safe_str(r.get('Spec_ID', ''))
            pn = safe_str(r.get('产品名称', ''))
            if sid.upper() == 'SPEC-RB800-KZP' and '主键盘' in pn:
                kb_total += safe_numeric(r.get('合同数量', 0))

        c_eqp = c_rows[c_rows['产品名称'].astype(str).str.strip() == '消防远程控制设备费']
        for _, e in c_eqp.iterrows():
            eqp_qty = safe_numeric(e.get('合同数量', 0))
            if kb_total <= THRESHOLD:
                print(f"  【错误】{cno}: 主键盘={kb_total} <= {THRESHOLD}, 不应有设备费(qty={eqp_qty})")
                eqp_issues += 1
            elif eqp_qty != kb_total - THRESHOLD:
                print(f"  【错误】{cno}: 主键盘={kb_total}, 设备费应={kb_total-THRESHOLD}, 实际={eqp_qty}")
                eqp_issues += 1
            else:
                print(f"  {cno}: OK (主键盘={kb_total}, 设备费={eqp_qty})")

    # 检查是否遗漏了设备费
    all_baogan = summary[summary['项目类型'].astype(str).str.strip() == '包干']
    for _, sr in all_baogan.iterrows():
        cno = safe_str(sr.get('合同编号', ''))
        if cno in eqp_contracts:
            continue
        if not cno:
            continue
        c_rows = detail[detail['合同编号'].astype(str).str.strip() == cno]
        kb_total = 0
        for _, r in c_rows.iterrows():
            sid = safe_str(r.get('Spec_ID', ''))
            pn = safe_str(r.get('产品名称', ''))
            if sid.upper() == 'SPEC-RB800-KZP' and '主键盘' in pn:
                kb_total += safe_numeric(r.get('合同数量', 0))
        if kb_total > THRESHOLD:
            print(f"  【遗漏】{cno}: 主键盘={kb_total} > {THRESHOLD}, 应补充设备费={kb_total-THRESHOLD}")
            eqp_issues += 1

    if eqp_issues == 0:
        print("  全部正确")
    else:
        issues.append(f"设备费问题 {eqp_issues}条")

    # ---- 3. 电话规则 ----
    print("\n--- 3. 电话规则 (每合同免费1个电话) ---")
    phone_issues = 0
    # 找所有有电话产品的合同
    phone_contracts = set()
    for _, r in detail.iterrows():
        sid = safe_str(r.get('Spec_ID', ''))
        pn = safe_str(r.get('产品名称', ''))
        if sid.upper() == 'SPEC-RB800-PHONE' and '电话' in pn:
            phone_contracts.add(safe_str(r.get('合同编号', '')))

    for cno in sorted(phone_contracts):
        c_rows = detail[detail['合同编号'].astype(str).str.strip() == cno]
        phone_total = 0
        phone_single_total = 0
        for _, r in c_rows.iterrows():
            sid = safe_str(r.get('Spec_ID', ''))
            pn = safe_str(r.get('产品名称', ''))
            if sid.upper() == 'SPEC-RB800-PHONE' and '电话' in pn:
                phone_total += safe_numeric(r.get('合同数量', 0))
                phone_single_total += safe_numeric(r.get('包干合同单采数量', 0))

        if phone_total > 0:
            expected_single = max(0, phone_total - 1)
            if phone_single_total != expected_single:
                print(f"  【错误】{cno}: 电话总={phone_total}, 单采合计应={expected_single}, 实际={phone_single_total}")
                phone_issues += 1
            elif phone_total > 1:
                print(f"  {cno}: OK (电话={phone_total}, 免费1, 单采={phone_single_total})")

    if phone_issues == 0:
        print("  全部正确")
    else:
        issues.append(f"电话规则问题 {phone_issues}条")

    # ---- 4. 关键词匹配 ----
    print("\n--- 4. 关键词匹配检查 ---")
    keywords = {
        '主键盘': 'SPEC-RB800-KZP',
        '总线盘': 'SPEC-RB800-ZXP',
        '总键盘': 'SPEC-RB800-ZXP',
        '电话': 'SPEC-RB800-PHONE',
        '多线盘': 'SPEC-RB800-DXP',
        '广播': 'SPEC-RB800-BROAD',
    }
    kw_issues = 0
    ac_rb800 = detail[detail['规格'].astype(str).str.strip().str.upper() == 'AC-RB800']
    for _, r in ac_rb800.iterrows():
        pn = safe_str(r.get('产品名称', ''))
        sid = safe_str(r.get('Spec_ID', ''))
        for kw, expected in keywords.items():
            if kw in pn:
                if sid.upper() != expected:
                    print(f"  【错误】{r['合同编号']} {pn[:40]}: Spec={sid}, 关键词'{kw}'应→{expected}")
                    kw_issues += 1
                break  # 只匹配第一个关键词

    if kw_issues == 0:
        print("  全部正确")
    else:
        issues.append(f"关键词匹配问题 {kw_issues}条")

    # ---- 5. 广播/电话是否都被计费（单采项目） ----
    print("\n--- 5. 计费覆盖检查 ---")
    uncounted = detail[
        (detail['是否计入包干'].astype(str).str.strip().isin(['', 'nan'])) |
        (detail['包干合同单采数量'].isna())
    ]
    print(f"未核算行(是否计入包干为空): {len(uncounted)}")
    for _, r in uncounted.iterrows():
        print(f"  {r['合同编号']} {r['产品名称']} Spec={r['Spec_ID']}")

    zero_price = detail[
        (detail['小计'].astype(float) == 0) &
        (detail['产品名称'].astype(str).str.strip().isin(['消防远程控制服务费', '消防远程控制包干费', '消防远程控制设备费']))
    ]
    print(f"费用行但小计为0: {len(zero_price)}")

    # ---- 6. 汇总统计 ----
    print(f"\n{'=' * 70}")
    print(f"总行数: {len(detail)}")
    print(f"合同数: {detail['合同编号'].astype(str).str.strip().nunique()}")

    total_subtotal = detail['小计'].astype(float).sum()
    print(f"总计小计: ¥{total_subtotal:,.0f}")

    by_type = detail.groupby(detail['合同类型'].astype(str).str.strip())['小计'].sum()
    print(f"按合同类型:")
    for ct, val in by_type.items():
        print(f"  {ct}: ¥{val:,.0f}")

    if issues:
        print(f"\n⚠ 发现 {len(issues)} 类问题:")
        for i in issues:
            print(f"  - {i}")
    else:
        print(f"\n 全部检查通过")

    return issues

if __name__ == '__main__':
    audit()

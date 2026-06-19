import os
import time
from datetime import datetime
from pathlib import Path
import pandas as pd

# 导入项目模块（会读取 .env）
import scheduler_api as sa

BACKUP_DIR = Path("backup_feishu")
BACKUP_DIR.mkdir(exist_ok=True)

def backup_table(table_id: str, friendly_name: str) -> str:
    try:
        df = sa.fetch_bitable_to_df(table_id)
    except Exception as e:
        raise RuntimeError(f"读取表 {friendly_name}({table_id}) 失败: {e}")
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    fname = BACKUP_DIR / f"{friendly_name}_{table_id}_{ts}.csv"
    df.to_csv(fname, index=False, encoding='utf-8-sig')
    return str(fname)


def main():
    print("开始生产写回流程：备份 -> 同步明细 -> 核算 -> 导出结果")

    # 检查表配置
    if not sa.TABLE_ID_FINANCE_DETAIL or not sa.TABLE_ID_FINANCE_SUMMARY:
        print("ERROR: 未在 .env 中配置 TABLE_ID_FINANCE_DETAIL 或 TABLE_ID_FINANCE_SUMMARY")
        return

    # 备份总表与明细表
    try:
        detail_backup = backup_table(sa.TABLE_ID_FINANCE_DETAIL, 'finance_detail')
        summary_backup = backup_table(sa.TABLE_ID_FINANCE_SUMMARY, 'finance_summary')
        print(f"已备份：明细 => {detail_backup}\n总表 => {summary_backup}")
    except Exception as e:
        print(f"备份失败，终止：{e}")
        return

    # 运行同步明细（会回写 Spec_ID 等）
    try:
        print("运行 _sync_finance_detail() ...")
        res_sync = sa._sync_finance_detail()
        print("_sync_finance_detail 返回:", res_sync)
    except Exception as e:
        print(f"同步明细失败：{e}")
        return

    # 等待短时
    time.sleep(1)

    # 运行核算并回写
    try:
        print("运行 _run_finance_calculate() ...")
        res_calc = sa._run_finance_calculate()
        print("_run_finance_calculate 返回:", res_calc)
    except Exception as e:
        print(f"核算失败：{e}")
        return

    # 再次导出明细表当前状态作为运行后快照
    try:
        post_backup = backup_table(sa.TABLE_ID_FINANCE_DETAIL, 'finance_detail_after')
        print(f"运行后明细表已备份：{post_backup}")
    except Exception as e:
        print(f"运行后备份失败：{e}")
        return

    print("生产写回流程完成。请查看备份文件并在飞书中确认明细表数据。")

if __name__ == '__main__':
    main()

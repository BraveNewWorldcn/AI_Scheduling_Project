#!/usr/bin/env python3
"""
AI排单系统启动脚本
"""

import uvicorn
import os
import sys

def main():
    # 检查必要的环境变量
    required_env_vars = [
        "FEISHU_APP_ID",
        "FEISHU_APP_SECRET",
        "BITABLE_APP_TOKEN",
        "TABLE_ID_ITEMS",
        "TABLE_ID_SKU",
        "TABLE_ID_INV",
        "TABLE_ID_SUMMARY",
        "TABLE_ID_DETAIL"
    ]

    missing_vars = []
    for var in required_env_vars:
        if not os.getenv(var):
            missing_vars.append(var)

    if missing_vars:
        print("❌ 缺少必要的环境变量:")
        for var in missing_vars:
            print(f"   - {var}")
        print("\n请设置这些环境变量后重新运行")
        sys.exit(1)

    print("✅ 环境变量检查通过")
    print("🚀 启动AI排单系统...")

    # 启动FastAPI服务器
    uvicorn.run(
        "scheduler_api:app",
        host="0.0.0.0",
        port=8000,
        reload=False,
        log_level="info"
    )

if __name__ == "__main__":
    main()
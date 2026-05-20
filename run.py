#!/usr/bin/env python3
"""
AI排单系统启动脚本
"""

import uvicorn
import os
import sys

# ----- 自动加载 .env 文件 -----
def _load_dotenv():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value

_load_dotenv()


def main():
    # 检查必要的环境变量
    required_vars = [
        ("FEISHU_APP_ID",       False),  # 非密钥，有默认值
        ("FEISHU_APP_SECRET",   True),   # 密钥，必须设置
        ("BITABLE_APP_TOKEN",   False),  # 非密钥，有默认值
        ("TABLE_ID_ITEMS",      False),
        ("TABLE_ID_SKU",        False),
        ("TABLE_ID_INV",        False),
        ("TABLE_ID_DETAIL",     False),
        ("TABLE_ID_RESERVATION",False),
    ]

    missing = []
    for var, is_secret in required_vars:
        if not os.getenv(var):
            if is_secret:
                missing.append(f"   - {var}（密钥，必须设置）")
            else:
                print(f"⚠️  {var} 未设置，将使用代码中的默认值")

    if missing:
        print("❌ 缺少必要的环境变量（密钥）:")
        for m in missing:
            print(m)
        print("\n请在项目 .env 文件中配置这些变量后重新运行")
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

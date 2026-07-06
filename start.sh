#!/bin/bash
# 启动 AI 排单服务
# 用法: ./start.sh
# 依赖: Python 3 + pip install -r requirements.txt
cd "$(dirname "$0")"
python scheduler_api.py

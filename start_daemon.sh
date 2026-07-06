#!/bin/bash
# AI 排单服务守护进程（自动重启）
# 用法: ./start_daemon.sh
# 退出: Ctrl+C 停止
cd "$(dirname "$0")"

echo "=== AI排单服务守护进程 ==="
while true; do
    echo "[$(date '+%H:%M:%S')] 启动服务..."
    python scheduler_api.py
    echo "[$(date '+%H:%M:%S')] 服务退出，3秒后自动重启..."
    sleep 3
done

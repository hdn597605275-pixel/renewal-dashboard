#!/bin/bash
cd "$(dirname "$0")"
echo "========================================"
echo "  会员续费工作台 - 局域网服务"
echo "========================================"
echo ""
IP=$(ipconfig getifaddr en0 2>/dev/null || ifconfig en0 | grep "inet " | awk '{print $2}')
if [ -z "$IP" ]; then IP=$(hostname); fi
echo "  员工请在浏览器打开："
echo "  http://$IP:8080"
echo ""
echo "  本机打开："
echo "  http://localhost:8080"
echo ""
echo "  Ctrl+C 停止服务"
echo "========================================"
python3 -m http.server 8080 --bind 0.0.0.0

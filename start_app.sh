#!/bin/bash
cd "$(dirname "$0")"

echo "=========================================================="
echo "   SỔ TÀI CHÍNH GIA ĐÌNH • HÒA & QUỲNH (MỤC TIÊU 27 TRIỆU)"
echo "=========================================================="
echo ""

# Tắt tiến trình cũ nếu đang chiếm cổng 5001
lsof -ti :5001 | xargs kill -9 2>/dev/null
pkill -f "cloudflared tunnel" 2>/dev/null

echo "1. Dang khoi dong server noi bo..."
python3 server.py &
SERVER_PID=$!

sleep 2
echo "2. Dang mo trinh duyet may Mac: http://localhost:5001"
open "http://localhost:5001" 2>/dev/null &

echo "3. Dang kich hoat duong link Online 4G (Cloudflare Tunnel)..."
./cloudflared tunnel --url http://127.0.0.1:5001 2>&1 | tee /tmp/cf_tunnel.log &
CLOUDFLARED_PID=$!

sleep 4
ONLINE_URL=$(grep -o 'https://[-a-zA-Z0-9]*\.trycloudflare\.com' /tmp/cf_tunnel.log | head -n 1)

echo ""
echo "=========================================================="
echo "  🎉 HE THONG DA SAN SANG SU DUNG!"
echo "  💻 May Mac noi bo:      http://localhost:5001"
echo "  📱 Dien thoai (Wifi):   http://$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1):5001"
if [ -n "$ONLINE_URL" ]; then
    echo "  🌐 Dien thoai (4G/5G): $ONLINE_URL"
fi
echo "=========================================================="
echo "Nhan Ctrl+C de dung server."

wait $SERVER_PID

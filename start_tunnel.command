#!/bin/bash
cd "$(dirname "$0")"

echo "=== AZURA Menu Tunnel ==="
echo ""

# Check/install cloudflared
if ! command -v cloudflared &>/dev/null; then
  echo "Installing cloudflared via brew..."
  brew install cloudflared
fi

# Start HTTP server
echo "Starting HTTP server on port 8080..."
python3 -m http.server 8080 &
HTTP_PID=$!
sleep 1

# Start tunnel
echo ""
echo "Starting Cloudflare tunnel..."
echo "Open the URL below on your phone:"
echo ""
cloudflared tunnel --url http://localhost:8080

# Cleanup
kill $HTTP_PID 2>/dev/null

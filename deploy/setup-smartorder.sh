#!/usr/bin/env bash
# 한익스프레스 스마트오더 라벨 - Lightsail(Ubuntu) 최초 설치 (독립 사이트)
set -e
REPO="https://github.com/sangjin94/hanex_smartorder.git"
APPDIR=/home/ubuntu/hanex_smartorder
DATADIR=/home/ubuntu/hanex_smartorder_data/masters

echo "[1/7] 패키지"
sudo apt update
sudo apt install -y python3-venv python3-pip nginx git

echo "[2/7] 소스"
if [ -d "$APPDIR/.git" ]; then git -C "$APPDIR" pull; else git clone "$REPO" "$APPDIR"; fi

echo "[3/7] venv + 의존성"
cd "$APPDIR"
python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

echo "[4/7] 마스터 영구 디렉터리(git 밖) + 시드 · 업로드 아카이브"
mkdir -p "$DATADIR" /home/ubuntu/hanex_smartorder_data/archive
for f in "$APPDIR"/masters/*.json; do
  bn=$(basename "$f"); [ -f "$DATADIR/$bn" ] || cp "$f" "$DATADIR/$bn"
done

echo "[5/7] systemd"
sudo cp "$APPDIR/deploy/smartorder.service" /etc/systemd/system/smartorder.service
sudo systemctl daemon-reload
sudo systemctl enable smartorder
sudo systemctl restart smartorder

echo "[6/7] nginx (독립 사이트)"
sudo cp "$APPDIR/deploy/nginx-smartorder.conf" /etc/nginx/sites-available/smartorder
sudo ln -sf /etc/nginx/sites-available/smartorder /etc/nginx/sites-enabled/smartorder
sudo nginx -t && sudo systemctl reload nginx

echo "[7/7] 확인"
sleep 1; curl -s http://127.0.0.1:8090/health && echo
IP=$(curl -s http://checkip.amazonaws.com || echo "<IP>")
echo "완료! 접속: http://$IP/"

#!/usr/bin/env bash
# 코드 갱신: git pull 후 의존성/서비스 갱신. 마스터는 영구 디렉터리라 보존됨.
set -e
APPDIR=/home/ubuntu/hanex_smartorder
git -C "$APPDIR" pull
"$APPDIR/venv/bin/pip" install -q -r "$APPDIR/requirements.txt"
sudo cp "$APPDIR/deploy/smartorder.service" /etc/systemd/system/smartorder.service
sudo systemctl daemon-reload
sudo systemctl restart smartorder
echo "갱신 완료: $(git -C "$APPDIR" rev-parse --short HEAD)"

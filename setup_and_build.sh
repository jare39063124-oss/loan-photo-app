#!/bin/bash
# ============================================
# 信贷外勤拍照 App - Buildozer 打包脚本 (WSL/Ubuntu)
# 使用方法:
#   1. 将整个 loan_photo_app 目录复制到 WSL 的 ~/loan_photo_app/
#   2. cd ~/loan_photo_app
#   3. chmod +x setup_and_build.sh
#   4. ./setup_and_build.sh
# ============================================

set -e

echo "=== 安装系统依赖 ==="
sudo apt-get update
sudo apt-get install -y \
    python3-pip \
    python3-venv \
    build-essential \
    git \
    zip \
    unzip \
    openjdk-17-jdk \
    autoconf \
    automake \
    libtool \
    libffi-dev \
    libssl-dev \
    cmake \
    libgstreamer1.0-dev \
    gstreamer1.0-plugins-base \
    gstreamer1.0-plugins-good \
    libsdl2-dev \
    libsdl2-image-dev \
    libsdl2-mixer-dev \
    libsdl2-ttf-dev \
    libportmidi-dev \
    libswscale-dev \
    libavformat-dev \
    libavcodec-dev \
    libmtdev-dev \
    xclip \
    xsel

echo "=== 创建虚拟环境 ==="
python3 -m venv venv
source venv/bin/activate

echo "=== 安装 Python 依赖 ==="
pip install --upgrade pip setuptools wheel
pip install buildozer
pip install kivy[base]
pip install openpyxl plyer pillow
pip install pyjnius
pip install requests

echo "=== 初始化 Buildozer ==="
if [ ! -f buildozer.spec ]; then
    buildozer init
fi

echo "=== 开始打包 APK (debug) ==="
buildozer android debug

echo ""
echo "=== 打包完成! ==="
echo "APK 文件位置: $(pwd)/bin/loanphoto-*-arm64-v8a-debug.apk"
echo "将此文件传输到手机即可安装"

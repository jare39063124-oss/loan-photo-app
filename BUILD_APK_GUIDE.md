# GitHub Actions 构建 Android APK 完整指南

## 一、项目结构

```
loan_photo_app/
├── .github/
│   └── workflows/
│       └── build_apk.yaml          # GitHub Actions 工作流配置
├── assets/
│   └── NotoSansSC.ttf              # 打包的中文字体(子集化)
├── main.py                          # Kivy 主程序
├── geocoder.py                      # GPS 逆地理编码
├── buildozer.spec                   # Buildozer 打包配置
└── .gitignore                       # Git 忽略规则
```

## 二、GitHub Actions 工作流 (.github/workflows/build_apk.yaml)

```yaml
name: Build Android APK

on:
  push:
    branches: [ main, master ]
  pull_request:
    branches: [ main, master ]
  workflow_dispatch:  # 允许手动触发

jobs:
  build:
    runs-on: ubuntu-22.04

    steps:
    - name: Checkout code
      uses: actions/checkout@v4

    - name: Set up Python
      uses: actions/setup-python@v5
      with:
        python-version: '3.11'

    - name: Install system dependencies
      run: |
        sudo apt-get update
        sudo apt-get install -y \
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

    - name: Set up Java
      uses: actions/setup-java@v4
      with:
        distribution: 'temurin'
        java-version: '17'

    - name: Install Python dependencies
      run: |
        python -m pip install --upgrade pip setuptools wheel
        pip install buildozer
        pip install kivy[base]
        pip install openpyxl plyer pillow
        pip install cython

    - name: Install Android SDK
      run: |
        mkdir -p $HOME/android-sdk/cmdline-tools
        cd /tmp
        wget -q https://dl.google.com/android/repository/commandlinetools-linux-14742923_latest.zip -O cmdline-tools.zip
        unzip -q cmdline-tools.zip -d $HOME/android-sdk/cmdline-tools/
        mv $HOME/android-sdk/cmdline-tools/cmdline-tools $HOME/android-sdk/cmdline-tools/latest
        echo "sdk.dir=$HOME/android-sdk" > $GITHUB_WORKSPACE/local.properties

    - name: Install Buildozer Android target dependencies
      run: |
        yes | $HOME/android-sdk/cmdline-tools/latest/bin/sdkmanager "platform-tools" "platforms;android-33" "build-tools;33.0.0" "ndk;25.2.9519653"

    - name: Build debug APK
      run: |
        buildozer android debug
      env:
        ANDROID_HOME: ${{ env.HOME }}/android-sdk
        ANDROID_SDK_ROOT: ${{ env.HOME }}/android-sdk

    - name: Upload APK artifact
      uses: actions/upload-artifact@v4
      with:
        name: loan-photo-apk
        path: bin/*.apk

    - name: Create Release
      if: github.event_name == 'push' && startsWith(github.ref, 'refs/tags/')
      uses: softprops/action-gh-release@v2
      with:
        files: bin/*.apk
      env:
        GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

## 三、Buildozer 配置 (buildozer.spec)

```ini
[app]
title = 信贷外勤拍照
package.name = loanphoto
package.domain = com.banktool
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf,otf
source.exclude_dirs = venv,__pycache__,tests,.git
version = 3.0.0

requirements = python3,kivy,openpyxl,plyer,pillow,android,pyjnius

# Android 配置
android.permissions = CAMERA,ACCESS_FINE_LOCATION,ACCESS_COARSE_LOCATION,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,READ_MEDIA_IMAGES
android.api = 33
android.minapi = 21
android.ndk = 25.2.9519653
android.arch = arm64-v8a
android.add_assets = assets/

fullscreen = 0
orientation = portrait

log_level = 2

[buildozer]
log_level = 2
warn_on_root = 1
```

## 四、关键注意事项（踩坑记录）

### 1. NDK 版本必须用完整版本号
- ✅ `"ndk;25.2.9519653"` — 这是 sdkmanager 能识别的完整包名
- ❌ `"ndk;25b"` — 简写无效
- ❌ `"ndk;28c"` — 新版 NDK 在旧 sdkmanager 中不可用
- 查询可用版本：`curl -s "https://dl.google.com/android/repository/repository2-3.xml" | grep -oP '<remotePackage path="ndk;[^"]*"'`

### 2. 不要安装 gstreamer 相关包
- `libgstreamer1.0-dev` 在 Ubuntu 22.04 runner 上依赖 `libunwind-dev`，会报 "held broken packages" 错误
- Kivy 拍照 App 不需要音视频支持，直接移除以下包：
  - libgstreamer1.0-dev
  - gstreamer1.0-plugins-base
  - gstreamer1.0-plugins-good

### 3. Fine-grained Token 权限
推送代码到 GitHub 的 Fine-grained PAT 需要以下权限：
- **Contents**: Read and write（推送代码 + 创建 workflow 文件）
- **Actions**: Read and write（触发/管理 workflow）
- **Workflows**: Read and write（修改 .github/workflows/ 下的文件）
- **Administration**: Read and write（可选，管理仓库设置）
- **Metadata**: Read-only（自动）

⚠️ 只修改权限不够，必须**重新生成 token**（fine-grained token 权限在生成时固化）

### 4. 字体打包方案
- 下载 Noto CJK 全量字体（~8MB）→ 用 fonttools 子集化到 28KB
- 放入 `assets/` 目录
- 在 `buildozer.spec` 中添加 `android.add_assets = assets/`
- 在代码中优先加载 `assets/NotoSansSC.ttf`，fallback 到系统字体

### 5. 跨设备兼容
- Android 13+ (API 33): 使用 `READ_MEDIA_IMAGES` 替代 `READ_EXTERNAL_STORAGE`
- Android 11-12 (API 30-32): 使用 `READ_EXTERNAL_STORAGE`
- Android 9-10 (API 28-29): 同时需要 `READ_EXTERNAL_STORAGE` + `WRITE_EXTERNAL_STORAGE`
- 在 `buildozer.spec` 中声明所有权限，运行时按版本动态请求

## 五、构建流程

1. 代码推送到 GitHub master 分支
2. GitHub Actions 自动触发（或手动在 Actions 页面点击 "Run workflow"）
3. 等待 5-10 分钟
4. 在 Actions 页面下载 artifact（zip 文件，内含 .apk）
5. 解压后传到手机安装

## 六、下载 APK 的几种方式

### 方式 1: Actions Artifacts（推荐）
- 进入仓库 → Actions → 点击成功的 workflow run → 底部 Artifacts 区域下载

### 方式 2: Releases（需要打 tag）
```bash
git tag v3.0.0
git push origin v3.0.0
```
workflow 会自动创建 Release 并上传 APK

### 方式 3: API 下载
```bash
# 获取最新成功的 run 的 artifact URL
curl -s "https://api.github.com/repos/USER/REPO/actions/runs?per_page=1&status=success" | jq '.workflow_runs[0].artifacts_url'
```

## 七、仓库信息

- 仓库: https://github.com/jare39063124-oss/loan-photo-app
- 项目路径: D:\hermes\loan_photo_app\
- 测试数据: C:\Users\Administrator\Desktop\走访客户测试数据.xlsx

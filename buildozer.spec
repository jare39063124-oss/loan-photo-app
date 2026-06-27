[app]
# 应用信息
title = 信贷外勤拍照
package.name = loanphoto
package.domain = com.banktool
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf,otf
source.exclude_dirs = venv,__pycache__,tests,.git
android.aars = 
# 包含 assets 目录
android.add_assets = assets/
version = 1.0.0

# 依赖
requirements = python3,kivy,openpyxl,plyer,pillow,android,pyjnius

# Android 配置
android.permissions = CAMERA,ACCESS_FINE_LOCATION,ACCESS_COARSE_LOCATION,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,READ_MEDIA_IMAGES
android.api = 33
android.minapi = 21
android.ndk = 25.2.9519653
android.arch = arm64-v8a

# 全屏
fullscreen = 0

# 屏幕方向
orientation = portrait

# 图标和启动屏
# icon.filename = %(source.dir)s/data/icon.png
# presplash.filename = %(source.dir)s/data/presplash.png

# 日志
log_level = 2

[buildozer]
log_level = 2
warn_on_root = 1

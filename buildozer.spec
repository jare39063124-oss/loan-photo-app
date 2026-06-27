[app]
title = 信贷外勤拍照
package.name = loanphoto
package.domain = com.banktool
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf,otf
source.exclude_dirs = venv,__pycache__,tests,.git
android.aars =
android.add_assets = assets/
version = 3.0.0

requirements = python3,kivy,openpyxl,plyer,pillow,android,pyjnius

android.permissions = CAMERA,ACCESS_FINE_LOCATION,ACCESS_COARSE_LOCATION,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,READ_MEDIA_IMAGES
android.api = 33
android.minapi = 21
android.ndk = 28c
android.build_tools_version = 33.0.0
android.accept_sdk_license = True
android.arch = arm64-v8a

fullscreen = 0
orientation = portrait
log_level = 2

[buildozer]
log_level = 2
warn_on_root = 1

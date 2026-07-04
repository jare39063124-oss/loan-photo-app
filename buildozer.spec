[app]
title = 资产盘点专项拍照工具
package.name = loanphoto
package.domain = com.banktool
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf,otf,xml
source.exclude_dirs = venv,__pycache__,tests,.git
android.aars =
version = 3.22.10

requirements = python3,kivy,openpyxl,plyer,pillow,android,pyjnius,requests,certifi,urllib3

android.permissions = CAMERA,INTERNET,ACCESS_NETWORK_STATE,ACCESS_FINE_LOCATION,ACCESS_COARSE_LOCATION,READ_MEDIA_IMAGES,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.api = 35
android.minapi = 24
android.ndk = 28c
android.build_tools_version = 35.0.0
android.accept_sdk_license = True
android.arch = arm64-v8a
android.gradle_dependencies = androidx.core:core:1.12.0
android.gradle_manifest_items = android:exported="true"
android.add_resources = res
android.extra_manifest_application_xml = res/extra_manifest_application.xml

fullscreen = 0
orientation = portrait
log_level = 2

[buildozer]
log_level = 2
warn_on_root = 1

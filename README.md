# 信贷外勤拍照工具

银行抵押物现场勘查拍照工具 — Android APP（Kivy + Python）

---

## 项目信息

| 项目 | 内容 |
|------|------|
| **项目名称** | loan-photo-app（信贷外勤拍照） |
| **版本** | v3.2 |
| **作者** | 王硕 |
| **联系方式** | 15940454123（同微信） |
| **本地路径** | `D:\hermes\loan_photo_app\` |
| **GitHub 仓库** | https://github.com/jare39063124-oss/loan-photo-app |
| **构建状态** | GitHub Actions 自动构建 |
| **APK 输出** | `D:\hermes\loan_photo_app\bin\loanphoto-*-arm64-v8a-debug.apk` |

---

## 功能概述

1. **Excel 数据导入** — 读取客户清单 Excel 文件（A=借款人名称、B=抵押物地址、C=抵押物性质、D=房号）
2. **四类拍照引导** — 远景 / 近景 / 内部 / 瑕疵，每种类型有拍摄提示
3. **自定义命名规则** — 通过模板变量自由组合文件名格式
4. **照片水印** — 自定义水印文字、位置、字号，可开关
5. **进度持久化** — 按借款人+地址哈希保存进度，跨 Excel 文件可识别
6. **照片管理** — 查看/删除已拍照片
7. **勘查报告生成** — 一键生成 Excel 日报表
8. **搜索与导航** — 搜索借款人、快速跳转行

---

## 技术栈

- **框架**: Kivy 2.3.1 + Python 3.11
- **Android 打包**: Buildozer 1.6.0 + python-for-android
- **图片处理**: Pillow（PIL）
- **Excel**: openpyxl
- **定位**: 高德地图逆地理编码 API + 离线 fallback
- **目标 SDK**: API 35（Android 15）
- **最低 SDK**: API 26（Android 8.0）
- **架构**: arm64-v8a
- **NDK**: r28c
- **权限**: CAMERA, ACCESS_FINE_LOCATION, READ_MEDIA_IMAGES

---

## Excel 格式

### 标准格式（推荐）

| A列（借款人名称） | B列（抵押物地址） | C列（抵押物性质） | D列（房号/可选） |
|:---:|:---:|:---:|:---:|
| 张三 | XX市XX路XX小区 | 住宅 | 3栋1单元502 |
| 李四 | XX市XX路XX大厦 | 商铺 | A栋101 |

- 程序只读取前 **4列**
- **B列不参与文件名命名**（仅作为水印地址参考）
- **D列可选**，填写房号/单元号用于区分同一地址下的不同抵押物

### 旧格式兼容

如果 Excel 只有 3 列（无 D列），程序自动兼容，`{unit}` 变量留空。

---

## 命名规则模板

### 默认模板

```
{date}-{borrower}-{unit}-{seq:02d}
```

示例输出：`20260627-张三-3栋1单元502-01.jpg`

### 可用变量

| 变量 | 说明 | 示例值 |
|------|------|--------|
| `{date}` | 拍摄日期（yyyymmdd） | 20260627 |
| `{borrower}` | 借款人名称（A列） | 张三 |
| `{address}` | 抵押物地址（B列） | XX路XX小区 |
| `{property_type}` | 抵押物性质（C列） | 住宅 |
| `{unit}` | 房号/单元号（D列） | 3栋1单元502 |
| `{seq}` | 序号（01开始） | 01, 02... |
| `{type}` | 拍照类型 | 远景, 近景, 内部, 瑕疵 |

### 示例模板

| 模板 | 输出示例 |
|------|----------|
| `{date}-{borrower}-{unit}-{seq:02d}` | `20260627-张三-502-01.jpg` |
| `{borrower}_{date}_{type}_{seq}` | `张三_20260627_远景_01.jpg` |
| `{date}-{borrower}-{property_type}-{seq}` | `20260627-张三-住宅-01.jpg` |

可在 App 设置页（⚙）中自定义。

---

## 水印设置

| 选项 | 说明 | 默认值 |
|------|------|--------|
| 水印开关 | 启用/禁用 | 启用 |
| 水印文字模板 | 支持 `{date}` `{address}` `{borrower}` `{property_type}` | `{date} {address}` |
| 位置 | 4个角落可选 | 右下角 |
| 字号 | 12~60 可调 | 28 |
| 不透明度 | 半透明黑色背景 | 170/255 |

---

## 安装与构建

### 从 GitHub Actions 下载 APK

1. 打开 https://github.com/jare39063124-oss/loan-photo-app/actions
2. 选择最新成功的构建
3. 下载 **loan-photo-apk** artifact
4. 解压得到 `.apk` 文件，传到手机安装

### 本地构建

```bash
# 克隆仓库
git clone https://github.com/jare39063124-oss/loan-photo-app.git
cd loan-photo-app

# 安装构建工具
pip install buildozer cython

# 构建 APK（需要 Linux/macOS WSL）
buildozer android debug
```

### 构建依赖

- Python 3.11+
- Java 17 (Temurin)
- Android SDK + NDK r28c（buildozer 自动管理）
- Linux 或 WSL（buildozer 不支持 Windows 原生）

---

## 项目结构

```
D:\hermes\loan_photo_app\
├── main.py                  # 主程序（~1500行）
├── geocoder.py              # 地名解析（高德 API + 离线）
├── buildozer.spec           # Buildozer 构建配置
├── local.properties         # Android SDK/NDK 路径（本地开发）
├── requirements.txt         # Python 依赖（若有）
├── .github/
│   └── workflows/
│       └── build_apk.yaml   # GitHub Actions 自动构建
├── assets/
│   ├── NotoSansSC.ttf       # 中文字体（28KB 子集化）
│   └── NotoSansSC-Regular.otf
├── data/
│   └── locations.json       # 离线地名数据库
└── bin/                     # 构建产物 APK
```

---

## 拍照类型说明

| 类型 | 拍摄内容 |
|------|----------|
| 🏞 远景 | 小区/厂区全貌、楼栋外立面、宗地全貌 |
| 🚪 近景 | 单元门口、楼层门牌、房号牌、宗地界桩 |
| 🏠 内部 | 室内全景、核心区域现状、厂房/设备整体 |
| ⚠️ 瑕疵 | 破损、漏水、违建、占用、查封等异常特写 |

---

## GitHub Token 配置

如果需要在本地执行 GitHub API 操作（如下载构建产物）：

```bash
# Token 存储在
C:\Users\Administrator\.github-token

# Git 远程 URL 已配置为
https://jare39063124-oss:{token}@github.com/jare39063124-oss/loan-photo-app.git

# bashrc 中设置了
export GITHUB_TOKEN=$(cat /c/Users/Administrator/.github-token)
export GH_TOKEN=$(cat /c/Users/Administrator/.github-token)
```

---

## 构建历史

| 日期 | 版本 | 说明 |
|------|------|------|
| 2026-06-27 | v3.0 | 初始 APK 构建成功，四类拍照引导 |
| 2026-06-27 | v3.1 | 欢迎页 + 设置页 + 自定义命名/水印 + UI 翻新 + 目标 API 35 |
| 2026-06-27 | v3.2 | 崩溃日志捕获 + D列房号 + 跨文件进度持久化 + 照片路径验证 |

### 踩坑记录（永久保存）

1. **SDK 下载文件名不匹配**：`wget` 默认文件名与 `unzip` 目标不一致导致 SDK 安装失败 → 加 `-O` 参数
2. **buildozer 忽略手动安装的 SDK/NDK**：buildozer 用自己的目录 `~/.buildozer/` → 删除手工配置，完全信任 buildozer
3. **NDK 版本号与下载 URL 不匹配**：`android.ndk=25.2.9519653` 导致构建无效 URL → 改为 `android.ndk=28c`
4. **build-tools 许可证未接受**：`android.accept_sdk_license = True` 解决
5. **Git 远程 token 截断**：PAT 包含下划线等字符导致 Bash/Python 字符串解析失败 → 用独立文件存储 token
6. **Android 16 兼容**：目标 API 从 33 提升至 35，build-tools 升级至 35.0.0

---

## 联系作者

**王硕**
- 电话 / 微信：15940454123
- 有问题请联系作者

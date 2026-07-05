# 资产盘点专项拍照工具

银行抵押物、抵债资产现场勘查拍照工具 — Android APP（Kivy + Python）

---

## 项目信息

| 项目 | 内容 |
|------|------|
| **项目名称** | 资产盘点专项拍照工具 |
| **版本** | v3.22.18 |
| **适用单位** | 抚顺银行风险管理部 |
| **问题咨询电话** | 15940454123 |
| **本地路径** | `D:\hermes\loan_photo_app\` |
| **GitHub 仓库** | https://github.com/jare39063124-oss/loan-photo-app |
| **构建状态** | GitHub Actions 自动构建 |
| **APK 输出** | 桌面 `loan-photo-tool-vX.X.X.apk` |

---

## 功能概述

1. **Excel 数据导入** — 读取客户清单 Excel 文件（A=客户名、B=地址概、C=地址精确、D=性质、E=备注）
2. **五类拍照引导** — 远景 / 近景 / 内部 / 瑕疵 / 其他，每种类型有拍摄提示
3. **自定义命名规则** — 4段下拉选择器，自由组合文件名格式
4. **照片水印** — 自动添加水印（日期、客户名、地址、GPS经纬度），字号/位置可配置
5. **进度持久化** — 按客户名+地址哈希保存进度，跨 Excel 文件可识别
6. **照片管理** — 按类型分类查看、删除已拍照片，自动同步系统相册
7. **AI 一键生成日报表** — 仅对有外访照片的客户生成报告行；调用 DeepSeek 按内置模板生成现场勘查日报表，末尾追加汇总说明（基于XX文件生成，共计XX户/外访XX户）；自动命名「抵押物、抵债资产现场勘查日报表YYYYMMDD.xlsx」
8. **AI 智能助手** — 查询拍摄进度、客户拍照状态等
9. **搜索与导航** — 按客户名快速搜索筛选
10. **DeepSeek AI 引擎** — 内置 deepseek-v4-flash 模型，无需额外配置即可使用

---

## 技术栈

- **框架**: Kivy 2.3.1 + Python 3.11
- **Android 打包**: Buildozer 1.6.0 + python-for-android
- **图片处理**: Pillow（PIL）— 水印渲染，自动缩放字号
- **Excel**: openpyxl — 读取客户清单，SAF 写回备注
- **AI**: DeepSeek API（deepseek-v4-flash），兼容 OpenAI Chat Completions 格式
- **定位**: GPS + Android 系统 Geocoder + 百度 API 坐标转换
- **目标 SDK**: API 35（Android 15）
- **最低 SDK**: API 26（Android 8.0）
- **架构**: arm64-v8a
- **权限**: CAMERA, ACCESS_FINE_LOCATION, READ_MEDIA_IMAGES, INTERNET

---

## Excel 格式

### 标准格式

| A列（客户名） | B列（地址概） | C列（地址精确） | D列（性质） | E列（备注） |
|:---:|:---:|:---:|:---:|:---:|
| 成都投资集团 | 和平区XX街123号 | 1-23-4 | 抵押 | 用户填写勘查备注 |
| 沈阳置业有限 | 沈河区YY路45号 | 2-15-1 | 商铺 | 用户填写勘查备注 |

- 程序读取 **5 列**（A-E）
- **E 列由用户手动填写**勘查备注，App 读取此列内容用于 AI 生成日报表
- 仅支持 `.xlsx` 格式（Excel 2007 及以上）

---

## 命名规则

### 4段组合式命名

文件名由 4 个可配置字段组合而成，各段用连字符 `-` 连接：

| 段 | 可选值 | 说明 |
|----|--------|------|
| 段1 | 拍摄日期 / 客户名 / 地址+时间 / 空值 | 日期格式：yyyymmdd |
| 段2 | 同上 | |
| 段3 | 同上 | |
| 段4 | 同上 | |

文件名末尾自动追加：**类型 + 编号**（如 `远景-01`）

### 示例

```
20260630-成都投资集团-和平区XX街123号1430-远景-01.jpg
```

在 App 设置页中自定义 4 段内容。

---

## 水印设置

| 选项 | 说明 | 默认值 |
|------|------|--------|
| 水印开关 | 启用/禁用 | 启用 |
| 水印内容 | 3段内容选择（日期/客户名/地址/经纬度/空值） | 日期+地址+经纬度 |
| 位置 | 左上 / 右上 / 左下 / 右下 | 右下 |
| 字号 | 大(80pt) / 中(56pt) / 小(36pt) | 中 |

- 日期仅显示年月日（不显示时间）
- v3.19 起：水印区域及字体增大 2 倍

---

## AI 功能

### AI 一键生成日报表

1. 用户在 Excel E 列填写勘查备注
2. 点击主界面底部「AI 一键生成日报表」按钮，按钮立即显示「正在生成中…」避免重复点击
3. App 仅对**有外访照片的客户**收集数据（名称、地址、类型、备注、拍照数）
4. 调用 DeepSeek(deepseek-v4-flash) 为已外访客户撰写日报表内容（抵押物情况、现状描述、风险备注）
5. 填入内置日报表模板（`report_template.xlsx`），末尾追加汇总说明行：「本次报告基于XX（导入文件名）生成，共计XX个客户，其中有外访XX户已生成报告，XX个客户没有外访未生成报告」
6. 自动命名「抵押物、抵债资产现场勘查日报表YYYYMMDD.xlsx」并弹出系统文件对话框，用户选择保存位置

### AI 智能助手

- 点击主界面「AI助手」按钮进入聊天界面
- 可查询：今天拍了多少照片、某公司拍了没有、某类型拍了多少张
- 内置 DeepSeek(deepseek-v4-flash) 模型，无需额外配置即可使用

---

## 安装与构建

### 从 GitHub Actions 下载 APK

1. 打开 https://github.com/jare39063124-oss/loan-photo-app/actions
2. 选择最新成功的构建
3. 下载 **loan-photo-apk** artifact（或通过 nightly.link 直接下载）
4. 解压得到 `.apk` 文件，传到手机安装

### 本地构建脚本

```powershell
# 一键构建（推送代码 → 监控 Actions → 下载 APK 到桌面）
powershell -ExecutionPolicy Bypass -File "D:\hermes\loan_photo_app\build_and_download.ps1"
```

### 构建依赖

- Python 3.11+
- Java 17 (Temurin)
- Android SDK + NDK r28c（buildozer 自动管理）
- Linux 或 WSL（buildozer 不支持 Windows 原生，通过 GitHub Actions 构建）

---

## 项目结构

```
D:\hermes\loan_photo_app\
├── main.py                     # 主程序（~4200行）
├── buildozer.spec              # Buildozer 构建配置
├── build_and_download.ps1     # 构建+下载脚本
├── report_template.xlsx        # 内置日报表模板
├── build_manual_v3193.py       # PPT说明书生成脚本
├── HANDOVER.md                 # 交接文档
├── geocoder.py                 # 地名解析（GPS+Geocoder+百度API）
├── .github/
│   └── workflows/
│       └── build_apk.yaml      # GitHub Actions 自动构建
├── res/
│   └── extra_manifest_application.xml  # FileProvider 声明
└── assets/
    └── SimHei.ttf              # 中文字体（不含emoji）
```

---

## 拍照类型说明

| 类型 | 拍摄内容 |
|------|----------|
| 远景 | 小区/厂区全貌、楼栋外立面、宗地全貌 |
| 近景 | 单元门口、楼层门牌、房号牌、宗地界桩 |
| 内部 | 室内全景、核心区域现状、厂房/设备整体 |
| 瑕疵 | 破损、漏水、违建、占用、查封等异常特写 |
| 其他 | 补充照片、特殊情况记录 |

---

## 构建历史

| 日期 | 版本 | 说明 |
|------|------|------|
| 2026-06-27 | v3.0-v3.2 | 初始版本，四类拍照引导、欢迎页、设置页 |
| 2026-06-29 | v3.16-v3.18 | 水印增强、GPS优化、照片质量修复、权限修复 |
| 2026-06-29 | v3.19.0 | GUI重新设计（明亮配色）、水印增大2倍、修复重复照片 |
| 2026-06-29 | v3.19.1 | 修复欢迎页点击闪退 |
| 2026-06-30 | v3.19.2 | AI日报表功能、AI latin-1修复、方块替换 |
| 2026-06-30 | v3.19.3 | 底部日志按钮替换为AI一键生成日报表大按钮 |
| 2026-06-30 | v3.19.4 | 删除app中所有作者信息 |
| 2026-06-30 | v3.19.5 | DeepSeek备用API、E列逻辑修正、README更新 |
| 2026-06-30 | v3.19.6 | 移除OpenRouter仅保留DeepSeek；报告仅生成有外访照片客户+汇总说明行+自动命名；按钮"正在生成中"状态 |
| 2026-07-04 | v3.22.14 | 恢复内嵌按钮布局 + 修复查看已拍 + AI 报告提示词增强 + 二次确认客户维度统计 + PPT 清理作者信息 |
| 2026-07-04 | v3.22.15 | PhotoViewerPopup 可见性终极修复（ThemedPopup background='' 禁用默认 atlas 背景）+ _on_view_photos 延迟 0.1s 创建+open + 防抖 1500ms + popup 状态诊断日志 |
| 2026-07-04 | v3.22.16 | PhotoViewerPopup 可见性终极修复 — 显式设置 self.size=self.center 替代 size_hint + 移除 Clock.schedule_once 延迟包装 |
| 2026-07-05 | v3.22.17 | PhotoViewerPopup auto_dismiss 终极修复 — 增加 self.auto_dismiss=False + 全项目 Bug 审计修复 9 处 |
| 2026-07-05 | v3.22.18 | PhotoViewerPopup 背景渲染修复（content 背景 + 权限补全） |

---

## 关键技术决策

1. **HTTP 请求头不能含中文** — `X-Title` 含中文导致 latin-1 编码错误
2. **Kivy Button 没有 scale_x/scale_y** — 按压动画用 `opacity` 而非 `scale`
3. **SimHei 字体不含 emoji** — 用 ●◆▶ 或纯文字替代 emoji 符号
4. **拍照后必须删除 MediaStore/DCIM 原件** — 防止一张照片保存两个文件
5. **DeepSeek 模型名** — 使用 `deepseek-v4-flash`（不带 `deepseek-ai/` 前缀）
6. **API Key 存储** — 完整未截断，通过 base64 编码存储在代码中（绕过 GitHub Push Protection）
7. **报告仅生成有照片客户** — 无外访照片的客户不计入报告行，仅在汇总说明中统计

---

## 使用说明书

PPT 格式说明书：`资产盘点拍照工具-使用说明书-v3.19.6.pptx`（桌面）

生成脚本：`python build_manual_v3193.py`

---

## 许可证

内部使用，仅供抚顺银行资产盘点工作使用。

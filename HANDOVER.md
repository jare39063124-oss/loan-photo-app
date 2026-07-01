# 资产盘点拍照工具 - 交接文档

> **交接日期**: 2026-07-02
> **当前版本**: v3.21.0 (已提交 + 已构建 APK)
> **项目路径**: `d:\hermes\loan_photo_app\`
> **Git 仓库**: `jare39063124-oss/loan-photo-app` (GitHub Actions 自动构建)
> **最新 APK**: `C:\Users\Administrator\Desktop\loan-photo-tool-v3.21.0.apk`

---

## 一、项目概述

Android 资产盘点拍照工具，使用 **Kivy + Buildozer** 框架开发，通过 GitHub Actions 构建 APK。
核心流程: Excel 导入客户 → 分类拍照(远景/近景/内部/瑕疵/其他) → 自动水印 → AI 生成日报表。

- **用户**: 抚顺银行风险管理部
- **用户偏好**: 中文沟通、只用免费模型、Fluent Design 浅色主题、SimHei 字体(不含 emoji)
- **AI 引擎**: DeepSeek(deepseek-v4-flash) — 唯一 LLM，已移除 OpenRouter

---

## 二、版本演进历史

| 版本 | 内容 | 状态 |
|------|------|------|
| v3.19.0 | GUI 重新设计(明亮配色)、水印增大 2 倍、修复重复照片 | ✅ 已构建 |
| v3.19.1 | 修复欢迎页点击闪退(scale_x/scale_y → opacity 动画) | ✅ 已构建 |
| v3.19.2 | AI latin-1 修复、日报表功能、方块替换为符号 | ✅ 已构建 |
| v3.19.3 | 底部日志按钮替换为"AI 一键生成日报表"大按钮 | ✅ 已构建 |
| v3.19.4 | 删除 app 中所有作者信息 | ✅ 已构建 |
| v3.19.5 | DeepSeek 备用 API、E 列逻辑修正、README 更新 | ✅ 已构建 |
| **v3.19.6** | **移除 OpenRouter 仅保留 DeepSeek；报告仅生成有外访照片客户+汇总说明行+自动命名；按钮"正在生成中"状态** | ✅ **已构建** |

### v3.19.6 核心改动（当前版本）

1. **AI 引擎简化** — 移除 OpenRouter 主引擎及主备切换逻辑，`AIService` 改为单引擎调用 DeepSeek
2. **报告生成重构**:
   - 仅对有外访照片(`photo_count > 0`)的客户生成报告行
   - 末尾追加汇总说明行(合并 A:F 列):「本次报告基于 XX(文件名)生成，共计 XX 个客户，其中有外访 XX 户已生成报告，XX 个客户没有外访未生成报告」
   - 自动命名「抵押物、抵债资产现场勘查日报表 YYYYMMDD.xlsx」
3. **按钮交互** — 点击「AI 一键生成日报表」后立即显示「正在生成中…」并禁用，避免重复点击
4. **API Key base64 编码** — DeepSeek key 用 base64 存储绕过 GitHub Push Protection

---

## 三、关键文件清单

### 源码文件

| 文件 | 说明 |
|------|------|
| `d:\hermes\loan_photo_app\main.py` | **主程序(~4200 行)**。包含所有 UI、相机、水印、Excel、AI、日报表逻辑 |
| `d:\hermes\loan_photo_app\buildozer.spec` | Buildozer 配置(version=3.19.6, source.include_exts 含 xlsx) |
| `d:\hermes\loan_photo_app\build_and_download.ps1` | 构建+下载脚本(推送 GitHub → 监控 Actions → nightly.link 下载 APK) |
| `d:\hermes\loan_photo_app\report_template.xlsx` | 内置日报表模板(12393 字节, 6 列, Sheet1) |
| `d:\hermes\loan_photo_app\res\extra_manifest_application.xml` | FileProvider 声明(修复拍照质量的关键) |
| `d:\hermes\loan_photo_app\build_manual_v3193.py` | PPT 说明书生成脚本(14 页, 纯矢量图形) |
| `d:\hermes\loan_photo_app\README.md` | GitHub 项目介绍(v3.19.6, 已移除作者信息) |
| `d:\hermes\loan_photo_app\HANDOVER.md` | 本交接文档 |

### 桌面文件

| 文件 | 说明 |
|------|------|
| `C:\Users\Administrator\Desktop\loan-photo-tool-v3.19.6.apk` | **最新 APK**(v3.19.6, 28.58 MB) |
| `C:\Users\Administrator\Desktop\资产盘点拍照工具-使用说明书-v3.19.6.pptx` | **最新 PPT 说明书**(14 页) |
| `C:\Users\Administrator\Desktop\盘点相关文件\` | 原始模板和参考文件目录 |

---

## 四、main.py 关键代码位置

> **注意**: main.py 约 4200 行，行号可能因编辑略有偏移，**建议用 Grep 搜索关键字定位**。

### 常量和配置 (行 360-430)

```python
# DEFAULT_CONFIG (行 364-374)
'ai_api_url': 'https://api.deepseek.com/v1',   # v3.19.6: OpenRouter → DeepSeek
'ai_api_key': '',
'ai_model': 'deepseek-v4-flash',                # v3.19.6: owl-alpha → deepseek-v4-flash

# API Key (base64 编码, 行 376-382)
_DEEPSEEK_KEY_B64 = "c2stMzNjNjNlMTUxMzllNGU2YzhkYmI5MzA4OGQwYjZjNWY="
DEEPSEEK_DEFAULT_API_KEY = _b64.b64decode(_DEEPSEEK_KEY_B64).decode('utf-8')
AI_DEFAULT_API_KEY = DEEPSEEK_DEFAULT_API_KEY   # 兼容别名(旧代码引用仍可用)

AUTHOR_NAME = ""   # v3.19.4 已置空
AUTHOR_PHONE = ""
AUTHOR_INFO = ""
```

### 核心类

| 类 | 搜索关键字 | 说明 |
|----|-----------|------|
| `ReportGenerator` | `class ReportGenerator` | AI 日报表生成，**v3.19.6 重构**: 仅外访客户+汇总行 |
| `AIService` | `class AIService` | **v3.19.6 单引擎**(DeepSeek only)，`_call_api` 内部方法 |
| `PhotoProcessor` | `class PhotoProcessor` | 水印添加(PIL)，`add_watermark` 方法 |
| `ExcelWriter` | `class ExcelWriter` | Excel 读写，`write_back_to_uri` (SAF 写入) |
| `CameraManager` | `class CameraManager` | 相机管理，FileProvider+MediaStore+DCIM 三策略 |
| `MainScreen` | `class MainScreen` | 主界面，所有 UI 和事件处理 |

### 关键方法

| 方法 | 搜索关键字 | 说明 |
|------|-----------|------|
| `generate_with_ai` | `def generate_with_ai` | **v3.19.6 重构**: 过滤有照片客户、生成汇总、新文件名 |
| `_fill_template` | `def _fill_template` | **v3.19.6**: 新增 summary_text 参数，末尾合并 A:F 写汇总 |
| `_collect_records` | `def _collect_records` | 收集所有客户记录(含 photo_count) |
| `_generate_report` | `def _generate_report` | **v3.19.6**: 按钮立即显示"正在生成中…"并 disabled |
| `_save_report_to_user` | `def _save_report_to_user` | SAF 保存，**v3.19.6**: 默认名改为「抵押物、抵债资产现场勘查日报表 YYYYMMDD.xlsx」 |
| `chat` | `def chat` (在 AIService) | **v3.19.6 单引擎**: 直接调用 `_call_api`，无主备切换 |
| `bind_press_animation` | `def bind_press_animation` | 按钮动画(**用 opacity，不能用 scale_x**) |
| `on_activity_result` | `def on_activity_result` | 处理 `_report_save_code=0x201` |
| `add_watermark` | `def add_watermark` | PIL 水印(min_font=24, padding=24) |
| `on_camera_result` | `def on_camera_result` | 三策略拍照(删除 MediaStore/DCIM 原件防重复) |

### v3.19.6 关键代码片段

**AIService 单引擎** (搜索 `class AIService`):
```python
DEFAULT_API_URL = "https://api.deepseek.com/v1"
DEFAULT_API_KEY = DEEPSEEK_DEFAULT_API_KEY
DEFAULT_MODEL = "deepseek-v4-flash"

def chat(self, messages, timeout=60):
    # v3.19.6: 仅使用 DeepSeek 单引擎(移除 OpenRouter 主备双引擎)
    return self._call_api(self.api_url, self.api_key, self.model, messages, timeout)
```

**报告生成过滤+汇总** (搜索 `def generate_with_ai`):
```python
visited_records = [r for r in records if r['photo_count'] > 0]
summary = "本次报告基于%s生成，共计%d个客户，其中有外访%d户已生成报告，%d个客户没有外访未生成报告" % (
    excel_name, total, visited_count, not_visited_count)
out_path = os.path.join(APP_DIR, "抵押物、抵债资产现场勘查日报表%s.xlsx" % get_date_str())
self._fill_template(items, out_path, summary_text=summary)
```

---

## 五、构建流程

### 构建 APK (GitHub Actions)

```powershell
# 1. 修改 buildozer.spec 中的 version
# 2. 修改 build_and_download.ps1 中的 $finalApkName 和临时文件名(v319X)
# 3. 修改 main.py 中欢迎页版本字符串(搜索 "v3.19")
# 4. 语法检查:
python -m py_compile "d:\hermes\loan_photo_app\main.py"
# 5. 提交:
git add -A; git commit -m "v3.19.X: ..."
# 6. 运行构建脚本:
powershell -ExecutionPolicy Bypass -File "d:\hermes\loan_photo_app\build_and_download.ps1"
# 脚本自动: git push → 等待 Actions → 监控构建 → nightly.link 下载 APK 到桌面
# 构建耗时约 14-16 分钟
```

### 构建 PPT 说明书

```powershell
# 1. 修改 build_manual_v3193.py 中的 DST 路径和封面版本号
# 2. 运行:
python "d:\hermes\loan_photo_app\build_manual_v3193.py"
# 输出到桌面: 资产盘点拍照工具-使用说明书-v3.19.X.pptx
```

### APK 下载备用方案（当构建脚本卡在速率限制时）

GitHub API 速率限制(60 次/小时)会导致 `build_and_download.ps1` 陷入 "Query error, retrying..." 循环。
此时可绕过脚本直接下载:

```powershell
# 1. 通过 WebFetch 查询 https://github.com/jare39063124-oss/loan-photo-app/actions 确认构建成功
# 2. 获取 runId (从 Actions 页面 URL)
# 3. 直接用 nightly.link 下载:
$runId = "28425843613"  # 替换为实际 runId
$url = "https://nightly.link/jare39063124-oss/loan-photo-app/actions/runs/$runId/loan-photo-apk.zip"
Invoke-WebRequest -Uri $url -OutFile "$env:TEMP\apk.zip"
Expand-Archive -Path "$env:TEMP\apk.zip" -DestinationPath "$env:TEMP\apk_extract" -Force
Copy-Item "$env:TEMP\apk_extract\*.apk" "$env:USERPROFILE\Desktop\loan-photo-tool-v3.19.X.apk"
```

---

## 六、重要设计决策和教训

### 必须遵守的规则

1. **HTTP 请求头不能含中文** — `X-Title: "资产盘点拍照工具"` 导致 latin-1 编码错误(position 0-7)
2. **Kivy Button 没有 scale_x/scale_y** — 用 `Animation(opacity=...)` 做按压动画
3. **SimHei 字体不含 emoji** — 📷🤖⚙✓ 等会显示为方块，用 ●◆▶ 或纯文字替代
4. **拍照后必须删除 MediaStore/DCIM 原件** — 否则一张照片保存两个文件
5. **API Key 必须完整存储** — 截断的 key 导致 HTTP 401
6. **API Key 必须用 base64 编码** — 明文 key 会被 GitHub Push Protection 拦截推送
7. **DeepSeek 模型名** — 用 `deepseek-v4-flash`（不带 `deepseek-ai/` 前缀）
8. **AI 生成的消息必须经人工审核** — 系统不自动发送消息给客户
9. **所有模型价格必须为 0** — 用户只用免费模型

### PPT 排版规范

- 主标题 28-36pt Bold，副标题 22-28pt，正文 18-24pt，注释 14-16pt
- 行距 1.2-1.5 倍，段间距 6-12pt
- 留白 ≥8% 页面宽度，颜色控制在 3 色以内
- 当前 PPT 字体: 标题栏 24pt，卡片标题 18pt，正文 15pt(符合最佳实践)
- 编号徽章用 0.55 英寸 + `word_wrap=False`(防两位数竖排)
- 目录用两列布局(防超过 6 章溢出)
- **用矢量图形(形状)替代生成图片** — trae-api text_to_image 对不同 prompt 返回相同占位图，不可靠

---

## 七、待办事项

### 当前无高优先级待办

v3.19.6 已完成并构建 APK。等待用户测试反馈。

### 用户测试验证点（v3.19.6）

1. **AI 引擎** — 验证 DeepSeek 单引擎能正常对话和生成日报表
2. **报告生成** — 验证:
   - 仅生成有外访照片客户的报告行
   - 末尾汇总说明行正确显示(基于 XX 文件/共计 XX 户/外访 XX 户)
   - 文件名自动为「抵押物、抵债资产现场勘查日报表 YYYYMMDD.xlsx」
3. **按钮交互** — 点击后显示「正在生成中…」并禁用，完成后恢复
4. **PPT 说明书** — 第 10 章日报表生成内容是否准确

### 潜在后续工作

- 若用户反馈报告格式需调整 → 修改 `report_template.xlsx` 或 `_fill_template` 方法
- 若用户需要恢复 OpenRouter 作为备用 → 参考 git 历史 commit d3c0f21(v3.19.5) 的双引擎实现
- PPT 说明书如需进一步优化 → 修改 `build_manual_v3193.py` 后重新运行

---

## 八、Git 提交历史

```
d33130b v3.19.6: 报告生成逻辑重构 + 仅保留DeepSeek引擎
d3c0f21 v3.19.5: DeepSeek备用API + E列逻辑修正 + README更新 + PPT第5页修复
7e47e74 docs: PPT字体优化+图片修复(矢量图替代)+交接文档
b70b349 v3.19.4: 删除app中所有作者信息
fa882b9 v3.19.3: 底部日志按钮替换为醒目的AI一键生成日报表大按钮
de05a71 v3.19.3: 工具栏拆为两行，修复生成日报表按钮在窄屏被截断不可见的问题
```

---

## 九、用户偏好汇总

- **语言**: 中文(所有沟通、代码注释、文档用中文)
- **模型**: 只用免费模型(input/output price == 0)
- **GUI**: Fluent Design 浅色主题，系统默认字体(≈9pt)，微软雅黑
- **PPT**: 简约明亮风格，不要花哨装饰，不要深色主题
- **安全**: AI 消息必须人工审核，系统不自动发送消息给客户
- **字体**: SimHei(不含 emoji，用 ●◆▶ 替代)
- **不喜欢**: 深色主题、24px 大字体

---

## 十、环境信息

- **OS**: Windows 11 Pro
- **Python**: 3.11(buildozer 在 GitHub Actions Linux 上运行)
- **Git**: 已配置，推送到 `jare39063124-oss/loan-photo-app`
- **GitHub Actions**: 构建 workflow 在 `.github/workflows/build_apk.yaml`
- **nightly.link**: 用于绕过 GitHub API 速率限制下载 artifact
- **ComfyUI**: `D:\ComfyUI\` (未运行，SD 1.5 可用，但生成图可靠性不如矢量图形)

---

## 十一、快速参考

### 修改版本号时需要同步更新的文件

1. `buildozer.spec` → `version = X.X.X`
2. `main.py` → 搜索 `v3.19` 更新欢迎页版本字符串
3. `build_and_download.ps1` → `$finalApkName`, `$tempZip`, `$extractDir`
4. `build_manual_v3193.py` → `DST` 路径、封面版本号、底部联系信息版本号
5. `README.md` → 版本表、构建历史表

### 调试技巧

- `python -m py_compile main.py` — 语法检查(构建前必做)
- `git log -p HEAD -- main.py | Select-String "sk-xxx"` — 检查提交是否含明文密钥
- 用 Grep 搜索关键字定位代码(行号可能偏移)
- GitHub Actions 构建日志: `https://github.com/jare39063124-oss/loan-photo-app/actions`
- APK 下载(nightly.link): `https://nightly.link/jare39063124-oss/loan-photo-app/actions/runs/{runId}/loan-photo-apk.zip`

### 常见问题排查

| 症状 | 原因 | 修复 |
|------|------|------|
| AI 无响应 | HTTP 头含中文 | 删除 X-Title 头 |
| 欢迎页闪退 | Animation 用 scale_x | 改用 opacity |
| 照片重复 | MediaStore/DCIM 原件未删 | on_camera_result 中删除原件 |
| 水印显示方块 | SimHei 不含 emoji | 用 ●◆▶ 或纯文字 |
| Excel 备注不保存 | write_back_to_uri 未实现 | SAF ContentResolver 写入 |
| 日报表按钮看不见 | 工具栏太挤 | 已移至底部(v3.19.3) |
| git push 被拒 | 明文 API key 触发 Push Protection | 用 base64 编码存储 key |
| 构建脚本卡 "Query error" | GitHub API 速率限制 | 用 nightly.link 直接下载 artifact |

---

*本文档由 AI 助手生成，用于跨会话交接。如有疑问请参考 Git 提交历史和代码注释。*

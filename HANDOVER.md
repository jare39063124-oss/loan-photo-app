# 资产盘点拍照工具 - 交接文档

> **交接日期**: 2026-06-30  
> **当前版本**: v3.19.4 (已提交,未构建APK)  
> **项目路径**: `d:\hermes\loan_photo_app\`  
> **Git 仓库**: `jare39063124-oss/loan-photo-app` (GitHub Actions 自动构建)  

---

## 一、项目概述

Android 资产盘点拍照工具,使用 **Kivy + Buildozer** 框架开发,通过 GitHub Actions 构建 APK。  
核心功能:Excel导入客户→分类拍照(远景/近景/内部/瑕疵/其他)→自动水印→AI生成日报表。

**用户**:抚顺银行风险管理部  
**用户偏好**:中文沟通、免费模型、Fluent Design 浅色主题、SimHei 字体(不含 emoji)

---

## 二、当前状态

### 已完成 (v3.19.0 → v3.19.4)

| 版本 | 内容 | 状态 |
|------|------|------|
| v3.19.0 | GUI重新设计(明亮配色)、水印增大2倍、修复重复照片、Excel备注修复 | ✅ 已构建 |
| v3.19.1 | 修复欢迎页点击闪退(scale_x/scale_y→opacity动画) | ✅ 已构建 |
| v3.19.2 | 作者改为抚顺银行风险管理部、AI latin-1修复、日报表功能、方块替换 | ✅ 已构建 |
| v3.19.3 | 工具栏拆两行→恢复单行、底部日志按钮替换为"AI一键生成日报表"大按钮 | ✅ 已构建 |
| v3.19.4 | **删除app中所有作者信息**(欢迎页作者卡片、设置页关于卡片、AUTHOR常量置空) | ⏳ 已提交,**未构建APK** |

### PPT 说明书状态

- **文件**: `C:\Users\Administrator\Desktop\资产盘点拍照工具-使用说明书-v3.19.3.pptx`
- **生成脚本**: `d:\hermes\loan_photo_app\build_manual_v3193.py`
- **14页, 63KB, 纯矢量图形(无外部图片)**
- v3.19.3 PPT 修复内容:
  - 目录页两列布局(修复11章后溢出)
  - 编号徽章 0.55英寸 + word_wrap=False(修复10/11/12竖排)
  - "适用场景"与"功能介绍"拆为独立卡片(修复重叠)
  - 配色改为白底+蓝色强调(简约明亮)
  - 3张损坏的API生成图替换为矢量图形(封面菱形装饰、水印位置示意图、日报表模拟表)
  - 字体层级:标题栏24pt、卡片标题18pt、正文15pt(符合PPT排版最佳实践)

---

## 三、关键文件清单

### 源码文件

| 文件 | 说明 |
|------|------|
| `d:\hermes\loan_photo_app\main.py` | **主程序(~4200行)**。包含所有UI、相机、水印、Excel、AI、日报表逻辑 |
| `d:\hermes\loan_photo_app\buildozer.spec` | Buildozer 配置(version=3.19.4, source.include_exts 含 xlsx) |
| `d:\hermes\loan_photo_app\build_and_download.ps1` | 构建+下载脚本(推送GitHub→监控Actions→nightly.link下载APK) |
| `d:\hermes\loan_photo_app\report_template.xlsx` | 内置日报表模板(12393字节,6列,Sheet1) |
| `d:\hermes\loan_photo_app\res\extra_manifest_application.xml` | FileProvider 声明(修复拍照质量的关键) |
| `d:\hermes\loan_photo_app\build_manual_v3193.py` | PPT说明书生成脚本 |
| `d:\hermes\loan_photo_app\HANDOVER.md` | 本交接文档 |

### 桌面文件

| 文件 | 说明 |
|------|------|
| `C:\Users\Administrator\Desktop\资产盘点拍照工具-使用说明书-v3.19.3.pptx` | 最新PPT说明书(14页) |
| `C:\Users\Administrator\Desktop\loan-photo-tool-v3.19.3.apk` | 最新APK(v3.19.3, 28.57MB) |
| `C:\Users\Administrator\Desktop\盘点相关文件\` | 原始模板和参考文件目录 |

---

## 四、main.py 关键代码位置

> **注意**: main.py 约4200行,行号可能因编辑略有偏移,建议用 Grep 搜索关键字定位。

### 常量和配置 (行 400-430)

```python
AUTHOR_NAME = ""          # v3.19.4 已置空
AUTHOR_PHONE = ""
AUTHOR_INFO = ""
THEME = { ... }           # 明亮浅色主题 (行417-430)
WATERMARK_FONT_SIZE_MAP = {"大": 80, "中": 56, "小": 36}  # 行403
```

### 核心类

| 类 | 位置(约) | 说明 |
|----|----------|------|
| `PhotoProcessor` | 行~780 | 水印添加(PIL), `add_watermark` 方法 |
| `ExcelWriter` | 行~780 | Excel读写, `write_back_to_uri` (SAF写入) |
| `ReportGenerator` | 行~801 | AI日报表生成, `generate_with_ai` 方法 |
| `CameraManager` | 行~1700 | 相机管理,FileProvider+MediaStore+DCIM三策略 |
| `AIService` | 行~1400 | OpenRouter API调用,**无X-Title头**(latin-1修复) |
| `MainScreen` | 行~2200 | 主界面,所有UI和事件处理 |

### 关键方法

| 方法 | 搜索关键字 | 说明 |
|------|-----------|------|
| `bind_press_animation` | `def bind_press_animation` | 按钮动画(**用opacity,不能用scale_x**) |
| `_generate_report` | `def _generate_report` | 日报表生成(线程+进度弹窗+SAF保存) |
| `_save_report_to_user` | `def _save_report_to_user` | ACTION_CREATE_DOCUMENT 保存日报表 |
| `on_activity_result` | `def on_activity_result` | 处理_report_save_code=0x201 |
| `add_watermark` | `def add_watermark` | PIL水印(min_font=24, padding=24) |
| `on_camera_result` | `def on_camera_result` | 三策略拍照(删除MediaStore/DCIM原件防重复) |

---

## 五、构建流程

### 构建 APK (GitHub Actions)

```powershell
# 1. 修改 buildozer.spec 中的 version
# 2. 修改 build_and_download.ps1 中的 $finalApkName 和临时文件名
# 3. 运行:
powershell -ExecutionPolicy Bypass -File "d:\hermes\loan_photo_app\build_and_download.ps1"
# 脚本自动: git push → 等待Actions → 监控构建 → nightly.link下载APK到桌面
# 构建耗时约16分钟, APK输出到桌面
```

### 构建 PPT 说明书

```powershell
python "d:\hermes\loan_photo_app\build_manual_v3193.py"
# 输出到桌面: 资产盘点拍照工具-使用说明书-v3.19.3.pptx
```

---

## 六、重要设计决策和教训

### 必须遵守的规则

1. **HTTP请求头不能含中文** — `X-Title: "资产盘点拍照工具"` 导致 latin-1 编码错误(position 0-7)
2. **Kivy Button 没有 scale_x/scale_y** — 用 `Animation(opacity=...)` 做按压动画
3. **SimHei 字体不含 emoji** — 📷🤖⚙✓ 等会显示为方块,用 ●◆▶ 或纯文字替代
4. **拍照后必须删除 MediaStore/DCIM 原件** — 否则一张照片保存两个文件
5. **API Key 必须完整存储** — 截断的 key 导致 HTTP 401
6. **AI生成的消息必须经人工审核** — 系统不自动发送消息给客户
7. **所有模型价格必须为0** — 用户只用免费模型

### v3.19.4 最新改动

- `AUTHOR_NAME`, `AUTHOR_PHONE`, `AUTHOR_INFO` 全部置空
- 欢迎页移除了 `author_card` (原显示 AUTHOR_INFO)
- 设置页移除了 `about_card` (原显示 AUTHOR_INFO 和 ℹ️ 图标)

---

## 七、待办事项

### 高优先级

1. **构建 v3.19.4 APK** — 作者信息删除已提交(commit b70b349),但尚未构建APK
   ```powershell
   # 修改 build_and_download.ps1: v3.19.3 → v3.19.4
   # 然后运行构建脚本
   ```

2. **用户测试 v3.19.4** — 验证:
   - 作者信息是否已从欢迎页和设置页消失
   - 底部"AI一键生成日报表"按钮是否正常工作
   - AI助手是否能正常对话(latin-1修复后)
   - 日报表生成功能是否完整

### 中优先级

3. **PPT说明书进一步优化** — 如果用户对字体排版仍不满意:
   - 检查 `build_manual_v3193.py` 中的字号设置
   - 参考搜索到的 PPT 排版规范:
     - 主标题 28-36pt Bold, 副标题 22-28pt, 正文 18-24pt, 注释 14-16pt
     - 行距 1.2-1.5倍, 段间距 6-12pt
     - 留白 ≥8% 页面宽度
     - 颜色控制在3色以内
   - 当前PPT字体: 标题栏24pt, 卡片标题18pt, 正文15pt(已在最佳实践范围内)

4. **ComfyUI图片生成** — ComfyUI安装在 `D:\ComfyUI\` 但未运行:
   - 可用模型: SD 1.5 (`v1-5-pruned-emaonly.safetensors` 在子目录中)
   - 需要用户手动启动 ComfyUI 或将checkpoint移到根目录
   - 也可使用内置API: `https://trae-api-cn.mchost.guru/api/ide/v1/text_to_image` (但对3个不同prompt返回了相同图片,可靠性存疑)

---

## 八、Git 提交历史

```
b70b349 v3.19.4: 删除app中所有作者信息
fa882b9 v3.19.3: 底部日志按钮替换为醒目的AI一键生成日报表大按钮
de05a71 v3.19.3: 工具栏拆为两行，修复生成日报表按钮在窄屏被截断不可见的问题
(v3.19.2 commits...)
```

---

## 九、用户偏好汇总

- **语言**: 中文(所有沟通、代码注释、文档用中文)
- **模型**: 只用免费模型(input/output price == 0)
- **GUI**: Fluent Design 浅色主题,系统默认字体(≈9pt),微软雅黑
- **PPT**: 简约明亮风格,不要花哨装饰
- **安全**: AI消息必须人工审核,系统不自动发送消息给客户
- **字体**: SimHei(不含emoji,用●◆▶替代)

---

## 十、环境信息

- **OS**: Windows 11 Pro
- **Python**: 需要确认版本(buildozer在Linux/WSL上运行,GitHub Actions构建)
- **Git**: 已配置,推送到 `jare39063124-oss/loan-photo-app`
- **ComfyUI**: `D:\ComfyUI\` (未运行,SD 1.5可用)
- **GitHub Actions**: 构建workflow在 `.github/workflows/` 目录
- **nightly.link**: 用于绕过GitHub API速率限制下载artifact

---

## 十一、快速参考

### 修改版本号时需要同步更新的文件

1. `buildozer.spec` → `version = X.X.X`
2. `main.py` → 搜索 `v3.19` 更新版本字符串(欢迎页)
3. `build_and_download.ps1` → `$finalApkName`, `$tempZip`, `$extractDir`, 输出消息

### 调试技巧

- `python -m py_compile main.py` — 语法检查(构建前必做)
- 用 Grep 搜索关键字定位代码(行号可能偏移)
- GitHub Actions 构建日志: `https://github.com/jare39063124-oss/loan-photo-app/actions`
- APK下载(nightly.link): `https://nightly.link/jare39063124-oss/loan-photo-app/actions/runs/{runId}/loan-photo-apk.zip`

### 常见问题排查

| 症状 | 原因 | 修复 |
|------|------|------|
| AI无响应 | HTTP头含中文 | 删除X-Title头 |
| 欢迎页闪退 | Animation用scale_x | 改用opacity |
| 照片重复 | MediaStore/DCIM原件未删 | on_camera_result中删除原件 |
| 水印显示方块 | SimHei不含emoji | 用●◆▶或纯文字 |
| Excel备注不保存 | write_back_to_uri未实现 | SAF ContentResolver写入 |
| 日报表按钮看不见 | 工具栏太挤 | 拆两行或放底部 |

---

*本文档由AI助手生成,用于跨会话交接。如有疑问请参考Git提交历史和代码注释。*

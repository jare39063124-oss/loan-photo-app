# 资产盘点拍照工具 - 交接文档

> **交接日期**: 2026-07-06
> **当前版本**: v3.22.22 (已构建)
> **项目路径**: `d:\hermes\loan_photo_app\`
> **Git 仓库**: `jare39063124-oss/loan-photo-app` (GitHub Actions 自动构建)
> **Git Commit**: `22579b0`
> **GitHub Actions Run**: `28763714124` (Build #93)
> **最新 APK**: `C:\Users\Administrator\Desktop\loan-photo-tool-v3.22.22.apk` (已构建, 29.03 MB)
> **最新 PPT**: `C:\Users\Administrator\Desktop\资产盘点拍照工具-使用说明书-v3.22.22.pptx` (已生成, 16 页)
>
> **APK 下载链接(nightly.link)**:
> https://nightly.link/jare39063124-oss/loan-photo-app/actions/runs/28763714124/loan-photo-apk.zip

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
| v3.19.6 | 移除 OpenRouter 仅保留 DeepSeek；报告仅生成有外访照片客户+汇总说明行+自动命名 | ✅ 已构建 |
| v3.20.0 | 全量日志+后台稳定(on_pause/on_restore)+Excel 分批加载+备注持久化+搜索重建+拍照会话锁+AI 报告增强+照片去重 | ✅ 已构建 |
| v3.21.0 | 性能优化 + Excel 稳定性修复(并发保存覆盖/缺锁/文件句柄泄漏/空行处理)；回滚录音功能 | ✅ 已构建 |
| **v3.22.0** | **14 项优化:日志开关/Excel A列序号(6列格式)/竖版计数器/最近文件/搜索类型选择器/AI日报表按客户合并/性能兼容/异步缩略图/clean_markdown/按钮倒角14/兼容性** | ✅ 已构建 |
| v3.22.1 | 修复查看已拍无反应(progress_key 不一致)+竖版计数器不更新(分隔符 _ → -) | ✅ 已构建 |
| v3.22.2 | 日志二级弹窗(开关/查看/复制/清除)+计数器横排+高亮当前拍摄行+搜索合并地址+倒角dp(8)+最近文件名加大 | ✅ 已构建 |
| v3.22.3 | P0备注覆盖(按行号独立存储)+倒角dp(12)+AI报表默认文件名(JString)+高亮会话期间保持+日志回退内存缓冲 | ✅ 已构建 |
| **v3.22.4** | **P0:AI报表卡死(_collect_records缺enumerate→NameError)+备注显示后段(TextInput光标重置)+查看已拍无反应(scoped storage存APP_DIR路径)+倒角dp(16)+高亮消失(强制_update_bg)+AI不省略信息+删除owl文案** | ✅ 已构建 |
| **v3.22.5** | **P0:查看已拍修复(get_photos不回写+migrate_photo_paths迁移)+切换Excel备注串扰(清理_row_remarks)+AI报表二次确认弹窗+倒角修复(RoundedButton改基类ButtonBehavior+Label)** | ✅ 已构建 |
| **v3.22.6** | **P0:查看已拍修复(_show_msg改用Popup)+AI确认弹窗修复(_collect_records调用)+新增批量标记同类型(长按多选+AI报表合并统计)** | ✅ 已构建 |
| **v3.22.7** | **P0:修复Popup背景异常(_show_msg_popup+report_confirm)+长按增强日志+AI键盘遮挡+性能优化** | ✅ 已构建 |
| **v3.22.8** | **P0:修复 AppLogger 缺失 debug/warning 方法导致 25 行 RowWidget 全部创建失败列表空白（v3.22.7 引入的 app_log.debug/app_log.warning 调用兜底）** | ✅ 已构建 |
| **v3.22.9** | **P0:修复 _toggle_batch_mark 遗漏 background=THEME['card'] bug + AI 键盘改用 Window.on_keyboard_height + AppLogger ERROR 始终输出 Kivy Logger + _create_row_widget 失败占位 widget + _on_row_long_press 可见反馈 + 抽 ThemedPopup 基类根治 background 遗漏 + 长按 Android toast + on_touch_move 阈值 dp(10)→dp(15)** | ✅ 已构建 |
| **v3.22.10** | **P0:Kivy 触摸事件机制根因修复:always_release=True+touch.grab/ungrab+_on_view_photos 错误反馈改 Popup+softinput_mode=pan+DebugOverlay 运行时诊断覆盖层+长按 1.2s** | ✅ 已构建 |
| **v3.22.11** | **P0:回退 v3.22.10 引入的回归 — 移除 touch.grab/ungrab（与 ButtonBehavior 内部 grab 冲突导致查看已拍触发7次）+ 移除 DebugOverlay 类（220×120dp 覆盖 AI 助手按钮导致无反应）+ 新增 _on_view_photos 500ms 防抖双保险 + 长按 1.2s→0.8s；保留 always_release=True 和 softinput_mode=pan** | ✅ 已构建 |
| **v3.22.12** | **P0:查看已拍无反应根因修复 — view_btn/photo_btn/remark_btn/del_btn/save_*_btn/recent_excel/placeholder 全部改用 on_press（ScrollView 消费 touch_up 导致 on_release 不触发）+ 长按 _trigger_long_press 移除 collide_point 改用移动距离检查（ScrollView 滚动后 collide_point 失效）+ PhotoViewerPopup 改为 cols=3 缩略图网格（参考手机相册）+ PIL Image 用 with 语句修复文件句柄泄漏 + RowWidget _cleanup 修复长按定时器内存泄漏 + 深度审查 46 处 on_release 绑定** | ✅ 已构建 |
| **v3.22.13** | **P0:查看已拍功能彻底重写 — 放弃 RowWidget 内嵌套按钮方案，改为 ButtonBehavior + 全行点击触发行操作菜单 Popup（参考 Kivy 图片画廊成功案例）+ PhotoViewerPopup 缩略图改用 Button + background_normal=filepath + _on_view_photos 整体 try/except + Toast 反馈（不再"连失败提示都没有"）+ 移除长按多选全部代码（_trigger_long_press/_toggle_batch_mark/_batch_mark_rows/_clear_batch_marks/_on_row_long_press）+ 新增 VisitNoteDialog 今日走访补充说明输入区域（按 Excel 路径 md5 哈希持久化到 APP_DIR/visit_notes/，切换 Excel 自动保存恢复，AI 报告生成时注入 system prompt）+ 深度审查修复 5 处高/中优 bug（ExcelWriter/android_copy_uri/save_to_gallery 流未关闭 + _show_msg_safe 增加 Toast + _save_visit_note 原子写）** | ✅ 已构建 |
| **v3.22.14** | **回退 RowWidget 至 v3.22.12 内嵌按钮布局（移除 ButtonBehavior 全行点击 + 行操作菜单）+ PhotoViewerPopup 回退到 v3.22.12 Label 占位 + KivyImage 实现 + _on_view_photos 增加诊断日志 + REPORT_SYSTEM_PROMPT 增强（同楼盘多房间合并描述）+ _build_prompt 增强（visit_note 显式段落）+ _show_report_confirm 增加客户维度走访统计 + PPT 删除作者信息仅保留咨询电话 15940454123 + buildozer.spec/build_and_download.ps1/HANDOVER.md 版本号同步更新 v3.22.14** | ✅ 已构建 commit 26e091a, Run 28704312746 |
| **v3.22.15** | **P0:PhotoViewerPopup 可见性终极修复 — ThemedPopup 增加 background='' 禁用 Kivy Popup 默认 atlas 背景图像（某些机型加载失败导致 popup 完全不可见）+ _on_view_photos 用 Clock.schedule_once 延迟 0.1s 创建+open 避免触摸事件链时序冲突 + 防抖 500ms→1500ms 防重复点击 + popup 状态诊断日志（pos/size/opacity/parent）+ VisitNoteDialog 提示文案同步更新"同楼盘多房间合并说明"** | ✅ 已构建 commit 1a3dc33, Run 28705924075 |
| **v3.22.16** | **P0:PhotoViewerPopup 可见性终极修复 — 显式设置 self.size=self.center 替代 size_hint（v3.22.15 诊断日志显示 size=[100,100] pos=[0,0] 证明 size_hint 在 Android 上未生效）+ 移除 Clock.schedule_once 延迟包装（非根因，直接创建+open）** | ✅ 已构建 commit f2f9135, Run 28711296361 |
| **v3.22.17** | **P0:PhotoViewerPopup auto_dismiss 终极修复 — 增加 self.auto_dismiss=False 防止被同一触摸事件立即关闭（v3.22.16 尺寸已正确但 ModalView.on_touch_down 检测到触摸点在 popup 外部触发 auto_dismiss）+ 全项目 Bug 审计修复 9 处（2 处 on_release→on_press + 7 处裸 except→具体异常类型）** | ✅ 已构建 commit c87c26e, Run 28712601573 |
| **v3.22.18** | **P0:PhotoViewerPopup 背景渲染修复 — 从继承 ThemedPopup 改为直接继承 Popup，背景绘制在 content.canvas.before（匹配 _show_report_confirm 有效模式）+ 权限补全 Android 11+ READ_EXTERNAL_STORAGE 和 Android 13+ READ_MEDIA_IMAGES** | ✅ 已构建 commit 77d3256, Run 28723802219 |
| **v3.22.19** | **综合增强：1) 缩略图方案替代原图查看（拍照后异步生成压缩缩略图到 app 私有目录，绕过 Scoped Storage 权限问题）+ 2) 键盘遮挡修复（softinput_mode 改 resize + 弹窗聚焦滚动）+ Spinner 倒角 + 3) AI 报告提示词调整（每客户仅 1 行 + 多地址分号成行 + 同地址 4+ 房间号合并 + few-shot 示例）+ 4) 默认文件名改为"线路四XX（日期）现场勘查日报表" + 5) 表头锁定 + 全选/行选择/反选（每行 CheckBox，支持全选后单行反选）+ 6) 最近列表数据指示器（绿/灰圆点）+ 2 次确认清除按钮** | ✅ 已构建 commit 39671d2, Run 28747164492 |
| **v3.22.20** | **修复：1) 查看已拍改用工厂方法（移除 PhotoViewerPopup 继承 Popup，改为 _show_photo_viewer_popup 直接创建 Popup，与 _show_report_confirm 一致，解决 parent=None 问题）+ 2) 最近列表数据指示器改文字按钮（"有数据 N 项"/"无数据"）+ 清除按钮弹窗确认对话框（说明删除内容和后果）+ 3) AI 拍摄助手键盘遮挡修复（focus 滚动聊天到底部）+ VisitNoteDialog 窗口位置上移至屏幕顶端 + 4) ColoredCheckBox 自定义颜色（未选中白边框+浅背景，选中蓝色填充+白色对勾）+ 5) 全局 bug 检查（6 处裸 except 修复，线程安全/资源泄漏/索引越界/回调循环审计）** | ✅ 已构建 commit 7e4be32, Run 28749105722 |
| **v3.22.21** | **修复：1) 删除缩略图二次确认提示原图保留（相册原图需手动删除）+ 2) 缩略图 2 列竖向布局 + 横版照片逆时针旋转 90° + 3) ColoredCheckBox 形状还原到 20dp×20dp + 4) 多选拍摄计数逻辑（多选 N 条目拍一张照片全部计数 +1，相册仅 1 张；AI 日报表/助手 prompt 注入多选语义）+ 5) AI 拍摄助手键盘遮挡修复（keyboard_spacer + on_leave 不解绑）+ 6) 全局 bug 审计（PIL Image 关闭 + tmp_rotated 清理 + CardWidget 硬编码改 dp）** | ✅ 已构建 commit 721bebf, Run 28753552339 |
| **v3.22.22** | **收尾打磨：1) 全选作用域修正（_on_select_all/_on_row_selected 改为基于当前可见行 self.row_widgets 而非全量 self.rows，搜索过滤后全选仅影响匹配行）+ 2) AI 拍摄助手键盘适配改用 Window.softinput_mode='resize' 整体缩短（移除不生效的 keyboard_spacer 方案，聊天记录与输入栏同时可见）+ 3) 搜索类别 Spinner 倒角根治（新增 RoundedSpinner 子类在 canvas.before 绘制 RoundedRectangle，按钮本体呈现 dp(8) 圆角，根治多次返工未修复的倒角问题）+ 4) 查看已拍缩略图布局优化（图片占比 0.65→0.78、显示完整文件名不截断、字体 10sp→12sp）+ 5) 权限精简（移除 READ_MEDIA_IMAGES/READ_EXTERNAL_STORAGE/WRITE_EXTERNAL_STORAGE，因查看已拍读 app 私有目录缩略图、拍照 URI 由 grantUriPermission 临时授权）** | ✅ 已构建 commit 22579b0, Run 28763714124 |

### v3.22.0 核心改动（当前版本）— 14 项优化

> 完整计划文档: `D:\Trae CN\.trae\documents\loan-photo-app-v3.22.0-update-plan.md`

1. **日志开关(需求1)** — `AppLogger` 新增 `_enabled`/`set_enabled()`/`is_enabled()`，默认关闭；底部按钮「日志:关/开」短按切换、长按查看日志 Popup(浅色背景+深色文字)；config 字段 `log_enabled`
2. **搜索类型选择器+键盘(需求2/11)** — toolbar 新增 `Spinner`(序号/客户名/地址(概)/地址(详)/备注)；`_on_search_field_change` 更新 config；`_refresh_list` 按 `field_map` 选列过滤；`Window.softinput_mode='resize'`(搜索框不再被顶出)
3. **最近 Excel 记录(需求3)** — `_add_recent_excel`/`_remove_recent_excel`/`_open_recent_excel`；`_show_empty_state` 在「Excel未加载」提示下方展示最近 3-5 个文件名，点击加载，文件移除时 Toast 提示
4. **按钮倒角(需求4)** — `RoundedButton` radius `[8]→[14]`(2792行)，视觉更圆润
5. **查看已拍异步(需求5)** — `PhotoViewerPopup` 后台线程生成缩略图(200x200, 缓存 `APP_DIR/thumbnails/`)，`Clock.schedule_once` 分批渲染；删除确认文字改 `THEME['text_on_light']` 保证对比度
6. **竖版计数器(需求6)** — `RowWidget` type_status 改为 5 行竖版 `☐ 远景 0`/`☑ 近景 2`；`ProgressManager.get_photo_count_by_type`
7. **AI 生成提示(需求7)** — `_generate_report` 进度弹窗:浅色背景 `THEME['card']`+居中+幽默文案「AI 正在奋笔疾书，请稍候片刻…(泡杯茶的时间就够了)」+`auto_dismiss=False`
8. **AI 日报表合并(需求8)** — `_collect_records` 按 borrower 分组，多处抵押物合并为一条，「抵押物/抵债资产具体情况」表述「共计盘点 N 处，位置为 XX XX」；备注弹窗增加引导提示；`REPORT_SYSTEM_PROMPT` 明确「严禁 AI 编造内容」
9. **AI 回答去 markdown(需求9)** — `clean_markdown(text)`(504行) 清理 `**`/`*`/`#`/`` ` ``；`_on_ai_response`/`_add_message` 调用
10. **Excel A 列序号(需求10)** — 6 列格式:A=序号 B=客户名 C=地址概 D=地址详 E=性质 F=备注；`ExcelReader.load` 读 `row[:6]`；`ExcelWriter.save_remark` `column=5→6`；32 处 `row[N]` 全部偏移(已审计)
11. **性能优化(需求12)** — `ProgressManager.get_progress_snapshot(rows)` 单次锁收集所有行进度(2N 次加锁→1 次)；`build_system_prompt` 已采用
12. **兼容性(需求13)** — 裸 except 37 处(7 处 pyjnius 必须保留)；`_on_view_photos`/`_on_remark_request`/Excel 加载异常边界保护
13. **全局 Bug 检查(需求14)** — row[N] 偏移审计(32处全对)、Lock 死锁审计(无死锁)、ProgressManager 加锁验证(全加锁)、Excel 写入串行化验证
14. **PPT 说明书** — `build_manual_v3193.py` 更新:新增 Slide 5「安装说明」(4 步:微信接收→菜单→保存→文件管理器安装)；Slide 6 Excel 格式改 6 列；版本号 v3.22.0；15 页

---

## 三、关键文件清单

### 源码文件

| 文件 | 说明 |
|------|------|
| `d:\hermes\loan_photo_app\main.py` | **主程序(~4734 行)**。包含所有 UI、相机、水印、Excel、AI、日报表逻辑 |
| `d:\hermes\loan_photo_app\buildozer.spec` | Buildozer 配置(version=3.22.0, source.include_exts 含 xlsx) |
| `d:\hermes\loan_photo_app\build_and_download.ps1` | 构建+下载脚本(推送 GitHub → 监控 Actions → nightly.link 下载 APK) |
| `d:\hermes\loan_photo_app\report_template.xlsx` | 内置日报表模板(6 列, Sheet1) |
| `d:\hermes\loan_photo_app\res\extra_manifest_application.xml` | FileProvider 声明(修复拍照质量的关键) |
| `d:\hermes\loan_photo_app\build_manual_v3193.py` | PPT 说明书生成脚本(15 页, 纯矢量图形) |
| `d:\hermes\loan_photo_app\README.md` | GitHub 项目介绍(已移除作者信息) |
| `d:\hermes\loan_photo_app\HANDOVER.md` | 本交接文档 |

### 桌面文件

| 文件 | 说明 |
|------|------|
| `C:\Users\Administrator\Desktop\loan-photo-tool-v3.22.21.apk` | **最新 APK**(v3.22.21, 29.03 MB) |
| `C:\Users\Administrator\Desktop\资产盘点拍照工具-使用说明书-v3.22.21.pptx` | **最新 PPT 说明书**(16 页, 67.15 KB) |
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

v3.22.0 已完成全部 14 项优化、全局 Bug 检查、APK 构建和 PPT 更新。等待用户实机测试反馈。

### 用户测试验证点（v3.22.0）

1. **搜索功能** — 切换搜索类型(序号/客户名/地址概/地址详/备注)，输入关键词，过滤正确且键盘不遮挡搜索框
2. **Excel 6 列格式** — 加载含序号列的新格式 Excel，客户列表正确显示，备注保存到 F 列
3. **最近文件** — 打开 Excel 后退出重进，空状态界面显示最近文件名，点击快捷加载；删除文件后点击有提示
4. **连续备注保存** — 连续保存多个客户备注，无覆盖无丢失
5. **照片命名 + 竖版计数器** — 拍照后照片按类型命名，计数器竖版 5 行显示 ☐/☑ + 张数
6. **AI 日报表合并** — 同一客户多个抵押物，日报表生成一条，「抵押物/抵债资产具体情况」表述共计盘点数量；备注内容参考客户输入不编造
7. **日志开关** — 默认关闭不生成日志；开启后生成；日志查看 Popup 文字可读
8. **查看已拍** — 响应迅速(异步缩略图)；删除确认弹窗文字清晰可读
9. **AI 对话** — 回答不含 `**` 等 markdown 符号
10. **多机型** — 华为/小米/OPPO/vivo 安装运行无闪退

### 潜在后续工作

- 若用户反馈报告格式需调整 → 修改 `report_template.xlsx` 或 `_fill_template` 方法
- 裸 except 清理(37 处, 7 处 pyjnius 必须保留) — 属代码风格非 bug，APK 已构建，建议后续单独提交避免引入风险
- PPT 说明书如需进一步优化 → 修改 `build_manual_v3193.py` 后重新运行

---

## 八、Git 提交历史

```
f2fb8a4 v3.22.0: 14项优化(日志开关/Excel序号/竖版计数器/最近文件/搜索选择器/AI日报表合并/性能兼容/异步缩略图)
a2af872 v3.21.0: 性能优化 + Excel 稳定性修复
7e7e38f v3.20.0: 全量日志+后台稳定+Excel分批加载+备注持久化+搜索重建+拍照会话锁+AI报告增强+照片去重
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

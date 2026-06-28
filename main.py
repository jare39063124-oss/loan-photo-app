"""
资产盘点专项拍照工具 App - v3.9.0
功能：
- 欢迎页 + 设置页
- 文件命名自选模式（4段下拉 X-X-X-X）
- 水印自选模式（3段下拉 X-X-X + 大中小字号 + 位置）
- 四类拍照引导（远景/近景/内部/瑕疵）+ 连续拍照（同类型手动按快门连拍）
- Excel A=客户名 B=抵押物地址（概） C=抵押物地址（精确门牌号） D=抵押物性质
- 进度持久化（跨Excel文件可识别已拍条目）
- 照片路径验证（缩略图不会引用失效路径）
- 崩溃日志收集
"""

import os
import json
import sys
import re
import traceback
from datetime import datetime
from collections import defaultdict
import threading
import hashlib

# ============================================================
# 崩溃日志捕获
# 注意：此处不能依赖 kivy.utils.platform，因为 kivy 尚未导入。
# 用环境变量做早期 Android 检测（python-for-android 会注入 ANDROID_ARGUMENT）。
# ============================================================
CRASH_LOG_FILE = None
_EARLY_IS_ANDROID = ('ANDROID_ARGUMENT' in os.environ) or (sys.platform == 'android')

def _resolve_android_app_dir():
    """返回 Android 上可写的 app 专属外部存储路径（无需权限、不受 scoped storage 限制）。

    路径 = /storage/emulated/0/Android/data/<package>/files/loan_photos
    用户可通过文件管理器访问：内部存储 → Android → data → <package> → files → loan_photos

    所有 Android 版本均无需运行时权限；Android 10+ scoped storage 下也可写。
    """
    if not _EARLY_IS_ANDROID:
        return os.path.join(os.path.expanduser('~'), 'loan_photos')

    # 1) 首选：App-specific external storage（无需权限）
    try:
        from jnius import autoclass
        PythonActivity = autoclass('org.kivy.android.PythonActivity')
        activity = PythonActivity.mActivity
        if activity is not None:
            ext_files_dir = activity.getExternalFilesDir(None)
            if ext_files_dir is not None:
                base = str(ext_files_dir.getAbsolutePath())
                if base and base.lower() != 'null':
                    return os.path.join(base, 'loan_photos')
    except Exception:
        pass

    # 2) 退路1：python-for-android 私有内部存储（一定可写，但用户在文件管理器看不到）
    try:
        home = os.path.expanduser('~')
        if home and os.path.isdir(home):
            return os.path.join(home, 'loan_photos')
    except Exception:
        pass

    # 3) 退路2：临时目录（最差兜底）
    import tempfile
    return os.path.join(tempfile.gettempdir(), 'loan_photos')

def setup_crash_handler():
    global CRASH_LOG_FILE
    try:
        log_dir = _resolve_android_app_dir()
        os.makedirs(log_dir, exist_ok=True)
        CRASH_LOG_FILE = os.path.join(log_dir, 'crash_log.txt')
    except Exception:
        # 最后兜底：写到当前目录（APK 内只读会失败，但不会抛出）
        try:
            log_dir = os.path.join(os.path.expanduser('~'), 'loan_photos')
            os.makedirs(log_dir, exist_ok=True)
            CRASH_LOG_FILE = os.path.join(log_dir, 'crash_log.txt')
        except Exception:
            CRASH_LOG_FILE = None

    def global_excepthook(exc_type, exc_value, exc_tb):
        msg = "="*50 + "\n"
        msg += "CRASH at %s\n" % datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        msg += "Type: %s\n" % exc_type.__name__
        msg += "Error: %s\n" % str(exc_value)
        msg += "Traceback:\n"
        msg += "".join(traceback.format_tb(exc_tb))
        msg += "\n"
        if CRASH_LOG_FILE:
            try:
                with open(CRASH_LOG_FILE, 'a', encoding='utf-8') as f:
                    f.write(msg)
            except Exception:
                pass
        # Also let original handler run
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    sys.excepthook = global_excepthook
    # Also capture Python thread exceptions
    threading.excepthook = lambda args: global_excepthook(args.exc_type, args.exc_value, args.exc_tb)

setup_crash_handler()

os.environ['KIVY_LOG_LEVEL'] = 'warning'

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.image import Image as KivyImage
from kivy.uix.checkbox import CheckBox
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.slider import Slider
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.clock import Clock, mainthread
from kivy.logger import Logger
from kivy.core.window import Window
from kivy.utils import platform
from kivy.graphics import Color, Rectangle, Line, RoundedRectangle
from kivy.metrics import dp, sp
from kivy.storage.jsonstore import JsonStore
from kivy.core.text import LabelBase
from kivy.resources import resource_find, resource_add_path
from kivy.lang import Builder
from kivy.uix.spinner import SpinnerOption

# ============================================================
# 中文字体加载（彻底解决中文豆腐块问题）
#
# 根本原因历史：
#   - 之前用的 NotoSansSC.ttf 只有 28KB，是损坏的占位文件，
#     一个真正的 CJK 中文字体至少需要 4-16MB。
#   - 现在使用 simhei.ttf（黑体，9.7MB），放在 main.py 同级目录，
#     buildozer 通过 source.include_exts=ttf 自动打包进 APK，
#     在 APK 内与 main.py 同目录，路径最简单最可靠。
#
# 三重保障：
#   1. LabelBase.register('Roboto', ...) 替换 Kivy 默认字体
#   2. KV 全局规则设置所有文本 widget 的 font_name，
#      特别包含 SpinnerOption（下拉菜单选项）
#   3. Android 系统字体 fallback（应对极端情况）
# ============================================================
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_ASSETS_DIR = os.path.join(_APP_DIR, 'assets')
resource_add_path(_APP_DIR)
if os.path.isdir(_ASSETS_DIR):
    resource_add_path(_ASSETS_DIR)

_FONT_PATH = None
for _fn in ['simhei.ttf', 'NotoSansSC.ttf', 'NotoSansSC-Regular.otf']:
    _p = os.path.join(_APP_DIR, _fn)
    if os.path.exists(_p) and os.path.getsize(_p) > 100000:
        _FONT_PATH = _p
        break
if not _FONT_PATH:
    for _fn in ['simhei.ttf']:
        _found = resource_find(_fn)
        if _found and os.path.exists(_found) and os.path.getsize(_found) > 100000:
            _FONT_PATH = _found
            break
if not _FONT_PATH:
    for _sp in [
        '/system/fonts/NotoSansCJK-Regular.ttc',
        '/system/fonts/NotoSansSC-Regular.otf',
        '/system/fonts/NotoSansCJKsc-Regular.otf',
        '/system/fonts/DroidSansFallback.ttf',
        '/system/fonts/SourceHanSansCN-Regular.otf',
        '/system/fonts/msyh.ttc',
    ]:
        if os.path.exists(_sp):
            _FONT_PATH = _sp
            break

if _FONT_PATH:
    try:
        LabelBase.register('Roboto', _FONT_PATH)
    except Exception as _e:
        _FONT_PATH = None

CHINESE_FONT = 'Roboto'

Builder.load_string('''
<Label>:
    font_name: 'Roboto'
<Button>:
    font_name: 'Roboto'
<Spinner>:
    font_name: 'Roboto'
<TextInput>:
    font_name: 'Roboto'
<CheckBox>:
    font_name: 'Roboto'
<SpinnerOption>:
    font_name: 'Roboto'
<Switch>:
    font_name: 'Roboto'
<Slider>:
    font_name: 'Roboto'
<ProgressBar>:
    font_name: 'Roboto'
''')

_diag_lines = ["=== Font diag ===", "__file__=%s" % __file__,
               "app_dir=%s" % _APP_DIR,
               "font_path=%s" % _FONT_PATH,
               "font_size=%d" % (os.path.getsize(_FONT_PATH) if _FONT_PATH and os.path.exists(_FONT_PATH) else 0)]
try:
    _diag_dir = _resolve_android_app_dir() if _EARLY_IS_ANDROID else os.path.expanduser('~')
    os.makedirs(_diag_dir, exist_ok=True)
    with open(os.path.join(_diag_dir, 'font_debug.log'), 'w', encoding='utf-8') as _f:
        _f.write('\n'.join(_diag_lines))
except Exception:
    pass

import openpyxl
from openpyxl import load_workbook
from PIL import Image as PILImage, ImageDraw, ImageFont

from geocoder import GeoCoder

# === 平台检测 ===
IS_ANDROID = platform == 'android'

if IS_ANDROID:
    from android.permissions import request_permissions, Permission
    try:
        from android import api_version
        ANDROID_API = api_version
    except:
        ANDROID_API = 30

    from jnius import PythonJavaClass, java_method

    class _UiRunnable(PythonJavaClass):
        __javainterfaces__ = ['java/lang/Runnable']
        def __init__(self, func):
            super().__init__()
            self._func = func
        @java_method('()V')
        def run(self):
            try:
                self._func()
            except Exception:
                Logger.error(traceback.format_exc())

    def run_on_ui_thread(func):
        from jnius import autoclass
        activity = autoclass('org.kivy.android.PythonActivity').mActivity
        if activity is not None:
            activity.runOnUiThread(_UiRunnable(func))
        else:
            func()
else:
    ANDROID_API = 0
    def run_on_ui_thread(func):
        func()


def get_status_bar_height_dp():
    """获取Android状态栏高度（dp），非Android返回0"""
    if not IS_ANDROID:
        return 0
    try:
        from jnius import autoclass
        activity = autoclass('org.kivy.android.PythonActivity').mActivity
        resources = activity.getResources()
        resource_id = resources.getIdentifier('status_bar_height', 'dimen', 'android')
        if resource_id > 0:
            px = resources.getDimensionPixelSize(resource_id)
            density = resources.getDisplayMetrics().density
            return int(px / density) + 2
    except:
        pass
    return 28


def setup_android_status_bar():
    """在Android上设置状态栏颜色、禁用edge-to-edge，防止内容被遮挡"""
    if not IS_ANDROID:
        return
    try:
        from jnius import autoclass
        activity = autoclass('org.kivy.android.PythonActivity').mActivity
        window = activity.getWindow()
        if ANDROID_API >= 30:
            window.setDecorFitsSystemWindows(True)
        # 状态栏颜色用深色（与主题bg一致: #1C1C24）
        window.setStatusBarColor(0xFF1C1C24)
        # 状态栏文字浅色（白色图标，适配深色背景）
        from jnius import autoclass as _ac
        View = _ac('android.view.View')
        decor = window.getDecorView()
        flags = decor.getSystemUiVisibility()
        # 清除 LIGHT_STATUS_BAR 标志（使状态栏图标为白色）
        flags = flags & ~View.SYSTEM_UI_FLAG_LIGHT_STATUS_BAR
        decor.setSystemUiVisibility(flags)
    except Exception as e:
        Logger.warning("setup_android_status_bar: %s" % e)


def android_copy_uri_to_app_dir(uri_str, dest_path):
    """通过Android ContentResolver从content URI或文件路径复制文件到app私有目录，
    解决scoped storage下PermissionError问题。
    支持: content:// URI 和 /storage/... 路径
    """
    from jnius import autoclass
    FileInputStream = autoclass('java.io.FileInputStream')
    FileOutputStream = autoclass('java.io.FileOutputStream')
    File = autoclass('java.io.File')
    Uri = autoclass('android.net.Uri')
    activity = autoclass('org.kivy.android.PythonActivity').mActivity
    resolver = activity.getContentResolver()

    if uri_str.startswith('content://'):
        uri = Uri.parse(uri_str)
        in_stream = resolver.openInputStream(uri)
    else:
        # 尝试直接用文件路径
        in_stream = FileInputStream(uri_str)

    out_stream = FileOutputStream(dest_path)
    buffer = bytearray(8192)
    while True:
        read = in_stream.read(buffer)
        if read == -1:
            break
        out_stream.write(buffer, 0, read)
    in_stream.close()
    out_stream.close()
    return dest_path

# === 目录 ===
# 复用 setup_crash_handler() 已解析出的 app 专属外部存储路径：
# /storage/emulated/0/Android/data/<package>/files/loan_photos/
# - 无需任何运行时权限（所有 Android 版本）
# - 不受 Android 10+ scoped storage 限制
# - 用户可通过文件管理器访问
APP_DIR = _resolve_android_app_dir()
try:
    os.makedirs(APP_DIR, exist_ok=True)
except Exception as e:
    # 极端情况下创建失败，回退到 app 私有内部存储（一定可写）
    APP_DIR = os.path.join(os.path.expanduser('~'), 'loan_photos')
    try:
        os.makedirs(APP_DIR, exist_ok=True)
    except Exception:
        pass

PROGRESS_FILE = os.path.join(APP_DIR, 'photo_progress.json')
CONFIG_FILE = os.path.join(APP_DIR, 'app_config.json')
FONT_PATH = _FONT_PATH if _FONT_PATH else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'simhei.ttf')

# === 默认配置 ===
# Excel 格式：A=客户名 B=抵押物地址（概） C=抵押物地址（精确门牌号） D=抵押物性质
DEFAULT_CONFIG = {
    'naming_segments': ['拍摄日期', '客户名', '地址+时间', '空值'],
    'watermark_enabled': True,
    'watermark_segments': ['拍摄时间', '地址名', '经纬度'],
    'watermark_position': 'bottom-right',
    'watermark_font_size': '中',
    'watermark_opacity': 170,
}

# === 命名段选项（用-连接成 X-X-X-X）===
NAMING_SEGMENT_OPTIONS = [
    "拍摄日期",
    "拍摄时间",
    "客户名",
    "抵押物地址（全）",
    "抵押物地址（概）",
    "抵押物地址（精确门牌号）",
    "地址+时间",
    "空值",
]

# === 水印段选项（3段，用-连接成 X-X-X）===
WATERMARK_SEGMENT_OPTIONS = [
    "经纬度",
    "拍摄时间",
    "地址名",
    "空值",
]

# === 水印字号 ===
WATERMARK_FONT_SIZE_OPTIONS = ["大", "中", "小"]
WATERMARK_FONT_SIZE_MAP = {"大": 40, "中": 28, "小": 18}

# === 水印位置 ===
WATERMARK_POSITION_OPTIONS = ['bottom-right', 'bottom-left', 'top-right', 'top-left']
WATERMARK_POSITION_LABELS = {'bottom-right': '右下', 'bottom-left': '左下',
                             'top-right': '右上', 'top-left': '左上'}
WATERMARK_POSITION_LABEL_TO_KEY = {v: k for k, v in WATERMARK_POSITION_LABELS.items()}

# === 作者信息 ===
AUTHOR_NAME = "王硕"
AUTHOR_PHONE = "15940454123（同微信）"
AUTHOR_INFO = f"作者：{AUTHOR_NAME}\n联系方式：{AUTHOR_PHONE}\n有问题请联系作者"

# === 颜色主题 ===
THEME = {
    'bg': (0.11, 0.11, 0.14, 1),
    'card': (0.16, 0.17, 0.21, 1),
    'accent': (0.22, 0.55, 0.85, 1),
    'accent_dark': (0.17, 0.42, 0.72, 1),
    'success': (0.25, 0.72, 0.32, 1),
    'danger': (0.85, 0.25, 0.25, 1),
    'text': (0.92, 0.92, 0.95, 1),
    'text_dim': (0.55, 0.55, 0.65, 1),
    'warning': (0.85, 0.80, 0.30, 1),
}

# === 拍照类型 ===
PHOTO_TYPES = [
    ("远景", "小区/厂区全貌、楼栋外立面"),
    ("近景", "单元门口、楼层门牌、房号牌"),
    ("内部", "室内全景、核心区域现状"),
    ("瑕疵", "破损、漏水、违建、占用特写"),
    ("其他", "其他需要记录的场景"),
]
PHOTO_TYPE_LABELS = ["远景", "近景", "内部", "瑕疵", "其他"]

# ============================================================
# 工具函数
# ============================================================

def get_system_date():
    return datetime.now()

def get_date_str():
    return get_system_date().strftime("%Y%m%d")

def get_time_str():
    return get_system_date().strftime("%H%M")

def get_date_display():
    return get_system_date().strftime("%Y年%m月%d日")

def get_datetime_str():
    return get_system_date().strftime("%Y-%m-%d %H:%M")

def get_full_datetime_str():
    return get_system_date().strftime("%Y-%m-%d %H:%M:%S")

def get_report_date_str():
    return get_system_date().strftime("%Y年%m月%d日")

def clean_filename(s):
    for ch in '/\\:*?"<>|':
        s = s.replace(ch, '_')
    return s.strip()

# ============================================================
# 配置管理器
# ============================================================

class AppConfig:
    """持久化应用配置"""
    def __init__(self):
        self.data = dict(DEFAULT_CONFIG)
        self._load()

    def _load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    self.data.update(saved)
            except Exception:
                pass

    def save(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            Logger.error(f"AppConfig.save: {e}")

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value
        self.save()

# ============================================================
# 进度管理器
# ============================================================

class ProgressManager:
    """拍照进度管理 - 用借款人+地址作为持久化key，跨Excel文件可识别"""
    def __init__(self, filepath=None):
        self.filepath = filepath or PROGRESS_FILE
        self.data = {}
        self.load()

    def _make_key(self, borrower, address=""):
        """基于借款人+地址生成持久化key"""
        raw = (borrower + "|" + address).strip("|")
        return hashlib.md5(raw.encode('utf-8')).hexdigest()[:16]

    def load(self):
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception:
                self.data = {}

    def save(self):
        try:
            with open(self.filepath, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            Logger.error(f"ProgressManager.save: {e}")

    def mark_photo(self, key, photo_path, photo_type=""):
        """key = _make_key(borrower, address)"""
        if key not in self.data:
            self.data[key] = {"photos": [], "types": {}, "timestamp": ""}
        # 验证并保存完整路径
        self.data[key]["photos"].append(os.path.abspath(photo_path))
        if photo_type:
            self.data[key]["types"][photo_type] = True
        self.data[key]["timestamp"] = get_full_datetime_str()
        self.save()

    def delete_photo(self, key, photo_index):
        if key in self.data and photo_index < len(self.data[key]["photos"]):
            self.data[key]["photos"].pop(photo_index)
            if not self.data[key]["photos"]:
                del self.data[key]
            self.save()

    def delete_all_photos(self, key):
        if key in self.data:
            del self.data[key]
            self.save()

    def is_photographed(self, key):
        return key in self.data and len(self.data[key].get("photos", [])) > 0

    def get_photos(self, key):
        """返回存在的照片路径列表（过滤掉已删除的）"""
        if key not in self.data:
            return []
        valid = []
        for p in self.data[key].get("photos", []):
            if os.path.exists(p):
                valid.append(p)
        # 清理失效路径
        if len(valid) != len(self.data[key].get("photos", [])):
            self.data[key]["photos"] = valid
            self.save()
        return valid

    def get_photo_count(self, key):
        return len(self.get_photos(key))

    def get_done_count(self, keys):
        """keys = [(borrower, address), ...]"""
        return sum(1 for b, a in keys if self.is_photographed(self._make_key(b, a)))

    def get_next_photo_index(self, key):
        if key in self.data:
            return len(self.data[key]["photos"])
        return 0

    def get_next_type_index(self, key, photo_type):
        """返回该客户指定类型的下一个照片编号（01开始，跨会话连续）。
        统计已有照片中文件名包含_{type}_NN模式的最大编号+1。
        """
        if key not in self.data:
            return 1
        max_idx = 0
        type_tag = "_%s_" % photo_type
        for p in self.data[key].get("photos", []):
            basename = os.path.basename(p)
            if type_tag in basename:
                try:
                    idx_str = basename.split(type_tag)[-1].split('.')[0].split('_')[0]
                    idx = int(idx_str)
                    if idx > max_idx:
                        max_idx = idx
                except:
                    pass
        return max_idx + 1

    def get_photo_types(self, key):
        return self.data.get(key, {}).get("types", {})

    def get_photo_type_summary(self, key):
        types = self.get_photo_types(key)
        done = sum(1 for t in PHOTO_TYPE_LABELS if types.get(t, False))
        return f"{done}/5"

# ============================================================
# Excel 读取器
# ============================================================

class ExcelReader:
    """读取Excel，A=客户名 B=抵押物地址（概） C=抵押物地址（精确门牌号） D=抵押物性质
    返回 rows 每项: [borrower, address_general, address_precise, property_type]
    """
    def __init__(self, file_path):
        self.file_path = file_path
        self.headers = []
        self.rows = []

    def load(self):
        wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
        ws = wb.active
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            cells = [str(cell).strip() if cell else "" for cell in row[:4]]
            if i == 0:
                self.headers = cells
            else:
                self.rows.append(cells)
        wb.close()
        # 自动判断表头
        if not self.headers:
            self.headers = ["客户名", "抵押物地址（概）", "抵押物地址（精确门牌号）", "抵押物性质"]
        return self.headers, self.rows

# ============================================================
# 报告生成器
# ============================================================

class ReportGenerator:
    def __init__(self, template_path=None):
        self.template_path = template_path

    def generate(self, excel_data, progress_mgr, output_path=None):
        headers, rows = excel_data
        try:
            wb = load_workbook(self.template_path) if self.template_path and os.path.exists(self.template_path) else None
        except:
            wb = None

        if wb is None:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws['A3'] = '序号'
            ws['B3'] = '日期'
            ws['C3'] = '勘查业务贷款人名称'
            ws['D3'] = '抵押物/抵债资产具体情况'
            ws['E3'] = '现状描述'
            ws['F3'] = '备注（是否存在发生风险的可能）'
        else:
            ws = wb['Sheet1']

        for row_idx in range(4, ws.max_row + 1):
            if row_idx < 25:
                for col in range(1, 7):
                    ws.cell(row=row_idx, column=col).value = None

        report_date = get_report_date_str()
        current_row = 4
        seq_num = 1

        for i, row in enumerate(rows):
            customer_name = row[0] if len(row) > 0 else ""
            address_general = row[1] if len(row) > 1 else ""
            address_precise = row[2] if len(row) > 2 else ""
            property_type = row[3] if len(row) > 3 else ""

            if customer_name:
                full_address = (address_general + address_precise).strip()
                ws.cell(row=current_row, column=1, value=seq_num)
                ws.cell(row=current_row, column=2, value=report_date)
                ws.cell(row=current_row, column=3, value=customer_name)
                ws.cell(row=current_row, column=4, value=full_address)
                ws.cell(row=current_row, column=5, value=property_type)
                ws.cell(row=current_row, column=6, value="")
                seq_num += 1
                current_row += 1

        if current_row <= 17:
            current_row = 17
        ws.cell(row=17, column=5, value="当天线路完成进度：100%")

        if output_path is None:
            output_path = os.path.join(APP_DIR, f"现场勘查日报表_{get_date_str()}.xlsx")

        try:
            wb.save(output_path)
            return output_path
        except:
            return None

# ============================================================
# 照片处理器
# ============================================================

class PhotoProcessor:
    """水印、命名、保存"""

    @staticmethod
    def build_watermark_text(segments, **kwargs):
        """根据水印段配置生成水印文本（X-X-X 格式）
        segments: ["经纬度"/"拍摄时间"/"地址名"/"空值", ...]
        kwargs: time_str(拍摄时间), address(地址), lat, lng(经纬度)
        """
        parts = []
        for seg in segments:
            val = ""
            if seg == "经纬度":
                lat = kwargs.get('lat', '')
                lng = kwargs.get('lng', '')
                if lat and lng:
                    val = "%s,%s" % (lng, lat)
                else:
                    val = "定位中"
            elif seg == "拍摄时间":
                val = kwargs.get('time_str', '')
            elif seg == "地址名":
                val = kwargs.get('address', '')
            elif seg == "空值":
                val = ""
            if val:
                parts.append(val)
        return "-".join(parts)

    @staticmethod
    def add_watermark(photo_path, config, **kwargs):
        """根据配置添加水印（段选择模式）"""
        if not config.get('watermark_enabled', True):
            return

        try:
            segments = config.get('watermark_segments', DEFAULT_CONFIG['watermark_segments'])
            text = PhotoProcessor.build_watermark_text(segments, **kwargs)
            if not text:
                return

            img = PILImage.open(photo_path)
            draw = ImageDraw.Draw(img)

            font_size_key = config.get('watermark_font_size', '中')
            font_size = WATERMARK_FONT_SIZE_MAP.get(font_size_key, 28)
            opacity = config.get('watermark_opacity', 170)
            position = config.get('watermark_position', 'bottom-right')

            font = PhotoProcessor._get_font(font_size)

            text_bbox = draw.textbbox((0, 0), text, font=font)
            tw = text_bbox[2] - text_bbox[0]
            th = text_bbox[3] - text_bbox[1] + 10

            padding = 12
            if position == 'bottom-right':
                tx = img.width - tw - padding
                ty = img.height - th - padding
            elif position == 'bottom-left':
                tx = padding
                ty = img.height - th - padding
            elif position == 'top-right':
                tx = img.width - tw - padding
                ty = padding
            else:  # top-left
                tx = padding
                ty = padding

            overlay = PILImage.new('RGBA', img.size, (0, 0, 0, 0))
            od = ImageDraw.Draw(overlay)
            od.rectangle([tx - 4, ty - 2, tx + tw + 4, ty + th + 2],
                         fill=(0, 0, 0, min(opacity, 255)))

            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            img = PILImage.alpha_composite(img, overlay)

            draw = ImageDraw.Draw(img)
            draw.text((tx, ty), text, font=font, fill=(255, 255, 255))

            final = img.convert('RGB') if img.mode == 'RGBA' else img
            final.save(photo_path, 'JPEG', quality=92)
        except Exception as e:
            Logger.error("PhotoProcessor.add_watermark: %s" % e)

    @staticmethod
    def _get_font(size):
        if os.path.exists(FONT_PATH):
            try:
                return ImageFont.truetype(FONT_PATH, size)
            except:
                pass
        font_paths = [
            "C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf",
            "msyh.ttc", "simhei.ttf",
            "/system/fonts/DroidSansFallback.ttf",
            "/system/fonts/NotoSansCJK-Regular.ttc",
        ]
        for fp in font_paths:
            try:
                return ImageFont.truetype(fp, size)
            except:
                continue
        return ImageFont.load_default()

    @staticmethod
    def generate_filename(segments, borrower="", address_general="", address_precise="",
                         property_type="", seq=0, date_str="", photo_type="", time_str=""):
        """根据段选择生成文件名（X-X-X-X 格式，空值段自动跳过）
        segments: ["拍摄日期"/"拍摄时间"/"客户名"/"抵押物地址（全）"/.../"地址+时间"/"空值", ...]
        抵押物地址（全） = 抵押物地址（概） + 抵押物地址（精确门牌号）
        地址+时间 = 抵押物地址（全） + 时间HHMM（无分隔符）
        """
        if not date_str:
            date_str = get_date_str()
        if not time_str:
            time_str = get_time_str()

        parts = []
        for seg in segments:
            val = ""
            if seg == "拍摄日期":
                val = date_str
            elif seg == "拍摄时间":
                val = date_str + time_str
            elif seg == "客户名":
                val = borrower if borrower else "未知"
            elif seg == "抵押物地址（全）":
                val = (address_general + address_precise).strip()
            elif seg == "抵押物地址（概）":
                val = address_general
            elif seg == "抵押物地址（精确门牌号）":
                val = address_precise
            elif seg == "地址+时间":
                full_addr = (address_general + address_precise).strip()
                val = full_addr + time_str if full_addr else time_str
            elif seg == "空值":
                val = ""
            val = clean_filename(val)
            if val:
                parts.append(val)

        filename = "-".join(parts)
        filename = clean_filename(filename)
        if not filename:
            filename = "photo_%s" % date_str
        if photo_type:
            filename = "%s-%s-%02d" % (filename, photo_type, seq)
        if not filename.endswith('.jpg'):
            filename += '.jpg'
        return filename

    @staticmethod
    def save_to_gallery(photo_path):
        if not IS_ANDROID:
            return
        try:
            from android import mActivity
            from android.provider import MediaStore
            from android.content import ContentValues
            from java.io import FileInputStream, FileOutputStream

            values = ContentValues()
            values.put(MediaStore.Images.Media.DATA, photo_path)
            values.put(MediaStore.Images.Media.MIME_TYPE, "image/jpeg")
            values.put(MediaStore.Images.Media.DISPLAY_NAME, os.path.basename(photo_path))

            resolver = mActivity.getContentResolver()
            uri = resolver.insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values)
            if uri:
                output_stream = resolver.openOutputStream(uri)
                input_stream = FileInputStream(photo_path)
                buffer = bytearray(1024)
                while True:
                    read = input_stream.read(buffer)
                    if read == -1:
                        break
                    output_stream.write(buffer, 0, read)
                input_stream.close()
                output_stream.close()
        except Exception as e:
            Logger.error(f"PhotoProcessor.save_to_gallery: {e}")

# ============================================================
# GPS 管理器（非阻塞缓存）
# ============================================================

class GpsManager:
    """异步获取 GPS 坐标并缓存，供水印经纬度段使用。
    不会阻塞主线程：在后台缓存最后一次定位结果。
    """
    def __init__(self):
        self.lat = ""
        self.lng = ""
        self._started = False
        if IS_ANDROID:
            self._start()

    def _start(self):
        try:
            from plyer import gps
            gps.configure(on_location=self._on_location)
            gps.start(min_time=3000, min_distance=0)
            self._started = True
        except Exception:
            self._started = False

    def _on_location(self, **kwargs):
        try:
            lat = kwargs.get('lat')
            lon = kwargs.get('lon')
            if lat is not None and lon is not None:
                self.lat = ("%.6f" % float(lat))
                self.lng = ("%.6f" % float(lon))
        except Exception:
            pass

    def get_coords(self):
        """返回 (lat, lng)，未定位时返回 ('', '')"""
        return self.lat, self.lng


# ============================================================
# 相机管理器
# ============================================================

class CameraManager:
    CAMERA_REQUEST_CODE = 0x123

    def __init__(self):
        self.photo_path = ""
        self.pending_callback = None
        self.status_callback = None
        self.gps = GpsManager()
        self.geocoder = GeoCoder()
        self._camera_launched = False
        self._media_uri = None
        self._launch_attempts = 0
        self._debug_log_path = ""
        self._log_lines = []
        self._max_log_lines = 50

    def _dbg(self, msg, show_toast=False):
        """Write debug message to log file and update UI status.
        show_toast=True to also show an Android Toast (use sparingly to avoid flashing)."""
        ts = get_system_date().strftime('%H:%M:%S')
        line = f"[{ts}] {msg}"
        Logger.info("CAMDBG: %s", msg)
        self._log_lines.append(line)
        if len(self._log_lines) > self._max_log_lines:
            self._log_lines = self._log_lines[-self._max_log_lines:]
        try:
            if not self._debug_log_path:
                self._debug_log_path = os.path.join(APP_DIR, "camera_debug.log")
            os.makedirs(APP_DIR, exist_ok=True)
            with open(self._debug_log_path, 'a', encoding='utf-8') as f:
                f.write(line + "\n")
        except:
            pass
        if self.status_callback:
            try:
                self.status_callback('\n'.join(self._log_lines[-10:]))
            except:
                pass
        if show_toast:
            self._toast(msg)

    def _toast(self, msg):
        """Show a native Android Toast message."""
        if not IS_ANDROID:
            Logger.info("Toast: %s", msg)
            return
        try:
            from jnius import autoclass
            PythonActivity = autoclass('org.kivy.android.PythonActivity')
            Toast = autoclass('android.widget.Toast')
            activity = PythonActivity.mActivity
            if activity:
                toast = Toast.makeText(activity, msg, Toast.LENGTH_LONG)
                toast.show()
        except Exception as e:
            Logger.warning("Toast failed: %s", str(e)[:80])

    def take_photo(self, callback, status_callback=None):
        """启动相机拍照，拍完后callback(photo_path)被调用，失败传None。
        status_callback(msg) 用于实时更新UI状态。"""
        self.pending_callback = callback
        self.status_callback = status_callback
        self._camera_launched = False
        self._media_uri = None
        self.photo_path = ""
        self._dbg("开始拍照流程...")
        if IS_ANDROID:
            self._dbg("正在启动相机...", show_toast=True)
            self._check_and_request_permission()
        else:
            self._dbg("桌面测试模式")
            self._simulate_photo()

    def _check_and_request_permission(self):
        """检查相机权限，有则直接启动相机；无则尝试请求，失败也直接启动（让系统处理）。"""
        self._dbg("检查相机权限...")
        has_perm = False
        try:
            from android.permissions import check_permission
            has_perm = check_permission('android.permission.CAMERA')
            self._dbg(f"权限检查结果: {has_perm}")
        except Exception as e:
            self._dbg(f"check_permission导入失败: {str(e)[:60]}，直接尝试启动")
            has_perm = False

        if has_perm:
            self._dbg("已有相机权限，启动相机")
            self._launch_camera_intent()
            return

        self._dbg("请求相机权限...")
        try:
            request_permissions(
                [Permission.CAMERA],
                self._on_permission_result
            )
            self._dbg("权限请求已发送（等待用户授权弹窗）")
            Clock.schedule_once(self._permission_timeout, 5)
        except Exception as e:
            self._dbg(f"request_permissions失败: {str(e)[:80]}，直接启动")
            self._launch_camera_intent()

    def _permission_timeout(self, dt):
        """5秒后如果还没启动相机，说明权限弹窗可能没出现，直接尝试启动。"""
        if not self._camera_launched and self.pending_callback:
            self._dbg("权限请求超时（5秒未响应），直接尝试启动相机")
            Clock.schedule_once(lambda dt2: self._launch_camera_intent(), 0)

    def _on_permission_result(self, permissions, grants):
        self._dbg(f"权限回调: perms={permissions}, grants={grants}")
        def _handle(dt):
            if any(grants):
                self._dbg("用户已授权，启动相机")
                self._launch_camera_intent()
            else:
                self._dbg("相机权限被拒绝，尝试直接启动")
                self._launch_camera_intent()
        Clock.schedule_once(_handle, 0)

    def _launch_camera_intent(self):
        """通过 Android Intent.ACTION_IMAGE_CAPTURE 调用系统相机。
        v3.9.0 修复：
        - FileProvider优先（Android官方推荐，HyperOS/MIUI兼容最好）
        - 移除FLAG_ACTIVITY_NEW_TASK（与startActivityForResult冲突）
        - 使用run_on_ui_thread确保在主线程执行
        - 保存到getExternalFilesDir/Pictures（与file_paths.xml配置匹配）
        - 详细错误日志持久化在底部面板（不弹Toast一闪而过）
        """
        def _do_launch():
            self._launch_attempts += 1
            PythonActivity = None
            Intent = None
            File = None
            Uri = None
            ActivityNotFoundException = None
            Environment = None
            try:
                from jnius import autoclass, cast
                PythonActivity = autoclass('org.kivy.android.PythonActivity')
                Intent = autoclass('android.content.Intent')
                File = autoclass('java.io.File')
                Uri = autoclass('android.net.Uri')
                Environment = autoclass('android.os.Environment')
                ActivityNotFoundException = autoclass('android.content.ActivityNotFoundException')
            except Exception as e:
                self._dbg(f"jnius类加载失败: {str(e)[:80]}", show_toast=True)
                if self.pending_callback:
                    cb = self.pending_callback
                    self.pending_callback = None
                    Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)
                return

            try:
                raw_activity = PythonActivity.mActivity
                if raw_activity is None:
                    self._dbg("错误：mActivity为None！", show_toast=True)
                    if self.pending_callback:
                        cb = self.pending_callback
                        self.pending_callback = None
                        Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)
                    return
                activity = cast('android.app.Activity', raw_activity)
                package_name = activity.getPackageName()
                self._dbg(f"Activity OK, pkg={package_name}, API={ANDROID_API}")

                self._dbg("正在构建照片文件...")
                photo_fname = f"capture_{get_system_date().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
                pics_dir = activity.getExternalFilesDir(Environment.DIRECTORY_PICTURES)
                if pics_dir is not None:
                    save_dir = str(pics_dir.getAbsolutePath())
                else:
                    cache_dir = activity.getExternalCacheDir()
                    save_dir = str(cache_dir.getAbsolutePath()) if cache_dir else APP_DIR
                os.makedirs(save_dir, exist_ok=True)
                self.photo_path = os.path.join(save_dir, photo_fname)
                self._dbg(f"保存目录: ...{save_dir[-30:]}")

                intent = Intent(Intent.ACTION_IMAGE_CAPTURE)
                intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                intent.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION)

                uri = None
                uri_type = "none"
                self._media_uri = None
                photo_file = None
                MediaStore_Images = None

                try:
                    photo_file = File(self.photo_path)
                    if photo_file.exists():
                        photo_file.delete()
                    photo_file.createNewFile()
                except Exception as e:
                    self._dbg(f"创建照片文件失败: {str(e)[:60]}")

                if photo_file is not None:
                    self._dbg("策略1：FileProvider content:// URI（官方推荐）")
                    for fp_cls_name in ['androidx.core.content.FileProvider', 'android.support.v4.content.FileProvider']:
                        try:
                            FileProvider = autoclass(fp_cls_name)
                            test_uri = FileProvider.getUriForFile(activity, package_name + ".fileprovider", photo_file)
                            if test_uri is not None:
                                uri = test_uri
                                uri_type = f"FileProvider({fp_cls_name.split('.')[-1]})"
                                self._dbg(f"FileProvider URI成功: {uri_type}")
                                break
                        except Exception as e:
                            emsg = str(e)[:80]
                            self._dbg(f"  {fp_cls_name.split('.')[-1]}失败: {emsg}")

                if uri is None and photo_file is not None and ANDROID_API < 26:
                    self._dbg("策略2：file:// URI（低版本兼容）")
                    try:
                        StrictMode = autoclass('android.os.StrictMode')
                        VmPolicyBuilder = autoclass('android.os.StrictMode$VmPolicy$Builder')
                        b = VmPolicyBuilder()
                        b.penaltyLog()
                        policy = b.build()
                        StrictMode.setVmPolicy(policy)
                    except:
                        pass
                    try:
                        uri = Uri.fromFile(photo_file)
                        uri_type = "file://"
                        self._dbg("file:// URI成功")
                    except Exception as e:
                        self._dbg(f"file:// URI失败: {str(e)[:60]}")

                if uri is None and ANDROID_API >= 29:
                    self._dbg("策略3：MediaStore content:// URI（Android 10+）")
                    try:
                        ContentValues = autoclass('android.content.ContentValues')
                        MediaStore_Images = autoclass('android.provider.MediaStore$Images$Media')
                        resolver = activity.getContentResolver()
                        cv = ContentValues()
                        cv.put(MediaStore_Images.DISPLAY_NAME, photo_fname)
                        cv.put(MediaStore_Images.MIME_TYPE, "image/jpeg")
                        media_uri = resolver.insert(MediaStore_Images.EXTERNAL_CONTENT_URI, cv)
                        if media_uri is not None:
                            uri = media_uri
                            self.photo_path = None
                            self._media_uri = uri
                            uri_type = "MediaStore"
                            self._dbg("MediaStore URI成功")
                        else:
                            self._dbg("MediaStore返回null")
                    except Exception as e:
                        self._dbg(f"MediaStore失败: {str(e)[:60]}")

                if uri is None:
                    self._dbg("错误：所有URI策略均失败！", show_toast=True)
                    self._dbg(f"photo_path={self.photo_path}", show_toast=True)
                    if self.pending_callback:
                        cb = self.pending_callback
                        self.pending_callback = None
                        Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)
                    return

                if uri_type == "MediaStore" and MediaStore_Images is not None:
                    intent.putExtra(MediaStore_Images.EXTRA_OUTPUT, uri)
                else:
                    intent.putExtra("output", uri)
                self._dbg(f"URI已设置(type={uri_type})")

                try:
                    pm = activity.getPackageManager()
                    resolves = pm.queryIntentActivities(intent, 0)
                    if resolves is not None and resolves.size() > 0:
                        self._dbg(f"找到{resolves.size()}个相机应用")
                        for i in range(resolves.size()):
                            try:
                                ri = resolves.get(i)
                                pkg = ri.activityInfo.packageName
                                activity.grantUriPermission(pkg, uri,
                                    Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_READ_URI_PERMISSION)
                            except:
                                pass
                    else:
                        self._dbg("未检测到相机应用列表（Android11+可见性限制），继续启动...")
                except Exception as e:
                    self._dbg(f"应用列表检查异常: {str(e)[:40]}")

                self._camera_launched = False
                launched = False
                launch_error = ""

                try:
                    self._dbg("正在启动相机（请稍候）...")
                    activity.startActivityForResult(intent, self.CAMERA_REQUEST_CODE)
                    launched = True
                    self._camera_launched = True
                    self._dbg("相机启动指令已发送")
                except ActivityNotFoundException as e:
                    launch_error = f"系统未找到相机应用(ActivityNotFound)"
                    self._dbg(launch_error)
                except Exception as e:
                    launch_error = f"启动异常: {type(e).__name__}: {str(e)[:100]}"
                    self._dbg(launch_error)

                if not launched:
                    self._dbg("直接启动失败，尝试通过选择器启动...")
                    try:
                        chooser_intent = Intent.createChooser(intent, "选择相机应用")
                        chooser_intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                        chooser_intent.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
                        activity.startActivityForResult(chooser_intent, self.CAMERA_REQUEST_CODE)
                        launched = True
                        self._camera_launched = True
                        self._dbg("相机选择器已弹出")
                    except ActivityNotFoundException as e:
                        launch_error = "系统未安装相机应用"
                        self._dbg(launch_error, show_toast=True)
                    except Exception as e:
                        launch_error = f"选择器也失败: {type(e).__name__}: {str(e)[:100]}"
                        self._dbg(launch_error, show_toast=True)

                if not launched:
                    self._dbg(f"相机启动失败！请截图此界面联系作者", show_toast=True)
                    self._dbg(f"错误详情: {launch_error}")
                    if self.pending_callback:
                        cb = self.pending_callback
                        self.pending_callback = None
                        Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)
                    return

            except Exception as e:
                err_msg = f"相机启动异常: {type(e).__name__}: {str(e)[:100]}"
                self._dbg(err_msg, show_toast=True)
                Logger.error(traceback.format_exc())
                self._camera_launched = False
                if self.pending_callback:
                    cb = self.pending_callback
                    self.pending_callback = None
                    Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)

        try:
            from android.runnable import run_on_ui_thread
            _do_launch_ui = run_on_ui_thread(_do_launch)
            _do_launch_ui()
        except ImportError:
            Clock.schedule_once(lambda dt: _do_launch(), 0)
        except Exception as e:
            self._dbg(f"UI线程调度失败: {str(e)[:60]}，直接执行")
            _do_launch()

    def on_camera_result(self, result_code):
        """由 MainScreen.on_activity_result 在收到 CAMERA_REQUEST_CODE 结果时调用。"""
        self._dbg(f"收到相机结果: result_code={result_code}, launched={self._camera_launched}")
        if not self._camera_launched:
            return
        self._camera_launched = False
        if result_code == -1:  # RESULT_OK
            if self._media_uri is not None:
                self._dbg("处理MediaStore返回的照片...")
                try:
                    from jnius import autoclass
                    PythonActivity = autoclass('org.kivy.android.PythonActivity')
                    activity = PythonActivity.mActivity
                    resolver = activity.getContentResolver()
                    istream = resolver.openInputStream(self._media_uri)
                    os.makedirs(APP_DIR, exist_ok=True)
                    dest_path = os.path.join(
                        APP_DIR, f"capture_{get_system_date().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
                    )
                    ByteArrayOutputStream = autoclass('java.io.ByteArrayOutputStream')
                    baos = ByteArrayOutputStream()
                    buf = bytearray(8192)
                    while True:
                        n = istream.read(buf)
                        if n <= 0:
                            break
                        baos.write(buf, 0, n)
                    istream.close()
                    with open(dest_path, 'wb') as f:
                        f.write(bytes(baos.toByteArray()))
                    baos.close()
                    self.photo_path = dest_path
                    self._dbg(f"MediaStore照片已保存: {os.path.getsize(dest_path)} bytes")
                except Exception as e:
                    self._dbg(f"MediaStore复制失败: {str(e)[:100]}")
                    self._media_uri = None
                    if self.pending_callback:
                        cb = self.pending_callback
                        self.pending_callback = None
                        Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)
                    return
                self._media_uri = None

            if self.photo_path and os.path.exists(self.photo_path) and os.path.getsize(self.photo_path) > 0:
                self._dbg(f"拍照成功！", show_toast=True)
                if self.pending_callback:
                    cb = self.pending_callback
                    self.pending_callback = None
                    Clock.schedule_once(lambda dt, cb=cb: cb(self.photo_path), 0)
                return
            else:
                self._dbg(f"照片文件不存在或为空", show_toast=True)
        else:
            self._dbg("用户取消了拍照")
        if self.pending_callback:
            cb = self.pending_callback
            self.pending_callback = None
            Clock.schedule_once(lambda dt, cb=cb: cb(None), 0)

    def _simulate_photo(self):
        """桌面端测试：生成一张模拟照片"""
        self.photo_path = os.path.join(APP_DIR, f"test_{get_system_date().strftime('%Y%m%d_%H%M%S')}.jpg")
        img = PILImage.new('RGB', (640, 480), (180, 180, 180))
        draw = ImageDraw.Draw(img)
        font = PhotoProcessor._get_font(24)
        now_str = get_full_datetime_str()
        draw.text((150, 200), "Test Photo", fill=(0, 0, 0), font=font)
        draw.text((150, 250), now_str, fill=(0, 0, 0), font=font)
        img.save(self.photo_path)
        if self.pending_callback:
            Clock.schedule_once(lambda dt: self.pending_callback(self.photo_path), 0.3)

    def get_location_name(self, lat, lng):
        """根据 GPS 坐标逆地理编码获取地名，供水印「地址名」段使用。
        在后台线程调用（HTTP/计算不会阻塞 UI）。
        无 GPS 时返回 'GPS未开启或无权限'。"""
        if lat and lng:
            try:
                name = self.geocoder.reverse_geocode(float(lng), float(lat))
                if name:
                    return name
                return "GPS定位解析失败"
            except Exception as e:
                Logger.error("CameraManager.get_location_name: %s" % e)
                return "GPS定位解析失败"
        return "GPS未开启或无权限"

# ============================================================
# 自定义小组件
# ============================================================

class CardWidget(BoxLayout):
    """卡片式容器"""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = [12, 10, 12, 10]
        self.spacing = 6
        with self.canvas.before:
            Color(*THEME['card'])
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[6])
        self.bind(pos=self._update_rect, size=self._update_rect)

    def _update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size


class SectionLabel(Label):
    """分区标题"""
    def __init__(self, text="", **kwargs):
        super().__init__(**kwargs)
        self.text = text
        self.font_size = '17sp'
        self.bold = True
        self.color = THEME['accent']
        self.size_hint_y = None
        self.height = dp(40)
        self.halign = 'left'
        self.valign = 'middle'


# ============================================================
# 欢迎页面
# ============================================================

class WelcomeScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = 'welcome'

        root = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(10))

        # 顶部状态栏留白（Android上为状态栏高度，桌面端为小间距）
        self._status_spacer = Label(size_hint_y=None, height=dp(get_status_bar_height_dp()))
        root.add_widget(self._status_spacer)

        # Logo 图标
        root.add_widget(Label(
            text="📷", font_size='72sp',
            size_hint_y=None, height=dp(90), color=THEME['accent'],
        ))

        # 标题
        root.add_widget(Label(
            text="资产盘点专项拍照工具", font_size='26sp',
            bold=True, color=THEME['text'],
            size_hint_y=None, height=dp(48),
        ))

        # 副标题
        root.add_widget(Label(
            text="银行抵押物现场勘查工具", font_size='15sp',
            color=THEME['text_dim'],
            size_hint_y=None, height=dp(32),
        ))

        # 版本
        root.add_widget(Label(
            text="v3.5.0", font_size='12sp',
            color=THEME['text_dim'],
            size_hint_y=None, height=dp(24),
        ))

        # 间距
        root.add_widget(Label(size_hint_y=None, height=dp(12)))

        # 功能简介卡片
        feat_card = CardWidget(size_hint_y=None)
        feat_card.bind(minimum_height=feat_card.setter('height'))
        features = [
            "• 四类拍照引导（远景/近景/内部/瑕疵）",
            "• 水印自选模式（段+位置+字号）",
            "• 文件命名自选模式（4段下拉）",
            "• 一键生成勘查日报表",
        ]
        for feat in features:
            feat_card.add_widget(Label(
                text=feat, font_size='15sp',
                color=THEME['text_dim'],
                size_hint_y=None, height=dp(32),
                halign='left', valign='middle',
                text_size=(None, dp(32)),
            ))
        root.add_widget(feat_card)

        # 弹性空间
        root.add_widget(Label(size_hint_y=1))

        # 作者信息
        author_card = CardWidget(size_hint_y=None, height=dp(80))
        author_card.add_widget(Label(
            text=AUTHOR_INFO, font_size='13sp',
            color=THEME['text_dim'], halign='center', valign='middle',
        ))
        root.add_widget(author_card)

        # 间距
        root.add_widget(Label(size_hint_y=None, height=dp(10)))

        # 进入按钮 - 大尺寸适合手指点击
        start_btn = Button(
            text="开始使用", font_size='22sp',
            size_hint_y=None, height=dp(60),
            background_color=THEME['accent'],
            background_normal='',
            color=(1, 1, 1, 1),
            bold=True,
        )
        start_btn.bind(on_release=self._go_main)
        root.add_widget(start_btn)

        self.add_widget(root)

    def _go_main(self, instance):
        self.manager.current = 'main'


# ============================================================
# 拍照类型选择弹窗
# ============================================================

class PhotoTypePopup(Popup):
    def __init__(self, on_select, **kwargs):
        super().__init__(**kwargs)
        self.title = "选择拍照类型"
        self.size_hint = (0.85, 0.55)
        self.auto_dismiss = True
        self.on_select = on_select

        layout = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(14))
        layout.add_widget(Label(text="请选择本次拍摄的照片类型：", font_size='16sp', size_hint_y=None, height=dp(40)))

        for type_name, type_desc in PHOTO_TYPES:
            btn = Button(
                text=f"{type_name}\n{type_desc}", font_size='16sp',
                size_hint_y=None, height=dp(72),
                background_color=THEME['accent'], background_normal='',
                halign='center', color=(1,1,1,1), bold=True,
            )
            btn.bind(on_release=lambda x, t=type_name: self._select(t))
            layout.add_widget(btn)

        cancel_btn = Button(text="取消", font_size='16sp', size_hint_y=None, height=dp(52),
                           background_color=(0.5, 0.5, 0.5, 1), background_normal='',
                           color=(1,1,1,1))
        cancel_btn.bind(on_release=self.dismiss)
        layout.add_widget(cancel_btn)
        self.content = layout

    def _select(self, photo_type):
        self.dismiss()
        if self.on_select:
            self.on_select(photo_type)


# ============================================================
# 照片查看弹窗
# ============================================================

class PhotoViewerPopup(Popup):
    def __init__(self, row_index, photos, delete_callback, **kwargs):
        super().__init__(**kwargs)
        self.title = f"已拍照片 ({len(photos)}张)"
        self.size_hint = (0.92, 0.8)
        self.row_index = row_index
        self.photos = photos
        self.delete_callback = delete_callback

        main_layout = BoxLayout(orientation='vertical', spacing=dp(8), padding=dp(10))
        scroll = ScrollView()
        list_layout = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        list_layout.bind(minimum_height=list_layout.setter('height'))

        for i, photo_path in enumerate(photos):
            item = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(130), spacing=dp(8))
            if os.path.exists(photo_path):
                item.add_widget(KivyImage(source=photo_path, size_hint_x=0.38, allow_stretch=True, keep_ratio=True))

            info_box = BoxLayout(orientation='vertical', spacing=dp(6), size_hint_x=0.62)
            info_box.add_widget(Label(text=os.path.basename(photo_path), font_size='12sp',
                                      halign='left', valign='top', size_hint_y=0.5,
                                      text_size=(None, None)))
            del_btn = Button(text="删除此照片", font_size='14sp', size_hint_y=0.5, height=dp(48),
                            background_color=THEME['danger'], background_normal='',
                            color=(1,1,1,1), bold=True)
            del_btn.bind(on_release=lambda x, idx=i: self._confirm_delete(idx))
            info_box.add_widget(del_btn)
            item.add_widget(info_box)
            list_layout.add_widget(item)

        scroll.add_widget(list_layout)
        main_layout.add_widget(scroll)

        del_all_btn = Button(text="删除全部照片（重拍）", font_size='15sp', size_hint_y=None, height=dp(52),
                            background_color=THEME['danger'], background_normal='',
                            color=(1,1,1,1), bold=True)
        del_all_btn.bind(on_release=self._confirm_delete_all)
        main_layout.add_widget(del_all_btn)

        close_btn = Button(text="关闭", font_size='16sp', size_hint_y=None, height=dp(52),
                          background_color=THEME['accent'], background_normal='',
                          color=(1,1,1,1), bold=True)
        close_btn.bind(on_release=self.dismiss)
        main_layout.add_widget(close_btn)
        self.content = main_layout

    def _confirm_delete(self, index):
        """删除单张照片前弹出二次确认"""
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(text=f"确定要删除这张照片吗？\n此操作不可撤销。",
                                 font_size='16sp', color=THEME['text'], halign='center'))
        btn_row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(10))
        popup = Popup(title="确认删除", size_hint=(0.8, 0.35), auto_dismiss=True)
        yes_btn = Button(text="确认删除", font_size='16sp', background_color=THEME['danger'],
                         background_normal='', color=(1,1,1,1), bold=True)
        no_btn = Button(text="取消", font_size='16sp', background_color=THEME['accent'],
                        background_normal='', color=(1,1,1,1), bold=True)
        def _do_delete(instance):
            popup.dismiss()
            self.delete_callback(self.row_index, index)
            self.dismiss()
        yes_btn.bind(on_release=_do_delete)
        no_btn.bind(on_release=popup.dismiss)
        btn_row.add_widget(yes_btn)
        btn_row.add_widget(no_btn)
        content.add_widget(btn_row)
        popup.content = content
        popup.open()

    def _confirm_delete_all(self, instance):
        """删除全部照片前弹出二次确认"""
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(text=f"确定要删除该客户的全部照片吗？\n此操作不可撤销！",
                                 font_size='16sp', color=THEME['danger'], halign='center'))
        btn_row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(10))
        popup = Popup(title="⚠ 危险操作确认", size_hint=(0.85, 0.35), auto_dismiss=True)
        yes_btn = Button(text="全部删除", font_size='16sp', background_color=THEME['danger'],
                         background_normal='', color=(1,1,1,1), bold=True)
        no_btn = Button(text="取消", font_size='16sp', background_color=THEME['accent'],
                        background_normal='', color=(1,1,1,1), bold=True)
        def _do_delete_all(instance):
            popup.dismiss()
            self.delete_callback(self.row_index, -1)
            self.dismiss()
        yes_btn.bind(on_release=_do_delete_all)
        no_btn.bind(on_release=popup.dismiss)
        btn_row.add_widget(yes_btn)
        btn_row.add_widget(no_btn)
        content.add_widget(btn_row)
        popup.content = content
        popup.open()


# ============================================================
# 设置页面
# ============================================================

class SettingsScreen(Screen):
    def __init__(self, app_config, **kwargs):
        super().__init__(**kwargs)
        self.name = 'settings'
        self.config = app_config

        main = BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(8))

        # 顶部状态栏留白
        main.add_widget(Label(size_hint_y=None, height=dp(get_status_bar_height_dp())))

        # 标题栏 - 增大高度和按钮
        title_bar = BoxLayout(size_hint_y=None, height=dp(56), spacing=dp(8))
        back_btn = Button(text="← 返回", font_size='18sp', size_hint_x=0.28,
                         background_color=THEME['accent'], background_normal='',
                         color=(1,1,1,1), bold=True,
                         size_hint_y=None, height=dp(52))
        back_btn.bind(on_release=self._go_back)
        title_bar.add_widget(back_btn)
        title_bar.add_widget(Label(text="设置", font_size='22sp', bold=True, color=THEME['text']))
        main.add_widget(title_bar)

        scroll = ScrollView()
        content = BoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None, padding=[dp(4), dp(4), dp(4), dp(30)])
        content.bind(minimum_height=content.setter('height'))

        # === 命名规则 ===
        naming_card = CardWidget(size_hint_y=None)
        naming_card.bind(minimum_height=naming_card.setter('height'))

        naming_card.add_widget(SectionLabel(text="📋 照片命名规则"))

        naming_card.add_widget(Label(text="格式：X-X-X-X（每段自选，空值段自动省略）",
                                     font_size='13sp', color=THEME['text_dim'],
                                     size_hint_y=None, height=dp(24)))

        cur_segments = self.config.get('naming_segments', DEFAULT_CONFIG['naming_segments'])
        self.naming_spinners = []
        for idx in range(4):
            row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
            row.add_widget(Label(text="第%d段" % (idx + 1), font_size='15sp',
                                 color=THEME['text_dim'], size_hint_x=0.2))
            cur_val = cur_segments[idx] if idx < len(cur_segments) else "空值"
            sp = Spinner(
                text=cur_val,
                values=NAMING_SEGMENT_OPTIONS,
                size_hint_x=0.8, font_size='15sp',
            )
            sp.bind(text=lambda inst, val, i=idx: self._on_naming_segment_change(i, val))
            self.naming_spinners.append(sp)
            row.add_widget(sp)
            naming_card.add_widget(row)

        preview = self._get_naming_preview()
        self.naming_preview_label = Label(text="预览：%s" % preview, font_size='13sp',
                                          color=THEME['text_dim'], size_hint_y=None, height=dp(28))
        naming_card.add_widget(self.naming_preview_label)

        save_naming_btn = Button(text="保存命名规则", font_size='16sp', size_hint_y=None, height=dp(52),
                                background_color=THEME['accent'], background_normal='',
                                color=(1,1,1,1), bold=True)
        save_naming_btn.bind(on_release=self._save_naming)
        naming_card.add_widget(save_naming_btn)

        content.add_widget(naming_card)

        # === 水印设置 ===
        watermark_card = CardWidget(size_hint_y=None)
        watermark_card.bind(minimum_height=watermark_card.setter('height'))

        watermark_card.add_widget(SectionLabel(text="💧 水印设置"))

        # 水印开关
        toggle_box = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
        toggle_box.add_widget(Label(text="启用水印", font_size='16sp', color=THEME['text'], size_hint_x=0.6))
        self.wm_check = CheckBox(active=self.config.get('watermark_enabled', True), size_hint_x=0.4)
        toggle_box.add_widget(self.wm_check)
        watermark_card.add_widget(toggle_box)

        watermark_card.add_widget(Label(text="水印格式：X-X-X（每段自选）",
                                        font_size='13sp', color=THEME['text_dim'],
                                        size_hint_y=None, height=dp(24)))

        # 水印段选择（3段）
        cur_wm_segments = self.config.get('watermark_segments', DEFAULT_CONFIG['watermark_segments'])
        self.wm_spinners = []
        for idx in range(3):
            row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
            row.add_widget(Label(text="第%d段" % (idx + 1), font_size='15sp',
                                 color=THEME['text_dim'], size_hint_x=0.2))
            cur_val = cur_wm_segments[idx] if idx < len(cur_wm_segments) else "空值"
            sp = Spinner(
                text=cur_val,
                values=WATERMARK_SEGMENT_OPTIONS,
                size_hint_x=0.8, font_size='15sp',
            )
            sp.bind(text=lambda inst, val, i=idx: self._on_wm_segment_change(i, val))
            self.wm_spinners.append(sp)
            row.add_widget(sp)
            watermark_card.add_widget(row)

        # 水印位置
        pos_box = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        pos_box.add_widget(Label(text="位置：", font_size='15sp', color=THEME['text'], size_hint_x=0.25))
        cur_pos = self.config.get('watermark_position', DEFAULT_CONFIG['watermark_position'])
        cur_pos_label = WATERMARK_POSITION_LABELS.get(cur_pos, '右下')
        self.pos_spinner = Spinner(
            text=cur_pos_label,
            values=list(WATERMARK_POSITION_LABELS.values()),
            size_hint_x=0.75, font_size='15sp',
        )
        pos_box.add_widget(self.pos_spinner)
        watermark_card.add_widget(pos_box)

        # 字号（大/中/小）
        font_size_box = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        font_size_box.add_widget(Label(text="字号：", font_size='15sp', color=THEME['text'], size_hint_x=0.25))
        cur_fs = self.config.get('watermark_font_size', DEFAULT_CONFIG['watermark_font_size'])
        if cur_fs not in WATERMARK_FONT_SIZE_OPTIONS:
            cur_fs = '中'
        self.font_spinner = Spinner(
            text=cur_fs,
            values=WATERMARK_FONT_SIZE_OPTIONS,
            size_hint_x=0.75, font_size='15sp',
        )
        font_size_box.add_widget(self.font_spinner)
        watermark_card.add_widget(font_size_box)

        save_wm_btn = Button(text="保存水印设置", font_size='16sp', size_hint_y=None, height=dp(52),
                            background_color=THEME['accent'], background_normal='',
                            color=(1,1,1,1), bold=True)
        save_wm_btn.bind(on_release=self._save_watermark)
        watermark_card.add_widget(save_wm_btn)

        content.add_widget(watermark_card)

        # === 关于 ===
        about_card = CardWidget(size_hint_y=None, height=dp(140))
        about_card.add_widget(SectionLabel(text="ℹ️ 关于"))
        about_card.add_widget(Label(text=AUTHOR_INFO, font_size='14sp',
                                    color=THEME['text_dim'], halign='center'))
        content.add_widget(about_card)

        content.add_widget(Label(size_hint_y=None, height=dp(20)))

        scroll.add_widget(content)
        main.add_widget(scroll)
        self.add_widget(main)

    def _on_naming_segment_change(self, idx, val):
        """命名段下拉变化时刷新预览"""
        self.naming_preview_label.text = "预览：%s" % self._get_naming_preview()

    def _get_naming_preview(self):
        """根据当前 4 个命名段下拉生成预览文件名"""
        segments = [sp.text for sp in self.naming_spinners]
        return PhotoProcessor.generate_filename(
            segments, borrower="张三",
            address_general="沈阳市和平区XX街123号", address_precise="1-23-4",
            property_type="住宅", seq=1, date_str="20260628", photo_type="远景",
            time_str="1430",
        )

    def _on_wm_segment_change(self, idx, val):
        """水印段下拉变化时刷新预览"""
        # 水印预览暂不单独显示，保存时生效
        pass

    def _save_naming(self, instance):
        segments = [sp.text for sp in self.naming_spinners]
        self.config.set('naming_segments', segments)
        self.naming_preview_label.text = "预览：%s" % self._get_naming_preview()
        self._show_toast("命名规则已保存")

    def _save_watermark(self, instance):
        segments = [sp.text for sp in self.wm_spinners]
        self.config.set('watermark_enabled', self.wm_check.active)
        self.config.set('watermark_segments', segments)
        pos_label = self.pos_spinner.text
        self.config.set('watermark_position',
                        WATERMARK_POSITION_LABEL_TO_KEY.get(pos_label, 'bottom-right'))
        self.config.set('watermark_font_size', self.font_spinner.text)
        self._show_toast("水印设置已保存")

    def _show_toast(self, msg):
        popup = Popup(title='', content=Label(text=msg, font_size='14sp'),
                     size_hint=(0.6, 0.2), auto_dismiss=True)
        popup.open()
        Clock.schedule_once(lambda dt: popup.dismiss(), 1.5)

    def _go_back(self, instance):
        self.manager.current = 'main'


# ============================================================
# 单行组件
# ============================================================

class RowWidget(BoxLayout):
    def __init__(self, row_index, borrower, address_general, address_precise, property_type,
                 progress_key, progress_mgr, photo_callback, view_photos_callback, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.size_hint_y = None
        self._base_height = dp(120)
        self.height = self._base_height
        self.padding = [dp(10), dp(8), dp(10), dp(8)]
        self.spacing = dp(4)

        self.row_index = row_index
        self.borrower = borrower
        self.address_general = address_general
        self.address_precise = address_precise
        self.property_type = property_type
        self.progress_key = progress_key
        self.photo_callback = photo_callback
        self.view_photos_callback = view_photos_callback
        self.progress_mgr = progress_mgr
        self.done = progress_mgr.is_photographed(progress_key)
        self.photo_count = progress_mgr.get_photo_count(progress_key)

        with self.canvas.before:
            Color(*THEME['card'])
            self.bg_rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[4])
        self.bind(pos=self._update_bg, size=self._update_bg)

        full_address = (address_general + address_precise).strip()
        addr_display = full_address if full_address else "（无地址）"
        if property_type:
            addr_display += "  [%s]" % property_type

        name_text = borrower if borrower else "（无客户名）"

        # 客户名（加粗，自动换行）
        self.name_label = Label(
            text=name_text,
            font_size='16sp',
            bold=True,
            color=THEME['success'] if self.done else THEME['text'],
            size_hint_y=None,
            halign='left', valign='middle',
        )
        self.name_label.bind(width=self._update_name_text_size)
        self.name_label.text_size = (None, None)
        self.name_label.bind(texture_size=self._update_heights)
        self.add_widget(self.name_label)

        # 地址+性质（自动换行，灰色小字）
        self.addr_label = Label(
            text=addr_display,
            font_size='13sp',
            color=THEME['text_dim'],
            size_hint_y=None,
            halign='left', valign='top',
        )
        self.addr_label.bind(width=self._update_addr_text_size)
        self.addr_label.text_size = (None, None)
        self.addr_label.bind(texture_size=self._update_heights)
        self.add_widget(self.addr_label)

        # 间隔
        self.add_widget(Label(size_hint_y=None, height=dp(4)))

        # 按钮行（拍照 + 查看已拍 + 类型计数）
        btn_row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))

        self.photo_btn = Button(
            text="拍照", font_size='17sp',
            size_hint_x=0.32,
            background_color=THEME['success'] if self.done else THEME['accent'],
            background_normal='',
            color=(1, 1, 1, 1), bold=True,
        )
        self.photo_btn.bind(on_release=self._on_photo)
        btn_row.add_widget(self.photo_btn)

        self.view_btn = Button(
            text="查看已拍(%d)" % self.photo_count if self.photo_count > 0 else "查看已拍",
            font_size='16sp', size_hint_x=0.50,
            background_color=THEME['accent_dark'] if self.photo_count > 0 else (0.3, 0.3, 0.4, 1),
            background_normal='',
            color=(1, 1, 1, 1), bold=True,
        )
        self.view_btn.bind(on_release=self._on_view_photos)
        btn_row.add_widget(self.view_btn)

        self.type_status = Label(text="", font_size='14sp',
                                color=THEME['success'] if self.done else THEME['text_dim'],
                                size_hint_x=0.18, bold=True)
        btn_row.add_widget(self.type_status)

        self.add_widget(btn_row)

        self._update_type_status()
        Clock.schedule_once(lambda dt: self._update_heights(), 0)

    def _update_name_text_size(self, instance, value):
        instance.text_size = (value, None)

    def _update_addr_text_size(self, instance, value):
        instance.text_size = (value, None)

    def _update_heights(self, *args):
        name_h = self.name_label.texture_size[1] if self.name_label.texture_size else dp(24)
        addr_h = self.addr_label.texture_size[1] if self.addr_label.texture_size else dp(20)
        self.name_label.height = max(dp(24), name_h + dp(4))
        self.addr_label.height = max(dp(20), addr_h + dp(4))
        self.height = self.name_label.height + self.addr_label.height + dp(52) + dp(16) + dp(4)

    def _update_bg(self, *args):
        self.bg_rect.pos = self.pos
        self.bg_rect.size = self.size

    def _update_type_status(self):
        types = self.progress_mgr.get_photo_types(self.progress_key)
        done_types = sum(1 for t in PHOTO_TYPE_LABELS if types.get(t, False))
        self.type_status.text = "%d/5" % done_types
        self.type_status.color = THEME['success'] if done_types >= 5 else (THEME['warning'] if done_types > 0 else THEME['text_dim'])

    def _on_photo(self, instance):
        self.photo_callback(self.row_index, self.borrower, self.address_general,
                           self.address_precise, self.property_type)

    def _on_view_photos(self, instance):
        self.view_photos_callback(self.row_index)

    def mark_done(self):
        self.done = True
        self.photo_count = self.progress_mgr.get_photo_count(self.progress_key)
        self.photo_btn.background_color = THEME['success']
        self.name_label.color = THEME['success']
        self.view_btn.text = "查看已拍(%d)" % self.photo_count
        self.view_btn.background_color = THEME['accent_dark'] if self.photo_count > 0 else (0.3, 0.3, 0.4, 1)
        self._update_type_status()
        Clock.schedule_once(lambda dt: self._update_heights(), 0)


# ============================================================
# 主界面
# ============================================================

class MainScreen(Screen):
    def __init__(self, app_config, **kwargs):
        super().__init__(**kwargs)
        self.name = 'main'
        self.config = app_config
        self.excel_path = ""
        self.headers = []
        self.rows = []  # [borrower, address_general, address_precise, property_type]
        self.progress_mgr = ProgressManager()
        self.camera_mgr = CameraManager()
        self.report_generator = ReportGenerator()
        self.row_widgets = []
        self._current_row = 0
        self._current_borrower = ""
        self._current_addr_general = ""
        self._current_addr_precise = ""
        self._current_property_type = ""
        self._current_photo_type = ""
        self._current_key = ""
        self._continuous_shooting = False
        self._photos_in_session = 0

        main = BoxLayout(orientation='vertical')
        self._build_ui(main)
        self.add_widget(main)

    def _build_ui(self, parent):
        parent.add_widget(Label(size_hint_y=None, height=dp(get_status_bar_height_dp())))

        title_bar = BoxLayout(size_hint_y=None, height=dp(56), spacing=dp(6), padding=[dp(10), dp(4), dp(10), dp(4)])
        title_bar.add_widget(Label(text="资产盘点专项拍照工具", font_size='18sp', bold=True, color=THEME['text'],
                                   size_hint_x=0.55, halign='left', valign='middle'))

        settings_btn = Button(text="⚙ 设置", font_size='16sp', size_hint_x=0.22,
                             background_color=THEME['accent'], background_normal='',
                             color=(1,1,1,1), bold=True)
        settings_btn.bind(on_release=self._go_settings)
        title_bar.add_widget(settings_btn)

        self.progress_label = Label(text="0/0", font_size='15sp', color=THEME['text_dim'],
                                    size_hint_x=0.23, halign='right', valign='middle')
        title_bar.add_widget(self.progress_label)
        parent.add_widget(title_bar)

        toolbar = BoxLayout(size_hint_y=None, height=dp(56), spacing=dp(6), padding=[dp(10), dp(4), dp(10), dp(4)])
        open_btn = Button(text="📂 打开Excel", font_size='15sp', size_hint_x=0.34,
                         background_color=THEME['accent'], background_normal='',
                         color=(1,1,1,1), bold=True)
        open_btn.bind(on_release=self._show_file_dialog)
        toolbar.add_widget(open_btn)

        self.search_input = TextInput(hint_text="搜索客户名…", multiline=False, font_size='15sp', size_hint_x=0.46)
        self.search_input.bind(text=self._on_search)
        toolbar.add_widget(self.search_input)

        search_btn = Button(text="搜索", font_size='14sp', size_hint_x=0.20,
                           background_color=THEME['accent'], background_normal='',
                           color=(1,1,1,1), bold=True)
        search_btn.bind(on_release=self._do_search)
        toolbar.add_widget(search_btn)
        parent.add_widget(toolbar)

        self.scroll_view = ScrollView(do_scroll_x=False, do_scroll_y=True)
        self.list_layout = GridLayout(cols=1, spacing=dp(6), size_hint_y=None, padding=[dp(8), dp(6), dp(8), dp(6)])
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        self.scroll_view.add_widget(self.list_layout)
        parent.add_widget(self.scroll_view)

        footer = BoxLayout(orientation='vertical', size_hint_y=None, height=dp(180), spacing=dp(4), padding=[dp(10), dp(6), dp(10), dp(6)])

        btn_row = BoxLayout(size_hint_y=None, height=dp(38), spacing=dp(6))
        report_btn = Button(text="（已屏蔽）生成报告", font_size='13sp',
                           background_color=(0.35, 0.35, 0.35, 1), background_normal='',
                           size_hint_x=0.55, color=(0.7, 0.7, 0.7, 1))
        report_btn.disabled = True
        btn_row.add_widget(report_btn)

        clear_log_btn = Button(text="清空日志", font_size='12sp',
                              background_color=(0.4, 0.4, 0.45, 1), background_normal='',
                              size_hint_x=0.45, color=(1,1,1,1))
        clear_log_btn.bind(on_release=self._clear_debug_log)
        btn_row.add_widget(clear_log_btn)
        footer.add_widget(btn_row)

        log_scroll = ScrollView(do_scroll_x=False, do_scroll_y=True, size_hint_y=1)
        self.status_label = Label(text="请点击「打开Excel」选择文件", font_size='12sp',
                                  color=THEME['warning'],
                                  halign='left', valign='top',
                                  size_hint_y=None, markup=False)
        self.status_label.bind(texture_size=self._update_log_label_size, width=lambda i, v: setattr(i, 'text_size', (v, None)))
        log_scroll.add_widget(self.status_label)
        footer.add_widget(log_scroll)
        parent.add_widget(footer)

        self._show_empty_state()

    def _show_empty_state(self):
        self.list_layout.clear_widgets()
        if not self.rows:
            msg = Label(
                text="暂无数据\n\n点击「打开Excel」加载客户清单",
                font_size='15sp', color=THEME['text_dim'],
                size_hint_y=None, height=200,
            )
            self.list_layout.add_widget(msg)

    def _build_header_row(self):
        header = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(36),
                           padding=[dp(10), dp(4), dp(10), dp(4)], spacing=dp(6))
        with header.canvas.before:
            Color(0.20, 0.22, 0.28, 1)
            self._header_rect = RoundedRectangle(pos=header.pos, size=header.size, radius=[4])
        header.bind(pos=self._update_header_rect, size=self._update_header_rect)
        header.add_widget(Label(text="客户名 / 抵押物地址", font_size='13sp', bold=True,
                                color=THEME['accent'], size_hint_x=0.6, halign='left'))
        header.add_widget(Label(text="操作", font_size='13sp', bold=True,
                                color=THEME['accent'], size_hint_x=0.4, halign='center'))
        return header

    def _update_header_rect(self, instance, *args):
        self._header_rect.pos = instance.pos
        self._header_rect.size = instance.size

    def _show_file_dialog(self, instance):
        """打开系统文件选择器。Android使用SAF(Intent.ACTION_OPEN_DOCUMENT)获取可访问URI，
        桌面端使用plyer.filechooser，最终fallback手动输入。"""
        if IS_ANDROID:
            self._android_open_file_picker()
        else:
            try:
                from plyer import filechooser
                filechooser.open_file(
                    filters=[("Excel 文件", "*.xlsx", "*.xls")],
                    on_selection=self._on_file_selected,
                )
            except Exception as e:
                Logger.warning("MainScreen.filechooser: %s, fallback to path input" % e)
                self._show_path_input_dialog()

    def _android_open_file_picker(self):
        """Android：使用Intent.ACTION_OPEN_DOCUMENT选择Excel文件，
        通过ContentResolver获取持久读取权限，然后复制到app私有目录。"""
        try:
            from jnius import autoclass
            PythonActivity = autoclass('org.kivy.android.PythonActivity')
            Intent = autoclass('android.content.Intent')
            activity = PythonActivity.mActivity

            intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
            intent.addCategory(Intent.CATEGORY_OPENABLE)
            intent.setType("*/*")
            intent.putExtra(Intent.EXTRA_MIME_TYPES, [
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
            ])
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            intent.addFlags(Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION)

            self._android_file_picker_code = 0x1001
            activity.startActivityForResult(intent, self._android_file_picker_code)
        except Exception as e:
            Logger.error("Android file picker error: %s" % e)
            self.status_label.text = "无法打开文件选择器"
            self.status_label.color = THEME['danger']

    def on_activity_result(self, request_code, result_code, intent):
        """处理Android Activity结果回调（文件选择器+相机）。"""
        if request_code == self.camera_mgr.CAMERA_REQUEST_CODE:
            self.camera_mgr.on_camera_result(result_code)
            return

        if request_code == getattr(self, '_android_file_picker_code', -1):
            if result_code != -1 or intent is None:
                return
            try:
                from jnius import autoclass
                uri = intent.getData()
                if uri is None:
                    return
                uri_str = str(uri.toString())

                PythonActivity = autoclass('org.kivy.android.PythonActivity')
                activity = PythonActivity.mActivity
                Intent = autoclass('android.content.Intent')
                resolver = activity.getContentResolver()
                try:
                    take_flags = Intent.FLAG_GRANT_READ_URI_PERMISSION
                    resolver.takePersistableUriPermission(uri, take_flags)
                except:
                    pass

                dest = os.path.join(APP_DIR, "_imported_excel.xlsx")
                android_copy_uri_to_app_dir(uri_str, dest)
                Clock.schedule_once(lambda dt: self._load_excel_path(dest), 0)
            except Exception as e:
                Logger.error("on_activity_result (file): %s" % e)
                self.status_label.text = "文件选择失败: %s" % str(e)[:40]
                self.status_label.color = THEME['danger']

    def _on_file_selected(self, selection):
        if not selection:
            return
        path = selection[0] if isinstance(selection, list) else str(selection)
        if path and os.path.exists(path):
            self._load_excel_path(path)
        elif path:
            self.status_label.text = "文件不存在：%s" % path[:40]
            self.status_label.color = THEME['danger']

    def _show_path_input_dialog(self):
        content = BoxLayout(orientation='vertical', spacing=8, padding=8)
        content.add_widget(Label(text="请输入 Excel 文件路径：", size_hint_y=None, height=dp(30),
                                 font_size='14sp'))
        path_input = TextInput(text=self.excel_path, multiline=False, font_size='13sp',
                               size_hint_y=None, height=dp(40))
        content.add_widget(path_input)
        load_btn = Button(text="加载", size_hint_y=None, height=dp(44),
                         background_color=THEME['accent'], background_normal='',
                         font_size='16sp', color=(1,1,1,1))
        popup = Popup(title='选择 Excel 文件', content=content, size_hint=(0.9, 0.35),
                      title_size='16sp')
        load_btn.bind(on_release=lambda x: (popup.dismiss(), self._load_excel_path(path_input.text)))
        content.add_widget(load_btn)
        popup.open()

    def _load_excel_path(self, path):
        if not path or not os.path.exists(path):
            self.status_label.text = "文件不存在！"
            self.status_label.color = THEME['danger']
            return
        self.excel_path = path
        try:
            reader = ExcelReader(path)
            self.headers, self.rows = reader.load()
            self.status_label.text = "✓ 已加载 %d 条记录" % len(self.rows)
            self.status_label.color = THEME['success']
            self._refresh_list()
        except Exception as e:
            err_msg = str(e)
            self.status_label.text = "加载失败: %s" % err_msg[:80]
            self.status_label.color = THEME['danger']
            Logger.error("Excel load failed: %s" % traceback.format_exc())

    def _refresh_list(self):
        self.list_layout.clear_widgets()
        self.row_widgets = []

        if self.rows:
            self.list_layout.add_widget(self._build_header_row())

        for i, row in enumerate(self.rows):
            borrower = row[0] if len(row) > 0 else ""
            address_general = row[1] if len(row) > 1 else ""
            address_precise = row[2] if len(row) > 2 else ""
            property_type = row[3] if len(row) > 3 else ""
            full_addr = (address_general + address_precise).strip()
            pk = self.progress_mgr._make_key(borrower, full_addr)

            rw = RowWidget(
                row_index=i, borrower=borrower,
                address_general=address_general, address_precise=address_precise,
                property_type=property_type,
                progress_key=pk, progress_mgr=self.progress_mgr,
                photo_callback=self._on_photo_request,
                view_photos_callback=self._on_view_photos,
            )
            self.list_layout.add_widget(rw)
            self.row_widgets.append(rw)

        self._update_progress()

    def _update_progress(self):
        total = len(self.rows)
        keys = []
        for r in self.rows:
            b = r[0] if len(r) > 0 else ""
            ag = r[1] if len(r) > 1 else ""
            ap = r[2] if len(r) > 2 else ""
            keys.append((b, (ag + ap).strip()))
        done = self.progress_mgr.get_done_count(keys)
        self.progress_label.text = "%d/%d" % (done, total)
        if total > 0:
            self.status_label.text = "进度 %d/%d" % (done, total)
            self.status_label.color = THEME['success'] if done == total else THEME['warning']

    def _on_search(self, instance, text=None):
        query = self.search_input.text.lower().strip()
        for rw in self.row_widgets:
            if not query or query in rw.borrower.lower():
                rw.opacity = 1
                rw.size_hint_y = None
                Clock.schedule_once(lambda dt, w=rw: w._update_heights(), 0)
            else:
                rw.opacity = 0
                rw.height = 0

    def _do_search(self, instance):
        self.search_input.focus = False
        self._on_search(None)

    def _clear_search(self, instance):
        self.search_input.text = ""
        self._on_search(None, "")

    def _clear_debug_log(self, instance):
        self.status_label.text = "日志已清空"
        self.status_label.color = THEME['text_dim']
        self.camera_mgr._log_lines = []

    def _update_log_label_size(self, instance, value):
        instance.height = max(dp(80), value[1] + dp(8))

    def _go_settings(self, instance):
        self._continuous_shooting = False
        self.manager.current = 'settings'

    def _on_photo_request(self, row_index, borrower, address_general, address_precise, property_type):
        self._current_row = row_index
        self._current_borrower = borrower
        self._current_addr_general = address_general
        self._current_addr_precise = address_precise
        self._current_property_type = property_type
        self._current_key = self.progress_mgr._make_key(borrower, (address_general + address_precise).strip())
        self._photos_in_session = 0
        popup = PhotoTypePopup(on_select=self._on_photo_type_selected)
        popup.open()

    def _on_photo_type_selected(self, photo_type):
        self._current_photo_type = photo_type
        self._continuous_shooting = True
        self._photos_in_session = 0
        self.status_label.text = "正在启动相机（%s），按返回键可结束拍摄" % photo_type
        self.status_label.color = THEME['warning']
        Clock.schedule_once(lambda dt: self.camera_mgr.take_photo(self._on_photo_done, self._camera_status_update), 0.3)

    def _camera_status_update(self, msg):
        """CameraManager回调：更新UI状态显示调试信息（多行日志）"""
        def _update(dt):
            self.status_label.text = msg
            self.status_label.color = THEME['warning']
            Clock.schedule_once(lambda dt2: self._scroll_log_to_bottom(), 0.05)
        Clock.schedule_once(_update, 0)

    def _scroll_log_to_bottom(self):
        """自动滚动日志到底部显示最新消息"""
        try:
            sv = self.status_label.parent
            if isinstance(sv, ScrollView):
                sv.scroll_y = 0
        except:
            pass

    def _launch_next_photo(self):
        """拍完一张后立即重新调起相机，继续拍摄同一类型，用户按快门拍下一张。"""
        if self._continuous_shooting:
            self.status_label.text = "准备下一张（%s）…按返回键结束" % self._current_photo_type
            self.status_label.color = THEME['warning']
            Clock.schedule_once(lambda dt: self.camera_mgr.take_photo(self._on_photo_done, self._camera_status_update), 0.3)

    def _on_photo_done(self, photo_path):
        if photo_path is None:
            self._continuous_shooting = False
            self.status_label.text = "拍照已取消"
            self.status_label.color = THEME['text_dim']
            self._refresh_row_done(self._current_row)
            return

        row_index = self._current_row
        borrower = self._current_borrower
        addr_general = self._current_addr_general
        addr_precise = self._current_addr_precise
        property_type = self._current_property_type
        photo_type = self._current_photo_type
        key = self._current_key

        self._photos_in_session += 1
        seq = self.progress_mgr.get_next_photo_index(key) + 1

        date_str = get_date_str()
        time_str = get_time_str()
        datetime_str = get_datetime_str()

        self.status_label.text = "处理中 (%s-%02d)…" % (photo_type, self._photos_in_session)
        self.status_label.color = THEME['warning']

        config_data = self.config.data
        naming_segments = self.config.get('naming_segments', DEFAULT_CONFIG['naming_segments'])
        lat, lng = self.camera_mgr.gps.get_coords()

        def _process():
            try:
                place_name = self.camera_mgr.get_location_name(lat, lng)

                PhotoProcessor.add_watermark(
                    photo_path, config_data,
                    time_str=datetime_str, address=place_name,
                    lat=lat, lng=lng,
                )

                filename = PhotoProcessor.generate_filename(
                    naming_segments, borrower, addr_general, addr_precise,
                    property_type, seq, date_str, photo_type, time_str,
                )
                new_path = os.path.join(APP_DIR, filename)
                if photo_path != new_path:
                    if os.path.exists(new_path):
                        name_base, ext = os.path.splitext(filename)
                        suffix = 2
                        while os.path.exists(new_path):
                            new_path = os.path.join(APP_DIR, "%s-%d%s" % (name_base, suffix, ext if ext else '.jpg'))
                            suffix += 1
                    os.rename(photo_path, new_path)

                PhotoProcessor.save_to_gallery(new_path)
                self.progress_mgr.mark_photo(key, new_path, photo_type)
                Clock.schedule_once(lambda dt: self._on_photo_saved(row_index, filename), 0)
            except Exception as e:
                Logger.error("MainScreen._on_photo_done: %s" % e)
                Logger.error(traceback.format_exc())
                err_msg = str(e)
                Clock.schedule_once(lambda dt: self._on_photo_failed(err_msg), 0)

        threading.Thread(target=_process, daemon=True).start()

    @mainthread
    def _on_photo_saved(self, row_index, filename):
        self._refresh_row_done(row_index)
        self.status_label.text = "✓ %s 已保存" % filename[:28]
        self.status_label.color = THEME['success']
        self._launch_next_photo()

    @mainthread
    def _on_photo_failed(self, err_msg):
        self._continuous_shooting = False
        self.status_label.text = "保存失败: %s" % err_msg[:40]
        self.status_label.color = THEME['danger']

    def _refresh_row_done(self, row_index):
        if 0 <= row_index < len(self.row_widgets):
            self.row_widgets[row_index].mark_done()
        self._update_progress()

    def _on_view_photos(self, row_index):
        if row_index >= len(self.rows):
            return
        row = self.rows[row_index]
        borrower = row[0] if len(row) > 0 else ""
        address_general = row[1] if len(row) > 1 else ""
        address_precise = row[2] if len(row) > 2 else ""
        full_addr = (address_general + address_precise).strip()
        key = self.progress_mgr._make_key(borrower, full_addr)
        photos = self.progress_mgr.get_photos(key)
        if not photos:
            self.status_label.text = "暂无照片"
            self.status_label.color = THEME['warning']
            return
        popup = PhotoViewerPopup(row_index=row_index, photos=photos,
                                 delete_callback=self._on_delete_photo)
        popup.open()

    def _on_delete_photo(self, row_index, photo_index):
        if row_index >= len(self.rows):
            return
        row = self.rows[row_index]
        borrower = row[0] if len(row) > 0 else ""
        address_general = row[1] if len(row) > 1 else ""
        address_precise = row[2] if len(row) > 2 else ""
        full_addr = (address_general + address_precise).strip()
        key = self.progress_mgr._make_key(borrower, full_addr)
        if photo_index == -1:
            # 删除全部：先删文件，再清进度
            photos = self.progress_mgr.get_photos(key)
            for p in photos:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except:
                    pass
            self.progress_mgr.delete_all_photos(key)
            self.status_label.text = "已删除该客户全部照片"
            self.status_label.color = THEME['warning']
        else:
            # 删除单张
            photos = self.progress_mgr.get_photos(key)
            if photo_index < len(photos):
                p = photos[photo_index]
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except:
                    pass
            self.progress_mgr.delete_photo(key, photo_index)
        Clock.schedule_once(lambda dt: self._refresh_row_done(row_index), 0)

    def _generate_report(self, instance):
        self.status_label.text = "报告功能暂未开放"
        self.status_label.color = THEME['text_dim']


# ============================================================
# App 入口
# ============================================================

class LoanPhotoApp(App):
    def build(self):
        self.title = "资产盘点专项拍照工具"
        Window.clearcolor = THEME['bg']
        Window.softinput_mode = 'pan'

        # 拦截Android返回键：边缘滑动不直接退出App，而是返回上一页面
        Window.bind(on_keyboard=self._on_keyboard)

        self.config = AppConfig()

        sm = ScreenManager(transition=SlideTransition(duration=0.25))
        self.sm = sm
        sm.add_widget(WelcomeScreen(name='welcome'))
        sm.add_widget(MainScreen(app_config=self.config, name='main'))
        sm.add_widget(SettingsScreen(app_config=self.config, name='settings'))
        sm.current = 'welcome'
        return sm

    def _on_keyboard(self, window, key, scancode, codepoint, modifier):
        if key == 27:
            current = self.sm.current
            if current == 'settings':
                self.sm.current = 'main'
                return True
            elif current == 'main':
                main_screen = self.sm.get_screen('main')
                if getattr(main_screen, '_continuous_shooting', False):
                    main_screen._continuous_shooting = False
                    main_screen.status_label.text = "已结束连续拍照"
                    main_screen.status_label.color = THEME['text_dim']
                    return True
                self.sm.current = 'welcome'
                return True
            elif current == 'welcome':
                if hasattr(self, '_back_pressed') and (datetime.now() - self._back_pressed).total_seconds() < 2:
                    self.stop()
                else:
                    self._back_pressed = datetime.now()
                    try:
                        from kivy.uix.popup import Popup
                        popup = Popup(title='',
                                     content=Label(text="再按一次退出", font_size='16sp'),
                                     size_hint=(0.5, 0.15), auto_dismiss=True)
                        popup.open()
                        Clock.schedule_once(lambda dt: popup.dismiss(), 1.5)
                    except:
                        pass
                return True
        return False

    def on_start(self):
        if IS_ANDROID:
            setup_android_status_bar()
            self._register_android_activity_result()
            self._request_permissions()

    def _register_android_activity_result(self):
        """注册Android activity结果回调，把文件选择器结果转发给MainScreen"""
        try:
            from android.activity import bind as activity_bind
            activity_bind(on_activity_result=self._on_android_activity_result)
        except Exception as e:
            Logger.warning("activity_bind failed: %s" % e)

    def _on_android_activity_result(self, request_code, result_code, intent):
        """Android activity结果统一入口，转发给当前screen"""
        if self.sm:
            current_screen = self.sm.current_screen
            if current_screen and hasattr(current_screen, 'on_activity_result'):
                try:
                    current_screen.on_activity_result(request_code, result_code, intent)
                except Exception as e:
                    Logger.error("on_activity_result error: %s", e)
            elif self.sm.current == 'main':
                main_screen = self.sm.get_screen('main')
                if hasattr(main_screen, 'on_activity_result'):
                    main_screen.on_activity_result(request_code, result_code, intent)

    def _request_permissions(self):
        try:
            if ANDROID_API >= 33:
                perms = [Permission.CAMERA, Permission.ACCESS_FINE_LOCATION,
                         Permission.ACCESS_COARSE_LOCATION]
            elif ANDROID_API >= 30:
                perms = [Permission.CAMERA, Permission.ACCESS_FINE_LOCATION,
                         Permission.ACCESS_COARSE_LOCATION]
            else:
                perms = [Permission.CAMERA, Permission.ACCESS_FINE_LOCATION,
                         Permission.ACCESS_COARSE_LOCATION,
                         Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE]
            request_permissions(perms)
        except:
            pass

    def on_pause(self):
        return True

    def on_resume(self):
        pass


if __name__ == '__main__':
    LoanPhotoApp().run()

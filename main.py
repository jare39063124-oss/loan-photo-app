"""
资产盘点专项拍照工具 App - v3.3
功能：
- 欢迎页 + 设置页
- 文件命名自选模式（4段下拉 X-X-X-X）
- 水印自选模式（3段下拉 X-X-X + 大中小字号 + 位置）
- 四类拍照引导（远景/近景/内部/瑕疵）
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
else:
    ANDROID_API = 0

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
FONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'NotoSansSC.ttf')

# === 默认配置 ===
# Excel 格式：A=客户名 B=抵押物地址（概） C=抵押物地址（精确门牌号） D=抵押物性质
DEFAULT_CONFIG = {
    'naming_segments': ['拍摄时间', '客户名', '抵押物地址（全）', '空值'],
    'watermark_enabled': True,
    'watermark_segments': ['拍摄时间', '地址名', '经纬度'],
    'watermark_position': 'bottom-right',
    'watermark_font_size': '中',
    'watermark_opacity': 170,
}

# === 命名段选项（4段，用-连接成 X-X-X-X）===
NAMING_SEGMENT_OPTIONS = [
    "拍摄时间",
    "客户名",
    "抵押物地址（全）",
    "抵押物地址（概）",
    "抵押物地址（精确门牌号）",
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
AUTHOR_PHONE = "15940454123"
AUTHOR_WECHAT = "15940454123（同微信）"
AUTHOR_INFO = f"作者：{AUTHOR_NAME}\n联系方式：{AUTHOR_PHONE}\n{AUTHOR_WECHAT}\n有问题请联系作者"

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
]
PHOTO_TYPE_LABELS = ["远景", "近景", "内部", "瑕疵"]

# ============================================================
# 工具函数
# ============================================================

def get_system_date():
    return datetime.now()

def get_date_str():
    return get_system_date().strftime("%Y%m%d")

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

    def get_photo_types(self, key):
        return self.data.get(key, {}).get("types", {})

    def get_photo_type_summary(self, key):
        types = self.get_photo_types(key)
        done = sum(1 for t in PHOTO_TYPE_LABELS if types.get(t, False))
        return f"{done}/4"

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
                         property_type="", seq=0, date_str="", photo_type=""):
        """根据段选择生成文件名（X-X-X-X 格式，空值段自动跳过）
        segments: ["拍摄时间"/"客户名"/"抵押物地址（全）"/.../"空值", ...]
        抵押物地址（全） = 抵押物地址（概） + 抵押物地址（精确门牌号）
        """
        if not date_str:
            date_str = get_date_str()

        parts = []
        for seg in segments:
            val = ""
            if seg == "拍摄时间":
                val = date_str
            elif seg == "客户名":
                val = borrower if borrower else "未知"
            elif seg == "抵押物地址（全）":
                val = (address_general + address_precise).strip()
            elif seg == "抵押物地址（概）":
                val = address_general
            elif seg == "抵押物地址（精确门牌号）":
                val = address_precise
            elif seg == "空值":
                val = ""
            val = clean_filename(val)
            if val:
                parts.append(val)

        filename = "-".join(parts)
        filename = clean_filename(filename)
        if not filename:
            filename = "photo_%s" % date_str
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
    def __init__(self):
        self.photo_path = ""
        self.pending_callback = None
        self.gps = GpsManager()
        self.geocoder = GeoCoder()
        self._plyer_failed = False

    def take_photo(self, callback):
        self.pending_callback = callback
        if IS_ANDROID:
            self._request_camera_permission()
        else:
            self._simulate_photo()

    def _request_camera_permission(self):
        try:
            request_permissions(
                [Permission.CAMERA, Permission.ACCESS_FINE_LOCATION],
                self._on_permission_result
            )
        except:
            self._simulate_photo()

    def _on_permission_result(self, permissions, grants):
        if all(grants):
            self._launch_camera()
        else:
            self._simulate_photo()

    def _launch_camera(self):
        if self._plyer_failed:
            self._launch_camera_intent()
            return
        try:
            from plyer import camera
            self.photo_path = os.path.join(
                APP_DIR, f"capture_{get_system_date().strftime('%Y%m%d_%H%M%S')}.jpg"
            )
            camera.take_picture(filename=self.photo_path, on_complete=self._on_photo_complete)
        except:
            self._plyer_failed = True
            self._launch_camera_intent()

    def _launch_camera_intent(self):
        try:
            from android import mActivity
            from android.content import Intent, FileProvider
            from android.net import Uri
            from java.io import File

            self.photo_path = os.path.join(
                APP_DIR, f"capture_{get_system_date().strftime('%Y%m%d_%H%M%S')}.jpg"
            )
            intent = Intent("android.media.action.IMAGE_CAPTURE")
            photo_file = File(self.photo_path)
            uri = FileProvider.getUriForFile(mActivity, mActivity.getPackageName() + ".fileprovider", photo_file)
            intent.putExtra("android.provider.extra_OUTPUT", uri)
            intent.addFlags(1)
            mActivity.startActivityForResult(intent, 0x123)
            Clock.schedule_once(self._check_intent_result, 3.0)
        except:
            self._simulate_photo()

    def _check_intent_result(self, dt):
        if os.path.exists(self.photo_path) and os.path.getsize(self.photo_path) > 0:
            if self.pending_callback:
                self.pending_callback(self.photo_path)
        else:
            Clock.schedule_once(self._check_intent_result, 2.0)

    def _on_photo_complete(self, path):
        if self.pending_callback and path and os.path.exists(path):
            self.pending_callback(path)
        else:
            self.pending_callback(None)

    def _simulate_photo(self):
        self.photo_path = os.path.join(APP_DIR, f"test_{get_system_date().strftime('%Y%m%d_%H%M%S')}.jpg")
        img = PILImage.new('RGB', (640, 480), (180, 180, 180))
        draw = ImageDraw.Draw(img)
        font = PhotoProcessor._get_font(24)
        now_str = get_full_datetime_str()
        draw.text((150, 200), "Test Photo", fill=(0, 0, 0), font=font)
        draw.text((150, 250), now_str, fill=(0, 0, 0), font=font)
        img.save(self.photo_path)
        if self.pending_callback:
            Clock.schedule_once(lambda dt: self.pending_callback(self.photo_path), 0.5)

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
        self.font_size = '14sp'
        self.bold = True
        self.color = THEME['accent']
        self.size_hint_y = None
        self.height = 30
        self.halign = 'left'
        self.text_size = (Window.width * 0.85, None)


# ============================================================
# 欢迎页面
# ============================================================

class WelcomeScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = 'welcome'
        layout = BoxLayout(orientation='vertical', padding=30, spacing=10)
        layout.pos_hint = {'center_x': 0.5, 'center_y': 0.5}

        # Logo 区域
        logo_area = BoxLayout(orientation='vertical', size_hint_y=0.4)
        logo_area.add_widget(Label(
            text="📷", font_size='80sp',
            size_hint_y=0.5, color=THEME['accent'],
        ))
        logo_area.add_widget(Label(
            text="欢迎使用", font_size='20sp',
            color=THEME['text_dim'], size_hint_y=0.15,
        ))
        logo_area.add_widget(Label(
            text="资产盘点专项拍照工具", font_size='26sp',
            bold=True, color=THEME['text'],
            size_hint_y=0.25,
        ))
        logo_area.add_widget(Label(
            text="银行抵押物现场勘查工具", font_size='14sp',
            color=THEME['text_dim'], size_hint_y=0.2,
        ))
        layout.add_widget(logo_area)

        # 版本信息
        layout.add_widget(Label(
            text="v3.3", font_size='12sp',
            color=THEME['text_dim'], size_hint_y=None, height=20,
        ))

        # 功能简介
        features = [
            "• 四类拍照引导（远景/近景/内部/瑕疵）",
            "• 水印自选模式（段+位置+字号）",
            "• 文件命名自选模式（4段下拉）",
            "• 一键生成勘查日报表",
        ]
        for feat in features:
            layout.add_widget(Label(
                text=feat, font_size='13sp',
                color=THEME['text_dim'], size_hint_y=None, height=22,
                halign='center',
            ))

        layout.add_widget(Label(size_hint_y=0.1))

        # 作者信息
        author_card = CardWidget(size_hint_y=None, height=80)
        author_card.add_widget(Label(
            text=AUTHOR_INFO, font_size='12sp',
            color=THEME['text_dim'], halign='center',
        ))
        layout.add_widget(author_card)

        layout.add_widget(Label(size_hint_y=0.05))

        # 进入按钮
        start_btn = Button(
            text="开始使用", font_size='18sp',
            size_hint_y=None, height=52,
            background_color=THEME['accent'],
            background_normal='',
        )
        start_btn.bind(on_release=self._go_main)
        layout.add_widget(start_btn)

        main_layout = FloatLayout()
        main_layout.add_widget(layout)
        self.add_widget(main_layout)

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

        layout = BoxLayout(orientation='vertical', spacing=8, padding=12)
        layout.add_widget(Label(text="请选择本次拍摄的照片类型：", font_size='14sp', size_hint_y=None, height=30))

        for type_name, type_desc in PHOTO_TYPES:
            btn = Button(
                text=f"{type_name}\n{type_desc}", font_size='13sp',
                size_hint_y=None, height=60,
                background_color=THEME['accent'], background_normal='',
                halign='center',
            )
            btn.bind(on_release=lambda x, t=type_name: self._select(t))
            layout.add_widget(btn)

        cancel_btn = Button(text="取消", font_size='14sp', size_hint_y=None, height=40,
                           background_color=(0.5, 0.5, 0.5, 1), background_normal='')
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
        self.size_hint = (0.9, 0.8)
        self.row_index = row_index
        self.photos = photos
        self.delete_callback = delete_callback

        main_layout = BoxLayout(orientation='vertical', spacing=8, padding=8)
        scroll = ScrollView()
        list_layout = GridLayout(cols=1, spacing=8, size_hint_y=None)
        list_layout.bind(minimum_height=list_layout.setter('height'))

        for i, photo_path in enumerate(photos):
            item = BoxLayout(orientation='horizontal', size_hint_y=None, height=110, spacing=8)
            if os.path.exists(photo_path):
                item.add_widget(KivyImage(source=photo_path, size_hint_x=0.35, allow_stretch=True, keep_ratio=True))

            info_box = BoxLayout(orientation='vertical', spacing=4, size_hint_x=0.65)
            info_box.add_widget(Label(text=os.path.basename(photo_path), font_size='11sp',
                                      halign='left', size_hint_y=0.4, text_size=(None, None)))
            del_btn = Button(text="删除此照片", font_size='12sp', size_hint_y=0.4,
                            background_color=THEME['danger'], background_normal='')
            del_btn.bind(on_release=lambda x, idx=i: self._delete_photo(idx))
            info_box.add_widget(del_btn)
            item.add_widget(info_box)
            list_layout.add_widget(item)

        scroll.add_widget(list_layout)
        main_layout.add_widget(scroll)

        del_all_btn = Button(text="删除全部照片（重拍）", font_size='14sp', size_hint_y=None, height=44,
                            background_color=THEME['danger'], background_normal='')
        del_all_btn.bind(on_release=self._delete_all)
        main_layout.add_widget(del_all_btn)

        close_btn = Button(text="关闭", font_size='14sp', size_hint_y=None, height=44)
        close_btn.bind(on_release=self.dismiss)
        main_layout.add_widget(close_btn)
        self.content = main_layout

    def _delete_photo(self, index):
        self.delete_callback(self.row_index, index)
        self.dismiss()

    def _delete_all(self, instance):
        self.delete_callback(self.row_index, -1)
        self.dismiss()


# ============================================================
# 设置页面
# ============================================================

class SettingsScreen(Screen):
    def __init__(self, app_config, **kwargs):
        super().__init__(**kwargs)
        self.name = 'settings'
        self.config = app_config

        main = BoxLayout(orientation='vertical', padding=12, spacing=8)

        # 标题栏
        title_bar = BoxLayout(size_hint_y=None, height=48, spacing=8)
        back_btn = Button(text="← 返回", font_size='14sp', size_hint_x=0.25,
                         background_color=THEME['accent'], background_normal='')
        back_btn.bind(on_release=self._go_back)
        title_bar.add_widget(back_btn)
        title_bar.add_widget(Label(text="设置", font_size='18sp', bold=True, color=THEME['text']))
        main.add_widget(title_bar)

        scroll = ScrollView()
        content = BoxLayout(orientation='vertical', spacing=10, size_hint_y=None, padding=[0, 0, 0, 20])
        content.bind(minimum_height=content.setter('height'))

        # === 命名规则（4段下拉自选模式 X-X-X-X）===
        naming_card = CardWidget(size_hint_y=None)
        naming_card.bind(minimum_height=naming_card.setter('height'))

        naming_card.add_widget(SectionLabel(text="📋 照片命名规则"))

        naming_card.add_widget(Label(text="格式：X-X-X-X（每段自选，空值段自动省略）",
                                     font_size='10sp', color=THEME['text_dim'],
                                     size_hint_y=None, height=18))

        cur_segments = self.config.get('naming_segments', DEFAULT_CONFIG['naming_segments'])
        self.naming_spinners = []
        for idx in range(4):
            row = BoxLayout(size_hint_y=None, height=40, spacing=6)
            row.add_widget(Label(text="第%d段" % (idx + 1), font_size='11sp',
                                 color=THEME['text_dim'], size_hint_x=0.18))
            cur_val = cur_segments[idx] if idx < len(cur_segments) else "空值"
            sp = Spinner(
                text=cur_val,
                values=NAMING_SEGMENT_OPTIONS,
                size_hint_x=0.82, font_size='11sp',
            )
            sp.bind(text=lambda inst, val, i=idx: self._on_naming_segment_change(i, val))
            self.naming_spinners.append(sp)
            row.add_widget(sp)
            naming_card.add_widget(row)

        preview = self._get_naming_preview()
        self.naming_preview_label = Label(text="预览：%s" % preview, font_size='11sp',
                                          color=THEME['text_dim'], size_hint_y=None, height=22)
        naming_card.add_widget(self.naming_preview_label)

        save_naming_btn = Button(text="保存命名规则", font_size='13sp', size_hint_y=None, height=40,
                                background_color=THEME['accent'], background_normal='')
        save_naming_btn.bind(on_release=self._save_naming)
        naming_card.add_widget(save_naming_btn)

        content.add_widget(naming_card)

        # === 水印设置（3段下拉自选模式 X-X-X）===
        watermark_card = CardWidget(size_hint_y=None)
        watermark_card.bind(minimum_height=watermark_card.setter('height'))

        watermark_card.add_widget(SectionLabel(text="💧 水印设置"))

        # 水印开关
        toggle_box = BoxLayout(size_hint_y=None, height=36, spacing=8)
        toggle_box.add_widget(Label(text="启用水印", font_size='13sp', color=THEME['text'], size_hint_x=0.6))
        self.wm_check = CheckBox(active=self.config.get('watermark_enabled', True), size_hint_x=0.4)
        toggle_box.add_widget(self.wm_check)
        watermark_card.add_widget(toggle_box)

        watermark_card.add_widget(Label(text="水印格式：X-X-X（每段自选）",
                                        font_size='10sp', color=THEME['text_dim'],
                                        size_hint_y=None, height=18))

        # 水印段选择（3段）
        cur_wm_segments = self.config.get('watermark_segments', DEFAULT_CONFIG['watermark_segments'])
        self.wm_spinners = []
        for idx in range(3):
            row = BoxLayout(size_hint_y=None, height=40, spacing=6)
            row.add_widget(Label(text="第%d段" % (idx + 1), font_size='11sp',
                                 color=THEME['text_dim'], size_hint_x=0.18))
            cur_val = cur_wm_segments[idx] if idx < len(cur_wm_segments) else "空值"
            sp = Spinner(
                text=cur_val,
                values=WATERMARK_SEGMENT_OPTIONS,
                size_hint_x=0.82, font_size='11sp',
            )
            sp.bind(text=lambda inst, val, i=idx: self._on_wm_segment_change(i, val))
            self.wm_spinners.append(sp)
            row.add_widget(sp)
            watermark_card.add_widget(row)

        # 水印位置
        pos_box = BoxLayout(size_hint_y=None, height=40, spacing=8)
        pos_box.add_widget(Label(text="位置：", font_size='12sp', color=THEME['text'], size_hint_x=0.3))
        cur_pos = self.config.get('watermark_position', DEFAULT_CONFIG['watermark_position'])
        cur_pos_label = WATERMARK_POSITION_LABELS.get(cur_pos, '右下')
        self.pos_spinner = Spinner(
            text=cur_pos_label,
            values=list(WATERMARK_POSITION_LABELS.values()),
            size_hint_x=0.7, font_size='12sp',
        )
        pos_box.add_widget(self.pos_spinner)
        watermark_card.add_widget(pos_box)

        # 字号（大/中/小）
        font_size_box = BoxLayout(size_hint_y=None, height=40, spacing=8)
        font_size_box.add_widget(Label(text="字号：", font_size='12sp', color=THEME['text'], size_hint_x=0.3))
        cur_fs = self.config.get('watermark_font_size', DEFAULT_CONFIG['watermark_font_size'])
        if cur_fs not in WATERMARK_FONT_SIZE_OPTIONS:
            cur_fs = '中'
        self.font_spinner = Spinner(
            text=cur_fs,
            values=WATERMARK_FONT_SIZE_OPTIONS,
            size_hint_x=0.7, font_size='12sp',
        )
        font_size_box.add_widget(self.font_spinner)
        watermark_card.add_widget(font_size_box)

        save_wm_btn = Button(text="保存水印设置", font_size='13sp', size_hint_y=None, height=40,
                            background_color=THEME['accent'], background_normal='')
        save_wm_btn.bind(on_release=self._save_watermark)
        watermark_card.add_widget(save_wm_btn)

        content.add_widget(watermark_card)

        # === 关于 ===
        about_card = CardWidget(size_hint_y=None, height=120)
        about_card.add_widget(SectionLabel(text="ℹ️ 关于"))
        about_card.add_widget(Label(text=AUTHOR_INFO, font_size='12sp',
                                    color=THEME['text_dim'], halign='center'))
        content.add_widget(about_card)

        content.add_widget(Label(size_hint_y=None, height=20))

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
            address_general="XX区", address_precise="XX路1号",
            property_type="住宅", seq=1, date_str=get_date_str(), photo_type="远景",
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
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = 68
        self.padding = [6, 4, 6, 4]
        self.spacing = 4

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

        # 信息区域
        info = BoxLayout(spacing=4, size_hint_x=0.6)
        info.padding = [4, 0, 0, 0]

        # 客户名
        name_box = BoxLayout(orientation='vertical')
        name_box.add_widget(Label(text="客户名", font_size='9sp', color=THEME['text_dim'],
                                  size_hint_y=0.3, halign='left', text_size=(None, None)))
        name_box.add_widget(Label(
            text=borrower[:12] + ("…" if len(borrower) > 12 else ""),
            font_size='13sp',
            color=THEME['success'] if self.done else THEME['text'],
            size_hint_y=0.7, halign='left', text_size=(None, None),
        ))
        info.add_widget(name_box)

        # 抵押物地址 + 性质
        prop_box = BoxLayout(orientation='vertical')
        prop_box.add_widget(Label(text="地址/性质", font_size='9sp', color=THEME['text_dim'],
                                  size_hint_y=0.3, halign='left', text_size=(None, None)))
        addr_short = (address_general + address_precise)[:10]
        if len((address_general + address_precise)) > 10:
            addr_short += "…"
        prop_label = addr_short
        if property_type:
            prop_label += " " + property_type[:6]
        prop_box.add_widget(Label(
            text=prop_label, font_size='11sp', color=THEME['text'], size_hint_y=0.7,
            halign='left', text_size=(None, None),
        ))
        info.add_widget(prop_box)

        self.add_widget(info)

        # 按钮区域
        btn_area = BoxLayout(spacing=4, size_hint_x=0.4)

        self.photo_btn = Button(
            text="📷", font_size='17sp', size_hint_x=0.35,
            background_color=THEME['success'] if self.done else THEME['accent'],
            background_normal='',
        )
        self.photo_btn.bind(on_release=self._on_photo)
        btn_area.add_widget(self.photo_btn)

        self.view_btn = Button(
            text="📂%d" % self.photo_count if self.photo_count > 0 else "📂",
            font_size='13sp', size_hint_x=0.35,
            background_color=THEME['accent_dark'] if self.photo_count > 0 else (0.3, 0.3, 0.4, 1),
            background_normal='',
        )
        self.view_btn.bind(on_release=self._on_view_photos)
        btn_area.add_widget(self.view_btn)

        summary = progress_mgr.get_photo_type_summary(progress_key)
        status = Label(text=f"✓{summary}" if self.done else "待拍", font_size='11sp',
                      color=THEME['success'] if self.done else THEME['text_dim'], size_hint_x=0.3)
        btn_area.add_widget(status)

        self.add_widget(btn_area)

    def _update_bg(self, *args):
        self.bg_rect.pos = self.pos
        self.bg_rect.size = self.size

    def _on_photo(self, instance):
        self.photo_callback(self.row_index, self.borrower, self.address_general,
                           self.address_precise, self.property_type)

    def _on_view_photos(self, instance):
        self.view_photos_callback(self.row_index)

    def mark_done(self):
        self.done = True
        self.photo_count = self.progress_mgr.get_photo_count(self.progress_key)
        self.photo_btn.background_color = THEME['success']
        self.view_btn.text = "📂%d" % self.photo_count
        for child in self.children:
            if isinstance(child, BoxLayout):
                for c in child.children:
                    if isinstance(c, Label) and ("✓" in c.text or "待拍" in c.text or "/" in c.text):
                        c.text = f"✓{self.progress_mgr.get_photo_type_summary(self.progress_key)}"
                        c.color = THEME['success']
                        break


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

        main = BoxLayout(orientation='vertical')
        self._build_ui(main)
        self.add_widget(main)

    def _build_ui(self, parent):
        # 顶部标题栏
        title_bar = BoxLayout(size_hint_y=None, height=48, spacing=6, padding=[6, 0, 6, 0])
        title_bar.add_widget(Label(text="资产盘点专项拍照工具", font_size='17sp', bold=True, color=THEME['text'],
                                   size_hint_x=0.6, halign='left', text_size=(None, None)))

        settings_btn = Button(text="⚙", font_size='18sp', size_hint_x=0.15,
                             background_color=THEME['accent'], background_normal='')
        settings_btn.bind(on_release=self._go_settings)
        title_bar.add_widget(settings_btn)

        self.progress_label = Label(text="0/0", font_size='12sp', color=THEME['text_dim'],
                                    size_hint_x=0.25, halign='right', text_size=(None, None))
        title_bar.add_widget(self.progress_label)
        parent.add_widget(title_bar)

        # 工具栏
        toolbar = BoxLayout(size_hint_y=None, height=44, spacing=6, padding=[6, 0, 6, 0])
        open_btn = Button(text="📂 打开Excel", font_size='13sp', size_hint_x=0.35,
                         background_color=THEME['accent'], background_normal='')
        open_btn.bind(on_release=self._show_file_dialog)
        toolbar.add_widget(open_btn)

        self.search_input = TextInput(hint_text="搜索借款人…", multiline=False, font_size='12sp', size_hint_x=0.5)
        self.search_input.bind(text=self._on_search)
        toolbar.add_widget(self.search_input)

        clear_btn = Button(text="×", font_size='16sp', size_hint_x=0.15,
                          background_color=(0.4, 0.4, 0.4, 1), background_normal='')
        clear_btn.bind(on_release=self._clear_search)
        toolbar.add_widget(clear_btn)
        parent.add_widget(toolbar)

        # 列表区域
        self.scroll_view = ScrollView()
        self.list_layout = GridLayout(cols=1, spacing=4, size_hint_y=None, padding=[4, 0, 4, 0])
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        self.scroll_view.add_widget(self.list_layout)
        parent.add_widget(self.scroll_view)

        # 底部状态栏
        footer = BoxLayout(size_hint_y=None, height=48, spacing=6, padding=[6, 0, 6, 0])

        report_btn = Button(text="📄 生成报告", font_size='13sp',
                           background_color=THEME['success'], background_normal='',
                           size_hint_x=0.35)
        report_btn.bind(on_release=self._generate_report)
        footer.add_widget(report_btn)

        self.status_label = Label(text="请打开 Excel 文件", font_size='12sp',
                                  color=THEME['warning'], size_hint_x=0.65)
        footer.add_widget(self.status_label)
        parent.add_widget(footer)

        # 空状态提示
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

    def _show_file_dialog(self, instance):
        content = BoxLayout(orientation='vertical', spacing=8)
        content.add_widget(Label(text="请输入 Excel 文件路径：", size_hint_y=None, height=30))
        path_input = TextInput(text=self.excel_path, multiline=False, font_size='13sp')
        content.add_widget(path_input)
        load_btn = Button(text="加载", size_hint_y=None, height=44,
                         background_color=THEME['accent'], background_normal='')
        popup = Popup(title='选择 Excel 文件', content=content, size_hint=(0.85, 0.4))
        load_btn.bind(on_release=lambda x: self._load_excel(path_input.text, popup))
        content.add_widget(load_btn)
        popup.open()

    def _load_excel(self, path, popup):
        popup.dismiss()
        if not os.path.exists(path):
            self.status_label.text = "文件不存在！"
            self.status_label.color = THEME['danger']
            return
        self.excel_path = path
        try:
            reader = ExcelReader(path)
            self.headers, self.rows = reader.load()
            self.status_label.text = f"✓ 已加载 {len(self.rows)} 条记录"
            self.status_label.color = THEME['success']
            self._refresh_list()
        except Exception as e:
            self.status_label.text = f"加载失败: {e}"
            self.status_label.color = THEME['danger']

    def _refresh_list(self):
        self.list_layout.clear_widgets()
        self.row_widgets = []

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
                rw.height = 68
                rw.opacity = 1
            else:
                rw.height = 0
                rw.opacity = 0

    def _clear_search(self, instance):
        self.search_input.text = ""
        self._on_search(None, "")

    def _go_settings(self, instance):
        self.manager.current = 'settings'

    def _on_photo_request(self, row_index, borrower, address_general, address_precise, property_type):
        self._current_row = row_index
        self._current_borrower = borrower
        self._current_addr_general = address_general
        self._current_addr_precise = address_precise
        self._current_property_type = property_type
        self._current_key = self.progress_mgr._make_key(borrower, (address_general + address_precise).strip())
        popup = PhotoTypePopup(on_select=self._on_photo_type_selected)
        popup.open()

    def _on_photo_type_selected(self, photo_type):
        self._current_photo_type = photo_type
        self.status_label.text = "拍照中 (%s)…" % photo_type
        self.status_label.color = THEME['warning']
        self.camera_mgr.take_photo(self._on_photo_done)

    def _on_photo_done(self, photo_path):
        if photo_path is None:
            self.status_label.text = "拍照失败！"
            self.status_label.color = THEME['danger']
            return

        row_index = self._current_row
        borrower = self._current_borrower
        addr_general = self._current_addr_general
        addr_precise = self._current_addr_precise
        property_type = self._current_property_type
        photo_type = self._current_photo_type
        key = self._current_key

        seq = self.progress_mgr.get_next_photo_index(key) + 1
        date_str = get_date_str()
        datetime_str = get_datetime_str()
        full_addr = (addr_general + addr_precise).strip()

        # 水印/重命名/保存都是耗时 IO，放后台线程避免 ANR 闪退
        self.status_label.text = "处理中…"
        self.status_label.color = THEME['warning']

        config_data = self.config.data
        naming_segments = self.config.get('naming_segments', DEFAULT_CONFIG['naming_segments'])
        lat, lng = self.camera_mgr.gps.get_coords()

        def _process():
            try:
                # 逆地理编码：GPS 坐标 -> 地名（后台线程，不阻塞 UI）
                # 水印「地址名」段使用 GPS 定位地名，无 GPS 显示提示
                place_name = self.camera_mgr.get_location_name(lat, lng)

                # 添加水印（段选择模式）
                PhotoProcessor.add_watermark(
                    photo_path, config_data,
                    time_str=datetime_str, address=place_name,
                    lat=lat, lng=lng,
                )

                # 生成文件名（段选择模式）
                filename = PhotoProcessor.generate_filename(
                    naming_segments, borrower, addr_general, addr_precise,
                    property_type, seq, date_str, photo_type,
                )
                new_path = os.path.join(APP_DIR, filename)
                if photo_path != new_path:
                    if os.path.exists(new_path):
                        name, ext = os.path.splitext(filename)
                        # 避免重名：追加序号
                        suffix = 2
                        while os.path.exists(new_path):
                            new_path = os.path.join(APP_DIR, "%s-%02d%s" % (name, suffix, ext))
                            suffix += 1
                    os.rename(photo_path, new_path)

                PhotoProcessor.save_to_gallery(new_path)
                self.progress_mgr.mark_photo(key, new_path, photo_type)
                Clock.schedule_once(lambda dt: self._on_photo_saved(row_index, filename), 0)
            except Exception as e:
                Logger.error("MainScreen._on_photo_done: %s" % e)
                err_msg = str(e)
                Clock.schedule_once(lambda dt: self._on_photo_failed(err_msg), 0)

        threading.Thread(target=_process, daemon=True).start()

    @mainthread
    def _on_photo_saved(self, row_index, filename):
        self._refresh_row_done(row_index)
        self.status_label.text = "✓ " + filename
        self.status_label.color = THEME['success']

    @mainthread
    def _on_photo_failed(self, err_msg):
        self.status_label.text = "保存失败: %s" % err_msg[:30]
        self.status_label.color = THEME['danger']

    def _refresh_row_done(self, row_index):
        if row_index < len(self.row_widgets):
            self.row_widgets[row_index].mark_done()
        self._update_progress()

    def _on_view_photos(self, row_index):
        if row_index >= len(self.rows):
            return
        row = self.rows[row_index]
        borrower = row[0] if len(row) > 0 else ""
        address = row[1] if len(row) > 1 else ""
        key = self.progress_mgr._make_key(borrower, address)
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
        address = row[1] if len(row) > 1 else ""
        key = self.progress_mgr._make_key(borrower, address)
        if photo_index == -1:
            self.progress_mgr.delete_all_photos(key)
        else:
            self.progress_mgr.delete_photo(key, photo_index)
        Clock.schedule_once(lambda dt: self._refresh_row_done(row_index), 0)

    def _generate_report(self, instance):
        if not self.rows:
            self.status_label.text = "请先加载 Excel 文件！"
            self.status_label.color = THEME['danger']
            return

        self.status_label.text = "生成报告中…"
        self.status_label.color = THEME['warning']

        def _do_generate():
            output_path = self.report_generator.generate(
                (self.headers, self.rows), self.progress_mgr,
            )
            Clock.schedule_once(lambda dt: self._on_report_done(output_path), 0)

        threading.Thread(target=_do_generate).start()

    def _on_report_done(self, output_path):
        if output_path:
            self.status_label.text = "✓ 报告已生成"
            self.status_label.color = THEME['success']
            content = BoxLayout(orientation='vertical', spacing=8)
            content.add_widget(Label(text="报告已生成：", font_size='14sp'))
            content.add_widget(Label(text=output_path, font_size='11sp',
                                    color=THEME['success'], text_size=(400, None)))
            if IS_ANDROID:
                content.add_widget(Label(text="文件已保存到设备", font_size='11sp'))
            popup = Popup(title='报告生成完成', content=content, size_hint=(0.8, 0.4))
            close_btn = Button(text="确定", size_hint_y=None, height=40)
            close_btn.bind(on_release=popup.dismiss)
            content.add_widget(close_btn)
            popup.open()
        else:
            self.status_label.text = "报告生成失败！"
            self.status_label.color = THEME['danger']


# ============================================================
# App 入口
# ============================================================

class LoanPhotoApp(App):
    def build(self):
        self.title = "资产盘点专项拍照工具"
        Window.clearcolor = THEME['bg']

        self.config = AppConfig()

        sm = ScreenManager(transition=SlideTransition(duration=0.25))
        sm.add_widget(WelcomeScreen(name='welcome'))
        sm.add_widget(MainScreen(app_config=self.config, name='main'))
        sm.add_widget(SettingsScreen(app_config=self.config, name='settings'))
        sm.current = 'welcome'
        return sm

    def on_start(self):
        if IS_ANDROID:
            self._request_permissions()

    def _request_permissions(self):
        try:
            if ANDROID_API >= 33:
                perms = [Permission.CAMERA, Permission.ACCESS_FINE_LOCATION,
                         Permission.ACCESS_COARSE_LOCATION, Permission.READ_MEDIA_IMAGES]
            elif ANDROID_API >= 30:
                perms = [Permission.CAMERA, Permission.ACCESS_FINE_LOCATION,
                         Permission.ACCESS_COARSE_LOCATION, Permission.READ_EXTERNAL_STORAGE]
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

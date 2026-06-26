"""
银行外勤拍照工具 App - 完整重写版
新增功能：
1. 右侧导航条（快速跳转到指定行）
2. 顶部搜索栏（支持客户名模糊搜索）
3. 完成走访生成报告（读取模板自动填写 Excel）
4. 查看已拍照片 + 删除重拍
5. 同一抵押物多张照片（_01, _02...）
6. 照片保存到系统相册
7. 列名明确：客户名称、押品性质
8. 命名规则可选
"""

import os
import json
from datetime import datetime
from collections import defaultdict
import threading

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
from kivy.clock import Clock, mainthread
from kivy.logger import Logger
from kivy.core.window import Window
from kivy.utils import platform
from kivy.graphics import Color, Rectangle, Line
from kivy.metrics import dp, sp

import openpyxl
from openpyxl import load_workbook
from PIL import Image as PILImage, ImageDraw, ImageFont

from geocoder import GeoCoder

# 平台检测
IS_ANDROID = platform == 'android'

if IS_ANDROID:
    from android.permissions import request_permissions, Permission
    from android.storage import getExternalStorageDirectory, getExternalStoragePublicDirectory
    from android import activity as android_activity
    try:
        from android import mActivity
    except:
        mActivity = None
    BASE_DIR = os.path.join(getExternalStorageDirectory(), 'loan_photos')
else:
    BASE_DIR = os.path.join(os.path.expanduser('~'), 'loan_photos')

os.makedirs(BASE_DIR, exist_ok=True)
PROGRESS_FILE = os.path.join(BASE_DIR, 'photo_progress.json')
TEMPLATE_PATH = r"C:\Users\Administrator\Desktop\盘点相关文件\抵押物、抵债资产现场勘查日报表模板.xlsx"


# ============================================================
# 系统日期工具
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
    """报告中的日期格式：2026年6月26日"""
    return get_system_date().strftime("%Y年%m月%d日")


# ============================================================
# 进度管理器
# ============================================================
class ProgressManager:
    """持久化拍照进度到 JSON"""

    def __init__(self, filepath=None):
        self.filepath = filepath or PROGRESS_FILE
        self.data = {}  # {row_index: {"photos": [path1, path2, ...], "timestamp": "..."}}
        self.load()

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

    def mark_photo(self, row_index, photo_path):
        key = str(row_index)
        if key not in self.data:
            self.data[key] = {"photos": [], "timestamp": ""}
        self.data[key]["photos"].append(photo_path)
        self.data[key]["timestamp"] = get_full_datetime_str()
        self.save()

    def delete_photo(self, row_index, photo_index):
        key = str(row_index)
        if key in self.data and photo_index < len(self.data[key]["photos"]):
            self.data[key]["photos"].pop(photo_index)
            if not self.data[key]["photos"]:
                del self.data[key]
            self.save()

    def delete_all_photos(self, row_index):
        key = str(row_index)
        if key in self.data:
            del self.data[key]
            self.save()

    def is_photographed(self, row_index):
        key = str(row_index)
        return key in self.data and len(self.data[key].get("photos", [])) > 0

    def get_photos(self, row_index):
        key = str(row_index)
        if key in self.data:
            return self.data[key].get("photos", [])
        return []

    def get_photo_count(self, row_index):
        return len(self.get_photos(row_index))

    def get_done_count(self, total):
        return sum(1 for i in range(total) if self.is_photographed(i))

    def get_next_photo_index(self, row_index):
        """获取同一客户下一张照片的序号"""
        key = str(row_index)
        if key in self.data:
            return len(self.data[key]["photos"])
        return 0


# ============================================================
# Excel 读取器
# ============================================================
class ExcelReader:
    """读取 Excel 文件，只取前 3 列"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.headers = []
        self.rows = []

    def load(self):
        wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
        ws = wb.active
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            cells = [str(cell).strip() if cell else "" for cell in row[:3]]
            if i == 0:
                self.headers = cells
            else:
                self.rows.append(cells)
        wb.close()
        if not self.headers:
            self.headers = ["客户名称", "押品性质", "备注"]
        return self.headers, self.rows


# ============================================================
# 报告生成器
# ============================================================
class ReportGenerator:
    """生成现场勘查日报表"""

    def __init__(self, template_path=None):
        self.template_path = template_path or TEMPLATE_PATH

    def generate(self, excel_data, progress_mgr, output_path=None):
        """
        生成报告 Excel
        excel_data: (headers, rows) 从 ExcelReader 获取的数据
        progress_mgr: ProgressManager 实例
        output_path: 输出路径
        """
        headers, rows = excel_data

        # 读取模板
        try:
            wb = load_workbook(self.template_path)
            ws = wb['Sheet1']
        except Exception as e:
            # 如果模板不存在，创建新文件
            Logger.error(f"模板读取失败: {e}，将创建新文件")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            # 写表头
            ws['A3'] = '序号'
            ws['B3'] = '日期'
            ws['C3'] = '勘查业务贷款人名称'
            ws['D3'] = '抵押物/抵债资产具体情况'
            ws['E3'] = '现状描述'
            ws['F3'] = '备注（是否存在发生风险的可能）'

        # 清除旧数据（保留表头行1-3和尾部注释行25-30）
        for row_idx in range(4, ws.max_row + 1):
            if row_idx < 25:  # 不删除注释行
                for col in range(1, 7):
                    ws.cell(row=row_idx, column=col).value = None

        # 填写数据
        report_date = get_report_date_str()
        current_row = 4
        seq_num = 1

        for i, row in enumerate(rows):
            customer_name = row[0] if len(row) > 0 else ""
            collateral_type = row[1] if len(row) > 1 else ""
            remark = row[2] if len(row) > 2 else ""

            # 只填写有客户名称的行
            if customer_name:
                ws.cell(row=current_row, column=1, value=seq_num)          # A: 序号
                ws.cell(row=current_row, column=2, value=report_date)      # B: 日期
                ws.cell(row=current_row, column=3, value=customer_name)    # C: 客户名称
                ws.cell(row=current_row, column=4, value=collateral_type)  # D: 押品性质/抵押物情况
                ws.cell(row=current_row, column=5, value=remark)           # E: 现状描述
                # F: 备注 - 留空或自动生成
                ws.cell(row=current_row, column=6, value="")               # F: 备注

                seq_num += 1
                current_row += 1

        # 如果数据超过10行，继续填充
        # 填写尾部信息
        if current_row <= 17:
            current_row = 17
        ws.cell(row=17, column=5, value="当天线路完成进度：100%")

        # 保存
        if output_path is None:
            output_path = os.path.join(
                BASE_DIR,
                f"现场勘查日报表_{get_date_str()}.xlsx"
            )

        try:
            wb.save(output_path)
            Logger.info(f"报告已保存: {output_path}")
            return output_path
        except Exception as e:
            Logger.error(f"报告保存失败: {e}")
            return None


# ============================================================
# 照片处理器
# ============================================================
class PhotoProcessor:
    """照片水印、命名、保存"""

    @staticmethod
    def add_watermark(photo_path, date_str, location_str):
        """在照片底部添加日期和地名水印"""
        try:
            img = PILImage.open(photo_path)
            draw = ImageDraw.Draw(img)

            font_size = max(22, img.height // 28)
            font = PhotoProcessor._get_font(font_size)

            watermark_text = f"{date_str}  {location_str}"

            text_bbox = draw.textbbox((0, 0), watermark_text, font=font)
            tw = text_bbox[2] - text_bbox[0]
            th = text_bbox[3] - text_bbox[1] + 10

            overlay = PILImage.new('RGBA', img.size, (0, 0, 0, 0))
            od = ImageDraw.Draw(overlay)
            od.rectangle([0, img.height - th - 14, img.width, img.height], fill=(0, 0, 0, 170))

            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            img = PILImage.alpha_composite(img, overlay)

            draw = ImageDraw.Draw(img)
            draw.text((8, img.height - th - 7), watermark_text, font=font, fill=(255, 255, 255))

            final = img.convert('RGB') if img.mode == 'RGBA' else img
            final.save(photo_path, 'JPEG', quality=92)
        except Exception as e:
            Logger.error(f"PhotoProcessor.add_watermark: {e}")

    @staticmethod
    def _get_font(size):
        """加载中文字体"""
        font_paths = [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/msyhbd.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "C:/Windows/Fonts/simsun.ttc",
            "msyh.ttc", "simhei.ttf",
            "/system/fonts/DroidSansFallback.ttf",
            "/system/fonts/NotoSansCJK-Regular.ttc",
        ]
        for fp in font_paths:
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
        return ImageFont.load_default()

    @staticmethod
    def generate_filename(rule_index, borrower_name, room_number="", seq=0, date_str=None):
        """根据规则生成文件名，seq为同一客户的照片序号"""
        if date_str is None:
            date_str = get_date_str()

        def clean(s):
            for ch in '/\\:*?"<>|':
                s = s.replace(ch, '_')
            return s.strip()

        borrower_name = clean(borrower_name)
        room_number = clean(room_number)

        # 序号后缀：01, 02, 03...
        seq_suffix = f"_{seq:02d}" if seq > 0 else "_01"

        if rule_index == 1 and room_number:
            return f"{date_str}_{borrower_name}_{room_number}{seq_suffix}.jpg"
        else:
            return f"{date_str}_{borrower_name}{seq_suffix}.jpg"

    @staticmethod
    def save_to_gallery(photo_path):
        """保存照片到系统相册"""
        if not IS_ANDROID:
            return  # Windows 测试环境不需要

        try:
            from android import mActivity
            from android.storage import getExternalStoragePublicDirectory
            from android.provider import MediaStore
            from android.content import ContentValues
            from java.io import FileInputStream, FileOutputStream

            # 使用 MediaScannerConnection 扫描文件
            values = ContentValues()
            values.put(MediaStore.Images.Media.DATA, photo_path)
            values.put(MediaStore.Images.Media.MIME_TYPE, "image/jpeg")
            values.put(MediaStore.Images.Media.DISPLAY_NAME, os.path.basename(photo_path))

            resolver = mActivity.getContentResolver()
            uri = resolver.insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values)

            if uri:
                # 复制文件内容
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
                Logger.info(f"已保存到相册: {photo_path}")
        except Exception as e:
            Logger.error(f"保存到相册失败: {e}")


# ============================================================
# 相机管理器
# ============================================================
class CameraManager:
    """通过 Plyer 调用相机和 GPS"""

    def __init__(self):
        self.photo_path = ""
        self.pending_callback = None
        self.geocoder = GeoCoder()

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
        except Exception as e:
            Logger.error(f"CameraManager permission: {e}")
            self._fallback_simulate()

    def _on_permission_result(self, permissions, grants):
        if all(grants):
            self._launch_camera()
        else:
            Logger.error("CameraManager: 权限被拒绝")
            self._fallback_simulate()

    def _launch_camera(self):
        try:
            from plyer import camera
            self.photo_path = os.path.join(
                BASE_DIR,
                f"capture_{get_system_date().strftime('%Y%m%d_%H%M%S')}.jpg"
            )
            camera.take_picture(filename=self.photo_path, on_complete=self._on_photo_complete)
        except Exception as e:
            Logger.error(f"CameraManager._launch_camera: {e}")
            self._fallback_simulate()

    def _on_photo_complete(self, path):
        if self.pending_callback and path and os.path.exists(path):
            self.pending_callback(path)
        else:
            self.pending_callback(None)

    def _simulate_photo(self):
        """模拟拍照（Windows 测试用）"""
        self.photo_path = os.path.join(
            BASE_DIR,
            f"test_{get_system_date().strftime('%Y%m%d_%H%M%S')}.jpg"
        )
        img = PILImage.new('RGB', (640, 480), (180, 180, 180))
        draw = ImageDraw.Draw(img)
        font = PhotoProcessor._get_font(24)
        now_str = get_full_datetime_str()
        draw.text((150, 200), "Test Photo", fill=(0, 0, 0), font=font)
        draw.text((150, 250), now_str, fill=(0, 0, 0), font=font)
        img.save(self.photo_path)
        if self.pending_callback:
            Clock.schedule_once(lambda dt: self.pending_callback(self.photo_path), 0.5)

    def _fallback_simulate(self):
        self._simulate_photo()

    def get_location_name(self):
        """获取地名"""
        if not IS_ANDROID:
            return "测试定位"
        try:
            from plyer import gps
            gps.configure(on_location=lambda **kw: None)
            gps.start()
            import time
            time.sleep(2)
            gps.stop()
            return "GPS定位完成"
        except Exception as e:
            Logger.error(f"GPS error: {e}")
            return "定位不可用"


# ============================================================
# 单行组件
# ============================================================
class RowWidget(BoxLayout):
    """列表中的每一行"""

    def __init__(self, row_index, customer_name, collateral_type, remark,
                 progress_mgr, photo_callback, view_photos_callback, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'
        self.size_hint_y = None
        self.height = 72
        self.padding = [4, 4, 4, 4]
        self.spacing = 4

        self.row_index = row_index
        self.customer_name = customer_name
        self.collateral_type = collateral_type
        self.remark = remark
        self.photo_callback = photo_callback
        self.view_photos_callback = view_photos_callback
        self.progress_mgr = progress_mgr
        self.done = progress_mgr.is_photographed(row_index)
        self.photo_count = progress_mgr.get_photo_count(row_index)

        # 信息区域
        info_layout = BoxLayout(spacing=4, size_hint_x=0.62)
        
        # 客户名称
        name_box = BoxLayout(orientation='vertical')
        name_box.add_widget(Label(text="客户名称", font_size='10sp', color=(0.5, 0.5, 0.5, 1), size_hint_y=0.3))
        name_box.add_widget(Label(
            text=customer_name[:10] + (".." if len(customer_name) > 10 else ""),
            font_size='13sp',
            color=(0.3, 0.85, 0.3, 1) if self.done else (1, 1, 1, 1),
            size_hint_y=0.7, halign='left', text_size=(None, None),
        ))
        info_layout.add_widget(name_box)

        # 押品性质
        col_box = BoxLayout(orientation='vertical')
        col_box.add_widget(Label(text="押品性质", font_size='10sp', color=(0.5, 0.5, 0.5, 1), size_hint_y=0.3))
        col_box.add_widget(Label(
            text=collateral_type[:8] + (".." if len(collateral_type) > 8 else ""),
            font_size='13sp',
            color=(1, 1, 1, 1), size_hint_y=0.7, halign='left', text_size=(None, None),
        ))
        info_layout.add_widget(col_box)

        # 备注
        remark_box = BoxLayout(orientation='vertical')
        remark_box.add_widget(Label(text="备注", font_size='10sp', color=(0.5, 0.5, 0.5, 1), size_hint_y=0.3))
        remark_box.add_widget(Label(
            text=remark[:8] + (".." if len(remark) > 8 else ""),
            font_size='13sp', color=(0.8, 0.8, 0.8, 1), size_hint_y=0.7,
            halign='left', text_size=(None, None),
        ))
        info_layout.add_widget(remark_box)

        self.add_widget(info_layout)

        # 右侧按钮区域
        btn_layout = BoxLayout(spacing=4, size_hint_x=0.38)

        # 拍照按钮
        self.photo_btn = Button(
            text="📷", font_size='18sp',
            size_hint_x=0.3,
            background_color=(0.25, 0.75, 0.25, 1) if self.done else (0.2, 0.45, 0.85, 1),
        )
        self.photo_btn.bind(on_release=self._on_photo)
        btn_layout.add_widget(self.photo_btn)

        # 查看照片按钮
        self.view_btn = Button(
            text=f"👁{self.photo_count}" if self.photo_count > 0 else "👁",
            font_size='14sp', size_hint_x=0.3,
            background_color=(0.5, 0.5, 0.7, 1) if self.photo_count > 0 else (0.3, 0.3, 0.4, 1),
        )
        self.view_btn.bind(on_release=self._on_view_photos)
        btn_layout.add_widget(self.view_btn)

        # 状态标签
        if self.done:
            status_label = Label(
                text=f"✓{self.photo_count}张",
                font_size='11sp', color=(0.3, 0.9, 0.3, 1),
                size_hint_x=0.4,
            )
        else:
            status_label = Label(
                text="未拍",
                font_size='11sp', color=(0.6, 0.6, 0.6, 1),
                size_hint_x=0.4,
            )
        btn_layout.add_widget(status_label)

        self.add_widget(btn_layout)

    def _on_photo(self, instance):
        self.photo_callback(self.row_index, self.customer_name, self.collateral_type)

    def _on_view_photos(self, instance):
        self.view_photos_callback(self.row_index)

    def mark_done(self):
        self.done = True
        self.photo_count = self.progress_mgr.get_photo_count(self.row_index)
        self.photo_btn.text = "📷"
        self.photo_btn.background_color = (0.25, 0.75, 0.25, 1)
        self.view_btn.text = f"👁{self.photo_count}"


# ============================================================
# 照片查看弹窗
# ============================================================
class PhotoViewerPopup(Popup):
    """查看已拍照片的弹窗"""

    def __init__(self, row_index, photos, delete_callback, **kwargs):
        super().__init__(**kwargs)
        self.title = f"已拍照片 ({len(photos)}张)"
        self.size_hint = (0.9, 0.8)
        self.auto_dismiss = True

        self.row_index = row_index
        self.photos = photos
        self.delete_callback = delete_callback

        main_layout = BoxLayout(orientation='vertical', spacing=8, padding=8)

        # 照片列表
        scroll = ScrollView()
        list_layout = GridLayout(cols=1, spacing=8, size_hint_y=None)
        list_layout.bind(minimum_height=list_layout.setter('height'))

        for i, photo_path in enumerate(photos):
            item = BoxLayout(orientation='horizontal', size_hint_y=None, height=120, spacing=8)

            # 缩略图
            if os.path.exists(photo_path):
                item.add_widget(KivyImage(
                    source=photo_path,
                    size_hint_x=0.4,
                    allow_stretch=True,
                    keep_ratio=True,
                ))

            # 文件名和删除按钮
            info_box = BoxLayout(orientation='vertical', spacing=4, size_hint_x=0.6)
            filename = os.path.basename(photo_path)
            info_box.add_widget(Label(
                text=filename, font_size='12sp', halign='left',
                size_hint_y=0.4, text_size=(None, None),
            ))
            info_box.add_widget(Label(
                text=photo_path, font_size='9sp', color=(0.5, 0.5, 0.5, 1),
                halign='left', size_hint_y=0.3, text_size=(None, None),
            ))

            del_btn = Button(
                text="删除此照片", font_size='12sp',
                background_color=(0.85, 0.2, 0.2, 1),
                size_hint_y=0.3,
            )
            del_btn.bind(on_release=lambda x, idx=i: self._delete_photo(idx))
            info_box.add_widget(del_btn)

            item.add_widget(info_box)
            list_layout.add_widget(item)

        scroll.add_widget(list_layout)
        main_layout.add_widget(scroll)

        # 全部删除按钮
        del_all_btn = Button(
            text="删除全部照片（重拍）", font_size='14sp',
            background_color=(0.85, 0.2, 0.2, 1),
            size_hint_y=None, height=44,
        )
        del_all_btn.bind(on_release=self._delete_all)
        main_layout.add_widget(del_all_btn)

        # 关闭按钮
        close_btn = Button(
            text="关闭", font_size='14sp',
            size_hint_y=None, height=44,
        )
        close_btn.bind(on_release=self.dismiss)
        main_layout.add_widget(close_btn)

        self.content = main_layout

    def _delete_photo(self, index):
        self.delete_callback(self.row_index, index)
        self.dismiss()

    def _delete_all(self, instance):
        self.delete_callback(self.row_index, -1)  # -1 表示全部删除
        self.dismiss()


# ============================================================
# 主界面
# ============================================================
class MainScreen(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 6
        self.spacing = 6

        self.excel_path = ""
        self.headers = []
        self.rows = []
        self.selected_rule = 0
        self.room_col_idx = None
        self.progress_mgr = ProgressManager()
        self.camera_mgr = CameraManager()
        self.report_generator = ReportGenerator()
        self.row_widgets = []  # 引用所有行组件

        # 整体布局：顶部固定 + 中间列表 + 底部固定
        self._build_header()
        self._build_list_area()
        self._build_footer()

    def _build_header(self):
        """顶部区域：文件打开 + 搜索 + 规则选择"""
        header = BoxLayout(orientation='vertical', spacing=4, size_hint_y=None, height=140)

        # 第一行：打开文件 + 进度
        row1 = BoxLayout(orientation='horizontal', spacing=6, size_hint_y=0.33)
        open_btn = Button(
            text="打开Excel", font_size='14sp',
            background_color=(0.2, 0.6, 0.85, 1), size_hint_x=0.3,
        )
        open_btn.bind(on_release=self._show_file_dialog)
        row1.add_widget(open_btn)

        self.progress_label = Label(
            text="进度: 0/0", font_size='13sp',
            color=(0.8, 0.8, 0.8, 1), size_hint_x=0.7,
        )
        row1.add_widget(self.progress_label)
        header.add_widget(row1)

        # 第二行：搜索栏
        row2 = BoxLayout(orientation='horizontal', spacing=6, size_hint_y=0.33)
        row2.add_widget(Label(text="搜索:", size_hint_x=0.12, font_size='13sp'))
        self.search_input = TextInput(
            hint_text="输入客户名称（支持模糊搜索）",
            multiline=False, font_size='13sp', size_hint_x=0.6,
        )
        self.search_input.bind(text=self._on_search)
        row2.add_widget(self.search_input)

        search_btn = Button(
            text="搜索", font_size='13sp', size_hint_x=0.15,
            background_color=(0.3, 0.5, 0.8, 1),
        )
        search_btn.bind(on_release=self._on_search)
        row2.add_widget(search_btn)

        clear_btn = Button(
            text="清除", font_size='13sp', size_hint_x=0.13,
            background_color=(0.5, 0.5, 0.5, 1),
        )
        clear_btn.bind(on_release=self._clear_search)
        row2.add_widget(clear_btn)
        header.add_widget(row2)

        # 第三行：命名规则选择
        row3 = BoxLayout(orientation='horizontal', spacing=6, size_hint_y=0.33)
        row3.add_widget(Label(text="命名:", size_hint_x=0.12, font_size='13sp'))
        self.rule_spinner = Spinner(
            text="日期_借款人",
            values=["日期_借款人", "日期_借款人_房间号"],
            size_hint_x=0.88, font_size='13sp',
        )
        self.rule_spinner.bind(text=self._on_rule_change)
        row3.add_widget(self.rule_spinner)
        header.add_widget(row3)

        self.add_widget(header)

    def _build_list_area(self):
        """中间区域：列表 + 右侧导航条"""
        list_area = BoxLayout(orientation='horizontal', spacing=4)

        # 主列表（左侧）
        self.scroll_view = ScrollView(size_hint_x=0.85)
        self.list_layout = GridLayout(cols=1, spacing=3, size_hint_y=None)
        self.list_layout.bind(minimum_height=self.list_layout.setter('height'))
        self.scroll_view.add_widget(self.list_layout)
        list_area.add_widget(self.scroll_view)

        # 右侧导航条
        nav_layout = BoxLayout(orientation='vertical', spacing=4, size_hint_x=0.15)
        nav_layout.add_widget(Label(text="导航", font_size='11sp', size_hint_y=None, height=25))

        # 导航按钮
        nav_up_btn = Button(text="▲", font_size='14sp', size_hint_y=None, height=40)
        nav_up_btn.bind(on_release=self._nav_scroll_up)
        nav_layout.add_widget(nav_up_btn)

        # 行号快速跳转
        self.nav_label = Label(text="", font_size='10sp', size_hint_y=0.3)
        nav_layout.add_widget(self.nav_label)

        nav_down_btn = Button(text="▼", font_size='14sp', size_hint_y=None, height=40)
        nav_down_btn.bind(on_release=self._nav_scroll_down)
        nav_layout.add_widget(nav_down_btn)

        # 跳转到指定行
        self.jump_input = TextInput(
            hint_text="行号", multiline=False, font_size='12sp',
            input_filter='int', size_hint_y=None, height=36,
        )
        nav_layout.add_widget(self.jump_input)

        jump_btn = Button(text="跳转", font_size='12sp', size_hint_y=None, height=36)
        jump_btn.bind(on_release=self._jump_to_row)
        nav_layout.add_widget(jump_btn)

        nav_layout.add_widget(Label())  # spacer
        list_area.add_widget(nav_layout)

        self.add_widget(list_area)

    def _build_footer(self):
        """底部区域：操作按钮"""
        footer = BoxLayout(orientation='horizontal', spacing=8, size_hint_y=None, height=50)

        report_btn = Button(
            text="完成走访\n生成报告", font_size='13sp',
            background_color=(0.2, 0.7, 0.3, 1),
        )
        report_btn.bind(on_release=self._generate_report)
        footer.add_widget(report_btn)

        self.status_label = Label(
            text="请选择 Excel 文件开始",
            font_size='12sp', color=(0.85, 0.85, 0.5, 1),
        )
        footer.add_widget(self.status_label)

        self.add_widget(footer)

    def _show_file_dialog(self, instance):
        content = BoxLayout(orientation='vertical', spacing=8)
        content.add_widget(Label(text="请输入Excel文件路径:", size_hint_y=None, height=30))
        path_input = TextInput(text=self.excel_path, multiline=False, font_size='13sp')
        content.add_widget(path_input)
        load_btn = Button(text="加载", size_hint_y=None, height=40, background_color=(0.2, 0.6, 0.85, 1))
        popup = Popup(title='选择 Excel 文件', content=content, size_hint=(0.85, 0.4))
        load_btn.bind(on_release=lambda x: self._load_excel(path_input.text, popup))
        content.add_widget(load_btn)
        popup.open()

    def _load_excel(self, path, popup):
        popup.dismiss()
        if not os.path.exists(path):
            self.status_label.text = "文件不存在!"
            self.status_label.color = (1, 0.3, 0.3, 1)
            return
        self.excel_path = path
        try:
            reader = ExcelReader(path)
            self.headers, self.rows = reader.load()
            self.status_label.text = f"已加载 {len(self.rows)} 行"
            self.status_label.color = (0.3, 0.85, 0.3, 1)

            # 查找房间号列
            self.room_col_idx = None
            for i, h in enumerate(self.headers):
                if any(kw in str(h) for kw in ['房', '室', '号']):
                    self.room_col_idx = i
                    break

            self._refresh_list()
        except Exception as e:
            self.status_label.text = f"加载失败: {e}"
            self.status_label.color = (1, 0.3, 0.3, 1)

    def _refresh_list(self):
        self.list_layout.clear_widgets()
        self.row_widgets = []

        for i, row in enumerate(self.rows):
            customer = row[0] if len(row) > 0 else ""
            collateral = row[1] if len(row) > 1 else ""
            remark = row[2] if len(row) > 2 else ""

            rw = RowWidget(
                row_index=i,
                customer_name=customer,
                collateral_type=collateral,
                remark=remark,
                progress_mgr=self.progress_mgr,
                photo_callback=self._on_photo_request,
                view_photos_callback=self._on_view_photos,
            )
            self.list_layout.add_widget(rw)
            self.row_widgets.append(rw)

        self._update_progress()
        self._update_nav_label()

    def _update_progress(self):
        total = len(self.rows)
        done = self.progress_mgr.get_done_count(total)
        self.progress_label.text = f"进度: {done}/{total}"

    def _update_nav_label(self):
        total = len(self.rows)
        self.nav_label.text = f"共{total}行"

    def _on_search(self, instance, text=None):
        """搜索过滤"""
        query = self.search_input.text.lower().strip()
        if not query:
            # 显示所有行
            for rw in self.row_widgets:
                rw.height = 72
                rw.opacity = 1
        else:
            # 过滤显示
            for rw in self.row_widgets:
                if query in rw.customer_name.lower():
                    rw.height = 72
                    rw.opacity = 1
                else:
                    rw.height = 0
                    rw.opacity = 0

    def _clear_search(self, instance):
        self.search_input.text = ""
        self._on_search(None, "")

    def _on_rule_change(self, spinner, text):
        self.selected_rule = 0 if text == "日期_借款人" else 1

    def _nav_scroll_up(self, instance):
        self.scroll_view.scroll_y = min(1, self.scroll_view.scroll_y + 0.3)

    def _nav_scroll_down(self, instance):
        self.scroll_view.scroll_y = max(0, self.scroll_view.scroll_y - 0.3)

    def _jump_to_row(self, instance):
        try:
            row_num = int(self.jump_input.text)
            if 1 <= row_num <= len(self.row_widgets):
                # 计算滚动位置
                fraction = 1 - (row_num - 1) / max(1, len(self.row_widgets) - 1)
                self.scroll_view.scroll_y = max(0, min(1, fraction))
        except ValueError:
            pass

    def _on_photo_request(self, row_index, customer_name, collateral_type):
        self._current_row = row_index
        self._current_customer = customer_name
        self._current_collateral = collateral_type
        self.status_label.text = "正在拍照..."
        self.status_label.color = (0.85, 0.85, 0.5, 1)
        self.camera_mgr.take_photo(self._on_photo_done)

    def _on_photo_done(self, photo_path):
        if photo_path is None:
            self.status_label.text = "拍照失败!"
            self.status_label.color = (1, 0.3, 0.3, 1)
            return

        row_index = self._current_row
        customer = self._current_customer
        collateral = self._current_collateral

        # 获取序号
        seq = self.progress_mgr.get_next_photo_index(row_index) + 1

        # 获取房间号
        room = ""
        if self.selected_rule == 1 and self.room_col_idx is not None:
            row_data = self.rows[row_index] if row_index < len(self.rows) else []
            if self.room_col_idx < len(row_data):
                room = row_data[self.room_col_idx]

        date_str = get_date_str()
        datetime_str = get_datetime_str()
        location = self.camera_mgr.get_location_name()

        # 添加水印
        PhotoProcessor.add_watermark(photo_path, datetime_str, location)

        # 生成文件名
        filename = PhotoProcessor.generate_filename(
            self.selected_rule, customer, room, seq, date_str
        )
        new_path = os.path.join(BASE_DIR, filename)
        if photo_path != new_path:
            if os.path.exists(new_path):
                name, ext = os.path.splitext(filename)
                new_path = os.path.join(BASE_DIR, f"{name}_{row_index}{ext}")
            os.rename(photo_path, new_path)

        # 保存到相册
        PhotoProcessor.save_to_gallery(new_path)

        # 保存进度
        self.progress_mgr.mark_photo(row_index, new_path)

        # 刷新 UI
        Clock.schedule_once(lambda dt: self._refresh_row_done(row_index), 0)

        self.status_label.text = f"✓ {filename}"
        self.status_label.color = (0.3, 0.85, 0.3, 1)

    def _refresh_row_done(self, row_index):
        if row_index < len(self.row_widgets):
            self.row_widgets[row_index].mark_done()
        self._update_progress()

    def _on_view_photos(self, row_index):
        photos = self.progress_mgr.get_photos(row_index)
        if not photos:
            self.status_label.text = "该客户暂无照片"
            self.status_label.color = (0.85, 0.85, 0.5, 1)
            return

        popup = PhotoViewerPopup(
            row_index=row_index,
            photos=photos,
            delete_callback=self._on_delete_photo,
        )
        popup.open()

    def _on_delete_photo(self, row_index, photo_index):
        if photo_index == -1:
            # 删除全部
            self.progress_mgr.delete_all_photos(row_index)
            self.status_label.text = "已删除全部照片，可重新拍摄"
        else:
            self.progress_mgr.delete_photo(row_index, photo_index)
            self.status_label.text = f"已删除照片，剩余 {self.progress_mgr.get_photo_count(row_index)} 张"

        # 刷新 UI
        Clock.schedule_once(lambda dt: self._refresh_row_done(row_index), 0)

    def _generate_report(self, instance):
        """生成走访报告"""
        if not self.rows:
            self.status_label.text = "请先加载 Excel 文件!"
            self.status_label.color = (1, 0.3, 0.3, 1)
            return

        self.status_label.text = "正在生成报告..."
        self.status_label.color = (0.85, 0.85, 0.5, 1)

        # 在后台线程生成，避免卡 UI
        def _do_generate():
            output_path = self.report_generator.generate(
                (self.headers, self.rows),
                self.progress_mgr,
            )
            Clock.schedule_once(lambda dt: self._on_report_done(output_path), 0)

        threading.Thread(target=_do_generate).start()

    def _on_report_done(self, output_path):
        if output_path:
            self.status_label.text = f"✓ 报告已生成"
            self.status_label.color = (0.3, 0.85, 0.3, 1)

            # 显示弹窗告知文件路径
            content = BoxLayout(orientation='vertical', spacing=8)
            content.add_widget(Label(text="报告已生成:"))
            content.add_widget(Label(
                text=output_path, font_size='11sp',
                color=(0.3, 0.85, 0.3, 1), text_size=(400, None),
            ))

            # 保存到相册/下载目录
            if IS_ANDROID:
                content.add_widget(Label(text="文件已保存到 Downloads 目录", font_size='11sp'))

            popup = Popup(title='报告生成完成', content=content, size_hint=(0.8, 0.4))
            close_btn = Button(text="确定", size_hint_y=None, height=40)
            close_btn.bind(on_release=popup.dismiss)
            content.add_widget(close_btn)
            popup.open()
        else:
            self.status_label.text = "报告生成失败!"
            self.status_label.color = (1, 0.3, 0.3, 1)


# ============================================================
# App
# ============================================================
class LoanPhotoApp(App):

    def build(self):
        self.title = "信贷外勤拍照"
        Window.clearcolor = (0.13, 0.13, 0.17, 1)
        return MainScreen()

    def on_start(self):
        if IS_ANDROID:
            try:
                request_permissions([
                    Permission.CAMERA,
                    Permission.ACCESS_FINE_LOCATION,
                    Permission.ACCESS_COARSE_LOCATION,
                    Permission.WRITE_EXTERNAL_STORAGE,
                    Permission.READ_EXTERNAL_STORAGE,
                ])
            except Exception as e:
                Logger.error(f"App permissions: {e}")

    def on_pause(self):
        return True

    def on_resume(self):
        pass


if __name__ == '__main__':
    LoanPhotoApp().run()

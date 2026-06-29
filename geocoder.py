"""
地名解析模块
v3.18.1: 支持多源解析：
  1. Android 系统 Geocoder（无需 key，小米/华为等国产机通常接入百度/高德后端）
  2. 用户配置的在线地图 API（高德 / 百度，需用户配置 key）
  3. 离线 fallback：内置中国区县坐标表 + 最近邻匹配

说明：小米自带相机水印上的文字地址，本质上就是系统通过百度/高德服务把 GPS 坐标
反解成地址文字。我们在 App 中通过 Android Geocoder 即可复用同一份系统能力，
无需集成百度地图 SDK（用户只要文字信息）。
"""

import os
import json
import math
import requests
import logging

# 安全日志：优先用 kivy.logger，回退到标准 logging（避免 NameError 闪退）
try:
    from kivy.logger import Logger
except Exception:
    Logger = logging.getLogger("geocoder")

# 高德地图逆地理编码 API
AMAP_API_KEY = ""
AMAP_GEOCODE_URL = "https://restapi.amap.com/v3/geocode/regeo"

# 百度地图逆地理编码 API
BAIDU_API_KEY = ""
BAIDU_GEOCODE_URL = "https://api.map.baidu.com/reverse_geocoding/v3/"

# 离线数据文件路径
OFFLINE_DB_PATH = os.path.join(os.path.dirname(__file__), 'data', 'locations.json')

# 离线区县坐标数据库（精简版，覆盖全国主要区县）
DEFAULT_LOCATIONS = [
    # 北京
    {"name": "北京市东城区", "lng": 116.418, "lat": 39.929},
    {"name": "北京市西城区", "lng": 116.366, "lat": 39.916},
    {"name": "北京市朝阳区", "lng": 116.486, "lat": 39.921},
    {"name": "北京市海淀区", "lng": 116.298, "lat": 39.959},
    {"name": "北京市丰台区", "lng": 116.287, "lat": 39.858},
    {"name": "北京市石景山区", "lng": 116.222, "lat": 39.906},
    {"name": "北京市通州区", "lng": 116.656, "lat": 39.902},
    {"name": "北京市昌平区", "lng": 116.231, "lat": 40.221},
    {"name": "北京市大兴区", "lng": 116.339, "lat": 39.729},
    {"name": "北京市顺义区", "lng": 116.654, "lat": 40.130},
    # 上海
    {"name": "上海市黄浦区", "lng": 121.490, "lat": 31.232},
    {"name": "上海市徐汇区", "lng": 121.434, "lat": 31.188},
    {"name": "上海市长宁区", "lng": 121.423, "lat": 31.219},
    {"name": "上海市静安区", "lng": 121.446, "lat": 31.229},
    {"name": "上海市普陀区", "lng": 121.396, "lat": 31.249},
    {"name": "上海市虹口区", "lng": 121.481, "lat": 31.265},
    {"name": "上海市杨浦区", "lng": 121.527, "lat": 31.270},
    {"name": "上海市浦东新区", "lng": 121.545, "lat": 31.222},
    {"name": "上海市闵行区", "lng": 121.381, "lat": 31.113},
    {"name": "上海市宝山区", "lng": 121.490, "lat": 31.405},
    {"name": "上海市嘉定区", "lng": 121.266, "lat": 31.376},
    # 广州
    {"name": "广州市越秀区", "lng": 113.264, "lat": 23.129},
    {"name": "广州市荔湾区", "lng": 113.234, "lat": 23.125},
    {"name": "广州市海珠区", "lng": 113.264, "lat": 23.105},
    {"name": "广州市天河区", "lng": 113.361, "lat": 23.135},
    {"name": "广州市白云区", "lng": 113.273, "lat": 23.168},
    {"name": "广州市黄埔区", "lng": 113.459, "lat": 23.108},
    {"name": "广州市番禺区", "lng": 113.384, "lat": 22.937},
    {"name": "广州市花都区", "lng": 113.220, "lat": 23.392},
    {"name": "广州市南沙区", "lng": 113.520, "lat": 22.770},
    {"name": "广州市从化区", "lng": 113.587, "lat": 23.547},
    {"name": "广州市增城区", "lng": 113.812, "lat": 23.261},
    # 深圳
    {"name": "深圳市罗湖区", "lng": 114.131, "lat": 22.548},
    {"name": "深圳市福田区", "lng": 114.055, "lat": 22.521},
    {"name": "深圳市南山区", "lng": 113.930, "lat": 22.533},
    {"name": "深圳市宝安区", "lng": 113.883, "lat": 22.755},
    {"name": "深圳市龙岗区", "lng": 114.248, "lat": 22.721},
    {"name": "深圳市盐田区", "lng": 114.237, "lat": 22.558},
    {"name": "深圳市龙华区", "lng": 114.045, "lat": 22.687},
    {"name": "深圳市坪山区", "lng": 114.349, "lat": 22.691},
    {"name": "深圳市光明区", "lng": 113.936, "lat": 22.776},
    # 杭州
    {"name": "杭州市上城区", "lng": 120.170, "lat": 30.248},
    {"name": "杭州市拱墅区", "lng": 120.142, "lat": 30.318},
    {"name": "杭州市西湖区", "lng": 120.129, "lat": 30.259},
    {"name": "杭州市滨江区", "lng": 120.211, "lat": 30.208},
    {"name": "杭州市萧山区", "lng": 120.264, "lat": 30.184},
    {"name": "杭州市余杭区", "lng": 120.299, "lat": 30.400},
    {"name": "杭州市富阳区", "lng": 119.960, "lat": 30.049},
    {"name": "杭州市临安区", "lng": 119.726, "lat": 30.234},
    # 南京
    {"name": "南京市玄武区", "lng": 118.797, "lat": 32.059},
    {"name": "南京市秦淮区", "lng": 118.795, "lat": 32.028},
    {"name": "南京市建邺区", "lng": 118.732, "lat": 32.004},
    {"name": "南京市鼓楼区", "lng": 118.783, "lat": 32.060},
    {"name": "南京市浦口区", "lng": 118.625, "lat": 32.059},
    {"name": "南京市栖霞区", "lng": 118.809, "lat": 32.108},
    {"name": "南京市雨花台区", "lng": 118.779, "lat": 31.995},
    {"name": "南京市江宁区", "lng": 118.840, "lat": 31.954},
    # 成都
    {"name": "成都市锦江区", "lng": 104.081, "lat": 30.657},
    {"name": "成都市青羊区", "lng": 104.064, "lat": 30.675},
    {"name": "成都市金牛区", "lng": 104.052, "lat": 30.692},
    {"name": "成都市武侯区", "lng": 104.036, "lat": 30.642},
    {"name": "成都市成华区", "lng": 104.089, "lat": 30.660},
    {"name": "成都市龙泉驿区", "lng": 104.341, "lat": 30.569},
    {"name": "成都市青白江区", "lng": 104.251, "lat": 30.880},
    {"name": "成都市新都区", "lng": 104.160, "lat": 30.835},
    {"name": "成都市温江区", "lng": 103.856, "lat": 30.682},
    {"name": "成都市双流区", "lng": 103.923, "lat": 30.573},
    {"name": "成都市郫都区", "lng": 103.887, "lat": 30.781},
    # 武汉
    {"name": "武汉市江岸区", "lng": 114.309, "lat": 30.593},
    {"name": "武汉市江汉区", "lng": 114.271, "lat": 30.600},
    {"name": "武汉市硚口区", "lng": 114.215, "lat": 30.570},
    {"name": "武汉市汉阳区", "lng": 114.233, "lat": 30.536},
    {"name": "武汉市武昌区", "lng": 114.316, "lat": 30.587},
    {"name": "武汉市青山区", "lng": 114.384, "lat": 30.637},
    {"name": "武汉市洪山区", "lng": 114.344, "lat": 30.501},
    {"name": "武汉市东西湖区", "lng": 114.137, "lat": 30.620},
    {"name": "武汉市蔡甸区", "lng": 114.039, "lat": 30.582},
    {"name": "武汉市江夏区", "lng": 114.321, "lat": 30.351},
    {"name": "武汉市黄陂区", "lng": 114.371, "lat": 30.874},
    {"name": "武汉市新洲区", "lng": 114.802, "lat": 30.842},
    # 重庆
    {"name": "重庆市万州区", "lng": 108.408, "lat": 30.807},
    {"name": "重庆市渝中区", "lng": 106.568, "lat": 29.553},
    {"name": "重庆市江北区", "lng": 106.553, "lat": 29.576},
    {"name": "重庆市沙坪坝区", "lng": 106.456, "lat": 29.541},
    {"name": "重庆市九龙坡区", "lng": 106.511, "lat": 29.528},
    {"name": "重庆市南岸区", "lng": 106.566, "lat": 29.523},
    {"name": "重庆市北碚区", "lng": 106.396, "lat": 29.805},
    {"name": "重庆市渝北区", "lng": 106.517, "lat": 29.630},
    {"name": "重庆市巴南区", "lng": 106.519, "lat": 29.382},
    # 西安
    {"name": "西安市新城区", "lng": 108.947, "lat": 34.265},
    {"name": "西安市碑林区", "lng": 108.940, "lat": 34.259},
    {"name": "西安市莲湖区", "lng": 108.933, "lat": 34.266},
    {"name": "西安市灞桥区", "lng": 109.040, "lat": 34.273},
    {"name": "西安市未央区", "lng": 108.947, "lat": 34.299},
    {"name": "西安市雁塔区", "lng": 108.929, "lat": 34.214},
    {"name": "西安市阎良区", "lng": 109.226, "lat": 34.662},
    {"name": "西安市临潼区", "lng": 109.214, "lat": 34.367},
    {"name": "西安市长安区", "lng": 108.949, "lat": 34.158},
    # 天津
    {"name": "天津市和平区", "lng": 117.195, "lat": 39.118},
    {"name": "天津市河东区", "lng": 117.227, "lat": 39.124},
    {"name": "天津市河西区", "lng": 117.215, "lat": 39.083},
    {"name": "天津市南开区", "lng": 117.155, "lat": 39.137},
    {"name": "天津市河北区", "lng": 117.184, "lat": 39.152},
    {"name": "天津市红桥区", "lng": 117.152, "lat": 39.162},
    {"name": "天津市滨海新区", "lng": 117.710, "lat": 39.039},
    {"name": "天津市东丽区", "lng": 117.314, "lat": 39.087},
    {"name": "天津市西青区", "lng": 117.012, "lat": 39.136},
    {"name": "天津市津南区", "lng": 117.385, "lat": 38.988},
    # 苏州
    {"name": "苏州市姑苏区", "lng": 120.619, "lat": 31.311},
    {"name": "苏州市虎丘区", "lng": 120.566, "lat": 31.313},
    {"name": "苏州市吴中区", "lng": 120.630, "lat": 31.269},
    {"name": "苏州市相城区", "lng": 120.618, "lat": 31.369},
    {"name": "苏州市吴江区", "lng": 120.638, "lat": 31.160},
    # 郑州
    {"name": "郑州市中原区", "lng": 113.647, "lat": 34.754},
    {"name": "郑州市二七区", "lng": 113.640, "lat": 34.730},
    {"name": "郑州市管城区", "lng": 113.677, "lat": 34.754},
    {"name": "郑州市金水区", "lng": 113.661, "lat": 34.766},
    {"name": "郑州市上街区", "lng": 113.308, "lat": 34.803},
    {"name": "郑州市惠济区", "lng": 113.625, "lat": 34.866},
    # 长沙
    {"name": "长沙市芙蓉区", "lng": 112.988, "lat": 28.196},
    {"name": "长沙市天心区", "lng": 112.973, "lat": 28.183},
    {"name": "长沙市岳麓区", "lng": 112.939, "lat": 28.203},
    {"name": "长沙市开福区", "lng": 113.007, "lat": 28.201},
    {"name": "长沙市雨花区", "lng": 112.998, "lat": 28.135},
    {"name": "长沙市望城区", "lng": 112.831, "lat": 28.351},
    # 东莞
    {"name": "东莞市莞城街道", "lng": 113.751, "lat": 23.049},
    {"name": "东莞市南城街道", "lng": 113.752, "lat": 23.020},
    {"name": "东莞市东城街道", "lng": 113.781, "lat": 23.049},
    {"name": "东莞市万江街道", "lng": 113.737, "lat": 23.044},
    {"name": "东莞市石碣镇", "lng": 113.970, "lat": 23.099},
    {"name": "东莞市石龙镇", "lng": 114.076, "lat": 23.107},
    {"name": "东莞市茶山镇", "lng": 113.884, "lat": 23.076},
    {"name": "东莞市石排镇", "lng": 113.968, "lat": 23.089},
    {"name": "东莞市企石镇", "lng": 114.022, "lat": 23.070},
    {"name": "东莞市横沥镇", "lng": 113.966, "lat": 23.028},
    {"name": "东莞市桥头镇", "lng": 114.014, "lat": 23.019},
    {"name": "东莞市谢岗镇", "lng": 114.058, "lat": 23.003},
    {"name": "东莞市东坑镇", "lng": 113.949, "lat": 23.032},
    {"name": "东莞市常平镇", "lng": 113.993, "lat": 23.001},
    {"name": "东莞市寮步镇", "lng": 113.875, "lat": 22.992},
    {"name": "东莞市樟木头镇", "lng": 114.086, "lat": 22.917},
    {"name": "东莞市大朗镇", "lng": 113.946, "lat": 22.939},
    {"name": "东莞市黄江镇", "lng": 114.121, "lat": 22.909},
    {"name": "东莞市清溪镇", "lng": 114.165, "lat": 22.844},
    {"name": "东莞市塘厦镇", "lng": 114.071, "lat": 22.807},
    {"name": "东莞市凤岗镇", "lng": 114.146, "lat": 22.740},
    {"name": "东莞市大岭山镇", "lng": 113.949, "lat": 22.899},
    {"name": "东莞市长安镇", "lng": 113.803, "lat": 22.815},
    # 佛山
    {"name": "佛山市禅城区", "lng": 113.121, "lat": 23.022},
    {"name": "佛山市南海区", "lng": 113.143, "lat": 23.029},
    {"name": "佛山市顺德区", "lng": 113.156, "lat": 22.805},
    {"name": "佛山市三水区", "lng": 112.897, "lat": 23.156},
    {"name": "佛山市高明区", "lng": 112.893, "lat": 22.901},
    # 合肥
    {"name": "合肥市瑶海区", "lng": 117.309, "lat": 31.865},
    {"name": "合肥市庐阳区", "lng": 117.265, "lat": 31.867},
    {"name": "合肥市蜀山区", "lng": 117.227, "lat": 31.828},
    {"name": "合肥市包河区", "lng": 117.309, "lat": 31.794},
    # 昆明
    {"name": "昆明市五华区", "lng": 102.707, "lat": 25.043},
    {"name": "昆明市盘龙区", "lng": 102.722, "lat": 25.076},
    {"name": "昆明市官渡区", "lng": 102.749, "lat": 25.015},
    {"name": "昆明市西山区", "lng": 102.665, "lat": 25.040},
    # 沈阳
    {"name": "沈阳市和平区", "lng": 123.417, "lat": 41.796},
    {"name": "沈阳市沈河区", "lng": 123.453, "lat": 41.806},
    {"name": "沈阳市大东区", "lng": 123.469, "lat": 41.805},
    {"name": "沈阳市皇姑区", "lng": 123.406, "lat": 41.821},
    {"name": "沈阳市铁西区", "lng": 123.352, "lat": 41.803},
    # 青岛
    {"name": "青岛市市南区", "lng": 120.382, "lat": 36.067},
    {"name": "青岛市市北区", "lng": 120.374, "lat": 36.087},
    {"name": "青岛市黄岛区", "lng": 120.198, "lat": 35.960},
    {"name": "青岛市崂山区", "lng": 120.469, "lat": 36.107},
    {"name": "青岛市李沧区", "lng": 120.433, "lat": 36.145},
    {"name": "青岛市城阳区", "lng": 120.396, "lat": 36.307},
    # 宁波
    {"name": "宁波市海曙区", "lng": 121.544, "lat": 29.868},
    {"name": "宁波市江北区", "lng": 121.551, "lat": 29.887},
    {"name": "宁波市北仑区", "lng": 121.844, "lat": 29.909},
    {"name": "宁波市镇海区", "lng": 121.716, "lat": 29.952},
    {"name": "宁波市鄞州区", "lng": 121.547, "lat": 29.831},
    # 福州
    {"name": "福州市鼓楼区", "lng": 119.306, "lat": 26.074},
    {"name": "福州市台江区", "lng": 119.314, "lat": 26.053},
    {"name": "福州市仓山区", "lng": 119.335, "lat": 26.035},
    {"name": "福州市马尾区", "lng": 119.455, "lat": 25.989},
    {"name": "福州市晋安区", "lng": 119.329, "lat": 26.070},
    # 大连
    {"name": "大连市中山区", "lng": 121.648, "lat": 38.917},
    {"name": "大连市西岗区", "lng": 121.612, "lat": 38.914},
    {"name": "大连市沙河口区", "lng": 121.594, "lat": 38.912},
    {"name": "大连市甘井子区", "lng": 121.583, "lat": 38.948},
    # 长春
    {"name": "长春市南关区", "lng": 125.323, "lat": 43.893},
    {"name": "长春市宽城区", "lng": 125.323, "lat": 43.943},
    {"name": "长春市朝阳区", "lng": 125.323, "lat": 43.838},
    {"name": "长春市二道区", "lng": 125.378, "lat": 43.865},
    # 哈尔滨
    {"name": "哈尔滨市道里区", "lng": 126.617, "lat": 45.757},
    {"name": "哈尔滨市南岗区", "lng": 126.665, "lat": 45.760},
    {"name": "哈尔滨市道外区", "lng": 126.649, "lat": 45.772},
    {"name": "哈尔滨市平房区", "lng": 126.587, "lat": 45.597},
    # 济南
    {"name": "济南市历下区", "lng": 117.032, "lat": 36.669},
    {"name": "济南市市中区", "lng": 116.976, "lat": 36.651},
    {"name": "济南市槐荫区", "lng": 116.672, "lat": 36.651},
    {"name": "济南市天桥区", "lng": 117.000, "lat": 36.692},
    {"name": "济南市历城区", "lng": 117.066, "lat": 36.680},
    # 厦门
    {"name": "厦门市思明区", "lng": 118.089, "lat": 24.445},
    {"name": "厦门市湖里区", "lng": 118.101, "lat": 24.513},
    {"name": "厦门市集美区", "lng": 118.089, "lat": 24.576},
    {"name": "厦门市海沧区", "lng": 118.037, "lat": 24.485},
    {"name": "厦门市同安区", "lng": 118.152, "lat": 24.723},
    {"name": "厦门市翔安区", "lng": 118.347, "lat": 24.618},
    # 南宁
    {"name": "南宁市兴宁区", "lng": 108.353, "lat": 22.818},
    {"name": "南宁市青秀区", "lng": 108.373, "lat": 22.818},
    {"name": "南宁市江南区", "lng": 108.160, "lat": 22.782},
    {"name": "南宁市西乡塘区", "lng": 108.307, "lat": 22.833},
    # 贵阳
    {"name": "贵阳市南明区", "lng": 106.715, "lat": 26.566},
    {"name": "贵阳市云岩区", "lng": 106.713, "lat": 26.588},
    {"name": "贵阳市花溪区", "lng": 106.690, "lat": 26.433},
    {"name": "贵阳市乌当区", "lng": 106.750, "lat": 26.630},
    # 昆明
    {"name": "昆明市五华区", "lng": 102.707, "lat": 25.043},
    {"name": "昆明市盘龙区", "lng": 102.722, "lat": 25.076},
    {"name": "昆明市官渡区", "lng": 102.749, "lat": 25.015},
    {"name": "昆明市西山区", "lng": 102.665, "lat": 25.040},
    # 兰州
    {"name": "兰州市城关区", "lng": 103.834, "lat": 36.061},
    {"name": "兰州市七里河区", "lng": 103.784, "lat": 36.066},
    {"name": "兰州市西固区", "lng": 103.626, "lat": 36.119},
    {"name": "兰州市安宁区", "lng": 103.719, "lat": 36.104},
    # 乌鲁木齐
    {"name": "乌鲁木齐市天山区", "lng": 87.617, "lat": 43.793},
    {"name": "乌鲁木齐市沙依巴克区", "lng": 87.598, "lat": 43.801},
    {"name": "乌鲁木齐市新市区", "lng": 87.575, "lat": 43.858},
    # 石家庄
    {"name": "石家庄市长安区", "lng": 114.540, "lat": 38.040},
    {"name": "石家庄市桥西区", "lng": 114.461, "lat": 38.004},
    {"name": "石家庄市新华区", "lng": 114.461, "lat": 38.051},
    {"name": "石家庄市裕华区", "lng": 114.519, "lat": 38.034},
    # 太原
    {"name": "太原市小店区", "lng": 112.566, "lat": 37.873},
    {"name": "太原市迎泽区", "lng": 112.558, "lat": 37.863},
    {"name": "太原市杏花岭区", "lng": 112.564, "lat": 37.881},
    {"name": "太原市尖草坪区", "lng": 112.527, "lat": 38.042},
    # 呼和浩特
    {"name": "呼和浩特市新城区", "lng": 111.665, "lat": 40.825},
    {"name": "呼和浩特市回民区", "lng": 111.623, "lat": 40.813},
    {"name": "呼和浩特市玉泉区", "lng": 111.674, "lat": 40.753},
    {"name": "呼和浩特市赛罕区", "lng": 111.696, "lat": 40.808},
    # 南昌
    {"name": "南昌市东湖区", "lng": 115.894, "lat": 28.682},
    {"name": "南昌市西湖区", "lng": 115.877, "lat": 28.657},
    {"name": "南昌市青云谱区", "lng": 115.914, "lat": 28.629},
    {"name": "南昌市青山湖区", "lng": 115.914, "lat": 28.687},
    # 乌鲁木齐
    {"name": "乌鲁木齐市天山区", "lng": 87.617, "lat": 43.793},
    {"name": "乌鲁木齐市沙依巴克区", "lng": 87.598, "lat": 43.801},
    # 拉萨
    {"name": "拉萨市城关区", "lng": 91.132, "lat": 29.652},
    # 银川
    {"name": "银川市兴庆区", "lng": 106.288, "lat": 38.487},
    {"name": "银川市西夏区", "lng": 106.251, "lat": 38.493},
    {"name": "银川市金凤区", "lng": 106.243, "lat": 38.478},
    # 西宁
    {"name": "西宁市城东区", "lng": 101.777, "lat": 36.618},
    {"name": "西宁市城中区", "lng": 101.777, "lat": 36.601},
    {"name": "西宁市城西区", "lng": 101.777, "lat": 36.595},
    {"name": "西宁市城北区", "lng": 101.785, "lat": 36.656},
    # 三亚
    {"name": "三亚市海棠区", "lng": 109.736, "lat": 18.308},
    {"name": "三亚市吉阳区", "lng": 109.517, "lat": 18.253},
    {"name": "三亚市天涯区", "lng": 109.465, "lat": 18.299},
    {"name": "三亚市崖州区", "lng": 109.168, "lat": 18.352},
]


# ============================================================
# 地名解析器
# ============================================================
class GeoCoder:
    """
    地名解析：将 GPS 坐标转换为地名
    策略（按优先级）：
    1. Android 系统 Geocoder（无需 key，小米/华为等国产机通常接入百度/高德后端）
    2. 用户配置的在线地图 API（高德/百度）
    3. 离线数据库最近邻匹配
    """

    def __init__(self, amap_key=None, baidu_key=None, provider="auto"):
        self.amap_key = amap_key or AMAP_API_KEY
        self.baidu_key = baidu_key or BAIDU_API_KEY
        self.provider = provider  # "auto", "amap", "baidu", "android", "offline"
        self.offline_locations = self._load_offline_db()
        self._cache = {}

    def _load_offline_db(self):
        """加载离线位置数据库"""
        if os.path.exists(OFFLINE_DB_PATH):
            try:
                with open(OFFLINE_DB_PATH, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return DEFAULT_LOCATIONS

    def reverse_geocode(self, lng, lat):
        """
        逆地理编码：经纬度 -> 地名
        返回: "省市区详细地址" 或 None
        """
        try:
            lng = float(lng)
            lat = float(lat)
        except Exception:
            return None

        cache_key = f"{lng:.5f},{lat:.5f}"
        if cache_key in self._cache:
            return self._cache[cache_key]

        result = None

        # 1. Android 系统 Geocoder（无需 key，优先使用；小米/华为通常接入百度/高德后端）
        if not result and self.provider in ("auto", "android"):
            result = self._android_geocoder(lng, lat)

        # 2. 在线 API（用户配置了 key 才使用）
        if not result and self.provider in ("auto", "amap") and self.amap_key and self.amap_key != "YOUR_AMAP_KEY_HERE":
            result = self._api_reverse_amap(lng, lat)
        if not result and self.provider in ("auto", "baidu") and self.baidu_key and self.baidu_key != "YOUR_BAIDU_KEY_HERE":
            result = self._api_reverse_baidu(lng, lat)

        # 3. 离线最近邻
        if not result:
            result = self._offline_nearest(lng, lat)

        if result:
            self._cache[cache_key] = result
        return result

    def _api_reverse_amap(self, lng, lat):
        """调用高德逆地理编码 API（带一次重试）"""
        for attempt in range(2):
            try:
                params = {
                    "key": self.amap_key,
                    "location": f"{lng},{lat}",
                    "extensions": "base",
                    "output": "JSON",
                }
                resp = requests.get(AMAP_GEOCODE_URL, params=params, timeout=5)
                data = resp.json()
                if data.get("status") == "1" and data.get("regeocode"):
                    formatted = data["regeocode"].get("formatted_address", "")
                    if formatted:
                        return formatted
                    addr = data["regeocode"].get("addressComponent", {})
                    province = addr.get("province", "")
                    city = addr.get("city", "")
                    district = addr.get("district", "")
                    if not city and province:
                        city = province
                    return f"{province}{city}{district}"
                break
            except Exception as e:
                Logger.warning(f"GeoCoder AMap API 尝试{attempt + 1}失败: {e}")
                if attempt == 0:
                    import time
                    time.sleep(0.5)
        return None

    def _api_reverse_baidu(self, lng, lat):
        """调用百度逆地理编码 API（带一次重试 + WGS-84 转 BD-09）
        Android GPS 返回的通常是 WGS-84。百度对 WGS-84 支持区县级别解析；
        若用户配置了 AK，先把 WGS-84 转成 BD-09 再请求，门牌级精度更高。
        """
        for attempt in range(2):
            try:
                # WGS-84 -> GCJ-02 -> BD-09，提升百度解析精度
                gcj_lng, gcj_lat = self._wgs84_to_gcj02(lng, lat)
                bd_lng, bd_lat = self._gcj02_to_bd09(gcj_lng, gcj_lat)
                params = {
                    "ak": self.baidu_key,
                    "location": f"{bd_lat},{bd_lng}",
                    "output": "json",
                    "coordtype": "bd09ll",
                }
                resp = requests.get(BAIDU_GEOCODE_URL, params=params, timeout=5)
                data = resp.json()
                if data.get("status") == 0 and data.get("result"):
                    result = data["result"]
                    formatted = result.get("formatted_address", "")
                    if formatted:
                        return formatted
                    comp = result.get("addressComponent", {})
                    province = comp.get("province", "")
                    city = comp.get("city", "")
                    district = comp.get("district", "")
                    street = comp.get("street", "")
                    return f"{province}{city}{district}{street}"
                break
            except Exception as e:
                Logger.warning(f"GeoCoder Baidu API 尝试{attempt + 1}失败: {e}")
                if attempt == 0:
                    import time
                    time.sleep(0.5)
        return None

    @staticmethod
    def _wgs84_to_gcj02(lng, lat):
        """WGS-84 转 GCJ-02（火星坐标）"""
        if GeoCoder._out_of_china(lng, lat):
            return lng, lat
        dlat = GeoCoder._transform_lat(lng - 105.0, lat - 35.0)
        dlng = GeoCoder._transform_lng(lng - 105.0, lat - 35.0)
        radlat = lat / 180.0 * math.pi
        magic = math.sin(radlat)
        magic = 1 - 0.00669342162296594323 * magic * magic
        sqrtmagic = math.sqrt(magic)
        dlat = (dlat * 180.0) / ((6378245.0 * (1 - 0.00669342162296594323)) / (magic * sqrtmagic) * math.pi)
        dlng = (dlng * 180.0) / (6378245.0 / sqrtmagic * math.cos(radlat) * math.pi)
        return lng + dlng, lat + dlat

    @staticmethod
    def _gcj02_to_bd09(lng, lat):
        """GCJ-02 转 BD-09（百度坐标）"""
        x = lng
        y = lat
        z = math.sqrt(x * x + y * y) + 0.00002 * math.sin(y * math.pi * 3000.0 / 180.0)
        theta = math.atan2(y, x) + 0.000003 * math.cos(x * math.pi * 3000.0 / 180.0)
        bd_lng = z * math.cos(theta) + 0.0065
        bd_lat = z * math.sin(theta) + 0.006
        return bd_lng, bd_lat

    @staticmethod
    def _transform_lat(lng, lat):
        ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * math.sqrt(abs(lng))
        ret += (20.0 * math.sin(6.0 * lng * math.pi) + 20.0 * math.sin(2.0 * lng * math.pi)) * 2.0 / 3.0
        ret += (20.0 * math.sin(lat * math.pi) + 40.0 * math.sin(lat / 3.0 * math.pi)) * 2.0 / 3.0
        ret += (160.0 * math.sin(lat / 12.0 * math.pi) + 320 * math.sin(lat * math.pi / 30.0)) * 2.0 / 3.0
        return ret

    @staticmethod
    def _transform_lng(lng, lat):
        ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * math.sqrt(abs(lng))
        ret += (20.0 * math.sin(6.0 * lng * math.pi) + 20.0 * math.sin(2.0 * lng * math.pi)) * 2.0 / 3.0
        ret += (20.0 * math.sin(lng * math.pi) + 40.0 * math.sin(lng / 3.0 * math.pi)) * 2.0 / 3.0
        ret += (150.0 * math.sin(lng / 12.0 * math.pi) + 300.0 * math.sin(lng / 30.0 * math.pi)) * 2.0 / 3.0
        return ret

    @staticmethod
    def _out_of_china(lng, lat):
        return lng < 72.004 or lng > 137.8347 or lat < 0.8293 or lat > 55.8271

    def _android_geocoder(self, lng, lat):
        """调用 Android 系统 Geocoder（无需 API key）
        小米/华为等国产 ROM 通常会把 Geocoder 后端接到百度或高德，
        因此无需额外 key 也能拿到中文地址。
        """
        try:
            from jnius import autoclass
            PythonActivity = autoclass('org.kivy.android.PythonActivity')
            activity = PythonActivity.mActivity
            if activity is None:
                return None
            Geocoder = autoclass('android.location.Geocoder')
            if not Geocoder.isPresent():
                Logger.warning("GeoCoder: 系统 Geocoder 不可用")
                return None

            Locale = autoclass('java.util.Locale')
            gc = Geocoder(activity, Locale.CHINA)
            # 部分 ROM 第一次调用可能超时，重试 2 次
            for attempt in range(2):
                try:
                    addresses = gc.getFromLocation(float(lat), float(lng), 1)
                    if addresses and addresses.size() > 0:
                        addr = addresses.get(0)
                        # 优先取完整地址行
                        try:
                            full = addr.getAddressLine(0)
                            if full:
                                return str(full)
                        except Exception:
                            pass
                        # 回退：拼接省市区街道门牌
                        parts = []
                        for method in ['getAdminArea', 'getLocality', 'getSubLocality',
                                       'getThoroughfare', 'getSubThoroughfare']:
                            try:
                                part = getattr(addr, method)()
                                if part:
                                    parts.append(str(part))
                            except Exception:
                                pass
                        if parts:
                            return "".join(parts)
                except Exception as e:
                    Logger.warning(f"GeoCoder Android Geocoder 尝试{attempt + 1}失败: {e}")
                    if attempt == 0:
                        import time
                        time.sleep(0.3)
        except Exception as e:
            Logger.error(f"GeoCoder Android Geocoder: {e}")
        return None

    def _offline_nearest(self, lng, lat):
        """离线最近邻匹配"""
        if not self.offline_locations:
            return None

        min_dist = float('inf')
        nearest = None

        for loc in self.offline_locations:
            d = self._haversine(lng, lat, loc["lng"], loc["lat"])
            if d < min_dist:
                min_dist = d
                nearest = loc

        if nearest:
            if min_dist > 100:
                return nearest["name"][:2] + "附近"
            return nearest["name"]
        return "未知位置"

    @staticmethod
    def _haversine(lng1, lat1, lng2, lat2):
        """计算两点间距离（km）"""
        R = 6371
        dlat = math.radians(lat2 - lat1)
        dlng = math.radians(lng2 - lng1)
        a = (math.sin(dlat / 2) ** 2 +
             math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) *
             math.sin(dlng / 2) ** 2)
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        return R * c

    def get_location_name(self, lng, lat):
        """获取位置名称的便捷方法"""
        if lng == 0 and lat == 0:
            return "定位不可用"
        return self.reverse_geocode(lng, lat) or "未知位置"


# ============================================================
# 测试
# ============================================================
if __name__ == "__main__":
    gc = GeoCoder()

    print("=== 离线定位测试 ===")
    tests = [
        (116.407, 39.904, "北京天安门附近"),
        (121.473, 31.230, "上海外滩附近"),
        (113.264, 23.129, "广州"),
        (114.058, 22.543, "深圳"),
        (120.155, 30.274, "杭州"),
        (0, 0, "无效坐标"),
    ]
    for lng, lat, desc in tests:
        name = gc.get_location_name(lng, lat)
        print(f"  {desc}: {name}")

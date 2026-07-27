"""
曲线计算模块

实现高速公路平曲线、竖曲线的核心计算功能
包括：线元法、缓和曲线、圆曲线、直线段等计算
"""

from typing import Tuple, List, Optional
from dataclasses import dataclass
import math


@dataclass
class CurveParams:
    """曲线参数数据类"""
    curve_type: str       # 曲线类型：直线/圆曲线/缓和曲线
    station: float        # K - 里程
    x_coord: float        # C - X 坐标
    y_coord: float        # D - Y 坐标
    azimuth: float        # F - 方位角 (度。分秒格式)
    radius: float         # R - 半径 (左偏负，右偏正)
    spiral_a: float       # A - 回旋参数
    spiral_l1: float      # L1 - 第一缓和曲线长
    spiral_l2: float      # L2 - 第二缓和曲线长


class AngleUtils:
    """角度工具类"""
    
    @staticmethod
    def dms_to_decimal(dms: float) -> float:
        """
        将度。分秒格式转换为十进制度
        
        Args:
            dms: 度。分秒格式的角度值 (如 153.240524 表示 153°24'05.24")
            
        Returns:
            十进制度数
        """
        if dms == 0:
            return 0.0
        
        sign = 1 if dms >= 0 else -1
        dms = abs(dms)
        
        degrees = int(dms)
        minutes = int((dms - degrees) * 100)
        seconds = ((dms - degrees) * 100 - minutes) * 100
        
        decimal_degrees = degrees + minutes / 60.0 + seconds / 3600.0
        return sign * decimal_degrees
    
    @staticmethod
    def decimal_to_dms(decimal: float) -> float:
        """
        将十进制度转换为度。分秒格式
        
        Args:
            decimal: 十进制度数
            
        Returns:
            度。分秒格式的角度值
        """
        if decimal == 0:
            return 0.0
        
        sign = 1 if decimal >= 0 else -1
        decimal = abs(decimal)
        
        degrees = int(decimal)
        minutes_decimal = (decimal - degrees) * 60
        minutes = int(minutes_decimal)
        seconds = (minutes_decimal - minutes) * 60
        
        dms = degrees + minutes / 100.0 + seconds / 10000.0
        return sign * dms
    
    @staticmethod
    def normalize_azimuth(azimuth: float) -> float:
        """
        规范化方位角到 0-360 度范围
        
        Args:
            azimuth: 方位角 (弧度)
            
        Returns:
            规范化的方位角
        """
        two_pi = 2 * math.pi
        azimuth = azimuth % two_pi
        if azimuth < 0:
            azimuth += two_pi
        return azimuth


class CoordinateCalculator:
    """坐标计算器"""
    
    def __init__(self):
        self.PI = math.pi
        self.angle_utils = AngleUtils()
    
    def calculate_azimuth(self, x1: float, y1: float, x2: float, y2: float) -> float:
        """
        计算两点间的方位角
        
        Args:
            x1, y1: 起点坐标
            x2, y2: 终点坐标
            
        Returns:
            方位角 (弧度)
        """
        dx = x2 - x1
        dy = y2 - y1
        
        if dx != 0 and dy != 0:
            azimuth = math.atan(dy / dx)
            if dx < 0:
                azimuth += self.PI
            elif dy < 0:
                azimuth += 2 * self.PI
        elif dx == 0:
            azimuth = self.PI / 2 if dy > 0 else 3 * self.PI / 2
        else:  # dy == 0
            azimuth = 0 if dx > 0 else self.PI
        
        return AngleUtils.normalize_azimuth(azimuth)
    
    def calculate_distance(self, x1: float, y1: float, x2: float, y2: float) -> float:
        """
        计算两点间的距离
        
        Args:
            x1, y1: 起点坐标
            x2, y2: 终点坐标
            
        Returns:
            距离
        """
        return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
    
    def polar_to_rectangular(self, x0: float, y0: float, distance: float, 
                            azimuth: float) -> Tuple[float, float]:
        """
        极坐标转直角坐标
        
        Args:
            x0, y0: 起点坐标
            distance: 距离
            azimuth: 方位角 (弧度)
            
        Returns:
            (x, y) 终点坐标
        """
        x = x0 + distance * math.cos(azimuth)
        y = y0 + distance * math.sin(azimuth)
        return (x, y)


class HorizontalCurve(CoordinateCalculator):
    """平曲线计算类"""
    
    def __init__(self):
        super().__init__()
        self.curve_params: List[CurveParams] = []
    
    def load_curve_params(self, params: List[CurveParams]):
        """加载曲线参数"""
        self.curve_params = params
    
    def find_curve_segment(self, station: float) -> int:
        """
        根据里程查找对应的曲线段索引
        
        Args:
            station: 里程值
            
        Returns:
            曲线段索引
        """
        if not self.curve_params:
            return -1
        
        idx = 0
        if self.curve_params[0].station < self.curve_params[-1].station:
            # 里程递增
            while (idx < len(self.curve_params) - 1 and 
                   station > self.curve_params[idx].station):
                idx += 1
        else:
            # 里程递减
            while (idx < len(self.curve_params) - 1 and 
                   station < self.curve_params[idx].station):
                idx += 1
        
        return max(0, idx)
    
    def calculate_line(self, station: float, segment_idx: int) -> Tuple[float, float, float]:
        """
        直线段坐标计算
        
        Args:
            station: 待求点里程
            segment_idx: 曲线段索引
            
        Returns:
            (x, y, azimuth) 坐标和方位角
        """
        param = self.curve_params[segment_idx]
        
        delta_station = station - param.station
        azimuth_rad = self.angle_utils.dms_to_decimal(param.azimuth) * self.PI / 180.0
        
        x = param.x_coord + delta_station * math.cos(azimuth_rad)
        y = param.y_coord + delta_station * math.sin(azimuth_rad)
        
        return (x, y, azimuth_rad)
    
    def calculate_circle(self, station: float, segment_idx: int) -> Tuple[float, float, float]:
        """
        圆曲线段坐标计算
        
        Args:
            station: 待求点里程
            segment_idx: 曲线段索引
            
        Returns:
            (x, y, azimuth) 坐标和方位角
        """
        param = self.curve_params[segment_idx]
        
        arc_length = station - param.station
        radius = param.radius
        
        # 圆心角 (弧度)
        central_angle = arc_length / abs(radius)
        
        # 判断左右偏
        direction = 1 if radius > 0 else -1
        
        # 起算方位角
        start_azimuth = self.angle_utils.dms_to_decimal(param.azimuth) * self.PI / 180.0
        
        # 切线方位角变化
        tangent_azimuth = start_azimuth + direction * central_angle
        
        # 弦切角
        chord_angle = central_angle / 2.0
        
        # 弦长
        chord_length = 2 * abs(radius) * math.sin(central_angle / 2.0)
        
        # 弦方位角
        chord_azimuth = start_azimuth + direction * chord_angle
        
        # 计算坐标
        x = param.x_coord + chord_length * math.cos(chord_azimuth)
        y = param.y_coord + chord_length * math.sin(chord_azimuth)
        
        return (x, y, tangent_azimuth)
    
    def calculate_spiral(self, station: float, segment_idx: int) -> Tuple[float, float, float]:
        """
        缓和曲线段坐标计算 (使用近似公式)
        
        Args:
            station: 待求点里程
            segment_idx: 曲线段索引
            
        Returns:
            (x, y, azimuth) 坐标和方位角
        """
        param = self.curve_params[segment_idx]
        
        l = station - param.station  # 缓和曲线上点到起点距离
        a = param.spiral_a
        r = param.radius
        
        if a == 0 or r == 0:
            # 退化情况，按直线处理
            return self.calculate_line(station, segment_idx)
        
        # 缓和曲线参数
        ls = abs(a * a / r)  # 缓和曲线全长
        
        # 切线角 (弧度)
        beta = l * l / (2 * a * a)
        
        # 起算方位角
        start_azimuth = self.angle_utils.dms_to_decimal(param.azimuth) * self.PI / 180.0
        
        # 方向判断
        direction = 1 if r > 0 else -1
        
        # 当前点方位角
        current_azimuth = start_azimuth + direction * beta
        
        # 使用级数展开计算坐标增量 (近似公式)
        # x = l - l^5/(40*A^4) + ...
        # y = l^3/(6*A^2) - l^7/(336*A^6) + ...
        x_local = l - (l ** 5) / (40 * a ** 4)
        y_local = (l ** 3) / (6 * a ** 2) - (l ** 7) / (336 * a ** 6)
        
        # 坐标转换
        cos_f = math.cos(start_azimuth)
        sin_f = math.sin(start_azimuth)
        
        if direction < 0:
            y_local = -y_local
        
        x = param.x_coord + x_local * cos_f - y_local * sin_f
        y = param.y_coord + x_local * sin_f + y_local * cos_f
        
        return (x, y, current_azimuth)
    
    def calculate_point(self, station: float) -> Optional[Tuple[float, float, float]]:
        """
        根据里程计算坐标和方位角
        
        Args:
            station: 待求点里程
            
        Returns:
            (x, y, azimuth) 或 None
        """
        if not self.curve_params:
            return None
        
        segment_idx = self.find_curve_segment(station)
        if segment_idx < 0:
            return None
        
        param = self.curve_params[segment_idx]
        
        if param.curve_type == "直线":
            return self.calculate_line(station, segment_idx)
        elif param.curve_type == "圆曲线":
            return self.calculate_circle(station, segment_idx)
        elif param.curve_type in ["缓和曲线", "1+ 圆 +2"]:
            return self.calculate_spiral(station, segment_idx)
        
        return None
    
    def inverse_calculation(self, target_x: float, target_y: float, 
                           initial_station: float) -> Tuple[float, float]:
        """
        坐标反算：已知坐标求里程和偏距
        
        使用逐步趋近法
        
        Args:
            target_x, target_y: 目标点坐标
            initial_station: 初始估算里程
            
        Returns:
            (station, offset) 里程和偏距
        """
        station = initial_station
        
        for _ in range(100):  # 最大迭代次数
            result = self.calculate_point(station)
            if result is None:
                break
                
            calc_x, calc_y, azimuth = result
            
            # 计算距离和方位角
            dx = target_x - calc_x
            dy = target_y - calc_y
            distance = math.sqrt(dx * dx + dy * dy)
            
            if distance < 0.001:  # 精度满足要求
                break
            
            line_azimuth = self.calculate_azimuth(calc_x, calc_y, target_x, target_y)
            
            # 计算沿路线方向的投影
            angle_diff = line_azimuth - azimuth
            station_correction = distance * math.cos(angle_diff)
            
            station += station_correction * 0.382  # 黄金分割系数
        
        # 计算最终偏距
        final_result = self.calculate_point(station)
        if final_result:
            calc_x, calc_y, azimuth = final_result
            dx = target_x - calc_x
            dy = target_y - calc_y
            line_azimuth = self.calculate_azimuth(calc_x, calc_y, target_x, target_y)
            angle_diff = line_azimuth - azimuth
            offset = distance * math.sin(angle_diff)
            if abs(angle_diff) > 0:
                offset = distance * math.sin(angle_diff) / abs(math.sin(angle_diff))
            else:
                offset = 0
        else:
            offset = 0
        
        return (station, offset)


class VerticalCurve:
    """竖曲线计算类"""
    
    def __init__(self):
        self.PI = math.pi
    
    def calculate_elevation(self, station: float, grade_points: List[Tuple[float, float]],
                           radius: float) -> float:
        """
        竖曲线高程计算
        
        Args:
            station: 待求点里程
            grade_points: 变坡点列表 [(里程，高程), ...]
            radius: 竖曲线半径
            
        Returns:
            设计高程
        """
        if not grade_points:
            return 0.0
        
        # 找到相邻的变坡点
        i = 0
        for j in range(len(grade_points) - 1):
            if grade_points[j][0] <= station <= grade_points[j + 1][0]:
                i = j
                break
        
        if i >= len(grade_points) - 1:
            # 在最后一个区间外
            i = len(grade_points) - 2
        
        # 计算前后坡度
        if i > 0:
            grade1 = (grade_points[i][1] - grade_points[i - 1][1]) / \
                     (grade_points[i][0] - grade_points[i - 1][0])
        else:
            grade1 = 0
        
        if i < len(grade_points) - 1:
            grade2 = (grade_points[i + 1][1] - grade_points[i][1]) / \
                     (grade_points[i + 1][0] - grade_points[i][0])
        else:
            grade2 = 0
        
        # 坡度差
        omega = grade2 - grade1
        
        # 切线长
        t = abs(radius * omega / 2)
        
        # 变坡点信息
        pvi_station = grade_points[i][0]
        pvi_elevation = grade_points[i][1]
        
        # 判断是凸曲线还是凹曲线
        is_convex = omega < 0
        
        # 计算切线高程
        if station <= pvi_station:
            tangent_elevation = pvi_elevation + grade1 * (station - pvi_station)
        else:
            tangent_elevation = pvi_elevation + grade2 * (station - pvi_station)
        
        # 计算竖曲线改正值
        x = abs(station - pvi_station)
        if x <= t:
            correction = x * x / (2 * radius)
            if is_convex:
                correction = -correction
            elevation = tangent_elevation + correction
        else:
            elevation = tangent_elevation
        
        return elevation


# 导出公共接口
__all__ = [
    'CurveParams',
    'AngleUtils',
    'CoordinateCalculator',
    'HorizontalCurve',
    'VerticalCurve'
]

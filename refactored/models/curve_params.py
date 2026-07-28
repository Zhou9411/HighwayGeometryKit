"""
曲线参数模型

定义曲线参数的数据结构和验证逻辑
"""

from dataclasses import dataclass, field
from typing import List, Optional
from enum import Enum


class CurveType(Enum):
    """曲线类型枚举"""
    LINE = "直线"
    CIRCLE = "圆曲线"
    SPIRAL = "缓和曲线"
    COMPOUND = "1+ 圆 +2"  # 复合曲线


@dataclass
class CurveSegment:
    """
    曲线段参数
    
    Attributes:
        curve_type: 曲线类型
        station: 起算点里程 (K)
        x_coord: 起算点 X 坐标 (C)
        y_coord: 起算点 Y 坐标 (D)
        azimuth: 起算点切线方位角 (F, 度。分秒格式)
        radius: 半径 (R, 左偏负右偏正)
        spiral_a: 回旋参数 (A)
        offset_x: X 方向偏移
        offset_y: Y 方向偏移
    """
    curve_type: CurveType
    station: float
    x_coord: float
    y_coord: float
    azimuth: float  # 度。分秒格式
    radius: float = 0.0
    spiral_a: float = 0.0
    offset_x: float = 0.0
    offset_y: float = 0.0
    
    def __post_init__(self):
        """数据验证和初始化"""
        if isinstance(self.curve_type, str):
            self.curve_type = CurveType(self.curve_type)
        
        # 验证必填字段
        if not self.station:
            raise ValueError("里程不能为空")
        if not self.x_coord:
            raise ValueError("X 坐标不能为空")
        if not self.y_coord:
            raise ValueError("Y 坐标不能为空")
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'curve_type': self.curve_type.value,
            'station': self.station,
            'x_coord': self.x_coord,
            'y_coord': self.y_coord,
            'azimuth': self.azimuth,
            'radius': self.radius,
            'spiral_a': self.spiral_a,
            'offset_x': self.offset_x,
            'offset_y': self.offset_y
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'CurveSegment':
        """从字典创建实例"""
        return cls(
            curve_type=CurveType(data.get('curve_type', '直线')),
            station=float(data.get('station', 0)),
            x_coord=float(data.get('x_coord', 0)),
            y_coord=float(data.get('y_coord', 0)),
            azimuth=float(data.get('azimuth', 0)),
            radius=float(data.get('radius', 0)),
            spiral_a=float(data.get('spiral_a', 0)),
            offset_x=float(data.get('offset_x', 0)),
            offset_y=float(data.get('offset_y', 0))
        )


@dataclass
class VerticalGradePoint:
    """
    竖曲线变坡点
    
    Attributes:
        station: 里程
        elevation: 高程
    """
    station: float
    elevation: float
    
    def __post_init__(self):
        """数据验证"""
        if not self.station:
            raise ValueError("里程不能为空")
        if self.elevation is None:
            raise ValueError("高程不能为空")


@dataclass
class VerticalCurveParams:
    """
    竖曲线参数
    
    Attributes:
        radius: 竖曲线半径
        grade_points: 变坡点列表
    """
    radius: float
    grade_points: List[VerticalGradePoint] = field(default_factory=list)
    
    def add_grade_point(self, station: float, elevation: float):
        """添加变坡点"""
        self.grade_points.append(VerticalGradePoint(station, elevation))
        # 按里程排序
        self.grade_points.sort(key=lambda p: p.station)
    
    def get_elevation_range(self) -> tuple:
        """获取高程范围"""
        if not self.grade_points:
            return (0.0, 0.0)
        elevations = [p.elevation for p in self.grade_points]
        return (min(elevations), max(elevations))


@dataclass
class WidthTransitionParams:
    """
    宽度渐变参数
    
    Attributes:
        start_station: 起点里程
        end_station: 终点里程
        start_width: 起点宽度
        end_width: 终点宽度
        left_offset: 左侧偏移
        right_offset: 右侧偏移
    """
    start_station: float
    end_station: float
    start_width: float
    end_width: float
    left_offset: float = 0.0
    right_offset: float = 0.0
    
    def calculate_width(self, station: float) -> float:
        """
        计算指定里程处的宽度
        
        Args:
            station: 里程
            
        Returns:
            宽度值
        """
        if station <= self.start_station:
            return self.start_width
        elif station >= self.end_station:
            return self.end_width
        else:
            ratio = (station - self.start_station) / (self.end_station - self.start_station)
            return self.start_width + ratio * (self.end_width - self.start_width)


@dataclass
class CrossSectionParams:
    """
    横断面参数
    
    Attributes:
        station: 里程
        design_line: 设计线数据 [(偏距，高程), ...]
        ground_line: 地面线数据 [(偏距，高程), ...]
        text_height: 文字高度
        font_name: 字体名称
    """
    station: float
    design_line: List[tuple] = field(default_factory=list)
    ground_line: List[tuple] = field(default_factory=list)
    text_height: float = 0.4
    font_name: str = "宋体"
    
    def add_design_point(self, offset: float, elevation: float):
        """添加设计线点"""
        self.design_line.append((offset, elevation))
        self.design_line.sort(key=lambda p: p[0])
    
    def add_ground_point(self, offset: float, elevation: float):
        """添加地面线点"""
        self.ground_line.append((offset, elevation))
        self.ground_line.sort(key=lambda p: p[0])


# 导出公共接口
__all__ = [
    'CurveType',
    'CurveSegment',
    'VerticalGradePoint',
    'VerticalCurveParams',
    'WidthTransitionParams',
    'CrossSectionParams'
]

#!/usr/bin/env python3
"""
高速公路坐标高程计算软件 - 重构版

主程序入口
演示核心计算功能的使用
"""

import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.curve import HorizontalCurve, VerticalCurve, AngleUtils
from models.curve_params import CurveSegment, CurveType, VerticalGradePoint


def demo_horizontal_curve():
    """演示平曲线计算"""
    print("=" * 60)
    print("平曲线计算演示")
    print("=" * 60)
    
    # 创建曲线参数 - 使用实际的有效数据
    segments = [
        CurveSegment(
            curve_type=CurveType.SPIRAL,
            station=1000.0,
            x_coord=50000.0,
            y_coord=50000.0,
            azimuth=153.240524,  # 度。分秒格式
            radius=-800.0,       # 左偏
            spiral_a=70.0
        ),
        CurveSegment(
            curve_type=CurveType.CIRCLE,
            station=1200.0,
            x_coord=49850.0,  # 提供初始估算值
            y_coord=50120.0,
            azimuth=0.0,
            radius=-800.0
        ),
        CurveSegment(
            curve_type=CurveType.LINE,
            station=1500.0,
            x_coord=49700.0,
            y_coord=50200.0,
            azimuth=0.0
        )
    ]
    
    # 创建计算器
    calculator = HorizontalCurve()
    
    # 加载参数 (需要转换为内部格式)
    from dataclasses import dataclass
    
    @dataclass
    class InternalParam:
        curve_type: str
        station: float
        x_coord: float
        y_coord: float
        azimuth: float
        radius: float
        spiral_a: float
    
    internal_params = [
        InternalParam(
            curve_type=s.curve_type.value,
            station=s.station,
            x_coord=s.x_coord,
            y_coord=s.y_coord,
            azimuth=s.azimuth,
            radius=s.radius,
            spiral_a=s.spiral_a
        ) for s in segments
    ]
    
    calculator.load_curve_params(internal_params)
    
    # 计算指定里程的坐标
    test_stations = [1000.0, 1100.0, 1200.0, 1300.0, 1400.0]
    
    print("\n里程\t\tX 坐标\t\tY 坐标\t\t方位角")
    print("-" * 60)
    
    for station in test_stations:
        result = calculator.calculate_point(station)
        if result:
            x, y, azimuth = result
            azimuth_deg = azimuth * 180 / 3.14159265358979
            print(f"K{station:.0f}\t\t{x:.3f}\t\t{y:.3f}\t\t{azimuth_deg:.2f}°")
        else:
            print(f"K{station:.0f}\t\t计算失败")
    
    # 坐标反算演示
    print("\n坐标反算演示:")
    target_x, target_y = 50100.0, 50050.0
    initial_station = 1050.0
    
    station, offset = calculator.inverse_calculation(target_x, target_y, initial_station)
    print(f"目标坐标：({target_x:.3f}, {target_y:.3f})")
    print(f"反算结果：里程 K{station:.3f}, 偏距 {offset:.3f}")


def demo_vertical_curve():
    """演示竖曲线计算"""
    print("\n" + "=" * 60)
    print("竖曲线计算演示")
    print("=" * 60)
    
    # 创建竖曲线计算器
    calculator = VerticalCurve()
    
    # 定义变坡点 [(里程，高程), ...]
    grade_points = [
        (1000.0, 100.0),
        (1500.0, 105.0),
        (2000.0, 102.0),
        (2500.0, 108.0)
    ]
    
    # 竖曲线半径
    radius = 10000.0
    
    # 计算指定里程的高程
    test_stations = [1000.0, 1250.0, 1500.0, 1750.0, 2000.0, 2250.0, 2500.0]
    
    print("\n里程\t\t设计高程")
    print("-" * 40)
    
    for station in test_stations:
        elevation = calculator.calculate_elevation(station, grade_points, radius)
        print(f"K{station:.0f}\t\t{elevation:.3f}")


def demo_angle_conversion():
    """演示角度转换"""
    print("\n" + "=" * 60)
    print("角度转换演示")
    print("=" * 60)
    
    utils = AngleUtils()
    
    test_values = [153.240524, 45.301500, 90.0, 180.0, 270.453000]
    
    print("\n度。分秒格式\t→\t十进制度")
    print("-" * 40)
    
    for dms in test_values:
        decimal = utils.dms_to_decimal(dms)
        back = utils.decimal_to_dms(decimal)
        print(f"{dms:.6f}\t→\t{decimal:.6f}°\t(验证：{back:.6f})")


def main():
    """主函数"""
    print("\n")
    print("*" * 60)
    print("   高速公路坐标高程计算软件 - Python 重构版")
    print("   Highway Coordinate & Elevation Calculation Software")
    print("*" * 60)
    print()
    
    try:
        # 运行演示
        demo_angle_conversion()
        demo_horizontal_curve()
        demo_vertical_curve()
        
        print("\n" + "=" * 60)
        print("演示完成!")
        print("=" * 60)
        print("\n提示：")
        print("- 查看 README.md 了解详细使用说明")
        print("- 运行测试：pytest tests/")
        print("- 代码检查：flake8 . && mypy .")
        
    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

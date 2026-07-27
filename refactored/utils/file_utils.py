"""
文件操作工具

实现曲线参数文件的导入导出功能
支持 VB6 原格式和 JSON 格式
"""

import json
import os
from typing import List, Dict, Any
from pathlib import Path

from models.curve_params import CurveSegment, CurveType, VerticalCurveParams, VerticalGradePoint


class FileHandler:
    """文件处理器"""
    
    def __init__(self):
        self.supported_extensions = ['.txt', '.json', '.dat']
    
    def load_curve_file(self, filepath: str) -> List[CurveSegment]:
        """
        加载曲线参数文件
        
        Args:
            filepath: 文件路径
            
        Returns:
            曲线段列表
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"文件不存在：{filepath}")
        
        ext = Path(filepath).suffix.lower()
        
        if ext == '.json':
            return self._load_json(filepath)
        else:
            # 尝试 VB6 原格式
            return self._load_vb6_format(filepath)
    
    def _load_vb6_format(self, filepath: str) -> List[CurveSegment]:
        """
        加载 VB6 原格式文件
        
        原格式每行包含：类型，Ki，J，C，D，F，R，A，B，O
        
        Args:
            filepath: 文件路径
            
        Returns:
            曲线段列表
        """
        segments = []
        
        with open(filepath, 'r', encoding='gbk', errors='ignore') as f:
            lines = f.readlines()
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith("'"):
                continue
            
            parts = line.split(',')
            if len(parts) < 5:
                continue
            
            try:
                # 解析 VB6 格式
                curve_type = parts[0].strip() if parts[0].strip() else "直线"
                station = float(parts[1].strip()) if len(parts) > 1 and parts[1].strip() else 0.0
                x_coord = float(parts[2].strip()) if len(parts) > 2 and parts[2].strip() else 0.0
                y_coord = float(parts[3].strip()) if len(parts) > 3 and parts[3].strip() else 0.0
                azimuth = float(parts[4].strip()) if len(parts) > 4 and parts[4].strip() else 0.0
                
                radius = float(parts[5].strip()) if len(parts) > 5 and parts[5].strip() else 0.0
                spiral_a = float(parts[6].strip()) if len(parts) > 6 and parts[6].strip() else 0.0
                
                segment = CurveSegment(
                    curve_type=curve_type,
                    station=station,
                    x_coord=x_coord,
                    y_coord=y_coord,
                    azimuth=azimuth,
                    radius=radius,
                    spiral_a=spiral_a
                )
                segments.append(segment)
            except (ValueError, IndexError) as e:
                print(f"解析行失败：{line}, 错误：{e}")
                continue
        
        return segments
    
    def _load_json(self, filepath: str) -> List[CurveSegment]:
        """
        加载 JSON 格式文件
        
        Args:
            filepath: 文件路径
            
        Returns:
            曲线段列表
        """
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        segments = []
        for item in data.get('segments', []):
            segment = CurveSegment.from_dict(item)
            segments.append(segment)
        
        return segments
    
    def save_curve_file(self, segments: List[CurveSegment], filepath: str, 
                       format: str = 'json'):
        """
        保存曲线参数文件
        
        Args:
            segments: 曲线段列表
            filepath: 文件路径
            format: 保存格式 ('json' 或 'vb6')
        """
        if format == 'json':
            self._save_json(segments, filepath)
        else:
            self._save_vb6_format(segments, filepath)
    
    def _save_json(self, segments: List[CurveSegment], filepath: str):
        """保存为 JSON 格式"""
        data = {
            'version': '1.0',
            'segments': [seg.to_dict() for seg in segments]
        }
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    def _save_vb6_format(self, segments: List[CurveSegment], filepath: str):
        """
        保存为 VB6 原格式
        
        格式：类型，Ki，J，C，D，F，R，A，B，O
        """
        with open(filepath, 'w', encoding='gbk') as f:
            for seg in segments:
                line = f"{seg.curve_type.value},{seg.station},{seg.x_coord},{seg.y_coord},{seg.azimuth},{seg.radius},{seg.spiral_a}\n"
                f.write(line)
    
    def export_to_cad_script(self, segments: List[CurveSegment], 
                            output_path: str,
                            precision: int = 3) -> str:
        """
        导出坐标到 CAD 脚本
        
        Args:
            segments: 曲线段列表
            output_path: 输出文件路径
            precision: 坐标精度
            
        Returns:
            生成的脚本内容
        """
        from core.curve import HorizontalCurve
        
        calculator = HorizontalCurve()
        calculator.load_curve_params([
            type('obj', (object,), {
                'curve_type': s.curve_type.value,
                'station': s.station,
                'x_coord': s.x_coord,
                'y_coord': s.y_coord,
                'azimuth': s.azimuth,
                'radius': s.radius,
                'spiral_a': s.spiral_a
            })() for s in segments
        ])
        
        script_lines = ["; Highway Coordinate Script", "; Generated by Refactored Software"]
        
        # 生成特征点
        for seg in segments:
            result = calculator.calculate_point(seg.station)
            if result:
                x, y, _ = result
                script_lines.append(f"POINT {round(x, precision)},{round(y, precision)},0")
        
        script_content = "\n".join(script_lines)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        return script_content
    
    def export_to_excel_csv(self, stations: List[float], 
                           coordinates: List[tuple],
                           output_path: str):
        """
        导出到 Excel CSV 格式
        
        Args:
            stations: 里程列表
            coordinates: 坐标列表 [(x, y, elevation), ...]
            output_path: 输出文件路径
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("里程，X 坐标，Y 坐标，高程\n")
            for station, coord in zip(stations, coordinates):
                x, y, elev = coord
                f.write(f"{station},{x:.3f},{y:.3f},{elev:.3f}\n")


# 导出公共接口
__all__ = ['FileHandler']

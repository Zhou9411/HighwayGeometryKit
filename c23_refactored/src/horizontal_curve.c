/**
 * @file horizontal_curve.c
 * @brief 平曲线计算函数实现
 * @version C23 Standard
 */

#include "horizontal_curve.h"
#include "coordinate_calculator.h"
#include "angle_utils.h"
#include <stdlib.h>
#include <string.h>
#include <math.h>

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

// 默认初始容量
#define DEFAULT_CAPACITY 16

int horizontal_curve_init(HorizontalCurveCalculator* calc) {
    if (calc == NULL) {
        return -1;
    }
    
    calc->segments = malloc(DEFAULT_CAPACITY * sizeof(CurveSegment));
    if (calc->segments == NULL) {
        return -1;
    }
    
    calc->segment_count = 0;
    calc->capacity = DEFAULT_CAPACITY;
    
    return 0;
}

void horizontal_curve_free(HorizontalCurveCalculator* calc) {
    if (calc != NULL && calc->segments != NULL) {
        free(calc->segments);
        calc->segments = NULL;
        calc->segment_count = 0;
        calc->capacity = 0;
    }
}

int horizontal_curve_load_params(HorizontalCurveCalculator* calc, 
                                 const CurveSegment* segments, 
                                 size_t count) {
    if (calc == NULL || segments == NULL || count == 0) {
        return -1;
    }
    
    // 如果容量不足，重新分配
    if (count > calc->capacity) {
        CurveSegment* new_segments = realloc(calc->segments, 
                                             count * sizeof(CurveSegment));
        if (new_segments == NULL) {
            return -1;
        }
        calc->segments = new_segments;
        calc->capacity = count;
    }
    
    // 复制数据
    memcpy(calc->segments, segments, count * sizeof(CurveSegment));
    calc->segment_count = count;
    
    return 0;
}

int find_curve_segment(const HorizontalCurveCalculator* calc, double station) {
    if (calc == NULL || calc->segment_count == 0) {
        return -1;
    }
    
    size_t idx = 0;
    const bool is_ascending = calc->segments[0].station < calc->segments[calc->segment_count - 1].station;
    
    if (is_ascending) {
        while (idx < calc->segment_count - 1 && station > calc->segments[idx].station) {
            idx++;
        }
    } else {
        while (idx < calc->segment_count - 1 && station < calc->segments[idx].station) {
            idx++;
        }
    }
    
    return (idx < calc->segment_count) ? (int)idx : 0;
}

int calculate_line(const HorizontalCurveCalculator* calc, double station, 
                   int segment_idx, CalculationResult* result) {
    if (calc == NULL || result == NULL || segment_idx < 0 || 
        (size_t)segment_idx >= calc->segment_count) {
        return -1;
    }
    
    const CurveSegment* param = &calc->segments[segment_idx];
    const double delta_station = station - param->station;
    const double azimuth_rad = deg_to_rad(dms_to_decimal(param->azimuth));
    
    result->x = param->x_coord + delta_station * cos(azimuth_rad);
    result->y = param->y_coord + delta_station * sin(azimuth_rad);
    result->azimuth = azimuth_rad;
    result->success = true;
    
    return 0;
}

int calculate_circle(const HorizontalCurveCalculator* calc, double station, 
                     int segment_idx, CalculationResult* result) {
    if (calc == NULL || result == NULL || segment_idx < 0 || 
        (size_t)segment_idx >= calc->segment_count) {
        return -1;
    }
    
    const CurveSegment* param = &calc->segments[segment_idx];
    const double arc_length = station - param->station;
    const double radius = param->radius;
    
    if (radius == 0.0) {
        return calculate_line(calc, station, segment_idx, result);
    }
    
    // 圆心角 (弧度)
    const double central_angle = arc_length / fabs(radius);
    
    // 方向判断
    const int direction = (radius > 0.0) ? 1 : -1;
    
    // 起算方位角
    const double start_azimuth = deg_to_rad(dms_to_decimal(param->azimuth));
    
    // 切线方位角
    const double tangent_azimuth = start_azimuth + direction * central_angle;
    
    // 弦切角
    const double chord_angle = central_angle / 2.0;
    
    // 弦长
    const double chord_length = 2.0 * fabs(radius) * sin(central_angle / 2.0);
    
    // 弦方位角
    const double chord_azimuth = start_azimuth + direction * chord_angle;
    
    // 计算坐标
    result->x = param->x_coord + chord_length * cos(chord_azimuth);
    result->y = param->y_coord + chord_length * sin(chord_azimuth);
    result->azimuth = tangent_azimuth;
    result->success = true;
    
    return 0;
}

int calculate_spiral(const HorizontalCurveCalculator* calc, double station, 
                     int segment_idx, CalculationResult* result) {
    if (calc == NULL || result == NULL || segment_idx < 0 || 
        (size_t)segment_idx >= calc->segment_count) {
        return -1;
    }
    
    const CurveSegment* param = &calc->segments[segment_idx];
    const double l = station - param->station;  // 缓和曲线上点到起点距离
    const double a = param->spiral_a;
    const double r = param->radius;
    
    // 退化情况，按直线处理
    if (a == 0.0 || r == 0.0) {
        return calculate_line(calc, station, segment_idx, result);
    }
    
    // 缓和曲线全长
    const double ls = fabs(a * a / r);
    
    // 切线角 (弧度)
    const double beta = l * l / (2.0 * a * a);
    
    // 起算方位角
    const double start_azimuth = deg_to_rad(dms_to_decimal(param->azimuth));
    
    // 方向判断
    const int direction = (r > 0.0) ? 1 : -1;
    
    // 当前点方位角
    const double current_azimuth = start_azimuth + direction * beta;
    
    // 使用级数展开计算坐标增量 (近似公式)
    const double l2 = l * l;
    const double l4 = l2 * l2;
    const double a2 = a * a;
    const double a4 = a2 * a2;
    const double a6 = a4 * a2;
    
    double x_local = l - l4 / (40.0 * a4);
    double y_local = l2 * l / (6.0 * a2) - l2 * l4 / (336.0 * a6);
    
    // 坐标转换
    const double cos_f = cos(start_azimuth);
    const double sin_f = sin(start_azimuth);
    
    if (direction < 0) {
        y_local = -y_local;
    }
    
    result->x = param->x_coord + x_local * cos_f - y_local * sin_f;
    result->y = param->y_coord + x_local * sin_f + y_local * cos_f;
    result->azimuth = current_azimuth;
    result->success = true;
    
    return 0;
}

int horizontal_curve_calculate_point(const HorizontalCurveCalculator* calc, 
                                     double station, 
                                     CalculationResult* result) {
    if (calc == NULL || result == NULL || calc->segment_count == 0) {
        return -1;
    }
    
    const int segment_idx = find_curve_segment(calc, station);
    if (segment_idx < 0) {
        return -1;
    }
    
    const CurveSegment* param = &calc->segments[segment_idx];
    
    switch (param->curve_type) {
        case CURVE_TYPE_LINE:
            return calculate_line(calc, station, segment_idx, result);
        case CURVE_TYPE_CIRCLE:
            return calculate_circle(calc, station, segment_idx, result);
        case CURVE_TYPE_SPIRAL:
        case CURVE_TYPE_COMPOUND:
            return calculate_spiral(calc, station, segment_idx, result);
        default:
            return calculate_line(calc, station, segment_idx, result);
    }
}

int horizontal_curve_inverse_calculation(const HorizontalCurveCalculator* calc,
                                         double target_x, double target_y,
                                         double initial_station,
                                         CalculationResult* result) {
    if (calc == NULL || result == NULL || calc->segment_count == 0) {
        return -1;
    }
    
    double station = initial_station;
    double distance = 0.0;
    
    // 最大迭代次数
    for (int iter = 0; iter < 100; iter++) {
        CalculationResult calc_result = {0};
        if (horizontal_curve_calculate_point(calc, station, &calc_result) != 0) {
            break;
        }
        
        const double calc_x = calc_result.x;
        const double calc_y = calc_result.y;
        const double azimuth = calc_result.azimuth;
        
        // 计算距离和方位角
        const double dx = target_x - calc_x;
        const double dy = target_y - calc_y;
        distance = sqrt(dx * dx + dy * dy);
        
        // 精度满足要求
        if (distance < 0.001) {
            break;
        }
        
        const double line_azimuth = calculate_azimuth(calc_x, calc_y, target_x, target_y);
        
        // 计算沿路线方向的投影
        const double angle_diff = line_azimuth - azimuth;
        const double station_correction = distance * cos(angle_diff);
        
        station += station_correction * 0.382;  // 黄金分割系数
    }
    
    // 计算最终偏距
    CalculationResult final_result = {0};
    if (horizontal_curve_calculate_point(calc, station, &final_result) == 0) {
        const double calc_x = final_result.x;
        const double calc_y = final_result.y;
        const double azimuth = final_result.azimuth;
        
        const double dx = target_x - calc_x;
        const double dy = target_y - calc_y;
        const double line_azimuth = calculate_azimuth(calc_x, calc_y, target_x, target_y);
        const double angle_diff = line_azimuth - azimuth;
        
        double offset = 0.0;
        if (fabs(angle_diff) > 1e-10) {
            offset = distance * sin(angle_diff);
            offset = (offset >= 0.0) ? fabs(offset) : -fabs(offset);
        }
        
        result->station = station;
        result->offset = offset;
        result->success = true;
        return 0;
    }
    
    return -1;
}

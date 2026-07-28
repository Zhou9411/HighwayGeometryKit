/**
 * @file vertical_curve.c
 * @brief 竖曲线计算函数实现
 * @version C23 Standard
 */

#include "vertical_curve.h"
#include <stdlib.h>
#include <string.h>
#include <math.h>

// 默认初始容量
#define DEFAULT_CAPACITY 16

int vertical_curve_params_init(VerticalCurveParams* params, size_t capacity) {
    if (params == NULL || capacity == 0) {
        return -1;
    }
    
    params->grade_points = malloc(capacity * sizeof(VerticalGradePoint));
    if (params->grade_points == NULL) {
        return -1;
    }
    
    params->radius = 0.0;
    params->grade_point_count = 0;
    params->capacity = capacity;
    
    return 0;
}

void vertical_curve_params_free(VerticalCurveParams* params) {
    if (params != NULL && params->grade_points != NULL) {
        free(params->grade_points);
        params->grade_points = NULL;
        params->grade_point_count = 0;
        params->capacity = 0;
    }
}

int vertical_curve_add_grade_point(VerticalCurveParams* params, 
                                   double station, double elevation) {
    if (params == NULL) {
        return -1;
    }
    
    // 如果容量不足，重新分配
    if (params->grade_point_count >= params->capacity) {
        const size_t new_capacity = params->capacity * 2;
        VerticalGradePoint* new_points = realloc(params->grade_points, 
                                                  new_capacity * sizeof(VerticalGradePoint));
        if (new_points == NULL) {
            return -1;
        }
        params->grade_points = new_points;
        params->capacity = new_capacity;
    }
    
    // 添加新点
    params->grade_points[params->grade_point_count].station = station;
    params->grade_points[params->grade_point_count].elevation = elevation;
    params->grade_point_count++;
    
    // 按里程排序 (简单插入排序)
    for (size_t i = params->grade_point_count - 1; i > 0; i--) {
        if (params->grade_points[i].station < params->grade_points[i - 1].station) {
            VerticalGradePoint temp = params->grade_points[i];
            params->grade_points[i] = params->grade_points[i - 1];
            params->grade_points[i - 1] = temp;
        } else {
            break;
        }
    }
    
    return 0;
}

int vertical_curve_calculate_elevation(double station, 
                                       const VerticalCurveParams* params,
                                       CalculationResult* result) {
    if (params == NULL || result == NULL || params->grade_point_count == 0) {
        return -1;
    }
    
    // 找到相邻的变坡点
    size_t i = 0;
    for (size_t j = 0; j < params->grade_point_count - 1; j++) {
        if (params->grade_points[j].station <= station && 
            station <= params->grade_points[j + 1].station) {
            i = j;
            break;
        }
    }
    
    // 边界处理
    if (i >= params->grade_point_count - 1) {
        i = params->grade_point_count - 2;
    }
    
    // 计算前后坡度
    double grade1 = 0.0;
    double grade2 = 0.0;
    
    if (i > 0) {
        grade1 = (params->grade_points[i].elevation - params->grade_points[i - 1].elevation) /
                 (params->grade_points[i].station - params->grade_points[i - 1].station);
    }
    
    if (i < params->grade_point_count - 1) {
        grade2 = (params->grade_points[i + 1].elevation - params->grade_points[i].elevation) /
                 (params->grade_points[i + 1].station - params->grade_points[i].station);
    }
    
    // 坡度差
    const double omega = grade2 - grade1;
    
    // 切线长
    const double t = fabs(params->radius * omega / 2.0);
    
    // 变坡点信息
    const double pvi_station = params->grade_points[i].station;
    const double pvi_elevation = params->grade_points[i].elevation;
    
    // 判断是凸曲线还是凹曲线
    const bool is_convex = omega < 0.0;
    
    // 计算切线高程
    double tangent_elevation;
    if (station <= pvi_station) {
        tangent_elevation = pvi_elevation + grade1 * (station - pvi_station);
    } else {
        tangent_elevation = pvi_elevation + grade2 * (station - pvi_station);
    }
    
    // 计算竖曲线改正值
    const double x = fabs(station - pvi_station);
    double elevation;
    
    if (x <= t) {
        double correction = x * x / (2.0 * params->radius);
        if (is_convex) {
            correction = -correction;
        }
        elevation = tangent_elevation + correction;
    } else {
        elevation = tangent_elevation;
    }
    
    result->elevation = elevation;
    result->success = true;
    
    return 0;
}

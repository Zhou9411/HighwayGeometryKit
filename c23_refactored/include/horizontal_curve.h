/**
 * @file horizontal_curve.h
 * @brief 平曲线计算函数声明
 * @version C23 Standard
 */

#ifndef HORIZONTAL_CURVE_H
#define HORIZONTAL_CURVE_H

#include "curve_types.h"

/**
 * @brief 平曲线计算器结构
 */
typedef struct {
    CurveSegment* segments;   ///< 曲线段数组
    size_t segment_count;     ///< 曲线段数量
    size_t capacity;          ///< 数组容量
} HorizontalCurveCalculator;

/**
 * @brief 初始化平曲线计算器
 * 
 * @param calc 计算器指针
 * @return int 成功返回 0，失败返回 -1
 */
int horizontal_curve_init(HorizontalCurveCalculator* calc);

/**
 * @brief 释放平曲线计算器资源
 * 
 * @param calc 计算器指针
 */
void horizontal_curve_free(HorizontalCurveCalculator* calc);

/**
 * @brief 加载曲线参数
 * 
 * @param calc 计算器指针
 * @param segments 曲线段数组
 * @param count 曲线段数量
 * @return int 成功返回 0，失败返回 -1
 */
int horizontal_curve_load_params(HorizontalCurveCalculator* calc, 
                                 const CurveSegment* segments, 
                                 size_t count);

/**
 * @brief 根据里程查找对应的曲线段索引
 * 
 * @param calc 计算器指针
 * @param station 里程值
 * @return int 曲线段索引，失败返回 -1
 */
[[nodiscard]] int find_curve_segment(const HorizontalCurveCalculator* calc, double station);

/**
 * @brief 直线段坐标计算
 * 
 * @param calc 计算器指针
 * @param station 待求点里程
 * @param segment_idx 曲线段索引
 * @param[out] result 计算结果
 * @return int 成功返回 0，失败返回 -1
 */
int calculate_line(const HorizontalCurveCalculator* calc, double station, 
                   int segment_idx, CalculationResult* result);

/**
 * @brief 圆曲线段坐标计算
 * 
 * @param calc 计算器指针
 * @param station 待求点里程
 * @param segment_idx 曲线段索引
 * @param[out] result 计算结果
 * @return int 成功返回 0，失败返回 -1
 */
int calculate_circle(const HorizontalCurveCalculator* calc, double station, 
                     int segment_idx, CalculationResult* result);

/**
 * @brief 缓和曲线段坐标计算
 * 
 * @param calc 计算器指针
 * @param station 待求点里程
 * @param segment_idx 曲线段索引
 * @param[out] result 计算结果
 * @return int 成功返回 0，失败返回 -1
 */
int calculate_spiral(const HorizontalCurveCalculator* calc, double station, 
                     int segment_idx, CalculationResult* result);

/**
 * @brief 根据里程计算坐标和方位角
 * 
 * @param calc 计算器指针
 * @param station 待求点里程
 * @param[out] result 计算结果
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int horizontal_curve_calculate_point(const HorizontalCurveCalculator* calc, 
                                                   double station, 
                                                   CalculationResult* result);

/**
 * @brief 坐标反算：已知坐标求里程和偏距
 * 
 * 使用逐步趋近法
 * 
 * @param calc 计算器指针
 * @param target_x 目标点 X 坐标
 * @param target_y 目标点 Y 坐标
 * @param initial_station 初始估算里程
 * @param[out] result 计算结果 (station 和 offset)
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int horizontal_curve_inverse_calculation(const HorizontalCurveCalculator* calc,
                                                       double target_x, double target_y,
                                                       double initial_station,
                                                       CalculationResult* result);

#endif /* HORIZONTAL_CURVE_H */

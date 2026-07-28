/**
 * @file vertical_curve.h
 * @brief 竖曲线计算函数声明
 * @version C23 Standard
 */

#ifndef VERTICAL_CURVE_H
#define VERTICAL_CURVE_H

#include "curve_types.h"

/**
 * @brief 初始化竖曲线参数
 * 
 * @param params 竖曲线参数指针
 * @param capacity 初始容量
 * @return int 成功返回 0，失败返回 -1
 */
int vertical_curve_params_init(VerticalCurveParams* params, size_t capacity);

/**
 * @brief 释放竖曲线参数资源
 * 
 * @param params 竖曲线参数指针
 */
void vertical_curve_params_free(VerticalCurveParams* params);

/**
 * @brief 添加变坡点
 * 
 * @param params 竖曲线参数指针
 * @param station 里程
 * @param elevation 高程
 * @return int 成功返回 0，失败返回 -1
 */
int vertical_curve_add_grade_point(VerticalCurveParams* params, 
                                   double station, double elevation);

/**
 * @brief 竖曲线高程计算
 * 
 * @param station 待求点里程
 * @param params 竖曲线参数
 * @param[out] result 计算结果 (elevation)
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int vertical_curve_calculate_elevation(double station, 
                                                     const VerticalCurveParams* params,
                                                     CalculationResult* result);

#endif /* VERTICAL_CURVE_H */

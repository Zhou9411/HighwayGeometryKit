/**
 * @file coordinate_calculator.h
 * @brief 坐标计算函数声明
 * @version C23 Standard (兼容 C17)
 */

#ifndef COORDINATE_CALCULATOR_H
#define COORDINATE_CALCULATOR_H

#include "curve_types.h"

/**
 * @brief 计算两点间的方位角
 * 
 * @param x1 起点 X 坐标
 * @param y1 起点 Y 坐标
 * @param x2 终点 X 坐标
 * @param y2 终点 Y 坐标
 * @return double 方位角 (弧度)
 */
double calculate_azimuth(double x1, double y1, double x2, double y2);

/**
 * @brief 计算两点间的距离
 * 
 * @param x1 起点 X 坐标
 * @param y1 起点 Y 坐标
 * @param x2 终点 X 坐标
 * @param y2 终点 Y 坐标
 * @return double 距离
 */
double calculate_distance(double x1, double y1, double x2, double y2);

/**
 * @brief 极坐标转直角坐标
 * 
 * @param x0 起点 X 坐标
 * @param y0 起点 Y 坐标
 * @param distance 距离
 * @param azimuth 方位角 (弧度)
 * @param[out] x 终点 X 坐标
 * @param[out] y 终点 Y 坐标
 */
void polar_to_rectangular(double x0, double y0, double distance, double azimuth, 
                         double* x, double* y);

#endif /* COORDINATE_CALCULATOR_H */

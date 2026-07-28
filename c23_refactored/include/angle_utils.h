/**
 * @file angle_utils.h
 * @brief 角度转换工具函数声明
 * @version C23 Standard (兼容 C17)
 */

#ifndef ANGLE_UTILS_H
#define ANGLE_UTILS_H

#include <math.h>

/**
 * @brief 将度分秒格式转换为十进制度
 * 
 * @param dms 度分秒格式的角度值 (如 153.240524 表示 153°24'05.24")
 * @return double 十进制度数
 */
double dms_to_decimal(double dms);

/**
 * @brief 将十进制度转换为度分秒格式
 * 
 * @param decimal 十进制度数
 * @return double 度分秒格式的角度值
 */
double decimal_to_dms(double decimal);

/**
 * @brief 规范化方位角到 0-2π范围
 * 
 * @param azimuth 方位角 (弧度)
 * @return double 规范化的方位角
 */
double normalize_azimuth(double azimuth);

/**
 * @brief 角度转弧度
 * 
 * @param degrees 角度值
 * @return double 弧度值
 */
static inline double deg_to_rad(double degrees) {
    return degrees * 0.017453292519943295;  // π / 180
}

/**
 * @brief 弧度转角度
 * 
 * @param radians 弧度值
 * @return double 角度值
 */
static inline double rad_to_deg(double radians) {
    return radians * 57.29577951308232;  // 180 / π
}

#endif /* ANGLE_UTILS_H */

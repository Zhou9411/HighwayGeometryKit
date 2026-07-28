/**
 * @file angle_utils.c
 * @brief 角度转换工具函数实现
 * @version C23 Standard
 */

#include "angle_utils.h"
#include <math.h>

// 数学常数定义
#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

double dms_to_decimal(double dms) {
    if (dms == 0.0) {
        return 0.0;
    }
    
    const double sign = (dms >= 0.0) ? 1.0 : -1.0;
    dms = fabs(dms);
    
    const double degrees = trunc(dms);
    const double minutes_decimal = (dms - degrees) * 100.0;
    const double minutes = trunc(minutes_decimal);
    const double seconds = (minutes_decimal - minutes) * 100.0;
    
    const double decimal_degrees = degrees + minutes / 60.0 + seconds / 3600.0;
    return sign * decimal_degrees;
}

double decimal_to_dms(double decimal) {
    if (decimal == 0.0) {
        return 0.0;
    }
    
    const double sign = (decimal >= 0.0) ? 1.0 : -1.0;
    decimal = fabs(decimal);
    
    const double degrees = trunc(decimal);
    const double minutes_decimal = (decimal - degrees) * 60.0;
    const double minutes = trunc(minutes_decimal);
    const double seconds = (minutes_decimal - minutes) * 60.0;
    
    const double dms = degrees + minutes / 100.0 + seconds / 10000.0;
    return sign * dms;
}

double normalize_azimuth(double azimuth) {
    const double two_pi = 2.0 * M_PI;
    azimuth = fmod(azimuth, two_pi);
    if (azimuth < 0.0) {
        azimuth += two_pi;
    }
    return azimuth;
}

/**
 * @file coordinate_calculator.c
 * @brief 坐标计算函数实现
 * @version C23 Standard (兼容 C17)
 */

#include "coordinate_calculator.h"
#include "angle_utils.h"
#include <math.h>

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

#ifndef M_PI_2
#define M_PI_2 (M_PI / 2.0)
#endif

double calculate_azimuth(double x1, double y1, double x2, double y2) {
    const double dx = x2 - x1;
    const double dy = y2 - y1;
    double azimuth;
    
    if (dx != 0.0 && dy != 0.0) {
        azimuth = atan(dy / dx);
        if (dx < 0.0) {
            azimuth += M_PI;
        } else if (dy < 0.0) {
            azimuth += 2.0 * M_PI;
        }
    } else if (dx == 0.0) {
        azimuth = (dy > 0.0) ? M_PI_2 : 3.0 * M_PI_2;
    } else {  // dy == 0.0
        azimuth = (dx > 0.0) ? 0.0 : M_PI;
    }
    
    return normalize_azimuth(azimuth);
}

double calculate_distance(double x1, double y1, double x2, double y2) {
    const double dx = x2 - x1;
    const double dy = y2 - y1;
    return sqrt(dx * dx + dy * dy);
}

void polar_to_rectangular(double x0, double y0, double distance, double azimuth, 
                         double* x, double* y) {
    *x = x0 + distance * cos(azimuth);
    *y = y0 + distance * sin(azimuth);
}

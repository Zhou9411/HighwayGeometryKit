/**
 * @file main.c
 * @brief 高速公路坐标高程计算软件 - C23 重构版主程序
 * @version C23 Standard
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>

#include "curve_types.h"
#include "angle_utils.h"
#include "coordinate_calculator.h"
#include "horizontal_curve.h"
#include "vertical_curve.h"
#include "file_utils.h"

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

// 打印分隔线
static void print_separator(const char* title, char ch) {
    printf("\n");
    for (int i = 0; i < 60; i++) putchar(ch);
    printf("\n");
    if (title != NULL) {
        printf("%s\n", title);
        for (int i = 0; i < 60; i++) putchar(ch);
    }
    printf("\n");
}

// 演示角度转换
static void demo_angle_conversion(void) {
    print_separator("角度转换演示", '=');
    
    const double test_values[] = {153.240524, 45.301500, 90.0, 180.0, 270.453000};
    const size_t count = sizeof(test_values) / sizeof(test_values[0]);
    
    printf("度。分秒格式\t→\t十进制度\n");
    printf("----------------------------------------\n");
    
    for (size_t i = 0; i < count; i++) {
        const double dms = test_values[i];
        const double decimal = dms_to_decimal(dms);
        const double back = decimal_to_dms(decimal);
        printf("%.6f\t→\t%.6f°\t(验证：%.6f)\n", dms, decimal, back);
    }
}

// 演示平曲线计算
static void demo_horizontal_curve(void) {
    print_separator("平曲线计算演示", '=');
    
    // 创建曲线参数
    CurveSegment segments[3] = {
        {
            .curve_type = CURVE_TYPE_SPIRAL,
            .station = 1000.0,
            .x_coord = 50000.0,
            .y_coord = 50000.0,
            .azimuth = 153.240524,
            .radius = -800.0,
            .spiral_a = 70.0,
            .offset_x = 0.0,
            .offset_y = 0.0
        },
        {
            .curve_type = CURVE_TYPE_CIRCLE,
            .station = 1200.0,
            .x_coord = 49850.0,
            .y_coord = 50120.0,
            .azimuth = 0.0,
            .radius = -800.0,
            .spiral_a = 0.0,
            .offset_x = 0.0,
            .offset_y = 0.0
        },
        {
            .curve_type = CURVE_TYPE_LINE,
            .station = 1500.0,
            .x_coord = 49700.0,
            .y_coord = 50200.0,
            .azimuth = 0.0,
            .radius = 0.0,
            .spiral_a = 0.0,
            .offset_x = 0.0,
            .offset_y = 0.0
        }
    };
    
    // 初始化计算器
    HorizontalCurveCalculator calc;
    if (horizontal_curve_init(&calc) != 0) {
        fprintf(stderr, "初始化平曲线计算器失败\n");
        return;
    }
    
    if (horizontal_curve_load_params(&calc, segments, 3) != 0) {
        fprintf(stderr, "加载曲线参数失败\n");
        horizontal_curve_free(&calc);
        return;
    }
    
    // 计算指定里程的坐标
    const double test_stations[] = {1000.0, 1100.0, 1200.0, 1300.0, 1400.0};
    const size_t count = sizeof(test_stations) / sizeof(test_stations[0]);
    
    printf("里程\t\tX 坐标\t\tY 坐标\t\t方位角\n");
    printf("------------------------------------------------------------\n");
    
    for (size_t i = 0; i < count; i++) {
        CalculationResult result = {0};
        if (horizontal_curve_calculate_point(&calc, test_stations[i], &result) == 0) {
            const double azimuth_deg = rad_to_deg(result.azimuth);
            printf("K%.0f\t\t%.3f\t\t%.3f\t\t%.2f°\n", 
                   test_stations[i], result.x, result.y, azimuth_deg);
        } else {
            printf("K%.0f\t\t计算失败\n", test_stations[i]);
        }
    }
    
    // 坐标反算演示
    printf("\n坐标反算演示:\n");
    const double target_x = 50100.0;
    const double target_y = 50050.0;
    const double initial_station = 1050.0;
    
    CalculationResult inverse_result = {0};
    if (horizontal_curve_inverse_calculation(&calc, target_x, target_y, 
                                             initial_station, &inverse_result) == 0) {
        printf("目标坐标：(%.3f, %.3f)\n", target_x, target_y);
        printf("反算结果：里程 K%.3f, 偏距 %.3f\n", 
               inverse_result.station, inverse_result.offset);
    } else {
        printf("坐标反算失败\n");
    }
    
    horizontal_curve_free(&calc);
}

// 演示竖曲线计算
static void demo_vertical_curve(void) {
    print_separator("竖曲线计算演示", '=');
    
    // 初始化竖曲线参数
    VerticalCurveParams params;
    if (vertical_curve_params_init(&params, 16) != 0) {
        fprintf(stderr, "初始化竖曲线参数失败\n");
        return;
    }
    
    params.radius = 10000.0;
    
    // 添加变坡点
    vertical_curve_add_grade_point(&params, 1000.0, 100.0);
    vertical_curve_add_grade_point(&params, 1500.0, 105.0);
    vertical_curve_add_grade_point(&params, 2000.0, 102.0);
    vertical_curve_add_grade_point(&params, 2500.0, 108.0);
    
    // 计算指定里程的高程
    const double test_stations[] = {1000.0, 1250.0, 1500.0, 1750.0, 2000.0, 2250.0, 2500.0};
    const size_t count = sizeof(test_stations) / sizeof(test_stations[0]);
    
    printf("里程\t\t设计高程\n");
    printf("----------------------------------------\n");
    
    for (size_t i = 0; i < count; i++) {
        CalculationResult result = {0};
        if (vertical_curve_calculate_elevation(test_stations[i], &params, &result) == 0) {
            printf("K%.0f\t\t%.3f\n", test_stations[i], result.elevation);
        } else {
            printf("K%.0f\t\t计算失败\n", test_stations[i]);
        }
    }
    
    vertical_curve_params_free(&params);
}

// 主函数
int main(int argc, char* argv[]) {
    printf("\n");
    printf("************************************************************\n");
    printf("   高速公路坐标高程计算软件 - C23 重构版\n");
    printf("   Highway Coordinate & Elevation Calculation Software\n");
    printf("************************************************************\n");
    
    // 运行演示
    demo_angle_conversion();
    demo_horizontal_curve();
    demo_vertical_curve();
    
    print_separator("演示完成!", '=');
    
    printf("\n提示:\n");
    printf("- 查看 README.md 了解详细使用说明\n");
    printf("- 编译：make 或 gcc -std=c23 -O2 -o main src/*.c\n");
    printf("- 测试：make test\n");
    
    return 0;
}

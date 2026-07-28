/**
 * @file curve_types.h
 * @brief 曲线类型和数据结构的定义
 * @version C23 Standard
 */

#ifndef CURVE_TYPES_H
#define CURVE_TYPES_H

#include <stdbool.h>
#include <stddef.h>

/**
 * @brief 曲线类型枚举
 */
typedef enum {
    CURVE_TYPE_LINE = 0,      ///< 直线段
    CURVE_TYPE_CIRCLE = 1,    ///< 圆曲线
    CURVE_TYPE_SPIRAL = 2,    ///< 缓和曲线 (回旋曲线)
    CURVE_TYPE_COMPOUND = 3   ///< 复合曲线 (1+ 圆 +2)
} CurveType;

/**
 * @brief 竖曲线变坡点
 */
typedef struct {
    double station;     ///< 里程
    double elevation;   ///< 高程
} VerticalGradePoint;

/**
 * @brief 平曲线段参数
 * 
 * 用于描述一个曲线段的完整参数
 */
typedef struct {
    CurveType curve_type;   ///< 曲线类型
    double station;         ///< 起算点里程 (K)
    double x_coord;         ///< 起算点 X 坐标 (C)
    double y_coord;         ///< 起算点 Y 坐标 (D)
    double azimuth;         ///< 起算点切线方位角 (F, 度分秒格式)
    double radius;          ///< 半径 (R, 左偏负右偏正)
    double spiral_a;        ///< 回旋参数 (A)
    double offset_x;        ///< X 方向偏移
    double offset_y;        ///< Y 方向偏移
} CurveSegment;

/**
 * @brief 竖曲线参数
 */
typedef struct {
    double radius;                  ///< 竖曲线半径
    VerticalGradePoint* grade_points; ///< 变坡点数组
    size_t grade_point_count;       ///< 变坡点数量
    size_t capacity;                ///< 数组容量
} VerticalCurveParams;

/**
 * @brief 宽度渐变参数
 */
typedef struct {
    double start_station;   ///< 起点里程
    double end_station;     ///< 终点里程
    double start_width;     ///< 起点宽度
    double end_width;       ///< 终点宽度
    double left_offset;     ///< 左侧偏移
    double right_offset;    ///< 右侧偏移
} WidthTransitionParams;

/**
 * @brief 横断面参数
 */
typedef struct {
    double station;         ///< 里程
    double* design_line;    ///< 设计线数据 [(偏距，高程), ...]
    double* ground_line;    ///< 地面线数据 [(偏距，高程), ...]
    size_t design_count;    ///< 设计线点数
    size_t ground_count;    ///< 地面线点数
    double text_height;     ///< 文字高度
    char font_name[64];     ///< 字体名称
} CrossSectionParams;

/**
 * @brief 计算结果结构
 */
typedef struct {
    bool success;           ///< 计算是否成功
    double x;               ///< X 坐标
    double y;               ///< Y 坐标
    double azimuth;         ///< 方位角 (弧度)
    double station;         ///< 里程
    double offset;          ///< 偏距
    double elevation;       ///< 高程
} CalculationResult;

#endif /* CURVE_TYPES_H */

/**
 * @file file_utils.h
 * @brief 文件操作工具函数声明
 * @version C23 Standard
 */

#ifndef FILE_UTILS_H
#define FILE_UTILS_H

#include "curve_types.h"

/**
 * @brief 加载曲线参数文件 (支持 VB6 原格式)
 * 
 * @param filepath 文件路径
 * @param[out] segments 曲线段数组指针
 * @param[out] count 曲线段数量
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int load_curve_file(const char* filepath, 
                                  CurveSegment** segments, 
                                  size_t* count);

/**
 * @brief 保存曲线参数文件为 JSON 格式
 * 
 * @param filepath 文件路径
 * @param segments 曲线段数组
 * @param count 曲线段数量
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int save_curve_file_json(const char* filepath, 
                                       const CurveSegment* segments, 
                                       size_t count);

/**
 * @brief 保存曲线参数文件为 VB6 原格式
 * 
 * @param filepath 文件路径
 * @param segments 曲线段数组
 * @param count 曲线段数量
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int save_curve_file_vb6(const char* filepath, 
                                      const CurveSegment* segments, 
                                      size_t count);

/**
 * @brief 导出坐标到 CAD 脚本
 * 
 * @param filepath 输出文件路径
 * @param segments 曲线段数组
 * @param count 曲线段数量
 * @param precision 坐标精度
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int export_to_cad_script(const char* filepath, 
                                       const CurveSegment* segments, 
                                       size_t count, 
                                       int precision);

/**
 * @brief 导出到 Excel CSV 格式
 * 
 * @param filepath 输出文件路径
 * @param stations 里程数组
 * @param coordinates 坐标数组 [(x, y, elevation), ...]
 * @param count 数据点数量
 * @return int 成功返回 0，失败返回 -1
 */
[[nodiscard]] int export_to_excel_csv(const char* filepath, 
                                      const double* stations, 
                                      const double (*coordinates)[3], 
                                      size_t count);

#endif /* FILE_UTILS_H */

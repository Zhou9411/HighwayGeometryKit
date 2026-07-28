/**
 * @file file_utils.c
 * @brief 文件操作工具函数实现
 * @version C23 Standard
 */

#include "file_utils.h"
#include "horizontal_curve.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <errno.h>

// 获取文件扩展名
static const char* get_extension(const char* filepath) {
    const char* dot = strrchr(filepath, '.');
    return (dot != NULL) ? dot : "";
}

int load_curve_file(const char* filepath, CurveSegment** segments, size_t* count) {
    if (filepath == NULL || segments == NULL || count == NULL) {
        return -1;
    }
    
    FILE* file = fopen(filepath, "r");
    if (file == NULL) {
        fprintf(stderr, "无法打开文件：%s\n", filepath);
        return -1;
    }
    
    // 初始容量
    size_t capacity = 16;
    *segments = malloc(capacity * sizeof(CurveSegment));
    if (*segments == NULL) {
        fclose(file);
        return -1;
    }
    *count = 0;
    
    const char* ext = get_extension(filepath);
    char line[1024];
    
    while (fgets(line, sizeof(line), file) != NULL) {
        // 跳过空行和注释
        char* trimmed = line;
        while (*trimmed == ' ' || *trimmed == '\t') trimmed++;
        if (*trimmed == '\0' || *trimmed == '\n' || *trimmed == '\'') {
            continue;
        }
        
        // 解析 CSV 格式：类型，Ki，J，C，D，F，R，A，B，O
        char curve_type_str[32] = {0};
        double station = 0.0, x_coord = 0.0, y_coord = 0.0, azimuth = 0.0;
        double radius = 0.0, spiral_a = 0.0;
        
        int parsed = sscanf(line, "%[^,],%lf,%lf,%lf,%lf,%lf,%lf",
                           curve_type_str, &station, &x_coord, &y_coord, 
                           &azimuth, &radius, &spiral_a);
        
        if (parsed < 5) {
            continue;  // 跳过无效行
        }
        
        // 确保容量足够
        if (*count >= capacity) {
            capacity *= 2;
            CurveSegment* new_segments = realloc(*segments, capacity * sizeof(CurveSegment));
            if (new_segments == NULL) {
                free(*segments);
                fclose(file);
                return -1;
            }
            *segments = new_segments;
        }
        
        // 解析曲线类型
        CurveType curve_type = CURVE_TYPE_LINE;
        if (strcmp(curve_type_str, "圆曲线") == 0) {
            curve_type = CURVE_TYPE_CIRCLE;
        } else if (strcmp(curve_type_str, "缓和曲线") == 0) {
            curve_type = CURVE_TYPE_SPIRAL;
        } else if (strcmp(curve_type_str, "1+ 圆 +2") == 0) {
            curve_type = CURVE_TYPE_COMPOUND;
        }
        
        // 填充曲线段
        (*segments)[*count].curve_type = curve_type;
        (*segments)[*count].station = station;
        (*segments)[*count].x_coord = x_coord;
        (*segments)[*count].y_coord = y_coord;
        (*segments)[*count].azimuth = azimuth;
        (*segments)[*count].radius = radius;
        (*segments)[*count].spiral_a = spiral_a;
        (*segments)[*count].offset_x = 0.0;
        (*segments)[*count].offset_y = 0.0;
        
        (*count)++;
    }
    
    fclose(file);
    return (*count > 0) ? 0 : -1;
}

int save_curve_file_json(const char* filepath, const CurveSegment* segments, size_t count) {
    if (filepath == NULL || segments == NULL || count == 0) {
        return -1;
    }
    
    FILE* file = fopen(filepath, "w");
    if (file == NULL) {
        return -1;
    }
    
    fprintf(file, "{\n");
    fprintf(file, "  \"version\": \"1.0\",\n");
    fprintf(file, "  \"segments\": [\n");
    
    for (size_t i = 0; i < count; i++) {
        const char* type_str;
        switch (segments[i].curve_type) {
            case CURVE_TYPE_LINE: type_str = "LINE"; break;
            case CURVE_TYPE_CIRCLE: type_str = "CIRCLE"; break;
            case CURVE_TYPE_SPIRAL: type_str = "SPIRAL"; break;
            case CURVE_TYPE_COMPOUND: type_str = "COMPOUND"; break;
            default: type_str = "LINE";
        }
        
        fprintf(file, "    {\n");
        fprintf(file, "      \"curve_type\": \"%s\",\n", type_str);
        fprintf(file, "      \"station\": %.6f,\n", segments[i].station);
        fprintf(file, "      \"x_coord\": %.6f,\n", segments[i].x_coord);
        fprintf(file, "      \"y_coord\": %.6f,\n", segments[i].y_coord);
        fprintf(file, "      \"azimuth\": %.6f,\n", segments[i].azimuth);
        fprintf(file, "      \"radius\": %.6f,\n", segments[i].radius);
        fprintf(file, "      \"spiral_a\": %.6f\n", segments[i].spiral_a);
        fprintf(file, "    }%s\n", (i < count - 1) ? "," : "");
    }
    
    fprintf(file, "  ]\n");
    fprintf(file, "}\n");
    
    fclose(file);
    return 0;
}

int save_curve_file_vb6(const char* filepath, const CurveSegment* segments, size_t count) {
    if (filepath == NULL || segments == NULL || count == 0) {
        return -1;
    }
    
    FILE* file = fopen(filepath, "w");
    if (file == NULL) {
        return -1;
    }
    
    for (size_t i = 0; i < count; i++) {
        const char* type_str;
        switch (segments[i].curve_type) {
            case CURVE_TYPE_LINE: type_str = "直线"; break;
            case CURVE_TYPE_CIRCLE: type_str = "圆曲线"; break;
            case CURVE_TYPE_SPIRAL: type_str = "缓和曲线"; break;
            case CURVE_TYPE_COMPOUND: type_str = "1+ 圆 +2"; break;
            default: type_str = "直线";
        }
        
        fprintf(file, "%s,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f\n",
                type_str,
                segments[i].station,
                segments[i].x_coord,
                segments[i].y_coord,
                segments[i].azimuth,
                segments[i].radius,
                segments[i].spiral_a);
    }
    
    fclose(file);
    return 0;
}

int export_to_cad_script(const char* filepath, const CurveSegment* segments, 
                         size_t count, int precision) {
    if (filepath == NULL || segments == NULL || count == 0) {
        return -1;
    }
    
    // 初始化计算器
    HorizontalCurveCalculator calc;
    if (horizontal_curve_init(&calc) != 0) {
        return -1;
    }
    
    if (horizontal_curve_load_params(&calc, segments, count) != 0) {
        horizontal_curve_free(&calc);
        return -1;
    }
    
    FILE* file = fopen(filepath, "w");
    if (file == NULL) {
        horizontal_curve_free(&calc);
        return -1;
    }
    
    fprintf(file, "; Highway Coordinate Script\n");
    fprintf(file, "; Generated by C23 Refactored Software\n\n");
    
    const char* fmt = (precision > 0 && precision <= 10) ? 
                      "POINT %.*f,%.*f,0\n" : "POINT %.3f,%.3f,0\n";
    
    // 生成特征点
    for (size_t i = 0; i < count; i++) {
        CalculationResult result = {0};
        if (horizontal_curve_calculate_point(&calc, segments[i].station, &result) == 0) {
            if (precision > 0 && precision <= 10) {
                fprintf(file, fmt, precision, result.x, precision, result.y);
            } else {
                fprintf(file, "POINT %.3f,%.3f,0\n", result.x, result.y);
            }
        }
    }
    
    fclose(file);
    horizontal_curve_free(&calc);
    return 0;
}

int export_to_excel_csv(const char* filepath, const double* stations, 
                        const double (*coordinates)[3], size_t count) {
    if (filepath == NULL || stations == NULL || coordinates == NULL || count == 0) {
        return -1;
    }
    
    FILE* file = fopen(filepath, "w");
    if (file == NULL) {
        return -1;
    }
    
    fprintf(file, "里程，X 坐标，Y 坐标，高程\n");
    
    for (size_t i = 0; i < count; i++) {
        fprintf(file, "%.3f,%.3f,%.3f,%.3f\n",
                stations[i],
                coordinates[i][0],
                coordinates[i][1],
                coordinates[i][2]);
    }
    
    fclose(file);
    return 0;
}

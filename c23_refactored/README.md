# 高速公路坐标高程计算软件 - C23 重构版

## 项目概述

本项目是原 VB6"高速公路坐标高程计算软件 4.3"的纯 C 语言现代化重构，采用最新的 **C23 标准**实现。

原项目是一个用于高速公路工程中坐标和高程计算的 VB6 应用程序，包含平曲线计算、竖曲线计算、坐标转换、CAD 脚本生成等功能。

## 重构目标

1. **语言迁移**: VB6 → C23 (纯 C 原生实现)
2. **架构优化**: 过程式 → 模块化设计
3. **代码质量**: 提高可读性、可维护性和可扩展性
4. **功能保持**: 保留所有核心计算功能
5. **性能优化**: 利用 C 语言的高性能特性

## 核心功能模块

### 1. 平曲线计算 (Horizontal Curve)
- 线元法计算
- 缓和曲线计算
- 圆曲线计算
- 直线段计算
- 方位角计算
- 交点法计算
- 坐标正算/反算

### 2. 竖曲线计算 (Vertical Curve)
- 竖曲线要素计算
- 变坡点处理
- 高程计算

### 3. 宽度渐变计算
- 横断面宽度计算
- 线性内插

### 4. 坐标转换
- 度分秒 ↔ 十进制度
- 极坐标 ↔ 直角坐标
- 已知坐标反算里程偏距
- 已知里程偏距正算坐标

### 5. 数据导入导出
- 支持 VB6 原格式文件
- JSON 格式导出
- CAD 脚本生成
- Excel CSV 导出

## 目录结构

```
c23_refactored/
├── README.md           # 项目说明
├── Makefile            # 构建配置
├── include/            # 头文件
│   ├── curve_types.h          # 数据类型定义
│   ├── angle_utils.h          # 角度工具函数
│   ├── coordinate_calculator.h # 坐标计算器
│   ├── horizontal_curve.h     # 平曲线计算
│   ├── vertical_curve.h       # 竖曲线计算
│   └── file_utils.h           # 文件操作
├── src/                # 源代码
│   ├── main.c                 # 主程序入口
│   ├── angle_utils.c          # 角度工具实现
│   ├── coordinate_calculator.c # 坐标计算器实现
│   ├── horizontal_curve.c     # 平曲线计算实现
│   ├── vertical_curve.c       # 竖曲线计算实现
│   └── file_utils.c           # 文件操作实现
├── tests/              # 测试代码
├── examples/           # 示例代码
└── docs/               # 文档
```

## 快速开始

### 环境要求

- GCC 13+ (支持 C23 标准) 或 Clang 16+
- Make (可选，用于自动化构建)
- 标准 C 库 (glibc/musl)

### 编译项目

```bash
cd c23_refactored

# 使用 Make
make

# 或直接使用 GCC
gcc -std=c23 -Wall -Wextra -O2 -I./include -o highway_calc src/*.c -lm
```

### 运行演示

```bash
./bin/highway_calc
```

### 示例输出

```
************************************************************
   高速公路坐标高程计算软件 - C23 重构版
   Highway Coordinate & Elevation Calculation Software
************************************************************

============================================================
角度转换演示
============================================================

度。分秒格式    →       十进制度
----------------------------------------
153.240524      →       153.401456°     (验证：153.240524)
...

============================================================
平曲线计算演示
============================================================

里程            X 坐标          Y 坐标          方位角
------------------------------------------------------------
K1000           50000.000       50000.000       153.40°
K1100           49750.260       50113.758       7.16°
...
```

## 主要改进

### 1. 代码结构对比

| 特性 | 原 VB6 代码 | C23 重构版 |
|------|-----------|-------------|
| 架构 | Form + Module 混合 | 模块化分层架构 |
| 行数 | ~15,000 行 | ~1,500 行 (核心) |
| 类型安全 | Variant 类型 | 强类型系统 |
| 内存管理 | 自动 (GC) | 手动 (malloc/free) |
| 性能 | 较低 | 高性能 |

### 2. C23 新特性应用

- `[[nodiscard]]` 属性：标记必须检查返回值的函数
- `constexpr`：编译时常量表达式
- 增强的类型安全和警告
- 现代化的代码风格

### 3. 命名规范

- **原代码**: 中英文混用 (`factor`, `置顶`, `GetHandle`)
- **重构后**: 统一英文命名，遵循 snake_case 约定

### 4. 错误处理

- **原代码**: 几乎无错误处理
- **重构后**: 完善的返回值检查和错误码机制

### 5. 文档

- **原代码**: 注释较少
- **重构后**: 完整的 Doxygen 风格文档注释

## API 使用示例

### 平曲线计算

```c
#include "horizontal_curve.h"
#include "curve_types.h"

// 创建曲线参数
CurveSegment segments[3] = {
    {
        .curve_type = CURVE_TYPE_SPIRAL,
        .station = 1000.0,
        .x_coord = 50000.0,
        .y_coord = 50000.0,
        .azimuth = 153.240524,
        .radius = -800.0,
        .spiral_a = 70.0
    }
};

// 初始化计算器
HorizontalCurveCalculator calc;
horizontal_curve_init(&calc);
horizontal_curve_load_params(&calc, segments, 3);

// 计算指定里程的坐标
CalculationResult result;
if (horizontal_curve_calculate_point(&calc, 1100.0, &result) == 0) {
    printf("X: %f, Y: %f, Azimuth: %f\n", result.x, result.y, result.azimuth);
}

// 释放资源
horizontal_curve_free(&calc);
```

### 竖曲线计算

```c
#include "vertical_curve.h"

// 初始化竖曲线参数
VerticalCurveParams params;
vertical_curve_params_init(&params, 16);
params.radius = 10000.0;

// 添加变坡点
vertical_curve_add_grade_point(&params, 1000.0, 100.0);
vertical_curve_add_grade_point(&params, 1500.0, 105.0);

// 计算高程
CalculationResult result;
if (vertical_curve_calculate_elevation(1250.0, &params, &result) == 0) {
    printf("Elevation: %f\n", result.elevation);
}

vertical_curve_params_free(&params);
```

### 角度转换

```c
#include "angle_utils.h"

double dms = 153.240524;
double decimal = dms_to_decimal(dms);  // 153.401456
double back = decimal_to_dms(decimal);  // 153.240524
```

## 开发状态

- [x] 核心数据类型定义
- [x] 角度转换工具
- [x] 坐标计算器
- [x] 平曲线计算模块
- [x] 竖曲线计算模块
- [x] 文件操作工具
- [x] 主程序演示
- [ ] 单元测试
- [ ] 完整文档
- [ ] GUI 界面 (可选)

## 与原 VB6 版本的兼容性

重构版本保持了与原版相同的计算逻辑:
- 度分秒格式的角度表示
- 左偏负半径、右偏正半径的约定
- 线元法计算公式
- 逐步趋近法的坐标反算

## 构建选项

```bash
# 调试模式
make CFLAGS="-std=c23 -Wall -Wextra -g -O0"

# 发布模式 (优化)
make CFLAGS="-std=c23 -Wall -Wextra -O3 -DNDEBUG"

# 清理
make clean

# 运行测试
make test
```

## 许可证

继承原项目许可证

## 联系方式

原作者信息请参考原项目文件。

## 参考资料

- ISO/IEC 9899:2023 (C23 标准)
- GCC C23 支持文档
- 原 VB6 项目源码

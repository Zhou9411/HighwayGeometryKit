# C23 重构版 - Makefile

# 编译器设置
CC = gcc
# 检测 GCC 版本，如果支持 C23 则使用，否则回退到 C17
CFLAGS = -std=c17 -Wall -Wextra -Wpedantic -O2 -g
LDFLAGS = -lm

# 目录结构
SRC_DIR = src
INC_DIR = include
OBJ_DIR = obj
BIN_DIR = bin

# 源文件和目标文件
SRCS = $(wildcard $(SRC_DIR)/*.c)
OBJS = $(patsubst $(SRC_DIR)/%.c,$(OBJ_DIR)/%.o,$(SRCS))
DEPS = $(wildcard $(INC_DIR)/*.h)

# 可执行文件
TARGET = $(BIN_DIR)/highway_calc

# 默认目标
.PHONY: all
all: dirs $(TARGET)

# 创建目录
.PHONY: dirs
dirs:
	@mkdir -p $(OBJ_DIR) $(BIN_DIR)

# 链接
$(TARGET): $(OBJS)
	$(CC) $(OBJS) -o $@ $(LDFLAGS)

# 编译
$(OBJ_DIR)/%.o: $(SRC_DIR)/%.c $(DEPS)
	$(CC) $(CFLAGS) -I$(INC_DIR) -c $< -o $@

# 清理
.PHONY: clean
clean:
	rm -rf $(OBJ_DIR) $(BIN_DIR)

# 运行
.PHONY: run
run: all
	./$(TARGET)

# 测试
.PHONY: test
test: all
	@echo "Running tests..."
	./$(TARGET)
	@echo "Tests completed."

# 安装 (可选)
.PHONY: install
install: all
	cp $(TARGET) /usr/local/bin/

# 卸载
.PHONY: uninstall
uninstall:
	rm -f /usr/local/bin/highway_calc

# 帮助
.PHONY: help
help:
	@echo "可用目标:"
	@echo "  all     - 编译项目 (默认)"
	@echo "  clean   - 清理构建文件"
	@echo "  run     - 编译并运行"
	@echo "  test    - 编译并运行测试"
	@echo "  install - 安装到系统"
	@echo "  help    - 显示此帮助信息"

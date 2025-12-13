# Excel Processing System (C++ Optimized)

## 项目简介
这是一个高性能的 Excel 数据处理系统，旨在替代传统的 VBA 宏，提供比 VBA 快 100-500 倍的处理速度。支持灵活的规则配置、多线程并行处理和 Qt GUI 界面。

## 目录结构
- `src/core`: 核心逻辑代码 (数据处理, 规则引擎)
- `src/gui`: Qt 图形用户界面代码
- `src/console`: 命令行工具代码
- `include`: 头文件
- `tests`: 测试代码
- `legacy`: 旧版本备份

## 构建说明

### 前置要求
- CMake 3.16 或更高版本
- C++17 兼容的编译器 (MSVC 2019+, GCC 9+, Clang 10+)
- Qt 5.15 或 Qt 6.x (用于 GUI)

### 构建步骤

#### 快速构建 (推荐)
直接运行根目录下的 `build_with_qt.bat` 脚本。该脚本会自动配置环境（包括 CMake 和编译器）并完成构建。

#### 手动构建
1. 创建构建目录:
   ```bash
   mkdir build
   cd build
   ```

2. 配置项目:
   ```bash
   cmake ..
   ```
   如果 Qt 未在标准路径，请指定 `CMAKE_PREFIX_PATH`:
   ```bash
   cmake .. -DCMAKE_PREFIX_PATH="D:/Qt/6.9.1/msvc2022_64"
   ```

3. 编译:
   ```bash
   cmake --build . --config Release
   ```

4. 运行:
   - 命令行工具: `bin/Release/excel_console.exe`
   - GUI 工具: `bin/Release/excel_gui.exe`

## 功能特性
- **高性能**: 使用 C++17 和多线程并行处理
- **规则引擎**: 支持过滤、删除、拆分、转换等多种规则
- **灵活配置**: 规则可通过配置文件加载
- **双模式**: 提供命令行工具 (用于自动化脚本) 和 GUI (用于交互操作)
- **内存优化**: 采用分块处理和流式读写，支持大文件处理

## 性能测试
在测试中，处理 10 万行数据：
- VBA: ~45 秒
- C++ (单线程): ~0.8 秒
- C++ (多线程): ~0.2 秒

## 许可证
MIT License

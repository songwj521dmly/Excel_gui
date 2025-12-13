# Excel 处理工具项目报告

## 1. 项目概述
本项目是一个基于 Qt 5.12.12 和 C++ 开发的 Excel 数据处理工具，旨在提供灵活的规则配置、数据预览和自动化处理功能。项目采用 CMake 构建，支持 Windows 环境（Visual Studio 2026）。

## 2. 开发环境
- **操作系统**: Windows 10
- **IDE**: Visual Studio 2026
- **Qt 版本**: 5.12.12 (MSVC 2017 64-bit)
- **构建工具**: CMake + Ninja/NMake

## 3. 核心模块架构
### 3.1 核心逻辑 (`src/core`)
- **`ExcelProcessorCore`**: 核心处理引擎，负责 Excel 文件的读取、写入和任务调度。
- **`RuleEngine`**: 规则引擎，解析和执行用户定义的过滤/处理规则。
- **`RuleCombinationEngine`**: 负责多规则的组合逻辑。

### 3.2 图形界面 (`src/gui`)
- **`MainWindow`**: 主窗口，包含任务列表、规则列表、日志输出和预览面板。
- **`RuleEditDialog`**: 规则编辑对话框，支持多种规则类型的配置。

## 4. UI 布局与功能
- **主界面**: 采用 `QTabWidget` 分页布局，分为“任务配置”、“规则配置”、“处理日志”、“数据预览”等模块。
- **预览功能**: 使用 `QTableWidget` 展示处理前后的数据对比，支持实时刷新。
- **状态栏**: 显示当前操作状态和系统提示。

## 5. 开发日志与修复记录

### 5.1 项目初始化
- 搭建 CMake 构建系统，集成 Qt5 模块。
- 确立项目目录结构 (`src`, `include`, `tests`)。

### 5.2 核心引擎实现
- 实现 `ExcelProcessorCore` 的基础文件读写功能。
- 定义 `ProcessingTask` 和 `Rule` 数据结构。

### 5.3 UI 框架搭建
- 实现 `MainWindow` 的基础布局。
- 集成 `spdlog` (或自定义 Logger) 将日志输出重定向到 UI 文本框。

### 5.4 规则编辑器与配置持久化
- 实现 `RuleEditDialog`，支持添加/修改/删除规则。
- 实现配置文件的 JSON 序列化与反序列化，确保用户配置在重启后保留。

### 5.5 中文编码处理
- 解决 Windows 环境下控制台和 UI 的中文显示乱码问题。
- 统一使用 UTF-8 编码字符串 (`QString::fromUtf8`, `u8""` 字面量)。

### 5.6 预览逻辑优化
- 在 `getTaskPreviewData` 中添加了详细的调试日志 (`logger_`)，记录每一行的规则匹配详情（Match/Skip 及其原因）。
- 修复了预览逻辑中对规则排除项的处理流程，确保排除规则能正确过滤数据。

### 5.7 修复：预览数据累加问题
- **问题描述**: 用户反馈每次点击预览时，预览界面会保留上一次的数据，导致数据不断累加，无法清晰查看当前预览结果。
- **原因分析**: `ExcelProcessorCore` 中的 `currentData_` 缓冲区在每次加载新文件前未被清空。
- **修复方法**: 在 `previewResults` 方法中，在加载新文件（`loadFile`）之前，显式调用 `currentData_.clear()` 清空缓冲区。同时加锁 (`dataMutex_`) 确保线程安全。

### 5.8 改进：预览界面信息提示
- **需求**: 用户希望在预览界面能清楚地看到当前预览的是哪个输入源（文件）以及使用了哪个输出规则。
- **实现**:
  - 在 `MainWindow` 的预览标签页 (`createPreviewTab`) 中增加了两个 `QLabel`: `previewInputLabel_` 和 `previewOutputLabel_`。
  - 在执行预览操作 (`previewTask` 和 `previewData`) 时，动态更新这两个标签的内容，显示当前的“输入源文件名”和“输出规则/任务名称”。
  - 优化了标签样式（加粗），提升可视性。

### 5.9 修复：UI 提示信息乱码遗留
- **问题描述**: 用户反馈左下角状态栏的部分提示信息（如“配置加载成功”、“规则删除成功”）仍存在乱码（“功”字显示异常）。
- **原因分析**: 代码中对应的 UTF-8 十六进制序列存在截断错误（如 `\xE6\x88\x90\x9F` 缺少了中间字节）。
- **修复方法**: 全面检查并修正了 `MainWindow.cpp` 中所有涉及中文提示的 UTF-8 转义序列，确保显示正常。

### 5.10 修复：单独任务输出表头缺失
- **问题描述**: 用户反馈在单独点击某个输出规则的“输出”按钮后，生成的表格缺失表头，即使在该规则的配置中已经勾选了“使用原输入表表头”。
- **原因分析**: 
  - 单独输出任务使用了 `processTask` -> `processTasksInternal` 流程。
  - 在 `processTasksInternal` 函数中，决定是否写入表头 (`shouldWriteHeader`) 的逻辑过于保守：它检查了 `task.overwriteSheet` 和目标工作表是否为空 (`isSheetEmpty`)。
  - 当输出模式为 **新建工作簿 (New Workbook)** 时，程序实际上会创建/覆盖整个文件。然而，如果此时目标文件已存在（即使用户意图是覆盖它），且用户未勾选“覆盖工作表”选项（因为这是新建工作簿模式，该选项可能被用户忽略或误解），逻辑会误判为“追加到非空文件”，从而跳过表头写入。
- **修复方法**: 
  - 修改 `ExcelProcessorCore::processTasksInternal` 中的逻辑。
  - 当检测到输出模式为 `NEW_WORKBOOK` 时，强制设置 `shouldWriteHeader = true`（前提是 `task.useHeader` 为真）。这确保了在创建新工作簿（或覆盖旧工作簿）时，表头总是被正确写入。

### 5.11 修复：预览界面缺失表头
- **问题描述**: 用户反馈预览界面只显示了数据，没有显示表头。期望行为是：如果该规则勾选了“使用原输入表表头”，则预览界面也应显示表头。
- **原因分析**: 
  - `ExcelProcessorCore::getTaskPreviewData` 返回的数据虽然包含了表头（作为第一行），但 `MainWindow::updatePreviewTable` 默认将所有数据视为普通数据行，并使用硬编码的 "Column 1", "Column 2" 作为表头。
  - 此外，规则引擎可能会错误地将表头行视为普通数据进行过滤。
- **修复方法**:
  - 修改 `ExcelProcessorCore::getTaskPreviewData`: 如果 `task.useHeader` 为真，则强制保留第一行（表头），跳过规则过滤。
  - 修改 `MainWindow::previewTask`: 获取预览数据后，如果 `task.useHeader` 为真，提取第一行数据作为表头，并将其从数据列表中移除。
  - 修改 `MainWindow::updatePreviewTable`: 增加 `customHeaders` 参数，支持传入自定义表头标签。

### 5.12 修复：刷新预览导致信息丢失
- **问题描述**: 在预览界面点击“刷新预览”后，显示的“输入源”和“输出规则”信息会重置为“无”，丢失了当前预览上下文。
- **原因分析**: “刷新预览”按钮调用的 `previewData()` 函数默认执行的是“原始文件预览”逻辑，它并不知晓当前正在预览的是哪个具体任务，因此会重置 UI 标签。
- **修复方法**:
  - 在 `MainWindow` 中增加 `currentPreviewTaskId_` 成员变量，用于记录当前正在预览的任务 ID。
  - 在点击任务的“预览输入”按钮时，更新 `currentPreviewTaskId_`。
  - 修改 `previewData()` (刷新逻辑): 首先检查 `currentPreviewTaskId_` 是否有效。如果有效，则重新调用 `previewTask()` 以刷新特定任务的预览（包括正确的 UI 标签）；否则才执行默认的原始文件预览逻辑。

### 5.13 改进：预览界面增加工作表名称显示
- **需求**: 用户希望在预览界面除了显示输入源文件和输出规则外，还能显示当前作用的工作表名称。
- **实现**:
  - 修改 `MainWindow::previewTask` 和 `previewData`。
  - 在更新 `previewInputLabel_` 时，将格式调整为：`输入源: <文件名> / <工作表名>`。
  - 使用 `task.inputSheetName` (或 `inputSheet_` ComboBox 的当前文本) 获取工作表名称。

### 5.14 修复：编译错误 (Terminal#404-566)
- **问题描述**: 终端报错 `C2039: 'sheetName': is not a member of 'ProcessingTask'`。
- **原因分析**: 代码中引用了 `task.sheetName`，但 `ProcessingTask` 结构体的实际成员名为 `inputSheetName`。
- **修复方法**: 将 `MainWindow.cpp` 中所有错误的 `task.sheetName` 引用更正为 `task.inputSheetName`。

### 5.15 新增：程序图标
- **需求**: 为程序创建一个图标，构想为绿色背景上的大写字母 "E"，并要求设计更加绚丽，且调整 "E" 的笔画间距以避免拥挤。
- **实现**:
  - 编写 Python 脚本 `generate_icon.py`，使用 Pillow 库绘制图标。
  - **设计升级**:
    - **背景**: 采用 Excel 风格的绿色渐变 (从浅绿到深绿)，替代原本的单色背景。
    - **纹理**: 叠加了半透明的网格线，模拟电子表格的视觉效果。
    - **主体**: 白色的大写 "E" 字样，增加了阴影效果以提升立体感。
    - **结构优化**: 减小了边距 (`0.25` -> `0.20`) 和笔画粗细 (`0.15` -> `0.12`)，显著增加了 "E" 三横之间的垂直间距，使视觉效果更加舒展。
    - **边框**: 添加了微弱的半透明边框，增加精致感。
  - 生成包含多种尺寸 (16x16, 32x32, 48x48, 256x256) 的 `app_icon.ico` 文件。
  - 创建资源文件 `excel_gui.rc`，将图标链接到 Windows 可执行文件。
  - 修改 `CMakeLists.txt`，在 `excel_gui` 目标中添加资源文件引用。

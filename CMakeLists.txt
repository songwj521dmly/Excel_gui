cmake_minimum_required(VERSION 3.16)
project(ExcelProcessor LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

# Output directories
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)

# Include directories
include_directories(${CMAKE_SOURCE_DIR}/include)

# FLTK (FetchContent) - DISABLED per user request
# include(FetchContent)
# set(FLTK_BUILD_TEST OFF CACHE BOOL " " FORCE)
# set(FLTK_BUILD_EXAMPLES OFF CACHE BOOL " " FORCE)
# set(FLTK_BUILD_FLUID OFF CACHE BOOL " " FORCE)
# set(OPTION_USE_GL OFF CACHE BOOL " " FORCE)
# set(OPTION_USE_THREADS ON CACHE BOOL " " FORCE)
# set(OPTION_USE_SYSTEM_LIBPNG OFF CACHE BOOL " " FORCE)
# set(OPTION_USE_SYSTEM_LIBJPEG OFF CACHE BOOL " " FORCE)
# set(OPTION_USE_SYSTEM_ZLIB OFF CACHE BOOL " " FORCE)

# FetchContent_Declare(
#   fltk
#   URL https://codeload.github.com/fltk/fltk/zip/refs/tags/release-1.3.9
#   DOWNLOAD_EXTRACT_TIMESTAMP TRUE
# )
# FetchContent_MakeAvailable(fltk)

# Find Qt (Try Qt6, fallback to Qt5)
option(OPTION_FORCE_QT5 "Force usage of Qt5" OFF)

if(NOT OPTION_FORCE_QT5)
    find_package(Qt6 COMPONENTS Core Gui Widgets Concurrent AxContainer QUIET)
endif()

if(Qt6_FOUND)
    message(STATUS "Found Qt6: ${Qt6_VERSION}")
    set(QT_LIB_PREFIX Qt6)
else()
    find_package(Qt5 COMPONENTS Core Gui Widgets Concurrent AxContainer REQUIRED)
    message(STATUS "Found Qt5: ${Qt5_VERSION}")
    set(QT_LIB_PREFIX Qt5)
endif()

# Export Qt Bin Directory for deployment script
if(QT_LIB_PREFIX STREQUAL "Qt6" AND TARGET Qt6::qmake)
    get_target_property(QT_QMAKE_EXECUTABLE Qt6::qmake IMPORTED_LOCATION)
    get_filename_component(QT_BIN_DIR "${QT_QMAKE_EXECUTABLE}" DIRECTORY)
elseif(QT_LIB_PREFIX STREQUAL "Qt5" AND TARGET Qt5::qmake)
    get_target_property(QT_QMAKE_EXECUTABLE Qt5::qmake IMPORTED_LOCATION)
    get_filename_component(QT_BIN_DIR "${QT_QMAKE_EXECUTABLE}" DIRECTORY)
else()
    message(WARNING "Could not determine Qt bin directory from targets.")
endif()

if(QT_BIN_DIR)
    file(WRITE "${CMAKE_BINARY_DIR}/qt_bin_dir.txt" "${QT_BIN_DIR}")
    message(STATUS "Qt Bin Dir exported to: ${CMAKE_BINARY_DIR}/qt_bin_dir.txt")
endif()

# Core Library
add_library(ExcelProcessorCore
    src/core/ExcelProcessorCore.cpp
    src/core/RuleEngine.cpp
    src/core/RuleCombinationEngine.cpp
)
target_link_libraries(ExcelProcessorCore PRIVATE ${QT_LIB_PREFIX}::Core ${QT_LIB_PREFIX}::Concurrent ${QT_LIB_PREFIX}::AxContainer)

# Console Application
add_executable(excel_console
    src/console/main.cpp
)
target_link_libraries(excel_console PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core)

# GUI Application
add_executable(excel_gui
    src/gui/main.cpp
    src/gui/MainWindow.cpp
    src/gui/MainWindow.h
    src/gui/RuleEditDialog.cpp
    src/gui/RuleEditDialog.h
)
if(WIN32)
    target_sources(excel_gui PRIVATE excel_gui.rc)
endif()

target_include_directories(excel_gui BEFORE PRIVATE ${CMAKE_SOURCE_DIR}/include)
target_link_libraries(excel_gui PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core ${QT_LIB_PREFIX}::Gui ${QT_LIB_PREFIX}::Widgets)

# Handle Window application on Windows (avoid console window for GUI)
if(WIN32)
    set_target_properties(excel_gui PROPERTIES WIN32_EXECUTABLE ON)
endif()

# Tests - DISABLED per user request
# enable_testing()
# add_executable(test_config_save_load
#     tests/test_config_save_load.cpp
# )
# target_link_libraries(test_config_save_load PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core)
# add_test(NAME test_config_save_load COMMAND test_config_save_load)
#
# add_executable(test_rule_engine
#     tests/test_rule_engine.cpp
# )
# target_link_libraries(test_rule_engine PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core)
# add_test(NAME test_rule_engine COMMAND test_rule_engine)
#
# # add_executable(test_rule_exclusion
# #     tests/test_rule_exclusion.cpp
# # )
# # target_link_libraries(test_rule_exclusion PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core)
# # add_test(NAME test_rule_exclusion COMMAND test_rule_exclusion)
#
# add_executable(test_config_persistence
#     tests/test_config_persistence.cpp
# )
# target_link_libraries(test_config_persistence PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core)
# add_test(NAME test_config_persistence COMMAND test_config_persistence)
#
# add_executable(test_granular_exclusion
#     tests/test_granular_exclusion.cpp
# )
# target_link_libraries(test_granular_exclusion PRIVATE ExcelProcessorCore ${QT_LIB_PREFIX}::Core)
# add_test(NAME test_granular_exclusion COMMAND test_granular_exclusion)

# FLTK GUI Application - DISABLED per user request
# add_executable(excel_fltk
#     src/fltk_gui/main.cpp
#     src/fltk_gui/FltkMainWindow.cpp
#     src/fltk_gui/FltkMainWindow.h
#     src/fltk_gui/RuleTable.cpp
#     src/fltk_gui/RuleTable.h
#     src/fltk_gui/EditRuleDialog.cpp
#     src/fltk_gui/EditRuleDialog.h
#     src/fltk_gui/PreviewTable.cpp
#     src/fltk_gui/PreviewTable.h
#     src/fltk_gui/TaskTable.cpp
#     src/fltk_gui/TaskTable.h
#     src/fltk_gui/EditTaskDialog.cpp
#     src/fltk_gui/EditTaskDialog.h
#     src/fltk_gui/ExclusionDialog.cpp
#     src/fltk_gui/ExclusionDialog.h
# )
# target_include_directories(excel_fltk PRIVATE ${CMAKE_SOURCE_DIR}/include ${fltk_SOURCE_DIR} ${fltk_BINARY_DIR})
# target_link_libraries(excel_fltk PRIVATE ExcelProcessorCore fltk fltk_images)
#
# # Handle Window application on Windows
# if(WIN32)
# #    set_target_properties(excel_fltk PROPERTIES WIN32_EXECUTABLE ON)
# endif()

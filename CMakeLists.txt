cmake_minimum_required(VERSION 3.16)
project(pptx_reader VERSION 0.1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE Release)
endif()

# ---------- 检测操作系统 ----------
if(WIN32)
    set(WINDOWS_BUILD TRUE)
else()
    set(WINDOWS_BUILD FALSE)
endif()

# ---------- 依赖：pugixml（git 子模块，位于 ./pugixml） ----------
set(PUGIXML_SRC_DIR "${CMAKE_CURRENT_SOURCE_DIR}/pugixml/src")
if(NOT EXISTS "${PUGIXML_SRC_DIR}/pugixml.hpp")
    message(FATAL_ERROR
        "未找到 pugixml。请先初始化子模块:\n"
        "  git submodule update --init --recursive\n"
        "期望路径: ${PUGIXML_SRC_DIR}/pugixml.hpp")
endif()
add_library(pugixml STATIC "${PUGIXML_SRC_DIR}/pugixml.cpp")
target_include_directories(pugixml PUBLIC "${PUGIXML_SRC_DIR}")
target_compile_options(pugixml PRIVATE -w)

# ---------- 依赖：libzip ----------
# 优先使用 vcpkg（Windows），否则使用 pkg-config（Linux/macOS）
if(WINDOWS_BUILD AND DEFINED ENV{VCPKG_ROOT})
    message(STATUS "使用 vcpkg 查找 libzip")
    find_path(LIBZIP_INCLUDE_DIR
        NAMES zip.h
        PATHS "$ENV{VCPKG_ROOT}/installed/x64-windows/include"
        REQUIRED
    )
    find_library(LIBZIP_LIBRARY
        NAMES zip libzip
        PATHS "$ENV{VCPKG_ROOT}/installed/x64-windows/lib"
        REQUIRED
    )
    
    if(LIBZIP_INCLUDE_DIR AND LIBZIP_LIBRARY)
        message(STATUS "找到 libzip: ${LIBZIP_LIBRARY}")
        add_library(LIBZIP::LIBZIP UNKNOWN IMPORTED)
        set_target_properties(LIBZIP::LIBZIP PROPERTIES
            IMPORTED_LOCATION "${LIBZIP_LIBRARY}"
            INTERFACE_INCLUDE_DIRECTORIES "${LIBZIP_INCLUDE_DIR}"
        )
    else()
        message(FATAL_ERROR "未找到 libzip，请确保 vcpkg 已安装 libzip: vcpkg install libzip")
    endif()
else()
    # Linux/macOS 使用 pkg-config
    message(STATUS "使用 pkg-config 查找 libzip")
    find_package(PkgConfig REQUIRED)
    pkg_check_modules(LIBZIP REQUIRED IMPORTED_TARGET libzip)
endif()

# ---------- pptx_reader 静态库 ----------
add_library(pptx_reader STATIC
    src/PptxArchive.cpp
    src/Paragraph.cpp
    src/Shape.cpp
    src/Slide.cpp
    src/Presentation.cpp
)
target_include_directories(pptx_reader PUBLIC
    "${CMAKE_CURRENT_SOURCE_DIR}/include"
)

# 链接依赖库
if(WINDOWS_BUILD AND DEFINED ENV{VCPKG_ROOT})
    target_link_libraries(pptx_reader PUBLIC pugixml LIBZIP::LIBZIP)
else()
    target_link_libraries(pptx_reader PUBLIC pugixml PkgConfig::LIBZIP)
endif()

target_compile_options(pptx_reader PRIVATE -Wall -Wextra)

# Windows 特定编译选项
if(WIN32)
    target_compile_options(pptx_reader PRIVATE /W4)
    # 禁用一些警告
    target_compile_options(pptx_reader PRIVATE /wd4251 /wd4275)
endif()

# ---------- 示例程序 ----------
add_executable(read_pptx examples/read_pptx.cpp)
target_link_libraries(read_pptx PRIVATE pptx_reader)

# ---------- 安装规则（可选） ----------
install(TARGETS pptx_reader
    ARCHIVE DESTINATION lib
    LIBRARY DESTINATION lib
    RUNTIME DESTINATION bin
)
install(DIRECTORY include/ DESTINATION include)
cmake_minimum_required(VERSION 3.20)
project(FilePropsTool LANGUAGES C CXX RC)

set(CMAKE_MFC_FLAG 1)                                    # MFC 静态链接 (无 /D_AFXDLL)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")  # /MT 静态 CRT
set(CMAKE_CONFIGURATION_TYPES "Debug;Release" CACHE STRING "" FORCE)

add_executable(FilePropsTool WIN32
    Src/main.cpp
    Src/miniz.c
    Res/_res.rc)
target_compile_definitions(FilePropsTool PRIVATE UNICODE _UNICODE)
target_compile_options(FilePropsTool PRIVATE /utf-8 /EHsc)

# 产物复制到源码根目录 (同旧 build.bat 行为)
add_custom_command(TARGET FilePropsTool POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different "$<TARGET_FILE:FilePropsTool>" "${CMAKE_SOURCE_DIR}/")

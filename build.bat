@echo off
rem FilePropsTool — CMake Preset 构建入口 (VS2019 x64, Release)
cd /d "%~dp0"
cmake --preset vs2019-x64 || exit /b 1
cmake --build --preset release

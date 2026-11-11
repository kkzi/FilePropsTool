# FilePropsTool

向导式的文件属性时间修改工具（Windows / MFC / 静态链接单文件）。

## 功能

- **第 1 屏**：拖放或点选文件；最近使用记录（单条 ✕ 删除、点击置顶，最多 5 条，存于 `%LOCALAPPDATA%\FilePropsTool\recent.txt`）
- **第 2 屏**：文件系统时间——创建 / 修改 / 访问（`SetFileTime`）
- **第 3 屏**：Office 文档属性（OpenXML：docx/xlsx/pptx 等）——作者、上次保存者、修订号、版本号、公司、经理、创建 / 修改 / 打印时间。直接重写 zip 内 `docProps/core.xml` / `app.xml`，非 OpenXML 文件自动跳过本屏
- XML 内部存 UTC，界面显示/输入本地时间；中文内容按 UTF-8 处理
- 保存后回到第 1 屏，可继续处理下一个文件

## 构建

VS2019 x64 + CMake Preset（静态 MFC、静态 CRT，产物无运行库依赖）：

```
cmake --preset vs2019-x64
cmake --build --preset release
```

或直接 `build.bat`。构建后 exe 自动复制到仓库根目录。

## 结构

```
Src/  main.cpp miniz.c miniz.h   源码（zip 重写用 miniz）
Res/  _res.rc smile.ico manifest 图标与资源
```

# FbdToPpt 发布编译目录

这个目录只包含编译桌面程序所需的源码、书籍适配器、内置 PPTX 模板和 PyInstaller 配置。

## 目录

- `tools/`：GUI、FBD 转 Markdown、公共资源和书籍适配器
- `tests/`：当前 GUI 动态调用的书籍生成模块；这里的文件名是历史兼容名称
- `static/templates/`：内置 PPTX 模板
- `packaging/fbd_to_ppt.spec`：PyInstaller 配置

原始 `.fbd`、`.tif`、历史 PPTX、调试目录和生成结果不会复制到这里。用户运行程序时可以从任意位置选择自己的 FBD 和模板文件。

## Windows 编译

请在 Windows 的 Anaconda/Miniconda 环境中打开 PowerShell，并进入本目录：

```powershell
conda activate <你的环境名>
python -m pip install -r requirements-release.txt
python -m PyInstaller --clean --noconfirm packaging/fbd_to_ppt.spec
```

输出目录为 `dist/FbdToPpt/`，其中的 `FbdToPpt.exe` 和 `static/templates/` 需要一起发布。

## macOS 编译

macOS 应在 macOS 本机的 Conda 环境中编译，不能在 Windows 或 WSL 中直接生成原生 macOS `.app`：

```bash
conda activate <你的环境名>
python -m pip install -r requirements-release.txt
python -m PyInstaller --clean --noconfirm packaging/fbd_to_ppt.spec
```

WSL 只能生成 Linux 版本，不能替代 Windows 编译 EXE 或 macOS 编译 APP。

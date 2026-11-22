# 肥鸭文档转换

**把拍歪的表格照片，变成能编辑的 Excel。**

拍一张表格照片，或者把 PDF 丢进来，几秒钟后得到一份保留行列结构的 Excel。全程在本机识别，图片和结果不上传云端。

官网：[feifei3yaya.github.io/fatduck-ocr](https://feifei3yaya.github.io/fatduck-ocr/) · 下载：[最新版本](https://github.com/feifei3yaya/fatduck-ocr/releases/latest)

![界面预览](docs/screenshot.png)

## 下载

| 版本 | 适合谁 | 下载 |
| --- | --- | --- |
| 安装包（推荐） | 大多数 Windows 用户，装完即可在开始菜单找到 | [下载安装包](https://github.com/feifei3yaya/fatduck-ocr/releases/latest) |
| 便携包 | 不想安装、放 U 盘里用 | [下载便携包](https://github.com/feifei3yaya/fatduck-ocr/releases/latest) |
| 源码 | 想自己改、自己构建 | [下载源码包](https://github.com/feifei3yaya/fatduck-ocr/releases/latest) |

- 系统要求：Windows 10 / 11（64 位）。macOS 版本在规划中。
- 免费、开源、无广告、无需注册账号。
- 首次运行如遇 Windows SmartScreen 蓝色提示，点「更多信息 → 仍要运行」即可；也可先运行随包附带的「解除锁定.bat」。原因是我们没有购买代码签名证书，详见 [使用说明_SmartScreen.txt](使用说明_SmartScreen.txt)。

## 三步上手

1. 打开肥鸭文档转换，把图片或 PDF 拖进窗口（也可以点「选择文件」或直接 Ctrl+V 粘贴截图）。
2. 右侧选择「导出模式」，一般保持默认的「图片转 Excel」。
3. 点「开始转换」，完成后点「打开目录」，Excel 就在那里。

识别过程不需要联网，不需要登录，也不需要把文件传到任何服务器。启动时的版本检查只查询版本号，可在设置中关闭。

## 能做什么

- **图片转 Excel**：识别照片、扫描件、截图里的文字和表格，按原版式逐行逐列还原；一页写成一个工作表。
- **PDF 转 Excel**：支持多页 PDF，逐页识别，不丢页。
- **查找单号**：粘贴完整或部分单号，在多张图片里找到它并定位到原图。
- **单号清单导出**：只汇总单号一列，适合批量核对。
- **待复核**：把识别没把握的单元格集中列出来，点一下就能放大看原图、对照复识结果，标记核对或重拍。
- **最近任务**：任务摘要和待复核快照存在本机，关掉程序再打开还能接着核对。
- **拍摄质量提示**：模糊、倾斜、反光会给出提示，帮你补拍。

支持格式：PNG / JPG / JPEG / BMP / TIF / TIFF / WEBP / PDF。单文件上限 100 MB，可一次拖入整个文件夹。

适合谁：需要把纸质表格、聊天截图、扫描 PDF 转成 Excel 的普通用户；需要批量核对单号的行政、财务、仓储、物流岗位；不想把资料上传到在线转换网站的人。

## 隐私

- 默认完全离线，OCR 在本机 CPU 上运行。
- 不收集图片、文件名、识别结果，不需要注册，没有统计埋点。
- 唯一的可选联网功能是「云端文本纠错」，默认关闭，需你在设置里主动开启；开启后只发送存疑单元格的局部裁剪图和识别文字，不发送整张图片。该功能需要你自己的 NVIDIA API Key，普通用户保持关闭即可。
- 「版本检查」默认开启，只向 GitHub 查询最新版本号，可在设置中关闭。
- 任务历史、配置、日志只写本机 `%APPDATA%` / `%LOCALAPPDATA%` 下的程序目录。

完整说明见 [privacy_notice.txt](privacy_notice.txt)。

## 常见问题

**会收费吗？**
不会。MIT 开源，免费使用，也没有计划做付费版。

**需要联网吗？**
识别不需要。只有你主动开启「云端文本纠错」或保持「版本检查」开启时才会联网。

**为什么 Windows 会弹蓝色警告？**
因为安装包没有代码签名证书。点「更多信息 → 仍要运行」即可，也可以先运行随包的「解除锁定.bat」。

**能识别手写体吗？**
对印刷体效果最好；工整手写可以尝试，潦草手写准确率有限。

**支持 Mac 吗？**
目前只支持 Windows 10/11 64 位，macOS 版本在规划中。

**识别不准怎么办？**
先看「待复核」里的提示。拍摄时尽量正对表格、光线均匀、避免反光，表格占满画面效果更好。

**会保存我的图片吗？**
转换用的临时文件在任务结束后清理；任务历史只保存摘要和缩略图，可在设置里清除。

更多问题见 [docs/FAQ.md](docs/FAQ.md)。

## 从源码运行

1. 安装 Python 3.12（64 位）。
2. 安装依赖：

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. 准备 OCR 运行时 `runtime\` 与模型 `models\PP-OCRv6\`（体积较大，未提交到仓库）：
   - `runtime\`：一个包含 `openvino`、`opencv-python`、`numpy`、`pypdfium2` 的 Python 环境。
     可在本机已有 OpenVINO venv 的基础上运行 `python build_runtime.py` 裁剪生成。
   - `models\PP-OCRv6\`：PaddleOCR 检测 / 识别模型转换为 OpenVINO IR 后的
     `det\*.xml|bin` 与 `rec\*.xml|bin`，可用 `python tools/ppocrv6_setup.py` 转换。
   - `assets\ppocrv6_dict.txt`、`assets\ppocrv5_dict.txt`：识别字库。

4. 启动：

   ```powershell
   python desktop.py
   ```

   浏览器调试：

   ```powershell
   python web_app.py
   ```

## 构建

```powershell
python build_exe.py        # dist\肥鸭文档转换.exe
python build_package.py    # dist\肥鸭文档转换_便携包_vX.Y.Z.zip
ISCC.exe installer.iss     # dist\肥鸭文档转换_安装包_vX.Y.Z.exe
```

构建前需要准备 `runtime\`、`models\PP-OCRv6\` 与两份字库，见上一节。

## 目录结构

| 路径 | 说明 |
| --- | --- |
| `desktop.py` | 桌面入口：本地服务 + WebView2 窗口 + 托盘 |
| `web_app.py` | 本地 FastAPI 服务与转换任务编排 |
| `engine.py` / `ocr_worker.py` | OCR 引擎与子进程工作器 |
| `md2excel.py` / `table_rebuild.py` | 表格重建与 Excel 导出 |
| `ai_calibrate.py` / `xcheck.py` / `nv_fix.py` | 规则校准、交叉校验、可选云端纠错 |
| `job_store.py` | 本机任务历史 |
| `core.py` | 程序名、隐私文案、配置与运行环境定位 |
| `web/` | 前端界面（HTML / CSS / JS） |
| `assets/` | 图标、吉祥物、字库 |
| `docs/` | 截图与 FAQ |
| `models/` | OCR 模型（不提交） |
| `runtime/` | OCR 运行时（不提交） |
| `build_exe.py` / `build_package.py` / `build_runtime.py` | 构建脚本 |
| `installer.iss` | Inno Setup 安装包脚本 |

## 参与贡献

欢迎提 Issue 和 PR。提交前请先看 [CONTRIBUTING.md](CONTRIBUTING.md)；安全问题请按 [SECURITY.md](SECURITY.md) 私下反馈，不要公开贴敏感图片或单号。

## 第三方组件

- [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) PP-OCRv6 模型（Apache-2.0）
- [OpenVINO](https://github.com/openvinotoolkit/openvino)（Apache-2.0）
- [FastAPI](https://github.com/fastapi/fastapi)（MIT）、[Uvicorn](https://github.com/encode/uvicorn)（BSD-3-Clause）
- [pywebview](https://github.com/r0x0r/pywebview)（BSD-3-Clause）、[openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)（MIT）

详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

## 许可证

[MIT](LICENSE) © 2026 肥鸭文档转换 contributors

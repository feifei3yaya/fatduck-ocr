# 第三方组件许可

肥鸭文档转换以 MIT 许可证发布，并使用了以下第三方组件。各组件的版权与许可归其作者所有。

## 运行期依赖

| 组件 | 用途 | 许可证 | 项目地址 |
| --- | --- | --- | --- |
| PaddleOCR PP-OCRv6 模型 | 文字检测与识别模型 | Apache-2.0 | https://github.com/PaddlePaddle/PaddleOCR |
| OpenVINO | CPU 推理运行时 | Apache-2.0 | https://github.com/openvinotoolkit/openvino |
| OpenCV | 图像处理 | Apache-2.0 | https://github.com/opencv/opencv |
| NumPy | 数值计算 | BSD-3-Clause | https://github.com/numpy/numpy |
| pypdfium2 | PDF 渲染 | Apache-2.0 / BSD-3-Clause | https://github.com/pypdfium2-team/pypdfium2 |
| FastAPI | 本地 API 服务 | MIT | https://github.com/fastapi/fastapi |
| Uvicorn | ASGI 服务器 | BSD-3-Clause | https://github.com/encode/uvicorn |
| Starlette | ASGI 框架 | BSD-3-Clause | https://github.com/encode/starlette |
| Pydantic | 数据校验 | MIT | https://github.com/pydantic/pydantic |
| python-multipart | 表单上传解析 | Apache-2.0 | https://github.com/Kludex/python-multipart |
| websockets | WebSocket 通信 | BSD-3-Clause | https://github.com/python-websockets/websockets |
| pywebview | 桌面窗口 | BSD-3-Clause | https://github.com/r0x0r/pywebview |
| pythonnet | .NET 互操作 | MIT | https://github.com/pythonnet/pythonnet |
| pystray | 系统托盘 | LGPL-3.0 | https://github.com/moses-palmer/pystray |
| Pillow | 图像处理 / 托盘图标 | MIT-CMU | https://github.com/python-pillow/Pillow |
| openpyxl | Excel 读写 | MIT | https://foss.heptapod.net/openpyxl/openpyxl |
| windnd | 文件拖拽 | MIT | https://github.com/CLOUGH/windnd |

## 构建期依赖

| 组件 | 用途 | 许可证 | 项目地址 |
| --- | --- | --- | --- |
| PyInstaller | 打包 exe | GPL-2.0 with exception | https://github.com/pyinstaller/pyinstaller |
| Inno Setup | 生成安装包 | Inno Setup License | https://jrsoftware.org/isinfo.php |

## 字体与图标

- 界面仅使用系统字体，不随包分发字体文件。
- 鸭头吉祥物、图标由项目作者提供，随项目以 MIT 许可证发布。

## 说明

- PyInstaller 的 GPL 例外允许将打包后的程序以其他许可证发布，打包产物本身不因此变为 GPL。
- pystray 以 LGPL-3.0 发布，本项目以动态依赖方式使用，未修改其源码。
- 如果你再分发本项目的安装包，请保留本文件与 `LICENSE`。

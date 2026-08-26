# OfficePicture

OfficePicture 是一个面向 Word、PowerPoint 和 Excel 桌面版的原生 VSTO 图片预览插件，不是 Web Add-in。

在文档中选中已经插入的图片并双击，插件会弹出独立的大图预览窗口。图片来自当前 Office 文档中的选中对象，不需要本地文件路径或 URL。

## 当前功能

- Word：支持嵌入式图片和浮动图片。
- PowerPoint：支持幻灯片中的嵌入图片。
- Excel：支持 Open XML 工作簿中的嵌入图片 Shape；通过 Excel 窗口双击消息触发。
- 预览窗口：适应窗口、100%、放大、缩小、鼠标滚轮缩放；缩放后可按住鼠标左键拖动查看，按 `Esc` 关闭。
- 缩放范围：10%～800%；较小图片在“适应窗口”模式下也会自动放大。
- 预览以当前 Office 窗口为父窗口，采用无标题栏的沉浸式预览层，不显示独立任务栏图标。
- 双击触发带有重入保护和关闭后的短暂去重，不会在关闭后再次弹出同一次预览。
- 直接读取 Office 文档内部的原始媒体，不调用复制命令、不访问系统剪贴板。
- “100%”使用原始媒体像素 1:1 绘制，不经过缩放插值。

图表、OLE 对象、组合对象和非图片 Shape 不会触发预览。

## 项目结构

- `OfficePicture.Core`：WinForms 预览弹窗、Open XML 原始媒体提取、Office 窗口适配和 Excel 双击钩子。
- `OfficePicture.WordAddIn`：Word VSTO Add-in。
- `OfficePicture.PowerPointAddIn`：PowerPoint VSTO Add-in。
- `OfficePicture.ExcelAddIn`：Excel VSTO Add-in。

## 产品与设计文档

完整的产品倒推资料位于 [`docs`](docs/README.md)，包括产品需求文档（PRD）、交互与视觉规范、可点击原型、技术设计、测试/发布/运维方案和需求追踪矩阵。文档明确区分当前 1.0 已实现能力与后续规划，可作为类似 Office 插件或单任务桌面工具的项目参考。

## 环境

- Visual Studio 2026，安装“Microsoft 365 开发”工作负载及 VSTO 工具。
- .NET Framework 4.7.2 Developer Pack（项目目标框架）。
- Microsoft Office 桌面版 Word、PowerPoint、Excel。
- Visual Studio Tools for Office Runtime。

当前开发机已检测到 Visual Studio Professional 2026、Office 2021 64 位、VSTO 工具和 .NET Framework 开发工具。

## 调试

1. 打开 `OfficePicture.sln`。
2. 将需要测试的 Word、PowerPoint 或 Excel Add-in 项目设为启动项目。
3. 按 `F5`，Visual Studio 会启动对应 Office 程序。
4. 新建或打开文档，插入图片。
5. 单击选中图片，再双击图片，确认出现“图片预览”弹窗。

如果构建提示 ClickOnce manifest 未签名，请打开对应 Add-in 项目的“属性 > 签名”，勾选“为 ClickOnce 清单签名”，创建本机测试证书。三个 Add-in 项目需分别配置一次。测试证书 `*.pfx` 已被 Git 忽略，不应提交。

## 发布

在 PowerShell 中运行：

```powershell
.\build\Publish.ps1 -Version 1.0.0.0
```

脚本会读取 Word 项目的本机签名证书，并为 Word、PowerPoint、Excel 生成统一签名的离线 ClickOnce 安装包，输出到 `publish`。`publish` 和私钥证书均不会提交到 Git。

Visual Studio 自动创建的测试证书仅适合本机或内部测试。向其他用户正式分发前，应换用受信任的代码签名证书。

## 许可证

本项目采用 [PolyForm Noncommercial License 1.0.0](LICENSE) 开源，仅允许非商业用途。商业使用、商业分发或商业服务集成不在授权范围内；如需商业授权，请联系版权持有人。

## Git

仓库已经初始化，`.gitignore` 已忽略 Visual Studio、VSTO 构建产物、用户配置和测试证书。

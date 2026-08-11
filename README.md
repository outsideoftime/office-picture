# OfficePicture

这是一个基于 VSTO/Office 原生对象模型的图片预览插件，不是 Web Add-in。

它会在 Word、PowerPoint、Excel 中显示原生 Custom Task Pane；当用户选中文档内已插入的图片时，插件读取 Office 的选中对象并复制图片数据到预览窗格。三种 Office 宿主分别使用独立的 VSTO Add-in 项目，共用 `OfficePicture.Core` 的预览控件。

## 项目结构

- `OfficePicture.Core`：WinForms 原生预览控件和剪贴板图片捕获逻辑。
- `OfficePicture.WordAddIn`：Word 图片选中事件与 Custom Task Pane。
- `OfficePicture.PowerPointAddIn`：PowerPoint 图片选中事件与 Custom Task Pane。
- `OfficePicture.ExcelAddIn`：Excel 图片选中事件与 Custom Task Pane。

## 环境检查结果

- Visual Studio Professional 2026 18.5.3：已安装且可启动。
- .NET Framework 4.8 Developer Pack：已安装。
- Office 2021 桌面版：已检测到。
- Git 2.48.1：已安装。
- 当前 Visual Studio 的 Office/VSTO 模板文件存在，但安装器组件查询未返回 Office Developer Tools；首次编译若提示找不到 `Microsoft.Office.Tools` 或 `OfficeTools` targets，需要在 Visual Studio Installer 中补装“Office/SharePoint 开发”工作负载/组件。

## 使用方式

1. 打开 `OfficePicture.sln`。
2. 在 Visual Studio Installer 中确认已安装：
   - Office/SharePoint 开发（Office Developer Tools）。
   - .NET Framework 4.8 targeting pack。
   - Visual Studio Tools for Office Runtime（VSTO Runtime）。
3. 分别将 `OfficePicture.WordAddIn`、`OfficePicture.PowerPointAddIn`、`OfficePicture.ExcelAddIn` 设为启动项目进行调试。VSTO 会启动对应的 Office 宿主。
4. 在 Office 文档中插入并选中图片，右侧会出现“图片预览”原生任务窗格。

## 当前行为

- Word：监听选区变化，支持 InlineShape 和浮动 Shape 图片。
- PowerPoint：监听形状选择，支持普通图片和链接图片。
- Excel：监听窗口选区变化并读取当前 Shape 图片。
- 预览使用的是 Office 当前选中的图片，不需要本地图片路径或 URL。

插件通过 Office 的复制接口捕获渲染后的图片，因此对常见 PNG、JPEG、截图、剪贴画和链接图片都适用。个别 OLE 对象、图表或受保护文档不是普通图片时会跳过预览。

## Git

仓库已初始化，`.gitignore` 已包含 Visual Studio、VSTO 构建产物和用户配置文件。

```powershell
git status
git add .
git commit -m "创建 Office 原生图片预览插件"
```

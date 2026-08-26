# OfficePicture 需求追踪矩阵

## 1. 使用说明

本表把 PRD 中的 P0 需求映射到现有代码证据、设计条目和测试用例。它既是倒推依据，也用于防止后续修改只更新代码而遗漏需求或测试。

状态含义：

- **已实现**：当前代码存在明确实现。
- **部分实现**：主路径已实现，但诊断、边界或兼容性仍有缺口。
- **待验证**：代码存在，尚需在真实 Office 环境执行验收。

## 2. 追踪表

| 需求 | 代码证据 | 交互/技术设计 | 测试 | 状态/说明 |
|---|---|---|---|---|
| FR-001 启动与事件注册 | 三个 `ThisAddIn.cs` 的 Startup/Shutdown；`OfficeDoubleClickHook.Dispose` | 技术 5、6、9 | TC-001/003/004/019 | 已实现，Excel Hook 失败诊断待补 |
| FR-002 支持对象识别 | `IsPicture`、`IsPictureSelection` | PRD 7；技术 6 | TC-001～005 | 已实现，类型数值应具名化 |
| FR-003 双击触发与去重 | `_previewOpen`、`_suppressPreviewUntilUtc`、`_callbackPending` | PRD 9；技术 5、9 | TC-012 | 已实现，400 ms 为经验值 |
| FR-004 图片捕获 | `ClipboardImageCapture.TryCapture`；宿主 Copy 委托 | 技术 7 | TC-001～004/015/020 | 已实现，异步剪贴板待强化 |
| FR-005 剪贴板保护 | `SnapshotClipboard` 与 `finally SetDataObject` | 技术 7、11 | TC-013～015 | 部分实现，并发覆盖风险明确 |
| FR-006 预览窗口 | `ImagePreviewForm` 构造、`ShowPreview`、`PlaceOverOwner` | UX 4、8、11；技术 8 | TC-001/003/004/017/018 | 已实现，DPI 需验证 |
| FR-007 自适应显示 | `Shown`/`Resize`、`FitImage`、`ApplyZoom` | UX 6/7；技术 8 | TC-006/011/018 | 已实现，Manual Resize 居中有缺口 |
| FR-008 缩放控制 | `ChangeZoom`、`SetZoom`、MouseWheel、KeyDown | UX 6；技术 8 | TC-007～011 | 已实现 |
| FR-009 关闭与释放 | 关闭按钮、Esc、`Dispose` | UX 6/10；技术 9 | TC-010/016/019 | 已实现，长稳待验证 |
| FR-010 安装与分发 | `build/Publish.ps1`、三个 Add-in `.csproj`、`publish/` | 技术 12；测试 8/9 | 安装/升级/卸载验收 | 部分实现，正式证书与 CI 待落地 |
| NFR-001 性能 | 同步内存路径，无网络 | 技术 7/9 | TC-006/016/020 | 待测量，无遥测 |
| NFR-002 稳定性 | 宿主 catch、Dispose、Hook 卸载 | 技术 9/10 | TC-012/015/016/019/020 | 部分实现，需长稳数据 |
| NFR-003 隐私安全 | 无网络/落盘代码；证书不入 Git | 技术 11/12 | 发布校验、TC-013～015 | 设计满足；签名流程需产品化 |
| NFR-004 兼容性 | .NET 4.7.2、AnyCPU、Interop 15.0 | 技术 13 | 环境矩阵、TC-017/018 | 待验证 |
| NFR-005 无障碍 | 可见文本、Esc、+/- | UX 10 | TC-010/018 | 部分实现，Tab/焦点/高对比待补 |
| NFR-006 可维护性 | Core 与三宿主项目分离 | 技术 3/4/16 | 代码审查 | 已有基础，缺自动化测试 |

## 3. 代码—产品事实索引

| 产品事实 | 直接来源 |
|---|---|
| 默认 1000×720、最小 480×360 | `ImagePreviewForm` 构造函数 |
| 无标题栏、无任务栏图标 | `FormBorderStyle.None`、`ShowInTaskbar=false` |
| 宿主内缩 12 px | `PlaceOverOwner` |
| 视口边距 20 px | Panel Padding 与位置计算 |
| 缩放 10%～800% | `MinZoom` / `MaxZoom` |
| 按钮步长 1.2、滚轮步长 1.15 | `ChangeZoom` 调用 |
| 首次与 Fit 模式 Resize 自动适应 | `Shown` / `Resize` 事件 |
| 关闭后 400 ms 去重 | 三个 `ThisAddIn.cs` |
| Word/PPT 宿主双击事件 | `WindowBeforeDoubleClick` 订阅 |
| Excel 当前线程消息 Hook | `SetWindowsHookEx(WH_GETMESSAGE)` |
| Excel 文档窗口筛选 | 窗口类名 `EXCEL7` |
| 剪贴板快照与恢复 | `ClipboardImageCapture` |
| 全本地、无源路径依赖 | 捕获与预览代码无网络/文件路径 |
| 三个 ClickOnce 包统一版本与签名参数 | `build/Publish.ps1` |

## 4. 已识别缺口清单

| ID | 缺口 | 影响 | 建议优先级 |
|---|---|---|---|
| GAP-01 | 捕获失败完全静默且无诊断 | 支持困难 | P1 |
| GAP-02 | 剪贴板恢复可能覆盖并发新内容 | 潜在用户数据风险 | P0/P1，先验证再修复 |
| GAP-03 | 无剪贴板延迟重试 | 某些大图/环境触发失败 | P1 |
| GAP-04 | 超大图无内存上限/降采样 | 宿主卡顿或 OOM | P1 |
| GAP-05 | Manual 模式 Resize 不主动重新居中 | 视觉一致性 | P2 |
| GAP-06 | 无系统化 DPI/高对比度支持证据 | 可用性与无障碍 | P1 |
| GAP-07 | `catch {}` 隐藏所有宿主错误 | 难诊断 | P1 |
| GAP-08 | 图片类型使用魔法数字 | 维护与兼容风险 | P2 |
| GAP-09 | 三宿主独立安装，正式签名依赖本机配置 | 交付复杂 | P1 |
| GAP-10 | 无自动化测试项目 | 回归成本高 | P1 |
| GAP-11 | 当前本机仅 Word 启用清单签名，PPT/Excel 使 Release 解决方案构建最终失败 | 发布阻断 | P0（发布前配置） |

## 5. 变更控制规则

后续每次功能变更至少完成：

1. 在 PRD 新增或修改带 ID 的 FR/NFR。
2. 在交互或技术文档记录状态、算法、数据流或架构变化。
3. 在本矩阵补充代码位置和测试用例。
4. 在测试文档增加正向、边界、失败与回归用例。
5. 若改变支持范围、权限、数据处理或部署方式，更新 README 与发布说明。

# PPTXMate

[English](./README.md) | [简体中文](./README.zh-CN.md)

PPTXMate 是一个开源的 PowerPoint AI 助手和 Office 加载项。它支持阅读、编辑、翻译和自动化 PPT/PPTX 演示文稿，并通过按任务获取上下文，减少不必要的模型 token 消耗。

> PPTXMate 是独立项目，与 Microsoft 没有隶属、合作或认可关系。

## 核心特性

- **按任务路由上下文**：文本、版式、创建和验证类请求只暴露各自需要的提示词、工具和演示文稿字段。
- **分层且有边界的读取**：先获取紧凑的幻灯片目录，再预览指定页面；文本、几何、格式和 OOXML 仅在需要时读取。
- **稳定的幻灯片标识**：写入操作以 PowerPoint `slide_id` 为目标，不会因删除或重排页面而使用过期页码。
- **达到限制后的恢复**：通过受限输出、对话压缩、自动续写、变更回执和重试前检查，让同一对话可以继续使用。
- **剪贴板图片**：可直接在聊天输入框中粘贴 PNG、JPEG、GIF、WebP 或 BMP 图片。
- **BYOK 模型服务**：支持 API 密钥、已支持的 OAuth 流程和兼容的自定义端点。
- **可选 CC Switch 集成**：将本地 `/v1` 端点路由至 CC Switch 当前选中的账号。
- **可选 macOS 生命周期监听器**：打开 PowerPoint 时启动本地 PPTXMate 路由，关闭 PowerPoint 时停止它。

PowerPoint 的实现细节和完整工具列表见 [`packages/powerpoint`](./packages/powerpoint)。

## 安装托管版加载项

GitHub Pages 工作流会发布任务窗格和公开加载项清单：

[`https://klaragraff.github.io/pptxmate/manifest.prod.xml`](https://klaragraff.github.io/pptxmate/manifest.prod.xml)

下载该清单文件后，在 PowerPoint 中旁加载。各平台操作步骤见 [PowerPoint 包 README](./packages/powerpoint/README.md#install)。

托管版支持常规 BYOK 和兼容的 HTTPS 端点。CC Switch 使用 localhost 路由，因此需要使用下方的本地模式。

## 本地开发

要求：

- Node.js 22.13 或更高版本
- pnpm 11.9.0 或更高版本
- 用于加载项测试的 Microsoft PowerPoint 桌面版

```bash
git clone https://github.com/KlaraGraff/pptxmate.git
cd pptxmate
pnpm install
pnpm start:ppt
```

PowerPoint 开发服务地址为 `https://localhost:3001`。

## 可选 CC Switch 路由

PPTXMate 可以将本地兼容 OpenAI 的端点代理到 CC Switch。CC Switch 是需要单独安装、运行和配置的应用；PPTXMate 不会存储 CC Switch 的账号凭据。

默认路由为：

```text
https://localhost:3001/v1 -> http://127.0.0.1:15721/v1
```

在 PPTXMate 设置中，按 CC Switch 支持的 API 类型和模型配置自定义端点，并将基础 URL 设为 `https://localhost:3001/v1`。如需修改地址或禁用路由，请在启动本地服务前使用：

```bash
PPTXMATE_CC_SWITCH_URL=http://127.0.0.1:25721 pnpm dev-server:ppt
PPTXMATE_CC_SWITCH_ENABLED=0 pnpm dev-server:ppt
```

请在 CC Switch 中切换账号，之后的 PPTXMate 请求会使用当前选中的账号。该路由会拒绝跨域浏览器请求和非本地主机请求头。

## 可选 macOS 自动启动和停止

macOS 监听器不保存模型或账号凭据。它会检测 `Microsoft PowerPoint` 进程：PowerPoint 打开时启动本地 PPTXMate 开发服务及 CC Switch 路由；PowerPoint 关闭时仅结束由它管理的服务进程组。它不会启动、停止或切换独立的 CC Switch 应用账号。若端口 `3001` 已被其他进程使用，监听器不会干预该进程。

```bash
./scripts/install-macos-powerpoint-watcher.sh
```

若要让监听器持久使用自定义 CC Switch 地址，或关闭该路由，请在安装时指定：

```bash
./scripts/install-macos-powerpoint-watcher.sh --cc-switch-url http://127.0.0.1:25721
./scripts/install-macos-powerpoint-watcher.sh --no-cc-switch
```

使用 `./scripts/install-macos-powerpoint-watcher.sh --dry-run` 可预览生成的 LaunchAgent，不会安装或修改本地文件。

卸载命令：

```bash
./scripts/uninstall-macos-powerpoint-watcher.sh
```

使用 `./scripts/uninstall-macos-powerpoint-watcher.sh --remove-logs` 可同时删除该监听器自己的日志文件。

安装脚本会自动定位当前仓库、Node.js 和 pnpm 路径，并在仓库之外生成当前用户专用的 LaunchAgent。发现旧的 OpenPPT 监听器时，它也会自动迁移。

## 仓库结构

PPTXMate 保留了上游 Office Agents 单仓库中的共享 SDK 和 UI 包，以保持兼容性：

| 包 | 用途 |
| --- | --- |
| [`@office-agents/powerpoint`](./packages/powerpoint) | PPTXMate PowerPoint 加载项、路由、幻灯片工具和清单 |
| [`@office-agents/core`](./packages/core) | 共享 Svelte 聊天界面和剪贴板附件处理 |
| [`@office-agents/sdk`](./packages/sdk) | Agent 运行时、上下文恢复、服务商、存储、VFS 和沙箱 |
| [`@office-agents/bridge`](./packages/bridge) | 可选的本地运行时检查桥接 |
| [`@office-agents/excel`](./packages/excel) | 保留的上游 Excel 加载项包 |
| [`@office-agents/word`](./packages/word) | 保留的上游 Word 加载项包 |

为尽量减少与上游的分叉，内部包名保持不变。PowerPoint 存储标识符 `openppt-*` 和 `OpenPPTDB_v1` 也予以保留，以使既有本地设置和对话在产品更名后继续可用。监听器脚本中出现的旧 OpenPPT LaunchAgent 标签，只用于迁移或删除旧监听器。

## 验证

```bash
pnpm test
pnpm typecheck
pnpm lint
pnpm validate
pnpm build
```

## 上游与许可证

PPTXMate 源自 [hewliyang/office-agents](https://github.com/hewliyang/office-agents)，并保留其 Git 历史。归属说明见 [NOTICE.md](./NOTICE.md)。

项目以 [MIT License](./LICENSE) 发布。

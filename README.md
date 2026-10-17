# 企业微信活动通知工具 v2.0

Windows 10/11，Python/Tkinter/ttk；Excel 数据表格 → 条件筛选 → 勾选活动/字段 → 自动分批校验 → 企业微信机器人 / 自建应用推送。

## 2.0 新增

- **图片表格**：先在后台渲染为 JPEG，测量编码后实际大小，反复尝试有界压缩、缩放；超过安全阈值则先按行、再按列拆分，重新校验，直到所有图片在 1,950,000 字节以内。JPEG 保证通过 Pillow 解码验证；正式发送前检查全部图片。绝不因超限默默裁剪 Excel 单元格内容。
- **Markdown V2 表格（仅群机器人）**：使用 UTF-8 实际长度动态拆分，安全阈值 3900 字节（平台文档上限 4096 字节）；拆分后重复表头和分隔行，并重新计算、复核总页数。用户不必手动设置行数和列数。
- **字段自由勾选**：原来的字体符号已替换成统一绘制的勾选图片；单击左侧复选框/空格切换，多选高亮行后可空格切换，支持当前筛选结果批量勾选、反选、清空、上/下调整字段排序。
- **性能**：Excel 读取、数据量较大的筛选、预览和全批消息校验、HTTP 推送由后台线程执行；Treeview 每次分块绘制 125 行，默认展示 2000 行，点击“加载更多行”增量显示，批量操作仍覆盖全部筛选结果。Pillow 和 requests 延迟到导入文件、预览或发送时加载，缩短初始打开界面所需时间。
- **安全**：先完成全批图片/Markdown 校验及接收映射校验，再展示发送确认；如遇极端超长单个字段无法拆分，取消本批发送，而不是发出不完整数据。密钥和 Webhook 不持久保存。
- **打包**：只提供 one-dir 便携文件夹模式，依赖保留在 EXE 外，避免启动时解包；每次构建包含 SHA256 校验文件。

## 发送的两种方法

**A 群机器人（Webhook）**：群中设置机器人，在软件中填入 `https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=...`；支持图片消息（base64 + 原图 MD5）、`markdown_v2` 表格。机器人的图片消息原始 JPG/PNG 需不超过 2 MB，本项目保留余量。此处会按每条消息间隔 3.2 秒控制速率；速率策略不保证不同网络环境一定不会被服务端限流。

**B 企业自建应用**：填写 CorpID、Secret、AgentID 和 UserID，图片先通过 `POST /cgi-bin/media/upload?type=image` 上传获得 media_id，再通过 `/cgi-bin/message/send` 发送。自建应用不走 `markdown_v2` 表格模式；请选“图片表格”。

> 重要：群机器人发到**所在群**，不能直接代表店端一对一私聊。跨服务店批量推送时应使用接收映射 CSV，`任务服务店编码,Webhook` 或 `任务服务店编码,UserID`。映射缺失将阻止发送，避免泄露店端数据。

## 快速使用

1. 导入 `.xlsx` / `.xlsm` / `.xls` / `.csv`（含多工作表）。
2. 选择活动状态、门店/区域、活动编号/名称、完成率阈值、任务数量等条件；点击 **执行筛选**。
3. 勾选结果记录（单击左侧复选框或空格）或 **勾选当前筛选结果**。
4. 第二页勾选需要推送的字段及调整顺序，选择**图片表格**或**Markdown表格**。行数/列数无需配置。
5. 点击 **校验并预览消息**，程序会预先生成全部消息并提示预计条数。若是图片，保存按钮可得到与实际推送相同的 JPEG 文件。
6. 配好接收目标，点击 **确认并发送**。软件会后台验证所有消息与所有接收目标，再弹出最终确认，联网发送。可以停止后续尚未发送的消息。

> 说明：自动分批可以处理整体超限，**不能保证任意无限长的单个单元格都可以完整展示**。如果拆到单条记录单个字段仍无法满足图片尺寸/体积或 Markdown 大小要求，会提示该字段并阻止整批发送。企业微信接口实际能否接收仍取决于密钥、账号权限、网络、调用频率等，未使用真实凭证在此环境测试。

## Windows EXE 构建

### GitHub Actions（无需在 Windows 上安装 Python）

推送到 GitHub 仓库后会自动构建，也可以进入 **Actions → Build Windows EXE → Run workflow** 手动启动。

- 产物 `WeComActivityNotifier-Windows-Portable` 是 ZIP，解压后保留完整的 `WeComActivityNotifier` 目录，运行其中的 `WeComActivityNotifier.exe`，无需额外安装 Python。依赖文件放在目录里，不会集成到 EXE 内。
- 每个产物包含 `SHA256.txt`，用于下载后校验二进制是否与构建时一致。

在 macOS Finder 中按 `Command + Shift + .` 显示 `.github` 等隐藏目录。GitHub 网页上传器对文件夹结构支持有限；建议在源码目录使用 `git add -A && git commit -m 'Update notifier v2' && git push` 推送（首次使用需要配置 Git 仓库）。不要把真实 Secret、Webhook 和接收映射提交到仓库。

### Windows 本地

双击 `build_windows.bat` 构建便携目录版。构建机需 Python 3.11/3.12 和联网安装依赖；生成的应用不要求目标机器安装 Python。

### Python 直接启动

```sh
python -m pip install -r requirements.txt
python main.py
```

### 回归测试

```sh
python -m unittest discover -s tests -v
```

## 浏览器提示 ZIP / EXE 是危险文件？

**不要直接关闭防病毒软件或添加整个下载目录白名单。** 先确认你的下载确实来自自己 GitHub 仓库的 Actions 页面：

1. 在浏览器下载记录中查看警告源头：Chrome/Edge 的下载防护、Windows SmartScreen，还是 Windows Defender 的病毒检出（威胁名称）。
2. 若只是“未常见下载 / 未知发布者”，该提示可能来自新构建、未签名的应用信誉不足，并非证明恶意，但也不是安全凭证；核对仓库源码和 GitHub 构建日志。
3. 解压后按 Windows PowerShell 执行 `Get-FileHash .\\WeComActivityNotifier.exe -Algorithm SHA256`，与随包的 `SHA256.txt` 比较。相同只证明与构建产物一致，**不等于证明无毒**。
4. 用 Windows 安全中心扫描目标目录；如果出现明确的 Trojan / malware 名称，不要强行运行，应提交微软安全情报样本分析或交由企业 IT 审核。
5. 正式发给更多同事使用时，考虑合法的 Authenticode 代码签名（微软 Artifact Signing 等），提升发布者信誉；即使已签名，首次发布仍可能有 SmartScreen 提示。

官方参考：https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation

## 其他说明

- 不包含定时发送，不会自动对企业微信鉴权。必须使用你有权限的正式接口。
- Excel 公式按已保存的结果缓存读取（公式数据无缓存时需要用 Excel 计算并保存）。
- 表格一行对应 Excel 一条活动记录；不同活动的任务台数不会自动去重成唯一车辆数。
- 部分老旧企业微信客户端对 Markdown V2 表格渲染可能存在差异，图片模式对手机可读性更稳定。
- 源码主要文件：`engine.py`（解析、分页校验、推送）和 `main.py`（界面、异步任务）。

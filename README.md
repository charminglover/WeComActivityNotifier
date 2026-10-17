# 企业微信活动通知工具

这是基于附件中的 `main.py`、`engine.py` 和演示工作簿整理的 Windows x64 打包工程。

## 构建结果

GitHub Actions 会生成单文件：

```text
WeComActivityNotifier.exe
```

目标系统为 Windows 10/11 x64。程序通过 PyInstaller `--onefile --windowed` 打包，运行电脑无需另装 Python。首次启动可能需要几秒钟解包。

## GitHub Actions 构建

将本目录内容提交到 GitHub 仓库的 `main` 或 `master` 分支后，工作流会在相关文件变更时自动构建。也可以在仓库的 **Actions → Build Windows EXE → Run workflow** 手动启动。

构建成功后，在对应运行记录页面的 **Artifacts** 区域下载 `WeComActivityNotifier-Windows-x64`，解压即可得到 EXE。演示工作簿只用于构建时验证导入、筛选和图片表格生成，不会被嵌入 EXE；用户可以在程序中选择自己的 Excel 文件。

## 功能与凭据

程序保留现有功能：Excel/CSV 导入、活动筛选与勾选、推送字段和分页设置、图片表格或 Markdown 消息生成、企业微信机器人或自建应用推送。Webhook、Secret 等发送凭据不会写入偏好设置文件。

## 本地 Windows 构建

在 Windows 安装 Python 3.12 后，可运行：

```powershell
python -m pip install -r requirements.txt
python build_smoke_test.py
python -m PyInstaller --noconfirm --clean --onefile --windowed --name WeComActivityNotifier --hidden-import xlrd --collect-all openpyxl --collect-all PIL --collect-all certifi main.py
```

产物位于 `dist\WeComActivityNotifier.exe`。

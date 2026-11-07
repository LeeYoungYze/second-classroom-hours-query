# 第二课堂学时自动查询工具

自动登录学校「第二课堂」系统，抓取各类学时数据，并生成一张带颜色标记、自动计算差值的 Excel 统计表。

> 说明：本仓库**不包含** `VC_redist.x64.exe` 运行库安装包，请按下方第一步从微软官方链接下载。

## 功能

- 支持**账号密码全自动登录**：自动识别算术验证码、自动重试
- 支持**微信扫码登录**：适合不想在配置文件里写密码的情况
- 自动抓取成绩单接口数据，无需手动复制
- 生成彩色 Excel 表格：毕业需求 / 目前学时 / 差值 / 完成度，未完成项红色、已完成项绿色
- 自动优先使用 Chrome，未安装则回退到 Edge

## 运行环境

- Windows 10 / 11
- Python 3.9 及以上
- Chrome 或 Edge 浏览器（任一即可）

## 第一步：安装微软运行库（必装）

程序依赖的 OCR 组件需要 Visual C++ 运行库，不装会直接闪退报错。

下载地址（微软官方，x64）：

**https://aka.ms/vs/17/release/vc_redist.x64.exe**

下载后双击，一路「下一步」装完，建议重启一次电脑。

## 第二步：安装 Python

到 https://www.python.org/ 下载安装包。安装第一页最下方务必勾选 **Add Python.exe to PATH**。

## 第三步：安装依赖

```bash
pip install -r requirements.txt
```

依赖只有三个：`DrissionPage`、`ddddocr`、`openpyxl`，其余由 pip 自动安装。

用 PyCharm / VS Code 打开项目时，编辑器一般也会提示安装缺少的库，点 Install 即可。

## 第四步：配置账号密码

打开 `account.json`：

**扫码登录**（默认，推荐）：

```json
{
  "YOUR_ID" : "",
  "YOUR_PWD" : ""
}
```

**全自动登录**：

```json
{
  "YOUR_ID" : "你的学号",
  "YOUR_PWD" : "你的密码"
}
```

注意：必须使用英文双引号，不能用中文引号，否则会报 `JSONDecodeError`。

## 第五步：运行

```bash
python main.py
```

运行前请先关闭任何正在打开的 `第二课堂学时统计.xlsx`。

如果 `account.json` 为空，程序会停在二维码界面：在弹出的浏览器里用手机扫码登录，然后**回到命令行按一次回车键**，程序才会继续抓取数据。

抓取成功后，当前目录下会生成 `第二课堂学时统计.xlsx`。

### 自定义保存位置（可选）

打开 `main.py`，拉到最下面，找到这一行：

```python
SAVE_PATH = ""
```

- **留空**（默认）：保存到当前文件夹，文件名仍是 `第二课堂学时统计.xlsx`
- **填文件夹**：`SAVE_PATH = r"D:\我的文件\第二课堂"`
- **填完整路径**：`SAVE_PATH = r"D:\我的文件\我的学时统计.xlsx"`

路径前面的 `r` 不要删，否则反斜杠会被当成转义符。文件夹不存在时程序会自动创建，并在控制台打印最终保存的完整路径，方便你直接找到文件。

## 常见问题

| 现象 | 原因与解决 |
| --- | --- |
| `ModuleNotFoundError: No module named 'xxx'` | 第三步没装依赖，或装到了别的 Python 环境 |
| 黑框一闪而过 / 找不到 python.exe | 没装第一步的 VC++ 运行库，或安装 Python 时没勾选 Add to PATH |
| `JSONDecodeError` | `account.json` 格式被改坏，检查引号是否为英文 |
| `Permission denied: '第二课堂学时统计.xlsx'` | 表格正开着，关掉 Excel 重新运行 |
| 扫码后黑框没反应 | 需要在命令行按一次回车键 |

## 目录结构

```
main.py              主程序
account.json         账号密码配置
requirements.txt     依赖清单
使用说明.txt          面向新手的图文步骤说明
README.md            本文件
```

## 免责声明

本项目仅用于学习交流与个人查询自己的学时数据，请勿用于任何商业用途或非法用途。使用产生的后果由使用者自行承担。

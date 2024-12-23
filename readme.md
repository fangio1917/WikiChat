# 使用步骤

### 1. 配置 API 密钥

首先，打开 `config.py` 文件，并输入你的 API 密钥。`config.py` 文件中有一个名为 apikey的变量，你需要将其替换为你自己的 API 密钥，并制定你的模型

```python
# config.py

apikey = "sk-" 
model = ""
```

### 2. 执行脚本

打开命令行，进入项目目录，并执行以下命令：

```bash
python ./auto_chat.py ./data/input/test.xlsx ./data/output/output.docx 2  
```

#### 参数说明：

* `input.xlsx`：你的输入文件，应该是一个有效的 `.xlsx` 文件，包含你希望处理的数据。
* `out.docx`：处理后的输出文件，将生成一个 `.docx` 格式的文件，包含处理后的结果。
* `mode` : 生成模式 1:only word; 2:only pic; 3: both

### 3. 文件结构

```bash
.
├── auto_chat.py
├── config.py
├── data
│   ├── input
│   │   └── Eng.xlsx
│   └── output
├── pic
├── readme.md
├── requirements.txt
├── test.py
└── utils.py
```

### 4. 安装依赖

运行命令行即可自行安装，你也可以通过注释来自定义

```python
#auto_chat.py:4

# 获取命令行参数
args = sys.argv
if len(args) != 3:
    print("请提供输入文件路径和输出文件路径作为命令行参数。eg: python auto_chat.py test.xlsx output.docx")
    sys.exit(1)
subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
```

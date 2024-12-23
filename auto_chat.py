import sys
import subprocess

# 获取命令行参数
args = sys.argv
if len(args) != 4:
    print("请提供输入文件路径和输出文件路径和模式作为命令行参数。eg: python auto_chat.py test.xlsx output.docx mode")
    sys.exit(1)
subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

mode = sys.argv[3]
# 1 only word
# 2 only pic
# 3 both
if mode not in ["1","2","3"]:
    print("mode only 1 2 3")
    sys.exit(1)



import wikipedia
import utils
import config as cfg
import openai

# 设置 OpenAI API 密钥
openai_cfg = cfg.Config()# Config 类的实例, 需要输入 open_api_key 和 model
openai.api_key = openai_cfg.open_api_key
model = openai_cfg.model

wikipedia.set_lang("zh")

# 处理输入文件
equipment_names,doc = utils.read_excel_and_generate_word(sys.argv[1])

# 遍历每个装备名称
for index, equipment in enumerate(equipment_names, start=1):
    if not isinstance(equipment, str) or equipment.strip() == "":
        print(f"跳过空装备名称，行号：{index}")
        continue

    print(f"正在处理装备：{equipment}")

    # 查询简介和作战运用
    try:

        # 使用 Wikipedia 库查询装备信息
        page = wikipedia.page(equipment)
        wiki = page.content

        if mode == "1" or mode == "3":
        # 数据源 1
            pormpt_1 = f"请为{equipment}生成一段简介，数据来源：简氏防务。"
            intro_1 = utils.openai_chat(model,pormpt_1)

            pormpt_2 = f"请为{equipment}的作战运用，数据来源：简氏防务。"
            operation_1 = utils.openai_chat(model,pormpt_2)

            # 数据源 2


            pormpt_3 = f"请为{equipment}生成一段简介，数据来源：维基百科。"+wiki
            intro_2 = utils.openai_chat(model,pormpt_3)
            pormpt_4 = f"请为{equipment}的作战运用，数据来源：维基百科。"+wiki
            operation_2 = utils.openai_chat(model,pormpt_4)


            utils.write_word(doc, f"{equipment} 简介", intro_1, intro_2)
            utils.write_word(doc, f"{equipment} 作战运用", operation_1, operation_2)

        if mode == "2" or mode == "3":
            pic = utils.get_wiki_first_pic(page)
            pic_local = f"./pic/{equipment}.jpg"
            utils.save_image_from_url(pic, pic_local)

        if mode == "1" or mode == "3":
            utils.write_pic_to_word(doc, pic_local)

    except Exception as e:
        print(f"查询失败：{e}")
        continue


if mode == "1" or mode == "3":
    # 保存 Word 文档
    word_output = sys.argv[2]
    doc.save(word_output)
    print(f"装备数据报告已保存到 {word_output}")

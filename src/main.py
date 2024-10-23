import datetime
import os
import argparse
import gradio as gr
from langchain_openai import ChatOpenAI
from langchain.prompts import ChatPromptTemplate
from langchain.schema import StrOutputParser
from input_parser import parse_input_text
from ppt_generator import generate_presentation
from template_manager import load_template, get_layout_mapping
from layout_manager import LayoutManager
from config import Config
from logger import LOG


def load_formatter_prompt():
    with open('prompts/formatter.txt', 'r') as file:
        return file.read()

def create_format_chain():
    llm = ChatOpenAI(model_name="gpt-4o-mini")
    prompt = ChatPromptTemplate.from_messages([
        ("system", load_formatter_prompt()),
        ("human", "{input}")
    ])
    chain = prompt | llm | StrOutputParser()
    return chain

format_chain = create_format_chain()

def format_content(input_text):
    return format_chain.invoke({"input": input_text})

def process_input(formatted_text):
    config = Config()
    prs = load_template(config.ppt_template)
    layout_mapping = get_layout_mapping(prs)
    layout_manager = LayoutManager(config.layout_mapping)

    powerpoint_data, presentation_title = parse_input_text(formatted_text, layout_manager)
    LOG.debug(f"解析的 PowerPoint 数据: {powerpoint_data}")

    now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    # output_pptx = f"outputs/{presentation_title}.pptx"
    output_pptx = f"outputs/output_{now}.pptx"
    generate_presentation(powerpoint_data, config.ppt_template, output_pptx)

    return f"演示文稿已生成: {output_pptx}"

def chatbot(message, history):
    if message.lower() in ["help", "instructions", "说明", "使用说明"]:
        help_text = """
欢迎使用 AI 辅助 PPT 生成工具！以下是使用说明：

1. 输入内容：直接输入您想要包含在演示文稿中的内容。可以是概要、大纲或详细内容。

2. AI 格式化：系统会自动格式化您的输入，组织成适合 PPT 的结构。

3. 审核和修改：检查 AI 格式化的结果。如果需要修改，直接输入新的内容或修改建议。

4. 添加更多内容：您可以多次输入来逐步构建您的演示文稿。

5. 生成 PPT：当您对内容满意时，输入 "Generate PPT" 来创建最终的 PowerPoint 文件。

6. 查看帮助：随时输入 "help" 或 "说明" 来查看这些使用说明。

开始吧！输入您的演示文稿内容，AI 将帮助您组织和格式化。
        """
        return help_text
    elif message.lower() == "generate ppt":
        last_formatted_content = history[-1][1].split("\n\n")[2]  # 获取上一次格式化的内容
        response = process_input(last_formatted_content)
        return f"{response}\n\n如需创建新的演示文稿，请直接输入新内容。输入 'help' 查看使用说明。"
    elif message.startswith("```md") and message.endswith("```"):
        # 检测到 Markdown 格式的输入
        md_content = message.strip("```md").strip("```").strip()
        response = process_input(md_content)
        return f"{response}\n\n您的 Markdown 内容已直接用于生成 PPT。如需创建新的演示文稿，请直接输入新内容。输入 'help' 查看使用说明。"
    else:
        formatted_content = format_content(message)
        return f"新内容已添加并格式化。请检查并确认：\n\n{formatted_content}\n\n如需继续添加内容，请直接输入。如果确认无误，请输入 'Generate PPT'。输入 'help' 查看使用说明。"

def launch_gradio():
    iface = gr.ChatInterface(
        chatbot,
        title="AI-Assisted PPT Generator",
        description="输入您的演示文稿内容，AI 将帮助您组织和格式化。确认内容后，输入 'Generate PPT' 生成演示文稿。输入 'help' 查看详细使用说明。",
        examples=[
            ["我想做一个关于人工智能在医疗领域应用的演示。主要包括三个方面：诊断辅助、药物研发和个性化治疗。"],
            ["再加上一个关于 AI 在医疗伦理方面的考虑"],
            ["Generate PPT"],
            ["help"],
        ],
        retry_btn=None,
        undo_btn=None,
        clear_btn="Clear",
    )
    iface.launch()

def main(input_file):
    # 原有的命令行处理逻辑
    config = Config()

    if not os.path.exists(input_file):
        LOG.error(f"{input_file} 不存在。")
        return
    
    with open(input_file, 'r', encoding='utf-8') as file:
        input_text = file.read()

    formatted_content = format_content(input_text)
    result = process_input(formatted_content)
    print(result)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='从 markdown 文件生成 PowerPoint 演示文稿。')
    parser.add_argument(
        'input_file',
        nargs='?',
        default='inputs/test_input.md',
        help='输入 markdown 文件的路径（默认: inputs/test_input.md）'
    )
    parser.add_argument(
        '--gradio',
        action='store_true',
        help='启动 Gradio 交互界面'
    )
    
    args = parser.parse_args()

    if args.gradio:
        launch_gradio()
    else:
        main(args.input_file)

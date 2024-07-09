import os
from langchain_openai import ChatOpenAI
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def get_correction_suggestion(text, type, api_url="https://api.rcouyi.com/v1/"):
    api_key = "sk-IwmKcBKY42zk7gqsE3B502F9699d476b808e573fE4A74206"
    # 设置环境变量
    os.environ["OPENAI_API_KEY"] = api_key

    chat = ChatOpenAI(model="qwen-turbo", openai_api_base=api_url)
    base_instruction1 = f"""
    你是一个语法检查助手，你的任务是检查一段文字是否出现错别字和语法错误,你应该检查我提供给你的的文字中的语法错误和错别字，只返回你的修改建议。
    文本也有可能没有错误,如果没有错误可以返回“没有错误”。
    {text}

    """
    base_instruction2 = f"""
        你是一个论文检查助手，我会给你提供参考文献的格式要求，你的任务是检查一段参考文献的引用是否出现错误，只返回你的修改建议,以为修改理由。你不需要关注文献的序号。
        参考文献的著录应符合国家标准，参考文献的序号左顶格，并用数字加方括号表示，如“[1]”。每一条参考文献著录均以“.”结束。具体各类参考文献的编排格式如下：
1．文献是期刊时，书写格式为：
[序号] 作者. 文章题目[J]. 期刊名, 出版年份，卷号(期数):起止页码.
2．文献是图书时，书写格式为：
[序号] 作者. 书名[M]. 版次. 出版地：出版单位，出版年份：起止页码.
3．文献是会议论文集时，书写格式为：
[序号] 作者. 文章题目[A].主编.论文集名[C], 出版地：出版单位，出版年份:起止页码.
4．文献是学位论文时，书写格式为：
[序号] 作者.论文题目[D].保存地：保存单位，年份.
5．文献是来自报告时，书写格式为：
[序号] 报告者.报告题目[R].报告地：报告会主办单位，报告年份.
6．文献是来自专利时，书写格式为：
[序号] 专利所有者.专利名称：专利国别，专利号[P].发布日期.
7．文献是来自国际、国家标准时，书写格式为：
[序号] 标准代号.标准名称[S].出版地：出版单位，出版年份.
8．文献来自报纸文章时，书写格式为：
[序号] 作者.文章题目[N].报纸名，出版日期（版次）.
9．文献来自电子文献时，书写格式为：
[序号] 作者.文献题目[电子文献及载体类型标识].电子文献的可获取地址，发表或更新日期/引用日期（可以只选择一项）.
电子参考文献建议标识：
［DB/OL］——联机网上数据库(database online)
［DB/MT］——磁带数据库(database on magnetic tape)
［M/CD］——光盘图书(monograph on CD-ROM)
［CP/DK］——磁盘软件(computer program on disk)
［J/OL］——网上期刊(serial online)
［EB/OL］——网上电子公告(electronic bulletin board online)

        {text}

        """
    if type:
        response = chat.invoke(base_instruction1)
    else:
        response = chat.invoke(base_instruction2)

    return response.content


def check_abstract_format(para, expected_font_name, expected_font_size_pt):
    """
    检查段落的格式是否符合预期，并在不符合时添加到paper_comments中。
    """
    for run in para.runs:
        run.font.name = expected_font_name
        # 设置字号为12磅
        run.font.size = Pt(expected_font_size_pt)
        if expected_font_name in ['宋体', '华文中宋']:
            run._element.rPr.rFonts.set(qn('w:eastAsia'), expected_font_name)
    return para


def check_title_format(paragraph, level):
    """
    检查标题的格式是否符合特定级别的要求。
    """
    if level == 1:
        font_name = '华文中宋'
        font_size = Pt(22)
        bold = True
        alignment = WD_ALIGN_PARAGRAPH.CENTER
        space_before = Pt(36)
        space_after = Pt(24)
        line_spacing = 1.5

    elif level == 2:
        font_name = '黑体'
        font_size = Pt(15)
        bold = True
        alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # 默认设置为两端对齐
        space_before = Pt(12)
        space_after = Pt(12)
        line_spacing = 1.5

    elif level == 3:
        font_name = '黑体'
        font_size = Pt(12)
        bold = True
        alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # 默认设置为两端对齐
        space_before = Pt(12)
        space_after = Pt(12)
        line_spacing = 1.5
    else:
        return  # 如果不是指定的标题级别，则不进行设置
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = space_before
    paragraph.paragraph_format.space_after = space_after
    paragraph.paragraph_format.line_spacing = line_spacing
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = font_size
        run.font.bold = bold
        # 针对中文设置字体
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    return paragraph


def check_text_format(paragraph):
    """
    检查文本的格式是否符合预期，并在不符合时添加到paper_comments中。
    """
    """
      设置段落为小四号宋体，具有特定的格式要求。
      """
    # 设置段落对齐方式为两端对齐
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # 设置段前段后间距为0
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    # 设置行距为1.5倍行距
    paragraph.paragraph_format.line_spacing = 1.5
    # 设置首行缩进为2个字符，大约等于36磅
    paragraph.paragraph_format.first_line_indent = Pt(36)

    # 遍历段落中的所有运行并设置字体样式
    for run in paragraph.runs:
        # print(run.font.name)
        # 设置字体为宋体
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        # 设置字体大小为12磅
        run.font.size = Pt(12)

    return paragraph


from annotate import add_comments
from llm import *
from tqdm import tqdm

def check_paper(doc, key_sections):
    paper_comments = []
    paper_comments = check_front(doc, key_sections['Abstract'], paper_comments)
    paper_comments, doc = check_abstract(doc, key_sections['Abstract'], key_sections['TextStart'], paper_comments)
    paper_comments, doc = check_text(doc, key_sections['TextStart'], key_sections['References'], paper_comments)
    paper_comments = check_references(doc, key_sections['References'],
                                      key_sections['Appendix'] if key_sections['Appendix'] is not None else
                                      key_sections['Acknowledgements'], paper_comments)

    doc.save('modified_document.docx')
    add_comments(paper_comments,'modified_document.docx')


def check_front(doc, pos, paper_comments):
    keywords = ['本科毕业论文（设计）']  # 包含所有要检查的关键词
    for i in tqdm(range(pos),desc="封面"):
        para = doc.paragraphs[i]
        text = para.text.strip()
        # 检查段落文本是否以任一关键词开头
        if any(text.startswith(kw) for kw in keywords):
            paper_comments.append({'text': text,
                                   'comment1': '这一段的下面的内容应该使用华文仿宋，并且字号为三号，下面应有下划线\n(若已经正确则忽略这条批注)'})
    return paper_comments


def check_abstract(doc, start_pos, end_pos, paper_comments):
    flag = None
    for i in tqdm(range(start_pos, end_pos), desc="摘要"):
        para = doc.paragraphs[i]
        text = para.text.strip()
        if len(text) > 30:
            advice1 = get_correction_suggestion(text, True)
            paper_comments.append({'text': text,
                                   'comment1': advice1})
        if text == '摘要':
            flag = True
            continue
        elif text == 'Abstract':
            flag = False
            continue
        if flag is not None:
            # 假设小四号字体大小约为12磅（Pt）
            if flag:  # 如果是中文摘要
                # advice2 = check_abstract_format(para, '宋体', 12)
                doc.paragraphs[i] = check_abstract_format(para, '宋体', 12)
            else:  # 如果是英文摘要
                doc.paragraphs[i] = check_abstract_format(para, 'Time New Roman', 12)

    return paper_comments, doc


def check_text(doc, start_pos, end_pos, paper_comments):
    for i in tqdm(range(start_pos, end_pos), desc="正文"):
        level = 0
        para = doc.paragraphs[i]
        text = para.text
        style_name = para.style.name

        # 检查是否为一级标题
        if style_name == 'Heading 1':
            level = 1
        # 检查是否为二级标题
        elif style_name == 'Heading 2':
            level = 2
        # 检查是否为三级标题
        elif style_name == 'Heading 3':
            level = 3
        if level != 0:
            doc.paragraphs[i] = check_title_format(para, level)
            continue
        if len(text) > 30:
            doc.paragraphs[i] = check_text_format(para)
            text = para.text.strip()
            advice1 = get_correction_suggestion(text, True)
            paper_comments.append({'text': para.text,
                                   'comment1': advice1})
    return paper_comments, doc


def check_references(doc, start_pos, end_pos, paper_comments):
    for i in tqdm(range(start_pos, end_pos), desc="参考文献"):
        para = doc.paragraphs[i]
        text = para.text
        if len(text) > 30:
            advice = get_correction_suggestion(text, False)
            paper_comments.append({'text': text,
                                   'comment1': advice})
    return paper_comments

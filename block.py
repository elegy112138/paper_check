

def find_key_sections(doc):
    """
    在文档中查找关键部分的位置：摘要、摘要后的第一个一级标题、参考文献、附录和致谢。
    返回一个字典，其中包含这些部分的起始位置。
    """
    positions = {
        'Abstract': None,
        'TextStart': None,
        'References': None,
        'Appendix': None,
        'Acknowledgements': None
    }
    found_abstract = False

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()

        # 查找“摘要”
        if text == '摘要':
            positions['Abstract'] = i
            found_abstract = True

        # “摘要”后的第一个一级标题，标记为正文开始
        elif found_abstract and para.style.name == 'Heading 1' and positions['Acknowledgements'] is None:
            positions['TextStart'] = i
            found_abstract = False  # Reset flag after finding first heading after abstract

        # 查找“参考文献”
        elif text == '参考文献':
            positions['References'] = i

        # 查找“附录”
        elif text.startswith('附录'):  # 附录可能有多个，如“附录A”，“附录B”
            positions['Appendix'] = i  # 这将更新到最后一个“附录”的位置

        # 查找“致谢”
        elif text == '致谢':
            positions['Acknowledgements'] = i

    return positions




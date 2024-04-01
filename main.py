import block
from docx import Document
from check_paper import check_paper

# 示例用法
doc_path = '41805054+基于深度强化学习的异构网络算力分配研究.docx'
document = Document(doc_path)
key_sections = block.find_key_sections(document)
check_paper(document, key_sections)
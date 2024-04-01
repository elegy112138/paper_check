from spire.doc import *
from spire.doc.common import *

# 创建一个 Document 类的对象
doc = Document()

# 加载一个 Word 文件
doc.LoadFromFile("41805054+基于深度强化学习的异构网络算力分配研究.docx")

# 获取特定的章节
section = doc.Sections[3]

# 获取特定的段落
paragraph = section.Paragraphs[3]


print(paragraph.Text)
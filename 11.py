# 假设 get_correction_suggestion 和 add_comment 是已经定义好的函数
from spire.doc import *

doc = Document()
doc.LoadFromFile('41805054+基于深度强化学习的异构网络算力分配研究.docx')
text='雷波等人[1]'
try:
    # 尝试查找要添加评论的文本
    text = doc.FindString(text, True, True)
    print(text)
    # 如果找到了文本，可以在这里调用 add_comment 或其他函数
except Exception as e:
    print(11)

    return

# 创建一个评论对象并设置评论的内容和作者
print(1111)
comment = Comment(doc)
comment.Body.AddParagraph().Text = '1111111111111111'
comment.Format.Author = "AI"

# 将找到的文本作为文本范围，并获取其所属的段落
range = text.GetAsOneRange()
paragraph = range.OwnerParagraph

# 将评论添加到段落中
paragraph.ChildObjects.Insert(paragraph.ChildObjects.IndexOf(range) + 1, comment)

# 创建评论起始标记和结束标记，并将它们设置为创建的评论的起始标记和结束标记
commentStart = CommentMark(doc, CommentMarkType.CommentStart)
commentEnd = CommentMark(doc, CommentMarkType.CommentEnd)
commentStart.CommentId = comment.Format.CommentId
commentEnd.CommentId = comment.Format.CommentId

# 在找到的文本之前和之后插入创建的评论起始和结束标记
paragraph.ChildObjects.Insert(paragraph.ChildObjects.IndexOf(range), commentStart)
paragraph.ChildObjects.Insert(paragraph.ChildObjects.IndexOf(range) + 1, commentEnd)

doc.SaveToFile("批注文本2.docx")
doc.Close()
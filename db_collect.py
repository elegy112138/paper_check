from pymongo import MongoClient

client = MongoClient('mongodb://localhost:50000/')

# 指定数据库
db = client['paper_check']



def get_collection(collection_name):
    return db[collection_name]
# 创建MongoDB客户端连接

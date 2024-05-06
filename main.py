import process
from config import get_collection
from flask import Flask, request, send_from_directory, jsonify
import os
from werkzeug.utils import secure_filename
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import pymongo
from bson.objectid import ObjectId
from check_paper1 import process_file
from flask_jwt_extended import JWTManager, jwt_required, get_jwt_identity,create_access_token
import datetime


executor = ThreadPoolExecutor(max_workers=1)

app = Flask(__name__)

# 设置用于存储上传和处理后的文件的目录
UPLOAD_FOLDER = 'uploads/'
PROCESSED_FOLDER = 'processed/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['JWT_SECRET_KEY'] = '2218023235'
jwt = JWTManager(app)

@app.route('/register', methods=['POST'])
def register():
    data = request.get_json()
    username = data.get('username')
    if username is None:
        return jsonify({"message": "Username is required"}), 400

    password = data.get('password')
    if username is None:
        return jsonify({"message": "password is required"}), 400
    collection = get_collection("user")
    if collection.find_one({'username': username}):
        return jsonify({"message": "Username already exists"}), 400

    user_data = {'username': username, 'password': password}
    collection.insert_one(user_data)
    return jsonify(message="register successfully"), 200

@app.route('/login', methods=['POST'])
def login():
    # 从请求中获取用户名和密码
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    collection = get_collection("user")
    # 在数据库中查找用户信息
    user = collection.find_one({'username': username, 'password': password})

    # 验证用户名和密码是否匹配
    if user:
        user_id = str(user['_id'])  # 获取用户的 _id
        # 生成 token，设置有效期为 2 分钟
        expires = datetime.timedelta(minutes=180)
        access_token = create_access_token(identity=user_id, expires_delta=expires)
        return jsonify(access_token=access_token,message='login successfully'), 200
    else:
        return jsonify({"message": "Invalid username or password"}), 401



def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['docx']


@app.route('/upload', methods=['POST'])
@jwt_required()
def upload_file():
    task_id = get_jwt_identity()
    process.cancel[task_id] = False
    # 'files[]' 是前端用来上传文件的字段名
    # 使用 'file' 而不是 'files[]' 获取文件列表
    file_objs = request.files.getlist('file')
    process.tasks[task_id] = {'total': 0, 'current': 0, 'status': '未开始','id':''}
    if not file_objs:
        return jsonify(error="No file uploaded or selected"), 400

    for file in file_objs:
        # 检查文件名是否符合要求
        if file and allowed_file(file.filename):
            pass
        else:
            return jsonify(error=f"Invalid file format for {file.filename}, only DOCX files are allowed."), 400

    collection=get_collection("paper")
    file_data=[]
    for file in file_objs:
        filename = secure_filename(file.filename)
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
        unique_filename = f"{timestamp}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 将日期时间格式化为字符串
        data={
            'date':date,
            "upload_file_path":file_path,
            "processed_file_path":os.path.join(app.config['PROCESSED_FOLDER'], unique_filename),
            'download_path':f'http://42.193.225.22:5000/static/processed_{unique_filename}',
            'status':1,
            'file_name':file.filename,
            'user_id':task_id
        }
        result = collection.insert_one(data)
        file_data.append({'id':str(result.inserted_id),'file_path':file_path,"unique_filename":unique_filename})
        # 如果文件上传成功，将其传递给异步函数处理
    executor.submit(process_file,file_data,task_id)

    return jsonify(message="Files uploaded successfully"), 200


@app.route('/upload', methods=['GET'])
@jwt_required()
def info():
    id = get_jwt_identity()

    if process.tasks.get(id) is not None:
        return {"data":process.tasks[id]},200
    return {"message":'id不能为空或没有此id对应的任务在进行'},404

@app.route('/fileList', methods=['GET'])
@jwt_required()
def get_file_list():
    data = request.args  # GET 请求通常从 URL 参数中获取数据
    current_user_id = get_jwt_identity()
    collection = get_collection("paper")
    page = int(data.get("page"))
    pagesize = int(data.get("pagesize"))
    total_count = collection.count_documents({"user_id": current_user_id})
    article_data = list(
        collection.find({"user_id":current_user_id}, {"upload_file_path": 0, "processed_file_path": 0}).sort("status", pymongo.ASCENDING).skip(
            (page - 1) * pagesize).limit(pagesize))

    for item in article_data:
        item['_id'] = str(item['_id'])
    # 根据结果返回不同的消息
    if article_data:
        return {"message": "success", "status": 1, "data": article_data, "total": total_count}, 200
    else:
        return {"message": "No paper found", "status": 0, "data": []}, 200

@app.route('/upload', methods=['DELETE'])
@jwt_required()
def delete_file():
    data = request.args  # GET 请求通常从 URL 参数中获取数据
    collection = get_collection("paper")
    _id = data.get("id")
    result=collection.find_one({"_id":ObjectId(_id)})
    if result:
        upload_file_path = result.get('upload_file_path')
        processed_file_path = result.get('processed_file_path')

        if upload_file_path and os.path.exists(upload_file_path):
            os.remove(upload_file_path)

        if processed_file_path and os.path.exists(processed_file_path):
            os.remove(processed_file_path)

        collection.delete_one({"_id": ObjectId(_id)})
        return {"message": "删除成功", "status": 1}, 200

    else:
        return {"message": "没有找到相关文件", "status": 0}, 200

@app.route('/upload', methods=['PUT'])
@jwt_required()
def set_cancel():
    current_user_id = get_jwt_identity()
    process.cancel[current_user_id]=True
    print('取消文件处理')
    return {"message": "取消成功"}, 200



@app.route('/static/<filename>')
def static_files(filename):
    return send_from_directory('processed', filename)


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    if not os.path.exists(PROCESSED_FOLDER):
        os.makedirs(PROCESSED_FOLDER)
    app.run(host='0.0.0.0', port=5000, debug=True)



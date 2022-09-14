import os
import json
from flask import Flask, flash, request, redirect, url_for, send_from_directory, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'static'
ALLOWED_EXTENSIONS = {'pptx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
CORS(app, supports_credentials=False)


def allowed_file(filename):
    '''
    判断文件后缀是否满足上传需求
    :param filename: 文件名
    :return: True or False
    '''
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# 获取周报文件列表
@app.route("/weeklyReports", methods=["GET"])
def weekly_reports():
    report_files = os.listdir('static/weeklyReports')
    result = []
    for file in report_files:
        item = dict({
            "name": str(file)
        })
        result.append(item)
    return jsonify(result)


# 获取月报文件列表
@app.route("/monthlyReports", methods=["GET"])
def monthly_reports():
    report_files = os.listdir('static/monthlyReports')
    result = []
    for file in report_files:
        item = dict({
            "name": str(file)
        })
        result.append(item)
    return jsonify(result)


# 上传周报
@app.route('/weeklyReports/upload', methods=['GET', 'POST'])
def weekly_reports_upload():
    file = request.files.get("file")
    # extra_args = dict(request.form)
    resp_data = {}
    if allowed_file(file.filename):
        if os.path.exists(app.config["UPLOAD_FOLDER"] + "/weeklyReports/" + file.filename):
            resp_data = {
                "msg": file.filename + "文件已存在!"
            }
        else:
            file.save(app.config["UPLOAD_FOLDER"] + "/weeklyReports/" + file.filename)
            resp_data = {
                "msg": file.filename + "文件上传成功!"
            }
    else:
        resp_data = {
            "msg": "只能上传后缀为.pptx的文件"
        }

    return jsonify(resp_data)


# 上传月报
@app.route('/monthlyReports/upload', methods=['GET', 'POST'])
def monthly_reports_upload():
    file = request.files.get("file")
    # extra_args = dict(request.form)
    resp_data = {}
    if allowed_file(file.filename):
        if os.path.exists(app.config["UPLOAD_FOLDER"] + "/monthlyReports/" + file.filename):
            resp_data = {
                "msg": file.filename + "文件已存在!"
            }
        else:
            file.save(app.config["UPLOAD_FOLDER"] + "/monthlyReports/" + file.filename)
            resp_data = {
                "msg": file.filename + "文件上传成功!"
            }
    else:
        resp_data = {
            "msg": "只能上传后缀为.pptx的文件"
        }

    return jsonify(resp_data)


# 下载周报
@app.route('/weeklyReports/download/<filename>', methods=["GET"])
def weekly_reports_download(filename):
    # send_static_file会在static目录下寻找文件
    return app.send_static_file("weeklyReports/" + filename)


# 下载月报
@app.route('/monthlyReports/download/<filename>', methods=["GET"])
def monthly_reports_download(filename):
    # send_static_file会在static目录下寻找文件
    return app.send_static_file("monthlyReports/" + filename)


@app.route("/")
def hello_world():
    return "<p>欢迎来到PPTX后台</p>"


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)

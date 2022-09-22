import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
from pptx import Presentation
from week_func import PresentationBuilder, WeaklyReports
from tool_func import generate_table_data


UPLOAD_FOLDER = 'static'
ALLOWED_EXTENSIONS = {'pptx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
CORS(app, supports_credentials=True)


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


@app.route("/weeklyReportsData", methods=["POST"])
def generate_weekly_report():
    # request.json是一个字典，接收post请求的data数据
    form_data = request.json
    # inspect_data
    inspect_data = []
    for i in form_data.get("inspect")[0].items():
        if i[0] != "index":
            inspect_data.append(i[1])

    # change_data
    change_data = []
    for i in form_data.get("change"):
        tmp = []
        if i.items():
            for j in i.items():
                if j[0] != "index":
                    tmp.append(j[1])
        change_data.append(tmp)

    # release_data
    release_data = []
    for i in form_data.get("release"):
        tmp = []
        if i.items():
            for j in i.items():
                if j[0] != "index":
                    tmp.append(j[1])
        release_data.append(tmp)

    # permissionManagement_data(有序号)
    permission_management_data = []
    for i in form_data.get("permissionManagement"):
        tmp = []
        if i.items():
            for j in i.items():
                if j[0] != "index":
                    tmp.append(j[1])
                else:
                    tmp.insert(0, str(int(j[1]) + 1))
        permission_management_data.append(tmp)

    # cooperation_data(有序号)
    cooperation_data = []
    for i in form_data.get("cooperation"):
        tmp = []
        if i.items():
            for j in i.items():
                if j[0] != "index":
                    tmp.append(j[1])
                else:
                    tmp.insert(0, str(int(j[1]) + 1))
        cooperation_data.append(tmp)

    # problem_data
    problem_data = []
    for i in form_data.get("problem"):
        tmp = []
        if i.items():
            for j in i.items():
                if j[0] != "index":
                    tmp.append(j[1])
        problem_data.append(tmp)

    # workingPlan_data(有序号)
    working_plan_data = []
    for i in form_data.get("workingPlan"):
        tmp = []
        if i.items():
            for j in i.items():
                if j[0] != "index":
                    tmp.append(j[1])
                else:
                    tmp.insert(0, str(int(j[1]) + 1))
        working_plan_data.append(tmp)

    # events_count
    events_count = [
        len(change_data),  # 变更
        len(permission_management_data),  # 资源权限管理
        len(cooperation_data),  # 配合操作
        len(release_data),  # 支撑发版
        len(problem_data),  # 问题和告警处理
        0  # 故障处理
    ]

    # 获取运行情况分析ppt中所需所需数据
    cluster_pie_data, cluster_table_data = generate_table_data()

    # 编写ppt
    # 这种打开方式适合ppt2007及最新，不适合ppt2003及以前。支持stringio/bytesio stream
    prs = Presentation("template.pptx")  # type: pptx.presentation.Presentation # 设置type，会有代码提示
    wr = WeaklyReports(prs)
    # 运维工作统计
    wr.slide_1(events_count)

    # 巡检
    wr.slide_2(inspect_data)

    # 变更
    wr.slide_3(change_data)

    # 支撑发版
    wr.slide_4(release_data)

    # 资源权限管理
    wr.slide_5(permission_management_data)

    # 配合操作
    wr.slide_6(cooperation_data)

    # 问题及告警
    wr.slide_7(problem_data)

    # 运行情况分析
    wr.slide_8(cluster_pie_data, cluster_table_data)
    #
    # 下周工作计划
    wr.slide_9(working_plan_data)

    # 保存pptx
    prs.save("haha.pptx")

    return jsonify({
        "msg": "ok"
    })


@app.route("/")
def hello_world():
    return "<p>欢迎来到PPTX后台</p>"


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)

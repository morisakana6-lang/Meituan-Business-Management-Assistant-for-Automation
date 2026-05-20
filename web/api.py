"""
美团经营宝 Web API 服务

提供 REST API 接口供前端页面调用

=== 启动方式 ===
python api.py

=== API 接口 ===

POST /api/summary/generate
请求体:
{
    "门店列表": ["1933643130", "1732110850"],
    "平台": 0,
    "门店评论平台": 1,
    "日期范围": {
        "begin": "2026-02-01",
        "end": "2026-03-31"
    }
}

响应: Excel文件下载
"""

from flask import Flask, request, jsonify, send_from_directory
import sys
import os

# 添加项目根目录到 path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

app = Flask(__name__, static_folder='.', static_url_path='')


@app.route('/')
def index():
    """返回汇总报表页面"""
    return send_from_directory('.', 'summary.html')


@app.route('/summary.html')
def summary_page():
    """返回汇总报表页面"""
    return send_from_directory('.', 'summary.html')


@app.route('/api/summary/generate', methods=['POST'])
def summary_generate():
    """
    生成汇总报表接口

    请求方法: POST
    Content-Type: application/json

    请求体:
    {
        "门店列表": ["1933643130", "1732110850"],
        "平台": 0,
        "门店评论平台": 1,
        "日期范围": {
            "begin": "2026-02-01",
            "end": "2026-03-31"
        }
    }

    响应: Excel文件下载
    """
    try:
        data = request.get_json()

        if not data:
            return jsonify({
                "status": "error",
                "message": "请求体不能为空"
            }), 400

        # 验证必填字段
        shop_ids = data.get('门店列表', [])
        if not shop_ids:
            return jsonify({
                "status": "error",
                "message": "门店列表不能为空"
            }), 400

        date_range = data.get('日期范围', {})
        begin_date = date_range.get('begin', '')
        end_date = date_range.get('end', '')

        if not begin_date or not end_date:
            return jsonify({
                "status": "error",
                "message": "日期范围不能为空"
            }), 400

        # 调用汇总报表生成器
        from summary_report_generator import generate_summary_report

        filepath = generate_summary_report(data)

        # 生成下载文件名
        is_multi_shop = len(shop_ids) > 1
        shop_label = "多门店" if is_multi_shop else str(shop_ids[0])
        filename = f"美团经营宝汇总_{shop_label}_{begin_date}_{end_date}.xlsx"

        # 读取文件为二进制后返回
        with open(filepath, 'rb') as f:
            file_data = f.read()

        from flask import make_response
        response = make_response(file_data)
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{filename}"
        response.headers['X-Suggested-Filename'] = filename
        return response

    except ValueError as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 400

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"服务器错误: {str(e)}"
        }), 500


if __name__ == '__main__':
    print("=" + "=" * 49)
    print("美团经营宝 Web 服务")
    print("=" + "=" * 49)
    print("访问地址: http://127.0.0.1:5000")
    print("按 Ctrl+C 停止服务")
    print("=" + "=" * 49)
    app.run(host='0.0.0.0', port=5000, debug=False)

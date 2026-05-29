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
    "评论统计平台": 1,
    "评论统计日期范围": 2,
    "日期范围": {
        "begin": "2026-02-01",
        "end": "2026-03-31"
    }
}

响应: Excel文件下载

POST /api/shop_review/generate
请求体:
{
    "门店列表": ["1933643130"],
    "平台": 1,
    "日期范围": {
        "begin": "2026-05-01",
        "end": "2026-05-20"
    }
}

响应: Excel文件下载
"""

from flask import Flask, request, jsonify, send_from_directory
import sys
import os
import json

# 添加项目根目录到 path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

app = Flask(__name__)


@app.route('/')
def index():
    """返回汇总报表页面"""
    return send_from_directory('.', 'summary.html')


@app.route('/summary.html')
def summary_page():
    """返回汇总报表页面"""
    return send_from_directory('.', 'summary.html')


@app.route('/tuiguangtong/<path:filename>')
def tuiguangtong_files(filename):
    """返回 tuiguangtong 目录下的文件"""
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    file_path = os.path.join(base_dir, 'tuiguangtong', filename)
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read(), 200, {'Content-Type': 'application/json'}


@app.route('/api/summary/generate', methods=['POST'])
def summary_generate():
    """
    生成汇总报表接口（4个Sheet：推广通、全站推广、简版看板、客流分析）

    请求方法: POST
    Content-Type: application/json

    请求体:
    {
        "门店列表": ["1933643130", "1732110850"],
        "平台": 0,
        "评论统计平台": 1,
        "评论统计日期范围": 2,
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
        # 使用 RFC 5987 编码处理中文文件名
        from urllib.parse import quote
        encoded_filename = quote(filename, safe='')
        response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{encoded_filename}"
        # X-Suggested-Filename 也需要 RFC 5987 编码，否则 latin-1 无法编码中文
        response.headers['X-Suggested-Filename'] = encoded_filename
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


@app.route('/api/shops', methods=['GET'])
def get_shops():
    """
    获取门店列表接口

    请求方法: GET
    Query参数:
        page: 页码，默认1
        page_size: 每页数量，默认100

    响应: JSON
    """
    try:
        page = int(request.args.get('page', 1))
        page_size = int(request.args.get('page_size', 100))

        if page < 1:
            page = 1
        if page_size < 1 or page_size > 500:
            page_size = 100

        # 读取 shop_mapping.json
        mapping_file = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'tuiguangtong', 'shop_mapping.json'
        )

        if not os.path.exists(mapping_file):
            return jsonify({
                "status": "error",
                "message": "门店映射文件不存在，请先更新门店数据"
            }), 404

        with open(mapping_file, 'r', encoding='utf-8') as f:
            mapping_data = json.load(f)

        shops = mapping_data.get('shops', [])
        total = len(shops)
        total_pages = (total + page_size - 1) // page_size

        # 分页
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        page_shops = shops[start_idx:end_idx]

        # 格式化返回数据
        result_shops = []
        for shop in page_shops:
            meituan_id = ''
            dianping_id = ''
            for id_info in shop.get('ids', []):
                if id_info.get('type') == 'meituan':
                    meituan_id = id_info.get('id', '')
                elif id_info.get('type') == 'dianping':
                    dianping_id = id_info.get('id', '')

            result_shops.append({
                'name': shop.get('name', ''),
                'city': shop.get('city', ''),
                'meituan_id': meituan_id,
                'dianping_id': dianping_id
            })

        return jsonify({
            "status": "success",
            "shops": result_shops,
            "total": total,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"服务器错误: {str(e)}"
        }), 500


@app.route('/api/shops/refresh', methods=['POST'])
def refresh_shops():
    """
    更新门店数据接口

    从美团API获取最新门店数据并更新 shop_mapping.json

    响应: JSON
    """
    try:
        from shop_info.client import ShopInfoClient

        client = ShopInfoClient()

        # 获取全量门店数据
        shops = client.fetch_all_shops(delay=0.3)

        if not shops:
            return jsonify({
                "status": "error",
                "message": "获取门店数据失败，请检查凭证是否有效"
            }), 500

        # 保存到 shop_mapping.json
        client.save_mapping(shops)

        return jsonify({
            "status": "success",
            "message": f"更新成功，共 {len(shops)} 家门店"
        })

    except FileNotFoundError as e:
        return jsonify({
            "status": "error",
            "message": f"凭证文件不存在: {str(e)}"
        }), 500

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"更新失败: {str(e)}"
        }), 500


@app.route('/api/shop_review/generate', methods=['POST'])
def shop_review_generate():
    """
    生成门店评论报表接口

    请求方法: POST
    Content-Type: application/json

    请求体:
    {
        "门店列表": ["1933643130"],
        "平台": 1,
        "日期范围": {
            "begin": "2026-05-01",
            "end": "2026-05-20"
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

        platform = data.get('平台', 1)

        # 调用门店评论报表生成器
        from shop_review.client import ShopReviewClient
        from shop_review.excel_generator import generate_report

        client = ShopReviewClient()
        filepath = generate_report(
            client=client,
            shop_ids=shop_ids,
            platform=platform,
            begin_date=begin_date,
            end_date=end_date
        )

        # 生成下载文件名
        is_multi_shop = len(shop_ids) > 1
        shop_label = "多门店" if is_multi_shop else str(shop_ids[0])
        filename = f"门店评论_{shop_label}_{begin_date}_{end_date}.xlsx"

        # 读取文件为二进制后返回
        with open(filepath, 'rb') as f:
            file_data = f.read()

        from flask import make_response
        response = make_response(file_data)
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        # 使用 RFC 5987 编码处理中文文件名
        from urllib.parse import quote
        encoded_filename = quote(filename, safe='')
        response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{encoded_filename}"
        # X-Suggested-Filename 也需要 RFC 5987 编码，否则 latin-1 无法编码中文
        response.headers['X-Suggested-Filename'] = encoded_filename
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

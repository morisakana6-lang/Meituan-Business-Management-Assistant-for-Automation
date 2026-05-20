"""
美团经营宝 Web API 服务

提供 REST API 接口供前端页面调用

=== 启动方式 ===
python api.py

=== API 接口 ===

POST /api/query
请求体:
{
    "search_key": "764510450",      # 门店ID或名称
    "platform": 0,                   # 0=美团, 1=点评
    "begin_date": "2026-05-01",     # 开始日期
    "end_date": "2026-05-15"        # 结束日期
}

响应:
{
    "status": "success",
    "data": {
        "shop_name": "丽减美瘦吧·减肥瘦身(新澳城店)",
        "shop_id": "764510450",
        "total_cost": 1234.56,
        "total_clicks": 5678,
        "total_exposure": 123456,
        "cpc": 0.22,
        "rows": [
            {"date": "05-01", "cost": "12.50", ...},
            ...
        ]
    }
}

错误响应:
{
    "status": "error",
    "message": "错误描述"
}
"""

from flask import Flask, request, jsonify, send_from_directory
import sys
import os

# 添加项目根目录到 path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tuiguangtong.shop_search import resolve_shop
from tuiguangtong.client import TuiguangtongClient
from tuiguangtong.excel_generator import ExcelGenerator
from datetime import datetime, timedelta

app = Flask(__name__, static_folder='.', static_url_path='')


def safe_float(val, default=0.0):
    """安全转换为浮点数，处理非数字字符串如'次日更新'"""
    try:
        return float(str(val).replace(',', ''))
    except (ValueError, TypeError):
        return default


def safe_int(val, default=0):
    """安全转换为整数，处理非数字字符串如'次日更新'"""
    try:
        return int(float(str(val).replace(',', '')))
    except (ValueError, TypeError):
        return default


@app.route('/')
def index():
    """返回前端页面"""
    return send_from_directory('.', 'index.html')


@app.route('/api/query', methods=['POST'])
def query():
    """
    查询数据接口

    请求方法: POST
    Content-Type: application/json

    请求体:
    {
        "search_key": "764510450" | "门店名称",
        "platform": 0 | 1,
        "begin_date": "2026-05-01",
        "end_date": "2026-05-15"
    }

    响应:
    {
        "status": "success",
        "data": {
            "shop_name": "丽减美瘦吧·减肥瘦身(新澳城店)",
            "shop_id": "764510450",
            "total_cost": 1234.56,
            "total_clicks": 5678,
            "total_exposure": 123456,
            "cpc": 0.22,
            "rows": [{"date": "汇总", "cost": "1234.56", ...}]
        }
    }
    """
    try:
        data = request.get_json()

        if not data:
            return jsonify({
                "status": "error",
                "message": "请求体不能为空"
            }), 400

        search_key = data.get('search_key', '').strip()
        platform = int(data.get('platform', 0))
        begin_date = data.get('begin_date', '').strip()
        end_date = data.get('end_date', '').strip()

        # 参数验证
        if not search_key:
            return jsonify({
                "status": "error",
                "message": "search_key 不能为空"
            }), 400

        if not begin_date or not end_date:
            return jsonify({
                "status": "error",
                "message": "begin_date 和 end_date 不能为空"
            }), 400

        # 解析门店
        shop_name, shop_id = resolve_shop(search_key, platform)

        # 创建客户端并获取数据
        client = TuiguangtongClient()
        generator = ExcelGenerator(shop_name=shop_name, shop_id=shop_id)
        generator._create_header()

        # 按日查询
        start = datetime.strptime(begin_date, "%Y-%m-%d")
        end = datetime.strptime(end_date, "%Y-%m-%d")
        current = start

        while current <= end:
            date_str = current.strftime("%Y-%m-%d")
            date_display = current.strftime("%m-%d")

            try:
                metrics = client.get_metrics(date_str, date_str, shop_id)
                generator.add_data_row(date_display, metrics)
            except Exception as e:
                generator.add_data_row(date_display, [])

            current += timedelta(days=1)

        # 准备结果数据 - 计算所有指标的汇总
        total_cost = 0
        total_cash_cost = 0
        total_exposure = 0
        total_clicks = 0
        total_orders = 0
        total_group_orders = 0
        total_browse = 0
        total_interested = 0

        for row_data in generator.data_rows:
            cost = safe_float(row_data.get('T30001', '0'))
            cash_cost = safe_float(row_data.get('T30047', '0'))
            clicks = safe_int(row_data.get('T30003', '0'))
            exposure = safe_int(row_data.get('T30002', '0'))
            orders = safe_int(row_data.get('T30049', '0'))
            group_orders = safe_int(row_data.get('T30020', '0'))
            browse = safe_int(row_data.get('T30005', '0'))
            interested = safe_int(row_data.get('T30083', '0'))

            total_cost += cost
            total_cash_cost += cash_cost
            total_clicks += clicks
            total_exposure += exposure
            total_orders += orders
            total_group_orders += group_orders
            total_browse += browse
            total_interested += interested

        # 计算平均 CPC 和点击率
        cpc = f"{total_cost / total_clicks:.2f}" if total_clicks > 0 else "0.00"
        click_rate = f"{total_clicks / total_exposure * 100:.2f}" if total_exposure > 0 else "0.00"

        # 只返回汇总行
        summary_row = {
            'date': '汇总',
            'cost': f"{total_cost:.2f}",
            'cash_cost': f"{total_cash_cost:.2f}",
            'exposure': str(total_exposure),
            'clicks': str(total_clicks),
            'click_rate': click_rate + '%',
            'cpc': cpc,
            'group_orders': str(total_group_orders),
            'orders': str(total_orders),
            'browse': str(total_browse),
            'interested': str(total_interested)
        }

        return jsonify({
            "status": "success",
            "data": {
                "shop_name": shop_name,
                "shop_id": shop_id,
                "total_cost": f"{total_cost:.2f}",
                "total_clicks": total_clicks,
                "total_exposure": total_exposure,
                "cpc": cpc,
                "rows": [summary_row]
            }
        })

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


@app.route('/api/download', methods=['POST'])
def download():
    """
    下载Excel报表接口（只含汇总行）

    请求方法: POST
    Content-Type: application/json

    请求体:
    {
        "search_key": "764510450" | "门店名称",
        "platform": 0 | 1,
        "begin_date": "2026-05-01",
        "end_date": "2026-05-15"
    }

    响应: Excel文件下载（只有一行汇总）
    """
    import tempfile
    import os

    try:
        data = request.get_json()

        if not data:
            return jsonify({
                "status": "error",
                "message": "请求体不能为空"
            }), 400

        search_key = data.get('search_key', '').strip()
        platform = int(data.get('platform', 0))
        begin_date = data.get('begin_date', '').strip()
        end_date = data.get('end_date', '').strip()

        # 参数验证
        if not search_key:
            return jsonify({
                "status": "error",
                "message": "search_key 不能为空"
            }), 400

        if not begin_date or not end_date:
            return jsonify({
                "status": "error",
                "message": "begin_date 和 end_date 不能为空"
            }), 400

        # 解析门店
        shop_name, shop_id = resolve_shop(search_key, platform)

        # 创建客户端并获取数据
        client = TuiguangtongClient()
        generator = ExcelGenerator(shop_name=shop_name, shop_id=shop_id)
        generator._create_header()

        # 计算所有指标汇总
        total_cost = 0
        total_cash_cost = 0
        total_exposure = 0
        total_clicks = 0
        total_orders = 0
        total_group_orders = 0
        total_browse = 0
        total_interested = 0

        # 按日查询
        start = datetime.strptime(begin_date, "%Y-%m-%d")
        end = datetime.strptime(end_date, "%Y-%m-%d")
        current = start

        while current <= end:
            date_str = current.strftime("%Y-%m-%d")

            try:
                metrics = client.get_metrics(date_str, date_str, shop_id)
                # 累加各指标
                for m in metrics:
                    mid = m.get("id", "")
                    val = safe_float(m.get("value", "0"))

                    if mid == "T30001":
                        total_cost += val
                    elif mid == "T30047":
                        total_cash_cost += val
                    elif mid == "T30002":
                        total_exposure += int(val)
                    elif mid == "T30003":
                        total_clicks += int(val)
                    elif mid == "T30049":
                        total_orders += int(val)
                    elif mid == "T30020":
                        total_group_orders += int(val)
                    elif mid == "T30005":
                        total_browse += int(val)
                    elif mid == "T30083":
                        total_interested += int(val)
            except Exception as e:
                pass

            current += timedelta(days=1)

        # 计算 CPC
        cpc = f"{total_cost / total_clicks:.2f}" if total_clicks > 0 else "0.00"

        # 汇总数据
        totals = {
            "cost": total_cost,
            "cash_cost": total_cash_cost,
            "exposure": total_exposure,
            "clicks": total_clicks,
            "cpc": cpc,
            "orders": total_orders,
            "group_orders": total_group_orders,
            "browse": total_browse,
            "interested": total_interested
        }

        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        # 清理文件名中的非法字符
        safe_name = "".join(c for c in shop_name if c.isalnum() or c in (' ', '-', '_', '(', ')'))
        filename = f"报表_{safe_name}_{begin_date}_{end_date}.xlsx"
        filepath = os.path.join(temp_dir, filename)

        # 保存Excel（只含汇总行）
        generator.save_summary_only(filepath, totals)

        return send_from_directory(
            temp_dir,
            filename,
            as_attachment=True,
            download_name=filename
        )

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
    print("=" * 50)
    print("美团经营宝 Web 服务")
    print("=" * 50)
    print("访问地址: http://127.0.0.1:5000")
    print("按 Ctrl+C 停止服务")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)

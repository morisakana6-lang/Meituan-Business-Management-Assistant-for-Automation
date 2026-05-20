"""
全站推广 API 测试脚本

测试新的 API 端点：https://e.dianping.com/shopdiy/report/datareport/ajax/cpcwsp/queryTabDataList
"""

import requests
import json
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from common import load_credentials


def test_quanzhan_api():
    """测试全站推广 API"""

    # 加载凭证
    cookie, mtgsig = load_credentials()

    # 构造请求
    url = "https://e.dianping.com/shopdiy/report/datareport/ajax/cpcwsp/queryTabDataList"

    params = {
        "beginDate": "2026-05-01",
        "endDate": "2026-05-14",
        "timeUnit": "day",
        "clientKey": "cpc.wsp.data.common",
        "compareEnabled": "0",
        "compareBeginDate": "2026-04-17",
        "compareEndDate": "2026-04-30",
        "groupUnit": "",
        "objectUnit": "",
        "tabIds": "T30001,T30047,T500001,T600008,T600009",
        "platform": "0",
        "smartShopFlag": "",
        "reportFunctionType": "",
        "shopIds": "0",
        "reportScene": "wsp_1_3",
        "yodaReady": "h5",
        "csecplatform": "4",
        "csecversion": "4.2.0",
        "mtgsig": mtgsig,
    }

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e.dianping.com/app/peon-isomorph-full-site/html/data-report-pc.html",
        "Cookie": cookie,
    }

    print("=" * 60)
    print("全站推广 API 测试")
    print("=" * 60)
    print(f"请求URL: {url}")
    print(f"参数: {json.dumps(params, indent=2, ensure_ascii=False)}")
    print("-" * 60)

    try:
        response = requests.get(url, params=params, headers=headers, timeout=30)
        print(f"状态码: {response.status_code}")
        print(f"响应头: {dict(response.headers)}")
        print("-" * 60)

        data = response.json()
        print(f"响应数据: {json.dumps(data, indent=2, ensure_ascii=False)[:3000]}")

        # 检查返回的数据结构
        print("\n" + "=" * 60)
        print("数据结构分析")
        print("=" * 60)

        if "code" in data:
            print(f"code: {data.get('code')}")
        if "msg" in data:
            msg = data.get("msg")
            if isinstance(msg, dict):
                print(f"msg 包含字段: {list(msg.keys())}")
                if "data" in msg:
                    data_field = msg.get("data")
                    if isinstance(data_field, dict):
                        print(f"data 包含字段: {list(data_field.keys())}")
                    elif isinstance(data_field, list):
                        print(f"data 是列表，长度: {len(data_field)}")
                        if data_field:
                            print(f"第一项: {json.dumps(data_field[0], ensure_ascii=False)[:500]}")
            elif isinstance(msg, list):
                print(f"msg 是列表，长度: {len(msg)}")
            else:
                print(f"msg: {str(msg)[:500]}")

        return data

    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"JSON解析失败: {e}")
        print(f"原始响应: {response.text[:1000]}")
        return None


if __name__ == "__main__":
    test_quanzhan_api()

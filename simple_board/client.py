"""
简版看板数据获取客户端

=== 板块归属 ===
简版看板板块，对应目录 simple_board/

=== API 信息 ===
URL: https://e.dianping.com/gateway/adviser/data
方法: GET

=== 指标分类 ===
概览数据 (simpleDataBoardOverview):
- 曝光人数
- 访问人数
- 下单券数
- 下单金额（原价）
- 核销券数
- 核销金额（原价）

交易数据 (simpleDataBoardTradeSummaryPC):
- 下单券数
- 下单金额（原价）
- 核销券数
- 核销金额（原价）

流量数据 (simpleDataBoardFlowSummaryPC):
- 曝光次数
- 曝光人数
- 访问次数
- 访问人数

=== 请求参数 ===
- source: 固定值 "1"
- device: 固定值 "pc"
- date: 日期范围，格式 "YYYY-MM-DD,YYYY-MM-DD"
- platform: 平台，0=全平台，1=点评，2=美团
- pageType: 固定值 "simpleDataBoard"
- shopIds: 门店ID
- yodaReady: 固定值 "h5"
- csecplatform: 固定值 "4"
- csecversion: 固定值 "4.2.0"
- mtgsig: 认证参数

=== 响应格式 ===
{
    "code": 200,
    "msg": "成功",
    "data": [
        {
            "uniqueMarker": "simpleDataBoardOverview",
            "templateId": "optionSummary",
            "body": [
                {"name": "曝光人数", "value": "78", "suffix": "人"},
                {"name": "访问人数", "value": "13", "suffix": "人"},
                ...
            ]
        },
        ...
    ]
}

=== 使用方法 ===
from simple_board.client import SimpleBoardClient

client = SimpleBoardClient()
data = client.get_metrics(
    begin_date="2026-05-01",
    end_date="2026-05-14",
    shop_id="1027064882511291",
    platform=0
)
"""

import requests
import json
import random
import time
from datetime import datetime, timedelta
from typing import List, Dict, Optional

import sys
import os

# 获取项目根目录
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)  # simple_board 的父目录就是项目根目录
sys.path.insert(0, _PROJECT_ROOT)

# 直接读取凭证文件，避免导入 playwright
def load_credentials():
    """直接从凭证文件读取"""
    cred_file = os.path.join(_PROJECT_ROOT, 'auth', 'credentials.json')
    if os.path.exists(cred_file):
        with open(cred_file, encoding='utf-8') as f:
            data = json.load(f)
            return data.get('cookie', ''), data.get('mtgsig', '')
    return '', ''

# 复用推广通的门店搜索模块
_SHOP_SEARCH_PATH = os.path.join(_PROJECT_ROOT, 'tuiguangtong', 'shop_search.py')
sys.path.insert(0, os.path.dirname(_SHOP_SEARCH_PATH))
from tuiguangtong.shop_search import resolve_shop


class SimpleBoardClient:
    """简版看板报表数据客户端"""

    # 简版看板 API 地址
    BASE_URL = "https://e.dianping.com/gateway/adviser/data"

    # 默认请求头
    DEFAULT_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e.dianping.com/app/peon-isomorph-full-site/html/data-report-pc.html",
    }

    # 概览指标名称映射
    OVERVIEW_METRIC_NAMES = {
        "曝光人数": "exposure_count",
        "访问人数": "visitor_count",
        "下单券数": "order_count",
        "下单金额（原价）": "order_amount",
        "核销券数": "verify_count",
        "核销金额（原价）": "verify_amount",
    }

    # 交易指标名称映射
    TRADE_METRIC_NAMES = {
        "下单券数": "order_count",
        "下单金额（原价）": "order_amount",
        "核销券数": "verify_count",
        "核销金额（原价）": "verify_amount",
    }

    # 流量指标名称映射
    FLOW_METRIC_NAMES = {
        "曝光次数": "exposure_times",
        "曝光人数": "exposure_count",
        "访问次数": "visit_times",
        "访问人数": "visitor_count",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def get_report(
        self,
        begin_date: str,
        end_date: str,
        shop_id: str,
        platform: str = "0",
        max_retries: int = 3,
    ) -> Dict:
        """
        获取简版看板报表数据

        Args:
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            max_retries: 最大重试次数

        Returns:
            API响应字典
        """
        date_range = f"{begin_date},{end_date}"

        params = {
            "source": "1",
            "device": "pc",
            "date": date_range,
            "platform": platform,
            "pageType": "simpleDataBoard",
            "optionType": "",
            "shopIds": shop_id,
            "excludeShopIds": "",
            "cityId": "",
            "prdIds": "",
            "spuId": "",
            "pageNum": "",
            "pageSize": "",
            "sign": "",
            "fromPage": "",
            "yodaReady": "h5",
            "csecplatform": "4",
            "csecversion": "4.2.0",
            "mtgsig": self.mtgsig,
        }

        self.session.headers.update({"Cookie": self.cookie})

        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    delay = random.uniform(2, 5)
                    print(f"  请求失败，{delay:.1f}秒后重试... (第{attempt + 1}次)")
                    time.sleep(delay)

                response = self.session.get(self.BASE_URL, params=params, timeout=30)
                data = response.json()

                if data.get("code") == 200:
                    return data

                # 401/403 错误时打印警告
                if data.get("code") in [401, 403]:
                    print(f"  认证错误 (code={data.get('code')})，请更新凭证")

                return data

            except requests.exceptions.RequestException as e:
                if attempt == max_retries - 1:
                    raise
                print(f"  网络错误: {e}")

        return {"code": -1, "msg": "重试次数耗尽"}

    def get_metrics(
        self,
        begin_date: str,
        end_date: str,
        shop_id: str,
        platform: str = "0",
        search_key: str = None,
    ) -> Dict:
        """
        获取简版看板指标数据

        Args:
            begin_date: 开始日期
            end_date: 结束日期
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            search_key: 门店搜索关键字（ID或名称），会覆盖shop_id

        Returns:
            指标字典，包含 overview, trade, flow 三类数据

        Raises:
            Exception: API返回错误时抛出
        """
        # 如果提供了search_key，通过门店搜索解析
        if search_key:
            _, resolved_shop_id = resolve_shop(search_key, int(platform))
            shop_id = resolved_shop_id

        data = self.get_report(
            begin_date=begin_date,
            end_date=end_date,
            shop_id=shop_id,
            platform=platform,
        )

        if data.get("code") != 200:
            raise Exception(f"API返回错误: code={data.get('code')}, msg={data.get('msg')}")

        result = {
            "overview": {},
            "trade": {},
            "flow": {},
        }

        # 按 uniqueMarker 分类解析数据
        for component in data.get("data", []):
            marker = component.get("uniqueMarker", "")
            body = component.get("body", {})

            # 处理 body 为列表的情况
            if isinstance(body, list):
                items = body
            elif isinstance(body, dict):
                items = []
            else:
                items = []

            if marker == "simpleDataBoardOverview":
                # 概览数据
                for item in items:
                    name = item.get("name", "")
                    value = item.get("value", "0")
                    key = self.OVERVIEW_METRIC_NAMES.get(name, name)
                    result["overview"][key] = value

            elif marker == "simpleDataBoardTradeSummaryPC":
                # 交易数据
                for item in items:
                    name = item.get("name", "")
                    value = item.get("value", "0")
                    key = self.TRADE_METRIC_NAMES.get(name, name)
                    result["trade"][key] = value

            elif marker == "simpleDataBoardFlowSummaryPC":
                # 流量数据
                for item in items:
                    name = item.get("name", "")
                    value = item.get("value", "0")
                    key = self.FLOW_METRIC_NAMES.get(name, name)
                    result["flow"][key] = value

        return result


if __name__ == "__main__":
    """测试简版看板API"""
    print("=" * 50)
    print("简版看板 API 测试")
    print("=" * 50)

    client = SimpleBoardClient()

    # 测试获取指标数据
    try:
        metrics = client.get_metrics(
            begin_date="2026-05-01",
            end_date="2026-05-14",
            shop_id="1027064882511291",
        )

        print("\n获取到的指标:")
        print(f"  概览数据: {metrics.get('overview')}")
        print(f"  交易数据: {metrics.get('trade')}")
        print(f"  流量数据: {metrics.get('flow')}")

    except Exception as e:
        print(f"错误: {e}")
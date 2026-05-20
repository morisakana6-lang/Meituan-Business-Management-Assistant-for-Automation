"""
客流分析数据获取客户端

=== 板块归属 ===
客流分析板块，对应目录 customer_flow/

=== API 信息 ===
URL: https://e.dianping.com/gateway/adviser/data
方法: GET

=== 指标分类 ===
变化趋势 (flowDataSummaryPC):
- 曝光人数
- 曝光次数
- 访问人数
- 访问次数
- 曝光访问转化率
- 意向转化人数
- 意向转化率
- 下单人数
- 留资人数
- 累计收藏人数
- 新增收藏人数
- 新增打卡人数

=== 请求参数 ===
- pageType: 固定值 "flowAnalysis"
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
            "uniqueMarker": "flowDataSummaryPC",
            "templateId": "optionSummary",
            "body": {
                "data": [
                    {"name": "曝光人数", "value": "2,149", "suffix": "人", ...},
                    ...
                ]
            }
        },
        ...
    ]
}

=== 使用方法 ===
from customer_flow.client import CustomerFlowClient

client = CustomerFlowClient()
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
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
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


class CustomerFlowClient:
    """客流分析报表数据客户端"""

    # 客流分析 API 地址
    BASE_URL = "https://e.dianping.com/gateway/adviser/data"

    # 默认请求头
    DEFAULT_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e.dianping.com/app/peon-isomorph-full-site/html/data-report-pc.html",
    }

    # 指标名称映射（英文键名）
    METRIC_NAMES = {
        "曝光人数": "exposure_count",
        "曝光次数": "exposure_times",
        "访问人数": "visitor_count",
        "访问次数": "visit_times",
        "曝光访问转化率": "exposure_visit_rate",
        "意向转化人数": "intention_count",
        "意向转化率": "intention_rate",
        "下单人数": "order_count",
        "留资人数": "retention_count",
        "累计收藏人数": "favorites_count",
        "新增收藏人数": "new_favorites_count",
        "新增打卡人数": "new_checkin_count",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def get_report(
        self,
        shop_id: str,
        platform: str = "0",
        begin_date: str = None,
        end_date: str = None,
        max_retries: int = 3,
    ) -> Dict:
        """
        获取客流分析报表数据

        Args:
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD（可选，默认近7天）
            end_date: 结束日期，格式 YYYY-MM-DD（可选，默认今天）
            max_retries: 最大重试次数

        Returns:
            API响应字典
        """
        from datetime import datetime, timedelta
        # 默认近7天
        if end_date is None:
            end_date = datetime.now().strftime("%Y-%m-%d")
        if begin_date is None:
            begin_date = (datetime.now() - timedelta(days=6)).strftime("%Y-%m-%d")
        date_range = f"{begin_date},{end_date}"

        params = {
            "source": "1",
            "device": "pc",
            "date": date_range,
            "platform": platform,
            "pageType": "flowAnalysis",
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

                response = self.session.post(self.BASE_URL, data=params, timeout=30)
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
        shop_id: str,
        platform: str = "0",
        begin_date: str = None,
        end_date: str = None,
    ) -> Dict:
        """
        获取客流分析指标数据

        Args:
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD（可选）
            end_date: 结束日期，格式 YYYY-MM-DD（可选）

        Returns:
            指标字典

        Raises:
            Exception: API返回错误时抛出
        """
        data = self.get_report(
            shop_id=shop_id,
            platform=platform,
            begin_date=begin_date,
            end_date=end_date,
        )

        if data.get("code") != 200:
            raise Exception(f"API返回错误: code={data.get('code')}, msg={data.get('msg')}")

        result = {}

        # 按 uniqueMarker 查找 flowDataSummaryPC
        for component in data.get("data", []):
            marker = component.get("uniqueMarker", "")

            if marker == "flowDataSummaryPC":
                body = component.get("body", {})
                items = body.get("data", [])

                for item in items:
                    name = item.get("name", "")
                    value = item.get("value", "0")
                    suffix = item.get("suffix", "")
                    key = self.METRIC_NAMES.get(name, name)
                    result[key] = value
                    # 同时保存suffix方便后续处理
                    if suffix:
                        result[f"{key}_suffix"] = suffix

                break

        return result


if __name__ == "__main__":
    """测试客流分析API"""
    print("=" * 50)
    print("客流分析 API 测试")
    print("=" * 50)

    client = CustomerFlowClient()

    # 测试参数
    shop_id = "1027064882511291"

    print(f"\n测试参数:")
    print(f"  门店ID: {shop_id}")
    print()

    try:
        # 获取报表数据
        print("正在请求 API...")
        data = client.get_report(
            shop_id=shop_id,
            platform="0"
        )

        print(f"API 返回 code: {data.get('code')}")
        print(f"API 返回 msg: {data.get('msg')}")
        print()

        # 解析并显示数据
        metrics = client.get_metrics(
            shop_id=shop_id,
            platform="0"
        )

        print("获取到的指标:")
        for key, value in metrics.items():
            if not key.endswith("_suffix"):
                print(f"  {key}: {value}")

    except Exception as e:
        print(f"错误: {e}")

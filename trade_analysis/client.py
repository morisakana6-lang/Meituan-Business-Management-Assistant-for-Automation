"""
交易分析数据获取客户端

=== 板块归属 ===
交易分析板块，对应目录 trade_analysis/

=== API 信息 ===
URL: https://e.dianping.com/mda/v5/trade
方法: POST

=== 指标分类 ===
下单环节:
- 下单人数
- 下单券数
- 下单金额（原价）
- 下单金额

核销环节:
- 核销人数
- 核销券数
- 核销金额（原价）
- 核销金额

退款环节:
- 退款券数
- 退款金额（原价）

=== 请求参数（form data）===
- source: 固定值 "1"
- device: 固定值 "pc"
- date: 日期范围，格式 "YYYY-MM-DD,YYYY-MM-DD"
- platform: 平台，0=全平台，1=点评，2=美团
- pageType: 固定值 "v5Trade"
- optionType: 固定值 "v5Trade"
- shopIds: 门店ID
- prdIds: 产品类型ID，"1,2,3,4,5,6,11,12,13,14,15,16,17,18,19,20"
- fromPage: 固定值 "storeKey"
- timeStamp: 时间戳

=== 响应格式 ===
{
    "code": 0,
    "success": true,
    "data": [{
        "componentId": "tradeOptionSummaryPC",
        "body": [
            {"name": "下单人数", "value": "10", "suffix": "人", ...},
            ...
        ]
    }]
}

=== 使用方法 ===
from trade_analysis.client import TradeAnalysisClient

client = TradeAnalysisClient()
data = client.get_metrics(
    shop_id="1933643130",
    begin_date="2026-05-01",
    end_date="2026-05-14"
)
"""

import requests
import json
import random
import time
from datetime import datetime, timedelta
from typing import Dict, Optional

import sys
import os

# 获取项目根目录
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)


def load_credentials():
    """直接从凭证文件读取"""
    cred_file = os.path.join(_PROJECT_ROOT, 'auth', 'credentials.json')
    if os.path.exists(cred_file):
        with open(cred_file, encoding='utf-8') as f:
            data = json.load(f)
            return data.get('cookie', ''), data.get('mtgsig', '')
    return '', ''


class TradeAnalysisClient:
    """交易分析报表数据客户端"""

    # 交易分析 API 地址
    BASE_URL = "https://e.dianping.com/mda/v5/trade"

    # 产品类型ID（交易分析需要）
    PRD_IDS = "1,2,3,4,5,6,11,12,13,14,15,16,17,18,19,20"

    # 默认请求头
    DEFAULT_HEADERS = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": "https://h5.dianping.com",
        "Referer": "https://h5.dianping.com/",
        "Sec-Ch-Ua": "\"Chromium\";v=\"148\", \"Google Chrome\";v=\"148\", \"Not/A)Brand\";v=\"99\"",
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": "\"Windows\"",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    }

    # 指标名称映射（中文到英文键名）
    METRIC_NAMES = {
        "下单人数": "order_person_count",
        "下单券数": "order_ticket_count",
        "下单金额（原价）": "order_original_amount",
        "下单金额": "order_amount",
        "核销人数": "redeem_person_count",
        "核销券数": "redeem_ticket_count",
        "核销金额（原价）": "redeem_original_amount",
        "核销金额": "redeem_amount",
        "退款券数": "refund_ticket_count",
        "退款金额（原价）": "refund_original_amount",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def _get_timestamp(self) -> str:
        """获取当前时间戳（毫秒）"""
        return str(int(datetime.now().timestamp() * 1000))

    def get_report(
        self,
        shop_id: str,
        platform: str = "0",
        begin_date: str = None,
        end_date: str = None,
        max_retries: int = 3,
    ) -> Dict:
        """
        获取交易分析报表数据

        Args:
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD（可选，默认近7天）
            end_date: 结束日期，格式 YYYY-MM-DD（可选，默认今天）
            max_retries: 最大重试次数

        Returns:
            API响应字典
        """
        # 默认近7天
        if end_date is None:
            end_date = datetime.now().strftime("%Y-%m-%d")
        if begin_date is None:
            begin_date = (datetime.now() - timedelta(days=6)).strftime("%Y-%m-%d")

        date_range = f"{begin_date},{end_date}"
        timestamp = self._get_timestamp()

        # URL参数
        params = {
            "yodaReady": "h5",
            "csecplatform": "4",
            "csecversion": "4.2.0",
            "mtgsig": self.mtgsig,
        }

        # 请求体数据（form data）
        data = {
            "source": "1",
            "device": "pc",
            "date": date_range,
            "platform": platform,
            "pageType": "v5Trade",
            "optionType": "v5Trade",
            "shopIds": shop_id,
            "excludeShopIds": "",
            "cityId": "",
            "prdIds": self.PRD_IDS,
            "spuId": "",
            "pageNum": "",
            "pageSize": "",
            "sign": "",
            "fromPage": "storeKey",
            "timeStamp": timestamp,
        }

        self.session.headers.update({"Cookie": self.cookie})

        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    delay = random.uniform(2, 5)
                    print(f"  请求失败，{delay:.1f}秒后重试... (第{attempt + 1}次)")
                    time.sleep(delay)

                response = self.session.post(
                    self.BASE_URL,
                    params=params,
                    data=data,
                    timeout=30
                )
                result = response.json()

                # 交易分析返回 code=0 表示成功
                if result.get("code") == 0 or result.get("success") is True:
                    return result

                # 401/403 错误时打印警告
                if result.get("code") in [401, 403]:
                    print(f"  认证错误 (code={result.get('code')})，请更新凭证")

                return result

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
        获取交易分析指标数据

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

        # 交易分析返回 code=0 表示成功
        if data.get("code") != 0 and data.get("success") is not True:
            raise Exception(f"API返回错误: code={data.get('code')}, msg={data.get('msg')}")

        result = {}

        # 解析响应数据
        # 数据结构: data[0].body[]
        for component in data.get("data", []):
            component_id = component.get("componentId", "")

            if component_id == "tradeOptionSummaryPC":
                body = component.get("body", [])

                for item in body:
                    name = item.get("name", "")
                    value = item.get("value", "0")
                    suffix = item.get("suffix", "")

                    if name in self.METRIC_NAMES:
                        result[self.METRIC_NAMES[name]] = value
                        # 同时保存suffix方便后续处理
                        if suffix:
                            result[f"{self.METRIC_NAMES[name]}_suffix"] = suffix

                break

        # 如果解析失败，尝试打印原始数据以便调试
        if not result:
            print(f"  警告: 未能解析数据格式，原始数据: {json.dumps(data, ensure_ascii=False)[:500]}")

        return result


if __name__ == "__main__":
    """测试交易分析API"""
    print("=" * 50)
    print("交易分析 API 测试")
    print("=" * 50)

    client = TradeAnalysisClient()

    # 测试参数
    shop_id = "1933643130"

    print(f"\n测试参数:")
    print(f"  门店ID: {shop_id}")
    print(f"  平台: 点评+美团 (0)")
    print()

    try:
        # 获取报表数据
        print("正在请求 API...")
        data = client.get_report(
            shop_id=shop_id,
            platform="0",
            begin_date="2026-05-08",
            end_date="2026-05-14"
        )

        print(f"API 返回 code: {data.get('code')}")
        print(f"API 返回 msg: {data.get('msg')}")
        print(f"API 返回 success: {data.get('success')}")
        print()

        # 解析并显示数据
        metrics = client.get_metrics(
            shop_id=shop_id,
            platform="0",
            begin_date="2026-05-08",
            end_date="2026-05-14"
        )

        print("=" * 40)
        print("交易分析指标:")
        print("=" * 40)
        for key, value in metrics.items():
            if not key.endswith("_suffix"):
                print(f"  {key}: {value}")
        print()

        print("=" * 60)
        print("API 测试成功!")
        print("=" * 60)

    except Exception as e:
        print(f"错误: {e}")

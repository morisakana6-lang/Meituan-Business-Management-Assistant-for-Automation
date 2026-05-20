"""
在线咨询分析数据获取客户端

=== 板块归属 ===
在线咨询分析板块，对应目录 online_consultation/

=== API 信息 ===
URL: https://e.dianping.com/mda/v5/onlineConsultant
方法: POST
Content-Type: application/x-www-form-urlencoded

=== Payload 参数 ===
- source: 1
- device: pc
- date: YYYY-MM-DD,YYYY-MM-DD（逗号分隔）
- platform: 0=全平台，1=点评，2=美团
- pageType: v5OnlineConsultant
- optionType: v5OnlineConsultant
- shopIds: 门店ID（多个用逗号分隔）
- storeKey: 门店Key
- timeStamp: 时间戳

=== 响应格式 ===
（待确认）

=== 使用方法 ===
from online_consultation.client import OnlineConsultationClient

client = OnlineConsultationClient()
data = client.get_metrics(
    shop_id="xxx",
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


def load_shop_config():
    """从 shop_config.json 加载配置"""
    config_file = os.path.join(_CURRENT_DIR, 'shop_config.json')
    if os.path.exists(config_file):
        with open(config_file, encoding='utf-8') as f:
            return json.load(f)
    return {}


def load_credentials():
    """直接从凭证文件读取"""
    cred_file = os.path.join(_PROJECT_ROOT, 'auth', 'credentials.json')
    if os.path.exists(cred_file):
        with open(cred_file, encoding='utf-8') as f:
            data = json.load(f)
            return data.get('cookie', ''), data.get('mtgsig', '')
    return '', ''


class OnlineConsultationClient:
    """在线咨询分析报表数据客户端"""

    # 在线咨询分析 API 地址
    BASE_URL = "https://e.dianping.com/mda/v5/onlineConsultant"

    # 默认请求头
    DEFAULT_HEADERS = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": "https://h5.dianping.com",
        "Referer": "https://h5.dianping.com/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    }

    # 指标名称映射（variable -> 中文名）
    METRIC_NAMES = {
        "ask_user_cnt": "在线咨询人数",
        "im_booking_uv": "在线咨询留资数",
        "im_booking_uv_rate": "咨询留资转化率",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def _build_url(self, mtgsig: str = None) -> str:
        """构建带mtgsig参数的URL"""
        sig = mtgsig or self.mtgsig
        if not sig:
            raise ValueError("mtgsig is required")
        return f"{self.BASE_URL}?yodaReady=h5&csecplatform=4&csecversion=4.2.0&mtgsig={sig}"

    def _get_timestamp(self) -> str:
        """获取当前时间戳（毫秒）"""
        return str(int(datetime.now().timestamp() * 1000))

    def get_report(
        self,
        shop_id: str,
        platform: str = "0",
        begin_date: str = None,
        end_date: str = None,
        store_key: str = None,
        max_retries: int = 3,
    ) -> Dict:
        """
        获取在线咨询分析报表数据

        Args:
            shop_id: 门店ID
            platform: 平台，0=全平台，1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD（可选，默认近7天）
            end_date: 结束日期，格式 YYYY-MM-DD（可选，默认今天）
            store_key: 门店Key（可选，默认使用shop_id）
            max_retries: 最大重试次数

        Returns:
            API响应字典
        """
        # 设置默认日期范围
        if not end_date:
            end_date = datetime.now().strftime("%Y-%m-%d")
        if not begin_date:
            begin_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")

        # date参数格式：YYYY-MM-DD,YYYY-MM-DD（逗号分隔）
        date_range = f"{begin_date},{end_date}"

        # 构建请求体（根据Payload）
        payload = {
            "source": "1",
            "device": "pc",
            "date": date_range,
            "platform": platform,
            "pageType": "v5OnlineConsultant",
            "optionType": "v5OnlineConsultant",
            "shopIds": shop_id,
            "storeKey": store_key or shop_id,
            "timeStamp": self._get_timestamp(),
        }

        url = self._build_url()
        self.session.headers["Cookie"] = self.cookie

        for attempt in range(max_retries):
            try:
                response = self.session.post(url, data=payload, timeout=30)
                if response.status_code == 200:
                    return response.json()
                elif response.status_code in (401, 403):
                    raise Exception("凭证已过期，请运行 python common/credentials.py --force 更新")
                else:
                    print(f"请求失败 (状态码: {response.status_code}), 响应: {response.text[:200]}")
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                print(f"请求异常: {e}")
                time.sleep(1)

        return {}

    def get_metrics(
        self,
        shop_id: str,
        platform: str = "0",
        begin_date: str = None,
        end_date: str = None,
    ) -> Dict:
        """
        获取在线咨询分析指标数据

        Args:
            shop_id: 门店ID
            platform: 平台，0=全平台（默认），1=点评，2=美团
            begin_date: 开始日期，格式 YYYY-MM-DD（可选）
            end_date: 结束日期，格式 YYYY-MM-DD（可选）

        Returns:
            指标字典 {
                "在线咨询人数": "123",
                "在线咨询留资数": "45",
                "咨询留资转化率": "36.59%"
            }

        Raises:
            Exception: API返回错误时抛出
        """
        response = self.get_report(shop_id, platform, begin_date, end_date)

        if not response or response.get("code") != 0:
            raise Exception(f"API返回错误: {response}")

        data_list = response.get("data", [])
        metrics = {}

        # 遍历所有组件，找到包含指标的body
        for data_item in data_list:
            body = data_item.get("body", [])
            for item in body:
                variable = item.get("variable", "")
                value = item.get("value", "")
                metric_name = self.METRIC_NAMES.get(variable, "")
                if metric_name:
                    metrics[metric_name] = value

        return metrics


if __name__ == "__main__":
    """测试在线咨询分析API"""
    print("=" * 50)
    print("在线咨询分析 API 测试")
    print("=" * 50)

    # 从配置文件加载
    config = load_shop_config()
    print(f"\n配置文件内容:")
    print(json.dumps(config, ensure_ascii=False, indent=2))

    shop_config = config.get("门店列表", [{}])[0] if config.get("门店列表") else {}
    shop_id = shop_config.get("shop_id", "1933643130")
    date_range = config.get("日期范围", {})
    begin_date = date_range.get("begin")
    end_date = date_range.get("end")
    platform = str(config.get("platform", 0))

    print(f"\n使用配置:")
    print(f"  门店ID: {shop_id}")
    print(f"  日期范围: {begin_date} 至 {end_date}")
    print(f"  平台: {platform}")

    client = OnlineConsultationClient()
    metrics = client.get_metrics(shop_id, platform, begin_date, end_date)
    print(f"\n获取到的指标:")
    for name, value in metrics.items():
        print(f"  {name}: {value}")

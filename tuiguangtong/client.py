"""
推广通数据获取客户端

=== 主要作用 ===
封装美团经营宝推广通板块的 API 调用，负责：
1. 从 common.credentials 加载 Cookie 和 mtgsig
2. 调用美团 API 获取报表数据
3. 解析返回数据为统一的指标格式

=== 板块归属 ===
推广通板块，对应目录 tuiguangtong/

=== API 信息 ===
- 端点: https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport
- 支持日期范围查询
- 支持门店筛选（shop_ids）

=== 核心类 ===
TuiguangtongClient
  - __init__(): 初始化，自动加载凭证
  - get_report(): 获取原始报表 JSON 数据（带重试机制）
  - get_metrics(): 获取简化后的指标列表

=== 使用方法 ===
# 方式1: 自动加载凭证
from tuiguangtong.client import TuiguangtongClient

client = TuiguangtongClient()
metrics = client.get_metrics("2025-05-01", "2025-05-31", shop_ids="764510450")

# 方式2: 手动传入凭证
client = TuiguangtongClient(cookie="...", mtgsig="...")

# 方式3: 配合 excel_generator 生成报表
from tuiguangtong.client import TuiguangtongClient
from tuiguangtong.excel_generator import generate_daily_report

client = TuiguangtongClient()
generate_daily_report(client, "2025-05-01", "2025-05-31", "764510450", "门店名称", "764510450")

=== 依赖 ===
- common.credentials.load_credentials(): 加载凭证
- requests: 发送 HTTP 请求
- datetime: 日期处理
"""

import requests
import json
import random
import time
from datetime import datetime, timedelta
from typing import List, Dict

import sys
sys.path.insert(0, '..')
from common import load_credentials


class TuiguangtongClient:
    """
    推广通报表数据客户端

    Attributes:
        cookie: 登录 Cookie 字符串
        mtgsig: mtgsig 签名参数
        session: requests.Session 实例，用于复用连接
        _mtgsig_updater: mtgsig 更新回调（当 401/403 时自动调用）
    """

    BASE_URL = "https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport"

    DEFAULT_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e.dianping.com/shopdiy-node/report",
    }

    METRIC_NAMES = {
        "T30001": "花费",
        "T30047": "现金花费",
        "T30002": "曝光",
        "T30003": "点击",
        "T30020": "团购订单量",
        "T30049": "订单量",
        "T30005": "商户浏览量",
        "T30006": "查看图片",
        "T30007": "查看评论",
        "T30039": "查看店铺信息",
        "T30009": "查看团购",
        "T30083": "感兴趣",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        """
        初始化客户端

        Args:
            cookie: Cookie 字符串（可选，不传则自动从 common.credentials 加载）
            mtgsig: mtgsig 签名参数（可选，不传则自动从 common.credentials 加载）

        Note:
            自动加载凭证优先级：传入参数 > load_credentials() 加载
        """
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)
        self._mtgsig_updater = None

    def get_report(
        self,
        begin_date: str,
        end_date: str,
        shop_ids: str = "0",
        platform: int = 0,
        compare_enabled: bool = False,
        max_retries: int = 3,
    ) -> Dict:
        """
        获取报表原始数据（带自动重试机制）

        Args:
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            shop_ids: 门店ID，"0" 表示全部门店
            platform: 平台，0=全平台，1=点评，2=美团
            compare_enabled: 是否启用对比（对比日期为 begin_date 前7天）
            max_retries: 最大重试次数

        Returns:
            API 响应的完整 JSON 数据，结构示例：
            {
                "code": 200,
                "msg": {
                    "total": [...],    # 指标列表
                    "dimension": "shop",
                    ...
                }
            }

        重试逻辑：
            1. 请求失败时随机延迟 2~5 秒后重试
            2. 遇到 401/403 认证错误，尝试调用 _mtgsig_updater 更新 mtgsig
            3. 最多重试 max_retries 次
        """
        compare_begin = compare_end = ""
        if compare_enabled:
            start_dt = datetime.strptime(begin_date, "%Y-%m-%d")
            compare_end_dt = start_dt - timedelta(days=1)
            compare_begin_dt = compare_end_dt - timedelta(days=6)
            compare_begin = compare_begin_dt.strftime("%Y-%m-%d")
            compare_end = compare_end_dt.strftime("%Y-%m-%d")

        params = {
            "dimension": "shop",
            "beginDate": begin_date,
            "endDate": end_date,
            "platform": str(platform),
            "compareEnabled": "1" if compare_enabled else "0",
            "compareBeginDate": compare_begin,
            "compareEndDate": compare_end,
            "objectUnit": "account",
            "groupUnit": "",
            "timeUnit": "day",
            "shopIds": shop_ids,
            "launchIds": "0",
            "launchPremiumIds": "0",
            "planIds": "0",
            "tabIds": "T30001,T30047,T30002,T30003,T30020,T30049,T30005,T30006,T30007,T30039,T30009,T30083",
            "reportFunctionType": "2",
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

                response = self.session.get(self.BASE_URL, params=params)
                data = response.json()

                if data.get("code") == 200:
                    return data

                if data.get("code") in [401, 403] and self._mtgsig_updater:
                    print(f"  检测到认证问题，尝试更新mtgsig...")
                    new_mtgsig = self._mtgsig_updater()
                    if new_mtgsig:
                        self.mtgsig = new_mtgsig
                        params["mtgsig"] = self.mtgsig
                        continue

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
        shop_ids: str = "0",
        platform: int = 0,
        compare_enabled: bool = False,
    ) -> List[Dict]:
        """
        获取指标数据列表

        Args:
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            shop_ids: 门店ID，"0" 表示全部门店
            platform: 平台，0=全平台，1=点评，2=美团
            compare_enabled: 是否启用对比

        Returns:
            指标列表，每项结构：
            {
                "id": "T30001",      # 指标ID
                "name": "花费",       # 指标名称
                "value": "1,234",    # 指标值（可能是带逗号的字符串）
                "unit": "元",         # 单位
                "group": "广告花费与成本",  # 分组
                "tips": ""           # 提示信息
            }

        Raises:
            Exception: 当 API 返回非 200 状态时抛出异常

        注意：
            - 如果 API 返回的 total 列表为空，会打印警告
            - 如果发现未知的指标ID，会打印警告提示更新 METRIC_NAMES
            - value 可能是带逗号的字符串（如 "1,234"），需自行处理
        """
        data = self.get_report(begin_date, end_date, shop_ids, platform, compare_enabled)

        if data.get("code") != 200:
            raise Exception(f"API返回错误: code={data.get('code')}, msg={data.get('msg')}")

        metrics = []
        all_items = data.get("msg", {}).get("total", [])

        if not all_items:
            print("\n⚠️ 警告: API返回的total列表为空")

        for item in all_items:
            metric_id = item.get("id")

            if metric_id not in self.METRIC_NAMES:
                print(f"\n⚠️ 发现新的指标ID: {metric_id} (name={item.get('name')})")

            metrics.append({
                "id": metric_id,
                "name": self.METRIC_NAMES.get(metric_id, item.get("name", "未知")),
                "value": item.get("value", "0"),
                "unit": item.get("unit", ""),
                "group": item.get("groupName", ""),
                "tips": item.get("tips", ""),
            })

        return metrics

"""
全站推广数据获取客户端

=== 板块归属 ===
全站推广板块，对应目录 quantianzhan/

=== API 信息 ===
URL: https://e.dianping.com/shopdiy/report/datareport/ajax/cpcwsp/queryTabDataList
方法: GET

=== 指标ID ===
| 指标ID | 名称 | 单位 |
|--------|------|------|
| T30001 | 花费 | 元 |
| T30047 | 现金花费 | 元 |
| T500001 | 全站浏览量 | 次 |
| T600008 | 下单券数 | 张 |
| T600009 | 每张券成本 | 元 |

=== 请求参数 ===
- beginDate: 开始日期 (YYYY-MM-DD)
- endDate: 结束日期 (YYYY-MM-DD)
- timeUnit: 时间粒度 (day)
- clientKey: 固定值 "cpc.wsp.data.common"
- compareEnabled: 是否对比 (0/1)
- tabIds: 指标ID列表 (逗号分隔)
- platform: 平台 (0=美团, 1=点评)
- shopIds: 门店ID (0表示账户级别)
- reportScene: 场景 (wsp_1_3)
- yodaReady: 固定值 "h5"
- csecplatform: 固定值 "4"
- csecversion: 固定值 "4.2.0"
- mtgsig: 认证参数

=== 响应格式 ===
{
    "code": 200,
    "msg": {
        "total": [
            {
                "id": "T30001",
                "name": "花费",
                "value": "123.45",
                "unit": "元",
                "groupName": "广告花费与成本",
                ...
            },
            ...
        ]
    }
}

=== 使用方法 ===
from quantianzhan.client import QuantianzhanClient

client = QuantianzhanClient()
data = client.get_metrics(
    begin_date="2026-05-01",
    end_date="2026-05-14",
    tab_ids="T30001,T30047,T500001,T600008,T600009"
)
"""

import requests
import json
import random
import time
from datetime import datetime, timedelta
from typing import List, Dict

import sys
import os
sys.path.insert(0, '..')
from common import load_credentials

# 复用推广通的门店搜索模块
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_SHOP_SEARCH_PATH = os.path.join(_CURRENT_DIR, '..', 'tuiguangtong', 'shop_search.py')
sys.path.insert(0, os.path.dirname(_SHOP_SEARCH_PATH))
from tuiguangtong.shop_search import resolve_shop


class QuantianzhanClient:
    """全站推广报表数据客户端"""

    # 全站推广 API 地址
    BASE_URL = "https://e.dianping.com/shopdiy/report/datareport/ajax/cpcwsp/queryTabDataList"

    # 默认请求头
    DEFAULT_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e.dianping.com/app/peon-isomorph-full-site/html/data-report-pc.html",
    }

    # 指标名称映射
    METRIC_NAMES = {
        "T30001": "花费",
        "T30047": "现金花费",
        "T500001": "全站浏览量",
        "T600008": "下单券数",
        "T600009": "每张券成本",
    }

    # 默认查询的指标
    DEFAULT_TAB_IDS = "T30001,T30047,T500001,T600008,T600009"

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def get_report(
        self,
        begin_date: str,
        end_date: str,
        tab_ids: str = None,
        platform: str = "0",
        shop_ids: str = "0",
        compare_enabled: bool = False,
        max_retries: int = 3,
    ) -> Dict:
        """
        获取全站推广报表数据

        Args:
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            tab_ids: 指标ID列表，逗号分隔，默认 "T30001,T30047,T500001,T600008,T600009"
            platform: 平台，0=美团，1=点评
            shop_ids: 门店ID，0表示账户级别
            compare_enabled: 是否启用数据对比
            max_retries: 最大重试次数

        Returns:
            API响应字典
        """
        if tab_ids is None:
            tab_ids = self.DEFAULT_TAB_IDS

        # 计算对比日期（上周同期）
        compare_begin = ""
        compare_end = ""
        if compare_enabled:
            start_dt = datetime.strptime(begin_date, "%Y-%m-%d")
            end_dt = datetime.strptime(end_date, "%Y-%m-%d")
            days = (end_dt - start_dt).days + 1
            compare_end_dt = start_dt - timedelta(days=1)
            compare_begin_dt = compare_end_dt - timedelta(days=days - 1)
            compare_begin = compare_begin_dt.strftime("%Y-%m-%d")
            compare_end = compare_end_dt.strftime("%Y-%m-%d")

        params = {
            "beginDate": begin_date,
            "endDate": end_date,
            "timeUnit": "day",
            "clientKey": "cpc.wsp.data.common",
            "compareEnabled": "1" if compare_enabled else "0",
            "compareBeginDate": compare_begin,
            "compareEndDate": compare_end,
            "groupUnit": "",
            "objectUnit": "",
            "tabIds": tab_ids,
            "platform": platform,
            "smartShopFlag": "",
            "reportFunctionType": "",
            "shopIds": shop_ids,
            "reportScene": "wsp_1_3",
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
        tab_ids: str = None,
        platform: str = "0",
        shop_ids: str = "0",
        search_key: str = None,
        compare_enabled: bool = False,
    ) -> List[Dict]:
        """
        获取指标数据列表

        Args:
            begin_date: 开始日期
            end_date: 结束日期
            tab_ids: 指标ID列表
            platform: 平台，0=美团，1=点评
            shop_ids: 门店ID，0表示全部门店汇总
            search_key: 门店搜索关键字（ID或名称），会覆盖shop_ids
            compare_enabled: 是否对比

        Returns:
            指标列表，每项包含 id, name, value, unit

        Raises:
            Exception: API返回错误时抛出
        """
        # 如果提供了search_key，通过门店搜索解析
        if search_key:
            _, resolved_shop_id = resolve_shop(search_key, int(platform))
            shop_ids = resolved_shop_id

        data = self.get_report(
            begin_date=begin_date,
            end_date=end_date,
            tab_ids=tab_ids,
            platform=platform,
            shop_ids=shop_ids,
            compare_enabled=compare_enabled,
        )

        if data.get("code") != 200:
            raise Exception(f"API返回错误: code={data.get('code')}, msg={data.get('msg')}")

        metrics = []
        total = data.get("msg", {}).get("total", [])

        if not total:
            print("  警告: API返回的total列表为空")

        for item in total:
            metric_id = item.get("id", "")
            metric_name = item.get("name", "")

            # 如果名称为空，尝试从映射表获取
            if not metric_name:
                metric_name = self.METRIC_NAMES.get(metric_id, "未知指标")

            metrics.append({
                "id": metric_id,
                "name": metric_name,
                "value": item.get("value", "0"),
                "unit": item.get("unit", ""),
                "group": item.get("groupName", ""),
                "tips": item.get("tips", ""),
            })

        return metrics


if __name__ == "__main__":
    """测试全站推广API"""
    print("=" * 50)
    print("全站推广 API 测试")
    print("=" * 50)

    client = QuantianzhanClient()

    # 测试获取指标数据
    try:
        metrics = client.get_metrics(
            begin_date="2026-05-01",
            end_date="2026-05-14",
        )

        print("\n获取到的指标:")
        for m in metrics:
            print(f"  {m['id']} | {m['name']} | 值: {m['value']} {m['unit']}")

    except Exception as e:
        print(f"错误: {e}")

"""
评论统计数据获取客户端

=== 板块归属 ===
评论统计板块，对应目录 review_statistics/

=== API 信息 ===
URL: https://e.dianping.com/gateway/adviser/data
方法: GET

=== 请求参数 ===
- source: 1
- device: pc
- date: YYYY-MM-DD,YYYY-MM-DD（逗号分隔）
- platform: 0
- pageType: reviewAnalysisV2
- shopIds: 门店ID

=== 需要获取的指标 ===
- 累计评价数
- 新增评价数
- 新增好评数
- 新增差评数
- 差评回复率
- 星级评分

=== 使用方法 ===
from review_statistics.client import ReviewStatisticsClient

client = ReviewStatisticsClient()
data = client.get_statistics(
    shop_id="1933643130",
    begin_date="2026-04-01",
    end_date="2026-04-30"
)
"""

import requests
import json
from datetime import datetime, timedelta
from typing import Dict, List

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


def load_shop_mapping():
    """从shop_mapping.json加载门店映射表"""
    mapping_file = os.path.join(_PROJECT_ROOT, 'tuiguangtong', 'shop_mapping.json')
    if os.path.exists(mapping_file):
        with open(mapping_file, encoding='utf-8') as f:
            data = json.load(f)
            shops = data.get('shops', [])
            # 构建ID到名称的映射
            mapping = {}
            for shop in shops:
                name = shop.get('name', '')
                ids = shop.get('ids', [])
                for id_info in ids:
                    shop_id = str(id_info.get('id', ''))
                    if shop_id:
                        mapping[shop_id] = name
            return mapping
    return {}


def get_shop_name(shop_id: str) -> str:
    """根据门店ID获取门店名称"""
    mapping = load_shop_mapping()
    return mapping.get(str(shop_id), '')


class ReviewStatisticsClient:
    """评论统计数据客户端"""

    BASE_URL = "https://e.dianping.com/gateway/adviser/data"

    DEFAULT_HEADERS = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "keep-alive",
        "Origin": "https://e.dianping.com",
        "Referer": "https://e.dianping.com/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    }

    def __init__(self, cookie: str = None, mtgsig: str = None):
        self.cookie = cookie or load_credentials()[0]
        self.mtgsig = mtgsig or load_credentials()[1]
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)

    def _build_url(self, shop_id: str, date_range: str, platform: str = "0") -> str:
        """构建完整URL"""
        base_url = f"{self.BASE_URL}?source=1&device=pc&date={date_range}&platform={platform}&pageType=reviewAnalysisV2&optionType=&shopIds={shop_id}&excludeShopIds=&cityId=&prdIds=&spuId=&pageNum=&pageSize=&sign=&fromPage=&yodaReady=h5&csecplatform=4&csecversion=4.2.0&mtgsig={self.mtgsig}"
        return base_url

    def get_statistics(
        self,
        shop_id: str,
        begin_date: str = None,
        end_date: str = None,
        platform: str = "0",
    ) -> Dict:
        """
        获取评论统计数据

        Args:
            shop_id: 门店ID
            begin_date: 开始日期，格式 YYYY-MM-DD
            end_date: 结束日期，格式 YYYY-MM-DD
            platform: 平台，0=全平台

        Returns:
            统计数据字典
        """
        # 设置默认日期
        if not end_date:
            end_date = datetime.now().strftime("%Y-%m-%d")
        if not begin_date:
            begin_date = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")

        date_range = f"{begin_date},{end_date}"
        url = self._build_url(shop_id, date_range, platform)
        self.session.headers["Cookie"] = self.cookie

        try:
            response = self.session.get(url, timeout=30)
            if response.status_code != 200:
                print(f"请求失败 (状态码: {response.status_code})")
                return {}

            result = response.json()
            if result.get("code") != 200:
                print(f"API错误: {result.get('msg')}")
                return {}

            return result

        except Exception as e:
            print(f"请求异常: {e}")
            return {}

    def parse_statistics(self, result: Dict) -> Dict:
        """
        解析统计数据

        Args:
            result: API原始响应

        Returns:
            统计数据字典
        """
        statistics = {}

        data_list = result.get("data", [])
        for data_item in data_list:
            if data_item.get("uniqueMarker") == "reviewAnalysisV2PCSummary":
                body = data_item.get("body", {})
                items = body.get("data", [])

                for item in items:
                    name = item.get("name", "")
                    value = item.get("value", "")

                    if name == "累计评价数":
                        statistics["累计评价数"] = value
                    elif name == "新增评价数":
                        statistics["新增评价数"] = value
                    elif name == "新增好评数":
                        statistics["新增好评数"] = value
                    elif name == "新增差评数":
                        statistics["新增差评数"] = value
                    elif name == "差评回复率":
                        statistics["差评回复率"] = value

        return statistics


if __name__ == "__main__":
    """测试评论统计API"""
    print("=" * 50)
    print("评论统计 API 测试")
    print("=" * 50)

    config_file = os.path.join(_CURRENT_DIR, 'shop_config.json')
    if os.path.exists(config_file):
        with open(config_file, encoding='utf-8') as f:
            config = json.load(f)

    search_key = config.get("search_key", "1933643130")
    if isinstance(search_key, list):
        shop_ids = search_key
    elif isinstance(search_key, str) and "," in search_key:
        shop_ids = [s.strip() for s in search_key.split(",")]
    else:
        shop_ids = [search_key]

    shop_id = shop_ids[0]
    platform = config.get("platform", 0)
    date_range = config.get("日期范围", {})
    begin_date = date_range.get("begin")
    end_date = date_range.get("end")

    print(f"\n使用参数:")
    print(f"  门店ID: {shop_ids}")
    print(f"  平台: {platform}")
    print(f"  日期: {begin_date} 至 {end_date}")

    client = ReviewStatisticsClient()
    result = client.get_statistics(shop_id, begin_date, end_date, str(platform))

    print(f"\n原始响应:")
    print(json.dumps(result, ensure_ascii=False, indent=2))
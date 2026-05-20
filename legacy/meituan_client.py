"""
美团经营宝数据获取客户端（旧版入口，参考用）

=== 主要作用 ===
这是项目旧版本的入口文件，用于：
1. 从 credentials.json 加载凭证
2. 调用美团 API 获取推广报表数据
3. 调用 Excel 生成器生成报表文件

=== 板块归属 ===
属于"推广通"板块，对应目录 tuiguangtong/

=== 核心类 ===
- MeituanReportClient: 美团经营宝报表数据客户端
  - get_report(): 获取原始报表 JSON 数据
  - get_metrics(): 获取简化后的指标列表
  - print_metrics(): 打印指标数据

=== API 信息 ===
- 端点: https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport
- 指标ID: T30001(花费), T30002(曝光), T30003(点击), T30047(现金花费) 等

=== 使用方法（已废弃，请使用 tuiguangtong/ 目录下的新版本） ===
# 1. 修改文件顶部的配置区域
SHOP_NAME = "门店名称"
SHOP_ID = "门店ID"
BEGIN_DATE = "2025-05-01"
END_DATE = "2025-06-30"

# 2. 运行
python meituan_client.py

# 3. 输出
- 控制台显示每日数据获取进度
- 生成 report_{shop_id}_{begin}_{end}.xlsx 文件

=== 与新版本 tuiguangtong/client.py 的区别 ===
| 功能 | 旧版 (legacy) | 新版 (tuiguangtong) |
|------|---------------|---------------------|
| 凭证加载 | 读取根目录 credentials.json | 调用 common.load_credentials() |
| Excel 生成 | 内嵌在文件中 | 独立的 excel_generator.py |
| 目录结构 | 扁平 | 按板块分离 |
| 输出目录 | 根目录 | tuiguangtong/reports/ |
"""

import requests
import json
import random
import time
from datetime import datetime, timedelta
from typing import List, Dict, Optional


class MeituanReportClient:
    """美团经营宝报表数据客户端"""

    # API地址
    BASE_URL = "https://e.dianping.com/shopdiy/report/datareport/pc/ajax/getBoardReport"

    # 默认请求头
    DEFAULT_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e.dianping.com/shopdiy-node/report",
    }

    # 指标ID到名称的映射
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

    def __init__(self, cookie: str, mtgsig: str):
        """
        初始化客户端

        Args:
            cookie: 登录Cookie
            mtgsig: mtgsig签名参数
        """
        self.cookie = cookie
        self.mtgsig = mtgsig
        self.session = requests.Session()
        self.session.headers.update(self.DEFAULT_HEADERS)
        self._mtgsig_updater = None  # 用于更新mtgsig的回调

    def get_report(
        self,
        begin_date: str,
        end_date: str,
        shop_ids: str = "0",
        compare_enabled: bool = False,
        max_retries: int = 3,
    ) -> Dict:
        """
        获取报表数据（带自动重试机制）

        Args:
            begin_date: 开始日期，格式YYYY-MM-DD
            end_date: 结束日期，格式YYYY-MM-DD
            shop_ids: 门店ID，0表示全部
            compare_enabled: 是否启用对比
            max_retries: 最大重试次数

        Returns:
            API响应的完整JSON数据
        """
        # 如果启用对比，设置对比日期为开始日期前7天
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
            "platform": "0",
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
                # 随机延迟，避免高频请求触发风控
                if attempt > 0:
                    delay = random.uniform(2, 5)
                    print(f"  请求失败，{delay:.1f}秒后重试... (第{attempt + 1}次)")
                    time.sleep(delay)

                response = self.session.get(self.BASE_URL, params=params)
                data = response.json()

                # 检查是否成功
                if data.get("code") == 200:
                    return data

                # 如果是认证问题，尝试更新mtgsig
                if data.get("code") in [401, 403] and self._mtgsig_updater:
                    print(f"  检测到认证问题，尝试更新mtgsig...")
                    new_mtgsig = self._mtgsig_updater()
                    if new_mtgsig:
                        self.mtgsig = new_mtgsig
                        params["mtgsig"] = self.mtgsig
                        continue

                # 其他错误直接返回
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
        compare_enabled: bool = False,
    ) -> List[Dict]:
        """
        获取指标数据列表

        Args:
            begin_date: 开始日期
            end_date: 结束日期
            shop_ids: 门店ID

        Returns:
            简化后的指标列表
        """
        data = self.get_report(begin_date, end_date, shop_ids, compare_enabled)

        if data.get("code") != 200:
            raise Exception(f"API返回错误: code={data.get('code')}, msg={data.get('msg')}")

        metrics = []
        all_items = data.get("msg", {}).get("total", [])

        # 如果没有数据，打印警告并显示所有返回的指标ID
        if not all_items:
            print("\n⚠️ 警告: API返回的total列表为空，可能原因:")
            print("  1. 指标ID已更新，需要更新 METRIC_NAMES 映射表")
            print("  2. 查询日期范围内无数据")
            print("  3. 门店ID不存在或无权限")

            # 打印原始返回数据的keys
            print(f"\n原始响应msg keys: {list(data.get('msg', {}).keys())}")
            print(f"原始响应: {json.dumps(data, ensure_ascii=False)[:500]}")

        for item in all_items:
            metric_id = item.get("id")

            # 如果ID不在已知映射中，打印警告
            if metric_id not in self.METRIC_NAMES:
                print(f"\n⚠️ 发现新的指标ID: {metric_id} (name={item.get('name')})")
                print(f"   请更新 METRIC_NAMES 映射表")

            metrics.append({
                "id": metric_id,
                "name": self.METRIC_NAMES.get(metric_id, item.get("name", "未知")),
                "value": item.get("value", "0"),
                "unit": item.get("unit", ""),
                "group": item.get("groupName", ""),
                "tips": item.get("tips", ""),
            })

        return metrics

    def print_metrics(self, metrics: List[Dict], title: str = ""):
        """打印指标数据"""
        if title:
            print(f"\n{'='*60}")
            print(f"{title}")
            print(f"{'='*60}")

        print(f"{'指标ID':<10} {'指标名称':<15} {'当前值':<15} {'单位':<8} {'分组'}")
        print("-" * 70)

        current_metrics = []
        for m in metrics:
            line = f"{m['id']:<10} {m['name']:<15} {m['value']:<15} {m['unit']:<8} {m['group']}"
            print(line)
            current_metrics.append({
                "id": m["id"],
                "name": m["name"],
                "value": m["value"],
                "unit": m["unit"],
            })

        return current_metrics


def main():
    import os
    import sys

    # ========== 配置区域 ==========

    # 凭证文件路径（由 get_credentials.py 生成）
    CREDENTIALS_FILE = "credentials.json"

    # 尝试从文件加载凭证，如果文件不存在或过期则使用下面的默认值
    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
            creds = json.load(f)
        COOKIE = creds.get("cookie", "")
        MTGSIG = creds.get("mtgsig", "")
        print(f"已从 {CREDENTIALS_FILE} 加载凭证 (更新时间: {creds.get('update_time', '未知')})")
    else:
        # 凭证文件不存在，使用下面的默认值（需要手动更新）
        print("错误: 未找到凭证文件，请先运行 python get_credentials.py 获取凭证")
        sys.exit(1)

    # 门店配置
    SHOP_NAME = "丽减美瘦吧·减肥瘦身(新澳城店)"
    SHOP_ID = "764510450"
    SHOP_IDS = SHOP_ID  # 门店ID

    # 日期范围配置
    BEGIN_DATE = "2025-05-01"
    END_DATE = "2025-06-30"
    # ==============================

    # 创建客户端
    client = MeituanReportClient(COOKIE, MTGSIG)

    # 设置mtgsig更新回调
    def update_mtgsig():
        """从登录态重新获取mtgsig"""
        print("  尝试重新获取 mtgsig...")
        os.system("python get_credentials.py --update")
        if os.path.exists(CREDENTIALS_FILE):
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                new_creds = json.load(f)
            return new_creds.get("mtgsig")
        return None

    client._mtgsig_updater = update_mtgsig

    # 导入 Excel 生成器
    try:
        from excel_generator import generate_daily_report
        print(f"开始生成 {SHOP_NAME} (ID: {SHOP_ID}) 的 {BEGIN_DATE} ~ {END_DATE} 每日报表...")
        output_file = generate_daily_report(
            client=client,
            begin_date=BEGIN_DATE,
            end_date=END_DATE,
            shop_ids=SHOP_IDS,
            shop_name=SHOP_NAME,
            shop_id=SHOP_ID
        )
        print(f"\nExcel 文件已生成: {output_file}")
    except ImportError:
        print("错误: 无法导入 excel_generator 模块")
        sys.exit(1)
    except Exception as e:
        print(f"生成 Excel 失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

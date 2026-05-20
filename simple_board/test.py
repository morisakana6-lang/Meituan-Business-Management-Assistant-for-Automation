"""
简版看板测试文件

用于测试 API 连接和数据获取
"""

import sys
import os
sys.path.insert(0, '..')

from client import SimpleBoardClient


def test_api():
    """测试简版看板API"""
    print("=" * 60)
    print("简版看板 API 测试")
    print("=" * 60)

    client = SimpleBoardClient()

    # 测试参数
    shop_id = "1027064882511291"
    begin_date = "2026-05-08"
    end_date = "2026-05-14"

    print(f"\n测试参数:")
    print(f"  门店ID: {shop_id}")
    print(f"  日期范围: {begin_date} ~ {end_date}")
    print()

    try:
        # 获取报表数据
        print("正在请求 API...")
        data = client.get_report(
            begin_date=begin_date,
            end_date=end_date,
            shop_id=shop_id,
            platform="0"
        )

        print(f"API 返回 code: {data.get('code')}")
        print(f"API 返回 msg: {data.get('msg')}")
        print()

        # 解析并显示数据
        metrics = client.get_metrics(
            begin_date=begin_date,
            end_date=end_date,
            shop_id=shop_id,
            platform="0"
        )

        print("=" * 40)
        print("概览数据 (simpleDataBoardOverview):")
        print("=" * 40)
        overview = metrics.get("overview", {})
        print(f"  曝光人数: {overview.get('exposure_count', 'N/A')}")
        print(f"  访问人数: {overview.get('visitor_count', 'N/A')}")
        print(f"  下单券数: {overview.get('order_count', 'N/A')}")
        print(f"  下单金额(原价): {overview.get('order_amount', 'N/A')}")
        print(f"  核销券数: {overview.get('verify_count', 'N/A')}")
        print(f"  核销金额(原价): {overview.get('verify_amount', 'N/A')}")
        print()

        print("=" * 40)
        print("交易数据 (simpleDataBoardTradeSummaryPC):")
        print("=" * 40)
        trade = metrics.get("trade", {})
        print(f"  下单券数: {trade.get('order_count', 'N/A')}")
        print(f"  下单金额(原价): {trade.get('order_amount', 'N/A')}")
        print(f"  核销券数: {trade.get('verify_count', 'N/A')}")
        print(f"  核销金额(原价): {trade.get('verify_amount', 'N/A')}")
        print()

        print("=" * 40)
        print("流量数据 (simpleDataBoardFlowSummaryPC):")
        print("=" * 40)
        flow = metrics.get("flow", {})
        print(f"  曝光次数: {flow.get('exposure_times', 'N/A')}")
        print(f"  曝光人数: {flow.get('exposure_count', 'N/A')}")
        print(f"  访问次数: {flow.get('visit_times', 'N/A')}")
        print(f"  访问人数: {flow.get('visitor_count', 'N/A')}")
        print()

        print("=" * 60)
        print("API 测试成功!")
        print("=" * 60)

        return True

    except Exception as e:
        print(f"API 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = test_api()
    exit(0 if success else 1)
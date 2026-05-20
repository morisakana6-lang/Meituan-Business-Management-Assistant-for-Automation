"""
客流分析测试文件

用于测试 API 连接和数据获取
"""

import sys
import os
sys.path.insert(0, '..')

from client import CustomerFlowClient


def test_api():
    """测试客流分析API"""
    print("=" * 60)
    print("客流分析 API 测试")
    print("=" * 60)

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

        print("=" * 40)
        print("客流分析指标:")
        print("=" * 40)
        print(f"  曝光人数: {metrics.get('exposure_count', 'N/A')}")
        print(f"  曝光次数: {metrics.get('exposure_times', 'N/A')}")
        print(f"  访问人数: {metrics.get('visitor_count', 'N/A')}")
        print(f"  访问次数: {metrics.get('visit_times', 'N/A')}")
        print(f"  曝光访问转化率: {metrics.get('exposure_visit_rate', 'N/A')}")
        print(f"  意向转化人数: {metrics.get('intention_count', 'N/A')}")
        print(f"  意向转化率: {metrics.get('intention_rate', 'N/A')}")
        print(f"  下单人数: {metrics.get('order_count', 'N/A')}")
        print(f"  留资人数: {metrics.get('retention_count', 'N/A')}")
        print(f"  累计收藏人数: {metrics.get('favorites_count', 'N/A')}")
        print(f"  新增收藏人数: {metrics.get('new_favorites_count', 'N/A')}")
        print(f"  新增打卡人数: {metrics.get('new_checkin_count', 'N/A')}")
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

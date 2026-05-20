"""
星级评分多门店ID测试脚本
使用实际payload参数测试多个门店ID，生成汇总报表
"""

import sys
import os

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)

from star_rating.client import StarRatingClient


def test_multi_shop_report():
    """测试多个门店ID，获取星级评分数据"""
    print("=" * 60)
    print("星级评分多门店ID测试")
    print("=" * 60)

    # 平台参数：0=全平台，1=点评，2=美团
    platform_names = {0: "全平台", 1: "点评", 2: "美团"}

    # 使用payload中的门店IDs
    shop_ids = ["1933643130", "57904793"]
    # 使用配置中的日期
    begin_date = "2026-04-01"
    end_date = "2026-04-30"
    platform = "0"  # 0=全平台

    print(f"\n测试门店IDs: {shop_ids}")
    print(f"日期范围: {begin_date} 至 {end_date}")
    print(f"平台: {platform} ({platform_names.get(int(platform), '未知')})")

    client = StarRatingClient()

    # 从shop_mapping获取门店名称
    from tuiguangtong.shop_search import load_mapping
    mapping_data = load_mapping()
    shops = mapping_data.get("shops", [])

    def get_shop_name(shop_id: str) -> str:
        for shop in shops:
            for id_info in shop.get("ids", []):
                if id_info.get("id") == shop_id:
                    return shop.get("name", f"门店{shop_id}")
        return f"门店{shop_id}"

    results = []

    # 方式1：单次请求多门店（测试是否支持）
    shop_ids_str = ",".join(shop_ids)
    print(f"\n=== 方式1: 单次请求多门店 ({shop_ids_str}) ===")
    try:
        result = client.get_statistics(shop_ids_str, begin_date, end_date, platform)
        if result.get("code") == 200:
            stats = client.parse_statistics(result)
            print(f"  合并数据: {stats if stats else '空数据（不支持多门店合并）'}")
        else:
            print(f"  错误: {result.get('msg')}")
    except Exception as e:
        print(f"  请求异常: {e}")

    # 方式2：逐个请求多门店
    print(f"\n=== 方式2: 逐个请求多门店 ===")
    for shop_id in shop_ids:
        print(f"\n  查询门店: {shop_id}...")

        try:
            result = client.get_statistics(shop_id, begin_date, end_date, platform)
            if result.get("code") == 200:
                stats = client.parse_statistics(result)
                shop_name = get_shop_name(shop_id)
                results.append({
                    "shop_id": shop_id,
                    "shop_name": shop_name,
                    "statistics": stats,
                })
                print(f"    {shop_name}: {stats}")
            else:
                print(f"    错误: {result.get('msg')}")
        except Exception as e:
            print(f"    请求异常: {e}")

    # 打印汇总
    print(f"\n{'='*60}")
    print("逐个查询结果汇总")
    print(f"{'='*60}")
    for result in results:
        stats = result['statistics']
        print(f"  {result['shop_name']}: 点评{stats.get('点评星级', '-')}星, 美团{stats.get('美团星级', '-')}星")

    return results


if __name__ == "__main__":
    test_multi_shop_report()

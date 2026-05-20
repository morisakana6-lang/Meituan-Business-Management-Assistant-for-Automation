"""
门店评论多门店ID测试脚本
使用实际payload参数测试多个门店ID，生成汇总报表
"""

import sys
import os

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)

from shop_review.client import ShopReviewClient


def test_multi_shop_report():
    """测试多个门店ID，获取评论数据"""
    print("=" * 60)
    print("门店评论多门店ID测试")
    print("=" * 60)

    # 平台参数：0=全平台，1=点评，2=美团
    platform_names = {0: "全平台", 1: "点评", 2: "美团"}

    # 使用payload中的门店IDs
    shop_ids = ["1933643130", "57904793"]
    # 使用配置中的日期
    begin_date = "2026-05-01"
    end_date = "2026-05-18"
    platform = "0"  # 0=全平台

    print(f"\n测试门店IDs: {shop_ids}")
    print(f"日期范围: {begin_date} 至 {end_date}")
    print(f"平台: {platform} ({platform_names.get(int(platform), '未知')})")

    client = ShopReviewClient()

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

    for shop_id in shop_ids:
        print(f"\n正在查询门店 {shop_id}...")

        try:
            reviews = client.get_reviews(
                shop_id=shop_id,
                platform=platform,
                begin_date=begin_date,
                end_date=end_date
            )

            shop_name = get_shop_name(shop_id)
            total_count = len(reviews)

            results.append({
                "shop_id": shop_id,
                "shop_name": shop_name,
                "total_count": total_count,
                "reviews": reviews,
            })
            print(f"  {shop_name}: 获取成功，共 {total_count} 条评论")

            # 显示前3条评论摘要
            for i, review in enumerate(reviews[:3]):
                detail = client.get_review_detail(review)
                print(f"    [{i+1}] {detail['create_time']} | 评分: {detail['star']}星 | 内容: {detail['content'][:30]}...")

            if len(reviews) > 3:
                print(f"    ... 还有 {len(reviews) - 3} 条评论")

        except Exception as e:
            print(f"  请求异常: {e}")

    # 打印汇总
    print(f"\n{'='*60}")
    print("汇总结果")
    print(f"{'='*60}")
    for result in results:
        print(f"  {result['shop_name']}: {result['total_count']} 条评论")

    return results


if __name__ == "__main__":
    test_multi_shop_report()

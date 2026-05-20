"""
门店评论 完整测试脚本
验证 API 调用 + Excel 生成
"""

import os
import sys
import json

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)

from shop_review.client import ShopReviewClient
from shop_review.excel_generator import generate_report


def load_shop_config():
    """从 shop_config.json 加载配置"""
    config_file = os.path.join(_CURRENT_DIR, 'shop_config.json')
    if os.path.exists(config_file):
        with open(config_file, encoding='utf-8') as f:
            return json.load(f)
    return {}


def test_api():
    """测试API调用"""
    print("=" * 60)
    print("门店评论 API 测试")
    print("=" * 60)

    config = load_shop_config()
    shop_id = config.get("search_key", "1933643130")
    # 支持数组或逗号分隔的字符串
    if isinstance(shop_id, list):
        shop_ids = shop_id
    elif isinstance(shop_id, str) and "," in shop_id:
        shop_ids = [s.strip() for s in shop_id.split(",")]
    else:
        shop_ids = [shop_id]
    shop_id = shop_ids[0]
    platform = config.get("platform", 1)
    date_range = config.get("日期范围", {})
    begin_date = date_range.get("begin")
    end_date = date_range.get("end")

    print(f"\n使用配置:")
    print(f"  门店ID: {shop_ids}")
    print(f"  日期范围: {begin_date} 至 {end_date}")
    print(f"  平台: {platform}")

    client = ShopReviewClient()
    reviews = client.get_reviews(shop_id, str(platform), begin_date, end_date)

    print(f"\n获取到 {len(reviews)} 条评论")

    if reviews:
        detail = client.get_review_detail(reviews[0])
        print(f"\n第一条评论详情:")
        for k, v in detail.items():
            print(f"  {k}: {v}")

    return client, shop_id, platform, begin_date, end_date


def test_excel(client, shop_id, platform, begin_date, end_date):
    """测试Excel生成"""
    print("\n" + "=" * 60)
    print("Excel 生成测试")
    print("=" * 60)

    try:
        filepath = generate_report(
            client=client,
            shop_id=shop_id,
            platform=platform,
            begin_date=begin_date,
            end_date=end_date,
        )
        print(f"\n[OK] Excel报表生成成功!")
        print(f"文件路径: {filepath}")

        if os.path.exists(filepath):
            file_size = os.path.getsize(filepath)
            print(f"文件大小: {file_size} bytes")
    except Exception as e:
        print(f"\n[FAIL] 报表生成失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    client, shop_id, platform, begin_date, end_date = test_api()
    test_excel(client, shop_id, platform, begin_date, end_date)
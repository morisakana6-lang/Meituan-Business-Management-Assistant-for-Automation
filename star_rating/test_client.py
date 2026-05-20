"""
星级评分 完整测试脚本
验证 API 调用 + Excel 生成
"""

import os
import sys
import json

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)

from star_rating.client import StarRatingClient, get_shop_name
from star_rating.excel_generator import generate_report


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
    print("星级评分 API 测试")
    print("=" * 60)

    config = load_shop_config()
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

    # 获取门店名称
    shop_name = get_shop_name(shop_id)

    print(f"\n使用配置:")
    print(f"  门店ID: {shop_ids}")
    print(f"  门店名称: {shop_name}")
    print(f"  日期范围: {begin_date} 至 {end_date}")
    print(f"  平台: {platform}")

    client = StarRatingClient()
    result = client.get_statistics(shop_id, begin_date, end_date, str(platform))
    statistics = client.parse_statistics(result)

    print(f"\n解析后的统计数据:")
    for k, v in statistics.items():
        print(f"  {k}: {v}")

    return client, shop_id, shop_name, platform, begin_date, end_date


def test_excel(client, shop_id, shop_name, platform, begin_date, end_date):
    """测试Excel生成"""
    print("\n" + "=" * 60)
    print("Excel 生成测试")
    print("=" * 60)

    try:
        filepath = generate_report(
            client=client,
            shop_id=shop_id,
            shop_name=shop_name,
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
    client, shop_id, shop_name, platform, begin_date, end_date = test_api()
    test_excel(client, shop_id, shop_name, platform, begin_date, end_date)

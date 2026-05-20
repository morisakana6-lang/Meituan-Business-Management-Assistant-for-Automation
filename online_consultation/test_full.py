"""
在线咨询分析 完整测试脚本
验证 API 调用 + Excel 生成
"""

import os
import sys
import json

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)

from online_consultation.client import OnlineConsultationClient
from online_consultation.excel_generator import generate_report


def load_shop_config():
    """从 shop_config.json 加载配置"""
    config_file = os.path.join(_CURRENT_DIR, 'shop_config.json')
    if os.path.exists(config_file):
        with open(config_file, encoding='utf-8') as f:
            return json.load(f)
    return {}


def test_full_flow():
    """测试完整流程：API调用 + Excel生成"""
    print("=" * 60)
    print("在线咨询分析 完整流程测试")
    print("=" * 60)

    # 从配置文件加载
    config = load_shop_config()
    shop_id = config.get("search_key", "1933643130")
    platform = config.get("platform", 0)
    date_range = config.get("日期范围", {})
    begin_date = date_range.get("begin")
    end_date = date_range.get("end")

    print(f"\n使用配置:")
    print(f"  门店ID (search_key): {shop_id}")
    print(f"  日期范围: {begin_date} 至 {end_date}")
    print(f"  平台: {platform}")

    # 创建客户端
    client = OnlineConsultationClient()

    # 生成报表
    print("\n开始生成Excel报表...")
    try:
        filepath = generate_report(
            client=client,
            shop_id=shop_id,
            shop_name=shop_id,  # 使用shop_id作为门店名称
            platform=platform,
            begin_date=begin_date,
            end_date=end_date,
        )
        print(f"\n[OK] Excel报表生成成功!")
        print(f"文件路径: {filepath}")

        # 检查文件是否存在
        if os.path.exists(filepath):
            file_size = os.path.getsize(filepath)
            print(f"文件大小: {file_size} bytes")
        else:
            print("[FAIL] 文件不存在")

    except Exception as e:
        print(f"\n[FAIL] 报表生成失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_full_flow()
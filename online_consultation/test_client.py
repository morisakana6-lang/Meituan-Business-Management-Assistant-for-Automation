"""
在线咨询分析 API 测试脚本

运行方式：
cd d:/code/meituan
python -c "
import sys
 sys.path.insert(0, '.')
 from online_consultation.test_client import test_online_consultation
 test_online_consultation()
"

或者直接：
python online_consultation/test_client.py
"""

import json
import os
import sys

# 确保项目根目录在路径中
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)

from online_consultation.client import OnlineConsultationClient


def test_online_consultation():
    """测试在线咨询分析 API"""

    print("=" * 60)
    print("在线咨询分析 API 测试")
    print("=" * 60)

    # 测试参数（使用用户截图中的实际值）
    test_shop_id = "1933643130"
    test_begin_date = "2026-05-10"
    test_end_date = "2026-05-16"
    test_platform = "0"

    print(f"\n测试门店ID: {test_shop_id}")
    print(f"测试日期范围: {test_begin_date} 至 {test_end_date}")
    print(f"测试平台: {test_platform}")

    # 加载凭证
    cred_file = os.path.join(_PROJECT_ROOT, 'auth', 'credentials.json')
    print(f"\n凭证文件: {cred_file}")
    print(f"凭证文件存在: {os.path.exists(cred_file)}")

    if os.path.exists(cred_file):
        with open(cred_file, encoding='utf-8') as f:
            cred_data = json.load(f)
            cookie = cred_data.get('cookie', '')
            mtgsig = cred_data.get('mtgsig', '')
            print(f"Cookie 长度: {len(cookie) if cookie else 0}")
            print(f"mtgsig 存在: {bool(mtgsig)}")
            if mtgsig:
                print(f"mtgsig 前50字符: {mtgsig[:50]}...")

    # 创建客户端
    print("\n创建客户端...")
    try:
        client = OnlineConsultationClient()
        print(f"  - Base URL: {client.BASE_URL}")
        print(f"  - Cookie 加载: {'成功' if client.cookie else '失败'}")
        print(f"  - mtgsig 加载: {'成功' if client.mtgsig else '失败'}")
    except Exception as e:
        print(f"  - 客户端创建失败: {e}")
        return

    # 发送请求
    print("\n发送请求...")
    try:
        response = client.get_report(
            shop_id=test_shop_id,
            platform=test_platform,
            begin_date=test_begin_date,
            end_date=test_end_date,
        )
        print(f"\n响应结果:")
        print(json.dumps(response, ensure_ascii=False, indent=2))
    except Exception as e:
        print(f"\n请求失败: {e}")
        return

    # 解析结果
    print("\n" + "=" * 60)
    print("结果解析")
    print("=" * 60)

    if response.get("code") == 0:
        print("[OK] API调用成功")

        try:
            metrics = client.get_metrics(
                shop_id=test_shop_id,
                platform=test_platform,
                begin_date=test_begin_date,
                end_date=test_end_date,
            )
            print("\n指标数据:")
            for name, value in metrics.items():
                print(f"  - {name}: {value}")
        except Exception as e:
            print(f"\n指标解析失败: {e}")
    else:
        print(f"[FAIL] API返回错误: code={response.get('code')}, msg={response.get('msg')}")


if __name__ == "__main__":
    test_online_consultation()
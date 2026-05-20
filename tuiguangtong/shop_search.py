"""
门店搜索模块

=== 主要作用 ===
提供门店查询功能，支持：
1. 按门店ID搜索（返回名称和平台类型）
2. 按门店名称搜索（精确匹配，返回ID列表）
3. platform 参数筛选（0=全平台，1=点评，2=美团）

=== 板块归属 ===
推广通板块，对应目录 tuiguangtong/

=== 核心函数 ===
search_by_id(): 按ID搜索门店
search_by_name(): 按名称搜索门店
resolve_shop(): 综合搜索，自动判断输入是ID还是名称

=== 使用方法 ===
from shop_search import search_by_id, search_by_name, resolve_shop

# 按ID搜索
result = search_by_id("764510450")
# -> {"name": "丽减美瘦吧·减肥瘦身(新澳城店)", "ids": [{"type": "meituan", "id": "764510450"}]}

# 按名称搜索
result = search_by_name("丽减美瘦吧·减肥瘦身(新澳城店)")
# -> {"name": "...", "ids": [...]}

# 自动判断并搜索
name, shop_id = resolve_shop("764510450", platform=0)
# -> ("丽减美瘦吧·减肥瘦身(新澳城店)", "764510450")

=== 配置文件 ===
shop_mapping.json: 门店映射表，包含所有门店的name和ids信息
"""

import json
import os
from typing import Dict, List, Tuple, Optional

# 获取当前文件所在目录
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
MAPPING_FILE = os.path.join(_CURRENT_DIR, "shop_mapping.json")


def load_mapping() -> Dict:
    """加载门店映射表"""
    if not os.path.exists(MAPPING_FILE):
        raise FileNotFoundError(f"门店映射表不存在: {MAPPING_FILE}")
    with open(MAPPING_FILE, encoding="utf-8") as f:
        return json.load(f)


def is_shop_id(s: str) -> bool:
    """
    判断输入是否为门店ID

    判断依据：纯数字字符串视为ID
    """
    return s.strip().isdigit()


def search_by_id(shop_id: str) -> List[Dict]:
    """
    按门店ID搜索

    Args:
        shop_id: 门店ID

    Returns:
        匹配的门店列表，每项包含 name 和 ids
        例如: [{"name": "...", "ids": [{"type": "meituan", "id": "xxx"}]}]
    """
    mapping = load_mapping()
    matches = []

    for shop in mapping.get("shops", []):
        for id_info in shop.get("ids", []):
            if id_info.get("id") == shop_id:
                # 避免重复添加同一门店
                if not any(s["name"] == shop["name"] for s in matches):
                    matches.append({"name": shop["name"], "ids": shop["ids"]})
                break

    return matches


def search_by_name(shop_name: str) -> Optional[Dict]:
    """
    按门店名称搜索（精确匹配）

    Args:
        shop_name: 门店名称

    Returns:
        匹配的门店信息，包含 name 和 ids
        如果未找到返回 None
    """
    mapping = load_mapping()

    for shop in mapping.get("shops", []):
        if shop.get("name") == shop_name:
            return {"name": shop["name"], "ids": shop["ids"]}

    return None


def select_id_by_platform(ids: List[Dict], platform: int) -> Optional[Dict]:
    """
    根据 platform 参数筛选 ID

    Args:
        ids: ID列表
        platform: 0=全平台，1=点评，2=美团

    Returns:
        筛选后的ID信息，如果没找到返回 None
    """
    platform_map = {1: "dianping", 2: "meituan"}
    target_type = platform_map.get(platform, "dianping")

    for id_info in ids:
        if id_info.get("type") == target_type:
            return id_info

    return None


def prompt_shop_selection(matches: List[Dict]) -> Dict:
    """
    当多个门店匹配时，让用户选择

    Args:
        matches: 匹配的门店列表

    Returns:
        用户选择的门店
    """
    print(f"\n搜索到多个匹配门店，请选择：")
    for i, shop in enumerate(matches, 1):
        print(f"  [{i}] {shop['name']}")

    while True:
        try:
            choice = input("请输入序号: ").strip()
            if choice == "":
                idx = 0
            else:
                idx = int(choice) - 1

            if 0 <= idx < len(matches):
                return matches[idx]
            else:
                print(f"无效选择，请输入 1-{len(matches)} 之间的数字")
        except ValueError:
            print("请输入有效数字")


def prompt_platform_selection(ids: List[Dict]) -> Dict:
    """
    当多个平台ID时，让用户选择

    Args:
        ids: ID列表

    Returns:
        用户选择的ID信息
    """
    print(f"\n该门店有多个平台ID，请选择：")
    for i, id_info in enumerate(ids, 1):
        platform_name = "点评" if id_info.get("type") == "dianping" else "美团"
        print(f"  [{i}] {platform_name} ID: {id_info.get('id')}")

    while True:
        try:
            choice = input("请选择平台 (默认1=点评): ").strip()
            if choice == "":
                idx = 1  # 默认选择点评（第二个）
            else:
                idx = int(choice) - 1

            if 0 <= idx < len(ids):
                return ids[idx]
            else:
                print(f"无效选择，请输入 1-{len(ids)} 之间的数字")
        except ValueError:
            print("请输入有效数字")


def resolve_shop(search_key: str, platform: int = 0) -> Tuple[str, str]:
    """
    综合搜索入口，自动判断输入是ID还是名称

    Args:
        search_key: 搜索关键字（ID或名称），传入"0"表示全部门店汇总
        platform: 平台选择，0=全平台，1=点评，2=美团

    Returns:
        Tuple[str, str]: (门店名称, 最终使用的门店ID)

    Raises:
        ValueError: 门店不存在或ID不存在时
    """
    # shop_ids=0 表示全部门店汇总
    if search_key == "0":
        return ("全部门店汇总", "0")

    if is_shop_id(search_key):
        # 按ID搜索
        matches = search_by_id(search_key)

        if not matches:
            # ID不在mapping中（新门店等情况），但仍然有效
            # 直接使用用户输入的ID，不受platform影响
            return (f"门店ID:{search_key}", search_key)

        if len(matches) > 1:
            # 多个门店匹配，让用户选择
            selected_shop = prompt_shop_selection(matches)
        else:
            selected_shop = matches[0]

        shop_name = selected_shop["name"]
        # 直接使用用户输入的ID，不受platform影响
        # 这是因为用户明确指定了某个平台的ID，不应该被platform参数覆盖
        return shop_name, search_key

    else:
        # 按名称搜索
        shop = search_by_name(search_key)

        if shop is None:
            raise ValueError(f"名称不存在: {search_key}")

        shop_name = shop["name"]
        ids = shop["ids"]

    # 根据platform筛选ID
    if len(ids) == 1:
        selected_id = ids[0]["id"]
    elif platform == 0:
        # 全平台且有多个ID时，让用户选择
        selected_id = prompt_platform_selection(ids)["id"]
    else:
        # 多个ID时，尝试按platform筛选
        id_info = select_id_by_platform(ids, platform)
        if id_info is None:
            # platform对应的ID不存在，让用户选择
            selected_id = prompt_platform_selection(ids)["id"]
        else:
            selected_id = id_info["id"]

    return shop_name, selected_id


def get_all_shop_names() -> List[str]:
    """获取所有门店名称（用于自动补全等场景）"""
    mapping = load_mapping()
    return [shop["name"] for shop in mapping.get("shops", [])]


if __name__ == "__main__":
    # 测试代码
    print("=== 门店搜索测试 ===\n")

    # 测试1: 按ID搜索
    print("测试1: 按ID搜索 '764510450'")
    matches = search_by_id("764510450")
    print(f"  结果: {matches}\n")

    # 测试2: 按名称搜索
    print("测试2: 按名称搜索 '丽减美瘦吧·减肥瘦身(新澳城店)'")
    shop = search_by_name("丽减美瘦吧·减肥瘦身(新澳城店)")
    print(f"  结果: {shop}\n")

    # 测试3: resolve_shop - ID
    print("测试3: resolve_shop('764510450', platform=0)")
    try:
        name, shop_id = resolve_shop("764510450", platform=0)
        print(f"  结果: name={name}, shop_id={shop_id}")
    except ValueError as e:
        print(f"  错误: {e}")
    print()

    # 测试4: resolve_shop - 名称
    print("测试4: resolve_shop('丽减美瘦吧·减肥瘦身(新澳城店)', platform=0)")
    try:
        name, shop_id = resolve_shop("丽减美瘦吧·减肥瘦身(新澳城店)", platform=0)
        print(f"  结果: name={name}, shop_id={shop_id}")
    except ValueError as e:
        print(f"  错误: {e}")
    print()

    # 测试5: 双ID门店
    print("测试5: resolve_shop('丽减美瘦吧·减肥瘦身（宝云街店）', platform=0)")
    try:
        name, shop_id = resolve_shop("丽减美瘦吧·减肥瘦身（宝云街店）", platform=0)
        print(f"  结果: name={name}, shop_id={shop_id}")
    except ValueError as e:
        print(f"  错误: {e}")

"""
简版看板Excel生成器

=== 板块归属 ===
简版看板板块，对应目录 simple_board/

=== 样式说明 ===
- 第1行: 大标题，简版看板(D-I列)橙色背景，流量数据(J-M列)橙色背景
- 第2行: 表头，灰色背景
- 第3行起: 数据行，白色背景
- 字体: 微软雅黑
- 数值右对齐，文本左对齐
"""

import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from typing import Dict

import sys
import os

# 获取项目根目录
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_CURRENT_DIR)
sys.path.insert(0, _PROJECT_ROOT)


# ============== 样式定义 ==============

# 字体
TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)  # 大标题
SECTION_FONT = Font(name="微软雅黑", bold=True, size=11)  # 分组标题
HEADER_FONT = Font(name="微软雅黑", bold=True, size=10)  # 表头
DATA_FONT = Font(name="微软雅黑", size=10)  # 数据

# 填充色
TITLE_FILL = PatternFill(start_color="FFE3F1D9", end_color="FFE3F1D9", fill_type="solid")  # 绿色（简版看板标题行）
FLOW_FILL = PatternFill(start_color="FFFBE5D5", end_color="FFFBE5D5", fill_type="solid")  # 橙色（流量数据标题行）
HEADER_FILL = PatternFill(start_color="FFE8E8E8", end_color="FFE8E8E8", fill_type="solid")  # 灰色（表头行）
WHITE_FILL = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")  # 白色（数据行）

# 对齐
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")
RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")

# 边框
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)


class ExcelGenerator:
    """简版看板Excel生成器"""

    def __init__(self, account_name: str = "简版看板"):
        self.account_name = account_name
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "简版看板报表"

    def _create_title_row(self, title: str = None):
        """创建标题行（第1行）"""
        # 简版看板分组标题 (A1:I1) - 包含门店ID、推广门店、时间
        self.ws.merge_cells("A1:I1")
        title_cell = self.ws["A1"]
        title_cell.value = title if title else "简版看板"
        title_cell.font = TITLE_FONT
        title_cell.fill = TITLE_FILL  # 绿色
        title_cell.alignment = CENTER_ALIGN

        # 流量分组标题 (J1:M1)
        self.ws.merge_cells("J1:M1")
        flow_cell = self.ws["J1"]
        flow_cell.value = "流量数据"
        flow_cell.font = TITLE_FONT
        flow_cell.fill = FLOW_FILL  # 橙色
        flow_cell.alignment = CENTER_ALIGN

        # 设置第1行行高
        self.ws.row_dimensions[1].height = 28

    def _create_header_row(self):
        """创建表头行（第2行）"""
        headers = [
            "门店ID",           # A
            "推广门店",         # B
            "时间",             # C
            "曝光人数",         # D - 简版看板
            "访问人数",         # E
            "下单券数",         # F
            "下单金额(原价)",   # G
            "核销券数",         # H
            "核销金额(原价)",   # I
            "曝光次数",         # J - 流量数据
            "曝光人数",         # K
            "访问次数",         # L
            "访问人数",         # M
        ]

        widths = [21.125, 28.875, 25.375, 10.0, 13.0, 13.0, 14.0, 10.0, 14.0, 10.0, 13.0, 13.0, 13.0]

        for col_idx, (header, width) in enumerate(zip(headers, widths), start=1):
            cell = self.ws.cell(row=2, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
            self.ws.column_dimensions[get_column_letter(col_idx)].width = width

        # 设置第2行行高
        self.ws.row_dimensions[2].height = 22

    def _get_widths(self):
        """获取列宽列表"""
        return [21.125, 28.875, 25.375, 10.0, 13.0, 13.0, 14.0, 10.0, 14.0, 10.0, 13.0, 13.0, 13.0]

    def _get_headers(self):
        """获取表头列表"""
        return [
            "门店ID",           # A
            "推广门店",         # B
            "时间",             # C
            "曝光人数",         # D - 简版看板
            "访问人数",         # E
            "下单券数",         # F
            "下单金额(原价)",   # G
            "核销券数",         # H
            "核销金额(原价)",   # I
            "曝光次数",         # J - 流量数据
            "曝光人数",         # K
            "访问次数",         # L
            "访问人数",         # M
        ]

    def _format_value(self, value: str) -> str:
        """格式化数值，移除逗号"""
        if isinstance(value, str):
            return value.replace(",", "")
        return value

    def add_data_row(
        self,
        shop_id: str,
        shop_name: str,
        date_str: str,
        metrics: Dict
    ):
        """
        添加一行数据

        Args:
            shop_id: 门店ID
            shop_name: 门店名称
            date_str: 时间描述（如"3天"）
            metrics: 指标字典，包含 overview, trade, flow 三类数据
        """
        overview = metrics.get("overview", {})
        trade = metrics.get("trade", {})
        flow = metrics.get("flow", {})

        row_data = [
            shop_id,                                           # A 门店ID
            shop_name,                                         # B 推广门店
            date_str,                                          # C 时间
            self._format_value(overview.get("exposure_count", "0")),   # D 曝光人数
            self._format_value(overview.get("visitor_count", "0")),   # E 访问人数
            self._format_value(trade.get("order_count", "0")),         # F 下单券数
            self._format_value(trade.get("order_amount", "0")),       # G 下单金额(原价)
            self._format_value(trade.get("verify_count", "0")),       # H 核销券数
            self._format_value(trade.get("verify_amount", "0")),       # I 核销金额(原价)
            self._format_value(flow.get("exposure_times", "0")),      # J 曝光次数
            self._format_value(flow.get("exposure_count", "0")),      # K 曝光人数
            self._format_value(flow.get("visit_times", "0")),         # L 访问次数
            self._format_value(flow.get("visitor_count", "0")),        # M 访问人数
        ]

        row_idx = self.ws.max_row + 1

        for col_idx, value in enumerate(row_data, start=1):
            cell = self.ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = DATA_FONT
            cell.fill = WHITE_FILL
            cell.border = THIN_BORDER

                    # 全部居中对齐
            cell.alignment = CENTER_ALIGN

        # 设置数据行行高
        self.ws.row_dimensions[row_idx].height = 18

    def save(self, filepath: str):
        """保存Excel文件"""
        # 检查文件是否被占用
        if os.path.exists(filepath):
            try:
                with open(filepath, 'a'):
                    pass
            except IOError:
                print(f"错误: 文件被占用，请关闭 '{filepath}' 后再重试")
                raise IOError(f"文件被占用: {filepath}")
        self.wb.save(filepath)
        print(f"Excel 文件已保存: {filepath}")


def generate_report(
    client,
    shop_ids: list,
    begin_date: str,
    end_date: str,
    platform: int = 0
) -> str:
    """
    生成简版看板报表（支持单门店和多门店）

    Args:
        client: SimpleBoardClient 实例
        shop_ids: 门店ID列表，支持单个或多个门店
        begin_date: 开始日期，格式 YYYY-MM-DD
        end_date: 结束日期，格式 YYYY-MM-DD
        platform: 平台选择，0=全平台，1=点评，2=美团

    Returns:
        生成的 Excel 文件路径
    """
    # 加载门店映射表
    from tuiguangtong.shop_search import load_mapping
    mapping_data = load_mapping()
    shops = mapping_data.get("shops", [])

    def get_shop_name(shop_id: str) -> str:
        if shop_id == "0":
            return "全部门店汇总数据"
        for shop in shops:
            for id_info in shop.get("ids", []):
                if id_info.get("id") == shop_id:
                    return shop.get("name", f"门店{shop_id}")
        return f"门店{shop_id}"

    generator = ExcelGenerator()

    # 判断单门店还是多门店
    is_multi_shop = len(shop_ids) > 1
    if is_multi_shop:
        generator._create_title_row("简版看板报表（多门店）")
    else:
        generator._create_title_row()

    generator._create_header_row()

    # 日期范围
    date_str = f"{begin_date}至{end_date}"
    print(f"开始生成简版看板报表: {begin_date} ~ {end_date}")

    # 遍历每个门店查询数据
    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  查询门店: {shop_name}...")

        try:
            metrics = client.get_metrics(
                begin_date=begin_date,
                end_date=end_date,
                shop_id=shop_id,
                platform=str(platform)
            )
            generator.add_data_row(shop_id, shop_name, date_str, metrics)
            print(f"    {shop_name}: 获取成功")
        except Exception as e:
            print(f"    {shop_name}: 获取失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_str, {})

    # 保存文件
    output_dir = os.path.join(_CURRENT_DIR, "reports")
    os.makedirs(output_dir, exist_ok=True)

    if is_multi_shop:
        output_file = f"{output_dir}/简版看板_多门店_{begin_date}_{end_date}.xlsx"
    else:
        output_file = f"{output_dir}/简版看板_{shop_ids[0]}_{begin_date}_{end_date}.xlsx"

    generator.save(output_file)
    return output_file


if __name__ == "__main__":
    """
    入口：直接运行 python excel_generator.py 即可生成报表

    使用方法：
        1. 修改 shop_config.json 中的 search_key（门店ID或名称）和日期
        2. 运行 python excel_generator.py

    search_key: 可以是单个门店ID、多个门店ID列表（JSON数组格式）或逗号分隔字符串
    platform: 0=全平台（默认），1=点评，2=美团
    """
    import json
    from client import SimpleBoardClient

    # 复用推广通的门店搜索模块
    _SHOP_SEARCH_PATH = os.path.join(_PROJECT_ROOT, 'tuiguangtong', 'shop_search.py')
    sys.path.insert(0, os.path.dirname(_SHOP_SEARCH_PATH))
    from shop_search import is_shop_id, search_by_id, resolve_shop

    # 读取配置
    config_path = os.path.join(os.path.dirname(__file__), "shop_config.json")
    with open(config_path, encoding="utf-8") as f:
        config = json.load(f)

    search_key = config.get("search_key", "")
    platform = config.get("platform", 0)
    date_range = config["日期范围"]

    # 解析 search_key（支持多种格式）
    if isinstance(search_key, list):
        shop_ids = search_key
    elif isinstance(search_key, str):
        if "," in search_key:
            shop_ids = [s.strip() for s in search_key.split(",")]
        else:
            shop_ids = [search_key]
    else:
        shop_ids = [str(search_key)]

    # 获取门店信息
    if len(shop_ids) == 1:
        if shop_ids[0] == "0":
            shop_name = "全部门店汇总数据"
        elif is_shop_id(str(shop_ids[0])):
            matches = search_by_id(str(shop_ids[0]))
            if matches:
                shop_name = matches[0]["name"]
            else:
                shop_name = f"门店ID:{shop_ids[0]}"
        else:
            shop_name, _ = resolve_shop(str(shop_ids[0]), platform)
        print(f"已选择门店: {shop_name}")
    else:
        first_key = shop_ids[0]
        if is_shop_id(str(first_key)):
            matches = search_by_id(str(first_key))
            if matches:
                shop_name = matches[0]["name"]
            else:
                shop_name = f"门店ID:{first_key}"
        else:
            shop_name, _ = resolve_shop(str(first_key), platform)
        print(f"已选择门店数量: {len(shop_ids)}")

    platform_names = {0: "全平台", 1: "点评", 2: "美团"}
    print(f"平台: {platform_names.get(platform, '全平台')}")
    print()

    client = SimpleBoardClient()

    generate_report(
        client=client,
        shop_ids=shop_ids,
        begin_date=date_range["begin"],
        end_date=date_range["end"],
        platform=platform
    )
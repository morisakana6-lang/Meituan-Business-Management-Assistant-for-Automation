"""
美团经营宝汇总报表生成器

将所有板块的报表合并到单个 Excel 文件中，每个板块一个 Sheet。

使用方法：
    python summary_report_generator.py

输出：
    config/reports/美团经营宝汇总_YYYY-MM-DD.xlsx
"""

import os
import sys
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# 获取项目根目录
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = _CURRENT_DIR
sys.path.insert(0, _PROJECT_ROOT)


# ============== 样式定义 ==============

TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
HEADER_FONT = Font(name="微软雅黑", bold=True, size=10)
DATA_FONT = Font(name="微软雅黑", size=10)

TITLE_FILL = PatternFill(start_color="FFFBE5D5", end_color="FFFBE5D5", fill_type="solid")
HEADER_FILL = PatternFill(start_color="FFE8E8E8", end_color="FFE8E8E8", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")
ERROR_FILL = PatternFill(start_color="FFFFE6E6", end_color="FFFFE6E6", fill_type="solid")
FLOW_FILL = PatternFill(start_color="FFE3F1D9", end_color="FFE3F1D9", fill_type="solid")  # 绿色（简版看板流量数据标题行）

CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")
RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)


class SummaryExcelGenerator:
    """汇总报表生成器"""

    def __init__(self):
        self.wb = Workbook()
        # 删除默认 sheet
        if "Sheet" in self.wb.sheetnames:
            self.wb.remove(self.wb["Sheet"])

    def add_sheet_from_generator(self, sheet_name: str, generator, title: str, headers: list, widths: list, data_rows: list, row_heights: list = None, title_style: str = "default"):
        """
        从已有 ExcelGenerator 的数据创建 Sheet

        Args:
            sheet_name: Sheet 名称
            generator: ExcelGenerator 实例（用于获取样式）
            title: 标题
            headers: 表头列表
            widths: 列宽列表
            data_rows: 数据行列表（每行是列表）
            row_heights: 每行的行高列表，默认 [28, 22, 18, ...]
            title_style: 标题样式，"default"=统一样式，"grouped"=分组样式（简版看板）
        """
        ws = self.wb.create_sheet(title=sheet_name)

        # 确定列数
        num_cols = max(len(headers), len(widths), 1)
        if data_rows:
            num_cols = max(num_cols, len(data_rows[0]))

        # 行高设置：默认 [标题行=28, 表头行=22, 数据行=18]
        if row_heights is None:
            row_heights = [28, 22] + [18] * len(data_rows) if data_rows else [28, 22]

        # 标题行
        last_col = get_column_letter(num_cols) if num_cols > 0 else "A"
        if title_style == "grouped" and num_cols >= 13:
            # 分组标题样式（简版看板）：A-I 绿色（FFE3F1D9），J-M 橙色（FFFBE5D5）
            ws.merge_cells(f"A1:I1")
            title_cell = ws["A1"]
            title_cell.value = title
            title_cell.font = TITLE_FONT
            title_cell.fill = FLOW_FILL  # 绿色
            title_cell.alignment = CENTER_ALIGN

            ws.merge_cells(f"J1:{last_col}1")
            flow_cell = ws["J1"]
            flow_cell.value = "流量数据"
            flow_cell.font = TITLE_FONT
            flow_cell.fill = TITLE_FILL  # 橙色
            flow_cell.alignment = CENTER_ALIGN
        else:
            # 默认统一标题样式
            ws.merge_cells(f"A1:{last_col}1")
            title_cell = ws["A1"]
            title_cell.value = title
            title_cell.font = TITLE_FONT
            title_cell.fill = TITLE_FILL
            title_cell.alignment = CENTER_ALIGN

        ws.row_dimensions[1].height = row_heights[0] if len(row_heights) > 0 else 28

        # 表头行
        for col_idx, (header, width) in enumerate(zip(headers, widths), start=1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        ws.row_dimensions[2].height = row_heights[1] if len(row_heights) > 1 else 22

        # 数据行
        for row_idx, row_data in enumerate(data_rows, start=3):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = DATA_FONT
                cell.fill = WHITE_FILL
                cell.border = THIN_BORDER
                cell.alignment = CENTER_ALIGN
            # 使用传入的行高，第3行对应 row_heights[2]，以此类推
            height_idx = row_idx - 1  # row_idx=3 -> idx=2
            ws.row_dimensions[row_idx].height = row_heights[height_idx] if height_idx < len(row_heights) else 18

    def add_error_sheet(self, sheet_name: str, module_name: str, error_msg: str):
        """添加错误提示 Sheet"""
        ws = self.wb.create_sheet(title=sheet_name)

        # 标题
        ws.merge_cells("A1:B1")
        title_cell = ws["A1"]
        title_cell.value = f"{module_name} - 数据获取失败"
        title_cell.font = TITLE_FONT
        title_cell.fill = ERROR_FILL
        title_cell.alignment = CENTER_ALIGN
        ws.row_dimensions[1].height = 28

        # 错误信息
        ws["A2"] = "错误信息"
        ws["A2"].font = HEADER_FONT
        ws["A2"].fill = HEADER_FILL
        ws["A2"].border = THIN_BORDER
        ws["A2"].alignment = CENTER_ALIGN

        ws["B2"] = error_msg
        ws["B2"].font = DATA_FONT
        ws["B2"].fill = WHITE_FILL
        ws["B2"].border = THIN_BORDER
        ws["B2"].alignment = LEFT_ALIGN

        ws.column_dimensions["A"].width = 15
        ws.column_dimensions["B"].width = 60

    def save(self, filepath: str):
        """保存 Excel 文件"""
        # 检查文件是否被占用
        if os.path.exists(filepath):
            try:
                with open(filepath, 'a'):
                    pass
            except IOError:
                print(f"错误: 文件被占用，请关闭 '{filepath}' 后再重试")
                raise IOError(f"文件被占用: {filepath}")
        self.wb.save(filepath)
        print(f"汇总报表已保存: {filepath}")


def load_main_config():
    """加载主配置文件"""
    config_path = os.path.join(_PROJECT_ROOT, "config", "main_config.json")
    with open(config_path, encoding="utf-8") as f:
        return json.load(f)


def get_shop_name(shop_id: str) -> str:
    """从门店映射表获取门店名称"""
    from tuiguangtong.shop_search import load_mapping
    mapping_data = load_mapping()
    shops = mapping_data.get("shops", [])

    if shop_id == "0":
        return "全部门店汇总数据"

    for shop in shops:
        for id_info in shop.get("ids", []):
            if id_info.get("id") == shop_id:
                return shop.get("name", f"门店{shop_id}")
    return f"门店{shop_id}"


# ============== 各板块数据获取 ==============

def generate_tuiguangtong_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """推广通汇总数据"""
    from tuiguangtong.client import TuiguangtongClient
    from tuiguangtong.excel_generator import ExcelGenerator

    client = TuiguangtongClient()
    generator = ExcelGenerator()
    generator._create_title_row("推广通报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            metrics = client.get_metrics(
                begin_date=begin_date,
                end_date=end_date,
                shop_ids=shop_id,
                platform=platform
            )
            generator.add_data_row(shop_id, shop_name, date_range, metrics)
            print("成功")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, [])
            fail_count += 1

    return generator, success_count, fail_count


def generate_review_statistics_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """评论统计汇总数据"""
    from review_statistics.client import ReviewStatisticsClient
    from review_statistics.excel_generator import ExcelGenerator

    client = ReviewStatisticsClient()
    generator = ExcelGenerator()
    generator._create_title_row("评论统计报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            result = client.get_statistics(shop_id, begin_date, end_date, str(platform))
            if result.get("code") == 200:
                statistics = client.parse_statistics(result)
                generator.add_data_row(shop_id, shop_name, date_range, statistics)
                print(f"成功 - 累计评价 {statistics.get('累计评价数', 0)}")
                success_count += 1
            else:
                print(f"失败 - {result.get('msg')}")
                generator.add_data_row(shop_id, shop_name, date_range, {})
                fail_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, {})
            fail_count += 1

    return generator, success_count, fail_count


def generate_star_rating_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """星级评分汇总数据"""
    from star_rating.client import StarRatingClient
    from star_rating.excel_generator import ExcelGenerator

    client = StarRatingClient()
    generator = ExcelGenerator()
    generator._create_title_row("星级评分报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            result = client.get_statistics(shop_id, begin_date, end_date, str(platform))
            if result.get("code") == 200:
                statistics = client.parse_statistics(result)
                generator.add_data_row(shop_id, shop_name, date_range, statistics)
                print(f"成功 - 点评{statistics.get('点评星级', '-')}星, 美团{statistics.get('美团星级', '-')}星")
                success_count += 1
            else:
                print(f"失败 - {result.get('msg')}")
                generator.add_data_row(shop_id, shop_name, date_range, {})
                fail_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, {})
            fail_count += 1

    return generator, success_count, fail_count


def generate_shop_review_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """门店评论汇总数据"""
    from shop_review.client import ShopReviewClient
    from shop_review.excel_generator import ExcelGenerator

    client = ShopReviewClient()
    generator = ExcelGenerator()
    generator._create_title_row("门店评论报表")
    generator._create_header_row()

    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            reviews = client.get_reviews(shop_id, str(platform), begin_date, end_date)
            # 获取平台名称
            platform_name = {0: "全平台", 1: "点评", 2: "美团"}.get(platform, str(platform))
            for review in reviews:
                detail = client.get_review_detail(review)
                generator.add_review_row(detail, platform_name)
            print(f"成功({len(reviews)}条)")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            fail_count += 1

    return generator, success_count, fail_count


def generate_online_consultation_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """在线咨询汇总数据"""
    from online_consultation.client import OnlineConsultationClient
    from online_consultation.excel_generator import ExcelGenerator

    client = OnlineConsultationClient()
    generator = ExcelGenerator()
    generator._create_title_row("在线咨询分析报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            metrics = client.get_metrics(
                shop_id=shop_id,
                platform=str(platform),
                begin_date=begin_date,
                end_date=end_date
            )
            generator.add_data_row(shop_id, shop_name, date_range, metrics)
            print("成功")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, {})
            fail_count += 1

    return generator, success_count, fail_count


def generate_trade_analysis_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """交易分析汇总数据"""
    from trade_analysis.client import TradeAnalysisClient
    from trade_analysis.excel_generator import ExcelGenerator

    client = TradeAnalysisClient()
    generator = ExcelGenerator()
    generator._create_title_row("交易分析报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            metrics = client.get_metrics(
                shop_id=shop_id,
                begin_date=begin_date,
                end_date=end_date,
                platform=str(platform)
            )
            generator.add_data_row(shop_id, shop_name, date_range, metrics)
            print("成功")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, {})
            fail_count += 1

    return generator, success_count, fail_count


def generate_customer_flow_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """客流统计汇总数据"""
    from customer_flow.client import CustomerFlowClient
    from customer_flow.excel_generator import ExcelGenerator

    client = CustomerFlowClient()
    generator = ExcelGenerator()
    generator._create_title_row("客流分析报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            metrics = client.get_metrics(
                shop_id=shop_id,
                begin_date=begin_date,
                end_date=end_date,
                platform=platform
            )
            generator.add_data_row(shop_id, shop_name, date_range, metrics)
            print("成功")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, {})
            fail_count += 1

    return generator, success_count, fail_count


def generate_quantianzhan_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """全站推广汇总数据"""
    from quantianzhan.client import QuantianzhanClient
    from quantianzhan.excel_generator import ExcelGenerator

    client = QuantianzhanClient()
    generator = ExcelGenerator()
    generator._create_title_row("全站推广报表")
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            metrics = client.get_metrics(
                begin_date=begin_date,
                end_date=end_date,
                platform=str(platform),
                shop_ids=shop_id
            )
            generator.add_data_row(shop_id, shop_name, date_range, metrics)
            print("成功")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, [])
            fail_count += 1

    return generator, success_count, fail_count


def generate_simple_board_summary(shop_ids: list, begin_date: str, end_date: str, platform: int):
    """简版看板汇总数据"""
    from simple_board.client import SimpleBoardClient
    from simple_board.excel_generator import ExcelGenerator

    client = SimpleBoardClient()
    generator = ExcelGenerator()
    generator._create_title_row()
    generator._create_header_row()

    date_range = f"{begin_date}至{end_date}"
    success_count = 0
    fail_count = 0

    for shop_id in shop_ids:
        shop_name = get_shop_name(shop_id)
        print(f"  [{shop_name}]...", end=" ")
        try:
            metrics = client.get_metrics(
                begin_date=begin_date,
                end_date=end_date,
                shop_id=shop_id,
                platform=str(platform)
            )
            generator.add_data_row(shop_id, shop_name, date_range, metrics)
            print("成功")
            success_count += 1
        except Exception as e:
            print(f"失败 - {e}")
            generator.add_data_row(shop_id, shop_name, date_range, {})
            fail_count += 1

    return generator, success_count, fail_count


# ============== 主流程 ==============

def generate_summary_report(config: dict = None) -> str:
    """
    生成汇总报表

    Args:
        config: 可选，配置字典。如果为 None，则从 main_config.json 加载。
               格式: {
                   "门店列表": ["id1", "id2"],
                   "平台": 0,
                   "门店评论平台": 1,
                   "日期范围": {"begin": "2026-01-01", "end": "2026-03-31"}
               }

    Returns:
        str: 生成的文件路径
    """
    print("=" * 60)
    print("美团经营宝汇总报表生成器")
    print("=" * 60)

    # 加载配置
    if config is None:
        config = load_main_config()
    shop_ids = config.get("门店列表", [])
    platform = config.get("平台", 0)
    shop_review_platform = config.get("门店评论平台", 2)  # 门店评论专用平台
    date_range = config.get("日期范围", {})
    begin_date = date_range.get("begin")
    end_date = date_range.get("end")

    print(f"\n配置信息:")
    print(f"  门店数量: {len(shop_ids)}")
    print(f"  平台: {platform}")
    print(f"  门店评论平台: {shop_review_platform}")
    print(f"  日期: {begin_date} 至 {end_date}")
    print()

    # 创建汇总生成器
    summary_generator = SummaryExcelGenerator()

    # 定义板块顺序和生成函数（使用标准名称）
    modules = [
        ("推广通", generate_tuiguangtong_summary),
        ("评论统计", generate_review_statistics_summary),
        ("星级评分", generate_star_rating_summary),
        ("门店评论", generate_shop_review_summary),
        ("在线咨询分析", generate_online_consultation_summary),
        ("交易分析", generate_trade_analysis_summary),
        ("客流统计", generate_customer_flow_summary),
        ("全站推广", generate_quantianzhan_summary),
        ("简版看板", generate_simple_board_summary),
    ]

    # 遍历各板块
    for module_name, generate_func in modules:
        print(f"\n[{module_name}]")
        try:
            # 门店评论使用专用平台
            effective_platform = shop_review_platform if module_name == "门店评论" else platform
            generator, success_count, fail_count = generate_func(
                shop_ids, begin_date, end_date, effective_platform
            )
            # 从生成器获取数据（统一从worksheet读取）
            headers = generator._get_headers() if hasattr(generator, '_get_headers') else []
            widths = generator._get_widths() if hasattr(generator, '_get_widths') else []
            data_rows = []
            for row in generator.ws.iter_rows(min_row=3, values_only=True):
                if any(cell is not None for cell in row):
                    data_rows.append(list(row))

            # 根据板块特性设置特殊参数
            extra_params = {}
            if module_name == "门店评论":
                # 门店评论需要更大的行高（评论内容换行）
                extra_params["row_heights"] = [28, 22] + [78] * len(data_rows) if data_rows else [28, 22]
            elif module_name == "简版看板":
                # 简版看板使用分组标题样式
                extra_params["title_style"] = "grouped"

            summary_generator.add_sheet_from_generator(
                sheet_name=module_name,
                generator=generator,
                title=f"{module_name}报表",
                headers=headers,
                widths=widths,
                data_rows=data_rows,
                **extra_params
            )
            print(f"  → 添加 Sheet 成功 (成功:{success_count}, 失败:{fail_count})")
        except Exception as e:
            print(f"  → 添加 Sheet 失败: {e}")
            summary_generator.add_error_sheet(module_name, module_name, str(e))

    # 保存文件
    output_dir = os.path.join(_PROJECT_ROOT, "config", "reports")
    os.makedirs(output_dir, exist_ok=True)

    # 判断是多门店还是单门店
    is_multi_shop = len(shop_ids) > 1
    shop_label = "多门店" if is_multi_shop else shop_ids[0]
    date_str = f"{begin_date}_{end_date}"
    output_file = os.path.join(output_dir, f"美团经营宝汇总_{shop_label}_{date_str}.xlsx")

    summary_generator.save(output_file)

    print("\n" + "=" * 60)
    print("汇总报表生成完成")
    print("=" * 60)

    return output_file


if __name__ == "__main__":
    generate_summary_report()

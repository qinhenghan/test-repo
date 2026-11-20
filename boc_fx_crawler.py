# -*- coding: utf-8 -*-
"""
首个爬虫作品：中国银行外汇牌价采集
====================================
数据源：https://www.boc.cn/sourcedb/whpj/  （中国银行官网 · 金融市场 · 外汇牌价，公开数据，无需登录）

功能：
  1. 抓取页面，解析当前全部货币的外汇牌价
  2. 字段清洗：空值处理、价格转 float、按货币名去重、排序
  3. 输出 Excel（带表头样式），并在终端打印摘要

运行：py boc_fx_crawler.py
依赖：requests, beautifulsoup4, openpyxl   （pip install requests beautifulsoup4 openpyxl）
"""

import re
import os
import datetime
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 数据源 URL（可加 ?pjname=... 查询历史，此处取实时牌价）
URL = "https://www.boc.cn/sourcedb/whpj/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
}

# 字段顺序（与页面表格列一一对应，来源：中国银行外汇牌价表）
FIELDS = [
    "货币名称",
    "现汇买入价",
    "现钞买入价",
    "现汇卖出价",
    "现钞卖出价",
    "中行折算价",
    "发布时间",   # 形如 2026/11/20 07:54:43
]


def clean_price(text: str):
    """清洗单个价格：去空白，空串返回 None，否则转 float（保留2位）。"""
    t = (text or "").strip().replace(",", "")
    if not t:
        return None
    try:
        return round(float(t), 2)
    except ValueError:
        return None


def parse_page(html: str):
    """解析 HTML，返回清洗后的记录列表（list[dict]）。"""
    soup = BeautifulSoup(html, "html.parser")
    rows = soup.find_all("tr")
    records = []

    for tr in rows:
        cells = [td.get_text(strip=True) for td in tr.find_all("td")]
        # 数据行特征：至少 8 个单元格，且第 7 个是"日期"（含 /）
        if len(cells) < 7 or "/" not in cells[6]:
            continue

        # 发布时间拆成"日期 + 时间"
        publish = cells[6]
        date_part = publish.split(" ")[0] if " " in publish else publish
        time_part = publish.split(" ")[1] if " " in publish else ""

        record = {
            "货币名称": cells[0],
            "现汇买入价": clean_price(cells[1]),
            "现钞买入价": clean_price(cells[2]),
            "现汇卖出价": clean_price(cells[3]),
            "现钞卖出价": clean_price(cells[4]),
            "中行折算价": clean_price(cells[5]),
            "发布时间": publish,
            "发布日期": date_part,   # 清洗：拆出日期
            "更新时间": time_part,   # 清洗：拆出时间
        }
        records.append(record)

    # 清洗：按货币名称去重（保留第一条）
    seen = set()
    deduped = []
    for rec in records:
        if rec["货币名称"] not in seen:
            seen.add(rec["货币名称"])
            deduped.append(rec)

    # 清洗：按货币名称排序
    deduped.sort(key=lambda r: r["货币名称"])
    return deduped


def save_excel(records, out_path):
    """输出 Excel：带样式的表头 + 数据行。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "外汇牌价"

    # 表头（含清洗后新增的两列）
    headers = FIELDS + ["发布日期", "更新时间"]
    ws.append(headers)

    # 表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="C0392B")  # 中国红表头
    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # 数据
    for rec in records:
        ws.append([rec[h] for h in headers])

    # 列宽自适应
    widths = [14, 12, 12, 12, 12, 12, 22, 12, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 冻结表头
    ws.freeze_panes = "A2"
    wb.save(out_path)
    return out_path


def main():
    print(f"[1/3] 抓取数据源：{URL}")
    resp = requests.get(URL, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    # 该页面为 UTF-8，显式解码
    html = resp.content.decode("utf-8", errors="ignore")
    print(f"      状态码 {resp.status_code}，页面 {len(html)} 字节")

    print("[2/3] 解析 + 清洗...")
    records = parse_page(html)
    if not records:
        raise SystemExit("未解析到数据，请检查页面结构或稍后重试。")
    print(f"      解析到 {len(records)} 种货币")

    today = datetime.date.today().strftime("%Y%m%d")
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"中国银行外汇牌价_{today}.xlsx")

    print("[3/3] 写入 Excel...")
    save_excel(records, out_path)
    print(f"      已保存：{out_path}")

    # 摘要
    print("\n=== 摘要（前 5 条，单位：100 外币兑人民币）===")
    for rec in records[:5]:
        print(f"  {rec['货币名称']:　<8} 现汇买入 {rec['现汇买入价']} | "
              f"现汇卖出 {rec['现汇卖出价']} | 中行折算 {rec['中行折算价']} | {rec['更新时间']}")
    print(f"... 共 {len(records)} 条，发布时间 {records[0]['发布时间']}")


if __name__ == "__main__":
    main()

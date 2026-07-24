#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NPS分析与转化报告生成脚本 v3
- 辅导维度 + 组长维度双表输出
- Excel 带 ColorScale 色阶
- 一个文件多 sheet（辅导维度 / 组长维度 / 汇总对比）
- NPS口径: 推荐者>=10, 贬损者<=7 (底表1-11分制)
- 数据校验: nps列值超出1-11时告警
"""

import pandas as pd
import sys
import os
import argparse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════
# 列名适配
# ═══════════════════════════════════════════════════════════

def _find_col(df, candidates):
    """在DataFrame列中查找第一个匹配的列名"""
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(f"未找到所需列，候选: {candidates}")


def _normalize_columns(df):
    """统一列名，支持三种底表格式"""
    mapping = {}
    rec_col = _find_col(df, ['【推荐值】', '推荐值', 'nps'])
    if rec_col != '推荐值':
        mapping[rec_col] = '推荐值'
    reg_col = _find_col(df, ['是否考虑报名长期班', '是否报名'])
    if reg_col != '是否报名':
        mapping[reg_col] = '是否报名'
    grade_col = _find_col(df, ['年级', 'grade_names'])
    if grade_col != '年级':
        mapping[grade_col] = '年级'
    tutor_col = _find_col(df, ['辅导', 'counselor_name', '辅导姓名'])
    if tutor_col != '辅导':
        mapping[tutor_col] = '辅导'
    leader_col = _find_col(df, ['组长', '组长姓名'])
    if leader_col != '组长':
        mapping[leader_col] = '组长'
    if mapping:
        df = df.rename(columns=mapping)
    return df


def _validate_nps_column(df, col='推荐值'):
    """
    校验 NPS 推荐值列是否在预期范围 1-11 内。
    如果存在超出范围的值，打印告警信息。
    返回: (是否正常, 异常数量, 异常值列表)
    """
    if col not in df.columns:
        return True, 0, []
    
    vals = df[col].dropna()
    invalid_mask = (vals < 1) | (vals > 11)
    invalid_count = invalid_mask.sum()
    
    if invalid_count > 0:
        invalid_vals = vals[invalid_mask].value_counts().to_dict()
        print(f'\n{"="*60}')
        print(f'⚠️  ⚠️  ⚠️  数据质量告警  ⚠️  ⚠️  ⚠️')
        print(f'{"="*60}')
        print(f'  NPS推荐值列 "{col}" 存在 {invalid_count} 条超出 1-11 范围的异常值')
        print(f'  异常值分布: {invalid_vals}')
        print(f'  总样本量: {len(df)}, 异常占比: {invalid_count/len(df)*100:.1f}%')
        print(f'  ⚠️  请检查底表数据是否正确！')
        print(f'  异常行将被自动剔除，NPS计算仅使用1-11分数据。')
        print(f'{"="*60}\n')
        return False, invalid_count, invalid_vals
    
    return True, 0, []


# ═══════════════════════════════════════════════════════════
# NPS 计算
# ═══════════════════════════════════════════════════════════

def calc_nps(group):
    """对一组用户计算NPS和转化指标
    
    NPS口径（底表1-11分制）:
      - 推荐者: >=10 (对应标准0-10量表的9-10)
      - 中立者: 8-9 (对应标准0-10量表的7-8)
      - 贬损者: <=7 (对应标准0-10量表的0-6)
      - NPS = 推荐者占比 - 贬损者占比
    """
    total = len(group)
    promoters = len(group[group['推荐值'] >= 10])
    detractors = len(group[group['推荐值'] <= 7])
    nps = (promoters - detractors) / total * 100
    return pd.Series({
        '样本量': total,
        'NPS': round(nps, 1),
        '推荐率%': round(promoters / total * 100, 1),
        '贬损率%': round(detractors / total * 100, 1),
        '报名率%': round(len(group[group['是否报名'] == 1]) / total * 100, 1),
        '考虑率%': round(len(group[group['是否报名'] == 2]) / total * 100, 1),
        '拒绝率%': round(len(group[group['是否报名'] == 3]) / total * 100, 1),
    })


# ═══════════════════════════════════════════════════════════
# Excel 生成（带色阶 + 多 sheet）
# ═══════════════════════════════════════════════════════════

HDR_FONT   = Font(name='Microsoft YaHei', bold=True, size=11, color='FFFFFF')
HDR_FILL   = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HDR_ALIGN  = Alignment(horizontal='center', vertical='center', wrap_text=True)
CELL_FONT  = Font(name='Microsoft YaHei', size=11)
CELL_ALIGN = Alignment(horizontal='center', vertical='center')
CELL_LEFT  = Alignment(horizontal='left', vertical='center')
BORDER     = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9'),
)

# 色阶定义: (方向, 高值色, 中值色, 低值色)
COLOR_SCALES = {
    'NPS':      ('正向', '00B050', 'FFFFFF', 'FF6666'),
    '推荐率%':  ('正向', '2F5496', 'FFFFFF', 'D6E4F0'),
    '贬损率%':  ('反向', 'FF4444', 'FFFFFF', '00B050'),
    '报名率%':  ('正向', '00B050', 'FFFFFF', 'FFC000'),
    '拒绝率%':  ('反向', 'FF4444', 'FFFFFF', '548235'),
    '考虑率%':  ('中性', None, None, None),
}

TUTOR_HEADERS  = ['辅导老师', '组长', '样本量', 'NPS', '推荐率%', '贬损率%', '报名率%', '考虑率%', '拒绝率%']
TUTOR_WIDTHS   = [14, 12, 10, 10, 10, 10, 10, 10, 10]
LEADER_HEADERS = ['组长', '样本量', 'NPS', '推荐率%', '贬损率%', '报名率%', '考虑率%', '拒绝率%']
LEADER_WIDTHS  = [14, 10, 10, 10, 10, 10, 10, 10]


def _write_sheet(ws, headers, data_rows, col_widths):
    """往工作表写入表头和数据，应用统一样式"""
    # 表头
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = HDR_ALIGN
        cell.border = BORDER

    # 数据行
    for r, row_data in enumerate(data_rows, 2):
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = CELL_FONT
            cell.alignment = CELL_LEFT if c == 1 else CELL_ALIGN
            cell.border = BORDER

    # 列宽
    for c, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.freeze_panes = 'A2'


def _add_color_scales(ws, data_start_row, data_end_row, headers):
    """给数据区域加 ColorScale 色阶"""
    n = data_end_row - data_start_row + 1
    if n <= 1:
        return

    for col_name in headers:
        cfg = COLOR_SCALES.get(col_name)
        if cfg is None or cfg[0] == '中性':
            continue
        direction, hi_color, mid_color, lo_color = cfg

        try:
            ci = headers.index(col_name) + 1  # 1-based
        except ValueError:
            continue

        col_letter = get_column_letter(ci)
        rng = f"{col_letter}{data_start_row}:{col_letter}{data_end_row}"

        if direction == '正向':
            ws.conditional_formatting.add(rng,
                ColorScaleRule(start_type='min', start_color=lo_color,
                               mid_type='percentile', mid_value=50, mid_color=mid_color,
                               end_type='max', end_color=hi_color))
        else:  # 反向
            ws.conditional_formatting.add(rng,
                ColorScaleRule(start_type='min', start_color=hi_color,
                               mid_type='percentile', mid_value=50, mid_color=mid_color,
                               end_type='max', end_color=lo_color))


def _build_tutor_rows(df_result):
    """从 tutor DataFrame 生成写入行"""
    rows = []
    for _, row in df_result.iterrows():
        rows.append([
            row['辅导'], row['组长'], int(row['样本量']),
            row['NPS'], row['推荐率%'], row['贬损率%'],
            row['报名率%'], row['考虑率%'], row['拒绝率%'],
        ])
    return rows


def _build_leader_rows(df_leader):
    """从 leader DataFrame 生成写入行"""
    rows = []
    for _, row in df_leader.iterrows():
        rows.append([
            row['组长'], int(row['样本量']),
            row['NPS'], row['推荐率%'], row['贬损率%'],
            row['报名率%'], row['考虑率%'], row['拒绝率%'],
        ])
    return rows


def _generate_excel_with_colors(result_tutor, result_leader, grade_label, output_dir):
    """生成带色阶的多 sheet Excel 文件"""
    wb = Workbook()

    tutor_rows = _build_tutor_rows(result_tutor)
    leader_rows = _build_leader_rows(result_leader)

    # ── Sheet 1: 辅导维度 ────────────────────────────
    ws1 = wb.active
    ws1.title = "辅导维度"
    _write_sheet(ws1, TUTOR_HEADERS, tutor_rows, TUTOR_WIDTHS)
    _add_color_scales(ws1, 2, len(tutor_rows) + 1, TUTOR_HEADERS)

    # ── Sheet 2: 组长维度 ────────────────────────────
    ws2 = wb.create_sheet("组长维度")
    _write_sheet(ws2, LEADER_HEADERS, leader_rows, LEADER_WIDTHS)
    _add_color_scales(ws2, 2, len(leader_rows) + 1, LEADER_HEADERS)

    # ── Sheet 3: 汇总对比 ────────────────────────────
    ws3 = wb.create_sheet("汇总对比")

    # 标题
    title1_row = 1
    ws3.merge_cells(start_row=title1_row, start_column=1,
                    end_row=title1_row, end_column=len(TUTOR_HEADERS))
    ws3.cell(row=title1_row, column=1,
             value=f'{grade_label} — 辅导维度 NPS（按NPS降序）').font = \
        Font(name='Microsoft YaHei', bold=True, size=13, color='2F5496')

    # 辅导表头
    tutor_hdr_row = title1_row + 1
    for c, h in enumerate(TUTOR_HEADERS, 1):
        cell = ws3.cell(row=tutor_hdr_row, column=c, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL
        cell.alignment = HDR_ALIGN; cell.border = BORDER

    for r, row_data in enumerate(tutor_rows):
        for c, val in enumerate(row_data, 1):
            cell = ws3.cell(row=r + tutor_hdr_row + 1, column=c, value=val)
            cell.font = CELL_FONT
            cell.alignment = CELL_LEFT if c == 1 else CELL_ALIGN
            cell.border = BORDER

    tutor_end_row = tutor_hdr_row + len(tutor_rows)
    _add_color_scales(ws3, tutor_hdr_row + 1, tutor_end_row, TUTOR_HEADERS)

    # 组长区（隔两行）
    leader_title_row = tutor_end_row + 3
    ws3.merge_cells(start_row=leader_title_row, start_column=1,
                    end_row=leader_title_row, end_column=len(LEADER_HEADERS))
    ws3.cell(row=leader_title_row, column=1,
             value=f'{grade_label} — 组长维度汇总').font = \
        Font(name='Microsoft YaHei', bold=True, size=13, color='2F5496')

    leader_hdr_row = leader_title_row + 1
    for c, h in enumerate(LEADER_HEADERS, 1):
        cell = ws3.cell(row=leader_hdr_row, column=c, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL
        cell.alignment = HDR_ALIGN; cell.border = BORDER

    for r, row_data in enumerate(leader_rows):
        for c, val in enumerate(row_data, 1):
            cell = ws3.cell(row=r + leader_hdr_row + 1, column=c, value=val)
            cell.font = CELL_FONT
            cell.alignment = CELL_LEFT if c == 1 else CELL_ALIGN
            cell.border = BORDER

    leader_end_row = leader_hdr_row + len(leader_rows)
    _add_color_scales(ws3, leader_hdr_row + 1, leader_end_row, LEADER_HEADERS)

    # 列宽
    for c, w in enumerate(TUTOR_WIDTHS, 1):
        ws3.column_dimensions[get_column_letter(c)].width = w

    # 保存
    excel_path = os.path.join(output_dir, f'{grade_label}NPS分析表.xlsx')
    wb.save(excel_path)
    print(f'Excel: {excel_path} ({len(tutor_rows)} 位辅导, {len(leader_rows)} 位组长)')

    return excel_path


# ═══════════════════════════════════════════════════════════
# HTML 生成
# ═══════════════════════════════════════════════════════════

def _html_section_table(rows_html, headers, section_title, note_text=''):
    """生成一个 HTML 区块（标题 + 表格）"""
    th_html = ''.join(f'<th>{h}</th>' for h in headers)
    return f'''
    <div class="section">
        <h2>{section_title}</h2>
        {note_text}
        <table><thead><tr>{th_html}</tr></thead><tbody>
        {''.join(rows_html)}
        </tbody></table>
    </div>'''


def _build_tutor_html_rows(result_tutor):
    """生成辅导维度的 HTML 行"""
    rows = []
    for _, row in result_tutor.iterrows():
        nps = row['NPS']
        rec = row['推荐率%']
        det = row['贬损率%']
        sig = row['报名率%']
        nos = row['拒绝率%']

        # NPS
        if nps >= 60: nbg, nc = '#2e7d32', 'white'
        elif nps >= 50: nbg, nc = '#43a047', 'white'
        elif nps >= 40: nbg, nc = '#66bb6a', 'white'
        elif nps >= 30: nbg, nc = '#a5d6a7', '#1b5e20'
        else: nbg, nc = '#c8e6c9', '#1b5e20'

        # 推荐
        if rec >= 75: rbg, rc = '#1976d2', 'white'
        elif rec >= 70: rbg, rc = '#1e88e5', 'white'
        elif rec >= 65: rbg, rc = '#42a5f5', 'white'
        elif rec >= 60: rbg, rc = '#64b5f6', '#0d47a1'
        else: rbg, rc = '#90caf9', '#0d47a1'

        # 贬损
        if det >= 22: dbg, dc = '#1976d2', 'white'
        elif det >= 18: dbg, dc = '#1e88e5', 'white'
        elif det >= 15: dbg, dc = '#42a5f5', 'white'
        elif det >= 12: dbg, dc = '#64b5f6', '#0d47a1'
        else: dbg, dc = '#90caf9', '#0d47a1'

        # 报名
        if sig >= 18: sbg, sc = '#2e7d32', 'white'
        elif sig >= 15: sbg, sc = '#43a047', 'white'
        elif sig >= 12: sbg, sc = '#66bb6a', '#1b5e20'
        elif sig >= 10: sbg, sc = '#a5d6a7', '#1b5e20'
        elif sig >= 8: sbg, sc = '#ffcc80', '#e65100'
        elif sig >= 5: sbg, sc = '#ffab91', '#bf360c'
        else: sbg, sc = '#ef9a9a', '#b71c1c'

        # 拒绝
        if nos >= 15: nbg2, nc2 = '#1976d2', 'white'
        elif nos >= 12: nbg2, nc2 = '#1e88e5', 'white'
        elif nos >= 8: nbg2, nc2 = '#42a5f5', 'white'
        elif nos >= 5: nbg2, nc2 = '#64b5f6', '#0d47a1'
        else: nbg2, nc2 = '#90caf9', '#0d47a1'

        rows.append(f'''<tr>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:left;font-weight:500;">{row['辅导']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:left;">{row['组长']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{nbg};color:{nc};font-weight:bold;">{nps:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{rbg};color:{rc};">{rec:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{dbg};color:{dc};">{det:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{sbg};color:{sc};font-weight:bold;">{sig:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;">{row['考虑率%']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{nbg2};color:{nc2};">{nos:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;color:#666;">{row['样本量']:.0f}</td>
        </tr>''')
    return rows


def _build_leader_html_rows(result_leader):
    """生成长维度的 HTML 行"""
    rows = []
    for _, row in result_leader.iterrows():
        nps = row['NPS']
        rec = row['推荐率%']
        sig = row['报名率%']
        nos = row['拒绝率%']

        if nps >= 60: nbg, nc = '#2e7d32', 'white'
        elif nps >= 50: nbg, nc = '#43a047', 'white'
        elif nps >= 40: nbg, nc = '#66bb6a', 'white'
        elif nps >= 30: nbg, nc = '#a5d6a7', '#1b5e20'
        else: nbg, nc = '#c8e6c9', '#1b5e20'

        if rec >= 75: rbg, rc = '#1976d2', 'white'
        elif rec >= 70: rbg, rc = '#1e88e5', 'white'
        elif rec >= 65: rbg, rc = '#42a5f5', 'white'
        elif rec >= 60: rbg, rc = '#64b5f6', '#0d47a1'
        else: rbg, rc = '#90caf9', '#0d47a1'

        if sig >= 18: sbg, sc = '#2e7d32', 'white'
        elif sig >= 15: sbg, sc = '#43a047', 'white'
        elif sig >= 12: sbg, sc = '#66bb6a', '#1b5e20'
        elif sig >= 10: sbg, sc = '#a5d6a7', '#1b5e20'
        elif sig >= 8: sbg, sc = '#ffcc80', '#e65100'
        elif sig >= 5: sbg, sc = '#ffab91', '#bf360c'
        else: sbg, sc = '#ef9a9a', '#b71c1c'

        if nos >= 15: nbg2, nc2 = '#1976d2', 'white'
        elif nos >= 12: nbg2, nc2 = '#1e88e5', 'white'
        elif nos >= 8: nbg2, nc2 = '#42a5f5', 'white'
        elif nos >= 5: nbg2, nc2 = '#64b5f6', '#0d47a1'
        else: nbg2, nc2 = '#90caf9', '#0d47a1'

        # 贬损
        det = row['贬损率%']
        if det >= 22: dbg, dc = '#1976d2', 'white'
        elif det >= 18: dbg, dc = '#1e88e5', 'white'
        elif det >= 15: dbg, dc = '#42a5f5', 'white'
        elif det >= 12: dbg, dc = '#64b5f6', '#0d47a1'
        else: dbg, dc = '#90caf9', '#0d47a1'

        rows.append(f'''<tr>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:left;font-weight:500;">{row['组长']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;color:#666;">{row['样本量']:.0f}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{nbg};color:{nc};font-weight:bold;">{nps:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{rbg};color:{rc};">{rec:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{dbg};color:{dc};">{det:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{sbg};color:{sc};font-weight:bold;">{sig:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;">{row['考虑率%']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{nbg2};color:{nc2};">{nos:.1f}%</td>
        </tr>''')
    return rows


def generate_html(result_tutor, result_leader, total_samples, grade_name, date_label):
    """生成带颜色标注的双维度 HTML 报告"""
    tutor_rows = _build_tutor_html_rows(result_tutor)
    leader_rows = _build_leader_html_rows(result_leader)

    tutor_section = _html_section_table(
        tutor_rows,
        TUTOR_HEADERS,
        f'{grade_name} — 辅导维度 NPS（按 NPS 降序）',
        f'<p class="note">样本量：{total_samples} 份 | {len(result_tutor)} 位辅导老师</p>'
    )
    leader_section = _html_section_table(
        leader_rows,
        LEADER_HEADERS,
        f'{grade_name} — 组长维度汇总',
    )

    return f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
    body {{ font-family: "Microsoft YaHei", "PingFang SC", Arial, sans-serif; background: #f8f9fa; padding: 30px; margin: 0; }}
    .container {{ max-width: 1050px; margin: 0 auto; background: white; padding: 32px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }}
    h1 {{ font-size: 22px; color: #1a1a1a; margin-bottom: 8px; text-align: center; font-weight: 600; }}
    .subtitle {{ font-size: 13px; color: #888; margin-bottom: 30px; text-align: center; }}
    .section {{ margin-bottom: 36px; }}
    .section h2 {{ font-size: 16px; color: #2F5496; margin-bottom: 6px; border-left: 4px solid #4472C4; padding-left: 10px; }}
    .note {{ font-size: 12px; color: #999; margin-bottom: 8px; }}
    table {{ width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; }}
    th {{ padding: 12px 14px; background: #f0f4f8; color: #333; font-weight: 600; text-align: center; border-bottom: 2px solid #d0d7de; border-top: 1px solid #e8eaed; font-size: 13px; }}
    th:first-child {{ border-top-left-radius: 8px; text-align: left; }}
    th:last-child {{ border-top-right-radius: 8px; }}
    tr:last-child td:first-child {{ border-bottom-left-radius: 8px; }}
    tr:last-child td:last-child {{ border-bottom-right-radius: 8px; }}
    tr:hover td {{ filter: brightness(0.97); }}
    td {{ transition: all 0.15s ease; }}
</style></head><body>
<div class="container">
<h1>{grade_name} NPS 与转化分析报告</h1>
<div class="subtitle">数据时间：{date_label}</div>
{tutor_section}
{leader_section}
</div></body></html>'''


# ═══════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════

def run(input_file, grade=None, min_samples=10, output_dir=None, sheet=0):
    """
    主函数：读取Excel，按年级和分组计算NPS与转化指标，输出Excel+HTML。

    参数:
        input_file:  Excel文件路径
        grade:       筛选年级，None表示全量
        min_samples: 最小样本量阈值，低于此值的辅导老师将被过滤
        output_dir:  输出目录，None则输出到input_file所在目录
        sheet:       Sheet名或索引（0=第一个sheet），默认0
    """
    if output_dir is None:
        output_dir = os.path.dirname(input_file) or '.'

    df = pd.read_excel(input_file, sheet_name=sheet)
    df = _normalize_columns(df)

    # ── 数据校验：nps值是否在 1-11 ──
    is_valid, inv_count, inv_vals = _validate_nps_column(df, '推荐值')
    if not is_valid:
        # 剔除异常行，只保留 1-11 分的数据
        valid_mask = (df['推荐值'] >= 1) & (df['推荐值'] <= 11) | df['推荐值'].isna()
        df = df[valid_mask].copy()

    # 筛选年级
    if grade and grade in df['年级'].unique():
        df_grade = df[df['年级'] == grade].copy()
        grade_label = grade
    else:
        df_grade = df.copy()
        grade_label = '全年级'

    total = len(df_grade)

    # ── 辅导维度 ──
    result_tutor = (
        df_grade.groupby(['辅导', '组长'], dropna=False)
        .apply(calc_nps, include_groups=False)
        .reset_index()
    )
    result_tutor = result_tutor[result_tutor['样本量'] >= min_samples].copy()
    result_tutor['组长'] = result_tutor['组长'].fillna('未分组')
    result_tutor = result_tutor.sort_values(
        ['组长', 'NPS'], ascending=[True, False]
    ).reset_index(drop=True)

    # ── 组长维度 ──
    result_leader = (
        df_grade.groupby('组长')
        .apply(calc_nps, include_groups=False)
        .reset_index()
    )
    result_leader['组长'] = result_leader['组长'].fillna('未分组')
    result_leader = result_leader.sort_values('NPS', ascending=False).reset_index(drop=True)

    # ── 输出 Excel ──
    excel_path = _generate_excel_with_colors(result_tutor, result_leader, grade_label, output_dir)

    # ── 输出 HTML ──
    date_label = pd.Timestamp.now().strftime('%Y年%m月')
    html = generate_html(result_tutor, result_leader, total, grade_label, date_label)
    html_path = os.path.join(output_dir, f'{grade_label}NPS分析表.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'HTML: {html_path}')

    return result_tutor, result_leader


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='NPS分析与转化报告生成 v2')
    parser.add_argument('input_file', help='Excel文件路径')
    parser.add_argument('--grade', '-g', default=None, help='筛选年级（如：一年级）')
    parser.add_argument('--min-samples', '-m', type=int, default=10, help='最小样本量阈值')
    parser.add_argument('--output-dir', '-o', default=None, help='输出目录')
    parser.add_argument('--sheet', '-s', default=0, help='Sheet名或索引（默认第一个sheet）')
    args = parser.parse_args()

    run(args.input_file, grade=args.grade, min_samples=args.min_samples,
        output_dir=args.output_dir, sheet=args.sheet)

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NPS分析与转化报告生成脚本
按辅导老师×组长维度计算NPS指标与转化意向，输出Excel和HTML报告。
"""

import pandas as pd
import sys
import os
import argparse
import math


def calc_nps(group):
    """对一组用户计算NPS和转化指标"""
    total = len(group)
    promoters = len(group[group['【推荐值】'] >= 9])
    detractors = len(group[group['【推荐值】'] <= 6])
    nps = (promoters - detractors) / total * 100
    return pd.Series({
        '样本量': total,
        'NPS': round(nps, 1),
        '推荐': round(promoters / total * 100, 1),
        '贬损': round(detractors / total * 100, 1),
        '要报': round(len(group[group['是否考虑报名长期班'] == 1]) / total * 100, 1),
        '考虑': round(len(group[group['是否考虑报名长期班'] == 2]) / total * 100, 1),
        '不报': round(len(group[group['是否考虑报名长期班'] == 3]) / total * 100, 1),
    })


def color_for_value(value, thresholds, colors):
    """根据阈值返回对应的背景色和文字色"""
    for (lo, hi), (bg, fg) in zip(thresholds, colors):
        if value >= lo and (hi is None or value < hi):
            return bg, fg
    return '#f5f5f5', '#333'


def generate_html(result, total_samples, grade_name, date_label):
    """生成带颜色标注的HTML报告"""
    def nps_colors(v):
        if v >= 60: return '#2e7d32', 'white'
        if v >= 50: return '#43a047', 'white'
        if v >= 40: return '#66bb6a', 'white'
        if v >= 30: return '#a5d6a7', '#1b5e20'
        return '#c8e6c9', '#1b5e20'

    def rec_colors(v):
        if v >= 75: return '#1976d2', 'white'
        if v >= 70: return '#1e88e5', 'white'
        if v >= 65: return '#42a5f5', 'white'
        if v >= 60: return '#64b5f6', '#0d47a1'
        return '#90caf9', '#0d47a1'

    def det_colors(v):
        if v >= 22: return '#1976d2', 'white'
        if v >= 18: return '#1e88e5', 'white'
        if v >= 15: return '#42a5f5', 'white'
        if v >= 12: return '#64b5f6', '#0d47a1'
        return '#90caf9', '#0d47a1'

    def sign_colors(v):
        if v >= 18: return '#2e7d32', 'white'
        if v >= 15: return '#43a047', 'white'
        if v >= 12: return '#66bb6a', '#1b5e20'
        if v >= 10: return '#a5d6a7', '#1b5e20'
        if v >= 8: return '#ffcc80', '#e65100'
        if v >= 5: return '#ffab91', '#bf360c'
        return '#ef9a9a', '#b71c1c'

    def nos_colors(v):
        if v >= 15: return '#1976d2', 'white'
        if v >= 12: return '#1e88e5', 'white'
        if v >= 8: return '#42a5f5', 'white'
        if v >= 5: return '#64b5f6', '#0d47a1'
        return '#90caf9', '#0d47a1'

    rows_html = []
    for _, row in result.iterrows():
        nbg, nc = nps_colors(row['NPS'])
        rbg, rc = rec_colors(row['推荐'])
        dbg, dc = det_colors(row['贬损'])
        sbg, sc = sign_colors(row['要报'])
        nbg2, nc2 = nos_colors(row['不报'])

        rows_html.append(f'''<tr>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:left;font-weight:500;">{row['辅导']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:left;">{row['组长']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{nbg};color:{nc};font-weight:bold;">{row['NPS']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{rbg};color:{rc};">{row['推荐']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{dbg};color:{dc};">{row['贬损']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{sbg};color:{sc};font-weight:bold;">{row['要报']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;">{row['考虑']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;background:{nbg2};color:{nc2};">{row['不报']:.1f}%</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e8eaed;text-align:center;color:#666;">{row['样本量']:.0f}</td>
        </tr>''')

    return f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
    body {{ font-family: "Microsoft YaHei", "PingFang SC", Arial, sans-serif; background: #f8f9fa; padding: 30px; margin: 0; }}
    .container {{ max-width: 1000px; margin: 0 auto; background: white; padding: 32px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }}
    h1 {{ font-size: 22px; color: #1a1a1a; margin-bottom: 8px; text-align: center; font-weight: 600; }}
    .subtitle {{ font-size: 13px; color: #888; margin-bottom: 24px; text-align: center; }}
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
<h1>{grade_name}辅导老师 NPS 与转化分析表</h1>
<div class="subtitle">样本量：{total_samples} 份问卷 | 数据时间：{date_label} | 按组长分组，组内按 NPS 降序排列</div>
<table><thead><tr>
    <th>辅导</th><th>组长</th><th>NPS</th><th>推荐</th><th>贬损</th><th>要报</th><th>考虑</th><th>不报</th><th>样本量</th>
</tr></thead><tbody>
{''.join(rows_html)}
</tbody></table></div></body></html>'''


def run(input_file, grade=None, min_samples=10, output_dir=None):
    """
    主函数：读取Excel，按年级和分组计算NPS与转化指标，输出Excel+HTML。

    参数:
        input_file:  Excel文件路径
        grade:       筛选年级，None表示全量
        min_samples: 最小样本量阈值，低于此值的辅导老师将被过滤
        output_dir:  输出目录，None则输出到input_file所在目录
    """
    if output_dir is None:
        output_dir = os.path.dirname(input_file) or '.'

    df = pd.read_excel(input_file)

    # 筛选年级
    if grade and grade in df['年级'].unique():
        df = df[df['年级'] == grade].copy()
        grade_label = grade
    else:
        grade_label = '全年级'

    total = len(df)

    # 分组计算
    result = df.groupby(['辅导', '组长'], dropna=False).apply(calc_nps, include_groups=False).reset_index()

    # 过滤小样本
    result = result[result['样本量'] >= min_samples].copy()
    result['组长'] = result['组长'].fillna('未分组')
    result = result.sort_values(['组长', 'NPS'], ascending=[True, False]).reset_index(drop=True)

    # 输出Excel
    excel_path = os.path.join(output_dir, f'{grade_label}NPS分析表.xlsx')
    result.to_excel(excel_path, index=False, sheet_name=f'{grade_label}NPS')
    print(f'Excel: {excel_path} ({len(result)} 位辅导老师)')

    # 输出HTML
    date_label = pd.Timestamp.now().strftime('%Y年%m月')
    html = generate_html(result, total, grade_label, date_label)
    html_path = os.path.join(output_dir, f'{grade_label}NPS分析表.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'HTML: {html_path}')

    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='NPS分析与转化报告生成')
    parser.add_argument('input_file', help='Excel文件路径')
    parser.add_argument('--grade', '-g', default=None, help='筛选年级（如：一年级）')
    parser.add_argument('--min-samples', '-m', type=int, default=10, help='最小样本量阈值')
    parser.add_argument('--output-dir', '-o', default=None, help='输出目录')
    args = parser.parse_args()

    run(args.input_file, grade=args.grade, min_samples=args.min_samples, output_dir=args.output_dir)

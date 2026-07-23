#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NPS 两表合并 + 分析一体化脚本 v1
================================================
Phase 1: 读取 NPS转化底表 + 行课底表 → 列名重命名 → 教学组三重键匹配
Phase 2: 年级筛选 → 输出合并后底表
Phase 3: 调用 nps_calculator 生成 NPS 汇总表（辅导+组长维度+色阶）+ HTML报告

依赖: nps-analysis skill (nps_calculator.py)
"""

import pandas as pd
import sys
import os
import argparse
import importlib.util

sys.stdout.reconfigure(encoding='utf-8')


# ═══════════════════════════════════════════════════════════
# 列名适配工具
# ═══════════════════════════════════════════════════════════

def _find_col(df, candidates, default=None):
    """在 DataFrame 列中查找第一个匹配的列名"""
    for c in candidates:
        if c in df.columns:
            return c
    if default is not None:
        return default
    raise KeyError(f"未找到所需列，候选: {candidates}")


# NPS转化底表列名重命名映射
NPS_RENAME_MAP = {
    'q1': 'nps',
    'q2': 'q2喜爱度',
    'q3': 'q3收获感',
    'q4': 'q4专注力',
    'q5': 'q5表达意愿',
    'q10': '是否考虑报名长期班',
}


# ═══════════════════════════════════════════════════════════
# 加载 nps_calculator 模块
# ═══════════════════════════════════════════════════════════

def _load_nps_calculator():
    """从多个位置查找并加载 nps_calculator 模块"""
    here = os.path.dirname(os.path.abspath(__file__))

    candidates = [
        # 同目录（bundled copy）
        os.path.join(here, 'nps_calculator.py'),
        # 本地 skill 安装: ~/.workbuddy/skills/nps-analysis/scripts/
        os.path.join(os.path.dirname(os.path.dirname(here)),
                     'nps-analysis', 'scripts', 'nps_calculator.py'),
        # GitHub 仓库结构: repo/scripts/nps_calculator.py
        # (merge_and_analyze.py 在 repo/nps-merge-analysis/scripts/ 下)
        os.path.join(os.path.dirname(os.path.dirname(here)),
                     'scripts', 'nps_calculator.py'),
    ]

    for path in candidates:
        if os.path.exists(path):
            spec = importlib.util.spec_from_file_location("nps_calculator", path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod

    raise FileNotFoundError(
        "未找到 nps_calculator.py\n"
        "请确保 nps-analysis skill 已安装，路径:\n"
        "  ~/.workbuddy/skills/nps-analysis/scripts/nps_calculator.py\n"
        "或将 nps_calculator.py 放在 merge_and_analyze.py 同目录"
    )


# ═══════════════════════════════════════════════════════════
# Phase 1: 两表合并
# ═══════════════════════════════════════════════════════════

def _auto_detect_sheet(file_path, required_cols, default_sheet=0):
    """
    自动检测包含所需列的 sheet。
    如果 default_sheet 已包含所需列，直接使用；
    否则遍历所有 sheet 找到包含最多所需列的那个。
    """
    xls = pd.ExcelFile(file_path)
    # 先试默认 sheet
    df_try = pd.read_excel(file_path, sheet_name=default_sheet, nrows=2)
    found = sum(1 for c in required_cols if c in df_try.columns)
    if found >= len(required_cols) * 0.5:
        return default_sheet

    # 遍历所有 sheet 找最佳匹配
    best_sheet = default_sheet
    best_count = 0
    for s in xls.sheet_names:
        df_s = pd.read_excel(file_path, sheet_name=s, nrows=2)
        cnt = sum(1 for c in required_cols if c in df_s.columns)
        if cnt > best_count:
            best_count = cnt
            best_sheet = s
    return best_sheet


def merge_tables(nps_file, xingke_file, nps_sheet=0, xk_sheet=0):
    """
    合并 NPS转化底表 + 行课底表

    步骤:
      1. 读取两个底表（自动检测 sheet）
      2. 重命名 NPS转化底表列名 (q1→nps 等)
      3. 三重键匹配教学组 (辅导名+年级+课程ID)
      4. 降级匹配 (辅导名+年级)
      5. 填充组长姓名列

    返回: 合并后的完整 DataFrame（未筛选年级）
    """
    # ── 1. 读取两个底表（自动检测 sheet）──
    print("Step 1: 读取底表...")

    # NPS底表: 需要 counselor_name/grade_names/fk_course 或 辅导/年级
    nps_required = ['counselor_name', 'grade_names', 'fk_course', '辅导', '年级']
    actual_nps_sheet = _auto_detect_sheet(nps_file, nps_required, nps_sheet)
    if actual_nps_sheet != nps_sheet:
        print(f"  NPS底表: 自动切换到 sheet '{actual_nps_sheet}'")
    df_nps = pd.read_excel(nps_file, sheet_name=actual_nps_sheet)

    # 行课底表: 需要 辅导名/年级/课程ID/教学组
    xk_required = ['辅导名', '年级', '课程ID', '教学组', 'counselor_name']
    actual_xk_sheet = _auto_detect_sheet(xingke_file, xk_required, xk_sheet)
    if actual_xk_sheet != xk_sheet:
        print(f"  行课底表: 自动切换到 sheet '{actual_xk_sheet}'")
    df_xk = pd.read_excel(xingke_file, sheet_name=actual_xk_sheet)

    print(f"  NPS转化底表: {len(df_nps)} 行")
    print(f"  行课底表: {len(df_xk)} 行")

    # ── 2. 重命名NPS转化底表列名 ──
    print("\nStep 2: 重命名NPS转化底表列名...")
    rename_map = {}
    for old, new in NPS_RENAME_MAP.items():
        if old in df_nps.columns:
            rename_map[old] = new
    if rename_map:
        df_nps = df_nps.rename(columns=rename_map)
        print(f"  已重命名: {rename_map}")
    else:
        print("  无需重命名（列名已是目标格式或不存在q1等列）")

    # ── 3. 识别列名 ──
    # NPS底表
    nps_tutor_col = _find_col(df_nps, ['counselor_name', '辅导', '辅导姓名'], 'counselor_name')
    nps_grade_col = _find_col(df_nps, ['grade_names', '年级'], 'grade_names')
    nps_course_col = _find_col(df_nps, ['fk_course', '课程ID', 'course_id'], 'fk_course')

    # 行课底表
    xk_tutor_col = _find_col(df_xk, ['辅导名', '辅导', 'counselor_name', '辅导姓名'], '辅导名')
    xk_grade_col = _find_col(df_xk, ['年级', 'grade_names'], '年级')
    xk_course_col = _find_col(df_xk, ['课程ID', 'fk_course', 'course_id'], '课程ID')
    xk_group_col = _find_col(df_xk, ['教学组', '组长', '组长姓名'], '教学组')

    print(f"\n  列名识别:")
    print(f"    NPS底表: 辅导={nps_tutor_col}, 年级={nps_grade_col}, 课程={nps_course_col}")
    print(f"    行课底表: 辅导={xk_tutor_col}, 年级={xk_grade_col}, 课程={xk_course_col}, 教学组={xk_group_col}")

    # ── 4. 提取教学组映射 ──
    print("\nStep 3: 提取教学组映射...")

    # 三重键: (辅导名, 年级, 课程ID)
    xk_3key = df_xk.groupby([xk_tutor_col, xk_grade_col, xk_course_col])[xk_group_col].first().reset_index()
    print(f"  三重键映射: {len(xk_3key)} 组合")

    # 双键降级: (辅导名, 年级)
    xk_2key = df_xk.groupby([xk_tutor_col, xk_grade_col])[xk_group_col].first().reset_index()
    print(f"  双键映射: {len(xk_2key)} 组合")

    # ── 5. 匹配教学组 ──
    print("\nStep 4: 匹配教学组...")

    # 三重键匹配
    df_nps['_mk3'] = list(zip(
        df_nps[nps_tutor_col], df_nps[nps_grade_col], df_nps[nps_course_col]
    ))
    xk_3key['_mk3'] = list(zip(
        xk_3key[xk_tutor_col], xk_3key[xk_grade_col], xk_3key[xk_course_col]
    ))
    dict_3key = dict(zip(xk_3key['_mk3'], xk_3key[xk_group_col]))
    df_nps['教学组'] = df_nps['_mk3'].map(dict_3key)

    matched_3 = df_nps['教学组'].notna().sum()
    pct_3 = matched_3 / len(df_nps) * 100 if len(df_nps) > 0 else 0
    print(f"  三重键匹配: {matched_3}/{len(df_nps)} ({pct_3:.1f}%)")

    # 降级匹配
    unmatched = df_nps['教学组'].isna()
    if unmatched.any():
        dict_2key = dict(zip(
            zip(xk_2key[xk_tutor_col], xk_2key[xk_grade_col]),
            xk_2key[xk_group_col]
        ))
        fb_keys = list(zip(
            df_nps.loc[unmatched, nps_tutor_col],
            df_nps.loc[unmatched, nps_grade_col]
        ))
        fb_matched = pd.Series(fb_keys).map(dict_2key)
        df_nps.loc[unmatched, '教学组'] = fb_matched.values
        matched_2 = df_nps['教学组'].notna().sum() - matched_3
        print(f"  降级匹配: +{matched_2}")

    still_unmatched = df_nps['教学组'].isna().sum()
    pct_un = still_unmatched / len(df_nps) * 100 if len(df_nps) > 0 else 0
    print(f"  最终未匹配: {still_unmatched}/{len(df_nps)} ({pct_un:.1f}%)")

    if still_unmatched > 0:
        unmatched_tutors = df_nps[df_nps['教学组'].isna()][nps_tutor_col].unique()
        print(f"  未匹配辅导: {list(unmatched_tutors)[:15]}")

    # ── 6. 填充组长姓名列 ──
    print("\nStep 5: 填充组长姓名列...")
    if '组长姓名' not in df_nps.columns:
        df_nps['组长姓名'] = None
    df_nps['组长姓名'] = df_nps['教学组']
    print(f"  组长姓名已填充（来源: 教学组）")

    # 清理临时列
    df_nps = df_nps.drop(columns=['_mk3'], errors='ignore')

    # 关键统计
    print(f"\n===== 合并结果 =====")
    print(f"总样本: {len(df_nps)}")
    print(f"辅导老师: {df_nps[nps_tutor_col].nunique()} 位")
    print(f"教学组: {df_nps['教学组'].nunique()} 个")
    print(f"教学组缺失: {df_nps['教学组'].isna().sum()} 行")

    grade_col_name = nps_grade_col
    print(f"\n各年级分布:")
    for g, cnt in df_nps.groupby(grade_col_name).size().sort_values(ascending=False).items():
        tutors = df_nps[df_nps[grade_col_name] == g][nps_tutor_col].nunique()
        print(f"  {g}: {cnt} 份, {tutors} 位辅导")

    return df_nps


# ═══════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════

def run(nps_file, xingke_file, grade=None, min_samples=10, output_dir=None,
        nps_sheet=0, xk_sheet=0):
    """
    完整流程：两表合并 → 年级筛选 → NPS分析 → 输出报告

    参数:
        nps_file:    NPS转化底表路径（必填）
        xingke_file: 行课底表路径（必填）
        grade:       筛选年级（如"一年级"），None=全量
        min_samples: 最小样本量阈值，默认 10
        output_dir:  输出目录，None 则输出到 nps_file 所在目录
        nps_sheet:   NPS底表 sheet 名或索引，默认 0
        xk_sheet:    行课底表 sheet 名或索引，默认 0

    返回:
        (merged_path, result_tutor, result_leader)
    """
    if output_dir is None:
        output_dir = os.path.dirname(nps_file) or '.'

    # ═══════════════════════════════════════════════════════
    # Phase 1: 两表合并
    # ═══════════════════════════════════════════════════════
    print("=" * 60)
    print("Phase 1: 两表合并")
    print("=" * 60)
    df_merged = merge_tables(nps_file, xingke_file, nps_sheet, xk_sheet)

    # ═══════════════════════════════════════════════════════
    # Phase 2: 年级筛选 + 保存合并底表
    # ═══════════════════════════════════════════════════════
    print("\n" + "=" * 60)
    print("Phase 2: 年级筛选 + 保存合并底表")
    print("=" * 60)

    grade_col = _find_col(df_merged, ['grade_names', '年级'], 'grade_names')
    if grade and grade in df_merged[grade_col].unique():
        df_grade = df_merged[df_merged[grade_col] == grade].copy()
        grade_label = grade
    else:
        df_grade = df_merged.copy()
        grade_label = '全年级'
        if grade:
            print(f"  ⚠️ 未找到年级「{grade}」，使用全量数据")

    print(f"  {grade_label}: {len(df_grade)} 行, "
          f"{df_grade[_find_col(df_grade, ['counselor_name', '辅导', '辅导姓名'], 'counselor_name')].nunique()} 位辅导")

    merged_path = os.path.join(output_dir, f'{grade_label}合并后底表.xlsx')
    df_grade.to_excel(merged_path, index=False, sheet_name=f'{grade_label}NPS底表')
    print(f"  合并底表已保存: {merged_path}")

    # ═══════════════════════════════════════════════════════
    # Phase 3: NPS 分析
    # ═══════════════════════════════════════════════════════
    print("\n" + "=" * 60)
    print("Phase 3: NPS 分析")
    print("=" * 60)

    nps_calc = _load_nps_calculator()
    result_tutor, result_leader = nps_calc.run(
        input_file=merged_path,
        grade=grade,
        min_samples=min_samples,
        output_dir=output_dir,
        sheet=0
    )

    # ═══════════════════════════════════════════════════════
    # 完成
    # ═══════════════════════════════════════════════════════
    nps_excel = os.path.join(output_dir, f'{grade_label}NPS分析表.xlsx')
    nps_html = os.path.join(output_dir, f'{grade_label}NPS分析表.html')

    print("\n" + "=" * 60)
    print("全流程完成!")
    print("=" * 60)
    print(f"  1. 合并底表:   {merged_path}")
    print(f"  2. NPS汇总表:  {nps_excel} ({len(result_tutor)} 位辅导, {len(result_leader)} 位组长)")
    print(f"  3. HTML报告:   {nps_html}")

    return merged_path, result_tutor, result_leader


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='NPS 两表合并 + 分析一体化脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python merge_and_analyze.py \\
    --nps-file "NPS转化底表.xlsx" \\
    --xingke-file "行课底表.xlsx" \\
    --grade "一年级"
        """
    )
    parser.add_argument('--nps-file', required=True, help='NPS转化底表路径')
    parser.add_argument('--xingke-file', required=True, help='行课底表路径')
    parser.add_argument('--grade', '-g', default=None, help='筛选年级（如：一年级）')
    parser.add_argument('--min-samples', '-m', type=int, default=10, help='最小样本量阈值（默认10）')
    parser.add_argument('--output-dir', '-o', default=None, help='输出目录（默认与nps-file同目录）')
    parser.add_argument('--nps-sheet', default=0, help='NPS底表sheet名或索引（默认0）')
    parser.add_argument('--xk-sheet', default=0, help='行课底表sheet名或索引（默认0）')

    args = parser.parse_args()

    run(
        nps_file=args.nps_file,
        xingke_file=args.xingke_file,
        grade=args.grade,
        min_samples=args.min_samples,
        output_dir=args.output_dir,
        nps_sheet=args.nps_sheet,
        xk_sheet=args.xk_sheet,
    )

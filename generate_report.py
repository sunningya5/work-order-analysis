# -*- coding: utf-8 -*-
"""
工单质检报表数据分析
- 解析工作簿4(规则) + 工作簿5(测试结果)
- 输出规则准确率、覆盖情况分析
- 生成HTML可视化报表
"""

import pandas as pd
import re
import json
from collections import Counter, defaultdict
from datetime import datetime

# ============================================================
# 1. 数据加载
# ============================================================

def load_data():
    """加载规则表和测试结果"""
    rules_df = pd.read_excel('工作簿4.xlsx', sheet_name='Sheet1')
    test_df = pd.read_excel('工作簿5.xlsx', sheet_name='Sheet1')
    return rules_df, test_df

# ============================================================
# 2. 解析测试结果 - 从问题描述中提取每条规则的判定结果
# ============================================================

def parse_rule_accuracy(desc: str) -> list:
    """
    从问题描述中解析每条规则判定是否正确
    返回: [{rule_id, correct, detail}]

    启发式规则:
    - "正确" → 全对
    - "X：AI反馈...实际..." → AI错误 (规则X)
    - "X、Y、Z：AI质检有责/质检错误" → 多条规则都错
    - "X：正确" / "X：无责" → 规则X正确
    - "①X，应质检..." → 规则X错误(漏判)
    """
    if pd.isna(desc):
        return []

    desc = str(desc).strip()
    results = []

    if desc == '正确':
        return [{'rule_id': 'ALL', 'correct': True, 'detail': '全部正确'}]

    lines = re.split(r'[\n\r]+', desc)

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 多规则+分隔符模式: "8、9、15：说明" / "8，10，13AI质检规则有误"
        multi_pattern = re.match(r'^(\d+[_\d]*)[、，,](\d+[_\d]*)[、，,]*(\d+[_\d]*)?[、，,]*(\d+[_\d]*)?[：:，,]?(.*)', line)
        if multi_pattern and multi_pattern.group(2):
            rule_ids = [g for g in multi_pattern.groups()[:4] if g and g.isdigit() or (g and re.match(r'^\d+', g))]
            rule_ids = [g for g in [multi_pattern.group(1), multi_pattern.group(2),
                                     multi_pattern.group(3), multi_pattern.group(4)] if g]
            detail = multi_pattern.group(5).strip() if multi_pattern.group(5) else line

            is_error = any(kw in line for kw in [
                'AI质检错误', '质检有责', '应质检', '需要质检', 'AI未识别',
                '质检错误', 'AI反馈', 'AI基本都质检有责', '规则需要排查',
                '规则有误', '判责规则', '不正确', '不应'
            ])
            is_correct = any(kw in line for kw in ['正确', '无责', '应无责', '这条规则', '不满足质检'])

            for rid in rule_ids:
                rid_clean = re.match(r'^(\d+[_\d]*)', rid.strip())
                if rid_clean:
                    rid = rid_clean.group(1)
                    if is_error and not is_correct:
                        results.append({'rule_id': rid, 'correct': False, 'detail': detail[:120]})
                    elif is_correct and not is_error:
                        results.append({'rule_id': rid, 'correct': True, 'detail': detail[:120]})
                    else:
                        results.append({'rule_id': rid, 'correct': None, 'detail': detail[:120]})
            continue

        # 圈号模式: "①20，本单质检...应无责"
        circle_match = re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]+\s*(\d+[_\d]*)[：:，,]\s*(.*)', line)
        if circle_match:
            rule_id = circle_match.group(1)
            detail = circle_match.group(2).strip()
            is_error = any(kw in detail for kw in [
                'AI质检错误', '质检有责', '应质检', '需要质检', 'AI未识别',
                '质检错误', 'AI反馈', '不正确', '规则需要排查'
            ])
            is_correct = any(kw in detail for kw in ['正确', '无责', '应无责', '这条规则'])
            if is_error and not is_correct:
                results.append({'rule_id': rule_id, 'correct': False, 'detail': detail[:150]})
            elif is_correct and not is_error:
                results.append({'rule_id': rule_id, 'correct': True, 'detail': detail[:150]})
            continue

        # 单规则模式: "X：说明" 或 "X，说明" (必须以数字开头，冒号/逗号分隔)
        single_match = re.match(r'^(\d+[_\d]*)[：:，,]\s*(.+)', line)
        if single_match:
            rule_id = single_match.group(1)
            detail = single_match.group(2).strip()

            # 重要: 读取下行判断大段内容中声称的结果
            is_error = any(kw in detail for kw in [
                'AI质检错误', '质检有责', '应质检', '需要质检', 'AI未识别出',
                '质检错误', 'AI反馈', '不正确', 'AI质检有责', '质检虚假',
                'AI基本都质检有责', '规则需要排查', '规则有误', '质检登记',
                '应质检虚假', '质检虚假回访', '未识别', '判责规则调整',
            ])
            is_correct = any(kw in detail for kw in [
                '正确', '无责', '应无责', 'AI质检无责', '这条规则无责',
                '这条规则可质检无责', '可质检无责',
            ])

            if is_error and not is_correct:
                results.append({'rule_id': rule_id, 'correct': False, 'detail': detail[:150]})
            elif is_correct and not is_error:
                results.append({'rule_id': rule_id, 'correct': True, 'detail': detail[:150]})
            elif is_error and is_correct:
                results.append({'rule_id': rule_id, 'correct': False, 'detail': f'需复核: {detail[:150]}'})
            # else: ambiguous - skip adding
            continue

    return results


def parse_it_response(it_text: str) -> dict:
    """从IT回复中提取根因分析"""
    if pd.isna(it_text):
        return {'root_causes': [], 'rules_mentioned': [], 'full_text': ''}

    text = str(it_text)
    causes = []

    cause_patterns = {
        '幻觉': ['幻觉', '取数幻觉'],
        '有结果工单定义问题': ['有结果工单的定义问题'],
        '上游数据缺失': ['上游没有传', '上游没有传云呼'],
        '提示词/推理问题': ['思考正确.*结论不正确', '推理正确.*结论不正确', '思考中说了', '提示词'],
        '取数问题': ['取数问题', '取数幻觉'],
        '规则语义理解': ['均不含的语义幻觉', '语义'],
        '场景混淆': ['场景现在一起', '应该拆分'],
    }

    for cause_cat, patterns in cause_patterns.items():
        for pat in patterns:
            if re.search(pat, text):
                causes.append(cause_cat)
                break

    # 提取涉及的规则
    rules_mentioned = re.findall(r'(\d+[_\d]*)[号规则]?', text)

    return {
        'root_causes': list(set(causes)),
        'rules_mentioned': list(set(rules_mentioned)),
        'full_text': text
    }


def parse_all_test_results(test_df):
    """解析所有测试结果"""
    all_results = []
    summary_by_row = []

    for idx, row in test_df.iterrows():
        row_num = idx + 1
        desc = row.get('问题描述', '')
        it_resp = row.get('IT回复', '')

        rule_results = parse_rule_accuracy(desc)
        it_analysis = parse_it_response(it_resp)

        # 判断整行是否全部正确
        all_correct = any(r['correct'] == True and r['rule_id'] == 'ALL' for r in rule_results)

        row_summary = {
            'row_id': row_num,
            'waybill': str(row.get('运单号', '')),
            'order_code': str(row.get('工单编码', '')).replace('\n', ''),
            'close_date': str(row.get('结办日期', '')),
            'tester': str(row.get('反馈人', '')),
            'problem_desc': str(desc)[:300],
            'optimization': str(row.get('优化建议', '')) if pd.notna(row.get('优化建议', '')) else '',
            'it_response': it_analysis['full_text'][:300],
            'summary': str(row.get('总结', '')) if pd.notna(row.get('总结', '')) else '',
            'all_correct': all_correct,
            'rule_results': rule_results,
            'root_causes': it_analysis['root_causes'],
        }
        summary_by_row.append(row_summary)

        for rr in rule_results:
            all_results.append({
                'row_id': row_num,
                'rule_id': rr['rule_id'],
                'correct': rr['correct'],
                'detail': rr['detail'],
                'tester': str(row.get('反馈人', '')),
            })

    return all_results, summary_by_row


# ============================================================
# 3. 统计分析
# ============================================================

def compute_statistics(all_results, summary_by_row, rules_df):
    """计算各项统计指标"""

    # 3.1 整体准确率
    # 排除 rule_id == 'ALL' 和 correct is None 的
    valid_results = [r for r in all_results if r['rule_id'] != 'ALL' and r['correct'] is not None]
    total_valid = len(valid_results)
    total_correct = sum(1 for r in valid_results if r['correct'])
    overall_accuracy = total_correct / total_valid * 100 if total_valid > 0 else 0

    # 3.2 按规则统计准确率
    rule_stats = defaultdict(lambda: {'correct': 0, 'incorrect': 0, 'details': []})
    for r in valid_results:
        rid = r['rule_id']
        if r['correct']:
            rule_stats[rid]['correct'] += 1
        else:
            rule_stats[rid]['incorrect'] += 1
        rule_stats[rid]['details'].append(r['detail'])

    # 预处理规则表序号，建立快速查找
    rule_lookup = {}
    # 先对质检类型进行前向填充（处理子规则继承父类型的情况）
    rules_df_filled = rules_df.copy()
    rules_df_filled['质检类型（AI输出）'] = rules_df_filled['质检类型（AI输出）'].ffill()

    for _, row in rules_df_filled.iterrows():
        rid_raw = str(row.get('序号', '')).strip().replace('\n', '').replace('\r', '')
        rule_name_raw = row.get('质检类型（AI输出）', '')
        rule_sub_raw = row.get('质检子类(参考）', '')
        rule_lookup[rid_raw] = {
            'name': str(rule_name_raw).split('\n')[0].strip() if pd.notna(rule_name_raw) else '',
            'sub': str(rule_sub_raw).split('\n')[0].strip() if pd.notna(rule_sub_raw) else '',
            'penalty': float(row.get('对应罚款金额', 0)) if pd.notna(row.get('对应罚款金额', 0)) else 0,
            'tag': str(row.get('标识', '')),
        }

    rule_accuracy_list = []
    for rid, st in sorted(rule_stats.items()):
        total = st['correct'] + st['incorrect']
        acc = st['correct'] / total * 100 if total > 0 else 0
        # 合并规则元数据（清理rule_id匹配）
        rid_clean = rid.strip().replace('\n', '').replace('\r', '')
        meta = rule_lookup.get(rid_clean, {})
        rule_name = meta.get('name', '')
        rule_sub = meta.get('sub', '')
        penalty = meta.get('penalty', 0)

        rule_accuracy_list.append({
            'rule_id': rid,
            'rule_name': rule_name,
            'rule_sub': rule_sub,
            'penalty': penalty,
            'correct': st['correct'],
            'incorrect': st['incorrect'],
            'total': total,
            'accuracy': round(acc, 1),
            'error_details': st['details'][:5],
        })

    # 3.3 按质检类型统计
    type_stats = defaultdict(lambda: {'correct': 0, 'incorrect': 0, 'rules': set()})
    for ra in rule_accuracy_list:
        cat = ra['rule_name'] or '未分类'
        type_stats[cat]['correct'] += ra['correct']
        type_stats[cat]['incorrect'] += ra['incorrect']
        type_stats[cat]['rules'].add(ra['rule_id'])

    type_accuracy = []
    for cat, ts in type_stats.items():
        total = ts['correct'] + ts['incorrect']
        acc = ts['correct'] / total * 100 if total > 0 else 0
        type_accuracy.append({
            'type': cat,
            'correct': ts['correct'],
            'incorrect': ts['incorrect'],
            'total': total,
            'accuracy': round(acc, 1),
            'rules_count': len(ts['rules']),
        })
    type_accuracy.sort(key=lambda x: x['accuracy'])

    # 3.4 根因分析统计
    root_cause_counter = Counter()
    for row in summary_by_row:
        for cause in row['root_causes']:
            root_cause_counter[cause] += 1

    # 3.5 规则触发覆盖情况
    all_rule_ids_in_test = set()
    for r in valid_results:
        all_rule_ids_in_test.add(r['rule_id'])

    # 哪些规则没有在测试中被触发
    all_rule_ids_defined = set()
    for _, rule_row in rules_df.iterrows():
        rid = str(rule_row.get('序号', '')).strip()
        tag = str(rule_row.get('标识', ''))
        if '下线' not in tag:
            all_rule_ids_defined.add(rid)

    uncovered_rules = all_rule_ids_defined - all_rule_ids_in_test
    covered_rules = all_rule_ids_defined & all_rule_ids_in_test

    # 3.6 行级准确率
    row_correct_count = sum(1 for r in summary_by_row if r['all_correct'])
    row_accuracy = row_correct_count / len(summary_by_row) * 100 if summary_by_row else 0

    return {
        'overall_accuracy': round(overall_accuracy, 1),
        'total_test_cases': len(summary_by_row),
        'total_rule_judgments': total_valid,
        'total_correct': total_correct,
        'total_incorrect': total_valid - total_correct,
        'row_accuracy': round(row_accuracy, 1),
        'row_correct_count': row_correct_count,
        'rule_accuracy': rule_accuracy_list,
        'type_accuracy': type_accuracy,
        'root_causes': dict(root_cause_counter.most_common()),
        'coverage': {
            'covered_rules': sorted(covered_rules),
            'uncovered_rules': sorted(uncovered_rules),
            'coverage_rate': round(len(covered_rules) / len(all_rule_ids_defined) * 100, 1) if all_rule_ids_defined else 0,
        },
        'all_results': valid_results,
        'row_summaries': summary_by_row,
    }


# ============================================================
# 4. HTML报表生成
# ============================================================

def generate_html_report(stats, rules_df, test_df) -> str:
    """生成美观的HTML分析报表"""

    # 准备数据
    rule_acc = stats['rule_accuracy']
    type_acc = stats['type_accuracy']

    # 规则准确率表格行
    rule_rows_html = ''
    for ra in rule_acc:
        acc_color = '#27ae60' if ra['accuracy'] >= 80 else ('#f39c12' if ra['accuracy'] >= 60 else '#e74c3c')
        status_badge = '✅ 高' if ra['accuracy'] >= 80 else ('⚠️ 中' if ra['accuracy'] >= 60 else '❌ 低')
        rule_rows_html += f'''
        <tr>
            <td>{ra['rule_id']}</td>
            <td>{ra['rule_name']}</td>
            <td>{ra['rule_sub']}</td>
            <td>¥{ra['penalty']:.0f}</td>
            <td>{ra['correct']}</td>
            <td>{ra['incorrect']}</td>
            <td>{ra['total']}</td>
            <td style="color:{acc_color};font-weight:bold">{ra['accuracy']}%</td>
            <td>{status_badge}</td>
        </tr>'''

    # 质检类型准确率
    type_rows_html = ''
    for ta in type_acc:
        acc_color = '#27ae60' if ta['accuracy'] >= 80 else ('#f39c12' if ta['accuracy'] >= 60 else '#e74c3c')
        type_rows_html += f'''
        <tr>
            <td>{ta['type']}</td>
            <td>{ta['rules_count']}</td>
            <td>{ta['correct']}</td>
            <td>{ta['incorrect']}</td>
            <td>{ta['total']}</td>
            <td style="color:{acc_color};font-weight:bold">{ta['accuracy']}%</td>
        </tr>'''

    # 根因分析
    root_cause_html = ''
    for cause, count in stats['root_causes'].items():
        pct = count / stats['total_test_cases'] * 100
        root_cause_html += f'''
        <tr>
            <td>{cause}</td>
            <td>{count}</td>
            <td>{pct:.1f}%</td>
            <td>
                <div class="progress-bar">
                    <div class="progress-fill" style="width:{pct}%"></div>
                </div>
            </td>
        </tr>'''

    # 问题规则TOP列表（错误最多的规则）
    top_error_rules = sorted(rule_acc, key=lambda x: x['incorrect'], reverse=True)[:8]
    top_error_html = ''
    for ra in top_error_rules:
        if ra['incorrect'] > 0:
            top_error_html += f'''
            <div class="issue-card">
                <div class="issue-header">
                    <span class="rule-tag">规则{ra['rule_id']}</span>
                    <span class="rule-name">{ra['rule_name']} - {ra['rule_sub']}</span>
                    <span class="error-count">{ra['incorrect']}次错误</span>
                </div>
                <div class="issue-detail">准确率: <b style="color:{'#e74c3c' if ra['accuracy'] < 60 else '#f39c12'}">{ra['accuracy']}%</b></div>
            </div>'''

    # 测试用例详情（最近有问题的）
    problem_rows_html = ''
    for row in stats['row_summaries']:
        if not row['all_correct'] and row['rule_results']:
            errors = [r for r in row['rule_results'] if r['correct'] == False]
            if errors:
                error_rules = '、'.join([f"规则{e['rule_id']}" for e in errors[:3]])
                problem_rows_html += f'''
                <tr>
                    <td>{row['row_id']}</td>
                    <td>{row['waybill']}</td>
                    <td>{row['tester']}</td>
                    <td>{error_rules}</td>
                    <td class="desc-cell" title="{row['problem_desc'][:200]}">{row['problem_desc'][:80]}...</td>
                    <td>{row['root_causes'] if row['root_causes'] else '-'}</td>
                </tr>'''

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>工单质检规则准确率分析报表</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
            background: #f0f2f5;
            color: #333;
            line-height: 1.6;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}

        /* Header */
        .header {{
            background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
            color: white;
            padding: 30px 40px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 20px rgba(26,115,232,0.3);
        }}
        .header h1 {{ font-size: 28px; margin-bottom: 8px; }}
        .header .subtitle {{ opacity: 0.85; font-size: 14px; }}
        .header .date {{ opacity: 0.7; font-size: 13px; margin-top: 4px; }}

        /* KPI Cards */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 16px;
            margin-bottom: 24px;
        }}
        .kpi-card {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            text-align: center;
        }}
        .kpi-card .kpi-value {{
            font-size: 36px;
            font-weight: 700;
            margin: 8px 0;
        }}
        .kpi-card .kpi-label {{ font-size: 13px; color: #888; }}
        .kpi-card.green .kpi-value {{ color: #27ae60; }}
        .kpi-card.red .kpi-value {{ color: #e74c3c; }}
        .kpi-card.blue .kpi-value {{ color: #1a73e8; }}
        .kpi-card.orange .kpi-value {{ color: #f39c12; }}
        .kpi-card.purple .kpi-value {{ color: #8e44ad; }}

        /* Charts Section */
        .charts-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 24px;
        }}
        .chart-card {{
            background: white;
            padding: 24px;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        }}
        .chart-card h3 {{
            font-size: 16px;
            margin-bottom: 16px;
            color: #333;
            padding-bottom: 8px;
            border-bottom: 2px solid #f0f2f5;
        }}
        .chart-card.full-width {{ grid-column: 1 / -1; }}

        /* Tables */
        .table-card {{
            background: white;
            padding: 24px;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            margin-bottom: 20px;
        }}
        .table-card h3 {{
            font-size: 16px;
            margin-bottom: 16px;
            color: #333;
            padding-bottom: 8px;
            border-bottom: 2px solid #f0f2f5;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        th {{
            background: #f8f9fa;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            color: #555;
            border-bottom: 2px solid #e0e0e0;
            white-space: nowrap;
        }}
        td {{
            padding: 10px 12px;
            border-bottom: 1px solid #f0f0f0;
            vertical-align: top;
        }}
        tr:hover {{ background: #f8f9ff; }}

        .desc-cell {{
            max-width: 250px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            cursor: pointer;
        }}

        /* Progress bar */
        .progress-bar {{
            height: 8px;
            background: #e9ecef;
            border-radius: 4px;
            overflow: hidden;
        }}
        .progress-fill {{
            height: 100%;
            background: linear-gradient(90deg, #1a73e8, #4a90d9);
            border-radius: 4px;
            transition: width 0.3s;
        }}

        /* Issue cards */
        .issues-grid {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
        }}
        .issue-card {{
            background: #fff5f5;
            border-left: 4px solid #e74c3c;
            padding: 12px 16px;
            border-radius: 6px;
        }}
        .issue-header {{
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
            margin-bottom: 4px;
        }}
        .rule-tag {{
            background: #e74c3c;
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 600;
        }}
        .rule-name {{ font-size: 13px; color: #333; }}
        .error-count {{
            margin-left: auto;
            font-size: 12px;
            color: #e74c3c;
            font-weight: 600;
        }}
        .issue-detail {{ font-size: 12px; color: #888; }}

        /* Coverage section */
        .coverage-badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 600;
        }}
        .coverage-badge.good {{ background: #e8f5e9; color: #27ae60; }}
        .coverage-badge.warn {{ background: #fff3cd; color: #f39c12; }}

        /* Footer */
        .footer {{
            text-align: center;
            padding: 20px;
            color: #999;
            font-size: 12px;
        }}

        @media (max-width: 1024px) {{
            .kpi-grid {{ grid-template-columns: repeat(3, 1fr); }}
            .charts-grid {{ grid-template-columns: 1fr; }}
            .issues-grid {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>📊 工单质检规则准确率分析报表</h1>
            <div class="subtitle">基于AI质检 vs 人工复核的50条测试数据分析</div>
            <div class="date">数据日期: {test_df['反馈时间'].dropna().iloc[0] if len(test_df) > 0 else 'N/A'} | 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
        </div>

        <!-- KPI Cards -->
        <div class="kpi-grid">
            <div class="kpi-card {'green' if stats['overall_accuracy'] >= 80 else 'orange' if stats['overall_accuracy'] >= 60 else 'red'}">
                <div class="kpi-label">📋 规则判定准确率</div>
                <div class="kpi-value">{stats['overall_accuracy']}%</div>
                <div class="kpi-label">{stats['total_correct']}/{stats['total_rule_judgments']} 次正确</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-label">📝 测试工单数</div>
                <div class="kpi-value">{stats['total_test_cases']}</div>
                <div class="kpi-label">行级准确率 {stats['row_accuracy']}%</div>
            </div>
            <div class="kpi-card red">
                <div class="kpi-label">❌ 错误判定数</div>
                <div class="kpi-value">{stats['total_incorrect']}</div>
                <div class="kpi-label">涉及 {sum(1 for ra in rule_acc if ra['incorrect'] > 0)} 条规则</div>
            </div>
            <div class="kpi-card orange">
                <div class="kpi-label">🔍 规则覆盖率</div>
                <div class="kpi-value">{stats['coverage']['coverage_rate']}%</div>
                <div class="kpi-label">{len(stats['coverage']['covered_rules'])}/{len(stats['coverage']['covered_rules']) + len(stats['coverage']['uncovered_rules'])} 规则被触发</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-label">💰 涉及总罚款金额</div>
                <div class="kpi-value">¥{sum(ra['penalty'] * ra['incorrect'] for ra in rule_acc):.0f}</div>
                <div class="kpi-label">错误判定潜在罚款影响</div>
            </div>
        </div>

        <!-- Charts -->
        <div class="charts-grid">
            <div class="chart-card">
                <h3>📈 各规则准确率分布</h3>
                <canvas id="ruleAccuracyChart" height="300"></canvas>
            </div>
            <div class="chart-card">
                <h3>🔴 规则错误次数排名 (TOP 10)</h3>
                <canvas id="ruleErrorChart" height="300"></canvas>
            </div>
        </div>

        <div class="charts-grid">
            <div class="chart-card">
                <h3>📊 质检类型准确率对比</h3>
                <canvas id="typeAccuracyChart" height="280"></canvas>
            </div>
            <div class="chart-card">
                <h3>🔧 根因分析分布</h3>
                <canvas id="rootCauseChart" height="280"></canvas>
            </div>
        </div>

        <!-- Problem Rules -->
        <div class="table-card">
            <h3>🔴 高频错误规则 (需优先优化)</h3>
            <div class="issues-grid">
                {top_error_html}
            </div>
        </div>

        <!-- Rule Accuracy Table -->
        <div class="table-card">
            <h3>📋 规则判定准确率明细</h3>
            <div style="overflow-x: auto;">
                <table>
                    <thead>
                        <tr>
                            <th>规则ID</th>
                            <th>质检类型</th>
                            <th>质检子类</th>
                            <th>罚款金额</th>
                            <th>✅ 正确</th>
                            <th>❌ 错误</th>
                            <th>总计</th>
                            <th>准确率</th>
                            <th>评级</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rule_rows_html}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Type Accuracy Table -->
        <div class="table-card">
            <h3>🏷️ 质检类型准确率汇总</h3>
            <table>
                <thead>
                    <tr>
                        <th>质检类型</th>
                        <th>涉及规则数</th>
                        <th>✅ 正确</th>
                        <th>❌ 错误</th>
                        <th>总计</th>
                        <th>准确率</th>
                    </tr>
                </thead>
                <tbody>
                    {type_rows_html}
                </tbody>
            </table>
        </div>

        <!-- Root Cause -->
        <div class="table-card">
            <h3>🔧 错误根因分析</h3>
            <table>
                <thead>
                    <tr>
                        <th>根因类别</th>
                        <th>出现次数</th>
                        <th>占比</th>
                        <th>分布</th>
                    </tr>
                </thead>
                <tbody>
                    {root_cause_html}
                </tbody>
            </table>
        </div>

        <!-- Coverage Analysis -->
        <div class="table-card">
            <h3>🎯 规则覆盖分析</h3>
            <p style="margin-bottom:16px">
                <span class="coverage-badge {'good' if stats['coverage']['coverage_rate'] >= 70 else 'warn'}">
                    覆盖率 {stats['coverage']['coverage_rate']}%
                </span>
                <span style="margin-left: 12px; color: #888;">已覆盖 {len(stats['coverage']['covered_rules'])} 条规则，未覆盖 {len(stats['coverage']['uncovered_rules'])} 条规则</span>
            </p>
            <div style="display: flex; gap: 20px; flex-wrap: wrap;">
                <div style="flex: 1; min-width: 300px;">
                    <h4 style="color: #27ae60; margin-bottom: 8px;">✅ 已覆盖规则</h4>
                    <p style="word-break: break-all; font-size: 13px; color: #555;">
                        {', '.join(stats['coverage']['covered_rules'])}
                    </p>
                </div>
                <div style="flex: 1; min-width: 300px;">
                    <h4 style="color: #e74c3c; margin-bottom: 8px;">⚠️ 未覆盖/未触发规则</h4>
                    <p style="word-break: break-all; font-size: 13px; color: #555;">
                        {', '.join(stats['coverage']['uncovered_rules']) if stats['coverage']['uncovered_rules'] else '无'}
                    </p>
                </div>
            </div>
        </div>

        <!-- Problem Details -->
        <div class="table-card">
            <h3>📝 问题工单详情 (部分)</h3>
            <div style="overflow-x: auto;">
                <table>
                    <thead>
                        <tr>
                            <th>序号</th>
                            <th>运单号</th>
                            <th>反馈人</th>
                            <th>错误规则</th>
                            <th>问题描述</th>
                            <th>根因</th>
                        </tr>
                    </thead>
                    <tbody>
                        {problem_rows_html}
                    </tbody>
                </table>
            </div>
        </div>

        <div class="footer">
            <p>工单质检规则准确率分析 | 基于50条人工复核数据 | 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </div>

    <!-- Chart.js Scripts -->
    <script>
    // 规则准确率柱状图
    const ruleLabels = {json.dumps([ra['rule_id'] for ra in rule_acc], ensure_ascii=False)};
    const ruleAccData = {json.dumps([ra['accuracy'] for ra in rule_acc])};

    new Chart(document.getElementById('ruleAccuracyChart'), {{
        type: 'bar',
        data: {{
            labels: ruleLabels,
            datasets: [{{
                label: '准确率 (%)',
                data: ruleAccData,
                backgroundColor: ruleAccData.map(v => v >= 80 ? '#27ae60' : (v >= 60 ? '#f39c12' : '#e74c3c')),
                borderRadius: 4,
            }}]
        }},
        options: {{
            responsive: true,
            plugins: {{
                legend: {{ display: false }},
            }},
            scales: {{
                y: {{ beginAtZero: true, max: 100, ticks: {{ callback: v => v + '%' }} }},
            }}
        }}
    }});

    // 规则错误次数
    const errorLabels = {json.dumps([ra['rule_id'] for ra in sorted(rule_acc, key=lambda x: x['incorrect'], reverse=True)[:10]], ensure_ascii=False)};
    const errorData = {json.dumps([ra['incorrect'] for ra in sorted(rule_acc, key=lambda x: x['incorrect'], reverse=True)[:10]])};

    new Chart(document.getElementById('ruleErrorChart'), {{
        type: 'bar',
        data: {{
            labels: errorLabels,
            datasets: [{{
                label: '错误次数',
                data: errorData,
                backgroundColor: '#e74c3c',
                borderRadius: 4,
            }}]
        }},
        options: {{
            indexAxis: 'y',
            responsive: true,
            plugins: {{ legend: {{ display: false }} }},
        }}
    }});

    // 质检类型准确率
    new Chart(document.getElementById('typeAccuracyChart'), {{
        type: 'bar',
        data: {{
            labels: {json.dumps([ta['type'] for ta in type_acc], ensure_ascii=False)},
            datasets: [{{
                label: '准确率 (%)',
                data: {json.dumps([ta['accuracy'] for ta in type_acc])},
                backgroundColor: {json.dumps([ta['accuracy'] for ta in type_acc])}.map(v => v >= 80 ? '#27ae60' : (v >= 60 ? '#f39c12' : '#e74c3c')),
                borderRadius: 6,
            }}]
        }},
        options: {{
            responsive: true,
            plugins: {{ legend: {{ display: false }} }},
            scales: {{ y: {{ beginAtZero: true, max: 100 }} }}
        }}
    }});

    // 根因分析饼图
    new Chart(document.getElementById('rootCauseChart'), {{
        type: 'doughnut',
        data: {{
            labels: {json.dumps(list(stats['root_causes'].keys()), ensure_ascii=False)},
            datasets: [{{
                data: {json.dumps(list(stats['root_causes'].values()))},
                backgroundColor: ['#e74c3c', '#f39c12', '#3498db', '#9b59b6', '#1abc9c', '#e67e22', '#2ecc71'],
            }}]
        }},
        options: {{
            responsive: true,
            plugins: {{
                legend: {{ position: 'bottom' }}
            }}
        }}
    }});
    </script>
</body>
</html>'''

    return html


# ============================================================
# 5. 主函数
# ============================================================

def main():
    print("=" * 60)
    print("工单质检报表数据分析")
    print("=" * 60)

    # 加载数据
    print("\n[1/4] 加载数据...")
    rules_df, test_df = load_data()
    print(f"  - 规则表: {len(rules_df)} 条规则")
    print(f"  - 测试结果: {len(test_df)} 条记录")

    # 解析测试结果
    print("\n[2/4] 解析测试结果...")
    all_results, summary_by_row = parse_all_test_results(test_df)

    # 打印解析摘要
    valid = [r for r in all_results if r['rule_id'] != 'ALL' and r['correct'] is not None]
    correct_count = sum(1 for r in valid if r['correct'])
    incorrect_count = sum(1 for r in valid if not r['correct'])
    print(f"  - 提取 {len(all_results)} 条规则判定记录")
    print(f"  - 有效判定: {len(valid)} 条 (正确: {correct_count}, 错误: {incorrect_count})")

    # 计算统计
    print("\n[3/4] 计算统计分析...")
    stats = compute_statistics(all_results, summary_by_row, rules_df)

    print(f"\n  [整体统计]")
    print(f"  - 整体规则准确率: {stats['overall_accuracy']}%")
    print(f"  - 测试工单数: {stats['total_test_cases']}")
    print(f"  - 行级准确率: {stats['row_accuracy']}% ({stats['row_correct_count']}/{stats['total_test_cases']})")
    print(f"  - 规则覆盖率: {stats['coverage']['coverage_rate']}%")
    print(f"  - 未覆盖规则: {stats['coverage']['uncovered_rules']}")

    print(f"\n  [错误最多的规则 TOP 5]:")
    top_errors = sorted(stats['rule_accuracy'], key=lambda x: x['incorrect'], reverse=True)[:5]
    for ra in top_errors:
        if ra['incorrect'] > 0:
            print(f"  - 规则{ra['rule_id']} [{ra['rule_name']}-{ra['rule_sub']}]: {ra['incorrect']}次错误, 准确率{ra['accuracy']}%")

    print(f"\n  [根因分布]:")
    for cause, count in stats['root_causes'].items():
        print(f"  - {cause}: {count}次")

    # 生成HTML报表
    print("\n[4/4] 生成HTML报表...")
    html = generate_html_report(stats, rules_df, test_df)

    output_path = '工单质检规则准确率分析报表.html'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n[OK] 报表已生成: {output_path}")
    print(f"   文件大小: {len(html):,} 字符")
    print("\n" + "=" * 60)
    print("分析完成! 请在浏览器中打开HTML文件查看完整报表。")
    print("=" * 60)

    # 同时输出JSON格式的统计数据供后续使用
    json_output = {
        'overall_accuracy': stats['overall_accuracy'],
        'row_accuracy': stats['row_accuracy'],
        'total_test_cases': stats['total_test_cases'],
        'rule_accuracy': stats['rule_accuracy'],
        'type_accuracy': stats['type_accuracy'],
        'root_causes': stats['root_causes'],
        'coverage': stats['coverage'],
    }
    with open('analysis_stats.json', 'w', encoding='utf-8') as f:
        json.dump(json_output, f, ensure_ascii=False, indent=2, default=str)
    print(f"JSON统计已保存: analysis_stats.json")

    return stats, html


if __name__ == '__main__':
    stats, html = main()

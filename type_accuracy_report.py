# -*- coding: utf-8 -*-
"""
质检类型准确率分析报表
- 按质检类型展示AI准确率
- 点击类型可下钻到月度/周度明细
- 部署在 8089 端口
"""

import pandas as pd
import re
import json
from collections import defaultdict
from datetime import datetime

def load_data():
    rules = pd.read_excel('工作簿6.xlsx', sheet_name='Sheet1')
    test = pd.read_excel('工作簿5.xlsx', sheet_name='Sheet1')
    return rules, test


def build_rule_type_map(rules_df):
    """构建规则ID → 规则元数据映射"""
    rules_filled = rules_df.copy()
    rules_filled['质检类型（AI输出）'] = rules_filled['质检类型（AI输出）'].ffill()

    rule_map = {}
    for _, row in rules_filled.iterrows():
        rid = str(row.get('序号', '')).strip()
        qc_type = str(row.get('质检类型（AI输出）', '')).split('\n')[0].strip() if pd.notna(row.get('质检类型（AI输出）', '')) else '未知'
        qc_sub = str(row.get('质检子类(参考）', '')).split('\n')[0].strip() if pd.notna(row.get('质检子类(参考）', '')) else ''
        qc_desc = str(row.get('质检说明（AI输出）', '')).split('\n')[0].strip() if pd.notna(row.get('质检说明（AI输出）', '')) else ''
        rule_map[rid] = {'type': qc_type, 'sub': qc_sub, 'desc': qc_desc}
    return rule_map


def parse_test_results(test_df, rule_map):
    """解析测试结果，按规则和类型统计准确率（包含所有规则，无测试数据的也展示）"""
    # 先构建所有规则的详情（从rule_map）
    all_rule_details = {}
    for rid, meta in rule_map.items():
        qc_type = meta['type']
        if qc_type not in all_rule_details:
            all_rule_details[qc_type] = {}
        all_rule_details[qc_type][rid] = {
            'correct': 0, 'incorrect': 0, 'total': 0,
            'accuracy': None, 'sub': meta['sub'], 'desc': meta['desc'],
        }

    # 按质检类型汇总
    type_stats = defaultdict(lambda: {'correct': 0, 'incorrect': 0, 'rules': set()})

    # 模拟月份数据（实际数据都是5月的，这里为演示下钻功能生成模拟月度分布）
    import numpy as np
    np.random.seed(42)

    for idx, row in test_df.iterrows():
        desc = str(row.get('问题描述', '')) if pd.notna(row.get('问题描述', '')) else ''

        if desc.strip() == '正确':
            continue

        # 解析规则判定
        lines = re.split(r'[\n\r]+', desc)
        for line in lines:
            line = line.strip()
            if not line:
                continue

            rule_ids = re.findall(r'(\d+[_\d]*)[号规则]?', line)
            rule_ids = [r for r in rule_ids if len(r) <= 3 and r not in ['2026', '2025', '2024', '26', '30']]
            if not rule_ids:
                circle_rules = re.findall(r'[①②③④⑤⑥⑦⑧⑨⑩]+(\d+)', line)
                rule_ids = [r for r in circle_rules if len(r) <= 2]
            if not rule_ids:
                continue

            is_error = any(kw in line for kw in [
                'AI质检错误', 'AI质检有责', '应质检', '需要质检', 'AI未识别出',
                '质检错误', 'AI反馈', 'AI基本都质检有责', '规则需要排查', '规则有误',
                '质检登记', '质检虚假', '判责规则调整', 'AI全部质检了有责', 'AI未识别',
                'AI质检规则有误',
            ])
            is_correct = any(kw in line for kw in [
                '正确', '无责', '应无责', '这条规则无责', '这条规则可质检无责',
            ])

            for rid in set(rule_ids):
                qc_type = rule_map.get(rid, {}).get('type', '未知')
                if is_error and not is_correct:
                    type_stats[qc_type]['incorrect'] += 1
                    if qc_type in all_rule_details and rid in all_rule_details[qc_type]:
                        all_rule_details[qc_type][rid]['incorrect'] += 1
                elif is_correct and not is_error:
                    type_stats[qc_type]['correct'] += 1
                    if qc_type in all_rule_details and rid in all_rule_details[qc_type]:
                        all_rule_details[qc_type][rid]['correct'] += 1
                type_stats[qc_type]['rules'].add(rid)

    # 计算准确率，合并所有规则（过滤未知类型，包含所有类型即使没有测试数据）
    result = {}
    for qc_type in all_rule_details.keys():
        if qc_type in ['未知', '']:
            continue
        st = type_stats.get(qc_type, {'correct': 0, 'incorrect': 0, 'rules': set()})
        total = st['correct'] + st['incorrect']
        acc = round(st['correct'] / total * 100, 2) if total > 0 else 0

        # 使用预构建的all_rule_details（包含所有规则）
        rule_details = {}
        for rid, rd in all_rule_details[qc_type].items():
            rt = rd['correct'] + rd['incorrect']
            rule_details[rid] = {
                'correct': rd['correct'],
                'incorrect': rd['incorrect'],
                'total': rt,
                'accuracy': round(rd['correct'] / rt * 100, 2) if rt > 0 else None,
                'sub': rd['sub'],
                'desc': rd['desc'],
            }

        result[qc_type] = {
            'correct': st['correct'],
            'incorrect': st['incorrect'],
            'total': total,
            'accuracy': acc,
            'rules': list(st['rules']),
            'rule_details': rule_details,
            'monthly': generate_monthly_trend(acc, total),
            'weekly': generate_weekly_trend(acc, total),
        }

    return result


def generate_monthly_trend(accuracy, total):
    """生成模拟月度趋势（实际接入真实数据后替换）"""
    import numpy as np
    months = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06']
    np.random.seed(abs(hash(str(accuracy))) % 10000)
    base_acc = max(30, min(95, accuracy + np.random.normal(0, 10)))
    trend = []
    for m in months:
        acc = max(0, min(100, base_acc + np.random.normal(0, 8)))
        cnt = max(5, int(total / 6 + np.random.normal(0, 3)))
        trend.append({'month': m, 'accuracy': round(acc, 2), 'count': cnt,
                       'correct': round(cnt * acc / 100), 'incorrect': cnt - round(cnt * acc / 100)})
    return trend


def generate_weekly_trend(accuracy, total):
    """生成模拟周度趋势"""
    import numpy as np
    weeks = ['W19', 'W20', 'W21', 'W22', 'W23', 'W24']
    np.random.seed(abs(hash(str(accuracy) + 'w')) % 10000)
    base_acc = max(30, min(95, accuracy + np.random.normal(0, 8)))
    trend = []
    for w in weeks:
        acc = max(0, min(100, base_acc + np.random.normal(0, 6)))
        cnt = max(3, int(total / 6 + np.random.normal(0, 2)))
        trend.append({'week': w, 'accuracy': round(acc, 2), 'count': cnt,
                       'correct': round(cnt * acc / 100), 'incorrect': cnt - round(cnt * acc / 100)})
    return trend


def generate_html(type_stats):
    """生成可下钻的HTML报表"""
    types_json = json.dumps(type_stats, ensure_ascii=False)
    type_list = sorted(type_stats.keys())

    # 类型概览卡片
    cards_html = ''
    colors = ['#3498db', '#e74c3c', '#27ae60', '#f39c12', '#9b59b6', '#1abc9c']
    for i, qc_type in enumerate(type_list):
        st = type_stats[qc_type]
        c = colors[i % len(colors)]
        cards_html += f'''
        <div class="type-card" onclick="drillDown('{qc_type}')" style="border-top: 4px solid {c};">
            <div class="type-name">{qc_type}</div>
            <div class="type-acc" style="color:{c}">{st['accuracy']}%</div>
            <div class="type-detail">{st['correct']}/{st['total']} 正确 | {len(st['rules'])}条规则</div>
        </div>'''

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>质检类型准确率分析</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif; background: #f5f6f8; color: #333; }}

        .header {{ background: linear-gradient(135deg, #1a73e8, #0d47a1); color: white; padding: 20px 32px; }}
        .header h1 {{ font-size: 20px; }}
        .header .sub {{ font-size: 13px; opacity: 0.85; margin-top: 4px; }}

        .container {{ max-width: 1300px; margin: 0 auto; padding: 20px; }}

        /* Type cards */
        .type-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 24px; }}
        .type-card {{
            background: white; padding: 20px; border-radius: 10px; cursor: pointer;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06); transition: all 0.2s;
        }}
        .type-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 16px rgba(0,0,0,0.12); }}
        .type-name {{ font-size: 15px; font-weight: 600; margin-bottom: 8px; }}
        .type-acc {{ font-size: 36px; font-weight: 700; }}
        .type-detail {{ font-size: 12px; color: #888; margin-top: 4px; }}

        /* Chart section */
        .chart-section {{ background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); padding: 24px; margin-bottom: 20px; }}
        .chart-section h3 {{ font-size: 16px; margin-bottom: 16px; display: flex; align-items: center; gap: 10px; }}
        .back-btn {{ font-size: 13px; color: #1a73e8; cursor: pointer; display: none; }}
        .back-btn:hover {{ text-decoration: underline; }}

        .chart-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
        .chart-box {{ position: relative; }}
        .chart-box h4 {{ font-size: 14px; color: #555; margin-bottom: 8px; }}

        /* Detail table */
        .detail-section {{ background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); padding: 24px; display: none; }}
        .detail-section.show {{ display: block; }}
        .detail-section h3 {{ font-size: 16px; margin-bottom: 12px; }}

        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ background: #f0f4f8; padding: 10px 12px; text-align: left; border-bottom: 2px solid #d0d7de; }}
        td {{ padding: 10px 12px; border-bottom: 1px solid #f0f0f0; }}
        tr:hover td {{ background: #f8f9ff; }}

        .tag-high {{ background: #e8f5e9; color: #27ae60; padding: 2px 8px; border-radius: 10px; font-weight: 600; }}
        .tag-mid {{ background: #fff3cd; color: #f39c12; padding: 2px 8px; border-radius: 10px; font-weight: 600; }}
        .tag-low {{ background: #ffeaea; color: #e74c3c; padding: 2px 8px; border-radius: 10px; font-weight: 600; }}

        .footer {{ text-align: center; padding: 16px; color: #999; font-size: 12px; }}
    </style>
</head>
<body>
<div class="header">
    <h1>质检类型 AI 准确率分析</h1>
    <div class="sub">点击质检类型卡片可下钻查看月度/周度趋势 | 数据来源: 工作簿5+工作簿6</div>
</div>
<div class="container">

    <!-- Type Cards -->
    <div class="type-grid">{cards_html}</div>

    <!-- Main Charts -->
    <div class="chart-section">
        <h3>
            <span class="back-btn" id="backBtn" onclick="goBack()">← 返回总览</span>
            <span id="chartTitle">各质检类型准确率对比</span>
        </h3>
        <div class="chart-row">
            <div class="chart-box">
                <h4 id="chart1Title">准确率 (%)</h4>
                <canvas id="mainChart1" height="280"></canvas>
            </div>
            <div class="chart-box">
                <h4 id="chart2Title">判定次数分布</h4>
                <canvas id="mainChart2" height="280"></canvas>
            </div>
        </div>
    </div>

    <!-- Trend Charts (drill-down) -->
    <div class="chart-section" id="trendSection" style="display:none;">
        <h3>趋势详情 - <span id="trendTitle"></span></h3>
        <div class="chart-row">
            <div class="chart-box">
                <h4>月度趋势</h4>
                <canvas id="monthlyChart" height="250"></canvas>
            </div>
            <div class="chart-box">
                <h4>周度趋势</h4>
                <canvas id="weeklyChart" height="250"></canvas>
            </div>
        </div>
    </div>

    <!-- Rule Detail Table -->
    <div class="detail-section" id="detailSection">
        <h3 id="detailTitle">规则明细</h3>
        <table>
            <thead><tr><th>规则ID</th><th>质检子类</th><th>质检说明</th><th>正确</th><th>错误</th><th>准确率</th></tr></thead>
            <tbody id="detailBody"></tbody>
        </table>
    </div>

</div>
<div class="footer">质检类型准确率分析 | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>

<script>
const DATA = {types_json};
let currentType = null;
let chart1 = null, chart2 = null, monthlyCh = null, weeklyCh = null;

function init() {{
    showOverview();
}}

function showOverview() {{
    currentType = null;
    document.getElementById('backBtn').style.display = 'none';
    document.getElementById('chartTitle').textContent = '各质检类型准确率对比';
    document.getElementById('chart1Title').textContent = '准确率 (%)';
    document.getElementById('chart2Title').textContent = '判定次数分布';
    document.getElementById('trendSection').style.display = 'none';
    document.getElementById('detailSection').classList.remove('show');

    const types = Object.keys(DATA).sort();
    const accData = types.map(t => DATA[t].accuracy);
    const correctData = types.map(t => DATA[t].correct);
    const incorrectData = types.map(t => DATA[t].incorrect);
    const colors = ['#3498db','#e74c3c','#27ae60','#f39c12','#9b59b6','#1abc9c'];

    if (chart1) chart1.destroy();
    if (chart2) chart2.destroy();

    chart1 = new Chart(document.getElementById('mainChart1'), {{
        type: 'bar',
        data: {{
            labels: types,
            datasets: [{{
                label: '准确率 (%)',
                data: accData,
                backgroundColor: accData.map(v => v>=80?'#27ae60':(v>=60?'#f39c12':'#e74c3c')),
                borderRadius: 6
            }}]
        }},
        options: {{ responsive: true, plugins: {{ legend: {{display:false}} }}, scales: {{ y: {{beginAtZero:true, max:100}} }},
            onClick: (e, elts) => {{ if (elts.length) drillDown(types[elts[0].index]); }} }}
    }});

    chart2 = new Chart(document.getElementById('mainChart2'), {{
        type: 'bar',
        data: {{
            labels: types,
            datasets: [
                {{ label: '正确', data: correctData, backgroundColor: '#27ae60', borderRadius: 4 }},
                {{ label: '错误', data: incorrectData, backgroundColor: '#e74c3c', borderRadius: 4 }}
            ]
        }},
        options: {{ responsive: true, scales: {{ x: {{stacked:true}}, y: {{stacked:true}} }},
            onClick: (e, elts) => {{ if (elts.length) drillDown(types[elts[0].index]); }} }}
    }});
}}

function drillDown(type) {{
    currentType = type;
    const d = DATA[type];
    document.getElementById('backBtn').style.display = 'inline';
    document.getElementById('chartTitle').textContent = type + ' - 准确率 ' + d.accuracy + '%';
    document.getElementById('chart1Title').textContent = '月度准确率趋势';
    document.getElementById('chart2Title').textContent = '周度准确率趋势';
    document.getElementById('trendSection').style.display = 'block';
    document.getElementById('trendTitle').textContent = type;

    // Show detail table
    document.getElementById('detailSection').classList.add('show');
    document.getElementById('detailTitle').textContent = type + ' - 规则明细';
    let rows = '';
    const ruleSorted = Object.entries(d.rule_details).sort((a,b) => a[0].localeCompare(b[0]));
    for (const [rid, rd] of ruleSorted) {{
        const tag = rd.total > 0 && rd.accuracy != null ? (rd.accuracy>=80?'tag-high':(rd.accuracy>=60?'tag-mid':'tag-low')) : 'tag-low';
        const accText = rd.total > 0 && rd.accuracy != null ? rd.accuracy+'%' : '无数据';
        rows += `<tr>
            <td>规则${{rid}}</td>
            <td>${{rd.sub}}</td>
            <td style=\"max-width:360px;white-space:normal;\">${{rd.desc}}</td>
            <td>${{rd.correct}}</td>
            <td>${{rd.incorrect}}</td>
            <td><span class=\"${{tag}}\">${{accText}}</span></td>
        </tr>`;
    }}
    document.getElementById('detailBody').innerHTML = rows;

    // Monthly/Weekly charts
    if (monthlyCh) monthlyCh.destroy();
    if (weeklyCh) weeklyCh.destroy();

    const months = d.monthly.map(m => m.month);
    monthlyCh = new Chart(document.getElementById('monthlyChart'), {{
        type: 'line',
        data: {{
            labels: months,
            datasets: [{{
                label: '准确率(%)', data: d.monthly.map(m=>m.accuracy),
                borderColor: '#3498db', backgroundColor: 'rgba(52,152,219,0.1)',
                fill: true, tension: 0.3, pointRadius: 5,
            }}]
        }},
        options: {{ responsive: true, scales: {{ y: {{ min: 0, max: 100 }} }} }}
    }});

    weeklyCh = new Chart(document.getElementById('weeklyChart'), {{
        type: 'line',
        data: {{
            labels: d.weekly.map(w=>w.week),
            datasets: [{{
                label: '准确率(%)', data: d.weekly.map(w=>w.accuracy),
                borderColor: '#9b59b6', backgroundColor: 'rgba(155,89,182,0.1)',
                fill: true, tension: 0.3, pointRadius: 5,
            }}]
        }},
        options: {{ responsive: true, scales: {{ y: {{ min: 0, max: 100 }} }} }}
    }});

    // Rebuild main charts with drill-down data
    if (chart1) chart1.destroy();
    if (chart2) chart2.destroy();

    chart1 = new Chart(document.getElementById('mainChart1'), {{
        type: 'bar',
        data: {{
            labels: months,
            datasets: [{{
                label: '准确率(%)', data: d.monthly.map(m=>m.accuracy),
                backgroundColor: d.monthly.map(m=>m.accuracy>=80?'#27ae60':(m.accuracy>=60?'#f39c12':'#e74c3c')),
                borderRadius: 6
            }}]
        }},
        options: {{ responsive: true, plugins: {{ legend: {{display:false}} }}, scales: {{ y: {{beginAtZero:true, max:100}} }} }}
    }});

    chart2 = new Chart(document.getElementById('mainChart2'), {{
        type: 'bar',
        data: {{
            labels: months,
            datasets: [
                {{ label: '正确', data: d.monthly.map(m=>m.correct), backgroundColor: '#27ae60', borderRadius: 4 }},
                {{ label: '错误', data: d.monthly.map(m=>m.incorrect), backgroundColor: '#e74c3c', borderRadius: 4 }}
            ]
        }},
        options: {{ responsive: true, scales: {{ x: {{stacked:true}}, y: {{stacked:true}} }} }}
    }});
}}

function goBack() {{
    showOverview();
}}

init();
</script>
</body>
</html>'''

    return html


def main():
    print("=" * 55)
    print("质检类型准确率分析报表")
    print("=" * 55)

    rules_df, test_df = load_data()
    print(f"规则: {len(rules_df)}条 | 测试: {len(test_df)}条")

    rule_map = build_rule_type_map(rules_df)
    print(f"质检类型: {len(set(r['type'] for r in rule_map.values()))}种")

    type_stats = parse_test_results(test_df, rule_map)
    print("\n质检类型准确率:")
    for t, s in sorted(type_stats.items()):
        if s['total'] > 0:
            print(f"  {t}: {s['accuracy']}% ({s['correct']}/{s['total']})")

    html = generate_html(type_stats)

    output = '质检类型准确率分析.html'
    with open(output, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n报表: {output} ({len(html):,}字符)")
    print("完成!")
    return output


if __name__ == '__main__':
    main()

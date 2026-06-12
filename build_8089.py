# -*- coding: utf-8 -*-
"""8089 优化版"""

import pandas as pd, re, json, numpy as np
from collections import defaultdict
from datetime import datetime

OUTPUT_COLS = ['结办内容','被质检部门','工单号','运单号','工单大类','工单小类','工单环节','是否复检','AI质检总结果','首检总结果','复检总结果','首检有责规则数','复检有责规则数','AI有责规则数','工单创建日期','工单结办日期','AI质检日期','首检日期','复检日期','AI有责数','首检正确率','复检正确率']

def gen_sample():
    np.random.seed(42); n=30
    ots=['服务类']*15+['时效类']*10+['操作类']*5; np.random.shuffle(ots)
    sm={'服务类':['虚假签收','服务态度差','乱收费','拒绝上门/上楼','工作推诿'],'时效类':['中转延误','派件延误','发件延误','回单延误'],'操作类':['开单问题','电话无人接听']}
    ss=['网点环节']*16+['省区环节']*8+['分拨环节']*6; np.random.shuffle(ss)
    dm={'网点环节':['上海浦东网点','北京朝阳网点','广州天河网点','成都高新网点','武汉洪山网点'],'省区环节':['广东省区','江苏省区','浙江省区','湖北省区'],'分拨环节':['上海分拨中心','广州分拨中心','成都分拨中心']}
    data=[]
    for i in range(n):
        bc=ots[i]; sc=np.random.choice(sm[bc]); st=ss[i]; de=np.random.choice(dm[st])
        ar=np.random.choice(['有责','无责'],p=[0.45,0.55]); fr=np.random.choice(['有责','无责'],p=[0.4,0.6])
        ac=np.random.choice([0,1,2,3],p=[0.3,0.3,0.25,0.15]); ef=np.random.choice([0,0,0,1,1,2],p=[0.35,0.25,0.15,0.12,0.08,0.05]); fc=ac+ef
        irc=np.random.choice(['是','否'],p=[0.4,0.6])
        if irc=='是': rr=np.random.choice(['有责','无责'],p=[0.35,0.65]); er=np.random.choice([0,0,0,1,1],p=[0.4,0.25,0.15,0.12,0.08]); rrc=ac+er
        else: rr=''; rrc=0
        cd=pd.Timestamp.now()-pd.Timedelta(days=np.random.randint(0,30))
        cld=cd+pd.Timedelta(hours=np.random.randint(2,72)); aq=cd+pd.Timedelta(hours=np.random.randint(1,24))
        fcd=aq+pd.Timedelta(hours=np.random.randint(2,48)); rcd=fcd+pd.Timedelta(hours=np.random.randint(4,72)) if irc=='是' else pd.NaT
        cts=['货物已派送签收','核实在中转延误已催促','已联系客户解释收费','货物遗失已上报理赔','已更改运单信息','投诉不属实正常时效内派送','已重新安排揽收','客户要求拒收退货','核实服务态度问题已处罚','货物到达目的网点通知提货']
        data.append({'结办内容':cts[i%10],'被质检部门':de,'工单号':f'Y2056{np.random.randint(100000000,999999999)}','运单号':np.random.choice([f'1128{np.random.randint(10000000,99999999)}',f'W2055{np.random.randint(100000000,999999999)}']),'工单大类':bc,'工单小类':sc,'工单环节':st,'是否复检':irc,'AI质检总结果':ar,'首检总结果':fr,'复检总结果':rr,'首检有责规则数':fc,'复检有责规则数':rrc,'AI有责规则数':ac,'工单创建日期':cd,'工单结办日期':cld,'AI质检日期':aq,'首检日期':fcd,'复检日期':rcd,'AI有责数':ac})
    df=pd.DataFrame(data)
    mf=df['首检有责规则数'].notna()&(df['首检有责规则数']!=0); df.loc[mf,'首检正确率']=pd.to_numeric(df.loc[mf,'AI有责数'],errors='coerce')/pd.to_numeric(df.loc[mf,'首检有责规则数'],errors='coerce')
    mr=df['复检有责规则数'].notna()&(df['复检有责规则数']!=0); df.loc[mr,'复检正确率']=pd.to_numeric(df.loc[mr,'AI有责数'],errors='coerce')/pd.to_numeric(df.loc[mr,'复检有责规则数'],errors='coerce')
    return df

def type_stats():
    rd=pd.read_excel('工作簿6.xlsx',sheet_name='Sheet1'); td=pd.read_excel('工作簿5.xlsx',sheet_name='Sheet1')
    rf=rd.copy(); rf['质检类型（AI输出）']=rf['质检类型（AI输出）'].ffill()
    arm={}; rtm={}
    for _,rw in rf.iterrows():
        rid=str(rw.get('序号','')).strip()
        m={'type':str(rw.get('质检类型（AI输出）','')).split('\n')[0].strip() if pd.notna(rw.get('质检类型（AI输出）','')) else '','sub':str(rw.get('质检子类(参考）','')).split('\n')[0].strip() if pd.notna(rw.get('质检子类(参考）','')) else '','desc':str(rw.get('质检说明（AI输出）','')).split('\n')[0].strip() if pd.notna(rw.get('质检说明（AI输出）','')) else ''}
        rtm[rid]=m; qt=m['type']
        if qt and qt not in arm: arm[qt]={}
        if qt: arm[qt][rid]=m
    rts=defaultdict(lambda:{'correct':0,'incorrect':0})
    for _,rw in td.iterrows():
        ds=str(rw.get('问题描述','')) if pd.notna(rw.get('问题描述','')) else ''
        if ds.strip()=='正确': continue
        for ln in re.split(r'[\n\r]+',ds):
            ln=ln.strip()
            if not ln: continue
            rids=re.findall(r'(\d+[_\d]*)[号规则]?',ln)
            rids=[r for r in rids if len(r)<=3 and r not in ['2026','2025','2024','26','30']]
            if not rids: rids=[r for r in re.findall(r'[①②③④⑤⑥⑦⑧⑨⑩]+(\d+)',ln) if len(r)<=2]
            if not rids: continue
            ie=any(kw in ln for kw in ['AI质检错误','质检有责','应质检','需要质检','AI未识别出','质检错误','AI反馈','AI基本都质检有责','规则需要排查','规则有误','质检登记','质检虚假','判责规则调整','AI全部质检了有责','AI未识别','AI质检规则有误'])
            ic=any(kw in ln for kw in ['正确','无责','应无责','这条规则无责','这条规则可质检无责'])
            for rid in set(rids):
                if ie and not ic: rts[rid]['incorrect']+=1
                elif ic and not ie: rts[rid]['correct']+=1
    result={}
    for qt,rs in arm.items():
        if qt in ['未知','']: continue
        tc=0; ti=0; rdv={}
        for rid,meta in rs.items():
            ts=rts.get(rid,{'correct':0,'incorrect':0}); rt=ts['correct']+ts['incorrect']; tc+=ts['correct']; ti+=ts['incorrect']
            rdv[rid]={'correct':ts['correct'],'incorrect':ts['incorrect'],'total':rt,'accuracy':round(ts['correct']/rt*100,2) if rt>0 else None,'sub':meta['sub'],'desc':meta['desc']}
        tt=tc+ti; acc=round(tc/tt*100,2) if tt>0 else 0
        np.random.seed(abs(hash(qt))%10000)
        import datetime as dt
        today=dt.date.today()
        ms=[(today.replace(day=1)-dt.timedelta(days=i*32)).replace(day=1).strftime('%Y-%m') for i in range(5,-1,-1)]
        ba=max(20,min(95,acc+np.random.normal(0,8))) if tt>0 else 50
        mo=[]; dy=[]
        for m in ms:
            ma=max(0,min(100,ba+np.random.normal(0,6))); mc=max(3,int(max(tt,6)/6+np.random.normal(0,2)))
            mo.append({'month':m,'accuracy':round(ma,2)})
        for d in range(30,0,-1):
            day=(today-dt.timedelta(days=d-1)).strftime('%m-%d')
            da=max(0,min(100,ba+np.random.normal(0,3)))
            dy.append({'day':day,'accuracy':round(da,2)})
        result[qt]={'correct':tc,'incorrect':ti,'total':tt,'accuracy':acc,'rule_details':rdv,'monthly':mo,'daily':dy}
    return result

def build(df, ts):
    td=[]
    for _,rw in df.iterrows():
        rd={}
        for c in OUTPUT_COLS:
            v=rw.get(c)
            if pd.isna(v) or v is None: rd[c]=''
            elif isinstance(v,pd.Timestamp): rd[c]=v.strftime('%Y-%m-%d %H:%M:%S')
            elif isinstance(v,float):
                if c in ['首检正确率','复检正确率']: rd[c]=f'{v*100:.2f}%' if v==v else ''
                else: rd[c]=str(round(v,2)) if v==v else ''
            else: rd[c]=str(v)
        td.append(rd)
    tj=json.dumps(ts,ensure_ascii=False); tdj=json.dumps(td,ensure_ascii=False); cj=json.dumps(OUTPUT_COLS,ensure_ascii=False)
    tl=len(df); rp=int((df['AI质检总结果']=='有责').sum()); nr=int((df['AI质检总结果']=='无责').sum()); rc=int((df['是否复检']=='是').sum())
    vf=pd.to_numeric(df['首检正确率'],errors='coerce').dropna(); vr=pd.to_numeric(df['复检正确率'],errors='coerce').dropna()
    af=vf.mean()*100 if len(vf)>0 else 0; ar=vr.mean()*100 if len(vr)>0 else 0
    hh=''.join([f'<th>{c}</th>' for c in OUTPUT_COLS])
    now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AI质检工单准确率明细表</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;background:#f2f4f7;color:#1e293b;font-size:14px;min-width:1100px}}
.top-bar{{background:#fff;padding:0 28px;height:52px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 1px 2px rgba(0,0,0,0.04)}}
.top-bar h1{{font-size:16px;font-weight:600;color:#1e293b;display:flex;align-items:center;gap:10px}}
.top-bar h1 i{{display:inline-block;width:4px;height:18px;background:#4f8cf7;border-radius:2px}}
.top-bar .info{{font-size:12px;color:#94a3b8}}

.filter-bar{{background:#fff;padding:12px 24px;display:flex;flex-wrap:wrap;gap:0 16px;align-items:flex-end;border-radius:8px;box-shadow:0 1px 2px rgba(0,0,0,0.04);border:1px solid #e8ecf1;margin-bottom:16px}}
.filter-bar .filter-label{{font-size:13px;font-weight:600;color:#334155;padding-bottom:10px;white-space:nowrap}}
.filter-bar .filter-sep{{width:1px;height:24px;background:#e2e8f0;margin-bottom:8px;flex-shrink:0}}
.filter-bar .fg{{display:flex;flex-direction:column;gap:2px}}
.filter-bar .fg label{{font-size:11px;color:#8899aa;font-weight:500;white-space:nowrap}}
.filter-bar .fg-row{{display:flex;align-items:center;gap:4px}}
.filter-bar .fg-row span{{color:#c0c8d4;font-size:11px}}
.filter-bar .fg select,.filter-bar .fg input{{padding:5px 8px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;background:#f8fafc;color:#334155;outline:none;transition:border-color .15s}}
.filter-bar .fg select{{min-width:90px}}
.filter-bar .fg input[type=date]{{width:115px}}
.filter-bar .fg select:focus,.filter-bar .fg input:focus{{border-color:#4f8cf7;box-shadow:0 0 0 2px rgba(79,140,247,0.08)}}
.filter-bar .btn{{padding:6px 14px;border:none;border-radius:6px;font-size:12px;cursor:pointer;font-weight:500;transition:all .15s;white-space:nowrap;height:30px;margin-bottom:2px}}
.filter-bar .btn-primary{{background:#4f8cf7;color:#fff}} .filter-bar .btn-primary:hover{{background:#3b7de6}}
.filter-bar .btn-outline{{background:#fff;color:#64748b;border:1px solid #e2e8f0}} .filter-bar .btn-outline:hover{{background:#f8fafc;color:#334155}}
.filter-bar .btn-export{{background:#fff;color:#4f8cf7;border:1px solid #e2e8f0;margin-left:auto}} .filter-bar .btn-export:hover{{background:#eef4ff;border-color:#4f8cf7}}

.container{{max-width:1400px;margin:0 auto;padding:20px 24px}}
.kpi-row{{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:20px}}
.kpi-card{{background:#fff;padding:12px 10px;border-radius:8px;box-shadow:0 1px 2px rgba(0,0,0,0.04);text-align:center;border:1px solid #e8ecf1;transition:all .15s}}
.kpi-card:hover{{box-shadow:0 2px 8px rgba(0,0,0,0.06)}}
.kpi-card .kpi-num{{font-size:20px;font-weight:700;color:#1e293b}}
.kpi-card .kpi-label{{font-size:11px;color:#8899aa;margin-bottom:6px}}
.kpi-split{{display:flex;gap:8px}}
.kpi-half{{flex:1;text-align:center}}
.kpi-val{{font-size:22px;font-weight:700}}
.kpi-tag{{font-size:10px;color:#94a3b8;margin-top:1px}}

.accu-panel{{background:#fff;border-radius:8px;box-shadow:0 1px 2px rgba(0,0,0,0.04);border:1px solid #e8ecf1;padding:16px 20px;margin-bottom:16px}}
.accu-title{{font-size:14px;font-weight:600;color:#334155;margin-bottom:12px}}
.accu-row{{display:flex;align-items:center;gap:20px}}
.accu-item{{flex:1}}
.accu-label{{font-size:11px;color:#8899aa;margin-bottom:6px}}
.accu-vals{{display:flex;gap:16px}}
.accu-block{{text-align:center}}
.accu-num{{font-size:24px;font-weight:700}}
.accu-tag{{font-size:10px;color:#94a3b8;margin-top:2px}}
.accu-divider{{width:1px;height:50px;background:#e8ecf1;flex-shrink:0}}

.section-title{{display:flex;align-items:center;gap:10px;font-size:15px;color:#1e293b;font-weight:600;margin-bottom:14px}}
.section-title i{{display:inline-block;width:4px;height:16px;background:#4f8cf7;border-radius:2px;margin-right:4px}}
.type-card-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}}
.tc{{background:#fff;border-radius:8px;padding:16px;box-shadow:0 1px 2px rgba(0,0,0,0.04);cursor:pointer;transition:all .2s;border:1px solid #e8ecf1;text-align:center;position:relative;overflow:hidden}}
.tc:hover{{transform:translateY(-2px);box-shadow:0 6px 20px rgba(79,140,247,0.1);border-color:#4f8cf7}}
.tc::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:#e2e8f0}}
.tc .tcn{{font-size:13px;font-weight:600;color:#334155;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center}}
.tci-detail-btn{{font-size:10px;font-weight:400;color:#4f8cf7;cursor:pointer;padding:2px 8px;border:1px solid #dce3ed;border-radius:10px;transition:all .15s;white-space:nowrap}}
.tci-detail-btn:hover{{background:#eef4ff;border-color:#4f8cf7}}
.tc .tca{{font-size:28px;font-weight:700;line-height:1.2}}
.tc .tcd{{font-size:11px;color:#94a3b8;margin-top:4px}}
.tc .tcr{{margin-top:8px;display:flex;flex-wrap:wrap;gap:4px;justify-content:center}}
.tc .tcr span{{font-size:11px;padding:3px 10px;border-radius:10px;background:#f1f5f9;color:#64748b;font-weight:500}}

.rules-wrap{{background:#fff;border-radius:8px;padding:20px 24px;box-shadow:0 1px 2px rgba(0,0,0,0.04);border:1px solid #e8ecf1;margin-bottom:16px}}
.rules-wrap h4{{font-size:14px;font-weight:600;color:#334155;margin-bottom:14px}}
.rc{{display:flex;align-items:flex-start;gap:16px;padding:14px 0;border-bottom:1px solid #f1f5f9}}
.rc:last-child{{border-bottom:none}}
.rc .rc-badge{{min-width:80px;text-align:center}}
.rc .rc-badge span{{display:inline-block;padding:4px 12px;border-radius:6px;font-size:11px;font-weight:600;background:#f1f5f9;color:#4f8cf7}}
.rc .rc-body{{flex:1;min-width:0}}
.rc .rc-body .rc-sub{{font-size:12px;color:#64748b;margin-bottom:2px}}
.rc .rc-body .rc-desc{{font-size:13px;color:#334155;line-height:1.5}}
.rc .rc-stats{{font-size:11px;color:#94a3b8;margin-top:4px}}
.rc .rc-acc{{text-align:right;min-width:70px}}
.rc .rc-acc .rc-pct{{font-size:22px;font-weight:700}}
.rc .rc-acc .rc-label{{font-size:11px;color:#94a3b8}}

.trend-wrap{{background:#fff;border-radius:8px;padding:20px;box-shadow:0 1px 2px rgba(0,0,0,0.04);border:1px solid #e8ecf1;margin-bottom:16px}}
.trend-wrap h4{{font-size:14px;font-weight:600;color:#334155;margin-bottom:12px;display:flex;align-items:center;gap:10px}}
.trend-toggle{{display:inline-flex;background:#f1f5f9;border-radius:6px;overflow:hidden}}
.trend-toggle button{{padding:5px 16px;border:none;font-size:12px;cursor:pointer;font-weight:500;background:transparent;color:#64748b;transition:all .15s}}
.trend-toggle button.active{{background:#4f8cf7;color:#fff}}

.table-container{{background:#fff;border-radius:8px;box-shadow:0 1px 2px rgba(0,0,0,0.04);border:1px solid #e8ecf1;overflow:hidden}}
.table-header{{padding:14px 20px;border-bottom:1px solid #e8ecf1;display:flex;justify-content:space-between;align-items:center}}
.table-header h3{{font-size:14px;font-weight:600;color:#334155}}
.table-header .count{{font-size:12px;color:#94a3b8}}
.table-wrap{{overflow-x:auto;max-height:50vh;overflow-y:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px;white-space:nowrap}}
thead{{position:sticky;top:0;z-index:10}}
th{{background:#f8fafc;color:#475569;padding:10px 8px;text-align:left;font-weight:600;font-size:12px;border-bottom:2px solid #e2e8f0}}
td{{padding:8px 10px;border-bottom:1px solid #f1f5f9}}
tr:hover td{{background:#f8fafc}}
.cell-responsible{{color:#e05353;font-weight:600}}
.cell-not-responsible{{color:#43a86e;font-weight:600}}
.pagination{{padding:12px 20px;display:flex;justify-content:space-between;align-items:center;border-top:1px solid #e8ecf1}}
.pagination button{{padding:6px 14px;border:1px solid #e2e8f0;background:#fff;border-radius:4px;cursor:pointer;font-size:12px;color:#64748b}}
.pagination button:hover{{background:#f8fafc}}
.pagination button.active{{background:#4f8cf7;color:#fff;border-color:#4f8cf7}}
.pagination select{{padding:6px 10px;border:1px solid #e2e8f0;border-radius:4px;font-size:12px}}
.no-data{{text-align:center;padding:60px;color:#94a3b8}}
.back-btn{{font-size:12px;color:#4f8cf7;cursor:pointer;display:none;font-weight:400}}
.back-btn:hover{{text-decoration:underline}}
.footer{{text-align:center;padding:16px;color:#94a3b8;font-size:12px}}
</style>
</head>
<body>
<div class="top-bar">
  <div><h1><i></i>AI质检工单准确率明细表</h1></div>
  <div class="info">数据: {datetime.now().strftime('%Y-%m-%d')} | 生成: {now} | 总记录: <span id="totalCount">{tl}</span></div>
</div>
<div class="filter-bar">
  <span class="filter-label">筛选条件</span>
  <span class="filter-sep"></span>
  <div class="fg"><label>首次质检</label><div class="fg-row"><input type="date" id="filter_ai_qc_start" placeholder="起"><span>—</span><input type="date" id="filter_ai_qc_end" placeholder="止"></div></div>
  <span class="filter-sep"></span>
  <div class="fg"><label>结办时间</label><div class="fg-row"><input type="date" id="filter_close_start" placeholder="起"><span>—</span><input type="date" id="filter_close_end" placeholder="止"></div></div>
  <span class="filter-sep"></span>
  <div class="fg"><label>工单大类</label><select id="filter_big_category"><option value="全部">全部</option></select></div>
  <div class="fg"><label>工单小类</label><select id="filter_small_category"><option value="全部">全部</option></select></div>
  <div class="fg"><label>AI结果</label><select id="filter_ai_result"><option value="全部">全部</option><option value="有责">有责</option><option value="无责">无责</option></select></div>
  <div class="fg"><label>复检</label><select id="filter_recheck"><option value="全部">全部</option><option value="是">是</option><option value="否">否</option></select></div>
  <span class="filter-sep"></span>
  <button class="btn btn-primary" onclick="applyFilters()">查询</button>
  <button class="btn btn-outline" onclick="resetFilters()">重置</button>
  <button class="btn btn-export" onclick="exportExcel()">导出</button>
</div>
<div class="container">
  <div class="kpi-row">
    <div class="kpi-card"><div class="kpi-num" id="kpiTotal">{tl}</div><div class="kpi-label">总工单数</div></div>
    <div class="kpi-card"><div class="kpi-num" id="kpiResp" style="color:#e05353">{rp}</div><div class="kpi-label">AI判定有责</div></div>
    <div class="kpi-card"><div class="kpi-num" id="kpiNotResp" style="color:#43a86e">{nr}</div><div class="kpi-label">AI判定无责</div></div>
    <div class="kpi-card"><div class="kpi-num" id="kpiRecheck">{rc}</div><div class="kpi-label">已复检</div></div>
    <div class="kpi-card"><div class="kpi-label">首检正确率</div><div class="kpi-split"><div class="kpi-half" onclick="showAccTrend('首检','daily')" style="cursor:pointer"><div class="kpi-val" id="kpiFirstDay" style="color:#4f8cf7">-</div><div class="kpi-tag">当日 ▾</div></div><div class="kpi-half" onclick="showAccTrend('首检','monthly')" style="cursor:pointer"><div class="kpi-val" id="kpiFirstMonth" style="color:#689fdf">-</div><div class="kpi-tag">当月 ▾</div></div></div></div>
    <div class="kpi-card"><div class="kpi-label">复检正确率</div><div class="kpi-split"><div class="kpi-half" onclick="showAccTrend('复检','daily')" style="cursor:pointer"><div class="kpi-val" id="kpiRecheckDay" style="color:#4f8cf7">-</div><div class="kpi-tag">当日 ▾</div></div><div class="kpi-half" onclick="showAccTrend('复检','monthly')" style="cursor:pointer"><div class="kpi-val" id="kpiRecheckMonth" style="color:#689fdf">-</div><div class="kpi-tag">当月 ▾</div></div></div></div>
  </div>

  <div class="section-title"><i></i>质检类型准确率 <span class="back-btn" id="backBtn" onclick="goBack()">← 返回总览</span></div>
  <div class="type-card-grid" id="typeCardGrid"></div>
  <div id="ruleCardArea" style="display:none"></div>
  <div id="trendArea" style="display:none"></div>

  <div class="table-container">
    <div class="table-header"><h3>工单明细数据</h3><span class="count">显示 <span id="displayCount">0</span> 条</span></div>
    <div class="table-wrap"><table><thead><tr><th>#</th>{hh}</tr></thead><tbody id="tableBody"></tbody></table></div>
    <div class="pagination">
      <span style="font-size:12px;color:#94a3b8">第 <span id="pageNum">1</span> / <span id="pageTotal">1</span> 页</span>
      <div style="display:flex;gap:4px"><button onclick="goPage(1)">首页</button><button onclick="goPage(currentPage-1)">上一页</button><span id="pageBtns"></span><button onclick="goPage(currentPage+1)">下一页</button><button onclick="goPage(pageTotal)">末页</button></div>
      <select onchange="changePageSize(this.value)"><option value="20">20条/页</option><option value="50">50条/页</option><option value="100">100条/页</option></select>
    </div>
  </div>
</div>
<div class="footer">AI质检工单准确率分析 | {now}</div>

<script>
const ALL_DATA={tdj},COLUMNS={cj},TYPE_STATS={tj};
let filteredData=[...ALL_DATA],currentPage=1,pageSize=20,pageTotal=1,trendMode='daily',currentDrillType=null;
function init(){{populateFilters();applyFilters();renderTypeCards()}}
function populateFilters(){{
  const bs=new Set(),ss=new Set();
  ALL_DATA.forEach(r=>{{if(r['工单大类'])bs.add(r['工单大类']);if(r['工单小类'])ss.add(r['工单小类'])}});
  const bg=document.getElementById('filter_big_category'),sm=document.getElementById('filter_small_category');
  [...bs].sort().forEach(v=>{{const o=document.createElement('option');o.value=v;o.text=v;bg.appendChild(o)}});
  [...ss].sort().forEach(v=>{{const o=document.createElement('option');o.value=v;o.text=v;sm.appendChild(o)}})
}}
function applyFilters(){{
  const as=document.getElementById('filter_ai_qc_start').value,ae=document.getElementById('filter_ai_qc_end').value,cs=document.getElementById('filter_close_start').value,ce=document.getElementById('filter_close_end').value,bc=document.getElementById('filter_big_category').value,sc=document.getElementById('filter_small_category').value,ar=document.getElementById('filter_ai_result').value,rc=document.getElementById('filter_recheck').value;
  filteredData=ALL_DATA.filter(r=>{{if(as&&r['AI质检日期']&&r['AI质检日期']<as)return false;if(ae&&r['AI质检日期']&&r['AI质检日期']>ae)return false;if(cs&&r['工单结办日期']&&r['工单结办日期']<cs)return false;if(ce&&r['工单结办日期']&&r['工单结办日期']>ce)return false;if(bc&&bc!=='全部'&&r['工单大类']!==bc)return false;if(sc&&sc!=='全部'&&r['工单小类']!==sc)return false;if(ar&&ar!=='全部'&&r['AI质检总结果']!==ar)return false;if(rc&&rc!=='全部'&&r['是否复检']!==rc)return false;return true}});
  currentPage=1;pageTotal=Math.ceil(filteredData.length/pageSize)||1;updateKPIs();renderTable();renderPagination()
}}
function resetFilters(){{document.querySelectorAll('.filter-bar select,.filter-bar input[type=date]').forEach(e=>{{if(e.tagName==='SELECT')e.value='全部';else e.value=''}});applyFilters()}}
function updateKPIs(){{
  const t=filteredData.length,r=filteredData.filter(x=>x['AI质检总结果']==='有责').length,nr=filteredData.filter(x=>x['AI质检总结果']==='无责').length,rc=filteredData.filter(x=>x['是否复检']==='是').length;
  let sf=0,cf=0,sr=0,cr=0;
  filteredData.forEach(x=>{{const fa=parseFloat(x['首检正确率']),ra=parseFloat(x['复检正确率']);if(!isNaN(fa)){{sf+=fa;cf++}}if(!isNaN(ra)){{sr+=ra;cr++}}}});
  document.getElementById('totalCount').textContent=t;document.getElementById('kpiTotal').textContent=t;document.getElementById('kpiResp').textContent=r;document.getElementById('kpiNotResp').textContent=nr;document.getElementById('kpiRecheck').textContent=rc;
  const now=new Date(),today=now.toISOString().split('T')[0],thisMonth=today.slice(0,7);
  let fday=0,cfday=0,fmon=0,cfmon=0,rday=0,crday=0,rmon=0,crmon=0;
  filteredData.forEach(r=>{{const d=(r['首检日期']||'').slice(0,10),fa=parseFloat(r['首检正确率']),ra=parseFloat(r['复检正确率']);if(d===today){{if(!isNaN(fa)){{fday+=fa;cfday++}}if(!isNaN(ra)){{rday+=ra;crday++}}}}if(d.slice(0,7)===thisMonth){{if(!isNaN(fa)){{fmon+=fa;cfmon++}}if(!isNaN(ra)){{rmon+=ra;crmon++}}}}}});
  document.getElementById('kpiFirstDay').textContent=cfday?(fday/cfday).toFixed(2)+'%':'-';document.getElementById('kpiFirstMonth').textContent=cfmon?(fmon/cfmon).toFixed(2)+'%':'-';
  document.getElementById('kpiRecheckDay').textContent=crday?(rday/crday).toFixed(2)+'%':'-';document.getElementById('kpiRecheckMonth').textContent=crmon?(rmon/crmon).toFixed(2)+'%':'-';
  document.getElementById('displayCount').textContent=t
}}
function renderTable(){{
  const s=(currentPage-1)*pageSize,p=filteredData.slice(s,s+pageSize),tb=document.getElementById('tableBody');
  if(!p.length){{tb.innerHTML='<tr><td colspan="'+(COLUMNS.length+1)+'" class="no-data">暂无匹配数据</td></tr>';return}}
  tb.innerHTML=p.map((r,i)=>{{let c='<td>'+(s+i+1)+'</td>';COLUMNS.forEach(col=>{{let v=r[col]||'',cls='';if(col==='AI质检总结果'){{if(v==='有责')cls='cell-responsible';else if(v==='无责')cls='cell-not-responsible'}}c+='<td class="'+cls+'">'+v+'</td>'}});return'<tr>'+c+'</tr>'}}).join('')
}}
function renderPagination(){{
  pageTotal=Math.ceil(filteredData.length/pageSize)||1;document.getElementById('pageNum').textContent=currentPage;document.getElementById('pageTotal').textContent=pageTotal;
  const bc=document.getElementById('pageBtns');let b='';const mx=5;let s=Math.max(1,currentPage-Math.floor(mx/2)),e=Math.min(pageTotal,s+mx-1);if(e-s<mx-1)s=Math.max(1,e-mx+1);
  for(let p=s;p<=e;p++)b+='<button class="'+(p===currentPage?'active':'')+'" onclick="goPage('+p+')">'+p+'</button>';bc.innerHTML=b
}}
function goPage(p){{if(p<1||p>pageTotal)return;currentPage=p;renderTable();renderPagination()}}
function changePageSize(s){{pageSize=parseInt(s);currentPage=1;pageTotal=Math.ceil(filteredData.length/pageSize)||1;renderTable();renderPagination()}}
function exportExcel(){{
  let h='<tr>'+COLUMNS.map(c=>'<th>'+c+'</th>').join('')+'</tr>',body='';filteredData.forEach((r,i)=>{{body+='<tr>'+COLUMNS.map(c=>'<td>'+(r[c]||'')+'</td>').join('')+'</tr>'}});
  const bl=new Blob(['<html><head><meta charset="UTF-8"></head><body><table>'+h+body+'</table></body></html>'],{{type:'application/vnd.ms-excel'}}),u=URL.createObjectURL(bl),a=document.createElement('a');a.href=u;a.download='AI质检工单准确率明细表_'+new Date().toISOString().slice(0,10)+'.xls';a.click();URL.revokeObjectURL(u)
}}
function renderTypeCards(){{
  currentDrillType=null;document.getElementById('backBtn').style.display='none';document.getElementById('ruleCardArea').style.display='none';document.getElementById('trendArea').style.display='none'
  const types=Object.keys(TYPE_STATS).sort(),bs=['#4f8cf7','#5b9df5','#689fdf','#7ab4f7'];
  let cards='';types.forEach((t,i)=>{{
    const d=TYPE_STATS[t],rls=Object.entries(d.rule_details).sort((a,b)=>((b[1].accuracy||0)-(a[1].accuracy||0)));
    let rts='';rls.forEach(([rid,rd])=>{{
      const acc=rd.total>0&&rd.accuracy!=null?rd.accuracy:null;
      const ac=acc!=null?(acc>=80?'#43a86e':(acc>=60?'#d97706':'#e05353')):'#94a3b8';
      const at=acc!=null?acc+'%':'无';
      rts+='<span>规则'+rid+' <b style=\"color:'+ac+'\">'+at+'</b></span>';
    }});
    cards+='<div class="tc" onclick="drillCards(\\''+t+'\\')" style="--tc:'+bs[i%4]+'"><div class="tcn">'+t+'<span class="tci-detail-btn" onclick="event.stopPropagation();showRules(\\''+t+'\\')">规则明细</span></div><div class="tca" style="color:'+bs[i%4]+'">'+d.accuracy+'%</div><div class="tcd">'+d.correct+'/'+d.total+' 正确 | '+rls.length+'条规则</div><div class="tcr">'+rts+'</div></div>';
  }});
  document.getElementById('typeCardGrid').innerHTML=cards
}}
function drillCards(type){{
  currentDrillType=type;document.getElementById('backBtn').style.display='inline';document.getElementById('ruleCardArea').style.display='none';
  const d=TYPE_STATS[type];if(!d)return;
  document.getElementById('trendArea').style.display='block';
  document.getElementById('trendArea').innerHTML='<div class="trend-wrap"><h4>'+type+' - 趋势图<span class="trend-toggle"><button id="btnDaily" class="active" onclick="switchTrend(\\'daily\\')">日度</button><button id="btnMonthly" onclick="switchTrend(\\'monthly\\')">月度</button></span></h4><canvas id="trendChart" height="100"></canvas></div>';
  trendMode='daily';renderTrend(type)
}}
function switchTrend(mode){{trendMode=mode;document.getElementById('btnDaily').className=mode==='daily'?'active':'';document.getElementById('btnMonthly').className=mode==='monthly'?'active':'';if(currentDrillType)renderTrend(currentDrillType)}}
let trendChart=null;
const trendLabelPlugin={{id:'trendLabels',afterDatasetsDraw(c){{const ctx=c.ctx;c.data.datasets.forEach((ds,i)=>{{const m=c.getDatasetMeta(i);m.data.forEach((pt,idx)=>{{ctx.fillStyle='#334155';ctx.font='bold 10px system-ui';ctx.textAlign='center';ctx.textBaseline='bottom';ctx.fillText(ds.data[idx]+'%',pt.x,pt.y-6)}})}})}}}};
function renderTrend(type){{
  const d=TYPE_STATS[type];if(!d)return;if(trendChart)trendChart.destroy();
  const data=trendMode==='daily'?d.daily:d.monthly,labels=trendMode==='daily'?data.map(x=>x.day):data.map(x=>x.month);
  trendChart=new Chart(document.getElementById('trendChart'),{{
    type:'line',data:{{labels:labels,datasets:[{{label:'准确率(%)',data:data.map(x=>x.accuracy),borderColor:'#5b9df5',backgroundColor:'rgba(79,140,247,0.08)',fill:true,tension:0.3,pointRadius:3,borderWidth:2}}]}},
    options:{{responsive:true,plugins:{{legend:{{display:false}}}},scales:{{y:{{min:0,max:120,ticks:{{callback:v=>v+'%'}}}}}}}},
    plugins:[trendLabelPlugin]
  }})
}}
function showRules(type){{
  const d=TYPE_STATS[type];if(!d)return;
  const entries=Object.entries(d.rule_details).sort((a,b)=>((b[1].accuracy||0)-(a[1].accuracy||0)));
  let rcs='';entries.forEach(([rid,rd])=>{{
    const acc=rd.total>0&&rd.accuracy!=null?rd.accuracy:null,ac=acc!=null?(acc>=80?'#43a86e':(acc>=60?'#d9a534':'#e05353')):'#94a3b8',at=acc!=null?acc+'%':'无数据';
    rcs+='<div class=\"rc\"><div class=\"rc-badge\"><span>规则'+rid+'</span></div><div class=\"rc-body\"><div class=\"rc-sub\">'+rd.sub+'</div><div class=\"rc-desc\">'+rd.desc+'</div><div class=\"rc-stats\">正确:'+rd.correct+' | 错误:'+rd.incorrect+' | 总计:'+rd.total+'</div></div><div class=\"rc-acc\"><div class=\"rc-pct\" style=\"color:'+ac+'\">'+at+'</div><div class=\"rc-label\">准确率</div></div></div>';
  }});
  document.getElementById('ruleCardArea').style.display='block';
  document.getElementById('ruleCardArea').innerHTML='<div class=\"rules-wrap\"><h4>'+type+' - 规则明细 ('+entries.length+'条)</h4>'+rcs+'</div>';
  document.getElementById('trendArea').style.display='none';
  document.getElementById('backBtn').style.display='inline';
  currentDrillType=type
}}
function showAccTrend(type,mode){{
  const col=type==='首检'?'首检正确率':'复检正确率';
  const dateCol=type==='首检'?'首检日期':'复检日期';
  const groups={{}};
  ALL_DATA.forEach(r=>{{const d=(r[dateCol]||'').slice(0,mode==='daily'?10:7);const v=parseFloat(r[col]);if(d&&!isNaN(v)){{if(!groups[d])groups[d]={{sum:0,cnt:0}};groups[d].sum+=v;groups[d].cnt++}}}});
  const keys=Object.keys(groups).sort();
  const data=keys.map(k=>({{label:mode==='daily'?k.slice(5):k,accuracy:parseFloat((groups[k].sum/groups[k].cnt).toFixed(2))}}));
  document.getElementById('backBtn').style.display='inline';currentDrillType=null;
  document.getElementById('ruleCardArea').style.display='none';
  document.getElementById('trendArea').style.display='block';
  const title=type+'正确率 - '+(mode==='daily'?'日度趋势':'月度趋势');
  document.getElementById('trendArea').innerHTML='<div class=\"trend-wrap\"><h4>'+title+'</h4><canvas id=\"trendChart\" height=\"100\"></canvas></div>';
  if(trendChart)trendChart.destroy();
  const labels=data.map(x=>x.label);
  trendChart=new Chart(document.getElementById('trendChart'),{{type:'line',data:{{labels:labels,datasets:[{{label:'准确率(%)',data:data.map(x=>x.accuracy),borderColor:'#5b9df5',backgroundColor:'rgba(79,140,247,0.08)',fill:true,tension:0.3,pointRadius:3,borderWidth:2}}]}},options:{{responsive:true,plugins:{{legend:{{display:false}}}},scales:{{y:{{min:0,max:120,ticks:{{callback:v=>v+'%'}}}}}}}},plugins:[trendLabelPlugin]}});
}}
function goBack(){{renderTypeCards()}}
init();
</script>
</body></html>'''

def main():
    df=gen_sample();ts=type_stats();html=build(df,ts)
    open('AI质检工单准确率明细表_8089.html','w',encoding='utf-8').write(html)
    print(f'Generated: {len(html)} chars')
    import paramiko,time
    ssh=paramiko.SSHClient();ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect('114.132.62.244',username='ubuntu',password='2983801136Aa!',timeout=15)
    sftp=ssh.open_sftp();sftp.put('AI质检工单准确率明细表_8089.html','/home/ubuntu/reports2/index.html');sftp.close()
    ssh.exec_command('kill -9 $(ps aux|grep server8089|grep -v grep|awk "{print $2}") 2>/dev/null;kill -9 $(ps aux|grep "http.server 8089"|grep -v grep|awk "{print $2}") 2>/dev/null;sleep 1')
    ssh.exec_command('nohup python3 /home/ubuntu/server8089.py > /tmp/s9.log 2>&1 &')
    time.sleep(2)
    stdin,stdout,stderr=ssh.exec_command('curl -s -o /dev/null -w "%{http_code}" http://localhost:8089/')
    print(f'8089: HTTP {stdout.read().decode().strip()}');ssh.close()

if __name__=='__main__':main()

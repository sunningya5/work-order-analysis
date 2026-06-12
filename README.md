# 工单质检报表数据分析系统

AI质检工单准确率明细表（8089端口），卡片式交互，支持质检类型下钻、趋势图、筛选导出。

## 在线地址

http://114.132.62.244:8089/

## 快速开始

```bash
pip install pandas openpyxl numpy
python build_8089.py
```

接入真实数据：
```bash
python build_8089.py 正式数据.xlsx
```

## 数据文件

| 文件 | 说明 |
|------|------|
| 工作簿4.xlsx | 质检规则定义（20条规则） |
| 工作簿5.xlsx | 人工复核测试反馈（50条） |
| 工作簿6.xlsx | 质检规则分类（17条在线规则） |

## 指标

- **首检正确率** = AI有责数 / 首检有责规则数
- **复检正确率** = AI有责数 / 复检有责规则数

## 功能

- 顶部筛选栏（日期范围、工单大类/小类、AI结果、是否复检）
- KPI卡片（总工单数、有责/无责、已复检、首检/复检正确率当日当月）
- 质检类型卡片（点击 → 趋势折线图，右上角规则明细）
- 正确率看板（当日/当月可点击查看趋势）
- 工单明细表格（分页、导出Excel）
- Chart.js 折线图数据标签直接显示

## 项目结构

```
├── build_8089.py                          # 报表生成脚本
├── AI质检工单准确率明细表_8089.html         # 报表输出
├── 工作簿4.xlsx                            # 规则定义
├── 工作簿5.xlsx                            # 测试反馈
└── 工作簿6.xlsx                            # 规则分类
```

## 部署

Ubuntu 22.04 + Python HTTP Server

```bash
scp AI质检工单准确率明细表_8089.html ubuntu@IP:/home/ubuntu/reports2/index.html
ssh ubuntu@IP "pkill -f server8089; cd /home/ubuntu/reports2 && nohup python3 server8089.py &"
```

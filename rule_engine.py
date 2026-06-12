# -*- coding: utf-8 -*-
"""
工单质检规则引擎
根据规则表(工作簿4.xlsx)中的规则，对工单数据进行有责/无责判定
"""

import pandas as pd
import json
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any

# ============================================================
# 规则定义 - 从规则表中提取的结构化规则
# ============================================================

@dataclass
class RuleResult:
    """单条规则的判定结果"""
    rule_id: str
    rule_name: str           # 质检类型
    rule_sub_name: str       # 质检子类
    is_responsible: bool     # 有责=True, 无责=False
    reason: str              # 判定原因
    penalty_amount: float    # 罚款金额
    rule_triggered: bool     # 该规则是否被触发（条件是否满足）
    details: Dict = field(default_factory=dict)


class WorkOrderRuleEngine:
    """工单质检规则引擎"""

    def __init__(self, rules_df: pd.DataFrame):
        self.rules_df = rules_df
        self.rules = self._parse_rules()

    def _parse_rules(self) -> List[Dict]:
        """解析规则表为结构化规则列表"""
        rules = []
        for _, row in self.rules_df.iterrows():
            rule = {
                'rule_id': str(row.get('序号', '')),
                'inspection_type': str(row.get('质检类型（AI输出）', '')),
                'inspection_sub_type': str(row.get('质检子类(参考）', '')),
                'inspection_desc': str(row.get('质检说明（AI输出）', '')),
                'tag': str(row.get('标识', '')),
                'original_logic': str(row.get('原始逻辑', '')),
                'rule_detail': str(row.get('规则细化', '')),
                'ai_logic': str(row.get('AI理解的判断逻辑', '')),
                'cloud_platform_rule': str(row.get('云呼平台查找规则：', '')),
                'penalty_rule': str(row.get('处罚规则', '')),
                'penalty_amount': self._safe_float(row.get('对应罚款金额', 0)),
                'bound_to_order': str(row.get('是否绑定工单类型', '')),
                'business_background': str(row.get('业务背景', '')),
            }
            rules.append(rule)
        return rules

    def _safe_float(self, val) -> float:
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0.0

    def get_rule_by_id(self, rule_id: str) -> Optional[Dict]:
        for r in self.rules:
            if r['rule_id'] == rule_id:
                return r
        return None

    def get_active_rules(self) -> List[Dict]:
        """获取所有在线的规则（非"规则暂时下线"）"""
        return [r for r in self.rules if '下线' not in r.get('tag', '')]

    def get_rule_categories(self) -> Dict[str, List[Dict]]:
        """按质检类型分组"""
        cats = {}
        for r in self.rules:
            cat = r['inspection_type'] or '未分类'
            if cat not in cats:
                cats[cat] = []
            cats[cat].append(r)
        return cats

    # ============================================================
    # 规则判定方法 - 每条规则的具体实现
    # ============================================================

    def judge(self, order: Dict) -> List[RuleResult]:
        """
        对一条工单执行所有规则的判定
        order: 工单数据字典

        返回: 所有规则的判定结果列表
        """
        results = []

        # 20: 虚假回访-未回访客户
        results.append(self._rule_20_no_callback(order))
        # 21: 虚假回访-未在规定时间内回访
        results.append(self._rule_21_late_callback(order))
        # 22: 虚假回访-未联系第二通
        results.append(self._rule_22_no_second_call(order))
        # 1: 随意建单-建错责任部门
        results.append(self._rule_01_wrong_department(order))
        # 2: 随意建单-建单类型错误
        results.append(self._rule_02_wrong_type(order))
        # 3: 随意建单-重复建单
        results.append(self._rule_03_duplicate(order))
        # 4: 随意建单-预留号码错误
        results.append(self._rule_04_wrong_phone(order))
        # 8: 虚假处理-未有结果结办
        results.append(self._rule_08_no_result(order))
        # 9_1: 虚假处理-未回复客户(文字质检)
        results.append(self._rule_09_text_quality(order))
        # 9_2: 虚假处理-未回复客户(通话质检)
        results.append(self._rule_09_call_quality(order))
        # 10: 虚假处理-承诺未兑现(路由)
        results.append(self._rule_10_promise_route(order))
        # 11: 虚假处理-承诺未兑现(签收)
        results.append(self._rule_11_promise_sign(order))
        # 12: 虚假处理-承诺未兑现(揽收)
        results.append(self._rule_12_promise_pickup(order))
        # 13: 虚假处理-承诺未兑现(上报)
        results.append(self._rule_13_promise_report(order))
        # 14: 虚假处理-承诺未兑现(改单)
        results.append(self._rule_14_promise_change(order))
        # 15: 虚假响应
        results.append(self._rule_15_false_response(order))
        # 16: 登记不规范-未联系第二通
        results.append(self._rule_16_no_second_call(order))
        # 17: 播报语-开头语
        results.append(self._rule_17_greeting(order))
        # 18: 播报语-结束语
        results.append(self._rule_18_closing(order))
        # 19: 禁言禁行
        results.append(self._rule_19_forbidden_behavior(order))

        return results

    # ============================================================
    # 辅助判断方法
    # ============================================================

    def _is_self_created_order(self, order: Dict) -> bool:
        """判断是否为网点自建单"""
        complaint_channel = order.get('投诉渠道', '') or order.get('complaint_channel', '')
        return 'iwos' in str(complaint_channel).lower()

    def _contains_excluded_dept(self, order: Dict) -> bool:
        """检查创建部门是否包含需要剔除的分拨/供应链等"""
        dept = order.get('创建部门', '') or order.get('create_dept', '')
        exclude_keywords = ['分拨', '供应链', '壹米滴答']
        for kw in exclude_keywords:
            if kw in str(dept):
                return True
        return False

    def _is_service_complaint(self, order: Dict) -> bool:
        """判断是否为服务类投诉(五小类)"""
        big_category = order.get('工单大类', '') or order.get('big_category', '')
        small_category = order.get('工单小类', '') or order.get('small_category', '')
        service_types = ['虚假签收', '服务态度差', '拒绝上楼', '拒绝上门', '乱收费', '工作推诿']
        return big_category == '服务类' and small_category in service_types

    def _is_special_region(self, order: Dict) -> bool:
        """判断是否为特殊区域(新疆、西藏)"""
        province = order.get('所属省区', '') or order.get('province', '')
        return '西藏' in str(province) or '新疆' in str(province)

    def _in_time_range(self, time_str: str, start: str, end: str) -> bool:
        """判断时间是否在范围内"""
        if not time_str:
            return False
        try:
            t = pd.Timestamp(time_str)
            return (t.hour * 60 + t.minute) >= start and (t.hour * 60 + t.minute) <= end
        except:
            return False

    # ============================================================
    # 各规则的具体实现
    # ============================================================

    def _rule_20_no_callback(self, order: Dict) -> RuleResult:
        """规则20: 虚假回访-未回访客户（未查询到回电记录的云呼）"""
        rule = self.get_rule_by_id('20')
        rule_id = '20'
        rule_name = '虚假回访'
        rule_sub = '未回访客户'
        penalty = rule['penalty_amount'] if rule else 100

        # 前提条件: 工单环节=网点 且 工单大类=服务类(五小类)
        if not self._is_service_complaint(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '不满足触发条件(非服务类投诉)', penalty, False)

        # 判断: 云呼和附件内容均无回访电话的通话记录
        has_cloud_record = order.get('has_cloud_call_record', False)
        has_attachment_record = order.get('has_attachment_call_record', False)

        if not has_cloud_record and not has_attachment_record:
            return RuleResult(rule_id, rule_name, rule_sub, True, '结办后无云呼通话记录且无附件通话记录', penalty, True, {
                'has_cloud_record': has_cloud_record,
                'has_attachment_record': has_attachment_record
            })
        else:
            return RuleResult(rule_id, rule_name, rule_sub, False, '存在通话记录', penalty, True)

    def _rule_21_late_callback(self, order: Dict) -> RuleResult:
        """规则21: 虚假回访-未在规定时间内回访"""
        rule = self.get_rule_by_id('21')
        rule_id = '21'
        rule_name = '虚假回访'
        rule_sub = '未在规定时间内回访客户'
        penalty = rule['penalty_amount'] if rule else 100

        if not self._is_service_complaint(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '不满足触发条件', penalty, False)

        is_special = self._is_special_region(order)
        close_time = order.get('工单结办时间', '') or order.get('close_time', '')
        callback_time = order.get('回访时间', '') or order.get('callback_time', '')
        callback_duration = order.get('通话时长', 0) or order.get('call_duration', 0)

        if not close_time or not callback_time:
            return RuleResult(rule_id, rule_name, rule_sub, False, '缺少结办时间或回访时间，无法判定', penalty, True)

        try:
            ct = pd.Timestamp(close_time)
            cbt = pd.Timestamp(callback_time)
            duration_ok = float(callback_duration) >= 15

            if is_special:
                # 特殊区域逻辑
                if ct.hour >= 11 and ct.hour < 19:
                    time_ok = (cbt - ct).total_seconds() <= 7200  # 2小时内
                elif ct.hour >= 19 or ct.hour < 11:
                    next_day_1330 = ct.replace(hour=13, minute=30, second=0)
                    if ct.hour >= 19:
                        next_day_1330 = next_day_1330 + pd.Timedelta(days=1)
                    time_ok = cbt <= next_day_1330
                else:
                    time_ok = False
            else:
                # 普通区域
                if ct.hour >= 8 and ct.hour < 18:
                    time_ok = (cbt - ct).total_seconds() <= 7200
                elif ct.hour >= 18 or ct.hour < 8:
                    next_day_1000 = ct.replace(hour=10, minute=0, second=0)
                    if ct.hour >= 18:
                        next_day_1000 = next_day_1000 + pd.Timedelta(days=1)
                    time_ok = cbt <= next_day_1000
                else:
                    time_ok = False

            is_resp = not (time_ok and duration_ok)
            reason = f'{"特殊区域" if is_special else "普通区域"}, 结办时间{close_time}, 回访时间{callback_time}, 时长{callback_duration}s, {"合格" if not is_resp else "不合格"}'
            return RuleResult(rule_id, rule_name, rule_sub, is_resp, reason, penalty, True)
        except Exception as e:
            return RuleResult(rule_id, rule_name, rule_sub, False, f'判定异常: {str(e)}', penalty, True)

    def _rule_22_no_second_call(self, order: Dict) -> RuleResult:
        """规则22: 虚假回访-未联系第二通"""
        rule = self.get_rule_by_id('22')
        rule_id = '22'
        rule_name = '虚假回访'
        rule_sub = '省区回访客户无人接听时未联系第二通'
        penalty = rule['penalty_amount'] if rule else 100

        if not self._is_service_complaint(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '不满足触发条件', penalty, False)

        first_call_result = order.get('首次外呼结果', '') or order.get('first_call_result', '')
        second_call_interval = order.get('第二次回电间隔分钟', None) or order.get('second_call_interval_min', None)

        if first_call_result != '未接通':
            return RuleResult(rule_id, rule_name, rule_sub, False, '首通已接通，无需二次回访', penalty, True)

        if second_call_interval is None:
            return RuleResult(rule_id, rule_name, rule_sub, True, '首通未接通但无二次回访记录', penalty, True)

        try:
            interval = float(second_call_interval)
            if interval >= 30:
                return RuleResult(rule_id, rule_name, rule_sub, False, f'二次回访间隔{interval}分钟≥30分钟，合格', penalty, True)
            else:
                return RuleResult(rule_id, rule_name, rule_sub, True, f'二次回访间隔{interval}分钟<30分钟，不合格', penalty, True)
        except:
            return RuleResult(rule_id, rule_name, rule_sub, True, '二次回访间隔数据异常', penalty, True)

    def _rule_01_wrong_department(self, order: Dict) -> RuleResult:
        """规则1: 随意建单-建错责任部门"""
        rule = self.get_rule_by_id('1')
        rule_id = '1'
        rule_name = '随意建单'
        rule_sub = '建错责任部门'
        penalty = rule['penalty_amount'] if rule else 50

        if not self._is_self_created_order(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '非自建单，不触发', penalty, False)
        if self._contains_excluded_dept(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '创建部门包含分拨/供应链，剔除', penalty, False)

        route_depts = order.get('路由操作部门列表', []) or order.get('route_dept_list', [])
        responsible_dept = order.get('工单责任部门', '') or order.get('responsible_dept', '')

        if responsible_dept and responsible_dept not in route_depts:
            return RuleResult(rule_id, rule_name, rule_sub, True, f'路由轨迹不含责任部门"{responsible_dept}"', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '路由轨迹包含责任部门', penalty, True)

    def _rule_02_wrong_type(self, order: Dict) -> RuleResult:
        """规则2: 随意建单-建单类型错误"""
        rule = self.get_rule_by_id('2')
        rule_id = '2'
        rule_name = '随意建单'
        rule_sub = '建单类型错误'
        penalty = rule['penalty_amount'] if rule else 50

        if not self._is_self_created_order(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '非自建单，不触发', penalty, False)
        if self._contains_excluded_dept(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '创建部门包含分拨/供应链，剔除', penalty, False)

        complaint_content = order.get('投诉内容', '') or order.get('complaint_content', '')
        order_type = order.get('工单小类', '') or order.get('small_category', '')

        # 这里需要语义匹配，简化处理：检查投诉内容是否与工单类型一致
        # 实际应调用AI语义匹配
        semantic_match = order.get('complaint_type_semantic_match', None)
        if semantic_match is False:
            return RuleResult(rule_id, rule_name, rule_sub, True, f'投诉内容描述与工单小类"{order_type}"不匹配', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_03_duplicate(self, order: Dict) -> RuleResult:
        """规则3: 随意建单-重复建单"""
        rule = self.get_rule_by_id('3')
        rule_id = '3'
        rule_name = '随意建单'
        rule_sub = '重复建单'
        penalty = rule['penalty_amount'] if rule else 50

        if not self._is_self_created_order(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '非自建单，不触发', penalty, False)
        if self._contains_excluded_dept(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '创建部门包含分拨/供应链，剔除', penalty, False)

        is_duplicate = order.get('is_duplicate_order', False)
        if is_duplicate:
            return RuleResult(rule_id, rule_name, rule_sub, True, '同一运单同一问题在自然日内存在重复建单', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_04_wrong_phone(self, order: Dict) -> RuleResult:
        """规则4: 随意建单-预留号码错误"""
        rule = self.get_rule_by_id('4')
        rule_id = '4'
        rule_name = '随意建单'
        rule_sub = '网点自建单预留号码错误'
        penalty = rule['penalty_amount'] if rule else 50

        if not self._is_self_created_order(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '非自建单，不触发', penalty, False)
        if self._contains_excluded_dept(order):
            return RuleResult(rule_id, rule_name, rule_sub, False, '创建部门包含分拨/供应链，剔除', penalty, False)

        callback_phone = str(order.get('回访号码', '') or order.get('callback_phone', ''))
        valid_phones = [
            str(order.get('网点联系电话', '') or order.get('branch_contact_phone', '')),
            str(order.get('网点客服电话', '') or order.get('branch_service_phone', '')),
            str(order.get('网点云呼电话', '') or order.get('branch_cloud_phone', '')),
            str(order.get('用户手机号', '') or order.get('user_phone', '')),
        ]

        if callback_phone and all(callback_phone != p for p in valid_phones if p):
            return RuleResult(rule_id, rule_name, rule_sub, True, f'回访号码{callback_phone}不在网点有效电话中', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_08_no_result(self, order: Dict) -> RuleResult:
        """规则8: 虚假处理-未有结果结办工单"""
        rule = self.get_rule_by_id('8')
        rule_id = '8'
        rule_name = '虚假处理'
        rule_sub = '未有结果结办工单'
        penalty = rule['penalty_amount'] if rule else 100

        close_detail = order.get('结办详情', '') or order.get('close_detail', '')
        invalid_keywords = ['核实中', '处理中', '已反馈', '对接中']

        if not close_detail:
            return RuleResult(rule_id, rule_name, rule_sub, True, '结办详情为空', penalty, True)

        for kw in invalid_keywords:
            if kw in close_detail:
                return RuleResult(rule_id, rule_name, rule_sub, True, f'结办详情含无效关键词"{kw}"', penalty, True)

        has_specific_result = order.get('has_specific_result', True)
        if not has_specific_result:
            return RuleResult(rule_id, rule_name, rule_sub, True, '无具体处理回复内容', penalty, True)

        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_09_text_quality(self, order: Dict) -> RuleResult:
        """规则9_1: 虚假处理-未回复客户(文字质检)"""
        rule = self.get_rule_by_id('9_1')
        rule_id = '9_1'
        rule_name = '虚假处理'
        rule_sub = '工单处理结果未回复客户'
        penalty = rule['penalty_amount'] if rule else 100

        detail_phone = order.get('结办详情电话号码', '') or order.get('detail_phone', '')
        callback_phone = order.get('回访号码', '') or order.get('callback_phone', '')

        if detail_phone and callback_phone and str(detail_phone) != str(callback_phone):
            return RuleResult(rule_id, rule_name, rule_sub, True, f'结办号码{detail_phone}≠回访号码{callback_phone}', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_09_call_quality(self, order: Dict) -> RuleResult:
        """规则9_2: 虚假处理-未回复客户(通话质检)"""
        rule = self.get_rule_by_id('9_2')
        rule_id = '9_2'
        rule_name = '虚假处理'
        rule_sub = '工单处理结果未回复客户'
        penalty = rule['penalty_amount'] if rule else 100

        order_stage = order.get('工单环节', '') or order.get('order_stage', '')

        if order_stage in ['省区环节', '分拨环节']:
            has_cloud = order.get('has_cloud_call_record_before30min', False)
            has_attach = order.get('has_attachment_call_record_before30min', False)
            if not has_cloud and not has_attach:
                return RuleResult(rule_id, rule_name, rule_sub, True, '省区/分拨环节: 结办前30分钟无云呼且无附件通话记录', penalty, True)
        elif order_stage == '网点环节':
            has_attach = order.get('has_attachment_call_record_before30min', False)
            if not has_attach:
                return RuleResult(rule_id, rule_name, rule_sub, True, '网点环节: 结办前30分钟无附件通话记录', penalty, True)

        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_10_promise_route(self, order: Dict) -> RuleResult:
        """规则10: 虚假处理-承诺未兑现(中转路由未更新)"""
        rule = self.get_rule_by_id('10')
        rule_id = '10'
        rule_name = '虚假处理'
        rule_sub = '承诺未兑现-中转路由未更新'
        penalty = rule['penalty_amount'] if rule else 100

        small_cat = order.get('工单小类', '') or order.get('small_category', '')
        target_types = ['回单延误', '退件延误', '中转延误', '发件延误']

        if small_cat not in target_types:
            return RuleResult(rule_id, rule_name, rule_sub, False, '不满足触发条件(非延误类型)', penalty, False)

        promise_time = order.get('承诺时间', '') or order.get('promise_time', '')
        has_route = order.get('has_route_in_promise', False)

        if not has_route:
            return RuleResult(rule_id, rule_name, rule_sub, True, f'承诺时间{promise_time}范围内无路由数据', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '承诺时间内有路由更新', penalty, True)

    def _rule_11_promise_sign(self, order: Dict) -> RuleResult:
        """规则11: 虚假处理-承诺未兑现(无签收记录)"""
        rule = self.get_rule_by_id('11')
        rule_id = '11'
        rule_name = '虚假处理'
        rule_sub = '承诺未兑现-无签收记录'
        penalty = rule['penalty_amount'] if rule else 100

        close_detail = order.get('结办详情', '') or order.get('close_detail', '')
        # 拒收/不要直接无责
        if '拒收' in str(close_detail) or '不要' in str(close_detail):
            return RuleResult(rule_id, rule_name, rule_sub, False, '结办详情含"拒收/不要"，直接判定无责', penalty, True)

        promise_date = order.get('承诺派送日期', '') or order.get('promise_delivery_date', '')
        has_sign = order.get('has_sign_record', False)
        sign_date = order.get('签收日期', '') or order.get('sign_date', '')

        if not has_sign:
            return RuleResult(rule_id, rule_name, rule_sub, True, f'承诺派送日期{promise_date}无签收记录', penalty, True)
        try:
            if pd.Timestamp(promise_date) < pd.Timestamp(sign_date):
                return RuleResult(rule_id, rule_name, rule_sub, True, f'签收日期{sign_date}晚于承诺日期{promise_date}', penalty, True)
        except:
            pass
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_12_promise_pickup(self, order: Dict) -> RuleResult:
        """规则12: 虚假处理-承诺未兑现(揽收无开单记录)"""
        rule = self.get_rule_by_id('12')
        rule_id = '12'
        rule_name = '虚假处理'
        rule_sub = '承诺未兑现-揽收无开单记录'
        penalty = rule['penalty_amount'] if rule else 100

        small_cat = order.get('工单小类', '') or order.get('small_category', '')
        if small_cat != '揽件超时':
            return RuleResult(rule_id, rule_name, rule_sub, False, '不满足触发条件(非揽件超时)', penalty, False)

        waybill = order.get('运单号', '') or order.get('waybill_no', '')
        if not str(waybill).startswith('W'):
            return RuleResult(rule_id, rule_name, rule_sub, False, '运单号非W开头', penalty, False)

        has_create_record = order.get('has_create_record', False)
        if not has_create_record:
            return RuleResult(rule_id, rule_name, rule_sub, True, '无路由开单记录', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_13_promise_report(self, order: Dict) -> RuleResult:
        """规则13: 虚假处理-承诺未兑现(未查询到上报记录)"""
        rule = self.get_rule_by_id('13')
        rule_id = '13'
        rule_name = '虚假处理'
        rule_sub = '承诺未兑现-未查询到上报记录'
        penalty = rule['penalty_amount'] if rule else 100

        close_detail = order.get('结办详情', '') or order.get('close_detail', '')
        has_claim_data = order.get('has_claim_data', False)

        report_keywords = ['已上报', '收集', '提供', '资料']
        has_report_claim = any(kw in str(close_detail) for kw in report_keywords)

        if has_report_claim and not has_claim_data:
            return RuleResult(rule_id, rule_name, rule_sub, True, '声称已上报但无理赔数据', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_14_promise_change(self, order: Dict) -> RuleResult:
        """规则14: 虚假处理-承诺未兑现(未查询到改单记录)"""
        rule = self.get_rule_by_id('14')
        rule_id = '14'
        rule_name = '虚假处理'
        rule_sub = '承诺未兑现-未查询到改单记录'
        penalty = rule['penalty_amount'] if rule else 100

        close_detail = order.get('结办详情', '') or order.get('close_detail', '')
        change_keywords = ['已更改', '已改单']
        has_change_claim = any(kw in str(close_detail) for kw in change_keywords)

        if not has_change_claim:
            return RuleResult(rule_id, rule_name, rule_sub, False, '不触发', penalty, False)

        has_change_record = order.get('has_change_record', False)
        change_time = order.get('change_create_time', '') or order.get('info_change_time', '')
        close_time = order.get('工单结办时间', '') or order.get('close_time', '')

        if not has_change_record:
            return RuleResult(rule_id, rule_name, rule_sub, True, '声称已改单但无改单记录', penalty, True)

        try:
            if change_time and close_time and pd.Timestamp(change_time) > pd.Timestamp(close_time):
                return RuleResult(rule_id, rule_name, rule_sub, True, f'改单时间{change_time}晚于结办时间{close_time}', penalty, True)
        except:
            pass
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_15_false_response(self, order: Dict) -> RuleResult:
        """规则15: 虚假响应"""
        rule = self.get_rule_by_id('15')
        rule_id = '15'
        rule_name = '虚假响应'
        rule_sub = '时效要求内首次响应客户告知投诉处理进度'
        penalty = rule['penalty_amount'] if rule else 100

        order_stage = order.get('工单环节', '') or order.get('order_stage', '')
        create_time = order.get('工单新增时间', '') or order.get('create_time', '')
        take_time = order.get('工单领单时间', '') or order.get('take_time', '')
        has_response_within_30min = order.get('has_response_within_30min', False)

        if not create_time or not take_time:
            return RuleResult(rule_id, rule_name, rule_sub, False, '缺少时间信息', penalty, True)

        if order_stage in ['省区环节', '分拨环节']:
            if not has_response_within_30min:
                return RuleResult(rule_id, rule_name, rule_sub, True, '新增-领单+30min内无响应记录', penalty, True)

        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_16_no_second_call(self, order: Dict) -> RuleResult:
        """规则16: 登记不规范-未联系第二通"""
        rule = self.get_rule_by_id('16')
        rule_id = '16'
        rule_name = '登记不规范'
        rule_sub = '工单处理电话回复客户无人接听时未联系第二通'
        penalty = rule['penalty_amount'] if rule else 10

        last_call_connected = order.get('last_call_connected', True)
        last_call_duration = order.get('last_call_duration', 0) or order.get('通话时长', 0)
        call_gap_sufficient = order.get('call_gap_sufficient', True)

        if last_call_connected and float(last_call_duration) >= 15:
            return RuleResult(rule_id, rule_name, rule_sub, False, '最后一通已接通且时长≥15s', penalty, True)

        if not call_gap_sufficient:
            return RuleResult(rule_id, rule_name, rule_sub, True, '未接通电话相邻间隔<30分钟', penalty, True)

        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_17_greeting(self, order: Dict) -> RuleResult:
        """规则17: 播报语-开头语未播报品牌"""
        rule = self.get_rule_by_id('17')
        rule_id = '17'
        rule_name = '播报语'
        rule_sub = '开头语未播报品牌'
        penalty = rule['penalty_amount'] if rule else 10

        has_greeting = order.get('has_brand_greeting', True)
        if not has_greeting:
            return RuleResult(rule_id, rule_name, rule_sub, True, '通话内容无品牌招呼语', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_18_closing(self, order: Dict) -> RuleResult:
        """规则18: 播报语-结束语未播报"""
        rule = self.get_rule_by_id('18')
        rule_id = '18'
        rule_name = '播报语'
        rule_sub = '结束语未播报'
        penalty = rule['penalty_amount'] if rule else 10

        has_closing = order.get('has_closing_script', True)
        if not has_closing:
            return RuleResult(rule_id, rule_name, rule_sub, True, '通话内容无结束语', penalty, True)
        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)

    def _rule_19_forbidden_behavior(self, order: Dict) -> RuleResult:
        """规则19: 禁言禁行"""
        rule = self.get_rule_by_id('19')
        rule_id = '19'
        rule_name = '禁言禁行'
        rule_sub = '禁言禁行'
        penalty = rule['penalty_amount'] if rule else 200

        has_negative = order.get('has_negative_emotion', False)
        has_forbidden_words = order.get('has_forbidden_keywords', False)

        forbidden_keywords = [
            '12315', '12345', '市场热线', '外投', '运管局', '维权', '曝光',
            '法律途径', '监管部门', '邮局', '市监局', '税务局', '发帖'
        ]

        call_content = order.get('通话内容', '') or order.get('call_content', '')
        found_keywords = [kw for kw in forbidden_keywords if kw in str(call_content)]

        if has_negative:
            return RuleResult(rule_id, rule_name, rule_sub, True, '通话存在负面情绪', penalty, True)
        if found_keywords:
            return RuleResult(rule_id, rule_name, rule_sub, True, f'通话含违禁词: {found_keywords}', penalty, True)
        if has_forbidden_words:
            return RuleResult(rule_id, rule_name, rule_sub, True, '通话存在禁言禁行行为', penalty, True)

        return RuleResult(rule_id, rule_name, rule_sub, False, '', penalty, True)


def load_rules(filepath: str = '工作簿4.xlsx') -> pd.DataFrame:
    """加载规则表"""
    return pd.read_excel(filepath, sheet_name='Sheet1')


def load_orders(filepath: str) -> pd.DataFrame:
    """加载工单数据"""
    return pd.read_excel(filepath)


if __name__ == '__main__':
    # 测试规则引擎
    rules_df = load_rules()
    engine = WorkOrderRuleEngine(rules_df)

    print(f"加载 {len(engine.rules)} 条规则")
    print(f"在线规则: {len(engine.get_active_rules())} 条")
    print(f"规则分类: {list(engine.get_rule_categories().keys())}")

    # 模拟一条工单数据
    test_order = {
        '工单大类': '服务类',
        '工单小类': '虚假签收',
        '工单环节': '网点环节',
        '投诉渠道': 'iwos',
        '结办详情': '核实中',
        'has_cloud_call_record': False,
        'has_attachment_call_record': False,
    }

    results = engine.judge(test_order)
    print(f"\n测试工单判定结果:")
    for r in results:
        if r.rule_triggered:
            status = '🔴 有责' if r.is_responsible else '🟢 无责'
            print(f"  规则{r.rule_id} [{r.rule_name}-{r.rule_sub}]: {status} | {r.reason} | 罚款{r.penalty_amount}元")

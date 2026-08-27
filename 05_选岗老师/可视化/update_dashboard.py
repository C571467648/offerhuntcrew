# -*- coding: utf-8 -*-
"""
投递看板自动更新脚本（固定接口）
================================
用法：更新投递总表（Excel）后，运行本脚本即可自动重新生成 投递看板.html。

    python update_dashboard.py

数据源（固定）：
  05_选岗老师/数据/秋招央国企岗位投递表.xlsx    （sheet: 央国企岗位投递表）
  05_选岗老师/数据/秋招私企投递记录表.xlsx      （sheet: 算子方向 / C++方向）

列名约定（勿改表头）：
  央国企：公司、网申日期、地点、岗位、是否投递、简历、笔试、一面、二面、三面、offer、链接、优先级、批次
  私企  ：公司、批次、测评截止时间、地点、优先级、岗位、是否投递、简历、测评、笔试、一面、二面、三面、offer、链接

状态推导：按 offer→三面→二面→一面→笔试→测评→简历 取最深非空阶段；
          阶段值：是→阶段名；通过→阶段+过；挂→阶段+挂（任何阶段"挂"=终态）。
是否投递：是→投递明细（看板"已投"）；非是/空→央国企按窗口归入"近期开闸·行动清单"（8-9月/9月/9-10月），私企不再展示关注池。

输出：05_选岗老师/可视化/投递看板.html（覆盖）
"""
import os
import re
import json
import datetime

# 用户姓名（看板标题显示）。按提示词指引由总架构师替换为使用者姓名，如 USER_NAME = '张同学'。
USER_NAME = '{{用户姓名}}'

BASE = os.path.dirname(os.path.abspath(__file__))          # 05_选岗老师/可视化
DATA_DIR = os.path.normpath(os.path.join(BASE, '..', '数据'))  # 05_选岗老师/数据
TIP_DIR = os.path.normpath(os.path.join(BASE, '..', '岗位情报'))  # 05_选岗老师/岗位情报
TEMPLATE = os.path.join(BASE, 'dashboard_template.html')
OUT = os.path.join(BASE, '投递看板.html')

SOE_FILE = '秋招央国企岗位投递表.xlsx'
PRV_FILE = '秋招私企投递记录表.xlsx'
TIPS_FILE = '新增岗位情报.xlsx'


def norm_cell(v):
    if v is None:
        return ''
    s = re.sub(r'\s+', ' ', str(v)).replace('\\n', ' ').strip()
    return '' if s.lower() == 'none' else s


def combo(stage, val):
    """把阶段列的值拼成状态串，如 简历+过 -> 简历过；笔试+待参加 -> 笔试待参加"""
    v = norm_cell(val)
    if not v:
        return None
    if v == '是':
        return stage
    if v == '通过':
        v = '过'
    return stage + v


def soe_status(row):
    for stage in ('offer', '三面', '二面', '一面', '笔试', '简历'):
        s = combo(stage, row.get(stage))
        if s:
            return s
    return '流程中'


def prv_status(row):
    for stage in ('offer', '三面', '二面', '一面', '笔试', '测评', '简历'):
        s = combo(stage, row.get(stage))
        if s:
            return s
    return '流程中'


def window_of(d):
    s = norm_cell(d)
    if '春招' in s:
        return '3-4月(春招)'
    if '8-9' in s or ('8' in s and '9' in s and '10' not in s and '11' not in s):
        return '8-9月'
    if '9-10' in s or '10' in s and '9' in s:
        return '9-10月'
    if '11-12' in s:
        return '11-12月'
    if '10-11' in s:
        return '10-11月'
    if '11' in s:
        return '11月'
    if '10' in s:
        return '10月'
    if '9' in s:
        return '9月'
    if '8' in s:
        return '8月'
    return '未标注'


BATCH_NORM = {'实习(练手)': '实习', '实习批次': '实习', '正式批次': '正式批'}


def norm_batch(b):
    v = norm_cell(b)
    if v.startswith('实习'):
        return '实习'
    return BATCH_NORM.get(v, v) or '-'


def load_soe(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], [], []
    header = list(rows[0])
    applied, pool, win_cnt = [], [], {}
    for r in rows[1:]:
        if not any(v is not None and str(v).strip() for v in r):
            continue
        d = dict(zip(header, r))
        co = norm_cell(d.get('公司'))
        if not co:
            continue
        rec = {
            'co': co,
            'date': norm_cell(d.get('网申日期')),
            'loc': norm_cell(d.get('地点')),
            'job': norm_cell(d.get('岗位')),
            'link': norm_cell(d.get('链接')),
            'pri': norm_cell(d.get('优先级')),
            'batch': norm_cell(d.get('批次')),
        }
        if norm_cell(d.get('是否投递')) == '是':
            rec['st'] = soe_status(d)
            applied.append(rec)
        else:
            rec['win'] = window_of(d.get('网申日期'))
            pool.append(rec)
            win_cnt[rec['win']] = win_cnt.get(rec['win'], 0) + 1
    # 窗口按固定顺序输出，未标注放最后
    order = ['8-9月', '9月', '9-10月', '10月', '10-11月', '11月', '11-12月', '3-4月(春招)', '未标注']
    keys = sorted(win_cnt, key=lambda k: order.index(k) if k in order else 99)
    windows = [{'name': k, 'n': win_cnt[k]} for k in keys]
    return applied, pool, windows


def load_prv(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out, pool = [], []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = list(rows[0])
        for r in rows[1:]:
            if not any(v is not None and str(v).strip() for v in r):
                continue
            d = dict(zip(header, r))
            co = norm_cell(d.get('公司'))
            if not co:
                continue
            rec = {
                'dir': sheet,
                'co': co,
                'batch': norm_batch(d.get('批次')),
                'loc': norm_cell(d.get('地点')),
                'pri': norm_cell(d.get('优先级')) or '-',
                'job': norm_cell(d.get('岗位')),
                'link': norm_cell(d.get('链接')),
            }
            if norm_cell(d.get('是否投递')) == '是':
                rec['st'] = prv_status(d)
                out.append(rec)
            else:
                pool.append(rec)   # 未投递 → 关注池（AI推荐）
    wb.close()
    return out, pool


TIPS_KEYS = ('公司', '性质', '届别', '岗位', '地点', '截止窗口', '匹配度', '备注', '入口')


def load_jobtips(path):
    """读 新增岗位情报.xlsx：分类列=央国企→soe 模块，私企→prv 模块；文件缺失时返回空（兼容开源版初始状态）。"""
    if not os.path.exists(path):
        return [], []
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    header = list(rows[0])
    soe, prv = [], []
    for r in rows[1:]:
        if not any(v is not None and str(v).strip() for v in r):
            continue
        d = dict(zip(header, r))
        co = norm_cell(d.get('公司'))
        if not co:
            continue
        rec = {k: norm_cell(d.get(k)) for k in TIPS_KEYS}
        (soe if norm_cell(d.get('分类')) == '央国企' else prv).append(rec)
    return soe, prv


def tips_date():
    """从最新情报清单文件名提取日期（如 2026-08-27_...xlsx → '8-27'），无则返回空串。"""
    if not os.path.isdir(TIP_DIR):
        return ''
    files = [f for f in os.listdir(TIP_DIR) if f.endswith('.xlsx') and '情报清单' in f]
    if not files:
        return ''
    files.sort(reverse=True)
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', files[0])
    return '%s-%s' % (m.group(2).lstrip('0'), m.group(3)) if m else ''


def main():
    soe_applied, soe_pool, soe_windows = load_soe(os.path.join(DATA_DIR, SOE_FILE))
    prv, prv_pool = load_prv(os.path.join(DATA_DIR, PRV_FILE))
    soe_tips, prv_tips = load_jobtips(os.path.join(TIP_DIR, TIPS_FILE))

    data = {
        'date': datetime.date.today().strftime('%Y-%m-%d'),
        'soe': {'applied': soe_applied, 'pool': soe_pool, 'windows': soe_windows},
        'prv': prv,
        'prv_pool': prv_pool,
        'soe_tips': soe_tips,
        'prv_tips': prv_tips,
        'tips_date': tips_date(),
    }
    payload = json.dumps(data, ensure_ascii=False, indent=1).replace('</', '<\\/')

    with open(TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()
    html = html.replace('/*__DATA__*/ null', '/*__DATA__*/ ' + payload)
    html = html.replace('{{USER_NAME}}', USER_NAME)

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)

    total_prv = len(prv)
    dead_prv = sum(1 for r in prv if r['st'].endswith('挂'))
    act_now = sum(1 for r in soe_pool if r['win'] in ('8-9月', '9月', '9-10月'))
    print('[OK] 投递看板已生成: %s' % OUT)
    print('     央国企: 已投 %d 家, 计划投 %d 家, 近期开闸行动清单 %d 家' % (len(soe_applied), len(soe_pool), act_now))
    print('     私企  : 共 %d 条, 已挂 %d 条' % (total_prv, dead_prv))
    print('     新增岗位情报: 央国企 %d 条, 私企 %d 条' % (len(soe_tips), len(prv_tips)))


if __name__ == '__main__':
    main()

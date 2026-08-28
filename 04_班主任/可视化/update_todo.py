# -*- coding: utf-8 -*-
"""班主任 · 待办看板 生成器（固定接口）
用法：
  python update_todo.py    # 读最新待办MD，生成 待办看板.html（班主任每次更新待办后执行）
勾选状态保存在浏览器本地（localStorage），无需本地服务、无需额外存储文件。
"""
import datetime
import json
import os
import re

ROOT = os.path.dirname(os.path.abspath(__file__))
TEACHER_DIR = os.path.dirname(ROOT)                                                      # 04_班主任
INPUT_DIR = os.path.join(TEACHER_DIR, '待办历史')                                        # 待办MD所在（历史归档+当天文件）
OUTPUT_HTML = os.path.join(ROOT, '待办看板.html')
TEMPLATE = os.path.join(ROOT, 'todo_template.html')
# 看板品牌名（开源版请由总架构师替换为使用者姓名）
USER_NAME = '{{用户姓名}}'


def find_latest_todo():
    if not os.path.isdir(INPUT_DIR):
        return None
    files = [f for f in os.listdir(INPUT_DIR) if f.endswith('_班主任_待办.md')]
    if not files:
        return None
    files.sort(reverse=True)  # 文件名以日期开头，倒序取最新
    return os.path.join(INPUT_DIR, files[0])


def parse_md(path):
    with open(path, encoding='utf-8') as f:
        text = f.read()
    # 日期优先取 MD 标题「今日待办（YYYY-MM-DD ...）」，兜底取文件名（保证每天更新后今日勾选自动隔离）
    m_date = re.search(r'^##\s*今日待办\s*[（(]\s*(\d{4}-\d{2}-\d{2})', text, re.M)
    date = m_date.group(1) if m_date else os.path.basename(path)[:10]
    sections = {'今日待办': {'range': '', 'items': []}, '本周待办': {'range': '', 'items': []}}
    cur = None
    for line in text.splitlines():
        m = re.match(r'^##\s*(今日待办|本周待办)\s*[（(](.*?)[)）]', line)
        if m:
            cur = m.group(1)
            sections[cur]['range'] = m.group(2)
            continue
        if cur and line.startswith('|') and '---' not in line:
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if len(cells) >= 3 and cells[0] and cells[0] != '事项':
                sections[cur]['items'].append({'text': cells[0], 'pri': cells[1], 'due': cells[2]})
    return date, sections


def iso_week(date_str):
    """由日期算 ISO 周年（如 2026-W35），用于本周待办勾选的周级隔离：同周保留、跨周清空"""
    try:
        y, m, d = map(int, date_str.split('-'))
        iso = datetime.date(y, m, d).isocalendar()
        return '%d-W%02d' % (iso[0], iso[1])
    except Exception:
        return date_str


def build():
    src = find_latest_todo()
    if not src:
        print('[待办看板] 未找到 *_班主任_待办.md（请先在 04_班主任/待办历史/ 落盘待办文档）')
        return False
    date, sections = parse_md(src)
    with open(TEMPLATE, encoding='utf-8') as f:
        html = f.read()
    payload = json.dumps({'date': date, 'weekKey': iso_week(date),
                          'today': sections['今日待办'], 'week': sections['本周待办']},
                         ensure_ascii=False)
    html = html.replace('const DATA = /*__DATA__*/ null;', 'const DATA = ' + payload + ';')
    html = html.replace('{{USER_NAME}}', USER_NAME)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    n = len(sections['今日待办']['items']) + len(sections['本周待办']['items'])
    print(f'[待办看板] 已生成：{OUTPUT_HTML}（{date}，今日{len(sections["今日待办"]["items"])}项 + 本周{len(sections["本周待办"]["items"])}项）')
    return True


if __name__ == '__main__':
    build()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会员续费工作台 · 数据更新脚本
用法:
  python3 update_data.py [--push] [excel_path]

从 Excel 的「会员明细」和「280会员明细」提取数据，注入 index.html。
- 跳过 2025-12 到期行
- V 列(col[21]) 判断续费状态 (Y=已续费)
- --push: 自动 git commit + push
"""
import sys
import os
import re
import json
import subprocess

try:
    import openpyxl
except ImportError:
    print('未找到 openpyxl，请使用 /usr/bin/python3 运行:')
    print('  /usr/bin/python3 update_data.py [--push] [excel_path]')
    sys.exit(1)

EMPS = ['蔡梦珊', '胡锆宇', '黄剑钦', '项波', '黄丹']
DEFAULT_EXCEL = '/Users/huangdan/Desktop/中山26年续费率.xlsx'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(BASE_DIR, 'index.html')


def parse_month(expiry):
    if isinstance(expiry, str):
        exp_str = expiry.replace('/', '-')
        try:
            parts = exp_str.split('-')
            return int(parts[0]), int(parts[1])
        except Exception:
            return None, None
    elif expiry:
        return expiry.year, expiry.month
    return None, None


def fmt_expiry(expiry):
    if isinstance(expiry, str):
        return expiry.replace('/', '-')
    elif expiry:
        return expiry.strftime('%Y-%m-%d')
    return '2026-12-31'


def classify_prod(amount):
    if amount <= 400:
        return '基础会员-400'
    elif amount <= 660:
        return '全功能-660'
    elif amount <= 680:
        return '全功能-680'
    return '全功能-980'


def extract_members(ws, amount_default):
    """提取会员数据，返回 (members_dict, counts_dict, monthly_dict, agency_stats_dict)"""
    members = {emp: [] for emp in EMPS}
    counts = {emp: [0, 0] for emp in EMPS}
    monthly = {emp: {} for emp in EMPS}
    agency = {emp: {} for emp in EMPS}

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        emp = str(row[1] or '').strip()
        if emp not in EMPS:
            continue
        y, m = parse_month(row[12])
        if (y, m) == (2025, 12):
            continue
        name = str(row[4] or '').strip()
        if not name:
            continue
        agency_name = str(row[0] or '').strip()
        amount = float(row[14]) if row[14] else amount_default
        is_renewed = str(row[21] or '').strip() == 'Y' if len(row) > 21 else False
        credit = str(row[5] or '').strip()
        is_ts = str(row[15] or '').strip() == '托收' if len(row) > 15 else False
        exp_str = fmt_expiry(row[12])
        exp_month = 12
        if isinstance(row[12], str):
            try:
                exp_month = int(exp_str.split('-')[1]) if '-' in exp_str else 12
            except Exception:
                exp_month = 12
        elif row[12]:
            exp_month = row[12].month

        counts[emp][0] += 1
        if is_renewed:
            counts[emp][1] += 1
        monthly[emp].setdefault(exp_month, [0, 0])
        monthly[emp][exp_month][0] += 1
        if is_renewed:
            monthly[emp][exp_month][1] += 1

        if amount_default == 0:
            prod = classify_prod(amount)
            members[emp].append({'n': name, 'c': credit, 'a': agency_name,
                                 'p': prod, 'pr': amount, 'ex': exp_str,
                                 'renewed': is_renewed, 'ts': is_ts})
        else:
            members[emp].append({'n': name, 'c': credit, 'a': agency_name,
                                 'ex': exp_str, 'p': amount,
                                 'renewed': is_renewed, 'ts': is_ts})

        if agency_name:
            ag = agency[emp].setdefault(agency_name, [0, 0, 0.0])
            ag[0] += 1
            if is_renewed:
                ag[1] += 1
            ag[2] += amount

    # 排序: 未续费在前，已续费在后（与现有数据一致）
    sorted_members = {}
    for emp in EMPS:
        pend = [x for x in members[emp] if not x['renewed']]
        renew = [x for x in members[emp] if x['renewed']]
        sorted_members[emp] = pend + renew

    sorted_agency = {emp: {a: s for a, s in sorted(agency[emp].items())} for emp in EMPS}
    return sorted_members, counts, monthly, sorted_agency


def build_monthly_str(monthly, emp):
    parts = []
    for m in sorted(monthly[emp]):
        v = monthly[emp][m]
        r = round(v[1] / v[0], 4) if v[0] else 0
        parts.append(f'{m}:{{t:{v[0]},Y:{v[1]},N:{v[0]-v[1]},r:{r}}}')
    return ','.join(parts)


def replace_block(c, name, value):
    start = c.find('const ' + name + ' = ')
    if start < 0:
        raise RuntimeError(f'未找到 const {name}')
    end = c.find(';', start)
    if end < 0:
        raise RuntimeError(f'未找到 {name} 的结束符')
    return c[:start] + f'const {name} = {value}' + c[end:]


def update_emp_data(c, counts, monthly):
    emps = EMPS
    for emp in emps:
        t, r = counts[emp]
        p = t - r
        rate = round(r / t, 4) if t > 0 else 0
        m_str = build_monthly_str(monthly, emp)

        idx = c.find("'" + emp + "': {")
        if idx < 0:
            idx = c.find('"' + emp + '": {')
        if idx < 0:
            raise RuntimeError(f'未找到 EMP_DATA 中的 {emp}')
        next_idx = len(c)
        for ne in emps:
            if ne != emp:
                pos = c.find("'" + ne + "': {", idx + 1)
                if pos > 0 and pos < next_idx:
                    next_idx = pos
        entry = c[idx:next_idx]
        entry = re.sub(r'total:\d+', 'total:' + str(t), entry)
        entry = re.sub(r'renewed:\d+', 'renewed:' + str(r), entry)
        entry = re.sub(r'pending:\d+', 'pending:' + str(p), entry)
        entry = re.sub(r'rate:0\.\d+', 'rate:' + str(rate), entry)
        ms = entry.find('monthly:{')
        me = entry.find('},r280:')
        if ms > 0 and me > 0:
            entry = entry[:ms] + 'monthly:{' + m_str + '}' + entry[me + 1:]
        c = c[:idx] + entry + c[next_idx:]
    return c


def main():
    args = sys.argv[1:]
    do_push = '--push' in args
    excel_path = DEFAULT_EXCEL
    for a in args:
        if a != '--push' and not a.startswith('-'):
            excel_path = a

    print(f'读取: {excel_path}')
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    # 会员明细
    ws = wb['会员明细']
    members, counts, monthly, agency = extract_members(ws, 0)
    agency_full = {emp: {a: s for a, s in sorted(agency[emp].items())} for emp in EMPS}
    emp_agencies = {emp: sorted(agency[emp].keys()) for emp in EMPS}

    # 280会员明细
    ws280 = wb['280会员明细']
    members280, counts280, _, agency280 = extract_members(ws280, 280)
    agency_280 = {emp: {a: s for a, s in sorted(agency280[emp].items())} for emp in EMPS}
    emp_agencies_280 = {emp: sorted(agency280[emp].keys()) for emp in EMPS}

    wb.close()

    with open(HTML, 'r') as f:
        c = f.read()

    c = replace_block(c, 'REAL_MEMBERS', json.dumps(members, ensure_ascii=False, separators=(',', ':')))
    c = replace_block(c, 'REAL_MEMBERS_280', json.dumps(members280, ensure_ascii=False, separators=(',', ':')))
    c = replace_block(c, 'AGENCY_FULL_STATS', json.dumps(agency_full, ensure_ascii=False, separators=(',', ':')))
    c = replace_block(c, 'EMP_AGENCIES', json.dumps(emp_agencies, ensure_ascii=False, separators=(',', ':')))
    c = replace_block(c, 'AGENCY_280_STATS', json.dumps(agency_280, ensure_ascii=False, separators=(',', ':')))
    c = replace_block(c, 'EMP_AGENCIES_280', json.dumps(emp_agencies_280, ensure_ascii=False, separators=(',', ':')))
    c = update_emp_data(c, counts, monthly)

    with open(HTML, 'w') as f:
        f.write(c)

    # 验证
    js = c[c.find('<script>') + len('<script>'):c.rfind('</script>')]
    braces = js.count('{') - js.count('}')
    print(f'JS braces diff: {braces}')
    for emp in EMPS:
        t, r = counts[emp]
        print(f'{emp}: 总{t} 已续{r} 续费率 {r/t*100 if t else 0:.1f}%')
    print(f'index.html 已更新: {os.path.getsize(HTML)/1024/1024:.2f} MB')

    if do_push:
        subprocess.run(['git', 'add', 'index.html'], cwd=BASE_DIR, check=True)
        subprocess.run(['git', 'commit', '-m', '更新续费数据'], cwd=BASE_DIR, check=True)
        subprocess.run(['git', 'push'], cwd=BASE_DIR, check=True)
        print('已提交并推送')


if __name__ == '__main__':
    main()

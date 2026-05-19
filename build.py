#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
4G模块选型工具 - 一键构建脚本（v3.0 - 以产品为主体架构）
用法：运行此脚本 → 读取 data.json + template.html → 生成 index.html（数据内嵌）

架构：
- data.json：存储国家数据（扁平列表）
- build.py：定义产品表PRODUCTS_DB + 模块表MODULES，自动计算国家-产品-模块匹配
- MODULES dict：每个模块的频段、GSM支持等
- PRODUCTS_DB：以产品为主体，记录每个产品支持的模块和类别
更新数据：编辑 PRODUCTS_DB 或 MODULES → 重新运行此脚本即可；也可通过Excel导入更新
"""

import json
import os

# ============================================================
# 路径配置
# ============================================================
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(TOOL_DIR, 'data.json')
TEMPLATE_FILE = os.path.join(TOOL_DIR, 'template.html')
OUTPUT_FILE = os.path.join(TOOL_DIR, 'index.html')

# ============================================================
# 模块定义（频段信息，每个模块的技术规格）
# ============================================================
MODULES = {
    'MC669-CN': {
        'name': "MC669-CN",
        'bands': {'FDD': ["B1","B3","B5","B8"], 'TDD': ["B34","B38","B39","B40","B41"]},
        'has_gsm': False,
        'region': "亚太",
        'note': "亚太区域模块"
    },
    'LE270-CN': {
        'name': "LE270-CN",
        'bands': {'FDD': ["B1","B3","B5","B8"], 'TDD': ["B34","B38","B39","B40","B41"]},
        'has_gsm': False,
        'region': "亚太",
        'note': "亚太区域模块"
    },
    'LE270-EU': {
        'name': "LE270-EU",
        'bands': {'FDD': ["B1","B3","B5","B7","B8","B20","B28"], 'TDD': ["B38","B40","B41"]},
        'has_gsm': False,
        'region': "欧洲",
        'note': "欧洲区域模块"
    },
    'LE270-GL': {
        'name': "LE270-GL",
        'bands': {'FDD': ["B1","B2","B3","B4","B5","B7","B8","B12","B13","B14","B17","B18","B19","B20","B25","B26","B28","B66"], 'TDD': ["B34","B38","B39","B40","B41"]},
        'has_gsm': False,
        'region': "全球",
        'note': "全球区域模块"
    },
    'L610-CN': {
        'name': "L610-CN",
        'bands': {'FDD': ["B1","B3","B5","B8"], 'TDD': ["B34","B39","B40","B41"]},
        'has_gsm': False,
        'region': "亚太",
        'note': "亚太区域模块"
    },
    'L610-EU': {
        'name': "L610-EU",
        'bands': {'FDD': ["B1","B3","B7","B8","B20","B28"], 'TDD': []},
        'has_gsm': True,
        'region': "欧洲",
        'note': "欧洲区域模块"
    },
    'L610-LA': {
        'name': "L610-LA",
        'bands': {'FDD': ["B1","B2","B3","B4","B5","B7","B8","B28","B66"], 'TDD': []},
        'has_gsm': True,
        'region': "拉美",
        'note': "拉美区域模块"
    },
    'ML307N-DC-DL': {
        'name': "ML307N-DC/DL",
        'bands': {'FDD': ["B1","B3","B5","B8"], 'TDD': ["B34","B38","B39","B40","B41"]},
        'has_gsm': False,
        'region': "亚太",
        'note': "亚太区域模块"
    },
}

PRODUCTS_DB = {
    'WD-219G': {
        'category': "共享产品",
        'modules': ["MC669-CN"]
    },
    'WD-210': {
        'category': "共享产品",
        'modules': ["LE270-CN","LE270-EU","LE270-GL"]
    },
    'WD-219K': {
        'category': "共享产品",
        'modules': ["LE270-CN","LE270-EU","LE270-GL"]
    },
    'WD-325': {
        'category': "两轮产品",
        'modules': ["L610-CN","L610-EU","L610-LA"]
    },
    'WD-300': {
        'category': "两轮产品",
        'modules': ["LE270-CN","LE270-EU","LE270-GL"]
    },
    'WD-110': {
        'category': "两轮产品",
        'modules': ["ML307N-DC-DL"]
    },
    'WD-281': {
        'category': "两轮产品",
        'modules': ["ML307N-DC-DL"]
    },
    'WD-282': {
        'category': "两轮产品",
        'modules': ["ML307N-DC-DL"]
    },
    'WD-280D': {
        'category': "两轮产品",
        'modules': ["ML307N-DC-DL"]
    },
}


def parse_bands_str(bands_str):
    """解析频段字符串为结构化数据"""
    result = {'FDD': [], 'TDD': [], 'WCDMA': []}
    if not bands_str:
        return result
    parts = bands_str.split()
    current_type = 'FDD'
    for part in parts:
        if part == 'FDD':
            current_type = 'FDD'
        elif part == 'TDD':
            current_type = 'TDD'
        elif part == 'WCDMA':
            current_type = 'WCDMA'
        elif '/' in part:
            for b in part.split('/'):
                bid = 'B' + b.replace('B', '')
                if bid not in result[current_type]:
                    result[current_type].append(bid)
        else:
            if part.upper().startswith('B') and len(part) > 1 and part[1:].isdigit():
                bid = part.upper()
                if bid not in result[current_type]:
                    result[current_type].append(bid)
    return result


def get_adaptation_level(module_info, country_bands):
    """计算模块对某国家的适配等级"""
    if not module_info.get('bands') or not country_bands:
        return {'grade': 'D', 'label': 'D 不推荐', 'pct': 0, 'detail': '无数据'}

    country_lte = set()
    for b in country_bands.get('FDD', []):
        country_lte.add(b)
    for b in country_bands.get('TDD', []):
        country_lte.add(b)

    module_lte = set()
    for b in module_info['bands'].get('FDD', []):
        module_lte.add(b)
    for b in module_info['bands'].get('TDD', []):
        module_lte.add(b)

    if not country_lte:
        return {'grade': 'A', 'label': 'A 高度适配', 'pct': 100, 'detail': '100%'}

    covered = len(country_lte & module_lte)
    total = len(country_lte)
    pct = round(covered / total * 100)
    has_gsm = module_info.get('has_gsm', False)

    if pct == 100 and has_gsm:
        return {'grade': 'S', 'label': 'S 完美适配', 'pct': pct, 'detail': '100%覆盖+GSM'}
    if pct == 100:
        return {'grade': 'A+', 'label': 'A+ 极佳适配', 'pct': pct, 'detail': '100%覆盖'}
    if pct >= 80 and has_gsm:
        return {'grade': 'A', 'label': 'A 高度适配', 'pct': pct, 'detail': f'{pct}%+GSM'}
    if pct >= 80 or (pct >= 60 and has_gsm):
        return {'grade': 'B+', 'label': 'B+ 良好适配', 'pct': pct, 'detail': f'{pct}%' + ('+GSM' if has_gsm else '')}
    if pct >= 40:
        return {'grade': 'B', 'label': 'B 基本可用', 'pct': pct, 'detail': f'{pct}%'}
    if pct >= 1:
        return {'grade': 'C', 'label': 'C 勉强可用', 'pct': pct, 'detail': f'{pct}%'}
    return {'grade': 'D', 'label': 'D 不推荐', 'pct': 0, 'detail': '无覆盖'}


def match_product_to_country(product_name, product_info, country_bands):
    """计算一个产品在某国家的最佳适配方案"""
    best_module = None
    best_adapt = None
    best_pct = -1

    for mod_key in product_info.get('modules', []):
        if mod_key not in MODULES:
            continue
        mod_info = MODULES[mod_key]
        adapt = get_adaptation_level(mod_info, country_bands)
        if adapt['pct'] > best_pct:
            best_pct = adapt['pct']
            best_module = mod_key
            best_adapt = adapt

    if best_module is None or best_pct == 0:
        return None

    return {
        'product': product_name,
        'category': product_info['category'],
        'module_key': best_module,
        'module_info': MODULES[best_module],
        'all_modules': product_info['modules'],
        'adapt': best_adapt
    }


def auto_match_country(country_bands):
    """对一个国家，匹配所有产品，按类别和适配等级排序"""
    grade_order = {'S': 0, 'A+': 1, 'A': 2, 'B+': 3, 'B': 4, 'C': 5, 'D': 6}
    result = {'共享产品': [], '两轮产品': []}

    for product_name, product_info in PRODUCTS_DB.items():
        match = match_product_to_country(product_name, product_info, country_bands)
        if match:
            cat = product_info['category']
            result[cat].append(match)

    for cat in result:
        result[cat].sort(key=lambda x: (grade_order.get(x['adapt']['grade'], 9), -x['adapt']['pct']))

    return result


def build():
    # 1. 读取国家数据
    if not os.path.exists(DATA_FILE):
        print("❌ 找不到数据文件：" + DATA_FILE)
        return False

    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        raw_data = json.load(f)

    countries_raw = raw_data.get('countries', [])
    print(f"✅ 从 {os.path.basename(DATA_FILE)} 读取 {len(countries_raw)} 个国家")

    # 2. 构建国家数据库（按产品匹配）
    countries_db = []

    for country_info in countries_raw:
        name = country_info[0]
        bands = parse_bands_str(country_info[2]) if len(country_info) > 2 else {}

        record = {
            'name': name,
            'region': country_info[1] if len(country_info) > 1 else '',
            'bands': bands,
            'assessment': country_info[3] if len(country_info) > 3 else '',
            'product_matches': auto_match_country(bands)
        }
        countries_db.append(record)

    # 按国家名排序
    countries_db.sort(key=lambda x: x['name'])

    # 统计
    total = sum(
        len(c['product_matches']['共享产品']) + len(c['product_matches']['两轮产品'])
        for c in countries_db
    )
    print(f"📊 匹配完成：共 {total} 条国家-产品关联")

    # 3. 读取模板并替换占位符
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()

    modules_json = json.dumps(MODULES, ensure_ascii=False)
    products_json = json.dumps(PRODUCTS_DB, ensure_ascii=False)
    countries_json = json.dumps(countries_db, ensure_ascii=False)

    html_content = template.replace(
        '{{MODULES_PLACEHOLDER}}',
        'const MODULE_DB = ' + modules_json + ';'
    ).replace(
        '{{PRODUCTS_PLACEHOLDER}}',
        'const PRODUCTS_DB = ' + products_json + ';'
    ).replace(
        '{{COUNTRIES_PLACEHOLDER}}',
        'const COUNTRY_DB = ' + countries_json + ';'
    )

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f'\n✨ 构建完成！')
    print(f'   📄 输出: {OUTPUT_FILE}')
    print(f'   🌍 国家: {len(countries_db)} 个')
    print(f'   📦 模块: {len(MODULES)} 种')
    print(f'   🛠️  产品: {len(PRODUCTS_DB)} 种')
    print(f'   📏 大小: {os.path.getsize(OUTPUT_FILE) / 1024:.1f} KB')
    print(f'\n💡 双击 index.html 即可打开使用！')
    return True


if __name__ == '__main__':
    build()
    # 暂停，让用户看到结果（双击运行时不会一闪而过）
    input('\n按回车键退出...')

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
4G模块选型工具 - 一键构建脚本（v2.0 - 自动匹配架构）
用法：运行此脚本 → 读取 data.json + template.html → 生成 index.html（数据内嵌）

架构：
- data.json：只存国家数据（扁平列表），不再按模块分Sheet
- build.py：自动计算每个国家能被哪些模块支持
- MODULES dict：定义所有模块的频段、GSM支持等
更新数据：编辑 data.json → 重新运行此脚本即可
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
# 模块定义（在此处添加新模块）
# ============================================================
MODULES = {
    'CN': {
        'name': 'MC669-CN / LE270-CN / ML307N-DC-DL',
        'short_name': 'CN三合一',
        'bands': {'FDD': ['B1', 'B3', 'B5', 'B8'], 'TDD': ['B34', 'B38', 'B39', 'B40', 'B41']},
        'has_gsm': False,
        'priority': 100,
        'region': '亚太/中东/非洲/大洋洲',
        'note': '不支持GSM，适合亚太地区'
    },
    'L610-CN': {
        'name': 'L610-CN',
        'short_name': 'L610-CN',
        'bands': {'FDD': ['B1', 'B3', 'B5', 'B8'], 'TDD': ['B34', 'B38', 'B39', 'B40', 'B41']},
        'has_gsm': False,
        'priority': 95,
        'region': '亚太/中东/非洲/大洋洲',
        'note': '不支持GSM，适合亚太地区'
    },
    'EU': {
        'name': 'L610-EU (支持GSM) / LE270-EU (不支持GSM)',
        'short_name': 'EU版',
        'bands': {'FDD': ['B1', 'B3', 'B7', 'B8', 'B20']},
        'has_gsm': True,
        'priority': 90,
        'region': '欧洲/东欧/中东/非洲',
        'note': 'L610-EU额外支持GSM'
    },
    'LA': {
        'name': 'L610-LA',
        'short_name': 'LA拉美版',
        'bands': {'FDD': ['B1', 'B2', 'B4', 'B5', 'B8'], 'TDD': ['B38', 'B40', 'B41'], 'WCDMA': ['B1', 'B2', 'B4', 'B5']},
        'has_gsm': True,
        'priority': 85,
        'region': '北美/拉丁美洲',
        'note': '支持GSM四频+WCDMA，适合北美和拉美'
    },
    'GL': {
        'name': 'LE270-GL',
        'short_name': 'GL全球版',
        'bands': {'FDD': ['B1','B2','B3','B4','B5','B7','B8','B12','B13','B17','B20','B25','B26','B28','B66'],
                  'TDD': ['B34','B38','B39','B40','B41']},
        'has_gsm': False,
        'priority': 80,
        'region': '全球',
        'note': '全球频段覆盖最广，成本较高'
    }
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


def auto_match_modules(country_bands):
    """
    自动判断一个国家的4G频段可以被哪些模块支持。
    
    匹配规则：模块的 FDD+TDD 频段集合 必须包含 该国的 FDD+TDD 频段的至少一部分
    （即有交集就认为该模块可用，让前端排序逻辑决定优先级）
    """
    matched = {}
    
    # 收集国家所有LTE频段
    country_lte = set()
    for band in country_bands.get('FDD', []):
        country_lte.add(band)
    for band in country_bands.get('TDD', []):
        country_lte.add(band)
    
    # 如果国家没有任何LTE频段，跳过
    if not country_lte:
        return matched
    
    for module_key, module_info in MODULES.items():
        module_lte = set()
        for band in module_info['bands'].get('FDD', []):
            module_lte.add(band)
        for band in module_info['bands'].get('TDD', []):
            module_lte.add(band)
        
        # 如果模块和国家有至少一个共同频段，则该模块可用
        if country_lte & module_lte:  # 交集非空
            matched[module_key] = module_info
    
    return matched


def build():
    # 1. 读取数据源（新格式：扁平国家列表）
    if not os.path.exists(DATA_FILE):
        print("❌ 找不到数据文件：" + DATA_FILE)
        return False
    
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        raw_data = json.load(f)
    
    countries_raw = raw_data.get('countries', [])
    print(f"✅ 从 {os.path.basename(DATA_FILE)} 读取 {len(countries_raw)} 个国家")

    # 2. 构建国家数据库（自动匹配模块）
    countries_db = []
    
    for country_info in countries_raw:
        name = country_info[0]
        
        record = {
            'name': name,
            'region': country_info[1] if len(country_info) > 1 else '',
            'bands': parse_bands_str(country_info[2]) if len(country_info) > 2 else {},
            'assessment': country_info[3] if len(country_info) > 3 else '',
            'supported_by': [],
            'modules_available': {}
        }
        
        # ★ 核心改动：自动匹配模块，不再依赖手动分配
        matched = auto_match_modules(record['bands'])
        record['modules_available'] = matched
        record['supported_by'] = list(matched.keys())
        
        countries_db.append(record)

    # 按国家名排序
    countries_db.sort(key=lambda x: x['name'])

    # 3. 统计匹配情况
    total_matches = sum(len(c['supported_by']) for c in countries_db)
    print(f"📊 自动匹配完成：共 {total_matches} 条国家-模块关联")

    # 4. 读取模板并替换占位符
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()

    modules_json = json.dumps(MODULES, ensure_ascii=False)
    countries_json = json.dumps(countries_db, ensure_ascii=False)

    html_content = template.replace(
        '{{MODULES_PLACEHOLDER}}',
        'const MODULE_DB = ' + modules_json + ';'
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
    print(f'   📏 大小: {os.path.getsize(OUTPUT_FILE) / 1024:.1f} KB')
    print(f'\n💡 双击 index.html 即可打开使用！')


if __name__ == '__main__':
    build()

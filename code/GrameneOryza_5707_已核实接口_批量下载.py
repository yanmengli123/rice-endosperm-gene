#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GrameneOryza v8 Nipponbare published-function genes: download 5707 records.

Verified from user's GrameneOryza Network screenshot (2026-10-28):
  GET https://data.gramene.org/oryza_v8/search
  q=*:* AND ((capabilities:pubs))
  fq=taxon_id:(39947013)
  response.numFound=5707

Requires Internet access on the user's PC. This script has NOT downloaded the
live dataset in the ChatGPT runtime; it validates every result at runtime.
"""
from __future__ import annotations
import csv
import json
import sys
import time
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit('请先安装依赖：py -m pip install requests')

BASE_URL = 'https://data.gramene.org/oryza_v8/search'
QUERY = '*:* AND ((capabilities:pubs))'
FILTER_QUERY = 'taxon_id:(39947013)'
EXPECTED_TOTAL = 5707
TAXON_ID = '39947013'
PAGE_SIZE = 200
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "GrameneOryza"


def get_page(session, start, rows):
    params = {
        'q': QUERY,
        'fq': FILTER_QUERY,
        'start': start,
        'rows': rows,
        'sort': 'id asc',
        'wt': 'json',
    }
    for attempt in range(1, 5):
        try:
            resp = session.get(BASE_URL, params=params, timeout=(15, 90))
            resp.raise_for_status()
            data = resp.json()['response']
            return int(data['numFound']), list(data['docs'])
        except (requests.RequestException, KeyError, ValueError, TypeError) as exc:
            if attempt == 4:
                raise RuntimeError(f'第 {start} 条开始的分页请求失败：{exc}') from exc
            print(f'第 {start} 条请求失败，重试 {attempt}/3：{exc}')
            time.sleep(attempt * 2)


def safe_cell(value):
    if value is None:
        return ''
    if isinstance(value, (list, dict)):
        text = json.dumps(value, ensure_ascii=False, separators=(',', ':'))
    else:
        text = str(value)
    # Prevent spreadsheet formula evaluation when opening TSV in Excel.
    if text.startswith(('=', '+', '-', '@')):
        return "'" + text
    return text


def main():
    print('GrameneOryza v8 / Curated publications / Nipponbare')
    print('API: ', BASE_URL)
    print('Query:', QUERY, '| Filter:', FILTER_QUERY)
    session = requests.Session()
    session.headers.update({'Accept': 'application/json', 'User-Agent': 'Rice-endosperm-research-curated-gene-export/1.0'})
    all_docs = []
    try:
        num_found, first_docs = get_page(session, 0, PAGE_SIZE)
        print(f'返回基因总数: {num_found}; 预期: {EXPECTED_TOTAL}')
        if num_found != EXPECTED_TOTAL:
            sys.exit('网站数据条数与截图不一致，停止导出。请确认筛选条件或数据库是否更新。')
        all_docs.extend(first_docs)
        while len(all_docs) < num_found:
            start = len(all_docs)
            current_total, docs = get_page(session, start, min(PAGE_SIZE, num_found - start))
            if current_total != num_found or not docs:
                sys.exit(f'数据总数变化或第 {start} 条分页返回空数据，停止导出。')
            all_docs.extend(docs)
            print(f'已下载 {len(all_docs)}/{num_found}')
            time.sleep(0.25)
    finally:
        session.close()

    if len(all_docs) != EXPECTED_TOTAL:
        sys.exit(f'记录数量不正确：{len(all_docs)}，未导出文件。')
    ids = []
    for pos, gene in enumerate(all_docs, start=1):
        if 'id' not in gene:
            sys.exit(f'第 {pos} 条缺少基因id，未导出文件。')
        if str(gene.get('taxon_id')) != TAXON_ID:
            sys.exit(f'第 {pos} 条taxon_id异常，未导出文件。')
        if 'pubs' not in gene.get('capabilities', []):
            sys.exit(f'第 {pos} 条缺少pubs功能标签，未导出文件。')
        ids.append(str(gene['id']))
    if len(set(ids)) != EXPECTED_TOTAL:
        sys.exit(f'基因ID去重后仅{len(set(ids))}个，未导出文件。')

    tsv_path = OUTPUT_DIR / 'GrameneOryza_Curated_Nipponbare_5707.tsv'
    raw_path = OUTPUT_DIR / 'GrameneOryza_Curated_Nipponbare_5707_raw.json'
    fields = ['Gene_ID', 'Gene_Symbol', 'Description', 'Synonyms', 'Model_Rep_ID', 'Biotype', 'Taxon_ID', 'Source']
    with tsv_path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fields, delimiter='\t')
        writer.writeheader()
        for g in all_docs:
            writer.writerow({
                'Gene_ID': safe_cell(g.get('id')),
                'Gene_Symbol': safe_cell(g.get('name')),
                'Description': safe_cell(g.get('description')),
                'Synonyms': safe_cell(g.get('synonyms')),
                'Model_Rep_ID': safe_cell(g.get('model_rep_id')),
                'Biotype': safe_cell(g.get('biotype')),
                'Taxon_ID': safe_cell(g.get('taxon_id')),
                'Source': 'GrameneOryza v8; curated publication',
            })
    raw_path.write_text(json.dumps(all_docs, ensure_ascii=False, indent=2), encoding='utf-8')

    print('\n下载成功，验证通过：5707条、5707个不重复基因、全部为日本晴且带pubs标签。')
    print('可用Excel打开：', tsv_path)
    print('原始完整JSON：', raw_path)
    print('注意：这5707条是文献注释基因，不等于5707条胚乳功能基因。')


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        sys.exit('用户中止。')
    except Exception as e:
        sys.exit(f'运行失败：{e}')

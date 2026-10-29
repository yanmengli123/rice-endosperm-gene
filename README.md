# rice-endosperm-gene

水稻胚乳功能基因多数据库整合数据（Oryza sativa）。

本仓库保存 7 个公开数据库的最终整理数据、序列、校验清单与抓取脚本，用于胚乳/淀粉/贮藏物质相关基因的检索、文献核对与跨库 ID 整合。

## 数据总览

| 目录 | 内容 | 规模 |
|---|---|---|
| `RAP-DB/` | IRGSP-1.0 注释、RAP-MSU 映射、Curated genes、Agri. genes | 40,249 转录本 / 36,061 位点；Curated 7,031；Agri 1,093 |
| `funPlantGenes/` | funPlantGenes（HENAU）水稻基因页面结构化数据与序列 | 7,407 个基因；24,972 条序列 |
| `GrameneOryza/` | GrameneOryza v8 文献功能注释基因 | 5,707 个基因；7,786 条基因–文献关联 |
| `FunRiceGenes/` | FunRiceGenes 官网原始表 + UTF-8 可读副本 | 4,567 个基因；45,929 条关键词；8,630 篇文献 |
| `Oryzabase/` | 6 类性状/QTL 基因表 | 2,577 条记录 |
| `KitBase/` | Kitaake 突变体补充表 | 1,504 个突变体 |
| `TRIM/` | 突变体表型与敲除补充表 | S3 96,748 行；S4 25,898 行 |
| `code/` | 抓取/处理脚本 | 5 个脚本 |

跨库按 RAP 位点合并后共 **7,745 个位点**。

## 目录说明

```text
.
├─ RAP-DB/
│  ├─ IRGSP-1.0_representative_annotation_2026-02-05.tsv
│  ├─ RAP-MSU_2026-02-05.txt.gz
│  ├─ Curated genes/        # xlsx / csv / json / references / endosperm hits
│  └─ Agri. genes/          # xlsx / csv / json / references / endosperm hits
├─ funPlantGenes/           # xlsx / csv / FASTA / official_fasta / official_records.jsonl.gz
├─ GrameneOryza/            # enriched xlsx / csv / raw.json / full.json / pubmed_metadata.json
├─ FunRiceGenes/            # 官网原始 3 表 + utf8/ 可读副本
├─ Oryzabase/               # 6 个 TSV
├─ KitBase/                 # 4 个 XLSX
├─ TRIM/                    # 3 个 DOCX + 2 个 XLSX
├─ code/                    # 爬虫与导出脚本
├─ README_DATA.md           # 数据使用说明（推荐先读）
├─ DATA_AUDIT.md            # 数据核验与清理结果
├─ FINAL_FILES_SHA256.json  # 最终文件 SHA-256 清单
├─ FINAL_VALIDATION.json    # 同步/清理验证结果
└─ OFFICIAL_SOURCE_MANIFEST.json  # 官网来源、时间与响应哈希
```

## 快速开始

1. 先读 `README_DATA.md`，了解每类数据“官网原样版本”和“便于分析版本”的区别。
2. 做跨库整合时，以 RAP Locus ID 为主键；MSU ID 通过 `RAP-MSU_2026-02-05.txt.gz` 转换。
3. 胚乳候选基因可从各库的 `*endosperm_hits*` 工作表/文件开始，再人工复核。
4. 做序列分析时，优先使用 `funPlantGenes/sequences_*.fa`（唯一 `record_id`）；需要与官网原始标题对应时使用 `funPlantGenes/official_fasta/`。
5. 完整性和来源核验见 `FINAL_FILES_SHA256.json` 与 `OFFICIAL_SOURCE_MANIFEST.json`。

## 已知源站异常（非抓取错误）

- funPlantGenes：140 个基因页面共用 `LOC_Os12g44380` 的序列；59 个基因存在源标题重复；241 条 CDS 长度不是 3 的倍数；4 条 DNA 栏目含蛋白字母。逐条见 `funPlantGenes/SOURCE_ISSUES.csv`。
- KitBase：原始文件中有 8 个 `#N/A`。
- FunRiceGenes：官方文件本身有 1 个 Unicode 替换字符、空映射和多重映射。
- Gramene：2 个旧 ID 不在当前 RAP 注释；147 个基因仅有 DOI/整条引用。

“与官网一致”指忠实保存源站内容，不代表源站注释均经实验确认。

## 数据来源

- RAP-DB: https://rapdb.dna.naro.go.jp/
- funPlantGenes: https://funplantgenes.henau.edu.cn/
- Gramene: https://www.gramene.org/
- FunRiceGenes: https://funricegenes.github.io/
- Oryzabase: https://shigen.nig.ac.jp/rice/oryzabase/
- KitBase: https://kitbase.ucdavis.edu/
- TRIM: 论文补充材料（Li et al., Plant Cell 2017, PMC5570018）

数据版权归各源数据库所有，使用请遵循各源站条款。

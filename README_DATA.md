# 数据使用说明

本目录已按 2026-10-28 的官网内容重新同步。详细结果见 `DATA_AUDIT.md`，官方下载地址、获取时间和 SHA-256 见 `OFFICIAL_SOURCE_MANIFEST.json`。

## 从哪里读取数据

`RAP-DB`、`Oryzabase`、`GrameneOryza`、`FunRiceGenes`、`funPlantGenes`、`KitBase`、`TRIM` 七个目录保存最终数据。爬虫代码统一在 `code/`，其说明见 `code/README.md`。

根据清理要求，网页缓存、备份、原始分页临时响应、下载日志、调试文件和重复压缩包已移除。最终表格、完整结构化记录、原样 FASTA、唯一 ID FASTA 及质量标记继续保留。来源 URL、获取时间、官网内容校验和见 `OFFICIAL_SOURCE_MANIFEST.json`；当前最终文件校验和见 `FINAL_FILES_SHA256.json`；同步验证结果见 `FINAL_VALIDATION.json`。

本次移除使用 Windows 回收站，项目目录不再保留这些副本；如需恢复，可从系统回收站找回。

## 官网原样数据与便于使用的导出

| 数据 | 官网原样内容 | 便于读取/分析的版本 |
|---|---|---|
| FunRiceGenes | 主目录三个 `.table.txt`，与官网下载字节一致，保留原始混合编码 | `FunRiceGenes/utf8/` 同名文件；UTF-8 BOM，编码转换清单可逆 |
| funPlantGenes 页面数据 | `official_records.jsonl.gz` 保存全部页面提取的字段、关联数据与序列；网页 HTML 缓存已删除 | CSV、Excel、`SOURCE_PAGES.json` 保存结构化视图与源页映射 |
| funPlantGenes 序列 | `funPlantGenes/official_fasta/*.fa`：保留源页序列标题和序列内容 | `funPlantGenes/sequences_*.fa`：序列内容相同，使用全局唯一的 `record_id` |
| RAP、KitBase、Oryzabase、TRIM | 最终 TSV、Excel、DOCX、JSON；RAP-MSU 只有压缩版，故保留该 `.gz` | 其余已解压且内容重复的压缩包已删除 |
| GrameneOryza | 主目录 `_raw.json` / `_full.json` 是官方分页 docs 合并；分页临时文件已删除 | Excel/CSV 为派生表 |
| PubMed 文献元数据 | `pubmed_metadata.json` 保存最终规范化元数据，包括完整作者列表；HTTP 批次临时响应已删除 | 文献 CSV/Excel；DOI 取明确的 DOI 标识 |

`FunPlantGenes_Oryza_sativa_sequences.csv` 和 Excel 的 `Sequences` 工作表保留 `fasta_header`（原始标题），并新增 `record_id` 与 `source_url`。主目录 FASTA 空格前的标识对应 `record_id`。唯一标识只解决关联和重名问题，不会替换序列内容或确认其生物学归属。

## 必须保留的官网自身问题

根据“与官网保存一致”的要求，以下内容保留原样，没有推测替换：

- funPlantGenes 的 140 个非 `loc_os12g44380` 页面共用该基因的序列、4 条核苷酸栏目中的蛋白序列、241 条非三倍数长度 CDS，以及源标题重复。
- KitBase 原文件中的 8 个 `#N/A`。
- FunRiceGenes 原文件中本身已经存在的 1 个 Unicode 替换字符 `�`，以及空映射、多重映射。

funPlantGenes 逐记录问题见 `funPlantGenes/SOURCE_ISSUES.csv`。这些记录可以用于回查官网；用于生物学序列分析时，应按问题类型审查。原始序列与唯一 ID 版本都没有删除这些记录。

官网原样一致是本次交付标准；它不等于源数据库的每一条科学注释都已被实验确认。

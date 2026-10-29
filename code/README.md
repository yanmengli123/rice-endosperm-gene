# 爬虫代码

代码统一保存在本目录；输出路径相对于工作区根目录定位，不受运行时当前目录影响。

| 脚本 | 数据输出目录 |
|---|---|
| `scrape_funplantgenes.py` | `../funPlantGenes/` |
| `scrape_curated_genes.py` | `../RAP-DB/Curated genes/` |
| `scrape_agri_genes.py` | `../RAP-DB/Agri. genes/` |
| `GrameneOryza_5707_已核实接口_批量下载.py` | `../GrameneOryza/` |
| `build_gramene_excel.py` | `../GrameneOryza/` |

依赖见 `requirements.txt`，建议 Python 3.11 或更新版本。

在工作区根目录运行示例：

```powershell
python -m pip install -r code/requirements.txt
python code/scrape_funplantgenes.py
python code/scrape_curated_genes.py --use-local
python code/scrape_agri_genes.py --use-local
python code/build_gramene_excel.py
```

本次清理已删除网页缓存和下载中间文件。重新运行 funPlantGenes 爬虫会重新下载目录页、基因页并生成运行缓存；清理后不能直接使用依赖已有网页缓存的 `--skip-download` 模式。RAP 的 `--use-local` 继续使用保留的最终官方 JSON。

下载脚本会更新各数据目录的对应输出。官网异常记录的含义及原样 FASTA/唯一 ID FASTA 的选择见根目录 `README_DATA.md`。

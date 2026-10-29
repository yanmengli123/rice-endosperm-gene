# FunRiceGenes 官网原文件

主目录中的 `geneInfo.table.txt`、`geneKeyword.table.txt`、`reference.table.txt` 均为本次官网下载的原始字节文件，分别含 4,567、45,929、8,630 条数据记录。

官网原文件包含混合编码。直接用严格 UTF-8 读取 keyword/reference 可能报编码错误，这不是下载不完整。需要 UTF-8 时读取 `utf8/` 中的同名副本。

UTF-8 副本保留有效 UTF-8 片段，对非法字节以 Windows-1252（未定义字节用 Latin-1）逐字节解码。每个转换位置和原字节记录于相邻 `.encoding.json` 文件，已验证可还原官网原文件。没有以 `errors="replace"` 丢弃内容。

reference 中保留 1 个官网原字节已经编码成 U+FFFD 的 `�`；没有猜测替换该源内容。此前转换额外产生的替换字符已通过重新获取原文件消除。

历史损坏文件和全部备份已按要求清理。本目录仅保留最终官网文件、UTF-8 阅读副本、可逆编码转换记录和本说明。

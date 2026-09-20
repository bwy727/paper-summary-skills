---
name: summary_answering
description: 回答我的问题。把用户的提问规范化、分解，按语义和关键词检索文献分组表 keyword_summary.xlsx 与文献总结表 paper_archiving.xlsx，整理出可回答该问题的文献名单并生成 markdown（.md），再在文献总结表中筛选呈现相关行。触发示例：「回答我的问题：XXX」「哪些文献讲了 XXX」「关于 XXX 的文献有哪些」。
---

# summary_answering — 文献问答

针对用户的问题，整理出可参考的文献名单。

## 脚本
`scripts/answering.py`（与本 SKILL.md 同级的 scripts 目录，用其绝对路径调用 `python`）。

## 流程
1. **规范化并分解问题**：把用户的提问（如「为什么 prior 很重要」）整理为核心子问题，并提炼检索用的关键词（语义关键词＋字面关键词）。
2. **检测并读取现有表格**（结果写入工作文件）：
   ```
   python <脚本绝对路径> dump --root D:\AIacademic\ReadPaper\xlsx_summary --json-out _work\answer_dump.json
   ```
   stdout 打印摘要 `{"has_keyword","has_archiving","json_out"}`；完整数据在 `_work\answer_dump.json`：
   `{"has_keyword":bool,"has_archiving":bool, "keyword":{...}, "archiving":{...}}`。用 Read 工具读取该文件。
3. **按可用性分支**：
   - **两表都在** → 同时检索两者。
   - **只有其一** → 用现有的检索，并在反馈中**明确提示缺失的另一份表**。
   - **两表都没有** → **停下来**，先向用户确认「确实没有 `keyword_summary.xlsx` 和 `paper_archiving.xlsx` 吗？」，并**申请许可**后再直接检索全部 markdown（`.md`）：
     ```
     python <脚本绝对路径> scanmd --folder D:\AIacademic\ReadPaper\notes_done --json-out _work\answer_scanmd.json
     ```
     未获许可不要自行扫描。
4. **筛选文献**：依据问题语义与关键词，从可用数据中整理出两类文献，并为每篇写一两句说明：
   - `direct`：可直接回答该问题的文献。`note` 用一两句说明**这篇文献如何回答该问题**。
   - `related`：与该话题/关键词相关的文献。`note` 用一两句说明**这篇文献在哪方面与该问题/话题相关**。
   - 说明依据现有数据（文献总结表的摘要/结果/思路，或文献分组表的关键词）撰写，文字精简、具体，避免空话。
5. **生成文献名单 markdown**：写入临时 `names.json`，每个条目为 `{"name":"xx","note":"一两句说明"}`（`name` 为论文标题，与文献总结表口径一致）：
   ```json
   {"question":"...",
    "direct":[{"name":"a","note":"该文用……实验直接证明……，回答了……"}],
    "related":[{"name":"b","note":"该文涉及……，与问题中的……话题相关"}]}
   ```
   再运行：
   ```
   python <脚本绝对路径> genmd --out <问题关键词>.md --data names.json
   ```
   生成的 markdown 中，每篇文献名下会附上对应说明（direct 显示「如何回答：…」，related 显示「相关之处：…」）。输出文件名默认取问题关键词（用户另有指定则用指定名），生成在输出文件夹 `D:\AIacademic\ReadPaper\xlsx_summary` 内。
6. **在文献总结表中筛选呈现相关行**：
   - 若 `has_archiving` 为真：
     ```
     python <脚本绝对路径> filter --archiving paper_archiving.xlsx --names names.json --out <问题关键词>_文献总结筛选.xlsx
     ```
     `--archiving` 与 `--out` 只给文件名时会自动落在输出文件夹 `D:\AIacademic\ReadPaper\xlsx_summary` 内。
   - 若没有文献总结表：向用户确认是否需要先生成 `paper_archiving.xlsx`（可调度 summary_archiving）。
7. 向用户报告：生成的 markdown 名单、筛选表路径、命中文献数；如有缺失的表也一并提示。删除临时 json（`names.json`、`_work\answer_*.json`）。

## 注意
- 文献名去重/匹配以「论文标题」（md 文件名去掉扩展名与开头的 `[note] ` 前缀、去首尾空格）为准。
- 不擅自猜测用户未提供的文件；缺表时按上面的确认流程处理。
- 路径约定：输入文件夹（无表时扫描的 md 来源）为 `D:\AIacademic\ReadPaper\notes_done`；输出文件夹为 `D:\AIacademic\ReadPaper\xlsx_summary`，`keyword_summary.xlsx`、`paper_archiving.xlsx`、`<问题关键词>.md`、`<问题关键词>_文献总结筛选.xlsx` 均在此。

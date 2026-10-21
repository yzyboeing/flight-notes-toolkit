# 已停用的工具（2026-10-21 规则重组 A7）

这些脚本服务于已作废的规则或一次性的迁移，不再被 `sync.sh` 或其他工具调用；留作历史参考，**不要在日常工作中使用**。

| 脚本 | 原用途 | 停用原因 |
|---|---|---|
| abbr_index.py | 生成附录「缩写与术语」 | 1019R6 撤销附录，缩写写在表下注 |
| apply_marks.py / remark_body.py / remark_quickref.py | 早期批量改颜色标记 | 一次性迁移，已完成 |
| audit_cells.py / audit_rules.py | SD-119 / 120 期间的排版审查 | 已由 check_layout / check_struct 取代 |
| blockify.py | 块索引（A-1）结构 | SD-199 / 201 全书取消块索引 |
| book_break.py | 全书预排版分页 | 由 fit_fix 的 P: 记录取代 |
| book_pdf.py / quickref_pdf.py / ipad_build.py | 竖版 / iPad 版 | 只出横版（SD-71） |
| docx2md.py / restructure.py | 早期从 Word 迁到 Markdown 源 | 一次性迁移，已完成 |
| muse_import_check.py | 噜噜（Muse）变更单导入 | 改为噜噜交 git 补丁（SD-181） |
| single_extra.py | 单册追加正文节（SD-116） | 速查改为独立成册（SD-146） |

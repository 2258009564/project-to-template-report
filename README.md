# project-to-template-report

面向 Codex 的 Word 模板报告技能：保留原模板的封面与版式，从真实材料编写正文，先审查文案，再通过 Microsoft Word 批量排版并核验。

## 能力

- 读取 Word 实际生效的格式，按语义角色生成正文、表格、图片和真实目录。
- 写入前单独审查语句、逻辑、事实与术语，修复后再写入 Word；完成后仍需逐页检查版式。
- 从技能目录外的本地档案复用姓名、学号等已知字段；获得用户持续授权后，将明确反馈记录为有适用范围的规则，在后续任务中读取。
- 区分已实现、已测试与规划内容，不凭空补全成果、日期、教师或组员信息。

## 安装与使用

将本仓库文件复制到 `$CODEX_HOME/skills/project-to-template-report`；未设置 `CODEX_HOME` 时使用用户目录下的 `.codex/skills/project-to-template-report`。重新开启会话以发现技能，然后提供模板和事实材料，例如：

> 使用 $project-to-template-report，按这个模板编写报告，保留封面和格式。先审稿，再生成 Word。

需要 Windows、Microsoft Word 和 PowerShell。Python 辅助工具用于 OOXML 结构核验；需要图片渲染时须另备可用的 DOCX 渲染工具，本仓库不包含 `render_docx.py`。可以使用 Codex documents 技能中的渲染器，或以 Word 导出的 PDF 渲染页面；最终分页以 Word 为准。项目代码取证及图示流程仅在任务需要时启用。

## 本地记忆与隐私

记忆位于 `<CODEX_HOME>/memories/extensions/ad_hoc/notes/`，没有设置环境变量时以 `~/.codex` 为根。个人档案用 `word-report-profile-v1` 标记，反馈学习用 `word-report-learning-v1` 标记。首次使用不会获得作者的任何个人档案；用户明确授权后才能持续记录反馈，授权本身也仅记在本地。

记录包含来源、日期、适用范围及核验状态。当前明确指令优先于旧记录，单份模板要求不自动升级为长期偏好。它是后续任务读取本地笔记的机制，不是训练模型，也不会在会话结束后自行运行。发布时不要复制记忆目录、实际报告或包含个人资料的日志；`.gitignore` 只能辅助，不能代替审查。

## 验证

在仓库根目录运行：

```powershell
python scripts/check_package.py
./scripts/test_word_report.ps1 -WorkDir "$env:TEMP/word-report-synthetic-test"
```

第一个检查发布文件类型、本地 Markdown 引用、Python 语法和常见泄露特征；不保证发现所有隐私信息。第二个在独立的隐藏 Word 实例中创建合成模板，验证格式提取、缓存、目录、混合格式拒绝及原文件保护。它不等于对真实报告的内容质量或跨会话记忆效果验证。最终文档仍须按技能执行审稿和逐页检查。

## 边界

内置生成器替换唯一指定段落到文末的内容，并保留边界前的区域；任意局部改写或格式迁移需要另写受限适配逻辑。请先明确可编辑范围，保留原件。工具不会把任意模板自动识别为可安全替换区域。

本项目以 MIT 许可证开放，见 [LICENSE](LICENSE)。

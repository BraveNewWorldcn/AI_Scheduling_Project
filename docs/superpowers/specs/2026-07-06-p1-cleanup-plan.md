# P1 代码质量与文件归档 — 执行计划 v2（终版）

**日期**: 2026-07-06
**状态**: 已评审，待执行
**评审**: code-review + implementation-risk 双审通过
**前置**: P0 已完成
**目标**: 清理调试日志、归档未跟踪文件

---

## 变更概要

| Commit | 内容 | 文件数 |
|--------|------|--------|
| Commit 1 | 删除 scheduler_api.py 中 4 处 DEBUG print | 1 |
| Commit 2 | 归档 23 个未跟踪文件 + 提交 2 个计划文档 | 25 |

---

## Commit 1：清理 DEBUG print

### 现状

4 处全在 `scheduler_api.py`，均经 reviewer 确认为独立行、无副作用、不在 try/except 内：

| 行号 | 代码 |
|------|------|
| 302 | `print(f"[DEBUG] 读取表 {table_id} 的列名: {list(df.columns)}")` |
| 1324 | `print(f"[DEBUG] billing_index 共 {len(billing_index)} 条规则...")` |
| 1332 | `print(f"[DEBUG] 已修复 {mask.sum()} 条消防远程控制设备费的规格...")` |
| 1511 | `print(f"[DEBUG] 备选匹配：产品=...")` |

**注意**：删除第 1324 行时，保留其后第 1325 行的 `# 修复消防远程控制设备费...` 注释。

### 执行

```bash
# 手动删除 4 行后：
git add scheduler_api.py
git commit -m "chore: 移除 scheduler_api 中的调试 print 语句"
git push origin main
```

---

## Commit 2：归档未跟踪文件

### 现状：28 个未跟踪文件（27 条目 + 1 目录含 2 文件）

**23 个需移动的文件**：

| 来源 | 目标 | 类别 |
|------|------|------|
| `boss_card_proposal_v3.md` | `docs/proposals/` | 设计提案 |
| `boss_card_review_summary.md` | `docs/proposals/` | 设计评审 |
| `confirm_ui_proposal.md` | `docs/proposals/` | 设计提案 |
| `design_proposal_fs_cards.md` | `docs/proposals/` | 设计提案 |
| `design_proposal_v2.md` | `docs/proposals/` | 设计提案 |
| `design_review_summary.md` | `docs/proposals/` | 设计评审 |
| `design_review_summary_v2.md` | `docs/proposals/` | 设计评审 |
| `notification_flow_proposal.md` | `docs/proposals/` | 设计提案 |
| `notification_flow_proposal_v2.md` | `docs/proposals/` | 设计提案 |
| `notification_flow_review_summary.md` | `docs/proposals/` | 设计评审 |
| `notification_flow_v2_review.md` | `docs/proposals/` | 设计评审 |
| `面试话术_排单引擎.md` | `docs/` | 文档 |
| `import_all_v3.py` | `scripts/` | 导入脚本 |
| `import_v2.py` | `scripts/` | 导入脚本 |
| `import_demo_standalone.py` | `scripts/` | 导入脚本 |
| `gen_demo_v2.py` | `scripts/` | 演示生成 |
| `confirm.html` | `static/` | Web 页面 |
| `index.html` | `static/` | Web 页面 |
| `测试集.xlsx` | `test_data/` | 测试数据 |
| `测试集_演示_demo.xlsx` | `test_data/` | 测试数据 |
| `测试集_演示_optimized.xlsx` | `test_data/` | 测试数据 |
| `测试集_演示_v2.xlsx` | `test_data/` | 测试数据 |
| `合同明细清单_demo.xls` | `test_data/` | 测试数据 |

**2 个原地提交的文件**（已在正确位置）：

| 文件 | 说明 |
|------|------|
| `docs/superpowers/specs/2026-07-06-p0-cleanup-plan.md` | P0 执行计划 |
| `docs/superpowers/specs/2026-07-06-p1-cleanup-plan.md` | P1 执行计划（本文件） |

**不处理（留 P2）**：`start.sh`、`start_daemon.sh`、`com.ai-scheduling.scheduler.plist`

### 已知问题（不阻塞 P1）

4 个脚本（`gen_demo_v2.py`、`import_all_v3.py`、`import_v2.py`、`import_demo_standalone.py`）含有硬编码的 macOS 绝对路径（`/Users/xiaowang/...`），在 Windows 上无法运行。移动操作不会恶化此问题。留 P2 统一修改为相对路径。

### 执行

使用 **Git Bash**（中文文件名在 PowerShell 下可能编码异常）：

```bash
# 创建目标目录
mkdir -p docs/proposals scripts static test_data

# 移动 23 个文件
git mv boss_card_proposal_v3.md docs/proposals/
git mv boss_card_review_summary.md docs/proposals/
git mv confirm_ui_proposal.md docs/proposals/
git mv design_proposal_fs_cards.md docs/proposals/
git mv design_proposal_v2.md docs/proposals/
git mv design_review_summary.md docs/proposals/
git mv design_review_summary_v2.md docs/proposals/
git mv notification_flow_proposal.md docs/proposals/
git mv notification_flow_proposal_v2.md docs/proposals/
git mv notification_flow_review_summary.md docs/proposals/
git mv notification_flow_v2_review.md docs/proposals/
git mv 面试话术_排单引擎.md docs/
git mv import_all_v3.py scripts/
git mv import_v2.py scripts/
git mv import_demo_standalone.py scripts/
git mv gen_demo_v2.py scripts/
git mv confirm.html static/
git mv index.html static/
git mv 测试集.xlsx test_data/
git mv 测试集_演示_demo.xlsx test_data/
git mv 测试集_演示_optimized.xlsx test_data/
git mv 测试集_演示_v2.xlsx test_data/
git mv 合同明细清单_demo.xls test_data/

# 提交 2 个计划文档（已在正确位置）
git add docs/superpowers/specs/2026-07-06-p0-cleanup-plan.md docs/superpowers/specs/2026-07-06-p1-cleanup-plan.md

# 提交
git commit -m "$(cat <<'EOF'
chore: 归档未跟踪文件到对应目录

- docs/proposals/ ← 11 份设计提案/评审文档
- docs/ ← 面试话术
- scripts/ ← 4 个导入/演示脚本
- static/ ← 2 个 Web 页面 (confirm.html, index.html)
- test_data/ ← 5 个测试数据集
- docs/superpowers/specs/ ← P0/P1 执行计划

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"

git push origin main
```

### 提交后目录结构

```
├── docs/
│   ├── proposals/          ← 11 份设计提案/评审
│   ├── specs/              ← P0/P1 执行计划（已提交）
│   └── 面试话术_排单引擎.md
├── scripts/                ← 4 个导入/演示脚本
├── static/                 ← 2 个 Web 页面
├── test_data/              ← 5 个测试数据集
├── start.sh                ← 未处理（P2）
├── start_daemon.sh         ← 未处理（P2）
└── com.ai-scheduling.scheduler.plist ← 未处理（P2）
```

---

## 验证清单

**Commit 1 后**:
- [ ] `grep -n '\[DEBUG\]' scheduler_api.py` 返回空
- [ ] `grep -rn '\[DEBUG\]' *.py` 返回空（全局确认）
- [ ] 第 1325 行注释 `# 修复消防远程控制设备费的规格...` 保留

**Commit 2 后**:
- [ ] `git status` 只剩 3 个 P2 未跟踪文件（start.sh、start_daemon.sh、.plist）
- [ ] `ls docs/proposals/` 有 11 个文件
- [ ] `ls scripts/` 有 4 个文件
- [ ] `ls static/` 有 2 个文件
- [ ] `ls test_data/` 有 5 个文件
- [ ] 根目录不再有散落的 .md/.py/.html/.xlsx/.xls 文件

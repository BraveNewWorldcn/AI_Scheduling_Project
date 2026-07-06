# P0 最小安全提交 — 执行计划 v3（终版）

**日期**: 2026-07-06
**状态**: 已评审，待执行
**评审**: code-review + implementation-risk 双审通过
**目标**: 止血 — 提交所有未提交的核心代码，清理误追踪文件，推送远程，统一换行符

---

## 背景

AI排单调度引擎项目经 Mac/Win 合并后，工作区积压大量未提交改动：

| 问题 | 数量 |
|------|------|
| 关键新文件未跟踪（shared.py） | 1 |
| 源文件有未提交修改 | 5 个，+1313/-670 行 |
| `__pycache__/` 和 `.vs/` 已从磁盘删除但索引仍追踪 | 38 个文件 |
| cache/ 下运行时产物被追踪 | 15 个文件 |
| Office 临时锁定文件被追踪 | 1 个（~$测试.xlsx） |
| 本地 commit 未推送 | 11 个 |
| CRLF/LF 换行符混用 | 全仓库 |

## 影响范围

- **阻断**: shared.py 缺失导致 checkout 后 ImportError，项目无法运行
- **数据风险**: 11 个未推送 commit + 工作区改动仅存在于本地硬盘
- **噪音**: 54 个已删除/运行时文件持续出现在 `git status` 中

---

## 执行步骤

### 第1步：一次性提交所有清理和核心代码（1个 commit）

将所有内容合并为**一个 commit**，确保没有中间断裂状态。

**暂存清单**:

| 文件/目录 | 操作 | 说明 |
|-----------|------|------|
| `shared.py` | 新增 | 列名映射唯一真源 + audit_schedule_results()（349行） |
| `ai_daily_agent.py` | 修改 | 采购/计划员通知卡片重构（+356行） |
| `scheduler_api.py` | 修改 | 排单引擎改动（+590行） |
| `import_orders.py` | 修改 | 订单导入改动（+201行） |
| `sf_shipping.py` | 修改 | 顺丰物流改动（+116行） |
| `.env.example` | 修改 | 配置项更新（+10行） |
| `.gitignore` | 修改 | 追加 `~$*` 和 `.vs/` 忽略规则 |
| `cache/` 下 15 个文件 | 停止追踪 | `git rm --cached -r cache/`，文件保留本地 |
| `__pycache__/` 下所有文件 | 停止追踪 | `git rm --cached -r __pycache__/` |
| `.vs/` 下 9 个文件 | 停止追踪 | `git rm --cached -r .vs/` |
| `~$测试.xlsx` | 停止追踪 | `git rm --cached`，Office 临时文件 |

**不暂存**:
- `测试.xlsx` — 已追踪的测试数据，用 `git checkout` 还原到 HEAD 版本
- 27 个未跟踪文件（设计提案、导入脚本变体、演示数据等）— P1 处理

**执行命令**（需要 bash，Windows 上可用 Git Bash）:

```bash
# 0) 推送前安全检查：确认远程没有新提交
git fetch origin main
echo "=== 远程领先本地的 commit（应为空）==="
git log --oneline HEAD..origin/main

# 1) 还原测试文件到 HEAD 版本（不提交其修改）
git checkout -- "测试.xlsx"

# 2) 暂存 6 个源代码文件
git add shared.py ai_daily_agent.py scheduler_api.py import_orders.py sf_shipping.py .env.example

# 3) 更新 .gitignore（追加 ~$* 和 .vs/ 规则）
# 检查是否已有这些规则，没有则追加
grep -q '~$*' .gitignore 2>/dev/null || echo '~$*' >> .gitignore
grep -q '\.vs/' .gitignore 2>/dev/null || echo '.vs/' >> .gitignore
git add .gitignore

# 4) 停止追踪所有运行时产物（--cached 只删索引，不删本地文件）
git rm --cached -r cache/
git rm --cached -r __pycache__/
git rm --cached -r .vs/ 2>/dev/null
git rm --cached "~$测试.xlsx" 2>/dev/null

# 5) 提交前验证
echo "=== 暂存区文件清单 ==="
git diff --cached --name-only

echo "=== shared.py 导入验证 ==="
python -c "from shared import OA_COLUMN_MAP, MAIN_FIELD_MAP, ITEMS_FIELD_MAP, INV_COLUMN_MAP, MAIN_COLUMNS, ITEMS_COLUMNS, INV_COLUMNS, audit_schedule_results; print('shared.py OK')"

echo "=== 语法检查 ==="
python -c "import ast; ast.parse(open('scheduler_api.py', encoding='utf-8').read()); print('scheduler_api.py OK')"
python -c "import ast; ast.parse(open('import_orders.py', encoding='utf-8').read()); print('import_orders.py OK')"
python -c "import ast; ast.parse(open('ai_daily_agent.py', encoding='utf-8').read()); print('ai_daily_agent.py OK')"
python -c "import ast; ast.parse(open('sf_shipping.py', encoding='utf-8').read()); print('sf_shipping.py OK')"

# 6) 提交
git commit -m "$(cat <<'EOF'
chore: 合并后清理 — 提交核心代码、停止追踪构建产物和缓存

- 新增 shared.py: 列名映射唯一真源 + 排单审计函数
- 更新 ai_daily_agent/scheduler_api/import_orders/sf_shipping
- 停止追踪 cache/、__pycache__/、.vs/（已在 .gitignore 中）
- .gitignore 新增 ~$*、.vs/ 忽略规则
- 更新 .env.example 配置项

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

**PowerShell 用户备选**（将 heredoc 替换为单行 `-m`，或直接用 Git Bash 执行上面的命令）:

```powershell
git commit -m 'chore: 合并后清理 — 提交核心代码、停止追踪构建产物和缓存

- 新增 shared.py: 列名映射唯一真源 + 排单审计函数
- 更新 ai_daily_agent/scheduler_api/import_orders/sf_shipping
- 停止追踪 cache/、__pycache__/、.vs/（已在 .gitignore 中）
- .gitignore 新增 ~$*、.vs/ 忽略规则
- 更新 .env.example 配置项

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>'
```

---

### 第2步：推送全部 commit

```bash
# 1) 确认推送范围（应为 12 个 commit）
git log --oneline origin/main..HEAD

# 2) 试运行
git push origin main --dry-run

# 3) 正式推送
git push origin main
```

**如果 push 被拒绝（non-fast-forward）**：说明远程有新提交。执行：
```bash
git pull --rebase origin main
# 解决冲突（如有）后重新推送
git push origin main
```

---

### 第3步：统一换行符 + 保护 git blame

创建 `.gitattributes`（仓库根目录）:

```
* text=auto
*.py text eol=lf
*.sh text eol=lf
*.bat text eol=crlf
*.ps1 text eol=crlf
*.md text eol=lf
*.csv text eol=lf
*.json text eol=lf
*.xlsx binary
*.xls binary
*.pkl binary
*.db binary
~$* binary
```

**执行命令**:

```bash
# 1) 写入 .gitattributes
# （用 Write 工具或编辑器写入上述内容）

# 2) 针对性重规范化（按文件类型逐个处理，避免误伤）
git add --renormalize '*.py' '*.md' '*.csv' '*.sh' '*.bat' '*.ps1' '*.json'

# 3) 暂存 .gitattributes 本身
git add .gitattributes

# 4) 审计：确认只有文本文件被重规范化
echo "=== 被重规范化的文件（应全部为 .py/.md/.csv/.sh/.bat/.ps1/.json）==="
git diff --cached --stat

# 5) 提交
git commit -m "$(cat <<'EOF'
chore: 统一换行符配置 (.gitattributes)

- Python/Shell/Markdown/CSV/JSON → LF
- Batch/PowerShell → CRLF
- 二进制文件 (.xlsx/.xls/.pkl/.db/~$*) 标记为 binary

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"

# 6) 创建 .git-blame-ignore-revs 并记录此 commit SHA
BLAME_SHA=$(git rev-parse HEAD)
echo "# 跳过换行符批量变更的 commit，避免 git blame 噪音" > .git-blame-ignore-revs
echo "$BLAME_SHA" >> .git-blame-ignore-revs
git add .git-blame-ignore-revs
git commit -m "$(cat <<'EOF'
chore: 添加 .git-blame-ignore-revs（排除换行符批量变更）

使用方式: git blame --ignore-revs-file .git-blame-ignore-revs <file>
可配置为默认: git config blame.ignoreRevsFile .git-blame-ignore-revs

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"

# 7) 推送这两个 commit
git push origin main
```

---

## 验证清单

**第1步提交前**:
- [ ] `git fetch origin main` 完成，`git log HEAD..origin/main` 为空
- [ ] `python -c "from shared import OA_COLUMN_MAP, MAIN_FIELD_MAP, ITEMS_FIELD_MAP, INV_COLUMN_MAP, MAIN_COLUMNS, ITEMS_COLUMNS, INV_COLUMNS, audit_schedule_results; print('OK')"` 通过
- [ ] 4 个源文件语法检查全部通过
- [ ] `git diff --cached --name-only` 确认暂存内容与清单一致

**第1步提交后**:
- [ ] `git status` 只剩 27 个预期未跟踪文件，无残留的 ` D` 状态
- [ ] cache/ 文件仍在本地磁盘（`ls cache/ai_schedule_*.csv` 存在）
- [ ] `git log -1 --stat` 确认 commit 包含所有预期文件

**第2步推送前**:
- [ ] `git log --oneline origin/main..HEAD` 显示 12 个 commit
- [ ] `git push origin main --dry-run` 无错误

**第2步推送后**:
- [ ] `git status` 显示 `Your branch is up to date with 'origin/main'`

**第3步提交前**:
- [ ] `git diff --cached --stat` 显示只有文本文件（.py/.md/.csv/.sh/.bat/.ps1/.json）被重规范化
- [ ] 无二进制文件出现在 diff 中

**第3步推送后**:
- [ ] `.git-blame-ignore-revs` 存在且包含正确的 commit SHA
- [ ] `git blame --ignore-revs-file .git-blame-ignore-revs shared.py` 正常工作

---

## 回滚方案

| 场景 | 操作 |
|------|------|
| **第1步提交了但发现遗漏** | `git reset --soft HEAD~1`（撤销 commit，改动回到暂存区） |
| **第1步 push 前发现问题** | `git reset --soft HEAD~1` 修正后重新 commit |
| **已推送，发现严重 bug** | `git revert <清理commit-SHA>` 然后 `git revert <换行符commit-SHA>` |
| **已推送，想完全撤销** | `git reset --hard origin/main~2` 然后 `git push --force`（需确认无人在此之上开发） |
| **换行符 commit 污染了 blame** | `git blame --ignore-revs-file .git-blame-ignore-revs <file>` 即可跳过，无需回滚 |
| **push 被拒绝（远程有新提交）** | `git pull --rebase origin main` 后重新 push |

---

## 不回滚 / 不做的事

- 不清理 DEBUG print — P1
- 不分类/归档 27 个未跟踪文件 — P1
- 不处理部署配置（start.sh、plist 等）— P2
- 不写 CLAUDE.md — P3
- 不写测试
- 不做功能变更
- 不修改 `测试.xlsx` 内容 — 仅还原到 HEAD 版本

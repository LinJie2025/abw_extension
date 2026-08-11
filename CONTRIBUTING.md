# Contributing Guide ｜ 贡献指南

本仓库采用 **GitHub Flow** 分支策略，请遵守以下规则。

## 铁律

1. `main` 永远保持可发布状态
2. **禁止直接 push 到 main**，一切变更必须走 Pull Request
3. PR 必须至少 1 人 Review 通过
4. 合并使用 **Squash Merge**（一个 PR 一个提交）
5. 提交消息必须符合 Conventional Commits（commitlint 已配置，违规会被拒绝）

## 分支命名

| 前缀 | 用途 | 示例 |
|------|------|------|
| `feat/` | 新功能 | `feat/user-auth` |
| `fix/` | 缺陷修复 | `fix/login-redirect` |
| `refactor/` | 重构 | `refactor/payment-service` |
| `docs/` | 文档 | `docs/api-guide` |
| `chore/` | 依赖/杂务 | `chore/deps-update` |

全小写、连字符分隔、简短有意义。

## 标准流程

```bash
# 1. 从最新 main 建分支
git checkout main && git pull
git checkout -b feat/my-feature

# 2. 开发，原子提交（commitlint 自动校验格式）
git add <files>
git commit -m "feat(auth): 支持扫码登录"

# 3. 开发中定期同步远端，避免冲突积压
git fetch origin
git rebase origin/main

# 4. 推送并开 PR
git push -u origin feat/my-feature
```

## 提交消息格式

```
<type>(<scope>): <subject>

<body: 为什么改>
```

type: `feat` `fix` `refactor` `docs` `chore` `test` `perf` `style` `build` `ci` `revert`

## 快速验证

```bash
npx commitlint --from HEAD~1   # 校验最近一次提交
```

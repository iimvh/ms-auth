## Agent skills

### Issue tracker

Issues live in GitHub Issues for `iimvh/ms-auth`, driven by the `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

The five canonical roles, each label string equal to its name. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context: one `GLOSSARY.md` plus `docs/adr/` at the repo root. See `docs/agents/domain.md`.

# No Comments Policy

- Ban and delete the vast majority of code comments.
- Agents copy existing comments. A comment explaining a temporary workaround spreads like a virus and fills the codebase with hacks and stale notes.
- Code must be self-documenting. Prefer clear names, types, structure, tests, and assertions over comments.
- Keep a comment only when it documents a real external constraint that cannot be expressed in code, types, tests, or lint rules (e.g. an unavoidable browser/platform quirk). Such cases are rare.
- Never use a comment to explain bad or surprising code. If the code needs a comment to be understood, refactor it instead.
- Prefer fixing the root cause over leaving a workaround + explanatory comment.
- When a genuine constraint must be preserved, encode it as a type, test, assertion, or lint rule, then delete the comment.
- Do not write narration comments, step-by-step comments, or restatements of what the code already does.
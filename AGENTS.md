# Repository instructions

This repository implements the Python `m365` command. Product usage guidance for
agents is in `docs/AGENTS.md`; human setup is in `README.md`.

- Preserve JSON stdout, secret-safe errors, account verification, change-key
  preconditions, the command lock, and the no-automatic-write-retry rule.
- Keep runtime browser state, mailbox exports, operation records, test accounts,
  and authentication material out of Git. Never log tokens or browser headers.
- Use synthetic data for offline tests. Live tests require an explicit account and
  authorization for the test traffic and cleanup. Only change labelled test items.
- Update command help, `docs/CLI.md`, `docs/AGENTS.md`, and relevant usage guides
  when changing the supported surface. Keep assessment reports and live test
  results private; public documentation describes current behavior and limits.
- Run the offline suite, lint, format checks, and a package build before release.
- Browser endpoints are undocumented. Fail conservatively on unknown outcomes or
  missing acknowledgements; do not add speculative write retries or fallbacks.

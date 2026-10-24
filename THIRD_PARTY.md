# Runtime dependencies

This repository ships the `m365-cli` Python package and its browser transport
script. Its runtime dependencies are Playwright for Python, Babel, timezone data, markdown-it-py (Markdown parsing),
nh3 (HTML sanitization), and the transitive
packages recorded in `uv.lock`. Their distributions contain their respective
license notices. Google Chrome is installed separately by the user.

There is no dependency on another mail, calendar, or Microsoft 365 CLI. All
supported operations, argument validation, result handling and retry protection
are implemented in this package.

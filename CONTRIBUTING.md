# Contributing

Follow [AGENTS.md](AGENTS.md) when changing the implementation. Human setup is in
[README.md](README.md); agent usage is in [docs/AGENTS.md](docs/AGENTS.md).

## Development environment

From a checkout, install the locked runtime and development dependencies:

```sh
uv sync --locked
```

The package requires Python 3.11 or newer. `uv.lock` records the dependency
versions used by the project environment. Use `uv run --locked m365` to run the
CLI from that environment.

## Checks before release

```sh
uv run --locked python -m unittest discover -s tests
uv run --locked ruff check m365_cli tests scripts
uv run --locked ruff format --check m365_cli tests scripts
uv run --locked python scripts/generate_reference.py
uv run --locked python -m build
```

The offline suite uses synthetic data and does not contact Microsoft. Review the
regenerated command reference and update the applicable usage guides when
changing commands or behavior. For a local coverage report:

```sh
uv run --locked python -m coverage run --source=m365_cli -m unittest discover -s tests
uv run --locked python -m coverage report -m
```

## Service compatibility

The packaged `m365 maintainer` commands provide read probes, explicit write
checks and local command-disable switches. See [docs/MAINTAINER.md](docs/MAINTAINER.md).
Live writes require an explicitly authorized account, destinations and cleanup;
only labelled disposable items may be changed. Do not publish assessment reports,
service exports, cleanup journals or authentication state. Public documentation
should describe current behavior, setup and limitations.

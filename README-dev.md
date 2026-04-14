# Development Notes

This repository keeps runtime and development dependencies separate:

- `requirements.txt` contains the packages required to run the billing scripts.
- `requirements-dev.txt` includes the runtime requirements plus developer tooling such as `ruff` and `black`.

## Setup

Activate the local virtual environment:

```bash
source venv/bin/activate
```

Install development dependencies:

```bash
pip install -r requirements-dev.txt
```

## Checks

Run linting locally:

```bash
ruff check .
```

Run formatting checks locally:

```bash
black --check .
```

Format files locally:

```bash
black .
```

## CI

GitHub Actions runs the same lint and format checks on pushes and pull requests:

- `.github/workflows/lint.yml`

## Notes

- `requirements-dev.txt` should be committed so contributors and CI use the same tool versions.
- `requirements.txt` should remain focused on script/runtime dependencies.

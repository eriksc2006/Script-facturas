# Contributing

## Branches

Use a short-lived branch:

```text
feature/<descripcion>
fix/<descripcion>
chore/<descripcion>
```

Do not commit directly to `main`.

## Before opening a Pull Request

Run:

```bash
black src tests
flake8 src tests
pytest --cov=src/effi_processor --cov-report=term-missing
bandit -r src -ll
pip-audit -r requirements.txt
```

## Pull Request checklist

- [ ] Tests added/updated.
- [ ] Existing tests pass.
- [ ] No secrets or production documents committed.
- [ ] `Plantilla_Importacion` remains formula-free.
- [ ] Effi headers remain byte-for-byte equivalent to the required headers.
- [ ] Audit behavior is preserved.
- [ ] Ambiguous catalog matches are not silently accepted.

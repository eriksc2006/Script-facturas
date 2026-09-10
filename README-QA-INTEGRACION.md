# Integración de los archivos QA

Copia los archivos respetando exactamente sus rutas relativas.

## Archivos nuevos

- `requirements-dev.txt` -> raíz
- `pyproject.toml` -> raíz
- `.pre-commit-config.yaml` -> raíz
- `.github/workflows/qa.yml`
- `.github/workflows/codeql.yml`
- `.github/workflows/dependency-review.yml`
- `.github/workflows/sbom.yml`
- `.github/dependabot.yml`
- `.github/CODEOWNERS`
- `.github/pull_request_template.md`
- `.github/ISSUE_TEMPLATE/bug_report.md`
- `.github/ISSUE_TEMPLATE/feature_request.md`
- `CODE_OF_CONDUCT.md` -> raíz
- `docs/QA.md`
- `tests/conftest.py`
- `tests/unit/test_repository_qa.py`

## Importante

El workflow `qa.yml` crea exactamente tres jobs llamados `test`, `quality` y `security`, porque esos son los nombres que deben registrarse como Required Status Checks en los Rulesets.

Antes de activar esos checks como obligatorios en GitHub, ejecuta el workflow al menos una vez para que los checks existan.

Las Actions están referenciadas por tags mayores para facilitar la instalación inicial. Para un endurecimiento adicional, GitHub recomienda fijarlas a SHA completo e inmutable.

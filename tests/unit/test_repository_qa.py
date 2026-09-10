from pathlib import Path


def test_effi_main_headers_are_documented(effi_main_headers):
    assert len(effi_main_headers) == 6
    assert effi_main_headers[0].startswith("Artículo (ID EFFI")
    assert effi_main_headers[-1] == "Código Effi Impuesto"


def test_qa_configuration_exists():
    root = Path(__file__).resolve().parents[2]
    required = [
        root / "requirements.txt",
        root / "requirements-dev.txt",
        root / "pyproject.toml",
        root / ".pre-commit-config.yaml",
        root / ".github" / "workflows" / "qa.yml",
        root / "docs" / "QA.md",
    ]
    missing = [str(path) for path in required if not path.exists()]
    assert not missing, f"Faltan archivos QA: {missing}"

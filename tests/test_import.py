def test_project_files_exist():
    from pathlib import Path
    assert Path('app_effi.py').exists()
    assert Path('.github/workflows/ci.yml').exists()

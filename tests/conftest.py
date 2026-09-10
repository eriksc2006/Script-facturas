import pytest


@pytest.fixture
def effi_main_headers():
    return [
        "Artículo (ID EFFI | Código de barras GTIN | Serie)",
        "Observación",
        "Cantidad *",
        "Precio ud. *",
        "Valor descuento total. *",
        "Código Effi Impuesto",
    ]

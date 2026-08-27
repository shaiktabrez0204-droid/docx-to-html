"""Adapter test."""
def test_import():
    from docx_to_html.adapter.mammoth_adapter import mammoth_to_html
    assert callable(mammoth_to_html)

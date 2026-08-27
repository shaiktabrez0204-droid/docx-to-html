"""Fixture tests."""
def test_fixtures_loaded():
    """Test that fixtures can be loaded."""
    import os
    fixtures_dir = os.path.join(os.path.dirname(__file__), 'fixtures')
    assert os.path.isdir(fixtures_dir)
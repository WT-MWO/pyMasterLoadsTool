import pytest
from pymasterloadstool import importer


class FakeLoadRecord:
    def __init__(self, type_int) -> None:
        self.Type = type_int


class TestImporter:

    @pytest.mark.parametrize("test_input,expected", [(1, "local"), (0, "global")])
    def test_read_cosystem(self, test_input, expected):
        assert importer.Importer._read_cosystem(test_input) == expected
        with pytest.raises(ValueError):
            importer.Importer._read_cosystem(None)

    @pytest.mark.parametrize("test_input,expected", [(0, "no"), (1, "generated")])
    def test_read_calcnode(self, test_input, expected):
        assert importer.Importer._read_calcnode(test_input) == expected
        with pytest.raises(ValueError):
            importer.Importer._read_calcnode(None)

    @pytest.mark.parametrize("test_input,expected", [(0, "absolute"), (1, "relative")])
    def test_read_relabs(self, test_input, expected):
        assert importer.Importer._read_relabs(test_input) == expected
        with pytest.raises(ValueError):
            importer.Importer._read_relabs(None)

    @pytest.mark.parametrize(
        "test_input,expected",
        [
            (0, True),
            (5, True),
            (26, True),
            (3, True),
            (7, True),
            (6, True),
            (69, True),
            (89, True),
            (28, True),
            (90, False),
        ],
    )
    def test_check_type(self, test_input, expected):
        load_record = FakeLoadRecord(test_input)
        assert importer.Importer._check_type(load_record) == expected
        with pytest.raises(ValueError):
            importer.Importer._check_type(FakeLoadRecord(None))

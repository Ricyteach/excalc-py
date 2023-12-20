import pytest
import faulthandler
import xlwings as xw
import collections.abc

TEST_WB = "test_book.xlsx"


@pytest.fixture(scope="session")
def xw_app():
    # setup
    faulthandler.disable()  # required for testing xlwings with pytest
    app = xw.App(visible=False)
    yield app
    # teardown
    app.quit()


@pytest.fixture
def wb(xw_app):
    return xw_app.books.open(TEST_WB)


@pytest.fixture
def input_rng1(wb):
    return wb.sheets[0].range("B1")


@pytest.fixture
def input_rng2(wb):
    return wb.sheets[0].range("B2")


@pytest.fixture
def output_rng(wb):
    return wb.sheets[0].range("B3")


@pytest.fixture
def input_adapter():
    def multiply_by_2(value):
        return value * 2

    return multiply_by_2


@pytest.fixture
def output_adapter():
    def make_set(value):
        if not isinstance(value, str) and isinstance(value, collections.abc.Iterable):
            return set(value)
        else:
            return {value}

    return make_set

import pytest
from excalc_py import adapt_function


@pytest.fixture
def inputs():
    return [1, 1]


@pytest.fixture
def expected_result():
    return {2}


@pytest.fixture
def my_func_calculation(input_adapter, output_adapter):
    @adapt_function([input_adapter, input_adapter], output_adapter)
    def my_func(a, b):
        return a, b

    return my_func


def test_adapt_for_units(my_func_calculation, inputs, expected_result):
    assert my_func_calculation(*inputs) == expected_result

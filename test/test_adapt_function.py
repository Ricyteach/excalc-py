import pytest
from excalc_py import adapt_function


@pytest.fixture
def inputs():
    return [1, 1]


@pytest.fixture
def expected_result():
    return {2}


@pytest.fixture
def my_func1_calculation(output_adapter, input_adapter):
    @adapt_function(output_adapter, input_adapter, input_adapter)
    def my_func(a, b, *args, **kwargs):
        return a, b

    return my_func


def test_adapt_for_units1(my_func1_calculation, inputs, expected_result):
    assert my_func1_calculation(*inputs) == expected_result


@pytest.fixture
def my_func2_calculation(output_adapter, input_adapter):
    @adapt_function(
        output_adapter, input_adapter, input_adapter, input_adapter, input_adapter
    )
    def my_func(a, b, *args, **kwargs):
        return a, b

    return my_func


def test_adapt_for_units2(my_func2_calculation, inputs, expected_result):
    assert my_func2_calculation(*inputs) == expected_result

from excalc_py import create_calculation, adapt_function
import pytest


@pytest.fixture
def decorated_func(input_rng1, input_rng2, output_rng, input_adapter, output_adapter):
    @adapt_function(output_adapter, input_adapter, input_adapter)
    @create_calculation(output_rng, input_rng1, input_rng2)
    def func(a, b):
        pass

    return func


@pytest.fixture
def inputs():
    return [1, 2]


@pytest.fixture
def expected_output():
    return {6}


def test_combined_decorators(decorated_func, inputs, expected_output):
    assert decorated_func(*inputs) == expected_output


def test_combined_decorators_b_None(decorated_func, inputs, expected_output, output_adapter, input_adapter):
    # NOTE: b is not supplied by python, and just uses the value already in Excel sheet (zero)
    assert decorated_func(inputs[0], b=None) == output_adapter(input_adapter(inputs[0]))

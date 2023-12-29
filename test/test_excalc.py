from excalc_py import create_calculation, _ExcelCalculation, SupportsCalculation
import pytest


@pytest.fixture
def inputs():
    return [1.1, 2.2]


@pytest.fixture
def expected_result():
    return 3.3


@pytest.fixture
def my_func_calculation1(input_rng1, input_rng2, output_rng):
    @create_calculation(
        output_rng,
        input_rng1,
        input_rng2,
    )
    def my_func(a, b):
        pass

    return my_func


def test_calculation_apply(
    my_func_calculation1: SupportsCalculation, input_rng1, input_rng2, inputs
):
    calc = my_func_calculation1.calculation
    calc.apply(*inputs)
    assert calc.parameter_rng_list[0].sheet.range(input_rng1.address).value == inputs[0]
    assert calc.parameter_rng_list[1].sheet.range(input_rng2.address).value == inputs[1]


def test_create_calculation(my_func_calculation1, inputs, expected_result):
    assert my_func_calculation1(*inputs) == pytest.approx(expected_result)


def test_create_calculation_kwargs(my_func_calculation1, inputs, expected_result):
    assert my_func_calculation1(**dict(zip("ab", inputs))) == pytest.approx(
        expected_result
    )

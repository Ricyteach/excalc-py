from excalc_py import create_calculation, _Calculation
import pytest


@pytest.fixture
def inputs():
    return [1.1, 2.2]


@pytest.fixture
def expected_result():
    return 3.3


@pytest.fixture
def calculation():
    return _Calculation()


@pytest.fixture
def my_func_calculation(input_rng1, input_rng2, output_rng):
    @create_calculation(output_rng, input_rng1, input_rng2, )
    def my_func(a, b):
        pass

    return my_func


def test_calculation_apply(calculation, input_rng1, input_rng2, inputs):
    calculation.input_rng_list = [input_rng1, input_rng2]
    calculation.apply(*inputs)
    assert (
        calculation.input_rng_list[0].sheet.range(input_rng1.address).value == inputs[0]
    )
    assert (
        calculation.input_rng_list[1].sheet.range(input_rng2.address).value == inputs[1]
    )


def test_create_calculation(my_func_calculation, inputs, expected_result):
    assert my_func_calculation(*inputs) == pytest.approx(expected_result)


def test_create_calculation_no_kwargs(my_func_calculation, inputs):
    with pytest.raises(TypeError):
        my_func_calculation(dict(zip("ab", inputs)))

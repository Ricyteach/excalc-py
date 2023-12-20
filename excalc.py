"""Turn your Excel calculators into python functions."""

__version__ = "0.0.1"

from functools import wraps
import xlwings as xw
import collections.abc


class _Calculation:
    """Object used to work with the Excel ranges and get the calculations."""

    _input_rng_list: list[xw.main.Range] = None
    _output_rng: xw.main.Range = None

    @property
    def input_rng_list(self):
        return self._input_rng_list

    @input_rng_list.setter
    def input_rng_list(self, value):
        self._input_rng_list = []
        for input_rng in value:
            if not isinstance(input_rng, xw.main.Range):
                raise TypeError(
                    f"Range objects are required input_rng_list, not {type(input_rng)}"
                )
            self._input_rng_list.append(input_rng)

    @property
    def output_rng(self):
        return self._output_rng

    @output_rng.setter
    def output_rng(self, value):
        if not isinstance(value, xw.main.Range):
            raise TypeError("Range object required for output_rng")
        self._output_rng = value

    def apply(self, *args, **kwargs):
        """Write the input args to the input ranges."""

        if kwargs:
            raise NotImplementedError("kwargs in calculation not supported")
        for arg, input_rng in zip(args, self.input_rng_list, strict=True):
            shape = input_rng.shape
            if len(shape) != 2:
                raise RuntimeError(f"unexpected type of range shape for {shape}")
            if not isinstance(arg, str) and isinstance(arg, collections.abc.Collection):
                if input_rng.shape == (1, 1):
                    # NOTE: xlwings already elegantly handles the case that arg is collection of shape (1,1)
                    raise TypeError(
                        f"not allowed to assign  type {type(arg).__name__} to singe cell {input_rng.address} because "
                        f"it contains multiple values"
                    )
                elif input_rng.shape[1] == 1:
                    # convert argument sequence to be compatible with a single column
                    input_rng.value = [[v] for v in arg]
                elif input_rng.shape[0] == 1:
                    # argument sequence is a single row
                    input_rng.value = [*arg]
                else:
                    # argument sequence needs to be 2d and not a 1d sequence of strings, nor a dict
                    # checking dict bc xlwings allows dicts; but see no harm in allowing dicts in other cases above
                    rows = []
                    for row in args:
                        if isinstance(row, str) or isinstance(row, dict):
                            raise TypeError(
                                f"2d range {input_rng.address} not compatible with 1d input {row}"
                            )
                        rows.append([*row])
                    input_rng.value = rows
            else:
                # single argument
                input_rng.value = arg

    def retrieve(self):
        """Get the output from the output range."""
        return self.output_rng.value

    def __call__(self, *args, **kwargs):
        self.apply(*args, **kwargs)
        return self.retrieve()


def create_calculation(input_rng_list, output_rng):
    """Decorator for turning Excel apps into functions.

    The functions use Excel in the background instead of computation being completed in python.
    """

    calc = _Calculation()
    calc.input_rng_list = input_rng_list
    calc.output_rng = output_rng

    def wrapper(func):
        @wraps(func)
        def wrapped(*args, **kwargs):
            return calc(*args, **kwargs)

        wrapped.calculation = calc
        return wrapped

    return wrapper


def adapt_function(input_adapter_list, output_adapter):
    """Decorator that allows modification of the inputs and outputs."""

    def wrapper(func):
        @wraps(func)
        def wrapped(*args, **kwargs):
            args = [
                input_adapter(arg)
                for input_adapter, arg in zip(input_adapter_list, args, strict=True)
            ]
            result = func(*args, **kwargs)
            return output_adapter(result)

        return wrapped

    return wrapper

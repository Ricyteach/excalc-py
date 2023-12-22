"""Turn your Excel calculators into python functions."""

__version__ = "0.0.2"

from functools import wraps
from inspect import signature, BoundArguments, Parameter
import xlwings as xw
import collections.abc


class _Calculation:
    """Object used to work with the Excel ranges and get the calculations."""

    _input_rng_list: list[xw.main.Range] = None
    _input_rng_dict: dict[str, xw.main.Range] = None
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
    def input_rng_dict(self):
        return self._input_rng_dict

    @input_rng_dict.setter
    def input_rng_dict(self, value):
        self._input_rng_dict = {}
        for k, input_rng in value.items():
            if not isinstance(input_rng, xw.main.Range):
                raise TypeError(
                    f"Range objects are required input_rng_list, not {type(input_rng)}"
                )
        self._input_rng_dict.update(value)

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
            raise NotImplementedError("kwargs in calculation not yet supported")

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


def create_calculation(output_rng, /, *input_rng_list, **input_rng_dict):
    """Decorator for turning Excel apps into functions.

    The functions use Excel in the background instead of computation being completed in python.
    """

    if input_rng_dict:
        raise NotImplementedError("**kwargs in create_calculation not yet supported")

    calc = _Calculation()
    calc.input_rng_list = input_rng_list
    calc.input_rng_dict = input_rng_dict
    calc.output_rng = output_rng

    def wrapper(func):
        @wraps(func)
        def wrapped(*args, **kwargs):
            return calc(*args, **kwargs)

        wrapped.calculation = calc
        return wrapped

    return wrapper


def do_nothing_adapter(item):
    return item


def adapt_function(output_adapter=None, /, *input_adapter_list, **kwargs):
    """Decorator that allows modification of the inputs and outputs."""

    # handle case of no output adapter
    if output_adapter is None:
        output_adapter = do_nothing_adapter

    def wrapper(func):
        sig = signature(func)

        # integrate kwargs adapters into the rest of the adapters
        partial_bound_adapters = sig.bind_partial(*input_adapter_list, **kwargs)

        full_input_adapter_list = []
        for param_name in sig.parameters.keys():
            adapter = partial_bound_adapters.arguments.pop(param_name, do_nothing_adapter)
            full_input_adapter_list.append(adapter)

        if len(sig.parameters) != len(full_input_adapter_list):
            raise TypeError(
                f"the length of the function {func.__name__} parameters list ({len(sig.parameters)}) must "
                f"match the length of the input adapter list ({len(full_input_adapter_list)})"
            )

        # create a modified input adapter list that takes into account *args and **kwargs supplied to called func
        modified_input_adapter_list = []
        for input_adapter, parameter in zip(full_input_adapter_list, sig.parameters.values(), strict=True):
            if parameter.kind is Parameter.VAR_POSITIONAL:

                def positional_adapter(pos_args):
                    return [input_adapter(pos_arg) for pos_arg in pos_args]

                input_adapter = positional_adapter
            elif parameter.kind is Parameter.VAR_KEYWORD:

                def keyword_adapter(kwd_args: dict):
                    return {k: input_adapter(v) for k, v in kwd_args.items()}

                input_adapter = keyword_adapter
            modified_input_adapter_list.append(input_adapter)

        @wraps(func)
        def wrapped(*args, **kwargs):
            bound_arguments = sig.bind(*args, **kwargs)
            bound_arguments.apply_defaults()
            bound_arguments.arguments = {
                k: adapter(v)
                for adapter, (k, v) in zip(
                    modified_input_adapter_list, bound_arguments.arguments.items()
                )
            }
            result = func(*bound_arguments.args, **bound_arguments.kwargs)
            return output_adapter(result)

        return wrapped

    return wrapper

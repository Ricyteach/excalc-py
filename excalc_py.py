"""Turn your Excel calculators into python functions.

Excalc currently requires xlwings. May support other python Excel libraries in the future.
"""

__version__ = "0.0.3"

from functools import wraps
from inspect import signature, Parameter, Signature, BoundArguments
from typing import Protocol
import xlwings as xw
import collections.abc


def _check_input_rng_list(input_rng_args_lst):
    for input_rng in input_rng_args_lst:
        if not isinstance(input_rng, xw.main.Range):
            raise TypeError(
                f"Range objects are required _ExcelCalculation args, not {type(input_rng)}"
            )


def _check_input_rng_dict(input_rng_kwargs_dct):
    for k, input_rng in input_rng_kwargs_dct.items():
        if not isinstance(input_rng, xw.main.Range):
            raise TypeError(
                f"Range objects are required _ExcelCalculation kwargs, not {type(input_rng)}"
            )


def _check_output_rng(output_rng_arg):
    if not isinstance(output_rng_arg, xw.main.Range):
        raise TypeError("Range object required for output_rng")


class _ExcelCalculation:
    """Object used to work with the Excel ranges and get the calculations.

    Calling the object with arguments performs the calculation using Excel.
    """

    signature: Signature = None
    bound_input_rngs: BoundArguments = None

    def __init__(self, func, output_rng, *input_rng_args, **input_rng_kwargs):
        self.signature = signature(func)
        _check_output_rng(output_rng)
        self.output_rng = output_rng
        _check_input_rng_list(input_rng_args)
        input_rng_list = self.input_rng_list = list(input_rng_args)
        _check_input_rng_dict(input_rng_kwargs)
        input_rng_dict = self.input_rng_dict = input_rng_kwargs
        sig = self.signature = signature(func)

        bound_input_rngs = self.bound_input_rngs = sig.bind(
            *input_rng_list, **input_rng_dict
        )

        # disallow *args or **kwargs in decorated functions (doesn't make any sense)
        if any(
            sig.parameters[k].kind in (Parameter.VAR_KEYWORD, Parameter.VAR_POSITIONAL)
            for k in bound_input_rngs.arguments.keys()
        ):
            starred = (
                repr(param)[repr(param).index("*"): -2]
                for param in sig.parameters.values()
                if param.kind in (Parameter.VAR_KEYWORD, Parameter.VAR_POSITIONAL)
            )
            raise TypeError(
                f"{', '.join(starred)} are not allowed in the decorated function signature"
            )

        self.parameter_rng_list = list(bound_input_rngs.arguments.values())

    def apply(self, *args, **kwargs):
        """Write the input args to the input ranges."""

        bound_args = self.signature.bind(*args, **kwargs)

        for arg, input_rng in zip(
            bound_args.arguments.values(),
            self.bound_input_rngs.arguments.values(),
            strict=True,
        ):
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


class SupportsCalculation(Protocol):
    calculation: _ExcelCalculation

    def __call__(self, *args, **kwargs):
        ...


def create_calculation(output_rng, /, *input_rng_list, **input_rng_dict):
    """Decorator for turning Excel apps into functions.

    The functions use Excel in the background instead of computation being completed in python.
    """

    def wrapper(func) -> SupportsCalculation:
        calc = _ExcelCalculation(func, output_rng, *input_rng_list, **input_rng_dict)

        @wraps(func)
        def wrapped(*args, **kwargs):
            return calc(*args, **kwargs)

        wrapped.calculation = calc
        return wrapped

    return wrapper


def do_nothing_adapter(item):
    """Default adapter when no adapter is supplied for an input."""

    return item


def adapt_function(output_adapter=None, /, *input_adapter_list, **input_adapter_dict):
    """Decorator that allows modification of the inputs and outputs."""

    # handle case of no output adapter
    if output_adapter is None:
        output_adapter = do_nothing_adapter

    def wrapper(func):
        sig = signature(func)

        # integrate kwargs adapters into the rest of the adapters
        partial_bound_adapters = sig.bind_partial(
            *input_adapter_list, **input_adapter_dict
        )

        full_input_adapter_list = []
        for param_name in sig.parameters.keys():
            bound_adapter = partial_bound_adapters.arguments.pop(
                param_name, do_nothing_adapter
            )
            full_input_adapter_list.append(bound_adapter)

        if len(sig.parameters) != len(full_input_adapter_list):
            raise TypeError(
                f"the length of the function {func.__name__} parameters list ({len(sig.parameters)}) must "
                f"match the length of the input adapter list ({len(full_input_adapter_list)})"
            )

        # create a modified input adapter list that takes into account *args and **kwargs supplied to called func
        modified_input_adapter_list = []
        for input_adapter, parameter in zip(
            full_input_adapter_list, sig.parameters.values(), strict=True
        ):
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

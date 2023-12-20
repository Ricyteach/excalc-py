# excalc

#### Turn your Excel calculators into python functions.

Say we have a Very Special Workbook (VSB) called `vsb.xlsx`. We are using this workbook to do some very special calculations.

Sheet1 of the VSB looks like this:

|   | A           | B     | C         | D                      | E |
|---|-------------|-------|-----------|------------------------|---|
| 1 | Diameter    | 1.0   | feet      |                        |   |
| 2 | Circle Area | 115.0 | sq inches | (rounded to nearest 5) |   |
| 3 |             |       |           |                        |   |

Granted, this is a weird calculation. But life is full of mysteries.

###### A Tour of the VSB

Obviously cell B1 is an input- a circle diameter (in feet). And B2 is an output- a circle area, in sq inches, and rounded to the nearest 5. Are you wondering what formula cell B2 contains? Don't worry about it. Life is also short.

##### Let's pythonize it!

Now, how can we turn this Excel calculator into a python function? Like so:

```python
import xlwings as xw
from excalc import create_calculation

# first, using xlwings, we'll get a reference to our sheet
sheet = xw.App().books.open("vsb.xlsx").sheets[0]  # Sheet 1 of vsb.xlsx

# and define where the inputs are
input_rng1 = sheet.range("B1")

# and the where the outputs are too
output_rng = sheet.range("B2")

# now just create a python function, and decorate it
@create_calculation([input_rng1], output_rng)
def weird_circle_area(diameter):
    pass
```

And use it like so:

```python
print(weird_circle_area(1.0))
# 115.0

# don't forget to close the Excel app!
sheet.book.app.quit()
```

Nice!

##### Ok great but this is useless without taking care of the units, Rick.

I agree. Boy... do those units make me nervous. How about a way to adapt the units for inputs and outputs...?

You got it:

```python
import xlwings as xw
from excalc import create_calculation, adapt_function

# as before, using xlwings, we'll get references to our input and output locations
sheet = xw.App(visible=False).books.open("vsb.xlsx").sheets[0]  # Sheet 1 of vsb.xlsx
input_rng1 = sheet.range("B1")
output_rng = sheet.range("B2")

# to handle units, we want to use `pint` unit objects (because `pint` rules)
from pint import UnitRegistry
U = UnitRegistry()

# so let's make a couple of adapters for our input and output

# for the input, which is in feet, the adapter looks like this
def input_feet_adapter(length):
    # just remove the feet!
    return length.to("ft").magnitude

# for the output, which is in sq in, the adapter looks like this
def output_sq_in_adapter(area_sq_in):
    # just add units to the output!
    area = area_sq_in * U.inches**2
    return area

# now just create a python function as before, but with the extra adapter decorator
@adapt_function([input_feet_adapter], output_sq_in_adapter)
@create_calculation([input_rng1], output_rng)
def weird_circle_area(diameter):
    pass
```

And use it like so:

```python
print(weird_circle_area(1.0*U.ft))
# 115 inch ** 2

# don't forget to close the Excel app!
sheet.book.app.quit()
```

VERY NICE!

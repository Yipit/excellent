# Excellent (v0.0.5)

Python library for writing CSV and XLS files from list of
[OrderedDict](http://docs.python.org/2/library/collections.html#collections.OrderedDict)
or just a list of lists containing 2-item tuples.

# Installing

```shell
pip install excellent
```

# Writing [Excel](http://en.wikipedia.org/wiki/Microsoft_Excel) files

```python
from collections import OrderedDict
from excellent import Writer, XL

output = open("superhero-database.xls", "wb")
sheet_manager = XL()

excel = Writer(sheet_manager, output)

sheet_manager.use_sheet("Weaknesses")
excel.write([
    OrderedDict([('Superhero', 'Superman'), ('Weakness', 'Kryptonite')]),
    OrderedDict([('Superhero', 'Spiderman'), ('Weakness', 'Maryjane')]),
])

sheet_manager.use_sheet("Special Powers")
excel.write([
    OrderedDict([('Superhero', 'Superman'), ('Super Powers', 'X-Ray vision')]),
    OrderedDict([('Superhero', 'Spiderman'), ('Super Powers', 'Release web')]),
    OrderedDict([('Superhero', 'Batman'), ('Super Powers', 'Money')]),
])

sheet_manager.use_sheet("Weaknesses")
excel.write([
    OrderedDict([('Superhero', 'Batman'), ('Weakness', 'Social Interactions')]),
])

# save it

excel.save()

# now open superhero-database.xls and be happy
```

![https://raw.github.com/Yipit/excellent/master/docs/superhero-database.png](https://raw.github.com/Yipit/excellent/master/docs/superhero-database.png)

# Writing CSV files
```python
from collections import OrderedDict
from excellent import Writer, CSV

output = open("pokemon-database.csv", "w")
csv = Writer(CSV(delimiter=';'), output)

csv.write([
    OrderedDict([('Name', 'Pikachu'), ('Type', 'Electric')]),
    OrderedDict([('Name', 'Jigglypuff'), ('Type', 'Normal')]),
    OrderedDict([('Name', 'Mew'), ('Type', 'Psychic')]),
])

# save it
csv.save()

# now open pokemon.csv and be happy
```

## pokemon-database.csv:

    Name;Type
    Pikachu;Electric
    Jigglypuff;Normal
    Mew;Psychic


# Styles

```python
excel.write([
    OrderedDict([('Superhero', 'Batman'), ('Weakness', 'Social Interactions')]),
    ],
    bold=True,
    bottom_border=True,
)
```

# Format Strings

```python
excel.write([
    OrderedDict([('Superhero', 'Batman'), ('Weakness', 'Social Interactions')]),
    ],
    format_string='"$"#,##0_);("$"#,##',
)
```

See [the examples](https://github.com/python-excel/xlwt/blob/master/xlwt/examples/num_formats.py) for uses.

# Custom Styles

```python
from xlwt import XFStyle, Alignment
alignment = Alignment()
alignment.horz = Alignment.HORZ_RIGHT
style = XFStyle()
style.alignment = alignment

excel.write([
    OrderedDict([('Superhero', 'Batman'), ('Weakness', 'Social Interactions')]),
    ],
    style=style,
)
```

## Default Style

```python
from xlwt import XFStyle, Alignment
alignment = Alignment()
alignment.horz = Alignment.HORZ_RIGHT
style = XFStyle()
style.alignment = alignment

backend = XL(default_style=style)
output = open("database.csv", "w")
writer = Writer(backend=backend, output=output)
```

# Exceptions

`Excellent.exceptions.TooManyRowsError`: This exception is raised if too many rows are written to an excel spreadsheet.

# Hacking

## Install dependencies

```shell
pip install -r requirements.txt
```

## Run the tests
```shell
make unit
make functional
```

# Licensing

Please check the [`COPYING`](COPYING) file distributed with this library

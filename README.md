# Excellent

Python library for writing CSV and XLS files from list of dictionaries

# Installing 

```shell
pip install excellent
```

# Writing [excel](http://en.wikipedia.org/wiki/Microsoft_Excel) files

```python
from collections import OrderedDict
from excellent import Writer, XL

output = open("superhero-database.xls", "wb")
excel = Writer(XL(), output)

excel.write([
    OrderedDict([('Superhero', 'Superman'), ('Weakness', 'Kryptonite')]),
    OrderedDict([('Superhero', 'Spiderman'), ('Weakness', 'Maryjane')]),
])

excel.write([
    OrderedDict([('Superhero', 'Batman'), ('Weakness', 'Social Interactions')]),
])

# save it

excel.save()

# now open superhero-database.xls and be happy
``` 




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

# Excellent

Python library for writing CSV and XLS files from list of dictionaries

# Installing 

```shell
pip install excellent
```

# Writing [Excel](http://en.wikipedia.org/wiki/Microsoft_Excel) files

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

![https://raw.github.com/Yipit/excellent/master/docs/superhero-database.png?login=suneel0101&token=79faadd827d16c56064ea3845850f7b8](https://raw.github.com/Yipit/excellent/master/docs/superhero-database.png?login=suneel0101&token=79faadd827d16c56064ea3845850f7b8)

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

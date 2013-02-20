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

![https://raw.github.com/Yipit/excellent/master/docs/superhero-database.png?login=suneel0101&token=79faadd827d16c56064ea3845850f7b8](https://raw.github.com/Yipit/excellent/master/docs/superhero-database.png?login=suneel0101&token=79faadd827d16c56064ea3845850f7b8)



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

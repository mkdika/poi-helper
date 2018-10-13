# Apache POI Helper

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](/LICENSE)

[Apache POI](https://poi.apache.org/) helper to ease and simplify spesific case usage.

## Features

#### 1. Collection to Plain Excel

Usage example:

```java
// data type is List of Person class
List<Person> data = personService.findAll();

// Instance helper and generate to excel file.
CollectionToPlainExcel coex = new CollectionToPlainExcel();
coex.setData(data);
coex.writeToExcel("c:\\temp\\data.xlsx");
```

_Other features will coming soon.._

## Copyright and License

Copyright 2018 Maikel Chandika (mkdika@gmail.com). Code released under the 
MIT License. See [LICENSE](/LICENSE) file.
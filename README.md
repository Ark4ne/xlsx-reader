# xlsx-reader
High performance XLSX Reader, with very low memory consumption.

# Installation
```
$ composer require ark4ne/xlsx-reader
```

# Usage
```php
$file = "my-calc.xlsx";

$reader = new \Ark4ne\XLSXReader\XLSXReader($file);

$reader->load();

foreach ($reader->read() as $row){
    // do stuff
}
```

# Performance

| 65K rows, 10cols |  load (s) | read (s) | max process mem | max php mem |
|------------------|----------:|---------:|----------------:|------------:|
| 327K strings     | 0.416     | 3.20     | 34.9M           | 27.9M       |
| 32K strings      | 0.044     | 3.29     | 15.0M           | 2.7M        |
| 512 strings      | 0.010     | 3.24     | 12.8M           | 0.8M        |

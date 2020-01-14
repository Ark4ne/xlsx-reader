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
| 327K strings     | 0.433     | 1.49     | 34.7M           | 27.9M       |
| 32K strings      | 0.044     | 1.51     | 14.6M           | 2.76M       |
| 512 strings      | 0.002     | 1.50     | 12.6M           | 0.82M       |

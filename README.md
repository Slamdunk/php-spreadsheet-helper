# Slam PhpSpreadsheet helper to create organized data table

[![Latest Stable Version](https://img.shields.io/packagist/v/slam/php-spreadsheet-helper.svg)](https://packagist.org/packages/slam/php-spreadsheet-helper)
[![Downloads](https://img.shields.io/packagist/dt/slam/php-spreadsheet-helper.svg)](https://packagist.org/packages/slam/php-spreadsheet-helper)
[![Integrate](https://github.com/Slamdunk/php-spreadsheet-helper/workflows/Integrate/badge.svg?branch=master)](https://github.com/Slamdunk/php-spreadsheet-helper/actions)
[![Code Coverage](https://codecov.io/gh/Slamdunk/php-spreadsheet-helper/coverage.svg?branch=master)](https://codecov.io/gh/Slamdunk/php-spreadsheet-helper?branch=master)

## Installation

`composer require slam/php-spreadsheet-helper`

## Usage

```php
use Slam\PhpSpreadsheetHelper as ExcelHelper;

require __DIR__ . '/vendor/autoload.php';

// Being an iterable, the data can be any dinamically generated content
// for example a PDOStatement set on unbuffered query
$users = new ArrayIterator([
    [
        'column_1' => 'John',
        'column_2' => '123.45',
        'column_3' => '2017-05-08',
    ],
    [
        'column_1' => 'Mary',
        'column_2' => '4321.09',
        'column_3' => '2018-05-08',
    ],
]);

$columnCollection = new ExcelHelper\ColumnCollection([
    new ExcelHelper\Column('column_1',  'User',     10,     new ExcelHelper\CellStyle\Text()),
    new ExcelHelper\Column('column_2',  'Amount',   15,     new ExcelHelper\CellStyle\Amount()),
    new ExcelHelper\Column('column_3',  'Date',     15,     new ExcelHelper\CellStyle\Date()),
]);

$filename = sprintf('%s/my_excel_%s.xls', __DIR__, uniqid());

$phpExcel = new ExcelHelper\TableWriter($filename);
$worksheet = $phpExcel->addWorksheet('My Users');

$table = new ExcelHelper\Table($worksheet, 0, 0, 'My Heading', $users);
$table->setColumnCollection($columnCollection);

$phpExcel->writeTableToWorksheet($table);
$phpExcel->close();
```

Result:

![Example](https://raw.githubusercontent.com/Slamdunk/php-spreadsheet-helper/master/example.png)

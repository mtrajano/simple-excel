# SimpleExcel

SimpleExcel is a wrapper around the PHPOffice/PHPExcel package that gives it a few extra features. Some of theses features include:

### Matrix Array Access
```php
//PHPExcel.php
$workbook = new PHPExcel();
$worksheet = $workbook->getActiveSheet();

$worksheet->setCellValue('A1', 'A test value');
```

```php
//SimpleExcel.php
$workbook = new Mtrajano\SimpleExcel\Workbook();
$worksheet = $workbook->getActiveSheet();

$worksheet[1]['A'] = 'A test value';
```

### Numeric Column Access
```php
//PHPExcel.php
$worksheet->setCellValue('B1', 'Another test value');
```

```php
//SimpleExcel.php
$worksheet[1][1] = 'Another test value';
```

### Seamless export
```php
//PHPExcel.php
$writer = new PHPExcel_Writer_Excel2007($workbook);
$writer->save("phpexcel.xlsx");

$writer = new PHPExcel_Writer_Excel5($workbook);
$writer->save("phpexcel.xls");
```

```php
//SimpleExcel.php
$workbook->save("simple.xlsx", Workbook::WRITE_Excel2007);
$workbook->save("simple.xls", Workbook::WRITE_Excel5);
```

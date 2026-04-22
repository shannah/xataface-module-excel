# Xataface Excel Module

Export result sets in XLS format.

## Synopsis

This module adds an action to export result sets as XLS files.  Xataface has always had a CSV export option, but it turns on MS Excel can have difficulty determining the charset encoding of CSV files.  
The solution, sadly, is to give Excel users the option of exporting result sets directly as XLS files.

## Requirements

- Xataface 2.0.4 (or at least include the modularization of export_csv action with this changeset:
    https://github.com/shannah/xataface/commit/34686c0f716773572bbd06274d518a3fb0953467
)
- PHP 5.6 or higher
- **PHP 8.0+:** Requires [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) installed via Composer (see Installation step 3 below). On PHP 5.6–7.x the bundled PHPExcel library is used automatically.

## License

LPGL

## Installation

1. Check out the git repo into your application's modules directory.  The module name should be "excel":
```
$ cd /path/to/myapp/modules
$ git clone https://github.com/shannah/xataface-module-excel.git excel
```

2. Add the following to the [_modules] section of your application's conf.ini file:
```
modules_excel=modules/excel/excel.php
```

3. **PHP 8.0+ only:** Install PhpSpreadsheet via Composer in the module directory:
```
$ cd /path/to/myapp/modules/excel
$ composer require phpoffice/phpspreadsheet
```
The module loads its own `vendor/autoload.php` automatically — no changes to your `index.php` are needed.

4. Confirm that the module has been installed by opening your application in the web browser, then (if you're using the g2 module) click on the "Export" button on the top toolbar (when in list view) and confirm that you have an option "Export XLS".


## Usage

Once you have installed the module you should have an Export XLS option under the "Export" menu in list view.  This works identical to the way that the Export CSV action works.  Perform a find
for the records you want to export, then select "Export XLS".  This will download the results as an excel file.

## Support

Direct support questions to the Xataface forum http://xataface.com/forum

## Credits

- Module written by Steve Hannah.
- Uses [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) on PHP 7.2+, with a bundled [PHPExcel](https://github.com/PHPOffice/PHPExcel/) fallback on older PHP versions.
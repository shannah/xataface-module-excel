<?php

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Excel export action for xataface list views.
 *
 * Supports two backends:
 *   - PhpSpreadsheet (preferred, requires PHP 7.2+).
 *   - Bundled PHPExcel 1.7.9 under lib/ (legacy fallback, PHP 5.6 - 7.1).
 *
 * On PHP 8+ without PhpSpreadsheet installed, construction fails fast
 * with an actionable error message, because the bundled PHPExcel is not
 * PHP 8 compatible.
 *
 * @author shannah
 */
import('actions/export_csv.php');

class actions_excel_export_xls extends dataface_actions_export_csv {

    const BACKEND_PHPSPREADSHEET = 'phpspreadsheet';
    const BACKEND_PHPEXCEL       = 'phpexcel';

    private $backend = null;
    private $book = null;
    private $row = 0;

    // Preserved for backward compatibility with any delegate code that
    // reads the spreadsheet off the action. Mirrors $this->book.
    public $objPHPExcel = null;

    public function __construct(){
        if (PHP_VERSION_ID >= 70200 && class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet')) {
            $this->backend = self::BACKEND_PHPSPREADSHEET;
        } elseif (PHP_VERSION_ID < 80000) {
            import('modules/excel/lib/PHPExcel.php');
            $this->backend = self::BACKEND_PHPEXCEL;
        } else {
            $moduleDir = dirname(__DIR__);
            throw new Exception(
                "The Xataface Excel module requires the PhpSpreadsheet library on PHP 8+.\n"
                . "The bundled PHPExcel library is not compatible with PHP 8.\n\n"
                . "To fix this, run:\n"
                . "  cd " . $moduleDir . "\n"
                . "  composer require phpoffice/phpspreadsheet\n\n"
                . "See the module README for details."
            );
        }
    }

    function writeRow($fh, $data, $query){
        $col = 0;
        foreach ($data as $val) {
            $this->setCell($col++, $this->row + 1, $val);
        }
        $this->row++;
    }

    private function setCell($col, $row, $val){
        if ($this->backend === self::BACKEND_PHPSPREADSHEET) {
            // PhpSpreadsheet's Coordinate helper is 1-indexed.
            $cellName = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col + 1) . $row;
            $this->book->getActiveSheet()->setCellValue($cellName, $val);
        } else {
            $cellName = self::getCellName($col, $row);
            $this->book->setActiveSheetIndex(0)->setCellValue($cellName, $val);
        }
    }

    static function getCellName($col, $row){
        $colCode = self::num2alpha($col);
        return $colCode.''.$row;
    }

    static function num2alpha($n) {
        $r = '';
        for ($i = 1; $n >= 0 && $i < 10; $i++) {
            $r = chr(0x41 + ($n % pow(26, $i) / pow(26, $i - 1))) . $r;
            $n -= pow(26, $i);
        }
        return $r;
    }

    function startFile($fh, $query){
        $author = Dataface_Application::getInstance()->getSiteTitle();
        if (class_exists('Dataface_AuthenticationTool')) {
            $author = Dataface_AuthenticationTool::getInstance()->getLoggedInUsername();
        }
        if ($author === null) {
            $author = '';
        }

        if ($this->backend === self::BACKEND_PHPSPREADSHEET) {
            $this->book = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
            $this->book->getProperties()->setCreator($author);
            $this->book->setActiveSheetIndex(0);
        } else {
            $this->book = new PHPExcel();
            $this->book->getProperties()->setCreator($author);
            $this->book->setActiveSheetIndex(0);
        }

        $this->objPHPExcel = $this->book;
    }

    function endFile($fh, $query){

    }

    function writeOutput($fh, $query){
        // Clear any output buffers that Xataface may have started,
        // so the binary XLS stream is sent directly to the client.
        while ( @ob_end_clean() );

        // Redirect output to a client's web browser (Excel5/Xls)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$query['-table'].'_results_'.date('Y_m_d_H_i_s').'.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        if ($this->backend === self::BACKEND_PHPSPREADSHEET) {
            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->book, 'Xls');
        } else {
            $writer = PHPExcel_IOFactory::createWriter($this->book, 'Excel5');
        }
        $writer->setPreCalculateFormulas(false);
        $writer->save('php://output');
        exit;
    }

}

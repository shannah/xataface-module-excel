<?php

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Description of excel_export_xls
 *
 * @author shannah
 */
import('modules/excel/lib/PHPExcel.php');
import('actions/export_csv.php');
class actions_excel_export_xls extends dataface_actions_export_csv {
        
        private $objPHPExcel = null;
        private $row = 0;
        
        
        public function __construct(){
            
        }
	
        function writeRow($fh, $data, $query){
            $col = 0;
            foreach ( $data as $val ){
                $cellName = $this->getCellName($col++, $this->row+1);
                $this->objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName, $val);
            }
            $this->row++;
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
            $this->objPHPExcel = new PHPExcel();
            $author = Dataface_Application::getInstance()->getSiteTitle();
            if ( class_exists('Dataface_AuthenticationTool') ){
                $author = Dataface_AuthenticationTool::getInstance()->getLoggedInUsername();
            }
            $this->objPHPExcel->getProperties()->setCreator($author);
                //->setLastModifiedBy("Maarten Balliauw")
                //->setTitle("Office 2007 XLSX Test Document")
                //->setSubject("Office 2007 XLSX Test Document")
                //->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                //->setKeywords("office 2007 openxml php")
                //->setCategory("Test result file");
            
            $this->objPHPExcel->setActiveSheetIndex(0);
        }
        
        function endFile($fh, $query){
            
        }
        
        function writeOutput($fh, $query){
            // Redirect output to a client’s web browser (Excel5)
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

            $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');
            $objWriter->setPreCalculateFormulas(false);
            $objWriter->save('php://output');
            exit;
            
        }
    

}



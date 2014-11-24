<?php defined('SYSPATH') or die('No direct access allowed.');

/**
 * PHP Excel library. Helper class to make and read spreadsheet easier
 * 
 * @package Koahana
 * @category spreadsheet
 * @author Katan, <Original>
 * @author Lord Mangila, <Modified_by>
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * 
 * @see https://github.com/rafsoaken/kohana-phpexcel (Flynsarmy, Dmitry Shovchko)
 * 
 */
class Kohana_Excel
{
    /**
     * @var PHPExcel
     */
    public $_excel;

    /**
     * @var array Valid types for PHPExcel
     */
    protected $options = array(
        'title' => 'New Spreadsheet',
        'subject' => 'New Spreadsheet',
        'description' => 'New Spreadsheet',
        'author' => 'None',
        'format' => 'Excel2007',
        'path' => './',
        'name' => 'NewSpreadsheet',
        'filename' => '', // Filename for read
        'csv_values' => array('delimiter' => ';', 'lineEnding' => "\r\n")// CSV file
    );

    /**
     * @var array file extentions
     */
    private $exts = array(
        'CSV' => 'csv',
        'PDF' => 'pdf',
        'Excel5' => 'xls',
        'Excel2007' => 'xlsx',
    );

    /**
     * @var array file mimes
     */
    private $mimes = array(
        'CSV' => 'text/csv',
        'PDF' => 'application/pdf',
        'Excel5' => 'application/vnd.ms-excel',
        'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    /**
     * Creates the spreadsheet class with given or default settings
     * @param array $options with optional parameters: title, subject, description, author
     * @return Excel 
     */
    public static function factory($options = array())
    {
        return new Excel($options);
    }

    /**
     * 
     * @access protected
     * 
     * @param array $options with optional parameters: title, subject, description, author
     */
    protected function __construct(array $options)
    {
        //get PHPExcel instance
        $this->_excel = new PHPExcel();

        //set options
        $this->set_options($options);
    }

    /**
     * call PHPExcel spreadsheet's function if not exist
     * 
     * @access public
     * @param string $method_name method name
     * @param mixed $arguments arguments
     */
    public function __call($method_name, $arguments)
    {
        $this->_excel->$method_name($arguments);
    }

    /**
     * @return Spreadsheet 
     */
    protected function set_properties()
    {
        $this->_excel->getProperties()
                ->setCreator($this->options['author'])
                ->setTitle($this->options['title'])
                ->setSubject($this->options['subject'])
                ->setDescription($this->options['description']);

        return $this;
    }

    /**
     * Add/Update options
     * @param Array $options
     * @return Spreadsheet 
     */
    protected function set_options(array $options) {
        $this->options = Arr::merge($this->options, $options);
        return $this;
    }

    /**
     * Get options
     * 
     * @access public
     * @return array options
     */
    public function get_options() {
        return $this->options;
    }

    public function set_active_sheet($index)
    {
        return $this->_excel->setActiveSheetIndex($index);
    }

    public function get_active_sheet()
    {
        return $this->_excel->getActiveSheet();
    }

    public function get_all_sheets()
    {
        return $this->_excel->getAllSheets();
    }

    public function set_data(array $datas, $multi_sheet = FALSE)
    {
        $dimension = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 
            'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
        $worksheet = $this->get_active_sheet();
        $_header = Kohana::$config->load('table_header')->as_array();
        $datas = array_values($datas);
        $excel = array();

        //array for witholding muti table
        if(empty($datas['0']['0']))
        {
            $excel[] = $datas;
        }
        if(empty($excel))
        {
            $excel = $datas;
        }

        $global_row = 0;
        $column = 1;
        $count = count($excel);   

        for($i=0;$i<$count;$i++)
        {
            $data =  $excel[$i];
            if(!empty($data['header']))
            {
                $global_row++;
                $column = 0;
                foreach ($data['header'] as $key => $val) 
                {
                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($column) . ($global_row);
                    $worksheet->setCellValue($coordinates, $key);
                    $worksheet->getStyle($coordinates)->getFont()->setBold(true);
                    $worksheet->getStyle($coordinates)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('EEEEEE');
                    $worksheet->getStyle($coordinates)->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($column+1) . ($global_row);
                    $worksheet->setCellValue($coordinates, $val);
                    $worksheet->getStyle($coordinates)->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                    $worksheet->getStyle($coordinates)->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                }       
            }
            if(!empty($data['0']) && is_array($data['0']))
            {
                $global_row++;
                $value = array_keys($data['0']);
                foreach ($value as $column => $val) 
                {
                    if(is_numeric($val))
                        break;
                    //change the header name
                    if(!empty($_header[$val]))
                    {
                        $val = $_header[$val];
                    }
                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($column) . ($global_row);

                    $worksheet->setCellValue($coordinates, $val);
                    $worksheet->getStyle($coordinates)->getFont()->setBold(true);
                    $worksheet->getStyle($coordinates)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('EEEEEE');
                    $worksheet->getStyle($coordinates)->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                    $worksheet->getRowDimension($global_row)->setRowHeight(30);
                    if(!empty($dimension[$column]))
                    {
                        $len = 12;
                        if($column<=10)
                        {
                            $len = 18;
                        }
                        $worksheet->getColumnDimension($dimension[$column])->setWidth($len); 
                   }
                }
            }
            foreach ($data as $row => $value) 
            {
                if(!is_numeric($row))
                    continue;
                //for the marketing overview
                if (!is_array($value)) 
                {

                    //for header
                    if($global_row==0)
                        $global_row++;
                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($row) . ($global_row);
                    $worksheet->setCellValue($coordinates, $value);
                    continue;
                }


                $global_row++;
                
                $value = array_values($value);
                foreach ($value as $column => $val) 
                {
                    if(is_array($val))continue;
                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($column) . ($global_row);
                    $worksheet->setCellValue($coordinates, $val);
                }
            }

            if(!empty($data['tailer']))
            {
                $column = 1;
                foreach ($data['tailer'] as $key => $val) 
                {
                    $global_row++;
                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($column) . ($global_row);
                    $worksheet->setCellValue($coordinates, $key);
                    $worksheet->getStyle($coordinates)->getFont()->setBold(true);
                    $worksheet->getStyle($coordinates)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('EEEEEE');
                    $worksheet->getStyle($coordinates)->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                    $coordinates = PHPExcel_Cell::stringFromColumnIndex($column+1) . ($global_row);
                    $worksheet->setCellValue($coordinates, $val);
                }
                
                    
            }
            $global_row++;
        }
        return;
    }

    /**
     * Send spreadsheet to browser without save to a file
     * @return void 
     */
    public function send($filename = 'file', $type = 'csv')
    {
        $this->set_properties();

        $objWriter = PHPExcel_IOFactory::createWriter($this->_excel, 'Excel2007');
        ob_end_clean();
        //$objWriter->setDelimiter(',');
        //$objWriter->setEnclosure('');
        //$objWriter->setUseBOM(TRUE);
        //$objWriter->setLineEnding("\r\n");
        //$objWriter->setSheetIndex(0);

        //$objwriter = new PHPExcel_Writer_Excel2007($this->_excel);  
        //header('Content-Type: text/csv');
         header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
        header('Cache-Control: max-age=0');
        // $objWriter = PHPExcel_IOFactory::createWriter($this->_excel, 'Excel5');

        // // header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // header('Content-type: application/vnd.ms-excel');
        // header('Content-Disposition: attachment; filename="file.xls"');

        return $objWriter->save('php://output');

        $response = Response::factory();
        $response->send_file(
                $this->save(), $this->options['name'] . '.' . $this->exts[$this->options['format']], // filename
                array(
            'mime_type' => $this->mimes[$this->options['format']]
        ));
    }
}
<?php
namespace strong2much\excel;

use Yii;
use yii\base\Component;
use yii\helpers\ArrayHelper;
use yii\helpers\FileHelper;
use PHPExcel;
use PHPExcel_Style_Alignment;
use PHPExcel_IOFactory;

/**
 * Manager class that handles with Excel.
 *
 * @author   Denis Tatarnikov <tatarnikovda@gmail.com>
 */
class Excel extends Component
{
    /**
     * @var string default format (@see static::$map)
     */
    public $defaultFormat = 'Excel2007';

    /**
     * @var array default properties
     */
    public $properties = [
        'title' => 'Yii2-Excel',
        'subject' => '',
        'description' => '',
        'keywords' => 'excel php',
        'category' => 'report',
        'creator' => 'strong2much',
        'company' => 'Ausenlab',
    ];

    /**
     * @var array map of format => extension
     */
    protected static $map = [
        'OpenDocument' => 'ods',
        'CSV' => 'csv',
        'HTML' => 'html',
        'Excel5' => 'xls',
        'Excel2007' => 'xlsx',
    ];

    /**
     * New instance
     * @param bool $setDefaultProperties initialize instance with default properties
     * @return PHPExcel
     */
    public function getInstance($setDefaultProperties = true)
    {
        $phpExcel = new PHPExcel();
        if($setDefaultProperties) {
            $phpExcel->getProperties()
                ->setCreator(ArrayHelper::getValue($this->properties, 'creator', ''))
                ->setLastModifiedBy(ArrayHelper::getValue($this->properties, 'creator', ''))
                ->setTitle(ArrayHelper::getValue($this->properties, 'title', ''))
                ->setSubject(ArrayHelper::getValue($this->properties, 'subject', ''))
                ->setDescription(ArrayHelper::getValue($this->properties, 'description', ''))
                ->setKeywords(ArrayHelper::getValue($this->properties, 'keywords', ''))
                ->setCategory(ArrayHelper::getValue($this->properties, 'category', ''))
                ->setCompany(ArrayHelper::getValue($this->properties, 'company', ''));
        }

        return $phpExcel;
    }

    /**
     * @param int $index the column index number (1-based)
     * @return string an excel column name
     */
    public function columnName($index)
    {
        $i = $index - 1;
        if ($i >= 0 && $i < 26) {
            return chr(ord('A') + $i);
        }

        if($i<0) {
            return 'A';
        }

        return (self::columnName($i / 26)).(self::columnName($i % 26 + 1));
    }

    /**
     * Writes array (one-dimension) of data to excel
     * @param PHPExcel $phpExcel reference to phpExcel
     * @param int $rowIndex index of row (1-based)
     * @param array $data data to write
     * @param array $style default style for the whole row
     */
    public function writeRow(PHPExcel &$phpExcel, $rowIndex, array $data, $style = [])
    {
        $activeSheet = $phpExcel->getActiveSheet();

        if($rowIndex<=0) {
            $rowIndex = 1;
        }

        $columnIndex = 1;
        foreach($data as $item) {
            $cell = $this->columnName($columnIndex).$rowIndex;
            $activeSheet->setCellValue($cell, $item);
            if(!empty($style)) {
                $activeSheet->getStyle($cell)->applyFromArray($style);
            }
            $columnIndex++;
        }
    }

    /**
     * Writes header data to excel. Header will be the at the first row
     * @param PHPExcel $phpExcel reference to phpExcel
     * @param array $data data to write
     * @param array $style additional style for the header
     */
    public function writeHeaderRow(PHPExcel &$phpExcel, array $data, $style = [])
    {
        $this->writeRow($phpExcel, 1, $data, ArrayHelper::merge([
            'font' => [
                'bold' => true
            ],
            'alignment' => [
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
            ],
        ], $style));

        $sheet = $phpExcel->getActiveSheet();
        $sheet->mergeCells();
        $sheet->freezePane($this->columnName(1).'2'); //A2
    }

    /**
     * Download file from PHPExcel instance
     * @param PHPExcel $phpExcel reference to phpExcel
     * @param string $fileName file name
     * @param string $format file export format
     */
    public function downloadFile(PHPExcel &$phpExcel, $fileName, $format = 'Excel2007')
    {
        if(!in_array($format, array_keys(static::$map))) {
            $format = $this->defaultFormat;
        }
        $fileName .= '.'.static::$map[$format];

        header('Content-Type: '.FileHelper::getMimeTypeByExtension($fileName));
        header('Content-Disposition: attachment;filename="'.$fileName.'"');
        header('Cache-Control: max-age=0');
        header('Cache-Control: max-age=1'); // If you're serving to IE 9, then the following may be needed
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $writer = PHPExcel_IOFactory::createWriter($phpExcel, $format);
        $writer->save('php://output');

        Yii::$app->end();
    }
}

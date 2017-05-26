<?php

namespace Mtrajano\SimpleExcel;

class Workbook extends \PHPExcel
{
    const WRITE_Excel2007 = 0;
    const WRITE_Excel5 = 1;
    const WRITE_CSV = 2;
    const WRITE_HTML = 3;
    const WRITE_PDF = 4;

    public function __construct()
    {
        parent::__construct();

        //remove default PHPExcel_Worksheet created and replace it with our version instead
        $this->removeSheetByIndex(0);
        $this->createSheet();
    }

    public function createSheet($iSheetIndex = null)
    {
        $newSheet = new Worksheet($this);
        $this->addSheet($newSheet, $iSheetIndex);
        return $newSheet;
    }

    public function save($filename, $type = self::WRITE_Excel2007)
    {
        $writer = $this->getWriter($type);

        $writer->save($filename);
    }

    public function getWriter($type)
    {
        switch($type) {
            case self::WRITE_Excel2007:
                return new \PHPExcel_Writer_Excel2007($this);
            case self::WRITE_Excel5:
                return new \PHPExcel_Writer_Excel5($this);
            case self::WRITE_CSV:
                return new \PHPExcel_Writer_CSV($this);
            case self::WRITE_HTML:
                return new \PHPExcel_Writer_HTML($this);
            case self::WRITE_PDF:
                return new \PHPExcel_Writer_PDF($this);
            default:
                throw new \PHPExcel_Exception("Invalid writer type passed");
        }
    }
}
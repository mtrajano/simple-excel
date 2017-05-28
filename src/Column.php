<?php

namespace Mtrajano\SimpleExcel;

class Column extends \PHPExcel_Worksheet_Column
{
    public static function getNumericIndex($column)
    {
        return \PHPExcel_Cell::columnIndexFromString($column) - 1;
    }
}
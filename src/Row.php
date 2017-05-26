<?php

namespace Mtrajano\SimpleExcel;

class Row extends \PHPExcel_Worksheet_Row implements \ArrayAccess
{
    private $worksheet;

    private $row;

    public function __construct(\PHPExcel_Worksheet $parent = null, $rowIndex = 1)
    {
        parent::__construct($parent, $rowIndex);

        $this->worksheet = $parent;
        $this->row = $rowIndex;
    }

    public function offsetExists($offset) {}

    public function offsetGet($offset)
    {
        return $this->getColumn($offset);
    }

    public function offsetSet($offset, $value)
    {
        return $this->getColumn($offset)->setValue($value);
    }

    public function offsetUnset($offset) {}

    public function getColumn($column)
    {
        if (is_null($column)) { //array push
            $column = $this->worksheet->getHighestColumn($this->row);
            $column = \PHPExcel_Cell::columnIndexFromString($column);
        } else if (is_string($column)) {
            $column = \PHPExcel_Cell::columnIndexFromString($column);
        }

        return $this->worksheet->getCellByColumnAndRow($column, $this->row);
    }
}
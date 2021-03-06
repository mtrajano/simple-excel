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

    public function offsetExists($offset)
    {
        return $this->getColumn($offset)->getValue() !== null;
    }

    public function offsetGet($offset)
    {
        return $this->getColumn($offset);
    }

    public function offsetSet($offset, $value)
    {
        $this->getColumn($offset)->setValue($value);
    }

    public function offsetUnset($offset)
    {
        $this->getColumn($offset)->setValue(null);
    }

    public function getColumn($column)
    {
        if (is_null($column)) { //array push
            $column = $this->getNextEmptyColumn();
        } else if (is_string($column)) {
            $column = Column::getNumericIndex($column);
        }

        return $this->worksheet->getCellByColumnAndRow($column, $this->row);
    }

    public function getNextEmptyColumn()
    {
        $column = $this->worksheet->getHighestColumn($this->row);

        if ($column === 'A' && $this->worksheet->getCell($column . $this->row)->getValue() === null) { //empty row
            return 0;
        }

        return Column::getNumericIndex($column) + 1;
    }
}
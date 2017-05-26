<?php

namespace Mtrajano\SimpleExcel;

class Worksheet extends \PHPExcel_Worksheet implements \ArrayAccess
{
    public function offsetExists($offset) {}

    public function offsetGet($offset)
    {
        return $this->getRow($offset);
    }

    public function offsetSet($offset, $value) {}

    public function offsetUnset($offset) {}

    public function getRow($row)
    {
        return new Row($this, $row);
    }
}
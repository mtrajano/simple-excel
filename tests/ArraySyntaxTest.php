<?php

namespace Mtrajano\SimpleExcel\Tests;

use PHPUnit\Framework\TestCase;
use Mtrajano\SimpleExcel\Workbook;

class ArraySyntaxTest extends TestCase
{
    protected $worksheet;

    public function setUp()
    {
        $workbook = new Workbook();
        $this->worksheet = $workbook->getActiveSheet();
    }

    public function test_string_column_access()
    {
        $this->worksheet[1]['A'] = 'Some test value';

        $this->assertEquals('Some test value', $this->worksheet->getCell('A1')->getValue());
    }

    public function test_numeric_column_access()
    {
        $this->worksheet[1][0] = 'Some other value';

        $this->assertEquals('Some other value', $this->worksheet->getCell('A1')->getValue());
    }

    public function test_array_push()
    {
        $values = ['abc', 'def', 'ghi'];

        foreach ($values as $val) {
            $this->worksheet[1][] = $val;
        }

        $this->assertEquals('abc', $this->worksheet->getCell('A1')->getValue());
        $this->assertEquals('def', $this->worksheet->getCell('B1')->getValue());
        $this->assertEquals('ghi', $this->worksheet->getCell('C1')->getValue());
    }
}
<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\Tests;

use PHPUnit\Framework\TestCase;
use Slam\PhpSpreadsheetHelper\CellStyle\Text;
use Slam\PhpSpreadsheetHelper\Column;
use Slam\PhpSpreadsheetHelper\ColumnCollection;
use Slam\PhpSpreadsheetHelper\Exception;

final class ColumnCollectionTest extends TestCase
{
    private Column $column;
    private ColumnCollection $collection;

    protected function setUp(): void
    {
        $this->column = new Column('foo', 'Foo', 10, new Text());
        $this->collection = new ColumnCollection(...[$this->column]);
    }

    public function testBaseFunctionalities(): void
    {
        self::assertArrayHasKey('foo', $this->collection);
        self::assertSame($this->column, $this->collection['foo']);
    }
}

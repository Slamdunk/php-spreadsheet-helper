<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\Tests;

use ArrayIterator;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\TestCase;
use Slam\PhpSpreadsheetHelper\Exception;
use Slam\PhpSpreadsheetHelper\Table;

final class TableTest extends TestCase
{
    private const EXCEL_DATA = ['a', 'b'];
    private Spreadsheet $phpExcel;
    private Worksheet $activeSheet;
    private Table $table;

    protected function setUp(): void
    {
        $this->phpExcel    = new Spreadsheet();
        $this->activeSheet = new Worksheet($this->phpExcel, 'sheet 1');
        $this->table = new Table(
            $this->activeSheet,
            3,
            12,
            'My Heading',
            self::EXCEL_DATA
        );
    }

    public function testRowAndColumn(): void
    {
        self::assertSame($this->activeSheet, $this->table->getActiveSheet());
        self::assertSame('My Heading', $this->table->getHeading());
        self::assertSame(self::EXCEL_DATA, $this->table->getData());

        self::assertNull($this->table->getDataRowStart());

        $this->table->incrementRow();
        $this->table->flagDataRowStart();
        $this->table->incrementRow();

        self::assertSame(3, $this->table->getRowStart());
        self::assertSame(4, $this->table->getDataRowStart());
        self::assertSame(5, $this->table->getRowEnd());
        self::assertSame(5, $this->table->getRowCurrent());

        $this->table->incrementColumn();
        $this->table->incrementColumn();

        self::assertSame(12, $this->table->getColumnStart());
        self::assertSame(14, $this->table->getColumnEnd());
        self::assertSame(14, $this->table->getColumnCurrent());

        $this->table->resetColumn();

        self::assertSame(12, $this->table->getColumnStart());
        self::assertSame(14, $this->table->getColumnEnd());
        self::assertSame(12, $this->table->getColumnCurrent());

        $this->table->setCount(0);
        self::assertCount(0, $this->table);
        self::assertTrue($this->table->isEmpty());

        $this->table->setCount(5);
        self::assertCount(5, $this->table);
        self::assertFalse($this->table->isEmpty());

        self::assertTrue($this->table->getFreezePanes());
        $this->table->setFreezePanes(false);
        self::assertFalse($this->table->getFreezePanes());

        self::assertNull($this->table->getWrittenColumnTitles());
        $columns = [
            'column_1' => 'Name',
            'column_2' => 'Surname',
        ];
        $this->table->setWrittenColumnTitles($columns);
        self::assertSame($columns, $this->table->getWrittenColumnTitles());
    }

    public function testTableCountMustBeSet(): void
    {
        $this->expectException(Exception\RuntimeException::class);

        $this->table->count();
    }

    public function testSplitTableIfNeeded(): void
    {
        $this->table->setFreezePanes(false);
        $newTable = $this->table->splitTableOnNewWorksheet();

        self::assertNotSame($this->table, $newTable);

        // The starting row must be the first of the new sheet
        self::assertSame(0, $newTable->getRowStart());
        self::assertSame(0, $newTable->getRowEnd());
        self::assertSame(0, $newTable->getRowCurrent());

        // The starting column must be the same of the previous sheet
        self::assertSame(12, $newTable->getColumnStart());
        self::assertSame(12, $newTable->getColumnEnd());
        self::assertSame(12, $newTable->getColumnCurrent());

        self::assertSame($this->table->getFreezePanes(), $newTable->getFreezePanes());
    }

    public function testFontRowAttributes(): void
    {
        self::assertSame(8, $this->table->getFontSize());
        self::assertNull($this->table->getRowHeight());
        self::assertFalse($this->table->getTextWrap());

        $this->table->setFontSize($fontSize = \mt_rand(10, 100));
        $this->table->setRowHeight($rowHeight = \mt_rand(10, 100));
        $this->table->setTextWrap(true);

        self::assertSame($fontSize, $this->table->getFontSize());
        self::assertSame($rowHeight, $this->table->getRowHeight());
        self::assertTrue($this->table->getTextWrap());
    }
}

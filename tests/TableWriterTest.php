<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\Tests;

use ArrayIterator;
use PhpOffice\PhpSpreadsheet;
use PHPUnit\Framework\TestCase;
use Slam\PhpSpreadsheetHelper\Column;
use Slam\PhpSpreadsheetHelper\CellStyle;
use Slam\PhpSpreadsheetHelper\ColumnCollection;
use Slam\PhpSpreadsheetHelper\Table;
use Slam\PhpSpreadsheetHelper\TableWriter;

final class TableWriterTest extends TestCase
{
    private function writeAndRead(PhpSpreadsheet\Spreadsheet $source): PhpSpreadsheet\Spreadsheet
    {
        $filename = __DIR__.'/tmp/test.xlsx';
        @unlink($filename);
        (new PhpSpreadsheet\Writer\Xlsx($source))->save($filename);
        $dest = (new PhpSpreadsheet\Reader\Xlsx())->load($filename);

        return $dest;
    }

    public function testPostGenerationDetails(): void
    {
        $source = new PhpSpreadsheet\Spreadsheet();
        $heading = \uniqid('Heading_');
        $table  = new Table($source->getActiveSheet(), 3, 4, $heading, [
            ['description' => 'AAA'],
            ['description' => 'BBB'],
        ]);

        (new TableWriter())->writeTableToWorksheet($table);

        self::assertSame(3, $table->getRowStart());
        self::assertSame(7, $table->getRowEnd());

        self::assertSame(5, $table->getDataRowStart());

        self::assertSame(4, $table->getColumnStart());
        self::assertSame(5, $table->getColumnEnd());

        self::assertCount(2, $table);
        self::assertSame(['description' => 'Description'], $table->getWrittenColumnTitles());

        $sheet = $this->writeAndRead($source)->getActiveSheet();
        
        self::assertSame($heading, $sheet->getCellByColumnAndRow(4, 3)->getValue());
        self::assertSame('Description', $sheet->getCellByColumnAndRow(4, 4)->getValue());
        self::assertSame('AAA', $sheet->getCellByColumnAndRow(4, 5)->getValue());
        self::assertSame('BBB', $sheet->getCellByColumnAndRow(4, 6)->getValue());
    }

    public function testHandleEncoding(): void
    {
        $textWithSpecialCharacters = \implode(' # ', [
            '€',
            'VIA MARTIRI DELLA LIBERTà 2',
            'FISSO20+OPZ.I¢CASA EURIB 3',
            'FISSO 20+ OPZIONE I°CASA EUR 6',
            '1° MAGGIO',
            'GIÀ XXXXXXX YYYYYYYYYYY',
            'FINANZIAMENTO 13/14¬ MENSILITà',

            'A \'\\|!"£$%&/()=?^àèìòùáéíóúÀÈÌÒÙÁÉÍÓÚ<>*ç°§[]@#{},.-;:_~` Z',
        ]);
        $heading = \sprintf('%s: %s', \uniqid('Heading_'), $textWithSpecialCharacters);
        $data    = \sprintf('%s: %s', \uniqid('Data_'), $textWithSpecialCharacters);

        $source = new PhpSpreadsheet\Spreadsheet();
        $sourceSheet = $source->getActiveSheet();
        $table  = new Table($sourceSheet, 1, 1, $heading, [
            ['description' => $data],
        ]);

        (new TableWriter())->writeTableToWorksheet($table);
        $sheet = $this->writeAndRead($source)->getActiveSheet();

        self::assertSame($sourceSheet->getTitle(), $sheet->getTitle());

        // Heading
        $value = $sheet->getCell('A1')->getValue();
        self::assertSame($heading, $value);

        // Data
        $value = $sheet->getCell('A3')->getValue();
        self::assertSame($data, $value);
    }

    public function testCellStyles(): void
    {
        $source = new PhpSpreadsheet\Spreadsheet();

        $columnCollection = new ColumnCollection([
            new Column('my_text', 'Foo1', 11, new CellStyle\Text()),
            new Column('my_perc', 'Foo2', 12, new CellStyle\Percentage()),
            new Column('my_inte', 'Foo3', 13, new CellStyle\Integer()),
            new Column('my_date', 'Foo4', 14, new CellStyle\Date()),
            new Column('my_amnt', 'Foo5', 15, new CellStyle\Amount()),
            new Column('my_itfc', 'Foo6', 16, new CellStyle\ItalianFiscalCode()),
            new Column('my_nodd', 'Foo7', 14, new CellStyle\Date()),
        ]);

        $table       = new Table($source->getActiveSheet(), 2, 1, \uniqid('Heading_'), [
            [
                'my_text' => 'text',
                'my_perc' => 3.45,
                'my_inte' => 1234567.8,
                'my_date' => '2017-03-02',
                'my_amnt' => 1234567.89,
                'my_itfc' => 'AABB',
                'my_nodd' => null,
            ],
        ]);
        $table->setColumnCollection($columnCollection);

        (new TableWriter())->writeTableToWorksheet($table);
        $firstSheet = $this->writeAndRead($source)->getActiveSheet();

        $expected   = [
            'A1' => null,
            'A2' => $table->getHeading(),

            'A3' => 'Foo1',
            'B3' => 'Foo2',
            'C3' => 'Foo3',
            'D3' => 'Foo4',
            'E3' => 'Foo5',
            'F3' => 'Foo6',
            'G3' => 'Foo7',

            'A4' => 'text',
            'B4' => 3.45,
            'C4' => 1234567.8,
            'D4' => 42796,
            'E4' => 1234567.89,
            'F4' => 'AABB',
            'G4' => NULL,
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = $firstSheet->getCell($cell)->getValue();
        }
        
        self::assertSame($expected, $actual);
    }

    public function testTablePagination(): void
    {
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $phpExcel->setRowsPerSheet(6);

        $activeSheet = $phpExcel->addWorksheet('names');
        $table       = new Excel\Helper\Table($activeSheet, 1, 2, \uniqid(), new ArrayIterator([
            ['description' => 'AAA'],
            ['description' => 'BBB'],
            ['description' => 'CCC'],
            ['description' => 'DDD'],
            ['description' => 'EEE'],
        ]));

        $returnTable = $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected   = [
            'C1' => null,
            'C2' => $table->getHeading(),
            'C3' => 'Description',
            'C4' => 'AAA',
            'C5' => 'BBB',
            'C6' => 'CCC',
            'C7' => null,
        ];

        foreach ($expected as $cell => $content) {
            self::assertSame($content, $firstSheet->getCell($cell)->getValue());
        }

        $secondSheet = $phpExcel->getSheet(1);
        $expected    = [
            'C1' => $returnTable->getHeading(),
            'C2' => 'Description',
            'C3' => 'DDD',
            'C4' => 'EEE',
            'C5' => null,
        ];

        foreach ($expected as $cell => $content) {
            self::assertSame($content, $secondSheet->getCell($cell)->getValue());
        }

        self::assertStringContainsString('names (', $firstSheet->getTitle());
        self::assertStringContainsString('names (', $secondSheet->getTitle());
    }

    public function testEmptyTable(): void
    {
        $emptyTableMessage = \uniqid('no_data_');

        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $phpExcel->setEmptyTableMessage($emptyTableMessage);

        $activeSheet = $phpExcel->addWorksheet(\uniqid());
        $table       = new Excel\Helper\Table($activeSheet, 0, 0, \uniqid(), new ArrayIterator([]));

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected   = [
            'A1' => $table->getHeading(),
            'A2' => null,
            'A3' => $emptyTableMessage,
            'A4' => null,
        ];

        foreach ($expected as $cell => $content) {
            self::assertSame($content, $firstSheet->getCell($cell)->getValue());
        }
    }

    public function testFontRowAttributesUsage(): void
    {
        $phpExcel    = new Excel\Helper\TableWorkbook($this->filename);
        $activeSheet = $phpExcel->addWorksheet(\uniqid());
        $table       = new Excel\Helper\Table($activeSheet, 0, 0, \uniqid(), new ArrayIterator([
            [
                'name'    => 'Foo',
                'surname' => 'Bar',
            ],
            [
                'name'    => 'Baz',
                'surname' => 'Xxx',
            ],
        ]));

        $table->setFontSize(12);
        $table->setRowHeight(33);
        $table->setTextWrap(true);

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $cell       = $firstSheet->getCell('A3');
        $style      = $cell->getStyle();

        self::assertSame('Foo', $cell->getValue());
        self::assertSame(12, (int) $style->getFont()->getSize());
        self::assertSame(33, (int) $firstSheet->getRowDimension($cell->getRow())->getRowHeight());
        self::assertTrue($style->getAlignment()->getWrapText());
    }

    /**
     * @dataProvider provideColumnStringFromIndexCases
     */
    public function testColumnStringFromIndex(int $index, string $columnString): void
    {
        self::assertSame($columnString, Excel\Helper\TableWorkbook::getColumnStringFromIndex($index));
    }

    public function provideColumnStringFromIndexCases(): array
    {
        return [
            [2, 'C'],
            [3, 'D'],
            [25, 'Z'],
            [26, 'AA'],
            [33, 'AH'],
            [701, 'ZZ'],
            [703, 'AAB'],
        ];
    }

    public function testColumnStringFromIndexExpectsPositiveValues(): void
    {
        $this->expectException(Excel\Exception\InvalidArgumentException::class);

        Excel\Helper\TableWorkbook::getColumnStringFromIndex(-1);
    }

    private function getPhpExcelFromFile(string $filename): PhpSpreadsheet\Spreadsheet
    {
        return PhpSpreadsheet\IOFactory::load($filename);
    }
}

<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\Tests;

use PhpOffice\PhpSpreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PHPUnit\Framework\TestCase;
use Slam\PhpSpreadsheetHelper\CellStyle;
use Slam\PhpSpreadsheetHelper\Column;
use Slam\PhpSpreadsheetHelper\ColumnCollection;
use Slam\PhpSpreadsheetHelper\Table;
use Slam\PhpSpreadsheetHelper\TableWriter;

final class TableWriterTest extends TestCase
{
    private function writeAndRead(PhpSpreadsheet\Spreadsheet $source): PhpSpreadsheet\Spreadsheet
    {
        $filename = __DIR__ . '/tmp/test.xlsx';
        @\unlink($filename);
        (new PhpSpreadsheet\Writer\Xlsx($source))->save($filename);

        return (new PhpSpreadsheet\Reader\Xlsx())->load($filename);
    }

    public function testPostGenerationDetails(): void
    {
        $source  = new PhpSpreadsheet\Spreadsheet();
        $heading = \uniqid('Heading_');
        $table   = new Table($source->getActiveSheet(), 3, 4, $heading, [
            ['description' => 'AAA'],
            ['description' => 'BBB'],
        ]);

        (new TableWriter())->writeTableToWorksheet($table);

        self::assertSame(3, $table->getRowStart());
        self::assertSame(6, $table->getRowEnd());

        self::assertSame(5, $table->getDataRowStart());

        self::assertSame(4, $table->getColumnStart());
        self::assertSame(4, $table->getColumnEnd());

        self::assertCount(2, $table);
        self::assertSame([4 => 'description'], $table->getWrittenColumn());

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

        $source      = new PhpSpreadsheet\Spreadsheet();
        $sourceSheet = $source->getActiveSheet();
        $table       = new Table($sourceSheet, 1, 1, $heading, [
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

        $columnCollection = new ColumnCollection(...[
            new Column('disorder', 'Foo8', 11, new CellStyle\Text()),

            new Column('my_text', 'Foo1', 11, new CellStyle\Text()),
            new Column('my_perc', 'Foo2', 12, new CellStyle\Percentage()),
            new Column('my_inte', 'Foo3', 13, new CellStyle\Integer()),
            new Column('my_date', 'Foo4', 14, new CellStyle\Date()),
            new Column('my_amnt', 'Foo5', 15, new CellStyle\Amount()),
            new Column('my_itfc', 'Foo6', 16, new CellStyle\Text()),
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

                'disorder'  => 'disorder',
                'no_column' => 'no_column',
            ],
        ]);
        $table->setColumnCollection($columnCollection);

        (new TableWriter())->writeTableToWorksheet($table);
        $firstSheet = $this->writeAndRead($source)->getActiveSheet();

        $expectedContent = [
            'A1' => null,
            'A2' => $table->getHeading(),

            'A3' => 'Foo1',
            'B3' => 'Foo2',
            'C3' => 'Foo3',
            'D3' => 'Foo4',
            'E3' => 'Foo5',
            'F3' => 'Foo6',
            'G3' => 'Foo7',
            'H3' => 'Foo8',
            'I3' => 'No Column',

            'A4' => 'text',
            'B4' => 3.45,
            'C4' => 1234567.8,
            'D4' => 42796.0,
            'E4' => 1234567.89,
            'F4' => 'AABB',
            'G4' => null,
            'H4' => 'disorder',
            'I4' => 'no_column',
        ];

        $expectedDataType = [
            'A1' => DataType::TYPE_NULL,
            'A2' => DataType::TYPE_STRING,

            'A3' => DataType::TYPE_STRING,
            'B3' => DataType::TYPE_STRING,
            'C3' => DataType::TYPE_STRING,
            'D3' => DataType::TYPE_STRING,
            'E3' => DataType::TYPE_STRING,
            'F3' => DataType::TYPE_STRING,
            'G3' => DataType::TYPE_STRING,
            'H3' => DataType::TYPE_STRING,
            'I3' => DataType::TYPE_STRING,

            'A4' => DataType::TYPE_STRING,
            'B4' => DataType::TYPE_NUMERIC,
            'C4' => DataType::TYPE_NUMERIC,
            'D4' => DataType::TYPE_NUMERIC,
            'E4' => DataType::TYPE_NUMERIC,
            'F4' => DataType::TYPE_STRING,
            'G4' => DataType::TYPE_NULL,
            'H4' => DataType::TYPE_STRING,
            'I4' => DataType::TYPE_STRING,
        ];

        $expectedNumberFormat = [
            'A1' => NumberFormat::FORMAT_GENERAL,
            'A2' => NumberFormat::FORMAT_GENERAL,

            'A3' => NumberFormat::FORMAT_GENERAL,
            'B3' => NumberFormat::FORMAT_GENERAL,
            'C3' => NumberFormat::FORMAT_GENERAL,
            'D3' => NumberFormat::FORMAT_GENERAL,
            'E3' => NumberFormat::FORMAT_GENERAL,
            'F3' => NumberFormat::FORMAT_GENERAL,
            'G3' => NumberFormat::FORMAT_GENERAL,
            'H3' => NumberFormat::FORMAT_GENERAL,
            'I3' => NumberFormat::FORMAT_GENERAL,

            'A4' => NumberFormat::FORMAT_GENERAL,
            'B4' => CellStyle\Percentage::FORMATCODE,
            'C4' => CellStyle\Integer::FORMATCODE,
            'D4' => NumberFormat::FORMAT_DATE_DDMMYYYY,
            'E4' => CellStyle\Amount::FORMATCODE,
            'F4' => NumberFormat::FORMAT_GENERAL,
            'G4' => NumberFormat::FORMAT_DATE_DDMMYYYY,
            'H4' => NumberFormat::FORMAT_GENERAL,
            'I4' => NumberFormat::FORMAT_GENERAL,
        ];

        $actualContent      = [];
        $actualDataType     = [];
        $actualNumberFormat = [];
        foreach ($expectedContent as $coordinate => $content) {
            $cell                            = $firstSheet->getCell($coordinate);
            $actualContent[$coordinate]      = $cell->getValue();
            $actualDataType[$coordinate]     = $cell->getDataType();
            $actualNumberFormat[$coordinate] = $cell->getStyle()->getNumberFormat()->getFormatCode();
        }

        self::assertSame($expectedContent, $actualContent);
        self::assertSame($expectedDataType, $actualDataType);
        self::assertSame($expectedNumberFormat, $actualNumberFormat);
    }

    public function testTablePagination(): void
    {
        $source = new PhpSpreadsheet\Spreadsheet();

        $worksheet = $source->getActiveSheet();
        $worksheet->setTitle('names');
        $table = new Table($worksheet, 2, 3, \uniqid('Heading_'), [
            ['description' => 'AAA'],
            ['description' => 'BBB'],
            ['description' => 'CCC'],
            ['description' => 'DDD'],
            ['description' => 'EEE'],
        ]);

        $tables     = (new TableWriter('', 6))->writeTableToWorksheet($table);
        $sheets     = $this->writeAndRead($source)->getAllSheets();
        $firstSheet = $sheets[0];

        $expected   = [
            'C1' => null,
            'C2' => $table->getHeading(),
            'C3' => 'Description',
            'C4' => 'AAA',
            'C5' => 'BBB',
            'C6' => 'CCC',
            'C7' => null,
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = $firstSheet->getCell($cell)->getValue();
        }
        self::assertSame($expected, $actual);

        $secondSheet = $sheets[1];
        $expected    = [
            'C1' => null,
            'C2' => $tables[1]->getHeading(),
            'C3' => 'Description',
            'C4' => 'DDD',
            'C5' => 'EEE',
            'C6' => null,
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = $secondSheet->getCell($cell)->getValue();
        }
        self::assertSame($expected, $actual);

        self::assertStringContainsString('names (', $firstSheet->getTitle());
        self::assertStringContainsString('names (', $secondSheet->getTitle());
    }

    public function testEmptyTable(): void
    {
        $emptyTableMessage = \uniqid('no_data_');
        $source            = new PhpSpreadsheet\Spreadsheet();

        $table = new Table($source->getActiveSheet(), 1, 1, \uniqid(), []);

        (new TableWriter($emptyTableMessage))->writeTableToWorksheet($table);
        $firstSheet = $this->writeAndRead($source)->getActiveSheet();

        $expected   = [
            'A1' => $table->getHeading(),
            'A2' => null,
            'A3' => $emptyTableMessage,
            'A4' => null,
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = $firstSheet->getCell($cell)->getValue();
        }

        self::assertSame($expected, $actual);
    }

    public function testFontRowAttributesUsage(): void
    {
        $source = new PhpSpreadsheet\Spreadsheet();
        $table  = new Table($source->getActiveSheet(), 1, 1, \uniqid(), [
            [
                'name'    => 'Foo',
                'surname' => 'Bar',
            ],
            [
                'name'    => 'Baz',
                'surname' => 'Xxx',
            ],
        ]);

        $table->setFontSize(12);
        $table->setRowHeight(33);
        $table->setTextWrap(true);

        (new TableWriter())->writeTableToWorksheet($table);
        $firstSheet = $this->writeAndRead($source)->getActiveSheet();

        $cell       = $firstSheet->getCell('A3');
        $style      = $cell->getStyle();

        self::assertSame('Foo', $cell->getValue());
        self::assertSame(12, (int) $style->getFont()->getSize());
        self::assertSame(33, (int) $firstSheet->getRowDimension($cell->getRow())->getRowHeight());
        self::assertTrue($style->getAlignment()->getWrapText());
    }
}

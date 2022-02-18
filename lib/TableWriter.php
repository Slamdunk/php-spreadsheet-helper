<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

final class TableWriter
{
    public const GREY_MEDIUM   = 43;
    public const GREY_LIGHT    = 42;
    private const SANITIZE_MAP = [
        '&amp;' => '&',
        '&lt;' => '<',
        '&gt;' => '>',
        '&apos;' => '\'',
        '&quot;' => '"',
    ];

    private CellStyle\Text $styleIdentity;
    private ?array $formats;

    public function __construct(
        private int $rowsPerSheet = 262144,
        private string $emptyTableMessage = ''
    )
    {
//        $this->setCustomColor(self::GREY_MEDIUM,    0xCC, 0xCC, 0xCC);
//        $this->setCustomColor(self::GREY_LIGHT,     0xE8, 0xE8, 0xE8);

        $this->styleIdentity = new CellStyle\Text();
    }

    /*
    public static function getColumnStringFromIndex(int $index): string
    {
        if ($index < 0) {
            throw new Exception\InvalidArgumentException('Column index must be equal or greater than zero');
        }

        static $indexCache = [];

        if (! isset($indexCache[$index])) {
            if ($index < 26) {
                $indexCache[$index] = \chr(65 + $index);
            } elseif ($index < 702) {
                $indexCache[$index] = \chr(64 + (int) ($index / 26))
                    . \chr(65 + $index % 26)
                ;
            } else {
                $indexCache[$index] = \chr(64 + (int) (($index - 26) / 676))
                    . \chr(65 + (int) ((($index - 26) % 676) / 26))
                    . \chr(65 + $index % 26)
                ;
            }
        }

        return $indexCache[$index];
    }
     */

    /**
     * @return Table[]
     */
    public function writeTableToWorksheet(Table $table): array
    {
        $this->writeTableHeading($table);
        $tables = [$table];

        $count      = 0;
        $headingRow = true;
        foreach ($table->getData() as $row) {
            ++$count;

            if ($table->getRowCurrent() >= $this->rowsPerSheet) {
                $table    = $table->splitTableOnNewWorksheet();
                $tables[] = $table;
                $this->writeTableHeading($table);
                $headingRow = true;
            }

            if ($headingRow) {
                $this->writeColumnsHeading($table, $row);

                $headingRow = false;
            }

            $this->writeRow($table, $row);
        }

        if (\count($tables) > 1) {
            \reset($tables);
            $table      = \current($tables);
            $firstSheet = $table->getActiveSheet();
            // In Excel the maximum length for a sheet name is 30
            $originalName = \mb_substr($firstSheet->name, 0, 21);

            $sheetCounter = 0;
            $sheetTotal   = \count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->setTitle(\sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal));
            }
        }

//        if ($table->getFreezePanes()) {
//            foreach ($tables as $table) {
//                $table->getActiveSheet()->freezePanes([$table->getRowStart() + 2, 0]);
//            }
//        }
//
//        if (0 === $count) {
//            $table->incrementRow();
//            $table->getActiveSheet()->writeString($table->getRowCurrent(), $table->getColumnCurrent(), $this->emptyTableMessage);
//            $table->incrementRow();
//        }

        $table->setCount($count);

        return $tables;
    }

    private function writeTableHeading(Table $table): void
    {
        $table->resetColumn();
        $table->getActiveSheet()->setCellValueExplicitByColumnAndRow(
            $table->getColumnCurrent(),
            $table->getRowCurrent(),
            $this->sanitize($table->getHeading()),
            DataType::TYPE_STRING
        );
        $table->incrementRow();
    }

    private function writeColumnsHeading(Table $table, array $row): void
    {
        $columnCollection = $table->getColumnCollection();
        $columnKeys       = \array_keys($row);
//        $this->generateFormats($table, $columnKeys, $columnCollection);

        $table->resetColumn();
        $titles = [];
        foreach ($columnKeys as $title) {
            $width    = 10;
            $newTitle = \ucwords(\str_replace('_', ' ', $title));

            if (null !== $columnCollection && null !== ($column = $columnCollection[$title])) {
                $width    = $column->getWidth();
                $newTitle = $column->getHeading();
            }

            $table->getActiveSheet()->getColumnDimensionByColumn($table->getColumnCurrent())->setWidth($width);
            $titles[$title] = $newTitle;

            $table->incrementColumn();
        }

        $this->writeRow($table, $titles, 'title');

        $table->setWrittenColumnTitles($titles);
        $table->flagDataRowStart();
    }

    private function writeRow(Table $table, array $row, ?string $type = null): void
    {
        $table->resetColumn();
        $sheet = $table->getActiveSheet();

        foreach ($row as $key => $content) {
            $cellStyle = $this->styleIdentity;
//            $format    = null;
//            if (isset($this->formats[$key])) {
//                if (null === $type) {
//                    $type = (
//                        ($table->getRowCurrent() % 2)
//                        ? 'zebra_light'
//                        : 'zebra_dark'
//                    );
//                }
//                $cellStyle = $this->formats[$key]['cell_style'];
//                $format    = $this->formats[$key][$type];
//            }

//            $write = 'write';
//            if (\get_class($cellStyle) === \get_class($this->styleIdentity)) {
//                $write = 'writeString';
//            }

            if ('title' !== $type) {
                $content = $cellStyle->decorateValue($content);
            }
            $content = $this->sanitize($content);

            $sheet->setCellValueExplicitByColumnAndRow(
                $table->getColumnCurrent(),
                $table->getRowCurrent(),
                $content,
                DataType::TYPE_STRING
            );

            $table->incrementColumn();
        }

        if (null !== ($rowHeight = $table->getRowHeight())) {
            $sheet->getRowDimension($table->getRowCurrent())->setRowHeight($rowHeight);
        }

        $table->incrementRow();
    }

    /**
     * @param mixed $value
     */
    private function sanitize($value): string
    {
        $value = \str_replace(
            \array_keys(self::SANITIZE_MAP),
            \array_values(self::SANITIZE_MAP),
            (string) $value
        );

        return $value;
    }

    private function generateFormats(Table $table, array $titles, ?ColumnCollectionInterface $columnCollection = null): void
    {
        $this->formats = [];
        foreach ($titles as $key) {
            $header = $this->addFormat();
            $header->setColor('black');
            $header->setSize($table->getFontSize());
            $header->setBold();
            $header->setFgColor(self::GREY_MEDIUM);
            $header->setTextWrap();
            $header->setAlign('center');

            $zebraLight = $this->addFormat();
            $zebraLight->setColor('black');
            $zebraLight->setSize($table->getFontSize());
            $zebraLight->setFgColor('white');

            $zebraDark = $this->addFormat();
            $zebraDark->setColor('black');
            $zebraDark->setSize($table->getFontSize());
            $zebraDark->setFgColor(self::GREY_LIGHT);

            if ($table->getTextWrap()) {
                $zebraLight->setTextWrap();
                $zebraLight->setAlign('top');

                $zebraDark->setTextWrap();
                $zebraDark->setAlign('top');
            }

            $this->formats[$key] = [
                'cell_style'    => null,
                'title'         => $header,
                'zebra_dark'    => $zebraLight,
                'zebra_light'   => $zebraDark,
            ];

            $cellStyle = $this->styleIdentity;
            if (isset($columnCollection) && isset($columnCollection[$key])) {
                $cellStyle = $columnCollection[$key]->getCellStyle();
            }

            $cellStyle->styleCell($zebraLight);
            $cellStyle->styleCell($zebraDark);

            $this->formats[$key]['cell_style'] = $cellStyle;
        }
    }
}

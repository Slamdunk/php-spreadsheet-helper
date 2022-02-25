<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;

final class TableWriter
{
    public const COLOR_HEADER_FONT = 'FFFFFF';
    public const COLOR_HEADER_FILL = '4472C4';
    public const COLOR_ODD_FILL    = 'D9E1F2';

    public const COLUMN_DEFAULT_WIDTH = 10;

    private const SANITIZE_MAP = [
        '&amp;'  => '&',
        '&lt;'   => '<',
        '&gt;'   => '>',
        '&apos;' => '\'',
        '&quot;' => '"',
    ];

    public function __construct(
        private string $emptyTableMessage = '',
        private int $rowsPerSheet = 262144
    ) {
    }

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

            if ($table->getRowCurrent() > $this->rowsPerSheet) {
                $table->setCount($count - 1);
                $count = 1;

                $table    = $table->splitTableOnNewWorksheet();
                $tables[] = $table;
                $this->writeTableHeading($table);
                $headingRow = true;
            }

            if ($headingRow) {
                $this->writeColumnsHeading($table, $row);

                $headingRow = false;
            }

            $this->writeRow($table, $row, false);
        }
        $table->setCount($count);

        if (\count($tables) > 1) {
            \reset($tables);
            $table      = \current($tables);
            $firstSheet = $table->getActiveSheet();
            // In Excel the maximum length for a sheet name is 30
            $originalName = \mb_substr($firstSheet->getTitle(), 0, 21);

            $sheetCounter = 0;
            $sheetTotal   = \count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->setTitle(\sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal));
            }
        }

        foreach ($tables as $table) {
            $columnCollection = $table->getColumnCollection();
            foreach ($table->getWrittenColumn() as $columnIndex => $columnKey) {
                if (! isset($columnCollection[$columnKey])) {
                    continue;
                }

                $dataRowStart = $table->getDataRowStart();
                \assert(null !== $dataRowStart);
                $columnCollection[$columnKey]->getCellStyle()->styleCell($table->getActiveSheet()->getStyleByColumnAndRow(
                    $columnIndex,
                    $dataRowStart,
                    $columnIndex,
                    $table->getRowEnd()
                ));
            }
        }

        if ($table->getFreezePanes()) {
            foreach ($tables as $table) {
                $table->getActiveSheet()->freezePaneByColumnAndRow(1, 2 + $table->getRowStart());
            }
        }

        if (0 !== $tables[0]->count()) {
            $conditional = $this->getZebraStripingStyle();
            foreach ($tables as $table) {
                $activeSheet = $table->getActiveSheet();
                $activeSheet->setAutoFilterByColumnAndRow(
                    $table->getColumnStart(),
                    $table->getDataRowStart() - 1,
                    $table->getColumnEnd(),
                    $table->getRowEnd()
                );
                $activeSheet->getStyleByColumnAndRow(
                    $table->getColumnStart(),
                    $table->getDataRowStart(),
                    $table->getColumnEnd(),
                    $table->getRowEnd()
                )->setConditionalStyles([$conditional]);
                $activeSheet->setSelectedCellByColumnAndRow(
                    $table->getColumnStart(),
                    $table->getDataRowStart()
                );
            }
        } else {
            $table->incrementRow();
            $table->getActiveSheet()->setCellValueExplicitByColumnAndRow(
                $table->getColumnCurrent(),
                $table->getRowCurrent(),
                $this->emptyTableMessage,
                DataType::TYPE_STRING
            );
            $table->incrementRow();
        }

        $table->setCount($count);

        return $tables;
    }

    private function writeTableHeading(Table $table): void
    {
        $defaultStyle = $table->getActiveSheet()->getParent()->getDefaultStyle();
        $defaultStyle->getFont()->setSize($table->getFontSize());
        $defaultStyle->getAlignment()->setWrapText($table->getTextWrap());

        $table->resetColumn();
        $table->getActiveSheet()->setCellValueExplicitByColumnAndRow(
            $table->getColumnCurrent(),
            $table->getRowCurrent(),
            $this->sanitize($table->getHeading()),
            DataType::TYPE_STRING
        );

        $headingStyle = $table->getActiveSheet()->getStyleByColumnAndRow(
            $table->getColumnCurrent(),
            $table->getRowCurrent()
        );
        $headingStyle->getAlignment()->setWrapText(false);
        $headingStyle->getFont()->setSize($table->getFontSize() + 2);

        $table->incrementRow();
    }

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function writeColumnsHeading(Table $table, array $row): void
    {
        $columnCollection = $table->getColumnCollection();
        $columnKeys       = \array_keys($row);

        $table->resetColumn();
        $writtenColumn = [];
        $titles        = [];
        foreach ($columnKeys as $columnKey) {
            $width    = self::COLUMN_DEFAULT_WIDTH;
            $newTitle = \ucwords(\str_replace('_', ' ', $columnKey));

            if (null !== ($column = $columnCollection[$columnKey] ?? null)) {
                $width    = $column->getWidth();
                $newTitle = $column->getHeading();
            }

            $table->getActiveSheet()->getColumnDimensionByColumn($table->getColumnCurrent())->setWidth($width);
            $writtenColumn[$table->getColumnCurrent()] = $columnKey;
            $titles[$columnKey]                        = $newTitle;

            $table->incrementColumn();
        }

        $this->writeRow($table, $titles, true);

        $table->setWrittenColumn($writtenColumn);
        $table->flagDataRowStart();
    }

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function writeRow(Table $table, array $row, bool $isTitle): void
    {
        $table->resetColumn();
        $sheet = $table->getActiveSheet();

        foreach ($row as $key => $content) {
            $content  = $this->sanitize($content);
            $dataType = DataType::TYPE_STRING;
            if (null === $content) {
                $dataType = DataType::TYPE_NULL;
            } elseif (
                ! $isTitle
                && 0 !== ($columnCollection = $table->getColumnCollection())->count()
                && isset($columnCollection[$key])
            ) {
                $dataType = $columnCollection[$key]->getCellStyle()->getDataType();
            }

            $sheet->setCellValueExplicitByColumnAndRow(
                $table->getColumnCurrent(),
                $table->getRowCurrent(),
                $content,
                $dataType
            );

            $table->incrementColumn();
        }

        if (null !== ($rowHeight = $table->getRowHeight())) {
            $sheet->getRowDimension($table->getRowCurrent())->setRowHeight($rowHeight);
        }

        if ($isTitle) {
            $titleStyle = $sheet->getStyleByColumnAndRow(
                $table->getColumnStart(),
                $table->getRowCurrent(),
                $table->getColumnEnd(),
                $table->getRowCurrent(),
            );
            $alignment = $titleStyle->getAlignment();
            $alignment->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $alignment->setVertical(Alignment::VERTICAL_CENTER);
            $alignment->setWrapText(true);
            $font = $titleStyle->getFont();
            $font->getColor()->setARGB(self::COLOR_HEADER_FONT);
            $font->setBold(true);
            $fill = $titleStyle->getFill();
            $fill->setFillType(Fill::FILL_SOLID);
            $fill->getStartColor()->setARGB(self::COLOR_HEADER_FILL);
            $fill->getEndColor()->setARGB(self::COLOR_HEADER_FILL);
        }

        $table->incrementRow();
    }

    private function sanitize(null|float|int|string $value): ?string
    {
        if (null === $value) {
            return null;
        }

        return \str_replace(
            \array_keys(self::SANITIZE_MAP),
            \array_values(self::SANITIZE_MAP),
            (string) $value
        );
    }

    private function getZebraStripingStyle(): Conditional
    {
        $conditional = new Conditional();
        $conditional->setConditionType(Conditional::CONDITION_EXPRESSION);
        $conditional->setOperatorType(Conditional::OPERATOR_EQUAL);
        $conditional->addCondition('MOD(ROW(),2)=0');
        $style = $conditional->getStyle();
        $fill  = $style->getFill();
        $fill->setFillType(Fill::FILL_SOLID);
        $fill->getStartColor()->setARGB(self::COLOR_ODD_FILL);
        $fill->getEndColor()->setARGB(self::COLOR_ODD_FILL);

        return $conditional;
    }
}

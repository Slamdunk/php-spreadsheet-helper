<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;

final class TableWriter
{
    public const COLOR_HEADER_FONT = 'FFFFFF';
    public const COLOR_HEADER_FILL = '4472C4';
    public const COLOR_ODD_FILL    = 'D9E1F2';
    public const COLOR_ODD_BORDER  = '8EA9DB';

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
            $originalName = \mb_substr($firstSheet->getTitle(), 0, 21);

            $sheetCounter = 0;
            $sheetTotal   = \count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->setTitle(\sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal));
            }
        }

        foreach ($tables as $table) {
            $index = 0;
            foreach ($table->getColumnCollection() as $column) {
                $dataRowStart = $table->getDataRowStart();
                \assert(null !== $dataRowStart);
                $column->getCellStyle()->styleCell($table->getActiveSheet()->getStyleByColumnAndRow(
                    $index + $table->getColumnStart(),
                    $dataRowStart,
                    $index + $table->getColumnStart(),
                    $table->getRowEnd()
                ));
                ++$index;
            }
        }

        if ($table->getFreezePanes()) {
            foreach ($tables as $table) {
                $table->getActiveSheet()->freezePaneByColumnAndRow(1, 2 + $table->getRowStart());
            }
        }

        if (0 !== $count) {
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
        $defaultStyle->getAlignment()->setWrapText(true);

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
        $titles = [];
        foreach ($columnKeys as $title) {
            $width    = 10;
            $newTitle = \ucwords(\str_replace('_', ' ', $title));

            if (0 !== $columnCollection->count() && null !== ($column = $columnCollection[$title])) {
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

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function writeRow(Table $table, array $row, ?string $type = null): void
    {
        $table->resetColumn();
        $sheet = $table->getActiveSheet();

        foreach ($row as $key => $content) {
            $content  = $this->sanitize($content);
            $dataType = DataType::TYPE_STRING;
            if (null === $content) {
                $dataType = DataType::TYPE_NULL;
            } elseif (
                'title' !== $type
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

        if ('title' === $type) {
            $headingStyle = $sheet->getStyleByColumnAndRow(
                $table->getColumnStart(),
                $table->getRowCurrent(),
                $table->getColumnEnd(),
                $table->getRowCurrent(),
            );
            $headingStyle->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $font = $headingStyle->getFont();
            $font->getColor()->setARGB(self::COLOR_HEADER_FONT);
            $font->setBold(true);
            $fill = $headingStyle->getFill();
            $fill->setFillType(Fill::FILL_SOLID);
            $fill->getStartColor()->setARGB(self::COLOR_HEADER_FILL);
            $fill->getEndColor()->setARGB(self::COLOR_HEADER_FILL);
        }

        $table->incrementRow();
    }

    /**
     * @param null|float|int|string $value
     */
    private function sanitize($value): ?string
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
        $bottom = $style->getBorders()->getBottom();
        $bottom->setBorderStyle(Border::BORDER_THIN);
        $bottom->getColor()->setARGB(self::COLOR_ODD_BORDER);
        $top = $style->getBorders()->getTop();
        $top->setBorderStyle(Border::BORDER_THIN);
        $top->getColor()->setARGB(self::COLOR_ODD_BORDER);

        return $conditional;
    }
}

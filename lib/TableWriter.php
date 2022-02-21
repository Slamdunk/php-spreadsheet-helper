<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class TableWriter
{
    private const SANITIZE_MAP = [
        '&amp;'  => '&',
        '&lt;'   => '<',
        '&gt;'   => '>',
        '&apos;' => '\'',
        '&quot;' => '"',
    ];

    public function __construct(
        private int $rowsPerSheet = 262144,
        private string $emptyTableMessage = ''
    ) {
//        $this->setCustomColor(self::GREY_MEDIUM,    0xCC, 0xCC, 0xCC);
//        $this->setCustomColor(self::GREY_LIGHT,     0xE8, 0xE8, 0xE8);
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

        if (0 === $count) {
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
}

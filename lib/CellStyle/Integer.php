<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\PhpSpreadsheetHelper\CellStyleInterface;

final class Integer implements CellStyleInterface
{
    public function getDataType(): string
    {
        return DataType::TYPE_NUMERIC;
    }

    public function styleCell(Style $style): void
    {
        $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $style->getNumberFormat()->setFormatCode('#,##0');
    }
}

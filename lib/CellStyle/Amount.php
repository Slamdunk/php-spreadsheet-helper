<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\PhpSpreadsheetHelper\CellStyleInterface;

final class Amount implements CellStyleInterface
{
    public function getDataType(): string
    {
        return DataType::TYPE_NUMERIC;
    }

    public function styleCell(Style $style): void
    {
        $style->getNumberFormat()->setFormatCode('#,##0.00');
    }
}

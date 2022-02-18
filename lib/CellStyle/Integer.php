<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\PhpSpreadsheetHelper\CellStyleInterface;

final class Integer implements CellStyleInterface
{
    public function decorateValue(mixed $value): mixed
    {
        return $value;
    }

    public function styleCell(Style $format): void
    {
        $format->setNumFormat('#,##0');
        $format->setAlign('center');
    }
}

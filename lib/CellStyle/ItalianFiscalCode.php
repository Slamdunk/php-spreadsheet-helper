<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\PhpSpreadsheetHelper\CellStyleInterface;
use Slam\Excel\Pear\Writer\Format;

final class ItalianFiscalCode implements CellStyleInterface
{
    public function decorateValue(mixed $value): mixed
    {
        return $value;
    }

    public function styleCell(Style $format): void
    {
        $format->setNumFormat('00000000000');
        $format->setAlign('left');
    }
}

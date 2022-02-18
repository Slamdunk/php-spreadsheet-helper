<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Style\Style;

interface CellStyleInterface
{
    public function decorateValue(mixed $value): mixed;

    public function styleCell(Style $format): void;
}

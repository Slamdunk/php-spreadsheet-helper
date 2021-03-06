<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Style\Style;

interface CellStyleInterface
{
    public function getDataType(): string;

    public function styleCell(Style $style): void;
}

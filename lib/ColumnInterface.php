<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

interface ColumnInterface
{
    public function getKey(): string;

    public function getHeading(): string;

    public function getWidth(): int;

    public function getCellStyle(): CellStyleInterface;
}

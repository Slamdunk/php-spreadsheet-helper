<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use ArrayObject;

/**
 * @extends ArrayObject<string, Column>
 */
final class ColumnCollection extends ArrayObject
{
    public function __construct(ColumnInterface ...$columns)
    {
        parent::__construct(array_combine(array_map(static function (ColumnInterface $column): string {
            return $column->getKey();
        }, $columns), $columns));
    }
}

<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use ArrayAccess;

/**
 * @extends ArrayAccess<string, ColumnInterface>
 */
interface ColumnCollectionInterface extends ArrayAccess
{
}

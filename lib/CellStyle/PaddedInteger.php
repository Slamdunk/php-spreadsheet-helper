<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\PhpSpreadsheetHelper\ContentDecoratorInterface;

final class PaddedInteger implements ContentDecoratorInterface
{
    private int $maxLength = 0;

    public function getDataType(): string
    {
        return DataType::TYPE_NUMERIC;
    }

    public function styleCell(Style $style): void
    {
        $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $style->getNumberFormat()->setFormatCode(\str_repeat('0', $this->maxLength));
    }

    public function decorate(?string $content): ?string
    {
        if (null !== $content) {
            $this->maxLength = \max($this->maxLength, \strlen($content));
        }

        return $content;
    }
}

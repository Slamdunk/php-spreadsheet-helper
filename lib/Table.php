<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

use Countable;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

final class Table implements Countable
{
    private Worksheet $activeSheet;
    private ?int $dataRowStart = null;
    private int $rowStart;
    private int $rowEnd;
    private int $rowCurrent;
    private int $columnStart;
    private int $columnEnd;
    private int $columnCurrent;
    private string $heading;

    /**
     * @var iterable<int, array<string, null|float|int|string>>
     */
    private iterable $data;
    private ColumnCollection $columnCollection;
    private bool $freezePanes = true;
    private int $fontSize     = 10;
    private ?int $rowHeight   = null;
    private bool $textWrap    = false;

    /**
     * @var array<int, string>
     */
    private array $writtenColumn       = [];
    private ?int $count                = null;

    /**
     * @param iterable<int, array<string, null|float|int|string>> $data
     */
    public function __construct(Worksheet $activeSheet, int $row, int $column, string $heading, iterable $data)
    {
        $this->activeSheet = $activeSheet;

        $this->rowStart   =
        $this->rowEnd     =
        $this->rowCurrent =
            $row
        ;

        $this->columnStart   =
        $this->columnEnd     =
        $this->columnCurrent =
            $column
        ;

        $this->heading = $heading;

        $this->data = $data;

        $this->columnCollection = new ColumnCollection();
    }

    public function getActiveSheet(): Worksheet
    {
        return $this->activeSheet;
    }

    public function getDataRowStart(): int
    {
        \assert(null !== $this->dataRowStart);

        return $this->dataRowStart;
    }

    public function flagDataRowStart(): void
    {
        $this->dataRowStart = $this->rowCurrent;
    }

    public function getRowStart(): int
    {
        return $this->rowStart;
    }

    public function getRowEnd(): int
    {
        return $this->rowEnd;
    }

    public function getRowCurrent(): int
    {
        return $this->rowCurrent;
    }

    public function incrementRow(): void
    {
        $this->rowEnd = \max($this->rowEnd, $this->rowCurrent);
        ++$this->rowCurrent;
    }

    public function getColumnStart(): int
    {
        return $this->columnStart;
    }

    public function getColumnEnd(): int
    {
        return $this->columnEnd;
    }

    public function getColumnCurrent(): int
    {
        return $this->columnCurrent;
    }

    public function incrementColumn(): void
    {
        $this->columnEnd = \max($this->columnEnd, $this->columnCurrent);
        ++$this->columnCurrent;
    }

    public function resetColumn(): void
    {
        $this->columnCurrent = $this->columnStart;
    }

    public function getHeading(): string
    {
        return $this->heading;
    }

    /**
     * @return iterable<int, array<string, null|float|int|string>>
     */
    public function getData(): iterable
    {
        return $this->data;
    }

    public function setColumnCollection(ColumnCollection $columnCollection): void
    {
        $this->columnCollection = $columnCollection;
    }

    public function getColumnCollection(): ColumnCollection
    {
        return $this->columnCollection;
    }

    public function setFreezePanes(bool $freezePanes): void
    {
        $this->freezePanes = $freezePanes;
    }

    public function getFreezePanes(): bool
    {
        return $this->freezePanes;
    }

    public function setFontSize(int $fontSize): void
    {
        $this->fontSize = $fontSize;
    }

    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    public function setRowHeight(?int $rowHeight): void
    {
        $this->rowHeight = $rowHeight;
    }

    public function getRowHeight(): ?int
    {
        return $this->rowHeight;
    }

    public function setTextWrap(bool $textWrap): void
    {
        $this->textWrap = $textWrap;
    }

    public function getTextWrap(): bool
    {
        return $this->textWrap;
    }

    /**
     * @param array<int, string> $writtenColumn
     */
    public function setWrittenColumn(array $writtenColumn): void
    {
        $this->writtenColumn = $writtenColumn;
    }

    /**
     * @return array<int, string>
     */
    public function getWrittenColumn(): array
    {
        return $this->writtenColumn;
    }

    public function setCount(int $count): void
    {
        $this->count = $count;
    }

    public function count(): int
    {
        if (null === $this->count) {
            throw new Exception(\sprintf('%s::setCount() have not been called yet', __CLASS__));
        }

        return $this->count;
    }

    public function isEmpty(): bool
    {
        return 0 === $this->count();
    }

    public function splitTableOnNewWorksheet(): self
    {
        $newTable = new self(
            $this->activeSheet->getParent()->createSheet(),
            $this->getRowStart(),
            $this->getColumnStart(),
            $this->getHeading(),
            $this->getData()
        );
        $newTable->setColumnCollection($this->getColumnCollection());
        $newTable->setFreezePanes($this->getFreezePanes());

        return $newTable;
    }
}

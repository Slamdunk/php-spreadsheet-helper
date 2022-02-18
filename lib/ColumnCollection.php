<?php

declare(strict_types=1);

namespace Slam\PhpSpreadsheetHelper;

final class ColumnCollection implements ColumnCollectionInterface
{
    /**
     * @var array<string, ColumnInterface>
     */
    private array $columns = [];

    public function __construct(array $columns)
    {
        foreach ($columns as $column) {
            $this->addColumn($column);
        }
    }

    private function addColumn(ColumnInterface $column): void
    {
        $this->columns[$column->getKey()] = $column;
    }

    public function offsetSet(mixed $offset, mixed $value): void
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    public function offsetExists(mixed $offset): bool
    {
        return isset($this->columns[$offset]);
    }

    public function offsetUnset(mixed $offset): void
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    public function offsetGet(mixed $offset): ?ColumnInterface
    {
        return $this->columns[$offset] ?? null;
    }
}

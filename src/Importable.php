<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Common\Type;
use Illuminate\Support\Collection;
use Box\Spout\Reader\SheetInterface;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Reader\Exception\ReaderNotOpenedException;

/**
 * Trait Importable.
 *
 * @property bool $with_header
 */
trait Importable
{
    /**
     * @var int
     */
    private $sheet_number = 1;

    /**
     * @param string $path
     *
     * @return string
     */
    abstract protected function getType($path): string;

    /**
     * @param ReaderEntityFactory|WriterEntityFactory $reader_or_writer
     *
     * @return mixed
     */
    abstract protected function setOptions(&$reader_or_writer);

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @return Collection
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     *
     * @throws IOException
     */
    public function import($path, callable $callback = null): Collection
    {
        $reader = $this->reader($path);

        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheet_number != $key) {
                continue;
            }
            $collection = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return collect($collection ?? []);
    }

    /**
     * @param string        $path
     * @param callable|null $callback
     *
     * @return Collection
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     *
     * @throws IOException
     */
    public function importSheets($path, callable $callback = null): Collection
    {
        $reader = $this->reader($path);

        $collections = [];
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            $collections[] = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return new SheetCollection($collections);
    }

    /**
     * @param $path
     *
     * @return \Box\Spout\Reader\CSV\Reader|\Box\Spout\Reader\ODS\Reader|\Box\Spout\Reader\XLSX\Reader
     *
     * @throws UnsupportedTypeException
     * @throws IOException
     */
    private function reader($path)
    {
        switch ($this->getType($path)) {
            case Type::XLSX:
                $reader = ReaderEntityFactory::createXLSXReader(); // replaces ReaderFactory::create(Type::XLSX)
                break;
            case Type::CSV:
                $reader = ReaderEntityFactory::createCSVReader();  // replaces ReaderFactory::create(Type::CSV)
                break;
            case Type::ODS:
                $reader = ReaderEntityFactory::createODSReader();  // replaces ReaderFactory::create(Type::ODS)
                break;
            default:
                throw new UnsupportedTypeException('Unsupported type ' . $this->getType($path));
        }

        $this->setOptions($reader);
        $reader->open($path);
        return $reader;
    }

    /**
     * @param SheetInterface $sheet
     * @param callable|null  $callback
     *
     * @return array
     */
    private function importSheet(SheetInterface $sheet, callable $callback = null): array
    {
        $headers = [];
        $collection = [];
        $count_header = 0;

        if ($this->with_header) {
            foreach ($sheet->getRowIterator() as $k => $row) {
                $rowAsArray = $row->toArray();
                if ($k == 1) {
                    $headers = $this->toStrings($rowAsArray);
                    $count_header = count($headers);
                    continue;
                }
                if ($count_header > $count_row = count($rowAsArray)) {
                    $rowAsArray = array_merge($rowAsArray, array_fill(0, $count_header - $count_row, null));
                } elseif ($count_header < $count_row = count($rowAsArray)) {
                    $rowAsArray = array_slice($rowAsArray, 0, $count_header);
                }
                if ($callback) {
                    if ($result = $callback(array_combine($headers, $rowAsArray))) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = array_combine($headers, $rowAsArray);
                }
            }
        } else {
            foreach ($sheet->getRowIterator() as $row) {
                $rowAsArray = $row->toArray();
                if ($callback) {
                    if ($result = $callback($rowAsArray)) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = $rowAsArray;
                }
            }
        }

        return $collection;
    }

    /**
     * @param array $values
     *
     * @return array
     */
    private function toStrings($values): array
    {
        foreach ($values as &$value) {
            if ($value instanceof \Datetime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value) {
                $value = (string) $value;
            }
        }

        return $values;
    }
}

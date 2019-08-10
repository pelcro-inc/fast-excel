<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Common\Type;
use Illuminate\Support\Collection;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Writer\WriterInterface;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Exception\InvalidArgumentException;
use Box\Spout\Writer\Exception\WriterNotOpenedException;
use Box\Spout\Writer\Exception\InvalidSheetNameException;

/**
 * Trait Exportable.
 *
 * @property bool $with_header
 * @property \Illuminate\Support\Collection $data
 */
trait Exportable
{
    /**
     * @var Style
     */
    private $header_style;

    /**
     * @param string $path
     *
     * @return string
     */
    abstract protected function getType($path): string;

    /**
     * @param ReaderInterface|WriterInterface $reader_or_writer
     *
     * @return mixed
     */
    abstract protected function setOptions(&$reader_or_writer);

    /**
     * @param $path
     * @param callable|null $callback
     *
     * @return mixed
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws InvalidSheetNameException
     * @throws WriterNotOpenedException
     */
    public function export($path, callable $callback = null)
    {
        self::exportOrDownload($path, 'openToFile', $callback);

        return realpath($path) ?: $path;
    }

    /**
     * @param $path
     * @param callable|null $callback
     *
     * @return \Symfony\Component\HttpFoundation\StreamedResponse|string
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws InvalidSheetNameException
     * @throws WriterNotOpenedException
     */
    public function download($path, callable $callback = null)
    {
        if (method_exists(response(), 'streamDownload')) {
            return response()->streamDownload(function () use ($path, $callback) {
                self::exportOrDownload($path, 'openToBrowser', $callback);
            });
        }
        self::exportOrDownload($path, 'openToBrowser', $callback);

        return '';
    }

    /**
     * @param $path
     * @param $function
     * @param callable|null $callback
     *
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws WriterNotOpenedException
     * @throws InvalidSheetNameException
     */
    private function exportOrDownload($path, $function, callable $callback = null): void
    {
        switch ($this->getType($path)) {
            case Type::XLSX:
                $writer = WriterEntityFactory::createXLSXWriter(); // replaces WriterFactory::create(Type::XLSX)
                break;
            case Type::CSV:
                $writer = WriterEntityFactory::createCSVWriter();  // replaces WriterFactory::create(Type::CSV)
                break;
            case Type::ODS:
                $writer = WriterEntityFactory::createODSWriter();  // replaces WriterFactory::create(Type::ODS)
                break;
            default:
                throw new UnsupportedTypeException('Unsupported type ' . $this->getType($path));
        }

        $this->setOptions($writer);
        $writer->$function($path);

        $has_sheets = ($writer instanceof \Box\Spout\Writer\XLSX\Writer || $writer instanceof \Box\Spout\Writer\ODS\Writer);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->data instanceof SheetCollection ? $this->data : collect([$this->data]);

        foreach ($data as $key => $collection) {
            if ($collection instanceof Collection) {
                // Apply callback
                if ($callback) {
                    $collection->transform(static function ($value) use ($callback) {
                        return $callback($value);
                    });
                }
                // Prepare collection (i.e remove non-string)
                $this->prepareCollection();
                // Add header row.
                if ($this->with_header) {
                    $first_row = $collection->first();
                    $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
                    if ($this->header_style) {
                        $row = WriterEntityFactory::createRowFromArray($keys, $this->header_style);
                    } else {
                        $row = WriterEntityFactory::createRowFromArray($keys);
                    }
                    $writer->addRow($row);
                }
                foreach ($collection->toArray() as $row) {
                    $writer->addRow(WriterEntityFactory::createRowFromArray(array_values((array)$row)));
                }
            }
            if (is_string($key)) {
                $writer->getCurrentSheet()->setName($key);
            }
            if ($has_sheets && $data->keys()->last() !== $key) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
        $writer->close();
    }

    /**
     * Prepare collection by removing non string if required.
     */
    protected function prepareCollection(): void
    {
        $need_conversion = false;
        $first_row = $this->data->first();

        if (!$first_row) {
            return;
        }

        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        if ($need_conversion) {
            $this->transform();
        }
    }

    /**
     * Transform the collection.
     */
    private function transform()
    {
        $this->data->transform(static function ($data) {
            return collect($data)->map(static function ($value) {
                return is_int($value) || is_float($value) || $value === null ? (string) $value : $value;
            })->filter(static function ($value) {
                return is_string($value);
            });
        });
    }

    /**
     * @param Style $style
     *
     * @return Exportable
     */
    public function headerStyle(Style $style)
    {
        $this->header_style = $style;

        return $this;
    }
}

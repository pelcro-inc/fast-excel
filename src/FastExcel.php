<?php

namespace Rap2hpoutre\FastExcel;

use Box\Spout\Common\Type;
use Box\Spout\Reader\CSV\Reader as CSVReader;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Writer\CSV\Writer as CSVWriter;
use Box\Spout\Writer\WriterInterface;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;

/**
 * Class FastExcel.
 */
class FastExcel
{
    use Importable, Exportable;

    /**
     * @var Collection
     */
    protected $data;

    /**
     * @var bool
     */
    private $with_header = true;

    /**
     * @var
     */
    private $csv_configuration = [
        'delimiter' => ',',
        'enclosure' => '"',
        'encoding'  => 'UTF-8',
        'bom'       => true,
    ];

    /**
     * FastExcel constructor.
     *
     * @param Collection $data
     */
    public function __construct($data = null)
    {
        $this->data = $data;
    }

    /**
     * Manually set data apart from the constructor.
     *
     * @param Collection $data
     *
     * @return FastExcel
     */
    public function data($data): self
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @param $path
     *
     * @return string
     */
    protected function getType($path): ?string
    {
        if (Str::endsWith($path, Type::CSV)) {
            return Type::CSV;
        } elseif (Str::endsWith($path, Type::ODS)) {
            return Type::ODS;
        } else {
            return Type::XLSX;
        }
    }

    /**
     * @param $sheet_number
     *
     * @return $this
     */
    public function sheet($sheet_number): self
    {
        $this->sheet_number = $sheet_number;

        return $this;
    }

    /**
     * @return $this
     */
    public function withoutHeaders(): self
    {
        $this->with_header = false;

        return $this;
    }

    /**
     * @param string $delimiter
     * @param string $enclosure
     * @param string $encoding
     * @param bool $bom
     *
     * @return $this
     */
    public function configureCsv($delimiter = ',', $enclosure = '"', $encoding = 'UTF-8', $bom = false): self
    {
        $this->csv_configuration = compact('delimiter', 'enclosure', 'encoding', 'bom');

        return $this;
    }

    /**
     * @param ReaderInterface|WriterInterface $reader_or_writer
     */
    protected function setOptions(&$reader_or_writer): void
    {
        if ($reader_or_writer instanceof CSVReader || $reader_or_writer instanceof CSVWriter) {
            $reader_or_writer->setFieldDelimiter($this->csv_configuration['delimiter']);
            $reader_or_writer->setFieldEnclosure($this->csv_configuration['enclosure']);
            if ($reader_or_writer instanceof CSVReader) {
                $reader_or_writer->setEncoding($this->csv_configuration['encoding']);
            }
            if ($reader_or_writer instanceof CSVWriter) {
                $reader_or_writer->setShouldAddBOM($this->csv_configuration['bom']);
            }
        }
    }
}

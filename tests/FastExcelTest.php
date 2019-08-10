<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Rap2hpoutre\FastExcel\FastExcel;
use Box\Spout\Common\Entity\Style\Color;
use Rap2hpoutre\FastExcel\SheetCollection;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\InvalidArgumentException;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Reader\Exception\ReaderNotOpenedException;
use Box\Spout\Writer\Exception\WriterNotOpenedException;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;

/**
 * Class FastExcelTest.
 */
class FastExcelTest extends TestCase
{
    /**
     * @throws IOException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     */
    public function testImportXlsx(): void
    {
        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx');
        $this->assertEquals($this->collection(), $collection);
    }

    /**
     * @throws IOException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     */
    public function testImportCsv(): void
    {
        $original_collection = $this->collection();

        $collection = (new FastExcel())->import(__DIR__.'/test2.csv');
        $this->assertEquals($original_collection, $collection);

        $collection = (new FastExcel())->configureCsv(';')->import(__DIR__.'/test1.csv');
        $this->assertEquals($original_collection, $collection);
    }

    /**
     * @param $file
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws ReaderNotOpenedException
     * @throws UnsupportedTypeException
     * @throws WriterNotOpenedException
     */
    private function export($file): void
    {
        $original_collection = $this->collection();

        (new FastExcel(clone $original_collection))->export($file);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws WriterNotOpenedException
     * @throws ReaderNotOpenedException
     */
    public function testExportXlsx(): void
    {
        $this->export(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws WriterNotOpenedException
     * @throws ReaderNotOpenedException
     */
    public function testExportCsv(): void
    {
        $this->export(__DIR__.'/test3.csv');
    }

    /**
     * @throws IOException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     */
    public function testExcelImportWithCallback(): void
    {
        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($value) {
            return [
                'test' => $value['col1'],
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $collection
        );

        $collection = (new FastExcel())->import(__DIR__.'/test1.xlsx', function ($value) {
            return new Dumb($value['col1']);
        });
        $this->assertEquals(
            collect([new Dumb('row1 col1'), new Dumb('row2 col1'), new Dumb('row3 col1')]),
            $collection
        );
    }

    /**
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     * @throws WriterNotOpenedException
     */
    public function testExcelExportWithCallback(): void
    {
        (new FastExcel(clone $this->collection()))->export(__DIR__.'/test2.xlsx', function ($value) {
            return [
                'test' => $value['col1'],
            ];
        });
        $this->assertEquals(
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            (new FastExcel())->import(__DIR__.'/test2.xlsx')
        );
        unlink(__DIR__.'/test2.xlsx');
    }

    /**
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     * @throws WriterNotOpenedException
     */
    public function testExportMultiSheetXLSX(): void
    {
        $file = __DIR__.'/test_multi_sheets.xlsx';
        $sheets = new SheetCollection([clone $this->collection(), clone $this->collection()]);
        (new FastExcel($sheets))->export($file);
        $this->assertEquals($this->collection(), (new FastExcel())->import($file));
        unlink($file);
    }

    /**
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     * @throws WriterNotOpenedException
     */
    public function testImportMultiSheetXLSX(): void
    {
        $collections = [
            collect([['test' => 'row1 col1'], ['test' => 'row2 col1'], ['test' => 'row3 col1']]),
            $this->collection(),
        ];
        $file = __DIR__.'/test_multi_sheets.xlsx';
        $sheets = new SheetCollection($collections);
        (new FastExcel($sheets))->export($file);

        $sheets = (new FastExcel())->importSheets($file);
        $this->assertInstanceOf(SheetCollection::class, $sheets);

        $this->assertEquals($collections[0], collect($sheets->first()));
        $this->assertEquals($collections[1], collect($sheets->all()[1]));

        unlink($file);
    }

    /**
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     * @throws WriterNotOpenedException
     */
    public function testExportWithHeaderStyle(): void
    {
        $original_collection = $this->collection();
        $style = (new StyleBuilder())
           ->setFontBold()
           ->setBackgroundColor(Color::YELLOW)
           ->build();
        $file = __DIR__.'/test-header-style.xlsx';
        (new FastExcel(clone $original_collection))
            ->headerStyle($style)
            ->export($file);
        $this->assertEquals($original_collection, (new FastExcel())->import($file));

        unlink($file);
    }
}

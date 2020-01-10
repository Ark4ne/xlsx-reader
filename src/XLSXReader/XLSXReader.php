<?php

namespace Ark4ne\XLSXReader;

use BadMethodCallException;
use DOMDocument;
use DOMNode;
use Exception;
use Generator;

/**
 * Class XLSXReader
 * @package App\XLSXParser
 */
class XLSXReader
{
    /** @var array */
    const XLSX_FORMATS = [
        0 => 'General',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',
        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',
        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',
    ];

    /** @var array */
    const DATE_TIME_CHARACTERS = ['e', 'd', 'h', 'm', 's', 'yy'];

    /** @var string */
    const TYPE_BOOLEAN = 'b';

    /** @var string */
    const TYPE_DATE = 'd';

    /** @var string */
    const TYPE_SHARED_STRING = 's';

    /** @var string */
    protected $file;

    /** @var bool */
    protected $date1904 = false;

    /** @var array */
    protected $shared = [];

    /** @var array */
    protected $formats = [];

    /** @var array */
    protected $worksheets = [];

    /** @var array */
    protected $worksheet;

    /**
     * XLSXReader constructor.
     *
     * @param string $file
     */
    public function __construct(string $file)
    {
        $this->file = $file;
    }

    /**
     * Load metadata
     *
     * @throws \Exception
     */
    public function load()
    {
        $this->loadSharedString();
        $this->loadFormats();
        $this->loadWorkbook();
    }

    /**
     * Select a sheet by is id
     *
     * @param int $index
     *
     * @throws \Exception
     */
    public function selectSheetByIndex(int $index)
    {
        if (!isset($this->worksheets[$index])) {
            throw new Exception("Worksheet index: $index not found.");
        }

        $this->worksheet = $this->worksheets[$index];
    }

    /**
     * Select a sheet by is id
     *
     * @param int $id
     *
     * @throws \Exception
     */
    public function selectSheetById(int $id)
    {
        foreach ($this->worksheets as $worksheet) {
            if ($worksheet['id'] === $id) {
                $this->worksheet = $worksheet;

                return;
            }
        }

        throw new Exception("Worksheet id: $id not found.");
    }

    /**
     * Select a sheet by is name
     *
     * @param string $name
     *
     * @throws \Exception
     */
    public function selectSheetByName(string $name)
    {
        foreach ($this->worksheets as $worksheet) {
            if ($worksheet['name'] === $name) {
                $this->worksheet = $worksheet;

                return;
            }
        }

        throw new Exception("Worksheet name: $name not found.");
    }

    /**
     * Read all row for selected worksheet
     *
     * @param int $start
     * @param int|null $end
     *
     * @throws \Exception
     * @return \Generator
     */
    public function read(int $start = 0, int $end = null): Generator
    {
        if ($start < 0)
            throw new BadMethodCallException("\$start must be a positive integer.");

        $end = $end ?? INF;

        if ($start > $end)
            throw new BadMethodCallException("\$start must be less then \$end.");


        if (!isset($this->worksheet['id'])) {
            $this->selectSheetByIndex(0);
        }

        $reader = new XMLReader;

        $file = "{$this->file}#xl/worksheets/sheet{$this->worksheet['id']}.xml";

        if (!$reader->open("zip://$file")) {
            throw new Exception("Can't open file $file.");
        }

        try {
            if (!$reader->find('row')) {
                throw new Exception("Can't find any row");
            }

            $row_count = $start;

            while ($start-- && $reader->next('row')) ;

            do {
                yield $row_count++ => $this->parseRow($reader->expand());
            } while ($end >= $row_count && $reader->next('row'));
        } finally {
            $reader->close();
        }
    }

    /**
     * Retrieve a row
     *
     * @param int $index
     *
     * @throws \Exception
     * @return array|null
     */
    public function row(int $index)
    {
        foreach ($this->read($index) as $idx => $row) {
            return $row;
        }

        return null;
    }

    /**
     * Parse a DOMNode to string[]
     *
     * @param \DOMNode $row
     *
     * @return array
     */
    protected function parseRow(DOMNode $row): array
    {
        $values = [];

        /** @var \DOMNodeList|\DOMNode[] $cols */
        $cols = $row->childNodes;

        $shared = $this->shared;
        $formats = $this->formats;

        foreach ($cols as $col) {
            $attrs = $col->attributes;
            $attr_type = $attrs->getNamedItem('t');
            $type = $attr_type ? $attr_type->nodeValue : null;

            if (null === $type) {
                $attr_style = $attrs->getNamedItem('s');
                $style = $attr_style ? $attr_style->nodeValue : null;
                $type = $formats[$style]['type'] ?? null;
            }

            switch ($type) {
                case self::TYPE_SHARED_STRING:
                    // shared string
                    $value = $shared[$col->nodeValue];
                    break;
                case self::TYPE_DATE:
                    // date
                    $value = $this->parseDate((float)$col->nodeValue);
                    break;
                case self::TYPE_BOOLEAN:
                    // boolean
                    $value = (bool)$col->nodeValue;
                    break;
                default:
                    // string / number
                    $value = null;

                    foreach ($col->childNodes as $node) {
                        /** @var DOMNode $node */
                        if ($node->nodeName === 'v') {
                            $value = $node->nodeValue;
                            break;
                        }
                    }

                    if ($value === null) {
                        $value = $col->nodeValue;
                    }

                    // Check for numeric values
                    if (is_numeric($value)) {
                        if (strpos($value, '.') === false) {
                            $value = (int)$value;
                        } else {
                            $value = (float)$value;
                        }
                    }
            }

            $values[] = $value;
        }

        return $values;
    }

    /**
     * Parse date
     *
     * @param float $value
     *
     * @return false|string
     */
    protected function parseDate(float $value)
    {
        $d = floor($value); // days since 1900 or 1904
        $t = $value - $d;

        if ($this->date1904) {
            $d += 1462;
        }

        $t = (abs($d) > 0)
            ? ($d - 25569) * 86400 + round($t * 86400)
            : round($t * 86400);

        return date('Y-m-d H:i:s', $t);
    }

    /**
     * Load shared strings
     *
     * @throws \Exception
     */
    protected function loadSharedString()
    {
        $file = "{$this->file}#xl/sharedStrings.xml";

        $reader = new XMLReader();

        if (!$reader->open("zip://$file")) {
            throw new Exception("Can't open file $file.");
        }

        while ($reader->find('t') && $reader->read() && $reader->nodeType === XMLReader::TEXT) {
            $this->shared[] = $reader->value;
        }

        $reader->close();
    }

    /**
     * Load all cells formats
     *
     * @throws \Exception
     */
    protected function loadFormats()
    {
        $styles = self::XLSX_FORMATS;

        $dom = $this->getDOMParts("xl/styles.xml");

        /** @var \DOMNodeList|\DOMNode[] $num_formats */
        $num_formats = $dom->getElementsByTagName('numFmts')[0]->childNodes;

        foreach ($num_formats as $format) {
            $attrs = $format->attributes;
            $styles[$attrs->getNamedItem('numFmtId')->nodeValue] = $attrs->getNamedItem('formatCode')->nodeValue;
        }

        /** @var \DOMNodeList|\DOMNode[] $cols */
        $cols = $dom->getElementsByTagName('cellXfs')[0]->childNodes;

        foreach ($cols as $col) {
            $attrs = $col->attributes;
            $nfi = $attrs->getNamedItem('numFmtId');
            $nfi = $nfi ? $nfi->nodeValue : null;

            $style = $styles[$nfi] ?? null;
            $type = null;
            if ($nfi !== '0' && $style) {
                $test = preg_replace('((?<!\\\)\[.+?(?<!\\\)\])', '', $style);

                foreach (self::DATE_TIME_CHARACTERS as $character) {
                    if (strpos($test, $character) !== false) {
                        $type = self::TYPE_DATE;
                        break;
                    }
                }
            }

            $this->formats[] = [
                'style' => $style,
                'type' => $type,
            ];
        }
    }

    /**
     * Load worksheets, and verify date format
     *
     * @throws \Exception
     */
    protected function loadWorkbook()
    {
        $dom = $this->getDOMParts('xl/workbook.xml');

        /** @var \DOMNodeList|\DOMNode[] $sheets */
        $sheets = $dom->getElementsByTagName('sheets')[0]->childNodes;

        foreach ($sheets as $sheet) {
            $attrs = $sheet->attributes;
            $this->worksheets[] = [
                'name' => $attrs->getNamedItem('name')->nodeValue,
                'id' => (int)$attrs->getNamedItem('sheetId')->nodeValue,
            ];
        }

        $date1904 = $dom->getElementsByTagName('date1904');

        if ($date1904->length) {
            $this->date1904 = (int)$date1904->item(0)->nodeValue === 1;
        }
    }


    /**
     * @param string $path
     *
     * @throws \Exception
     * @return \DOMDocument
     */
    protected function getDOMParts(string $path): DOMDocument
    {
        $dom = new DOMDocument;

        if ($dom->load("zip://{$this->file}#$path") !== true) {
            throw new Exception("Can't open file {$this->file}#$path.");
        }

        return $dom;
    }
}

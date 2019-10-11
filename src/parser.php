<?php

namespace SimpleXlsxParser;

use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

use Iterator;
use XMLReader;
use DOMDocument;
use ZipArchive;
use SimpleXMLElement;

class Parser implements Iterator
{
  private $sharedStringsCache = [];
  private $formatsCache = [];

  private $zip = [];
  private $tmp_dir = '';

  private $_reader = null;
  private $_index = 1;
  const _TAG = 'row';

  public function __construct($inputFileName)
  {
    $this->tmp_dir = ini_get('upload_tmp_dir') ? ini_get('upload_tmp_dir') : sys_get_temp_dir();
    $this->tmp_dir .= '/romkamix_parser_pamyat_unzip';

    // Unzip
    $zip = new ZipArchive();
    $zip->open($inputFileName);
    $zip->extractTo($this->tmp_dir);

    if (file_exists($this->tmp_dir . '/xl/sharedStrings.xml'))
    {
      $sharedStrings = simplexml_load_file($this->tmp_dir . '/xl/sharedStrings.xml');

      foreach ($sharedStrings->si as $sharedString) {
        $this->sharedStringsCache[] = (string)$sharedString->t;
      }

      unset($sharedStrings);
    }

    if (file_exists($this->tmp_dir . '/xl/styles.xml'))
    {
      $styles = simplexml_load_file($this->tmp_dir . '/xl/styles.xml');

      $customFormats = array();

      if ($styles->numFmts)
      {
        foreach ($styles->numFmts->numFmt as $numFmt)
        {
          $customFormats[(int) $numFmt['numFmtId']] = (string)$numFmt['formatCode'];
        }
      }

      if ($styles->cellXfs)
      {
        foreach ($styles->cellXfs->xf as $xf)
        {
          $numFmtId = (int) $xf['numFmtId'];

          if (isset($customFormats[$numFmtId])) {
            $this->formatsCache[] = $customFormats[$numFmtId];
            continue;
          }

          if (in_array($numFmtId, array('14'))) {
            $this->formatsCache[] = 'dd.mm.yyyy';
            continue;
          }

          $this->formatsCache[] = NumberFormat::builtInFormatCode($numFmtId);
        }
      }

      unset($styles);
      unset($customFormats);
    }

    $this->rewind();
  }

  public function current()
  {
    $row = array();

    $doc = new DOMDocument;
    $node = simplexml_import_dom($doc->importNode($this->_reader->expand(), true));

    foreach ($node->c as $cell)
    {
      $value = isset($cell->v) ? (string) $cell->v : '';

      if (isset($cell['t']) && $cell['t'] == 's')
      {
        $value = $this->sharedStringsCache[$value];
      }

      if (!empty($value) && isset($cell['s'])
          && isset($this->formatsCache[(string) $cell['s']]))
      {
        $value = NumberFormat::toFormattedString($value, $this->formatsCache[(string) $cell['s']]);
      }

      [$cellColumn, $cellRow] = Coordinate::coordinateFromString($cell['r']);
      $cellColumnIndex = Coordinate::columnIndexFromString($cellColumn);

      $row[$cellColumnIndex] = $value;
    }

    return $row;
  }

  public function key()
  {
    return $this->_index;
  }

  public function next() //: void;
  {
    $this->_reader->next(self::_TAG);
    $this->_index++;
  }

  public function rewind() //: void;
  {
    $this->_reader = new XMLReader;
    $this->_reader->open($this->tmp_dir . '/xl/worksheets/sheet1.xml');

    while ($this->_reader->read() && $this->_reader->name !== self::_TAG);

    $this->_index = 1;
  }

  public function valid() //: bool
  {
    return ($this->_reader->name === self::_TAG);
  }
}
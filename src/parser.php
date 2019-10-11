<?php

namespace SimpleXlsxParser;

use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

use XMLReader;
use ZipArchive;
use SimpleXMLElement;

class Parser
{
  private $sharedStringsCache = [];
  private $formatsCache = [];

  private $zip = [];
  private $tmp_dir = '';

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
  }

  public function rows($page = 1, $limit = 2000)
  {
    $rows = array();

    $xmlreader = new \XMLReader;
    $xmlreader->open($this->tmp_dir . '/xl/worksheets/sheet1.xml');

    $doc = new \DOMDocument;

    $row_id = 0;

    while ($xmlreader->read() && $xmlreader->name !== 'row');
    while ($xmlreader->name === 'row')
    {
      $row_id++;

      if (($row_id <= ($page - 1) * $limit))
      {
        $xmlreader->next('row');
        continue;
      }

      if ($row_id > $page * $limit)
      {
        break;
      }

      $node = simplexml_import_dom($doc->importNode($xmlreader->expand(), true));
      $row = array();

      echo $row_id . "\n";

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

        $row[$cellColumnIndex] = self::formatString($value);
      }

      $rows[] = $row;

      $xmlreader->next('row');
    }

    return $rows;
  }

  public static function formatString($str)
  {
    return trim(preg_replace('/\s+/', ' ', $str));
  }
}
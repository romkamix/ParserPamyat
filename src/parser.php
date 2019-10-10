<?php

namespace Pamyat;

use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class Parser
{
  public static function xls($inputFileName, $start = null, $limit = null)
  {
    $data = [];

    if (!is_file($inputFileName))
    {
      return [];
    }

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');

    $filter = new ReadFilter(9,15,range('G','K'));

    $reader->setReadFilter($filter);

    ini_set('memory_limit', '256M');
    $spreadsheet = $reader->load($inputFileName);

    return '';


    $tmp_dir = ini_get('upload_tmp_dir') ? ini_get('upload_tmp_dir') : sys_get_temp_dir();
    $tmp_dir .= '/romkamix_parser_pamyat_unzip';

    // Unzip
    $zip = new \ZipArchive();
    $zip->open($inputFile);
    $zip->extractTo($tmp_dir);

    $xml = simplexml_load_file($tmp_dir . '/xl/sharedStrings.xml');

    $strings = array();
    foreach ($xml->children() as $item) {
      $strings[] = (string)$item->t;
    }

    $datetype = array();
    $xml = simplexml_load_file($tmp_dir . '/xl/styles.xml');

    $formats = array();
    foreach ($xml as $item) {
      if ($item->getName() == 'numFmts') {
        foreach ($item->children() as $numFmt) {
          $formats[(string) $numFmt['numFmtId']] = (string) $numFmt['formatCode'];
        }
      }

      if ($item->getName() == 'cellXfs') {
        $i = -1;
        foreach ($item->children() as $xf) {
          $i++;

          $numFmtId = (string) $xf['numFmtId'];

          if (isset($formats[$numFmtId])) {
            $datetype[$i] = $formats[$numFmtId];
            continue;
          }

          if (in_array($numFmtId, array('14'))) {
            $datetype[$i] = 'dd.mm.yyyy';
            continue;
          }

          $datetype[$i] = NumberFormat::builtInFormatCode($numFmtId);
        }
      }
    }

    $params = array();

    $xmlreader = new \XMLReader;
    $xmlreader->open($tmp_dir . '/xl/worksheets/sheet1.xml');

    $doc = new \DOMDocument;

    while ($xmlreader->read() && $xmlreader->name !== 'row');
    while ($xmlreader->name === 'row') {
      $node = simplexml_import_dom($doc->importNode($xmlreader->expand(), true));

      $row_id = (int)$node['r'];
      echo $row_id . "\n";

      $row = array();

      foreach ($node->c as $cell) {
        $value = isset($cell->v) ? (string) $cell->v : '';

        if (isset($cell['t']) && $cell['t'] == 's') {
          $value = $strings[$value];
        }

        if (!empty($value) && isset($cell['s']) && isset($datetype[(string) $cell['s']])) {
          $value = NumberFormat::toFormattedString($value, $datetype[(string) $cell['s']]);
        }

        [$cellColumn, $cellRow] = Coordinate::coordinateFromString($cell['r']);
        $cellColumnIndex = Coordinate::columnIndexFromString($cellColumn);

        $row[$cellColumnIndex] = self::formatString($value);
      }

      $data[] = $row;

      $xmlreader->next('row');
      continue;

      // die();

			if ($row_id == 1) {
				foreach ($row as $param) {
					$param = $this->mb_ucfirst($param);

					if (empty($param)) {
						$params[] = '';
						continue;
					}

					$this->db->setQuery(
						'SELECT `param_id` '.
						'FROM `#__iws_param` '.
						'WHERE `param_name` = "' . $this->db->escape($param) . '"'
					);

					$this->db->query();

					if ($param_id = $this->db->loadAssoc()) {
						$param_id = $param_id['param_id'];
					} else {
						$this->db->setQuery(
							'INSERT INTO `#__iws_param` (`param_name`) '.
							'VALUES ("' . $this->db->escape($param) . '")'
						);
						$this->db->query();

						$param_id = $this->db->insertid();
					}

					if ($param == 'Фотография') $paramPhotoId = $param_id;

					$this->db->setQuery(
						'INSERT IGNORE INTO `#__iws_param_to_group` (`param_id`, `group_id`) '.
						'VALUES ("' . (int)$param_id . '", "' . (int)$group_id . '")'
					);
					$this->db->query();

					$params[] = $param_id;
				}

				$xmlreader->next('row');
				continue;
			}

			if (empty($params)) return;

			$object_name = md5($inputFile.$row_id);

			$this->db->setQuery(
				'SELECT `object_id` '.
				'FROM `#__iws_object` '.
				'WHERE `object_name` = "' . $this->db->escape($object_name) . '"'
			);

			$this->db->query();

			if ($object_id = $this->db->loadAssoc()) {
				$object_id = $object_id['object_id'];
			} else {
				$this->db->setQuery(
					'INSERT INTO `#__iws_object` (`object_name`) '.
					'VALUES ("' . $this->db->escape($object_name) . '")'
				);
				$this->db->query();
				$object_id = $this->db->insertid();
			}

			$this->db->setQuery(
				'INSERT IGNORE INTO `#__iws_object_to_group` (`object_id`, `group_id`) '.
				'VALUES ("' . (int)$object_id . '", "' . (int)$group_id . '")'
			);
			$this->db->query();

			// $this->db->setQuery(
				// 'DELETE FROM `#__iws_param_to_object` '.
				// 'WHERE `object_id` = "' . (int)$object_id . '"'
			// );
			// $this->db->query();

			// Reading values
      $values = array();

			foreach($params as $i => $param) {
				if (empty($param) || !isset($row[$i]) || empty($row[$i])) continue;

				if ($param == $paramPhotoId) {
					$images = glob($dirImages . '**/' . $row[$i] . '*');

					$image = array_shift($images);

					$source = realpath($image);

					if (!$source) continue;

					$dest = $dirImagesDest . md5(dirname($source));

					if (!is_dir($dest)) mkdir($dest, 0777, true);

					copy($source, $dest . '/' . basename($source));

					$row[$i] = '/images/iws/' . md5(dirname($source)) . '/' . basename($source);
				}

				$values[$param] = $row[$i];
			}

			$insert = array();

			foreach ($values as $param_id => $value) {
				$insert[] = '("' . (int)$param_id . '", "' . (int)$object_id . '", "' . $this->db->escape($value) . '")';
			}

			if (!empty($insert)) {
				$this->db->setQuery(
					'INSERT INTO `#__iws_param_to_object` (`param_id`, `object_id`, `value`) '.
					'VALUES ' . implode(', ', $insert)
				);
				$this->db->execute();
			}

			$xmlreader->next('row');
		}

    unset($xmlreader);
    unset($datetype);
    unset($formats);
    unset($strings);

    return $data;
















    // // Unzip
    // $zip = new \ZipArchive();
    // $zip->open($inputFile);
    // $zip->extractTo($tmp_dir);

    // $xml = simplexml_load_file($tmp_dir . '/xl/sharedStrings.xml');
    // $strings = array();

    // foreach ($xml->children() as $item)
    // {
    //   $strings[] = (string)$item->t;
    // }

    // $datetype = array();
    // $xml = simplexml_load_file($tmp_dir . '/xl/styles.xml');

    // $formats = array();
    // foreach ($xml as $item)
    // {
    //   if ($item->getName() == 'numFmts')
    //   {
    //     foreach ($item->children() as $numFmt)
    //     {
    //       $formats[(string) $numFmt['numFmtId']] = (string) $numFmt['formatCode'];
    //     }
    //   }

    //   if ($item->getName() == 'cellXfs')
    //   {
    //     $i = 0;
    //     foreach ($item->children() as $xf)
    //     {
    //       $datetype[$i] = NumberFormat::builtInFormatCode((string) $xf['numFmtId']);

    //       if (isset($formats[(string) $xf['numFmtId']]))
    //       {
    //         $datetype[$i] = $formats[(string) $xf['numFmtId']];
    //       }

    //       $i++;
    //     }
    //   }
    // }

    // $xmlreader = new \XMLReader;
    // $xmlreader->open($tmp_dir . '/xl/worksheets/sheet1.xml');

    // $doc = new \DOMDocument;

    // while ($xmlreader->read() && $xmlreader->name !== 'row');
    // $xmlreader->next('row');

    // $count = 0;
    // $totaltime = microtime(true);
    // while ($xmlreader->name === 'row')
    // {
    //   $start = microtime(true);

    //   if (++$count % 2000 == 0) {
    //       echo 'parsed: ' . $count . ' rows for ' . round(microtime(true) - $totaltime, 2) . ' seconds';
    //       $totaltime = $start;
    //   }

    //   $node = simplexml_import_dom($doc->importNode($xmlreader->expand(), true));

    //   $column = 0;
    //   $elements = array();

    //   foreach ($config['elements'] as $key => $element)
    //   {
    //     $cell = $node->c[$column];

    //     if (array_search($key, $elements_keys) != self::getNumByLetter($cell['r']))
    //     {
    //       $elements[$key] = array(array('value' => ''));
    //       continue;
    //     }

    //     $value = isset($cell->v) ? (string) $cell->v : '';

    //     if (isset($cell['t']) && $cell['t'] == 's')
    //     {
    //       $value = $strings[$value];
    //     }

    //     if (!empty($value) && isset($cell['s']) && isset($datetype[(string) $cell['s']]))
    //     {
    //       $value = NumberFormat::toFormattedString($value, $datetype[(string) $cell['s']]);
    //     }

    //     $elements[$key] = array(array('value' => $value));

    //     $column++;
    //   }

    //   $empty_row = true;
    //   foreach($elements as $key => $value)
    //   {
    //     if (!empty(trim($value[0]['value'])))
    //     {
    //       $empty_row = false;
    //       break;
    //     }
    //   }

    //   $item = new StdClass();
    //   $item->application_id = $this->app['id'];
    //   $item->type = $config['type'];
    //   $item->name = str_replace('.', '-', (string) microtime(true)); // Изменить на ФИО
    //   $item->alias = str_replace('.', '-', (string) microtime(true));
    //   $item->created = date('Y-m-d H:i:s');
    //   $item->modified = $item->created;
    //   $item->modified_by = $this->user->id;
    //   $item->publish_up = date('Y-m-d H:i:s', time()-2*60*60);
    //   $item->publish_down = '0000-00-00 00:00:00';
    //   $item->priority = 0;
    //   $item->hits = 0;
    //   $item->state = 1;
    //   $item->access = 1;
    //   $item->created_by = $this->user->id;
    //   $item->created_by_alias = '';
    //   $item->searchable = 1;
    //   $item->elements = json_encode($elements, JSON_FORCE_OBJECT);
    //   $item->params = $params;

    //   if (!$empty_row && $this->db->insertObject('#__zoo_item', $item))
    //   {
    //     $item_id = $this->db->insertid();

    //     $obj = new StdClass();
    //     $obj->item_id = $item_id;
    //     $obj->category_id = $category_id;
    //     $this->db->insertObject('#__zoo_category_item', $obj);
    //     unset($obj);

    //     $query = $this->db->getQuery(true);

    //     $columns = array('item_id', 'element_id', 'value');
    //     $values = array();

    //     foreach ($elements as $key => $value)
    //     {
    //       $value = trim($value[0]['value']);
    //       if (empty($value))
    //       {
    //         continue;
    //       }

    //       $values[] = $this->db->quote($item_id).', '.
    //                   $this->db->quote($key).', '.
    //                   $this->db->quote($value);
    //     }

    //     if (!empty($values))
    //     {
    //       $query->insert($this->db->quoteName('#__zoo_search_index'));
    //       $query->columns($columns);
    //       $query->values($values);
    //       $this->db->setQuery($query);
    //       $this->db->query();
    //     }
    //   }

    //   $time = microtime(true) - $start;

    //   if ($time < $delta)
    //   {
    //     usleep(($delta - $time) * 1000000);
    //   }

    //   $xmlreader->next('row');
    // }
  }

  public static function formatString($str)
  {
    return trim(preg_replace('/\s+/', ' ', $str));
  }
}
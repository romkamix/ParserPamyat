<?php

// defined('_JEXEC') or die;


// class Parser extends JApplicationCli {
// 	protected $db;

// 	function __construct() {
// 		$this->db = JFactory::getDbo();
// 	}

// 	public function execute() {
// 		$this->db->setQuery(
// 			'DELETE FROM `#__zoo_item` '.
// 			'WHERE `type` IN ("cm-soldiers", "cm-soldiers-marshevye-roty", "cm-soldiers-voinskie-chasti-arkhangelskogo-okruga", "cm-soldiers-vologzhane-pogibshie-v-plenu")'
// 		);
// 		$this->db->query();

// 		$this->db->setQuery(
// 			'DELETE FROM `#__zoo_search_index` '.
// 			'WHERE `item_id` NOT IN ('.
// 				'SELECT `id` '.
// 				'FROM `#__zoo_item` '.
// 			')'
// 		);
// 		$this->db->query();

// 		$this->db->setQuery(
// 			'DELETE FROM `#__zoo_category_item` '.
// 			'WHERE `item_id` NOT IN ('.
// 				'SELECT `id` '.
// 				'FROM `#__zoo_item` '.
// 			')'
// 		);
// 		$this->db->query();

// 		$dir = BASE.'/Базы данных/*/*';
// 		$files = glob($dir);

// 		foreach ($files as $file) {
// 			if (!is_file($file)) continue;

// 			if (!$this->parseExcel($file)) continue;

// 			$dest = BASE . '/backup/' . date('Ymd_His') . '/';

// 			if (!is_dir($dest)) mkdir($dest, 0777, true);

// 			if (copy($file, $dest . basename($file))) {
// 				unlink($file);
// 			}
// 		}
// 	}

// 	public static function getNumByLetter($letters) {
// 		$alphabet = array('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z');
// 		$c_a = count($alphabet);

// 		$index = 0;

// 		$letters = str_split(strtolower(preg_replace('/[0-9]/','',$letters)));
// 		$c = count($letters) - 1;
// 		foreach ($letters as $k => $l) {
// 			$index += pow($c_a, $c - $k)*(array_search($l, $alphabet) + 1);
// 		}

// 		return ($index - 1);
// 	}

// 	public function formatString($string) {
// 		return trim(preg_replace('/\s+/', ' ', $string));
// 	}

// 	function mb_ucfirst($string) {
// 		return mb_strtoupper(mb_substr($string, 0, 1)) . mb_substr($string, 1);
// 	}

// 	private function readRow($row) {
// 		return $row;
// 	}

// 	public function parseExcel($inputFileName) {
// 		$group_name = explode('/', dirname($inputFileName));
// 		$group_name = end($group_name);

// 		$this->db->setQuery(
// 			'SELECT `group_id` '.
// 			'FROM `#__iws_group` '.
// 			'WHERE `group_name` = "' . $this->db->escape($group_name) . '"'
// 		);

// 		$this->db->query();

// 		if ($group_id = $this->db->loadAssoc()) {
// 			$group_id = $group_id['group_id'];
// 		} else {
// 			$this->db->setQuery(
// 				'INSERT INTO `#__iws_group` (`group_name`) '.
// 				'VALUES ("' . $this->db->escape($group_name) . '")'
// 			);
// 			$this->db->query();
// 			$group_id = $this->db->insertid();
// 		}

// 		$this->db->setQuery( 'DELETE FROM `#__iws_object_to_group` WHERE `group_id` = "' . (int)$group_id . '"');
// 		$this->db->query();

// 		$this->db->setQuery( 'DELETE FROM `#__iws_param_to_group` WHERE `group_id` = "' . (int)$group_id . '"');
// 		$this->db->query();

// 		$this->db->setQuery(
// 			'DELETE FROM `#__iws_param_to_object` '.
// 			'WHERE `object_id` NOT IN ('.
// 				'SELECT DISTINCT `object_id` FROM `#__iws_object_to_group`'.
// 			')'
// 		);
// 		$this->db->query();

// 		$dirImages = BASE . '/Фотографии/';
// 		$dirImagesDest = dirname(BASE) . '/images/iws/';

// 		$paramPhotoId = 0;

// 		$dir = BASE.'/parseTmp';

// 		if (!file_exists($dir)) mkdir($dir, 0777, true);

// 		// Unzip
// 		$zip = new ZipArchive();
// 		$zip->open($inputFileName);
// 		$zip->extractTo($dir);

// 		$xml = simplexml_load_file($dir.'/xl/sharedStrings.xml');

// 		$strings = array();
// 		foreach ($xml->children() as $item) {
// 			$strings[] = (string)$item->t;
// 		}

// 		$datetype = array();
// 		$xml = simplexml_load_file($dir.'/xl/styles.xml');

// 		$formats = array();
// 		foreach ($xml as $item) {
// 			if ($item->getName() == 'numFmts') {
// 				foreach ($item->children() as $numFmt) {
// 					$formats[(string) $numFmt['numFmtId']] = (string) $numFmt['formatCode'];
// 				}
// 			}

// 			if ($item->getName() == 'cellXfs') {
// 				$i = -1;
// 				foreach ($item->children() as $xf) {
// 					$i++;

// 					$numFmtId = (string) $xf['numFmtId'];

// 					if (isset($formats[$numFmtId])) {
// 						$datetype[$i] = $formats[$numFmtId];
// 						continue;
// 					}

// 					if (in_array($numFmtId, array('14'))) {
// 						$datetype[$i] = 'dd.mm.yyyy';
// 						continue;
// 					}

// 					$datetype[$i] = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::builtInFormatCode($numFmtId);
// 				}
// 			}
// 		}

// 		$params = array();

// 		$xmlreader = new XMLReader;
// 		$xmlreader->open($dir.'/xl/worksheets/sheet1.xml');

// 		$doc = new DOMDocument;

// 		while ($xmlreader->read() && $xmlreader->name !== 'row');
// 		while ($xmlreader->name === 'row') {
// 			$node = simplexml_import_dom($doc->importNode($xmlreader->expand(), true));

// 			$row_id = (int)$node['r'];
// 			echo $row_id . "\n";

// 			$row = array();

// 			foreach ($node->c as $cell) {
// 				$value = isset($cell->v) ? (string) $cell->v : '';

// 				if (isset($cell['t']) && $cell['t'] == 's') {
// 					$value = $strings[$value];
// 				}

// 				if (!empty($value) && isset($cell['s']) && isset($datetype[(string) $cell['s']])) {
// 					$value = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString($value, $datetype[(string) $cell['s']]);
// 				}

// 				$row[self::getNumByLetter($cell['r'])] = $this->formatString($value);
// 			}

// 			if ($row_id == 1) {
// 				foreach ($row as $param) {
// 					$param = $this->mb_ucfirst($param);

// 					if (empty($param)) {
// 						$params[] = '';
// 						continue;
// 					}

// 					$this->db->setQuery(
// 						'SELECT `param_id` '.
// 						'FROM `#__iws_param` '.
// 						'WHERE `param_name` = "' . $this->db->escape($param) . '"'
// 					);

// 					$this->db->query();

// 					if ($param_id = $this->db->loadAssoc()) {
// 						$param_id = $param_id['param_id'];
// 					} else {
// 						$this->db->setQuery(
// 							'INSERT INTO `#__iws_param` (`param_name`) '.
// 							'VALUES ("' . $this->db->escape($param) . '")'
// 						);
// 						$this->db->query();

// 						$param_id = $this->db->insertid();
// 					}

// 					if ($param == 'Фотография') $paramPhotoId = $param_id;

// 					$this->db->setQuery(
// 						'INSERT IGNORE INTO `#__iws_param_to_group` (`param_id`, `group_id`) '.
// 						'VALUES ("' . (int)$param_id . '", "' . (int)$group_id . '")'
// 					);
// 					$this->db->query();

// 					$params[] = $param_id;
// 				}

// 				$xmlreader->next('row');
// 				continue;
// 			}

// 			if (empty($params)) return;

// 			$object_name = md5($inputFileName.$row_id);

// 			$this->db->setQuery(
// 				'SELECT `object_id` '.
// 				'FROM `#__iws_object` '.
// 				'WHERE `object_name` = "' . $this->db->escape($object_name) . '"'
// 			);

// 			$this->db->query();

// 			if ($object_id = $this->db->loadAssoc()) {
// 				$object_id = $object_id['object_id'];
// 			} else {
// 				$this->db->setQuery(
// 					'INSERT INTO `#__iws_object` (`object_name`) '.
// 					'VALUES ("' . $this->db->escape($object_name) . '")'
// 				);
// 				$this->db->query();
// 				$object_id = $this->db->insertid();
// 			}

// 			$this->db->setQuery(
// 				'INSERT IGNORE INTO `#__iws_object_to_group` (`object_id`, `group_id`) '.
// 				'VALUES ("' . (int)$object_id . '", "' . (int)$group_id . '")'
// 			);
// 			$this->db->query();

// 			// $this->db->setQuery(
// 				// 'DELETE FROM `#__iws_param_to_object` '.
// 				// 'WHERE `object_id` = "' . (int)$object_id . '"'
// 			// );
// 			// $this->db->query();

// 			// Reading values
// 			$values = array();
// 			foreach($params as $i => $param) {
// 				if (empty($param) || !isset($row[$i]) || empty($row[$i])) continue;

// 				if ($param == $paramPhotoId) {
// 					$images = glob($dirImages . '**/' . $row[$i] . '*');

// 					$image = array_shift($images);

// 					$source = realpath($image);

// 					if (!$source) continue;

// 					$dest = $dirImagesDest . md5(dirname($source));

// 					if (!is_dir($dest)) mkdir($dest, 0777, true);

// 					copy($source, $dest . '/' . basename($source));

// 					$row[$i] = '/images/iws/' . md5(dirname($source)) . '/' . basename($source);
// 				}

// 				$values[$param] = $row[$i];
// 			}

// 			$insert = array();

// 			foreach ($values as $param_id => $value) {
// 				$insert[] = '("' . (int)$param_id . '", "' . (int)$object_id . '", "' . $this->db->escape($value) . '")';
// 			}

// 			if (!empty($insert)) {
// 				$this->db->setQuery(
// 					'INSERT INTO `#__iws_param_to_object` (`param_id`, `object_id`, `value`) '.
// 					'VALUES ' . implode(', ', $insert)
// 				);
// 				$this->db->execute();
// 			}

// 			$xmlreader->next('row');
// 		}

// 		unset($xmlreader);
// 		unset($datetype);
// 		unset($formats);
// 		unset($strings);

// 		return true;
// 	}
// }

?>
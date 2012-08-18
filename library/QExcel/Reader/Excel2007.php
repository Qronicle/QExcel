<?php
/**
 * QExcel
 *
 * Original Excel2007 reader by PHPExcel 1.7.7 (http://www.codeplex.com/PHPExcel)
 * Stripped off all styling, referencing and formula code.
 *
 * @package     QExcel_Reader
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 */
/**
 * Excel 2007 Reader
 *
 * @package     QExcel_Reader
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-17 20:33
 * @author      ruud.seberechts
 * @author      PHPExcel
 */
class QExcel_Reader_Excel2007 extends QExcel_Reader_ReaderAbstract
{
    /**
     * Can the Excel 2003 Reader open the file?
     *
     * @param string $filename      The desired file
     * @return bool                 Readable
     * @throws Exception            Invalid file
     */
	public function canRead($filename)
	{
		// Check if file exists
		if (!file_exists($filename)) {
			throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
		}

		// Check if zip class exists
		if (!class_exists('ZipArchive')) {
			throw new Exception("ZipArchive library is not enabled");
		}

		$xl = false;
		// Load file
		$zip = new ZipArchive;
		if ($zip->open($filename) === true) {
			// check if it is an OOXML archive
			$rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels"));
			if ($rels !== false) {
				foreach ($rels->Relationship as $rel) {
					switch ($rel["Type"]) {
						case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
							if (basename($rel["Target"]) == 'workbook.xml') {
								$xl = true;
							}
							break;

					}
				}
			}
			$zip->close();
		}

		return $xl;
	}


	protected static function _castToBool($c)
    {
//		echo 'Initial Cast to Boolean<br />';
		$value = isset($c->v) ? (string) $c->v : null;
		if ($value == '0') {
			return false;
		} elseif ($value == '1') {
			return true;
		} else {
			return (bool)$c->v;
		}
		return $value;
	}	//	function _castToBool()


	protected static function _castToError($c)
    {
//		echo 'Initial Cast to Error<br />';
		return isset($c->v) ? (string) $c->v : null;;
	}	//	function _castToError()


	protected static function _castToString($c)
    {
//		echo 'Initial Cast to String<br />';
		return isset($c->v) ? (string) $c->v : null;;
	}	//	function _castToString()


	public function _getFromZipArchive(ZipArchive $archive, $fileName = '')
	{
		// Root-relative paths
		if (strpos($fileName, '//') !== false)
		{
			$fileName = substr($fileName, strpos($fileName, '//') + 1);
		}
		$fileName = PHPExcel_Shared_File::realpath($fileName);

		// Apache POI fixes
		$contents = $archive->getFromName($fileName);
		if ($contents === false)
		{
			$contents = $archive->getFromName(substr($fileName, 1));
		}

		return $contents;
	}


	/**
	 * Reads names of the worksheets from a file, without parsing the whole file to a PHPExcel object
	 *
	 * @param 	string 		$pFilename
	 * @throws 	Exception
	 */
	public function listWorksheetNames($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		$worksheetNames = array();

		$zip = new ZipArchive;
		$zip->open($pFilename);

		$rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels")); //~ http://schemas.openxmlformats.org/package/2006/relationships");
		foreach ($rels->Relationship as $rel) {
			switch ($rel["Type"]) {
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
					$xmlWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "{$rel['Target']}"));  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");

					if ($xmlWorkbook->sheets) {
						foreach ($xmlWorkbook->sheets->sheet as $eleSheet) {
							// Check if sheet should be skipped
							$worksheetNames[] = (string) $eleSheet["name"];
						}
					}
			}
		}

		$zip->close();

		return $worksheetNames;
	}


    /**
     * Load workbook
     *
     * Loads the specified Excel2003XML file
     *
     * @param string $filename  The file that should be loaded
     * @return QExcel_Workbook  The loaded workbook
     * @throws Exception        Invalid file
     */
	public function load($filename)
	{
		// Check if file exists
		if (!file_exists($filename)) {
			throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
		}

		// Initialisations
		$zip = new ZipArchive;
		$zip->open($filename);

		$rels = simplexml_load_string($this->_getFromZipArchive($zip, "_rels/.rels")); //~ http://schemas.openxmlformats.org/package/2006/relationships");
		foreach ($rels->Relationship as $rel) {
			switch ($rel["Type"]) {

				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
					$dir = dirname($rel["Target"]);
					$relsWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/_rels/" . basename($rel["Target"]) . ".rels"));  //~ http://schemas.openxmlformats.org/package/2006/relationships");
					$relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

					$sharedStrings = array();
					$xpath = self::array_item($relsWorkbook->xpath("rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings']"));
					$xmlStrings = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$xpath[Target]"));  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					if (isset($xmlStrings) && isset($xmlStrings->si)) {
						foreach ($xmlStrings->si as $val) {
							if (isset($val->t)) {
								$sharedStrings[] = PHPExcel_Shared_String::ControlCharacterOOXML2PHP( (string) $val->t );
							} elseif (isset($val->r)) {
								$sharedStrings[] = $this->_parseRichText($val);
							}
						}
					}

					$worksheets = array();
					foreach ($relsWorkbook->Relationship as $ele) {
						if ($ele["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
							$worksheets[(string) $ele["Id"]] = $ele["Target"];
						}
					}

					$styles 	= array();
					$xpath = self::array_item($relsWorkbook->xpath("rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles']"));
					$xmlStyles = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$xpath[Target]")); //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					$numFmts = null;
					if ($xmlStyles && $xmlStyles->numFmts[0]) {
						$numFmts = $xmlStyles->numFmts[0];
					}
					if (isset($numFmts) && ($numFmts !== NULL)) {
						$numFmts->registerXPathNamespace("sml", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					}

					$xmlWorkbook = simplexml_load_string($this->_getFromZipArchive($zip, "{$rel['Target']}"));  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");

					// Set base date
					if ($xmlWorkbook->workbookPr) {
						PHPExcel_Shared_Date::setExcelCalendar(PHPExcel_Shared_Date::CALENDAR_WINDOWS_1900);
						if (isset($xmlWorkbook->workbookPr['date1904'])) {
							$date1904 = (string)$xmlWorkbook->workbookPr['date1904'];
							if ($date1904 == "true" || $date1904 == "1") {
								PHPExcel_Shared_Date::setExcelCalendar(PHPExcel_Shared_Date::CALENDAR_MAC_1904);
							}
						}
					}

					if ($xmlWorkbook->sheets) {
						foreach ($xmlWorkbook->sheets->sheet as $eleSheet) {

                            $sheetName = (string) $eleSheet["name"];

							// Check if sheet should be skipped
							if (isset($this->_loadSheetsOnly) && !in_array($sheetName, $this->_loadSheetsOnly)) {
								continue;
							}

                            $sheet = $this->_workbook->addSheet($sheetName);

							$fileWorksheet = $worksheets[(string) self::array_item($eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];
							$xmlSheet = simplexml_load_string($this->_getFromZipArchive($zip, "$dir/$fileWorksheet"));  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");

							if ($xmlSheet && $xmlSheet->sheetData && $xmlSheet->sheetData->row) {
								foreach ($xmlSheet->sheetData->row as $row) {

									foreach ($row->c as $c) {
										$r 					= (string) $c["r"];
										$cellDataType 		= (string) $c["t"];
										$value				= null;
										$calculatedValue 	= null;
	//
										// Read cell!
										switch ($cellDataType) {
											case "s":
	//											echo 'String<br />';
												if ((string)$c->v != '') {
													$value = $sharedStrings[intval($c->v)];

													if ($value instanceof PHPExcel_RichText) {
														$value = clone $value;
													}
												} else {
													$value = '';
												}

												break;
											case "b":
	//											echo 'Boolean<br />';
												if (!isset($c->f)) {
													$value = self::_castToBool($c);
												} else {
													// Formula
                                                    $value = self::_castToBool($c);
													/*$this->_castToFormula($c,$r,$cellDataType,$value,$calculatedValue,$sharedFormulas,'_castToBool');
													if (isset($c->f['t'])) {
														$att = array();
														$att = $c->f;
														$docSheet->getCell($r)->setFormulaAttributes($att);
													}*/
	//												echo '$calculatedValue = '.$calculatedValue.'<br />';
												}
												break;
											case "inlineStr":
	//											echo 'Inline String<br />';
												$value = $this->_parseRichText($c->is);

												break;
											case "e":
	//											echo 'Error<br />';
												if (!isset($c->f)) {
													$value = self::_castToError($c);
												} else {
													// Formula
                                                    $value = self::_castToError($c);
													//$this->_castToFormula($c,$r,$cellDataType,$value,$calculatedValue,$sharedFormulas,'_castToError');
	//												echo '$calculatedValue = '.$calculatedValue.'<br />';
												}

												break;

											default:
												if (!isset($c->f)) {
	//												// Not a formula
													$value = self::_castToString($c);
												} else {
                                                    //@todo report formula usage not supported
                                                    $value = self::_castToString($c);
												}

												break;
										}

										// Check for numeric values
										if (is_numeric($value) && $cellDataType != 's') {
											if ($value == (int)$value) $value = (int)$value;
											elseif ($value == (float)$value) $value = (float)$value;
											elseif ($value == (double)$value) $value = (double)$value;
										}

										// Rich text?
										if ($value instanceof PHPExcel_RichText) {
											$value = $value->getPlainText();
										}

                                        $row = $this->getRowFromCellName($r);
                                        $col = $this->getColFromCellName($r);

                                        $sheet->setCell($row, $col, $this->getCellValue($value, $cellDataType));
									}
								}
							}
						}
					}
				break;
			}

		}

		$zip->close();

		return $this->_workbook;
	}

	protected function _parseRichText($is = null)
    {
		$value = new PHPExcel_RichText();

		if (isset($is->t)) {
			$value->createText( PHPExcel_Shared_String::ControlCharacterOOXML2PHP( (string) $is->t ) );
		} else {
			foreach ($is->r as $run) {
				if (!isset($run->rPr)) {
					$objText = $value->createText( PHPExcel_Shared_String::ControlCharacterOOXML2PHP( (string) $run->t ) );

				} else {
					$objText = $value->createTextRun( PHPExcel_Shared_String::ControlCharacterOOXML2PHP( (string) $run->t ) );
				}
			}
		}

		return $value;
	}


	protected static function array_item($array, $key = 0)
    {
		return (isset($array[$key]) ? $array[$key] : null);
	}
}

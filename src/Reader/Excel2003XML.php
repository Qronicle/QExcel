<?php

namespace QExcel\Reader;

use Exception;
use QExcel\Cell\DataType;
use QExcel\RichText;
use QExcel\Shared\ExcelDate;
use QExcel\Shared\ExcelString;
use QExcel\Workbook;

/**
 * Excel 2003 XML Reader
 *
 * @package     QExcel_Reader
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-18 15:12
 * @author      ruud.seberechts
 */
class Excel2003XML extends AbstractReader
{
    protected $_charSet = 'UTF-8';

    public function _init()
    {
        $this->_defaultOptions = array(
            'loadSheet' => null,
        );
    }

    /**
     * Can the Excel 2003 XML Reader open the file?
     *
     * @param string $filename      The desired file
     * @return bool                 Readable
     * @throws Exception            Invalid file
     */
    public function canRead(string $filename): bool
	{

		//	Office					xmlns:o="urn:schemas-microsoft-com:office:office"
		//	Excel					xmlns:x="urn:schemas-microsoft-com:office:excel"
		//	XML Spreadsheet			xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
		//	Spreadsheet component	xmlns:c="urn:schemas-microsoft-com:office:component:spreadsheet"
		//	XML schema 				xmlns:s="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882"
		//	XML data type			xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"
		//	MS-persist recordset	xmlns:rs="urn:schemas-microsoft-com:rowset"
		//	Rowset					xmlns:z="#RowsetSchema"
		//

		$signature = array(
				'<?xml version="1.0"',
				'<?mso-application progid="Excel.Sheet"?>'
			);

		// Check if file exists
		if (!file_exists($filename)) {
			throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
		}

		// Read sample data (first 2 KB will do)
		$fh = fopen($filename, 'r');
		$data = fread($fh, 2048);
		fclose($fh);

		$valid = true;
		foreach($signature as $match) {
			// every part of the signature must be present
			if (strpos($data, $match) === false) {
				$valid = false;
				break;
			}
		}

		//	Retrieve charset encoding
		if(preg_match('/<?xml.*encoding=[\'"](.*?)[\'"].*?>/um',$data,$matches)) {
			$this->_charSet = strtoupper($matches[1]);
		}
//		echo 'Character Set is ',$this->_charSet,'<br />';

		return $valid;
	}


    /**
     * Get the sheet names of a workbook
     *
     * @abstract
     * @param string $filename  The file that should be loaded
     * @return array
     * @throws Exception
     */
	public function getSheetNames(string $filename): array
	{
		// Check if file exists
		if (!file_exists($filename)) {
			throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
		}
		if (!$this->canRead($filename)) {
			throw new Exception($filename . " is an Invalid Spreadsheet file.");
		}

		$worksheetNames = array();

		$xml = simplexml_load_file($filename);
		$namespaces = $xml->getNamespaces(true);

		$xml_ss = $xml->children($namespaces['ss']);
		foreach($xml_ss->Worksheet as $worksheet) {
			$worksheet_ss = $worksheet->attributes($namespaces['ss']);
			$worksheetNames[] = self::_convertStringEncoding((string) $worksheet_ss['Name'],$this->_charSet);
		}

		return $worksheetNames;
	}

    /**
     * Load workbook
     *
     * Loads the specified Excel2003XML file
     *
     * @param string $filename  The file that should be loaded
     * @return Workbook  The loaded workbook
     * @throws Exception        Invalid file
     */
    public function load(string $filename): Workbook
	{
        if (!$this->canRead($filename)) {
			throw new Exception($filename . " is an Invalid Spreadsheet file.");
		}

		$xml = simplexml_load_file($filename);
		$namespaces = $xml->getNamespaces(true);

        $sheetId = 0;
		$xml_ss = $xml->children($namespaces['ss']);

		foreach($xml_ss->Worksheet as $worksheet) {

            // Initialize worksheet
            $sheet = $this->_workbook->addSheet();
			$worksheet_ss = $worksheet->attributes($namespaces['ss']);

            // Set worksheet name
            if (isset($worksheet_ss['Name'])) {
                $sheet->name = self::_convertStringEncoding((string) $worksheet_ss['Name'],$this->_charSet);
            }

            // Skip worksheet when not in loadSheetsOnly
			if (($this->getLoadSheetsOnly()) && (isset($worksheet_ss['Name'])) &&
				(!in_array($worksheet_ss['Name'], $this->getLoadSheetsOnly()))) {
				continue;
			}

            $row = 0;
			if (isset($worksheet->Table->Row)) {
				foreach($worksheet->Table->Row as $rowData) {
					$row_ss = $rowData->attributes($namespaces['ss']);
					if (isset($row_ss['Index'])) {
                        $row = (integer) $row_ss['Index'] - 1;
					}

                    $column = 0;
					foreach($rowData->Cell as $cell) {

						$cell_ss = $cell->attributes($namespaces['ss']);
						if (isset($cell_ss['Index'])) {
                            $column = $cell_ss['Index']-1;
						}
						if (isset($cell->Data)) {
							$cellValue = $cellData = $cell->Data;
							$type = DataType::TYPE_NULL;
							$cellData_ss = $cellData->attributes($namespaces['ss']);
							if (isset($cellData_ss['Type'])) {
								$cellDataType = $cellData_ss['Type'];
								switch ($cellDataType)
                                {
									case 'String' :
											$cellValue = self::_convertStringEncoding($cellValue,$this->_charSet);
											$type = DataType::TYPE_STRING;
											break;
									case 'Number' :
											$type = DataType::TYPE_NUMERIC;
											$cellValue = (float) $cellValue;
											if (floor($cellValue) == $cellValue) {
												$cellValue = (integer) $cellValue;
											}
											break;
									case 'Boolean' :
											$type = DataType::TYPE_BOOL;
											$cellValue = ($cellValue != 0);
											break;
									case 'DateTime' :
											$type = DataType::TYPE_NUMERIC;
											$cellValue = ExcelDate::PHPToExcel(strtotime($cellValue));
											break;
									case 'Error' :
											$type = DataType::TYPE_ERROR;
											break;
								}
							}
                            $sheet->setCell($row, $column, $this->getCellValue($cellValue, $type));
						}
                        $column++;
					}
                    $row++;
				}
			}
			$sheetId++;
		}

		// Return
		return $this->_workbook;
	}


    protected static function identifyFixedStyleValue($styleList,&$styleAttributeValue) {
        $styleAttributeValue = strtolower($styleAttributeValue);
        foreach($styleList as $style) {
            if ($styleAttributeValue == strtolower($style)) {
                $styleAttributeValue = $style;
                return true;
            }
        }
        return false;
    }


    /**
     * pixel units to excel width units(units of 1/256th of a character width)
     * @param pxs
     * @return
     */
    protected static function _pixel2WidthUnits($pxs) {
        $UNIT_OFFSET_MAP = array(0, 36, 73, 109, 146, 182, 219);

        $widthUnits = 256 * ($pxs / 7);
        $widthUnits += $UNIT_OFFSET_MAP[($pxs % 7)];
        return $widthUnits;
    }


    /**
     * excel width units(units of 1/256th of a character width) to pixel units
     * @param widthUnits
     * @return
     */
    protected static function _widthUnits2Pixel($widthUnits) {
        $pixels = ($widthUnits / 256) * 7;
        $offsetWidthUnits = $widthUnits % 256;
        $pixels += round($offsetWidthUnits / (256 / 7));
        return $pixels;
    }


    protected static function _hex2str($hex) {
        return chr(hexdec($hex[1]));
    }

	protected static function _convertStringEncoding($string,$charset) {
		if ($charset != 'UTF-8') {
			return ExcelString::ConvertEncoding($string,'UTF-8',$charset);
		}
		return $string;
	}


	protected function _parseRichText($is = '') {
		$value = new RichText();

		$value->createText(self::_convertStringEncoding($is,$this->_charSet));

		return $value;
	}

}

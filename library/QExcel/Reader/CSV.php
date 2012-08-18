<?php
/**
 * QExcel
 *
 * Original CSV reader by PHPExcel 1.7.7 (http://www.codeplex.com/PHPExcel)
 * Modded to work with the QExcel classes and added auto delimiter
 *
 * @package     QExcel_Reader
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 */


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
class QExcel_Reader_CSV extends QExcel_Reader_ReaderAbstract
{
	/**
	 * Input encoding
	 *
	 * @access	private
	 * @var	string
	 */
	private $_inputEncoding	= '';

	/**
	 * Delimiter
     *
     * null = auto
	 *
	 * @access	private
	 * @var string
	 */
	private $_delimiter		= null;

	/**
	 * Enclosure
	 *
	 * @access	private
	 * @var	string
	 */
	private $_enclosure		= '"';

	/**
	 * Line ending
	 *
	 * @access	private
	 * @var	string
	 */
	private $_lineEnding	= PHP_EOL;

    /**
     * Can the CSV Reader open the file?
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

		return true;
	}	//	function canRead()

	/**
	 * Set input encoding
	 *
	 * @access	public
	 * @param string $pValue Input encoding
	 */
	public function setInputEncoding($pValue = 'UTF-8')
	{
		$this->_inputEncoding = $pValue;
		return $this;
	}	//	function setInputEncoding()


	/**
	 * Get input encoding
	 *
	 * @access	public
	 * @return string
	 */
	public function getInputEncoding()
	{
		return $this->_inputEncoding;
	}	//	function getInputEncoding()


    /**
     * Load workbook
     *
     * Loads the specified CSV file
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

        if (!$this->_delimiter) {
            $this->detectDelimiter($filename);
        }

        $sheet = $this->_workbook->addSheet('Sheet 1');
        $sheet->active = true;

		$lineEnding = ini_get('auto_detect_line_endings');
		ini_set('auto_detect_line_endings', true);

		// Open file
		$fileHandle = fopen($filename, 'r');
		if ($fileHandle === false) {
			throw new Exception("Could not open file $filename for reading.");
		}

		// Skip BOM, if any
		switch ($this->_inputEncoding) {
			case 'UTF-8':
				fgets($fileHandle, 4) == "\xEF\xBB\xBF" ?
					fseek($fileHandle, 3) : fseek($fileHandle, 0);
				break;
			case 'UTF-16LE':
				fgets($fileHandle, 3) == "\xFF\xFE" ?
					fseek($fileHandle, 2) : fseek($fileHandle, 0);
				break;
			case 'UTF-16BE':
				fgets($fileHandle, 3) == "\xFE\xFF" ?
					fseek($fileHandle, 2) : fseek($fileHandle, 0);
				break;
			case 'UTF-32LE':
				fgets($fileHandle, 5) == "\xFF\xFE\x00\x00" ?
					fseek($fileHandle, 4) : fseek($fileHandle, 0);
				break;
			case 'UTF-32BE':
				fgets($fileHandle, 5) == "\x00\x00\xFE\xFF" ?
					fseek($fileHandle, 4) : fseek($fileHandle, 0);
				break;
			default:
				break;
		}

		$escapeEnclosures = array( "\\" . $this->_enclosure,
								   $this->_enclosure . $this->_enclosure
								 );

		// Loop through each line of the file in turn
        $row = 0;
		while (($rowData = fgetcsv($fileHandle, 0, $this->_delimiter, $this->_enclosure)) !== FALSE) {
            $column = 0;
			foreach($rowData as $rowDatum) {
				if ($rowDatum != '') {
					// Unescape enclosures
					$rowDatum = str_replace($escapeEnclosures, $this->_enclosure, $rowDatum);

					// Convert encoding if necessary
					if ($this->_inputEncoding !== 'UTF-8') {
						$rowDatum = PHPExcel_Shared_String::ConvertEncoding($rowDatum, 'UTF-8', $this->_inputEncoding);
					}

                    $sheet->setCell($row, $column, $rowDatum);
				}
                $column++;
			}
            $row++;
		}

		// Close file
		fclose($fileHandle);
		ini_set('auto_detect_line_endings', $lineEnding);

		return $this->_workbook;
    }

    public function detectDelimiter($filename)
    {
        $possibleDelimiters = array(',', ';', "\t");
        $fileContent = substr(file_get_contents($filename), 0, 5000);
        $mostWithEnclosureDelimiter = null;
        $mostWithoutEnclosureDelimiter = null;
        $mostWithEnclosure = -1;
        $mostWithoutEnclosure = -1;
        foreach ($possibleDelimiters as $delimiter) {
            $withEnclosure = substr_count($fileContent, $delimiter.$this->_enclosure);
            $withoutEnclosure = substr_count($fileContent, $delimiter);
            if ($withEnclosure > $mostWithEnclosure) {
                $mostWithEnclosure = $withEnclosure;
                $mostWithEnclosureDelimiter = $delimiter;
            }
            if ($withoutEnclosure > $mostWithoutEnclosure) {
                $mostWithoutEnclosure = $withoutEnclosure;
                $mostWithoutEnclosureDelimiter = $delimiter;
            }
        }
        if ($mostWithEnclosure) {
            $this->setDelimiter($mostWithEnclosureDelimiter);
        } else {
            $this->setDelimiter($mostWithoutEnclosureDelimiter);
        }
    }


	/**
	 * Get delimiter
	 *
	 * @access	public
	 * @return string
	 */
	public function getDelimiter() {
		return $this->_delimiter;
	}	//	function getDelimiter()


	/**
	 * Set delimiter
	 *
	 * @access	public
	 * @param	string	$pValue		Delimiter, defaults to ,
	 * @return	PHPExcel_Reader_CSV
	 */
	public function setDelimiter($pValue = ',') {
		$this->_delimiter = $pValue;
		return $this;
	}	//	function setDelimiter()


	/**
	 * Get enclosure
	 *
	 * @access	public
	 * @return string
	 */
	public function getEnclosure() {
		return $this->_enclosure;
	}	//	function getEnclosure()


	/**
	 * Set enclosure
	 *
	 * @access	public
	 * @param	string	$pValue		Enclosure, defaults to "
	 * @return PHPExcel_Reader_CSV
	 */
	public function setEnclosure($pValue = '"') {
		if ($pValue == '') {
			$pValue = '"';
		}
		$this->_enclosure = $pValue;
		return $this;
	}	//	function setEnclosure()


	/**
	 * Get line ending
	 *
	 * @access	public
	 * @return string
	 */
	public function getLineEnding() {
		return $this->_lineEnding;
	}	//	function getLineEnding()


	/**
	 * Set line ending
	 *
	 * @access	public
	 * @param	string	$pValue		Line ending, defaults to OS line ending (PHP_EOL)
	 * @return PHPExcel_Reader_CSV
	 */
	public function setLineEnding($pValue = PHP_EOL) {
		$this->_lineEnding = $pValue;
		return $this;
	}	//	function setLineEnding()
}

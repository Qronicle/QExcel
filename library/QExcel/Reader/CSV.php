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
 * CSV Reader
 *
 * <b>Options</b>
 * <ul>
 *  <li><i>encoding</i><br/>The input file's encoding (eg. 'UTF-8', 'UTF-16LE', 'UTF-16BE', 'UTF-32LE'). Will use iconv/mbstring for conversion<br/>Default: 'UTF-8'</li>
 *  <li><i>delimiter</i><br/>The CSV delimiter (eg. ';', ',', '\t').<br/>Default: null (auto-detect)</li>
 *  <li><i>possibleDelimiters</i><br/>The delimiters that will be tried when auto-detetecting the CSV delimiter<br/>Default: [';', ',', '\t']</li>
 *  <li><i>enclosure</i>The CSV enclosure character (eg '"').<br/>Default: '"'</li>
 * </ul>
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
     * Initialize default options
     */
    public function _init()
    {
        $this->_defaultOptions = array(
            'encoding'           => 'UTF-8',
            'delimiter'          => null,
            'possibleDelimiters' => array(',', ';', "\t"),
            'enclosure'          => '"',
//          'lineEnding'         => PHP_EOL,
        );
    }

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
	}

    /**
     * Get the sheet names of a workbook
     *
     * @abstract
     * @param string $filename  The file that should be loaded
     * @return array
     * @throws Exception
     */
    public function getSheetNames($filename)
    {
        // Check if file exists
        if (!file_exists($filename)) {
            throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
        }

        return array('Sheet 1');
    }


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

        if (!$this->getOption('delimiter')) {
            $this->detectDelimiter($filename);
        }

        $sheet = $this->_workbook->addSheet('Sheet 1');

        $lineEnding = ini_get('auto_detect_line_endings');
        ini_set('auto_detect_line_endings', true);

        // Open file
        $fileHandle = fopen($filename, 'r');
        if ($fileHandle === false) {
            throw new Exception("Could not open file $filename for reading.");
        }

        // Skip BOM, if any
        switch ($this->getOption('encoding')) {
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

        $escapeEnclosures = array( "\\" . $this->getOption('enclosure'),
            $this->getOption('enclosure') . $this->getOption('enclosure')
        );

        // Loop through each line of the file in turn
        $row = 0;
        while (($rowData = fgetcsv($fileHandle, 0, $this->getOption('delimiter'), $this->getOption('enclosure'))) !== FALSE) {
            $column = 0;
            foreach($rowData as $rowDatum) {
                if ($rowDatum != '') {
                    // Unescape enclosures
                    $rowDatum = str_replace($escapeEnclosures, $this->getOption('enclosure'), $rowDatum);

                    // Convert encoding if necessary
                    if ($this->getOption('encoding') !== 'UTF-8') {
                        $rowDatum = PHPExcel_Shared_String::ConvertEncoding($rowDatum, 'UTF-8', $this->getOption('encoding'));
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

    /**
     * Detect the file's delimiter
     *
     * The delimiter option will be set after invoking this method
     * Could use some testing :)
     *
     * @param string $filename
     */
    public function detectDelimiter($filename)
    {
        $possibleDelimiters = $this->getOption('possibleDelimiters');
        $fileContent = substr(file_get_contents($filename), 0, 5000);
        $mostWithEnclosureDelimiter = null;
        $mostWithoutEnclosureDelimiter = null;
        $mostWithEnclosure = -1;
        $mostWithoutEnclosure = -1;
        foreach ($possibleDelimiters as $delimiter) {
            $withEnclosure = substr_count($fileContent, $delimiter . $this->getOption('enclosure'));
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
            $this->setOption('delimiter', $mostWithEnclosureDelimiter);
        } elseif ($mostWithoutEnclosure) {
            $this->setOption('delimiter', $mostWithoutEnclosureDelimiter);
        } else {
            $this->setOption('delimiter', ';');
        }
    }
}

<?php
/**
 * QExcel
 *
 * QExcel is heavily based on PHPExcel (http://www.codeplex.com/PHPExcel)
 *
 * @package     QExcel_Reader
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 */

/**
 * Abstract Reader
 *
 * All readers should extend this class and implement the canRead and load methods
 *
 * @abstract
 * @package     QExcel_Reader
 * @copyright   2012 Qronicle
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-17 19:32
 * @author      ruud.seberechts
 */
abstract class QExcel_Reader_ReaderAbstract
{
    /**
     * @var QExcel_Workbook
     */
    protected $_workbook;

    /**
     * Create a new reader
     */
    public function __construct()
    {
        $this->_workbook = new QExcel_Workbook();
    }

    /**
     * Can the current reader open the file?
     *
     * @abstract
     * @param string $filename  The file that should be tested
     * @return bool             Readable
     */
    abstract function canRead($filename);

    /**
     * Load a file into the Workbook format
     *
     * @abstract
     * @param string $filename  The file that should be loaded
     * @return QExcel_Workbook
     */
    abstract function load($filename);

    /**
     * Get cell value
     *
     * Get the value for a specific data type
     *
     * @todo Check the actual usefulness of this method
     *
     * @param $value
     * @param string $cellDataType
     * @return bool|float|mixed|string
     * @throws Exception
     */
    public function getCellValue($value, $cellDataType = '')
    {
        if ($cellDataType == '') {
            return $value;
        }
        switch ($cellDataType)
        {
            case PHPExcel_Cell_DataType::TYPE_STRING2:
            case PHPExcel_Cell_DataType::TYPE_STRING:
            case PHPExcel_Cell_DataType::TYPE_NULL:
            case PHPExcel_Cell_DataType::TYPE_INLINE:
                return PHPExcel_Cell_DataType::checkString($value);

            case PHPExcel_Cell_DataType::TYPE_NUMERIC:
                return (float) $value;

            case PHPExcel_Cell_DataType::TYPE_FORMULA:
                return '=FORMULA(' . (string) $value . ')';

            case PHPExcel_Cell_DataType::TYPE_BOOL:
                return (bool) $value;

            case PHPExcel_Cell_DataType::TYPE_ERROR:
                return PHPExcel_Cell_DataType::checkErrorCode($value);

            default:
                throw new Exception('Invalid datatype: ' . $cellDataType);
                break;
        }
    }

    /**
     * Get the row index from a cell name
     *
     * For example the row index of 'Z3' equals 2
     *
     * @param string $cellName  The cell name
     * @return int              The row index
     * @throws Exception        Invalid cell name format
     */
    public function getRowFromCellName($cellName)
    {
        $row = null;
        for ($i = 0; $i < strlen($cellName); $i++) {
            $row = substr($cellName, $i);
            if (!is_numeric($row)){
                continue;
            }
            break;
        }
        if (!$row) {
            throw new Exception("Could not determine row for cell $cellName");
        }

        return intval($row) - 1;
    }

    /**
     * Get the column index from a cell name
     *
     * For example the column index of 'Z3) equals 25
     *
     * @param string $cellName  The cell name
     * @return int              The column index
     * @throws Exception        Invalid cell name format
     */
    public function getColFromCellName($cellName)
    {
        $col = 0;
        $multiplier = 1;
        for ($i = strlen($cellName)-1; $i >= 0; $i--) {
            $char = substr($cellName,$i,1);
            if (is_numeric($char)) {
                continue;
            }
            $col += ord($char)-64 * $multiplier;
            $multiplier += 26;
        }
        if (!$col) {
            throw new Exception("Could not determine col for cell $cellName");
        }

        return $col - 1;
    }


    /**
     * Restrict which sheets should be loaded?
     * This property holds an array of worksheet names to be loaded. If null, then all worksheets will be loaded.
     *
     * @var array of string
     */
    protected $_loadSheetsOnly = null;

    public function getSheets()
    {
        return $this->_sheets;
    }

    /**
     * Get which sheets to load
     * Returns either an array of worksheet names (the list of worksheets that should be loaded), or a null
     *		indicating that all worksheets in the workbook should be loaded.
     *
     * @return mixed
     */
    public function getLoadSheetsOnly()
    {
        return $this->_loadSheetsOnly;
    }


    /**
     * Set which sheets to load
     *
     * @param mixed $value
     *		This should be either an array of worksheet names to be loaded, or a string containing a single worksheet name.
     *		If NULL, then it tells the Reader to read all worksheets in the workbook
     *
     * @return QExcel_Reader_Excel2007
     */
    public function setLoadSheetsOnly($value = null)
    {
        $this->_loadSheetsOnly = is_array($value) ?
            $value : array($value);
        return $this;
    }


    /**
     * Set all sheets to load
     *		Tells the Reader to load all worksheets from the workbook.
     *
     * @return QExcel_Reader_Excel2007
     */
    public function setLoadAllSheets()
    {
        $this->_loadSheetsOnly = null;
        return $this;
    }
}

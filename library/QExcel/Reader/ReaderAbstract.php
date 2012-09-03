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
     * Custom options
     *
     * @var array
     */
    protected $_options = array();

    /**
     * Default option values
     *
     * This array should be initialized in the init method and contain all possible options as a key
     *
     * @var array
     */
    protected $_defaultOptions = array();

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
        $this->_init();
    }

    /**
     * Internal initialization method
     *
     * Use this instead of overriding the constructor,
     * for example to set the default options
     */
    protected function _init()
    {
        // Override me
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
     * Get the sheet names of a workbook
     *
     * @abstract
     * @param string $filename  The file that should be loaded
     * @return array
     */
    abstract function getSheetNames($filename);

    /**
     * Get all option names that can be set
     *
     * @return array
     */
    public function getOptions()
    {
        return array_keys($this->_defaultOptions);
    }

    /**
     * Get the default options
     *
     * @return array
     */
    public function getDefaultOptions()
    {
        return $this->_defaultOptions;
    }

    /**
     * Get an option value
     *
     * @param string $key     The option name
     * @param null   $default The default value that should be returned in case this is not an option
     * @return mixed          The option value
     */
    public function getOption($key, $default = null)
    {
        return array_key_exists($key, $this->_options) ?
            $this->_options[$key] :
            (isset($this->_defaultOptions[$key]) ? $this->_defaultOptions[$key] : $default);
    }

    /**
     * Set an option value
     *
     * @param string $key   The option name
     * @param mixed  $value The new option value
     * @return bool         Whether the option value was set (FALSE in case option doesn't exist)
     */
    public function setOption($key, $value)
    {
        if (array_key_exists($key, $this->_defaultOptions)) {
            $this->_options[$key] = $value;
            return true;
        }
        return false;
    }

    /**
     * Set multiple option values
     *
     * Uses setOption for each array entry
     *
     * @param array $options
     */
    public function setOptions(array $options)
    {
        foreach ($options as $key => $value) {
            $this->setOption($key, $value);
        }
    }

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
            throw new Exception("Could not determine column for cell $cellName");
        }

        return $col - 1;
    }

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
        $sheets = $this->getOption('loadSheet');
        if (is_null($sheets)) return null;
        return is_array($sheets) ? $sheets : array($sheets);
    }
}

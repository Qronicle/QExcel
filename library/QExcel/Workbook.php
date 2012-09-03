<?php
/**
 * QExcel
 *
 * QExcel is heavily based on PHPExcel (http://www.codeplex.com/PHPExcel)
 *
 * @package     QExcel
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 */

/**
 * Workbook
 *
 * @package     QExcel
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-18 20:14
 * @author      ruud.seberechts
 */
class QExcel_Workbook
{
    /**
     * Worksheets
     *
     * @var array
     */
    protected $_sheets = array();

    /**
     * Get all worksheets
     *
     * @return array
     */
    public function getSheets()
    {
        return $this->_sheets;
    }

    /**
     * Get worksheet
     *
     * @param string|int $sheet     The sheet name or index
     * @return QExcel_Worksheet     The demanded worksheet. NULL in case there was none.
     */
    public function getSheet($sheetName)
    {
        if (is_numeric($sheetName)) {
            return isset($this->_sheets[$sheetName]) ? $this->_sheets[$sheetName] : null;
        }
        foreach ($this->_sheets as $sheet) {
            if ($sheet->name == $sheetName) {
                return $sheet;
            }
        }
        return null;
    }

    /**
     * Add worksheet
     *
     * This method is used internally by the readers
     *
     * @param string $name      The worksheet name - optional
     * @return QExcel_Worksheet
     */
    public function addSheet($name = '')
    {
        $nSheet = new QExcel_Worksheet();
        $this->_sheets[] = $nSheet;
        $nSheet->name = $name;
        return $nSheet;
    }
}

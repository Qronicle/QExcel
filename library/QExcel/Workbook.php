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
     * Add worksheet
     *
     * @param string $name      The worksheet name - optional
     * @return QExcel_Worksheet
     */
    public function addSheet($name = '')
    {
        $nSheet = new QExcel_Worksheet();
        $this->_sheets[] = $nSheet;
        $nSheet->sheetId = count($this->_sheets);
        $nSheet->name = $name;
        return $nSheet;
    }
}

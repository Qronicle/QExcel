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
 * Workbook.php
 *
 * @package     QExcel
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-18 21:00
 * @author      ruud.seberechts
 */
class QExcel_Worksheet
{
    public $sheetId = 0;
    public $name = '';
    public $data = null;
    public $active = false;

    public $numCols = 0;
    public $numRows = 0;

    public function setCell($row, $col, $value)
    {
        $this->numCols = max($col+1, $this->numCols);
        $this->numRows = max($row+1, $this->numRows);
        $this->data[$row][$col] = $value;
    }
}

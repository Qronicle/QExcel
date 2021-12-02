<?php

namespace QExcel;

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
class Worksheet
{
    /**
     * Worksheet name
     *
     * @var string
     */
    public $name = '';

    /**
     * Worksheet data grid
     *
     * @var array
     */
    protected $_data = array();

    /**
     * Column counter
     *
     * @var int
     */
    protected $_numCols = 0;

    /**
     * Row counter
     *
     * @var int
     */
    protected $_numRows = 0;

    /**
     * Get the worksheet name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set worksheet cell value
     *
     * Is used internally when populating the worksheet
     *
     * @param int $row
     * @param int $col
     * @param mixed $value
     */
    public function setCell($row, $col, $value)
    {
        $this->_numCols = max($col+1, $this->_numCols);
        $this->_numRows = max($row+1, $this->_numRows);
        $this->_data[$row][$col] = $value;
    }

    /**
     * Get the amount of columns
     *
     * @return int
     */
    public function getNumColumns()
    {
        return $this->_numCols;
    }

    /**
     * Get the amount of rows
     *
     * @return int
     */
    public function getNumRows()
    {
        return $this->_numRows;
    }

    /**
     * Get the value of a single cell
     *
     * @param int $row
     * @param int $column
     * @return mixed        The cell value, NULL in case the cell was not populated
     */
    public function getCellValue($row, $column)
    {
        return isset($this->_data[$row][$column]) ? $this->_data[$row][$column] : null;
    }

    /**
     * Get all worksheet cell data
     *
     * @return array    array[$row][$column]
     */
    public function getData()
    {
        return $this->_data;
    }

    /**
     * Get all worksheet cell data in an associative array
     *
     * @param array|null $columns
     * @return array    array[$row][$columnTitle => $columnValue]
     */
    public function getAssocData($columns = null)
    {
        $assocData = $this->_data;

        // Create columns from first row
        if (is_null($columns)) {
            if (!$assocData) {
                return [];
            }
            $columns = array_map(function($value) {
                return str_replace(' ', '_', mb_strtolower($value));
            }, array_shift($assocData));
        }

        // Use columns as keys for all row data
        foreach ($assocData as $row => $values) {
            // Fix issue where empty cells are not available in the data array
            if (count($values) != count($columns)) {
                $newValues = [];
                for ($i = 0; $i < count($columns); $i++) {
                    $newValues[$i] = $values[$i] ?? '';
                }
                $values = $newValues;
            }
            $assocData[$row] = array_combine($columns, $values);
        }
        return $assocData;
    }
}

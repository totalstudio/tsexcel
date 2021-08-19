<?php
namespace TotalStudio\TSExcel\View\Helper;

use Cake\Collection\Collection;
use Cake\Database\Expression\QueryExpression;
use Cake\I18n\Date;
use Cake\I18n\FrozenDate;
use Cake\I18n\FrozenTime;
use Cake\I18n\Time;
use Cake\ORM\Entity;
use Cake\ORM\Query;
use Cake\ORM\ResultSet;
use Cake\View\Helper;
use Cake\View\View;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Exception;
use TSExcel\View\ExcelView;

class ExcelHelper extends Helper
{
    /**
     * Format in which dates are exported to excel
     * set it globally in the bootstrap file or pass it as config-Variable
     *
     * @var ExcelView $view
     */
    private $view = null;

    /**
     * Constructor
     *
     * @param View $View
     * @param array $config
     */
    public function __construct(View $View, array $config = array())
    {
        parent::__construct($View, $config);
        $this->view = $this->getView();

    }

    /**
     * add new Worksheet
     *
     * @param mixed $data can be Query, Entity, Collection or flat Array
     * @param string $name
     * @return void
     * @throws Exception
     */
    public function addWorksheet($data = null, $name = '')
    {

        // add empty sheet to Workbook
        $this->addSheet($name);

        if (is_array($data)) {
            $data = $this->prepareCollectionData(collection($data));
        } elseif ($data instanceof Entity) {
            $data = $this->prepareEntityData($data);
        } elseif ($data instanceof Query) {
            $data = $this->prepareCollectionData(collection($data->toArray()));
        } elseif ($data instanceof ResultSet) {
            $data = $this->prepareCollectionData(collection($data->toArray()));
        } else {
            $data = $this->prepareCollectionData($data);
        }

        // add the Data
        $this->addData($data);

        //auto-sizing of the columns
        $highestColumn = $this->view->getSpreadsheet()->getActiveSheet()->getHighestColumn();
        foreach (range('A', $highestColumn) as $column) {
            $this->view->getSpreadsheet()->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
        }

    }

    /**
     * converts a Collection into a flat Array
     * properties are extracted from first item und inserted in first row
     *
     * @param mixed $collection \Cake\Collection\Collection | \Cake\ORM\Query
     * @return array
     */
    public function prepareCollectionData(Collection $collection = null)
    {

        /* extract keys from first item */
        $first = $collection->first();
        if (is_array($first)) {
            $data = [array_keys($first)];
        } else {
            $data = [array_keys($first->toArray())];
        }

        /* add data */
        foreach ($collection as $row) {

            if (is_array($row)) {
                $data[] = array_values($row);
            } else {
                $data[] = array_values($row->toArray());
            }
        }
        return $data;
    }

    /**
     * converts a Entity into a flat Array
     * properties are inserted in first row
     *
     * @param Entity $entity
     * @return array
     */
    public function prepareEntityData(Entity $entity = null)
    {
        $entityArray = $entity->toArray();
        $data = [array_keys($entityArray)];
        $data[] = array_values($entityArray);

        return $data;
    }

    /**
     * adds data to a worksheet
     *
     * @param array $array
     * @param array $options if set row and column, data entry starts there
     * @return void
     * @throws Exception
     */
    public function addData(array $array = [], array $options = [])
    {
        $rowIndex = isset($options['row']) ? $options['row'] : 1;
        foreach ($array as $row) {
            $columnIndex = isset($options['column']) ? $options['column'] : 0;
            foreach ($row as $cell) {
                $this->_addCellData($cell, $columnIndex, $rowIndex);
                $columnIndex++;
            }
            $rowIndex++;
        }
    }

    /**
     * Fills in the data in a cell.
     * respects data type
     *
     * @param mixed $cell
     * @param int $columnIndex
     * @param int $rowIndex
     * @return void
     * @throws Exception
     */
    protected function _addCellData($cell = null, $columnIndex = 1, $rowIndex = 1)
    {
        if (is_array($cell)) {
            $cell = null; // adding cells of this Type is useless
            return;
        }
        if ($cell instanceof Date or $cell instanceof Time or $cell instanceof FrozenDate or $cell instanceof FrozenTime) {
            $cell = $cell->i18nFormat($this->__dateformat);  // Dates must be converted for Excel
            $this->view->getSpreadsheet()->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex)->setValueExplicit($cell, DataType::TYPE_STRING);
            return;
        }
        if ($cell instanceof QueryExpression) {
            $cell = null;  // @TODO find a way to get the Values and insert them into the Sheet
            return;
        }
        if (is_string($cell)) {
            $this->view->getSpreadsheet()->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex)->setValueExplicit($cell, DataType::TYPE_STRING);
            return;
        }
        $this->view->getSpreadsheet()->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex)->setValueExplicit($cell, DataType::TYPE_NUMERIC);

    }

    /**
     * create empty Sheet and add some Metadata
     *
     * @param string $title
     * @return void
     * @throws Exception
     */
    public function addSheet($title = '')
    {
        $this->view->getSpreadsheet()->createSheet();
        $this->view->getSpreadsheet()->setActiveSheetIndex(1);
        $this->view->getSpreadsheet()->getActiveSheet()->setTitle($title);
        $this->view->getSpreadsheet()->getProperties()->setTitle($title);
        $this->view->getSpreadsheet()->getProperties()->setSubject($title. ' ' . date('d.m.Y H:i'));
    }

    /**
     * Set the Name of Excel-File
     *
     * @param string $filename
     */
    public function setFilename(string $filename)
    {
        $this->view->setFilename($filename);
    }
}
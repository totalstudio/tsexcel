<?php

namespace TotalStudio\TSExcel\Controller\Component;

use Cake\Controller\Component;
use Cake\Controller\ComponentRegistry;
use Cake\ORM\Exception\MissingTableClassException;


/**
 * CakePHP Import Component
 *
 * @author TotalStudio <info@totalstudio.hu>
 */
class ImportComponent extends Component
{

    /**
     * reads a file PHPExcel can understand and converts a contained worksheet into an array
     * which can be used to build entities. If the File contains more than one worksheet and it is not named like the Controller
     * you have to provide the name or index of the desired worksheet in the options array.
     * If you set $options['append'] to true, the primary key will be deleted.
     * @todo Find a way to handle primary keys other than id.
     *
     * @param string $file name of Excel-File with full path. Must be of a readable Filetype (xls, xlsx, csv, ods)
     * @param array $options Override Worksheet name, set append Mode
     * @return array The Array has the same structure as provided by request->data
     * @throws MissingTableClassException
     */
    public function prepareEntityData($file = null, array $options = [])
    {

        /**  load and configure PHPExcelReader  * */
        $fileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file);

        $reader = $PhpExcelReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($fileType);
        $reader->setReadDataOnly(true);

        if ($fileType !== 'CSV') {  // csv-files have only one 'worksheet'
            $reader->setLoadAllSheets();


            /** identify worksheets in file * */
            $worksheets = $reader->listWorksheetNames($file);

            $worksheetToLoad = null;

            if (count($worksheets) === 1) {
                $worksheetToLoad = $worksheets[0];  //first option: if there is only one worksheet, use it
            } elseif (isset($options['worksheet']) && is_int($options['worksheet']) && isset($worksheets[$options['worksheet']])) {
                $worksheetToLoad = $worksheets[$options['worksheet']]; //second option: select a fixed worksheet index
            } elseif (isset($options['worksheet'])) {
                $worksheetToLoad = $options['worksheet']; //third option: desired worksheet was provided as option
            } else {
                $worksheetToLoad = $this->_registry->getController()->name; //last option: try to load worksheet with the name of current controller
            }
            if (!in_array($worksheetToLoad, $worksheets)) {
                throw new MissingTableClassException(__('No proper named worksheet found'));
            }

            /** load the sheet and convert data to an array */
            $reader->setLoadSheetsOnly($worksheetToLoad);
        }

        $PhpExcel = $reader->load($file);
        $data = $PhpExcel->getSheet(0)->toArray();

        /** convert data for building entities */
        $result = [];
        $properties = array_shift($data); //first row columns are the properties
        foreach($properties as $propertyCol => $property){
            if(empty($property)){
                $properties[$propertyCol] = $propertyCol;
            }
        }

        foreach ($data as $row) {
            $record = array_combine($properties, $row);
            if (isset($record['modified'])) {
                unset($record['modified']);
            }
            if (isset($options['type']) && $options['type'] == 'append' && isset($record['id'])) {
                unset($record['id']);
            }
            $result[] = $record;
        }

        /** log in debug mode */
        $this->log(count($result) . ' records were extracted from File ' . $file, 'debug');

        return $result;
    }

}

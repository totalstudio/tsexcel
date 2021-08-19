<?php

use Cake\Core\Configure;
use Cake\Datasource\ConnectionManager;
use Cake\Event\EventManager;

/**
 * use sqlite for temporary Data
 * borrowed idea from debug_kit
 */
if(!ConnectionManager::getConfig('excel')){
    ConnectionManager::setConfig('excel', [
        'className' => 'Cake\Database\Connection',
        'driver' => 'Cake\Database\Driver\Sqlite',
        'database' => TMP . 'excel.sqlite',
        'encoding' => 'utf8',
        'cacheMetadata' => true,
        'quoteIdentifiers' => false,
    ]);
}

/**
 * load and prepare RequestHandler in all Controllers
 */
EventManager::instance()
        ->on('Controller.initialize', function (Cake\Event\Event $event) {
            $controller = $event->getSubject();
            if ($controller->components()->has('RequestHandler')) {
                $controller->RequestHandler->setConfig('viewClassMap.xlsx', 'TotalStudio/TSExcel.Excel');
            }
        }
);

/**
 * because Cake3 returns instances of Cake\i18n\Time for datetime-fields, it is necessary to tell the Helper
 * in which format they should be exported to excel
 */
Configure::write('excel.creator', 'TotalStudio');
Configure::write('excel.generator', 'TotalStudio CakePHP 4 Excel Generator');

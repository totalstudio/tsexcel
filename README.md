# TSExcel plugin for CakePHP

##About
This plugin inspired by CewiExcel Cake PHP plugin (https://github.com/cewi/excel) and uses PHPSpreadsheet (https://github.com/PHPOffice/PhpSpreadsheet)

## Installation

You can install this plugin into your CakePHP application using [composer](https://getcomposer.org).

The recommended way to install composer packages is:

add

```
"repositories": [
         {
            "type": "vcs",
            "url": "https://github.com/totalstudio/tsexcel"
        }
    ] 
```

to your composer.json because this package is not on packagist. Then in your console:

```
composer require total-studio/t-s-excel
```

should fetch the plugin. Load the Plugin in the bootstrap() method in Application.php as ususal:

```
$this->addPlugin('TotalStudio/TSExcel');
```

RequestHandler Component is configured by the Plugin's bootstrap file. But you could do this also in a controller's initialize method, e.g.:
```
public function initialize()
	{
        	parent::initialize();
        	$this->loadComponent('RequestHandler', [
            		'viewClassMap' => ['xlsx' => 'TotalStudio/TSExcel.Excel']
        	]);
        }
```

Be careful: RequestHandlerComponent is already loaded in your AppController by default. Adapt the settings to your needs.

You need to set up parsing for the xlsx extension. Add the following to your config/routes.php file before any route or scope definition:
```
Router::extensions('xlsx');
```

you can configure this also within a scope:
```
$routes->setExtensions(['xlsx']);
```

(Setting this in the plugin's config/routes.php file is currently broken. So you do have to provide the code in the application's config/routes.php file)

You further have to provide a layout for the generated Excel-Files. Add a folder xlsx in src/Template/Layout/ subdirectory and within that folder a file named default.ctp with this minimum content:
```  
<?= $this->fetch('content') ?>
```  

## Usage of ExcelHelper
Has a Method 'addworksheet' which takes a ResultSet, an Entity, a Collection of Entities or an Array of Data and creates a worksheet from the data. Properties of the Entities, or the keys of the first record in the array are set as column-headers in first row of the generated worksheet. Be careful if you use non-standard column-types. The Helper actually works only with strings, numbers and dates. 

Register xlsx-Extension in config/routes.php file before the routes that should be affected:
```
Router::extensions(['xlsx']);
```

Example (assumed you have an article model and controller with the usual index-action) 

Include the helper in ArticlesController:
```
public $helpers = ['TotalStudio/TSExcel.Excel'];
```

add a Folder 'xlsx' in Template/Articles and create the file 'index.ctp' in this Folder. Include this snippet of code to get an excel-file with a single worksheet called 'Articles':        
```    
$this->Excel->addWorksheet($articles, 'Articles');
```    
    
create the link to generate the file somewhere in your app: 
```
<?= $this->Html->link(__('Excel'), ['controller' => 'Articles', 'action' => 'index', '_ext'=>'xlsx']); ?>
```

done. The name of the file will be the lowercase plural of the entity with '.xslx' added, e.g. 'articles.xlsx'. If you like to change that, add
```
$this->Excel->setFilename('foo');
```
in the Template file. The filename now will be 'foo.xlslx'. 


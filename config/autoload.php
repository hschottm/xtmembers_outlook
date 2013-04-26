<?php

/**
 * Contao Open Source CMS
 * 
 * Copyright (C) 2005-2013 Leo Feyer
 * 
 * @package Xtmembers_outlook
 * @link    https://contao.org
 * @license http://www.gnu.org/licenses/lgpl-3.0.html LGPL
 */


/**
 * Register the classes
 */
ClassLoader::addClasses(array
(
	// Classes
	'Contao\OutlookTools' => 'system/modules/xtmembers_outlook/classes/OutlookTools.php',
));


/**
 * Register the templates
 */
TemplateLoader::addFiles(array
(
	'be_import_outlook' => 'system/modules/xtmembers_outlook/templates',
));

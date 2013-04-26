<?php if (!defined('TL_ROOT')) die('You can not access this file directly!');

/**
 * TYPOlight webCMS
 * Copyright (C) 2005 Leo Feyer
 *
 * This program is free software: you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation, either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program. If not, please visit the Free
 * Software Foundation website at http://www.gnu.org/licenses/.
 *
 * PHP version 5
 * @copyright  Helmut Schottmüller 2009 
 * @author     Helmut Schottmüller 
 * @package    xtmembers_outlook 
 * @license    LGPL 
 * @filesource
 */


$GLOBALS['TL_LANG']['tl_member']['outlook_field'] = array('Outlook field', 'Name of the export field for Microsoft Outlook contact data.');
$GLOBALS['TL_LANG']['tl_member']['member_field'] = array('Member field', 'Name of the TYPOlight member field.');
$GLOBALS['TL_LANG']['tl_member']['start_export'] = 'Export';
$GLOBALS['TL_LANG']['tl_member']['start_import'] = 'Import';
$GLOBALS['TL_LANG']['tl_member']['info_export'] = 'Please assign the TYPOlight member fields to the corresponding Outlook fields that are used for the Outlook Excel import of Outlook contacts. Your selection will be saved and is available for future exports.';
$GLOBALS['TL_LANG']['tl_member']['outlooksource'] = array('File source', 'Please choose the Outlook export file (.csv) you want to import from the files directory.');
$GLOBALS['TL_LANG']['tl_member']['outlook_import_groups'] = array('Member groups', 'Please select the member groups that should be associated with the imported members.');
$GLOBALS['TL_LANG']['tl_member']['outlook_import_newsletter'] = array('Newsletter', 'Please select the newsletters that should be associated with the imported members.');
?>
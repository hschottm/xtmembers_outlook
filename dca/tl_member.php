<?php

/**
 * @copyright  Helmut Schottmüller
 * @author     Helmut Schottmüller <https://github.com/hschottm/xtmembers_outlook>
 * @license    LGPL
 */

/**
 * Table tl_member
 */
$GLOBALS['TL_DCA']['tl_member']['list']['global_operations']['outlook_export'] = 
	array(
		'label'               => &$GLOBALS['TL_LANG']['MSC']['outlook_export'],
		'href'                => 'key=outlook_export',
		'class'               => 'header_export',
		'attributes'          => 'onclick="Backend.getScrollOffset();"'
	);

$GLOBALS['TL_DCA']['tl_member']['list']['global_operations']['outlook_import'] = 
	array(
		'label'               => &$GLOBALS['TL_LANG']['MSC']['outlook_import'],
		'href'                => 'key=outlook_import',
		'class'               => 'header_import',
		'attributes'          => 'onclick="Backend.getScrollOffset();"'
	);

$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource'] = array
(
	'label'                   => &$GLOBALS['TL_LANG']['tl_content']['source'],
	'eval'                    => array('fieldType'=>'radio', 'files'=>true, 'filesOnly'=>true, 'extensions'=>'csv')
);

/**
 * Class tl_member_export
 *
 * Provide methods that are used for import and export of member data
 * @copyright  Helmut Schottmüller
 * @author     Helmut Schottmüller <https://github.com/hschottm/xtmembers_outlook>
 * @package    Controller
 */
class tl_member_outlook extends tl_member
{
}

?>
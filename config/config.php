<?php

/**
 * @copyright  Helmut Schottmüller
 * @author     Helmut Schottmüller <https://github.com/hschottm/xtmembers_outlook>
 * @license    LGPL
 */

/**
 * -------------------------------------------------------------------------
 * BACK END MODULES
 * -------------------------------------------------------------------------
 */

$GLOBALS['BE_MOD']['accounts']['member']['outlook_export'] = array('OutlookTools', 'exportMembers');
$GLOBALS['BE_MOD']['accounts']['member']['outlook_import'] = array('OutlookTools', 'importMembers');
$GLOBALS['BE_MOD']['accounts']['member']['stylesheet'] = 'system/modules/xtmembers_outlook/assets/outlook.css';

?>
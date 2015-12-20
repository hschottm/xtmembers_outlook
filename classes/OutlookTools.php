<?php

/**
 * @copyright  Helmut Schottmüller
 * @author     Helmut Schottmüller <https://github.com/hschottm/xtmembers_outlook>
 * @license    LGPL
 */

namespace Contao;

/**
 * Class OutlookTools
 *
 * Provide methods to handle import and export of member data to/from Outlook
 * @copyright  Helmut Schottmüller
 * @author     Helmut Schottmüller <https://github.com/hschottm/xtmembers_outlook>
 * @package    Controller
 */
class OutlookTools extends Backend
{
	protected $blnSave = true;

	protected function importMembersFromXML($filename)	
	{
		$isTLFile = false;
		$isMember = false;
		$tags = array();
		$tag = "";
		$members = array();
		$member = null;
		$xml = new XMLReader(); 
		$xml->open(TL_ROOT . '/' . $filename);
		while($xml->read()) 
		switch ($xml->nodeType) 
		{ 
			case XMLReader::END_ELEMENT:
				switch ($xml->name)
				{
					case 'TYPOlight':
					case 'member':
						$isMember = false;
						array_pop($tags);
						array_push($members, $member);
						break;
					default:
						array_pop($tags);
						break;
				}
				break;
			case XMLReader::ELEMENT: 
				switch ($xml->name)
				{
					case 'TYPOlight':
						$isTLFile = true;
						break;
					case 'member':
						$isMember = true;
						$member = array();
						break;
				}
				if ($isTLFile && $isMember)
				{
					$tag = $xml->name;
					if (!$xml->isEmptyElement)
					{
						array_push($tags, $tag);
						if (count($tags) == 2)
						{
							if($xml->hasAttributes)
							{
								while($xml->moveToNextAttribute()) 
								{
									if (strcmp($xml->name, 'unserialized') == 0) $member[$tag] = $xml->value;
								}
							}
						}
					}
				}
				break; 
			case XMLReader::TEXT: 
			case XMLReader::CDATA:
				if (count($tags) == 2)
				{
					$member[$tag] = $xml->value; 
				}
				break;
		} 
		$xml->close();

		// Get all default values for the new entry
		$defaults = array();
		foreach ($GLOBALS['TL_DCA']['tl_member']['fields'] as $k=>$v)
		{
			if (isset($v['default']))
			{
				$defaults[$k] = is_array($v['default']) ? serialize($v['default']) : $v['default'];
			}
		}
		foreach ($members as $member)
		{
			$set = array();
			foreach ($member as $field => $value)
			{
				if ($this->Database->fieldExists($field, 'tl_member'))
				{
					$set[$field] = $value;
				}
			}
			// Set passed values
			if (is_array($set) && count($set))
			{
				$set = array_merge($set, $defaults);
			}

			$set['tstamp'] = time();
			$objInsertStmt = $this->Database->prepare("INSERT INTO tl_member %s")
				->set($set)
				->execute();
		}
		$this->redirect(str_replace('&key=import', '', $this->Environment->request));
	}
	
	protected function writeArray(XMLWriter &$writer, $array)
	{
		foreach ($array as $key => $value)
		{
			if (is_array($value))
			{
				$writer->startElement($key);
				$this->writeArray($writer, $value);
				$writer->endElement();
			}
			else
			{
				$writer->startElement('item');
				$writer->writeAttribute('key', $key);
				$writer->text($value);
				$writer->endElement();
			}
		}
	}
	
	public function exportMembers()
	{
		if ($this->Input->get('key') != 'outlook_export')
		{
			$this->redirect(str_replace('&key=outlook_export', '', $this->Environment->request));
		}

		if ($this->Input->post('FORM_SUBMIT') == 'tl_export_outlook')
		{
			$export_settings = array();
			for ($i = 1; $i <= 92; $i++)
			{
				$export_settings[$i] = $this->Input->post('outlook' . $i);
			}
			$this->Config->update("\$GLOBALS['TL_CONFIG']['outlook_export']", serialize($export_settings));
			$objMember = $this->Database->prepare("SELECT * FROM tl_member ORDER BY lastname,firstname")->execute();
			if ($objMember->numRows)
			{
				$xls = new \xlsexport();
				$sheet = utf8_decode($GLOBALS['TL_LANG']['MSC']['outlook_contacts']);
				$xls->addworksheet($sheet);
				$intRowCounter = 1;
				$intColCounter = 0;
				for ($c = 1; $c <= 92; $c++)
				{
					$xls->setcell(array("sheetname" => $sheet,"row" => 0, "col" => $c-1, "data" => utf8_decode($GLOBALS['TL_LANG']['MSC']['ol_field' . $c])));
				}
				while ($objMember->next())
				{
					$data = $objMember->row();
					for ($c = 1; $c <= 92; $c++)
					{
						$field = $export_settings[$c];
						switch ($field)
						{
							case 'tags':
								$arrTags = $this->Database->prepare("SELECT tag FROM tl_tag WHERE id = ? AND from_table = ?")
									->execute($data['id'], 'tl_member')->fetchEach('tag');
								if (is_array($arrTags))
								{
									$output = join($arrTags, ',');
								}
								else
								{
									$output = '';
								}
								break;
							case 'gender':
								$output = $GLOBALS['TL_LANG']['MSC'][$data[$field]];
								break;
							case 'country':
								$output = $GLOBALS['TL_LANG']['CNT'][$data[$field]];
								break;
							case 'dateOfBirth':
								$d = new Date($data[$field]);
								$output = $d->date;
								break;
							default:
								$output = $data[$field];
								break;
						}
 						$xls->setcell(array("sheetname" => $sheet,"row" => $intRowCounter, "col" => $c-1, "data" => utf8_decode($output)));
					}
					$intRowCounter++;
				}
				$xls->sendFile($this->safefilename('export') . ".xls");
				exit;
			}
		}

		// Return form
		$result = '
<div id="tl_buttons">
<a href="'.ampersand(str_replace('&key=outlook_export', '', $this->Environment->request)).'" class="header_back" title="'.specialchars($GLOBALS['TL_LANG']['MSC']['backBT']).'">'.$GLOBALS['TL_LANG']['MSC']['backBT'].'</a>
</div>

<h2 class="sub_headline">'.$GLOBALS['TL_LANG']['MSC']['export_member'][0].'</h2>'.$this->getMessages().'

<form action="'.ampersand($this->Environment->request, ENCODE_AMPERSANDS).'" id="tl_export_outlook" class="tl_form" method="post">
<div class="tl_formbody_edit">
<input type="hidden" name="FORM_SUBMIT" value="tl_export_outlook" />
<input type="hidden" name="REQUEST_TOKEN" value="' . REQUEST_TOKEN . '" />

<div class="tl_tbox">
  <div>' . $GLOBALS['TL_LANG']['tl_member']['info_export'] . '</div>';
	
	$result .= '<table>';
	$result .= '<thead><tr><th>' . $GLOBALS['TL_LANG']['tl_member']['outlook_field'][0] . '</th>';
	$result .= '<th>' . $GLOBALS['TL_LANG']['tl_member']['member_field'][0] . '</th></tr></thead><tbody>';
	$values = deserialize($GLOBALS['TL_CONFIG']['outlook_export'], true);
	for ($i = 1; $i <= 92; $i++)
	{
		$objSelect = $this->getOutlookFieldWidget($i, $values[$i]);
		$result .= '<tr>';
		$result .= '<td>' . $objSelect->generateLabel() . '</td>';
		$result .= '<td>' . $objSelect->generate() . '</td>';
		$result .= '</tr>';
	}
	$result .= '</tbody></table>';
	$result .= '</div>

</div>

<div class="tl_formbody_submit">

<div class="tl_submit_container">
<input type="submit" name="export" id="save" class="tl_submit" alt="export member" accesskey="s" value="'.specialchars($GLOBALS['TL_LANG']['tl_member']['start_export']).'" />
</div>

</div>
</form>';
		return $result;
	}
	
	protected function safefilename($filename) 
	{
		$search = array('/ß/','/ä/','/Ä/','/ö/','/Ö/','/ü/','/Ü/','([^[:alnum:]._])');
		$replace = array('ss','ae','Ae','oe','Oe','ue','Ue','_');
		return preg_replace($search,$replace,$filename);
	}
	
	public function importMembers()
	{
		if ($this->Input->get('key') != 'outlook_import')
		{
			$this->redirect(str_replace('&key=outlook_import', '', $this->Environment->request));
		}

		$this->import('BackendUser', 'User');
		$class = $this->User->uploader;

		// See #4086
		if (!class_exists($class))
		{
			$class = 'FileUpload';
		}

		$objUploader = new $class();

		$this->Template = new BackendTemplate('be_import_outlook');

		$class = $this->User->uploader;

		// See #4086
		if (!class_exists($class))
		{
			$class = 'FileUpload';
		}

		$objUploader = new $class();
		$this->Template->markup = $objUploader->generateMarkup();
		$this->Template->outlooksource = $this->getFileTreeWidget();
		$this->Template->membergroup = $this->getMembergroupsWidget();
		$this->Template->encoding = $this->getEncodingWidget();
		$this->Template->newsletter = $this->getNewsletterWidget();
		$this->Template->hrefBack = ampersand(str_replace('&key=outlook_import', '', $this->Environment->request));
		$this->Template->goBack = $GLOBALS['TL_LANG']['MSC']['goBack'];
		$this->Template->headline = $GLOBALS['TL_LANG']['MSC']['outlook_import'][0];
		$this->Template->request = ampersand($this->Environment->request, ENCODE_AMPERSANDS);
		$this->Template->submit = specialchars($GLOBALS['TL_LANG']['MSC']['continue']);

		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook')
		{
			$arrFiles = array();
			$strFile = $this->Input->post('filepath');
			
			// Skip folders
			if (is_dir(TL_ROOT . '/' . $strFile))
			{
				\Message::addError(sprintf($GLOBALS['TL_LANG']['ERR']['importFolder'], basename($strFile)));
			}

			$objFile = new \File($strFile, true);

			if ($objFile->extension != 'txt' && $objFile->extension != 'csv')
			{
				\Message::addError(sprintf($GLOBALS['TL_LANG']['ERR']['filetype'], $objFile->extension));
				continue;
			}

			$arrFiles[] = $strFile;

			if (empty($arrFiles))
			{
				\Message::addError($GLOBALS['TL_LANG']['ERR']['emptyUpload']);
				$this->reload();
			}
			else if (count($arrFiles) > 1)
			{
				\Message::addError($GLOBALS['TL_LANG']['ERR']['only_one_file']);
				$this->reload();
			}
			else
			{
				$file = new \File($arrFiles[0], true);
				$output = $this->importMembersFromCSV($file->path);
				return $output;
			}
		}
		// Create import form
		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook_fileselection')
		{
			$arrUploaded = $objUploader->uploadTo('system/tmp');
			if (empty($arrUploaded))
			{
				\Message::addError($GLOBALS['TL_LANG']['ERR']['emptyUpload']);
				$this->reload();
			}

			$arrFiles = array();

			foreach ($arrUploaded as $strFile)
			{
				// Skip folders
				if (is_dir(TL_ROOT . '/' . $strFile))
				{
					\Message::addError(sprintf($GLOBALS['TL_LANG']['ERR']['importFolder'], basename($strFile)));
					continue;
				}

				$objFile = new \File($strFile, true);

				if ($objFile->extension != 'txt' && $objFile->extension != 'csv')
				{
					\Message::addError(sprintf($GLOBALS['TL_LANG']['ERR']['filetype'], $objFile->extension));
					continue;
				}

				$arrFiles[] = $strFile;
			}
			if (empty($arrFiles))
			{
				\Message::addError($GLOBALS['TL_LANG']['ERR']['emptyUpload']);
				$this->reload();
			}
			else if (count($arrFiles) > 1)
			{
				\Message::addError($GLOBALS['TL_LANG']['ERR']['only_one_file']);
				$this->reload();
			}
			else
			{
				$file = new \File($arrFiles[0], true);
				$filename = $file->path;
				$groups = $this->Template->membergroup->value;
				$newsletters = $this->Template->newsletter->value;
				// set session values
				$this->Session->set('outlook_fileid', $filename);
				$this->Session->set('outlook_groups', $groups);
				$this->Session->set('outlook_newsletters', $newsletters);
				$this->Session->set('outlook_encoding', $this->Input->post('encoding'));
				$output = $this->importMembersFromCSV($file->path);
				return $output;
			}
		}
		return $this->Template->parse();
	}

	public function importMembersFromCSV($filepath)
	{
				if ($this->Input->get('key') != 'outlook_import')
				{
					$this->redirect(str_replace('&key=outlook_import', '', $this->Environment->request));
				}

				$encoding = $this->Session->get('outlook_encoding');

				$parser = new CSVParser(TL_ROOT . '/' . $filepath, (strlen($encoding) > 0) ? $encoding : 'UTF-8');
				$fields = $parser->extractHeader();
				
				foreach ($fields as $idx => $field)
				{
					$fields[$idx] = trim(str_replace('"', '', $field));
				}
				if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook')
				{
					$import_settings = array();
					$i = 0;
					$idx_email = -1;
					$idx_firstname = -1;
					$idx_lastname = -1;
					while (array_key_exists('outlook' . $i, $_POST))
					{
						array_push($import_settings, $this->Input->post('outlook' . $i));
						$foundpair = trimsplit('::', $this->Input->post('outlook' . $i));
						switch ($foundpair[1])
						{
							case 'email':
								$idx_email = $i;
								break;
							case 'firstname':
								$idx_firstname = $i;
								break;
							case 'lastname':
								$idx_lastname = $i;
								break;
						}
						$i++;
					}
					$this->Config->update("\$GLOBALS['TL_CONFIG']['outlook_import']", serialize($import_settings));
					$idx = 1;
					while ($entities = $parser->getDataArray())
					{
						foreach ($entities as $ent_idx => $ent_value)
						{
							$ent_value = preg_replace("/(^\")|(\"$)/", "", $ent_value);
							$entities[$ent_idx] = $ent_value;
						}
						$where = array();
						$where_values = array();
						if ($idx_email >= 0)
						{
							if (strlen($entities[$idx_email]))
							{
								array_push($where, 'email = ?');
								array_push($where_values, $entities[$idx_email]);
							}
						}
						if ($idx_firstname >= 0)
						{
							if (strlen($entities[$idx_firstname]))
							{
								array_push($where, 'firstname = ?');
								array_push($where_values, $entities[$idx_firstname]);
							}
						}
						if ($idx_lastname >= 0)
						{
							if (strlen($entities[$idx_lastname]))
							{
								array_push($where, 'lastname = ?');
								array_push($where_values, $entities[$idx_lastname]);
							}
						}
						$where_text = join($where, ' AND ');
						if (strlen($where_text)) $where_text = ' WHERE ' . $where_text;
						
						$set = array();
						$set['tstamp'] = time();
						$set['dateAdded'] = time();
						$objInsertStmt = $this->Database->prepare("INSERT INTO tl_member %s")
							->set($set)
							->execute();
						if ($objInsertStmt->affectedRows)
						{
							$insertID = $objInsertStmt->insertId;
							$member = null;
							$member = FrontendUser::getInstance();
							$member->findBy('id', $insertID);
							$member->allGroups = deserialize($this->Session->get('outlook_groups'), true);
							$member->newsletter = deserialize($this->Session->get('outlook_newsletters'), true);
							$member->tstamp = time();
							$tags = array();
							foreach ($entities as $fieldidx => $field)
							{
								$fieldname = $import_settings[$fieldidx];
								$foundpair = trimsplit('::', $fieldname);
								$fieldname = $foundpair[1];
								if (strlen($fieldname))
								{
									switch ($fieldname)
									{
										case 'tags':
											$tags = trimsplit(",", $field);
											break;
										case 'dateOfBirth':
//											$d = new Date($data[$field]);
//											$output = $d->date;
											break;
									}
									$member->$fieldname = $field;
								}
							}
							$last_insert_id = $member->save();
							foreach (deserialize($this->Session->get('outlook_newsletters'), true) as $channel)
							{
								if (strlen($member->email))
								{
									$this->Database->prepare("INSERT INTO tl_newsletter_recipients (pid, tstamp, email, active) VALUES (?, ?, ?, ?)")
										->execute($channel, time(), $member->email, 1);
								}
							}
							if ($this->Database->tableExists('tl_tag'))
							{
								foreach ($tags as $tag)
								{
									$this->Database->prepare("INSERT INTO tl_tag (id, tag, from_table) VALUES (?, ?, ?)")
										->execute($last_insert_id, $tag, 'tl_member');
								}
							}
						}
					}
					$this->redirect(str_replace('&key=outlook_import', '', $this->Environment->request));
				}

				// Return form
				$result = '
		<div id="tl_buttons">
		<a href="'.ampersand(str_replace('&key=import', '', $this->Environment->request)).'" class="header_back" title="'.specialchars($GLOBALS['TL_LANG']['MSC']['backBT']).'">'.$GLOBALS['TL_LANG']['MSC']['backBT'].'</a>
		</div>

		<h2 class="sub_headline">'.$GLOBALS['TL_LANG']['MSC']['import_member'][0].'</h2>'.$this->getMessages().'

		<form action="'.ampersand($this->Environment->request, ENCODE_AMPERSANDS).'" id="tl_import_outlook" class="tl_form" method="post">
		<div class="tl_formbody_edit">
		<input type="hidden" name="filepath" value="'.$filepath.'" />
		<input type="hidden" name="FORM_SUBMIT" value="tl_import_outlook" />
		<input type="hidden" name="REQUEST_TOKEN" value="' . REQUEST_TOKEN . '" />

		<div class="tl_tbox">
		  <div>' . $GLOBALS['TL_LANG']['tl_member']['info_import'] . '</div>';

			$result .= '<table>';
			$result .= '<thead><tr><th>' . $GLOBALS['TL_LANG']['tl_member']['outlook_field'][0] . '</th>';
			$result .= '<th>' . $GLOBALS['TL_LANG']['tl_member']['member_field'][0] . '</th></tr></thead><tbody>';
			$oi = deserialize($GLOBALS['TL_CONFIG']['outlook_import'], true);
			$values = array();
			foreach ($oi as $val)
			{
				$foundpair = trimsplit('::', $val);
				$values[$foundpair[0]] = $foundpair[1];
			}
			$index = 0;
			foreach ($fields as $field)
			{
				$objSelect = $this->getMemberFieldWidget($index, $field, $values[$field]);
				$result .= '<tr>';
				$result .= '<td>' . $objSelect->generateLabel() . '</td>';
				$result .= '<td>' . $objSelect->generate() . '</td>';
				$result .= '</tr>';
				$index++;
			}
			$result .= '</tbody></table>';
			$result .= '</div>

		</div>

		<div class="tl_formbody_submit">

		<div class="tl_submit_container">
		<input type="submit" name="import" id="save" class="tl_submit" alt="import member" accesskey="s" value="'.specialchars($GLOBALS['TL_LANG']['tl_member']['start_import']).'" />
		</div>

		</div>
		</form>';
				return $result;
	}
	
	/**
	 * Return the status widget as object
	 * @param mixed
	 * @return object
	 */
	protected function getOutlookFieldWidget($index, $value=null)
	{
		$widget = new SelectMenu();
		$this->loadDataContainer('tl_member');

		$widget->id = 'outlook' . $index;
		$widget->name = 'outlook' . $index;
		$widget->mandatory = false;
		$widget->value = $value;
		$widget->label = $GLOBALS['TL_LANG']['MSC']['ol_field' . $index];

		$arrOptions = array();

		$arrOptions[] = array('value'=> '', 'label'=> '-');
		foreach ($GLOBALS['TL_DCA']['tl_member']['fields'] as $fieldname => $data)
		{
			if (is_array($data['label']))
			{
				$arrOptions[] = array('value'=>$fieldname, 'label'=>$data['label'][0]);
			}
		}

		$widget->options = $arrOptions;

		// Valiate input
		if ($this->Input->post('FORM_SUBMIT') == 'tl_export_outlook')
		{
			$widget->validate();

			if ($widget->hasErrors())
			{
				$this->blnSave = false;
			}
		}

		return $widget;
	}

	/**
	 * Return the status widget as object
	 * @param mixed
	 * @return object
	 */
	protected function getMemberFieldWidget($index, $label, $value=null)
	{
		$widget = new SelectMenu();
		$this->loadDataContainer('tl_member');

		$widget->id = 'outlook' . $index;
		$widget->name = 'outlook' . $index;
		$widget->mandatory = false;
		$widget->value = $label . '::' . $value;
		$widget->label = $label;

		$arrOptions = array();

		$arrOptions[] = array('value'=> '', 'label'=> '-');
		foreach ($GLOBALS['TL_DCA']['tl_member']['fields'] as $fieldname => $data)
		{
			if (is_array($data['label']))
			{
				$arrOptions[] = array('value'=>$label . '::' . $fieldname, 'label'=>$data['label'][0]);
			}
		}

		$widget->options = $arrOptions;

		// Valiate input
		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook')
		{
			$widget->validate();

			if ($widget->hasErrors())
			{
				$this->blnSave = false;
			}
		}

		return $widget;
	}

	/**
	 * Return the file tree widget as object
	 * @param mixed
	 * @return object
	 */
	protected function getFileTreeWidget($value=null)
	{
		$widget = new FileTree();

		$widget->id = 'outlooksource';
		$widget->name = 'outlooksource';
		$widget->strTable = 'tl_member';
		$widget->strField = 'outlooksource';
		$widget->mandatory = true;
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['fieldType'] = 'radio';
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['files'] = true;
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['filesOnly'] = true;
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['extensions'] = 'csv,txt';
		$widget->value = $value;

		$widget->label = $GLOBALS['TL_LANG']['tl_member']['outlooksource'][0];

		if ($GLOBALS['TL_CONFIG']['showHelp'] && strlen($GLOBALS['TL_LANG']['tl_member']['outlooksource'][1]))
		{
			$widget->help = $GLOBALS['TL_LANG']['tl_member']['outlooksource'][1];
		}

		// Valiate input
		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook_fileselection')
		{
			$widget->validate();
			if ($widget->hasErrors())
			{
				$this->blnSave = false;
			}
		}

		return $widget;
	}

	protected function getEncodingWidget($value=null)
	{
		$widget = new SelectMenu();

		$widget->id = 'encoding';
		$widget->name = 'encoding';
		$widget->mandatory = true;
		$widget->value = $value;
		$widget->label = $GLOBALS['TL_LANG']['tl_member']['encoding'][0];

		if ($GLOBALS['TL_CONFIG']['showHelp'] && strlen($GLOBALS['TL_LANG']['tl_member']['encoding'][1]))
		{
			$widget->help = $GLOBALS['TL_LANG']['tl_member']['encoding'][1];
		}

		$arrOptions = array(
			array('value' => 'UTF-8', 'label' => 'UTF-8'),
			array('value' => 'ISO-8859-1', 'label' => 'ISO-8859-1 (Windows)'),
				array('value' => 'UCS-4', 'label' => 'UCS-4'),
				array('value' => 'UCS-4BE', 'label' => 'UCS-4BE'),
				array('value' => 'UCS-4LE', 'label' => 'UCS-4LE'),
				array('value' => 'UCS-2', 'label' => 'UCS-2'),
				array('value' => 'UCS-2BE', 'label' => 'UCS-2BE'),
				array('value' => 'UCS-2LE', 'label' => 'UCS-2LE'),
				array('value' => 'UTF-32', 'label' => 'UTF-32'),
				array('value' => 'UTF-32BE', 'label' => 'UTF-32BE'),
				array('value' => 'UTF-32LE', 'label' => 'UTF-32LE'),
				array('value' => 'UTF-16', 'label' => 'UTF-16'),
				array('value' => 'UTF-16BE', 'label' => 'UTF-16BE'),
				array('value' => 'UTF-16LE', 'label' => 'UTF-16LE'),
				array('value' => 'UTF-7', 'label' => 'UTF-7'),
				array('value' => 'UTF7-IMAP', 'label' => 'UTF7-IMAP'),
				array('value' => 'ASCII', 'label' => 'ASCII'),
				array('value' => 'EUC-JP', 'label' => 'EUC-JP'),
				array('value' => 'SJIS', 'label' => 'SJIS'),
				array('value' => 'eucJP-win', 'label' => 'eucJP-win'),
				array('value' => 'SJIS-win', 'label' => 'SJIS-win'),
				array('value' => 'ISO-2022-JP', 'label' => 'ISO-2022-JP'),
				array('value' => 'ISO-2022-JP-MS', 'label' => 'ISO-2022-JP-MS'),
				array('value' => 'CP932', 'label' => 'CP932'),
				array('value' => 'CP51932', 'label' => 'CP51932'),
				array('value' => 'JIS', 'label' => 'JIS'),
				array('value' => 'JIS-ms', 'label' => 'JIS-ms'),
				array('value' => 'CP50220', 'label' => 'CP50220'),
				array('value' => 'CP50220raw', 'label' => 'CP50220raw'),
				array('value' => 'CP50221', 'label' => 'CP50221'),
				array('value' => 'CP50222', 'label' => 'CP50222'),
				array('value' => 'ISO-8859-2', 'label' => 'ISO-8859-2'),
				array('value' => 'ISO-8859-3', 'label' => 'ISO-8859-3'),
				array('value' => 'ISO-8859-4', 'label' => 'ISO-8859-4'),
				array('value' => 'ISO-8859-5', 'label' => 'ISO-8859-5'),
				array('value' => 'ISO-8859-6', 'label' => 'ISO-8859-6'),
				array('value' => 'ISO-8859-7', 'label' => 'ISO-8859-7'),
				array('value' => 'ISO-8859-8', 'label' => 'ISO-8859-8'),
				array('value' => 'ISO-8859-9', 'label' => 'ISO-8859-9'),
				array('value' => 'ISO-8859-10', 'label' => 'ISO-8859-10'),
				array('value' => 'ISO-8859-13', 'label' => 'ISO-8859-13'),
				array('value' => 'ISO-8859-14', 'label' => 'ISO-8859-14'),
				array('value' => 'ISO-8859-15', 'label' => 'ISO-8859-15'),
				array('value' => 'byte2be', 'label' => 'byte2be'),
				array('value' => 'byte2le', 'label' => 'byte2le'),
				array('value' => 'byte4be', 'label' => 'byte4be'),
				array('value' => 'byte4le', 'label' => 'byte4le'),
				array('value' => 'BASE64', 'label' => 'BASE64'),
				array('value' => 'HTML-ENTITIES', 'label' => 'HTML-ENTITIES'),
				array('value' => '7bit', 'label' => '7bit'),
				array('value' => '8bit', 'label' => '8bit'),
				array('value' => 'EUC-CN', 'label' => 'EUC-CN'),
				array('value' => 'CP936', 'label' => 'CP936'),
				array('value' => 'HZ', 'label' => 'HZ'),
				array('value' => 'EUC-TW', 'label' => 'EUC-TW'),
				array('value' => 'CP950', 'label' => 'CP950'),
				array('value' => 'BIG-5', 'label' => 'BIG-5'),
				array('value' => 'EUC-KR', 'label' => 'EUC-KR'),
				array('value' => 'UHC (CP949)', 'label' => 'UHC (CP949)'),
				array('value' => 'ISO-2022-KR', 'label' => 'ISO-2022-KR'),
				array('value' => 'Windows-1251 (CP1251)', 'label' => 'Windows-1251 (CP1251)'),
				array('value' => 'Windows-1252 (CP1252)', 'label' => 'Windows-1252 (CP1252)'),
				array('value' => 'CP866 (IBM866)', 'label' => 'CP866 (IBM866)'),
				array('value' => 'KOI8-R', 'label' => 'KOI8-R'),
				array('value' => 'ArmSCII-8 (ArmSCII8)', 'label' => 'ArmSCII-8 (ArmSCII8)')
		);
		if (version_compare(phpversion(), '5.4.0', '>=')) {
			$arrOptions[] = array('value' => 'SJIS-mac', 'label' => 'SJIS-mac (alias: MacJapanese)');
			$arrOptions[] = array('value' => 'SJIS-Mobile#DOCOMO', 'label' => 'SJIS-Mobile#DOCOMO (alias: SJIS-DOCOMO)');
			$arrOptions[] = array('value' => 'SJIS-Mobile#KDDI', 'label' => 'SJIS-Mobile#KDDI (alias: SJIS-KDDI)');
			$arrOptions[] = array('value' => 'SJIS-Mobile#SOFTBANK', 'label' => 'SJIS-Mobile#SOFTBANK (alias: SJIS-SOFTBANK)');
			$arrOptions[] = array('value' => 'UTF-8-Mobile#DOCOMO', 'label' => 'UTF-8-Mobile#DOCOMO (alias: UTF-8-DOCOMO)');
			$arrOptions[] = array('value' => 'UTF-8-Mobile#KDDI-A', 'label' => 'UTF-8-Mobile#KDDI-A');
			$arrOptions[] = array('value' => 'UTF-8-Mobile#KDDI-B', 'label' => 'UTF-8-Mobile#KDDI-B (alias: UTF-8-KDDI)');
			$arrOptions[] = array('value' => 'UTF-8-Mobile#SOFTBANK', 'label' => 'UTF-8-Mobile#SOFTBANK (alias: UTF-8-SOFTBANK)');
			$arrOptions[] = array('value' => 'ISO-2022-JP-MOBILE#KDDI', 'label' => 'ISO-2022-JP-MOBILE#KDDI (alias: ISO-2022-JP-KDDI)');
			$arrOptions[] = array('value' => 'GB18030', 'label' => 'GB18030');
		}
		$widget->options = $arrOptions;

		// Valiate input
		if (\Input::post('FORM_SUBMIT') == 'tl_import_outlook_fileselection')
		{
			$widget->validate();

			if ($widget->hasErrors())
			{
				$this->blnSave = false;
			}
		}

		return $widget;
	}

	/**
	 * Return the member groups widget as object
	 * @param mixed
	 * @return object
	 */
	protected function getMembergroupsWidget($value=null)
	{
		$widget = new CheckBox();

		$widget->id = 'membergroup';
		$widget->name = 'membergroup';
		$widget->mandatory = false;
		$widget->label = $GLOBALS['TL_LANG']['tl_member']['outlook_import_groups'][0];
		$widget->multiple = true;
		$options = array();
		$objMember = $this->Database->prepare("SELECT id, name FROM tl_member_group ORDER BY name")->execute();
		while ($objMember->next())
		{
			$data = $objMember->row();
			array_push($options, array('value' => $data['id'], 'label' => $data['name']));
		}
		$widget->options = $options;
		$widget->value = $value;
		if ($GLOBALS['TL_CONFIG']['showHelp'] && strlen($GLOBALS['TL_LANG']['tl_member']['outlook_import_groups'][1]))
		{
			$widget->help = $GLOBALS['TL_LANG']['tl_member']['outlook_import_groups'][1];
		}

		// Valiate input
		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook_fileselection')
		{
			$widget->validate();

			if ($widget->hasErrors())
			{
				$this->blnSave = false;
			}
		}

		return $widget;
	}

	/**
	 * Return the member groups widget as object
	 * @param mixed
	 * @return object
	 */
	protected function getNewsletterWidget($value=null)
	{
		$widget = new CheckBox();

		$widget->id = 'newsletter';
		$widget->name = 'newsletter';
		$widget->mandatory = false;
		$widget->label = $GLOBALS['TL_LANG']['tl_member']['outlook_import_newsletter'][0];
		$widget->multiple = true;
		$options = array();
		$objMember = $this->Database->prepare("SELECT id, title FROM tl_newsletter_channel ORDER BY title")->execute();
		while ($objMember->next())
		{
			$data = $objMember->row();
			array_push($options, array('value' => $data['id'], 'label' => $data['title']));
		}
		$widget->options = $options;
		$widget->value = $value;
		if ($GLOBALS['TL_CONFIG']['showHelp'] && strlen($GLOBALS['TL_LANG']['tl_member']['outlook_import_newsletter'][1]))
		{
			$widget->help = $GLOBALS['TL_LANG']['tl_member']['outlook_import_newsletter'][1];
		}

		// Valiate input
		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook_fileselection')
		{
			$widget->validate();

			if ($widget->hasErrors())
			{
				$this->blnSave = false;
			}
		}

		return $widget;
	}

}
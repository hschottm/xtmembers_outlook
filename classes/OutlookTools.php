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

		$this->Template = new BackendTemplate('be_import_outlook');

		$this->Template->outlooksource = $this->getFileTreeWidget();
		$this->Template->membergroup = $this->getMembergroupsWidget();
		$this->Template->newsletter = $this->getNewsletterWidget();
		$this->Template->hrefBack = ampersand(str_replace('&key=outlook_import', '', $this->Environment->request));
		$this->Template->goBack = $GLOBALS['TL_LANG']['MSC']['goBack'];
		$this->Template->headline = $GLOBALS['TL_LANG']['MSC']['outlook_import'][0];
		$this->Template->request = ampersand($this->Environment->request, ENCODE_AMPERSANDS);
		$this->Template->submit = specialchars($GLOBALS['TL_LANG']['MSC']['continue']);

		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook')
		{
			$output = $this->importMembersFromCSV();
			return $output;
		}
		// Create import form
		if ($this->Input->post('FORM_SUBMIT') == 'tl_import_outlook_fileselection' && $this->blnSave)
		{
			$filename = $this->Template->outlooksource->value;
			$groups = $this->Template->membergroup->value;
			$newsletters = $this->Template->newsletter->value;
			// set session values
			$this->Session->set('outlook_fileid', $filename);
			$this->Session->set('outlook_groups', $groups);
			$this->Session->set('outlook_newsletters', $newsletters);
			$output = $this->importMembersFromCSV();
			return $output;
		}
		return $this->Template->parse();
	}

	public function importMembersFromCSV()
	{
				if ($this->Input->get('key') != 'outlook_import')
				{
					$this->redirect(str_replace('&key=outlook_import', '', $this->Environment->request));
				}

				$f = \FilesModel::findOneById($this->Session->get('outlook_fileid'));
				$file = new \File($f->path);
				$data = $file->getContent();
				$chunks = preg_split("/((?<=\")|(?<=,))[\r\n]+((?=\")|(?=,))/", $data);
				$fields = preg_split("/(,(?=,))|((?<=,),)|(,(?=\"))|((?<=\"),)/", $chunks[0]);
//				$chunks = preg_split("/[\r\n]+/", $data);
//				$fields = trimsplit(";", $chunks[0]);
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
					foreach ($chunks as $idx => $line)
					{
						if ($idx > 0)
						{
							$entities = preg_split("/(,(?=,))|((?<=,),)|(,(?=\"))|((?<=\"),)/", $line);
//							$entities = trimsplit(";", $line);
							foreach ($entities as $ent_idx => $ent_value)
							{
								$ent_value = preg_replace("/(^\")|(\"$)/", "", $ent_value);
								$entities[$ent_idx] = $ent_value;
							}
							$where = array();
							$where_values = array();
							if ($idx_email >= 0)
							{
								if (strlen(utf8_encode($entities[$idx_email])))
								{
									array_push($where, 'email = ?');
									array_push($where_values, utf8_encode($entities[$idx_email]));
								}
							}
							if ($idx_firstname >= 0)
							{
								if (strlen(utf8_encode($entities[$idx_firstname])))
								{
									array_push($where, 'firstname = ?');
									array_push($where_values, utf8_encode($entities[$idx_firstname]));
								}
							}
							if ($idx_lastname >= 0)
							{
								if (strlen(utf8_encode($entities[$idx_lastname])))
								{
									array_push($where, 'lastname = ?');
									array_push($where_values, utf8_encode($entities[$idx_lastname]));
								}
							}
							$where_text = join($where, ' AND ');
							if (strlen($where_text)) $where_text = ' WHERE ' . $where_text;
							$member = null;
							$member = FrontendUser::getInstance();
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
									$field = utf8_encode($field);
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
				$objSelect = $this->getMemberFieldWidget($index, utf8_encode($field), $values[utf8_encode($field)]);
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
		$widget->mandatory = true;
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['fieldType'] = 'radio';
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['files'] = true;
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['filesOnly'] = true;
		$GLOBALS['TL_DCA']['tl_member']['fields']['outlooksource']['eval']['extensions'] = 'csv';
		$widget->strTable = 'tl_member';
		$widget->strField = 'outlooksource';
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
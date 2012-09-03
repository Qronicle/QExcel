<?php
/**
 * QExcel
 *
 * Original Excel5 reader by PHPExcel 1.7.7 (http://www.codeplex.com/PHPExcel)
 * Stripped off all styling, referencing and formula code.
 *
 * Original file header of ParseXL (used as the base for the PHPExcel class):
 * --------------------------------------------------------------------------
 *
 * Adapted from Excel_Spreadsheet_Reader developed by users bizon153,
 * trex005, and mmp11 (SourceForge.net)
 * http: *sourceforge.net/projects/phpexcelreader/
 * Primary changes made by canyoncasa (dvc) for ParseXL 1.00 ...
 *	 Modelled moreso after Perl Excel Parse/Write modules
 *	 Added Parse_Excel_Spreadsheet object
 *		 Reads a whole worksheet or tab as row,column array or as
 *		 associated hash of indexed rows and named column fields
 *	 Added variables for worksheet (tab) indexes and names
 *	 Added an object call for loading individual woorksheets
 *	 Changed default indexing defaults to 0 based arrays
 *	 Fixed date/time and percent formats
 *	 Includes patches found at SourceForge...
 *		 unicode patch by nobody
 *		 unpack("d") machine depedency patch by matchy
 *		 boundsheet utf16 patch by bjaenichen
 *	 Renamed functions for shorter names
 *	 General code cleanup and rigor, including <80 column width
 *	 Included a testcase Excel file and PHP example calls
 *	 Code works for PHP 5.x
 *
 * Primary changes made by canyoncasa (dvc) for ParseXL 1.10 ...
 * http://sourceforge.net/tracker/index.php?func=detail&aid=1466964&group_id=99160&atid=623334
 *	 Decoding of formula conditions, results, and tokens.
 *	 Support for user-defined named cells added as an array "namedcells"
 *		 Patch code for user-defined named cells supports single cells only.
 *		 NOTE: this patch only works for BIFF8 as BIFF5-7 use a different
 *		 external sheet reference structure
 *
 * @package     QExcel_Reader
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 */
/**
 * Excel 5 Reader
 *
 * @package     QExcel_Reader
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-18 14:38
 * @author      ruud.seberechts
 * @author      PHPExcel
 */
class QExcel_Reader_Excel5 extends QExcel_Reader_ReaderAbstract
{
	// ParseXL definitions
	const XLS_BIFF8						= 0x0600;
	const XLS_BIFF7						= 0x0500;
	const XLS_WorkbookGlobals			= 0x0005;
	const XLS_Worksheet					= 0x0010;

	// record identifiers
	const XLS_Type_FORMULA				= 0x0006;
	const XLS_Type_EOF					= 0x000a;
	const XLS_Type_PROTECT				= 0x0012;
	const XLS_Type_OBJECTPROTECT		= 0x0063;
	const XLS_Type_SCENPROTECT			= 0x00dd;
	const XLS_Type_PASSWORD				= 0x0013;
	const XLS_Type_HEADER				= 0x0014;
	const XLS_Type_FOOTER				= 0x0015;
	const XLS_Type_EXTERNSHEET			= 0x0017;
	const XLS_Type_DEFINEDNAME			= 0x0018;
	const XLS_Type_VERTICALPAGEBREAKS	= 0x001a;
	const XLS_Type_HORIZONTALPAGEBREAKS	= 0x001b;
	const XLS_Type_NOTE					= 0x001c;
	const XLS_Type_SELECTION			= 0x001d;
	const XLS_Type_DATEMODE				= 0x0022;
	const XLS_Type_EXTERNNAME			= 0x0023;
	const XLS_Type_LEFTMARGIN			= 0x0026;
	const XLS_Type_RIGHTMARGIN			= 0x0027;
	const XLS_Type_TOPMARGIN			= 0x0028;
	const XLS_Type_BOTTOMMARGIN			= 0x0029;
	const XLS_Type_PRINTGRIDLINES		= 0x002b;
	const XLS_Type_FILEPASS				= 0x002f;
	const XLS_Type_FONT					= 0x0031;
	const XLS_Type_CONTINUE				= 0x003c;
	const XLS_Type_PANE					= 0x0041;
	const XLS_Type_CODEPAGE				= 0x0042;
	const XLS_Type_DEFCOLWIDTH 			= 0x0055;
	const XLS_Type_OBJ					= 0x005d;
	const XLS_Type_COLINFO				= 0x007d;
	const XLS_Type_IMDATA				= 0x007f;
	const XLS_Type_SHEETPR				= 0x0081;
	const XLS_Type_HCENTER				= 0x0083;
	const XLS_Type_VCENTER				= 0x0084;
	const XLS_Type_SHEET				= 0x0085;
	const XLS_Type_PALETTE				= 0x0092;
	const XLS_Type_SCL					= 0x00a0;
	const XLS_Type_PAGESETUP			= 0x00a1;
	const XLS_Type_MULRK				= 0x00bd;
	const XLS_Type_MULBLANK				= 0x00be;
	const XLS_Type_DBCELL				= 0x00d7;
	const XLS_Type_XF					= 0x00e0;
	const XLS_Type_MERGEDCELLS			= 0x00e5;
	const XLS_Type_MSODRAWINGGROUP		= 0x00eb;
	const XLS_Type_MSODRAWING			= 0x00ec;
	const XLS_Type_SST					= 0x00fc;
	const XLS_Type_LABELSST				= 0x00fd;
	const XLS_Type_EXTSST				= 0x00ff;
	const XLS_Type_EXTERNALBOOK			= 0x01ae;
	const XLS_Type_DATAVALIDATIONS		= 0x01b2;
	const XLS_Type_TXO					= 0x01b6;
	const XLS_Type_HYPERLINK			= 0x01b8;
	const XLS_Type_DATAVALIDATION		= 0x01be;
	const XLS_Type_DIMENSION			= 0x0200;
	const XLS_Type_BLANK				= 0x0201;
	const XLS_Type_NUMBER				= 0x0203;
	const XLS_Type_LABEL				= 0x0204;
	const XLS_Type_BOOLERR				= 0x0205;
	const XLS_Type_STRING				= 0x0207;
	const XLS_Type_ROW					= 0x0208;
	const XLS_Type_INDEX				= 0x020b;
	const XLS_Type_ARRAY				= 0x0221;
	const XLS_Type_DEFAULTROWHEIGHT 	= 0x0225;
	const XLS_Type_WINDOW2				= 0x023e;
	const XLS_Type_RK					= 0x027e;
	const XLS_Type_STYLE				= 0x0293;
	const XLS_Type_FORMAT				= 0x041e;
	const XLS_Type_SHAREDFMLA			= 0x04bc;
	const XLS_Type_BOF					= 0x0809;
	const XLS_Type_SHEETPROTECTION		= 0x0867;
	const XLS_Type_RANGEPROTECTION		= 0x0868;
	const XLS_Type_SHEETLAYOUT			= 0x0862;
	const XLS_Type_XFEXT				= 0x087d;
	const XLS_Type_UNKNOWN				= 0xffff;

	/**
	 * Workbook stream data. (Includes workbook globals substream as well as sheet substreams)
	 *
	 * @var string
	 */
	private $_data;

	/**
	 * Size in bytes of $this->_data
	 *
	 * @var int
	 */
	private $_dataSize;

	/**
	 * Current position in stream
	 *
	 * @var integer
	 */
	private $_pos;

	/**
	 * Worksheet that is currently being built by the reader.
	 *
	 * @var QExcel_Worksheet
	 */
	private $_phpSheet;

	/**
	 * BIFF version
	 *
	 * @var int
	 */
	private $_version;

	/**
	 * Codepage set in the Excel file being read. Only important for BIFF5 (Excel 5.0 - Excel 95)
	 * For BIFF8 (Excel 97 - Excel 2003) this will always have the value 'UTF-16LE'
	 *
	 * @var string
	 */
	private $_codepage;

	/**
	 * Worksheets
	 *
	 * @var array
	 */
	private $_aSheets;

	/**
	 * REF structures. Only applies to BIFF8.
	 *
	 * @var array
	 */
	private $_ref;

	/**
	 * Defined names
	 *
	 * @var array
	 */
	private $_definedname;

	/**
	 * Shared strings. Only applies to BIFF8.
	 *
	 * @var array
	 */
	private $_sst;

	/**
	 * Objects. One OBJ record contributes with one entry.
	 *
	 * @var array
	 */
	private $_objs;

	/**
	 * Text Objects. One TXO record corresponds with one entry.
	 *
	 * @var array
	 */
	private $_textObjects;

	/**
	 * Keep track of XF index
	 *
	 * @var int
	 */
	private $_xfIndex;

	/**
	 * Mapping of XF index (that is a cell XF) to final index in cellXf collection
	 *
	 * @var array
	 */
	private $_mapCellXfIndex;

	/**
	 * Mapping of XF index (that is a style XF) to final index in cellStyleXf collection
	 *
	 * @var array
	 */
	private $_mapCellStyleXfIndex;

    public function _init()
    {
        $this->_defaultOptions = array(
            'loadSheet' => null,
        );
    }

    /**
     * Can the Excel 5 Reader open the file?
     *
     * @param string $filename      The desired file
     * @return bool                 Readable
     * @throws Exception            Invalid file
     */
	public function canRead($filename)
	{
		// Check if file exists
		if (!file_exists($filename)) {
			throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
		}

		try {
			// Use ParseXL for the hard work.
			$ole = new PHPExcel_Shared_OLERead();

			// get excel data
			$res = $ole->read($filename);
			return true;

		} catch (Exception $e) {
			return false;
		}
	}


    /**
     * Get the sheet names of a workbook
     *
     * @abstract
     * @param string $filename  The file that should be loaded
     * @return array
     * @throws Exception
     */
	public function getSheetNames($filename)
	{
		// Check if file exists
		if (!file_exists($filename)) {
			throw new Exception("Could not open " . $filename . " for reading! File does not exist.");
		}

		$worksheetNames = array();

		// Read the OLE file
		$this->_loadOLE($filename);

		// total byte size of Excel data (workbook global substream + sheet substreams)
		$this->_dataSize = strlen($this->_data);

		$this->_pos		= 0;
		$this->_aSheets	= array();

		// Parse Workbook Global Substream
		while ($this->_pos < $this->_dataSize) {
			$code = self::_GetInt2d($this->_data, $this->_pos);

			switch ($code) {
				case self::XLS_Type_BOF:	$this->_readBof();		break;
				case self::XLS_Type_SHEET:	$this->_readSheet();	break;
				case self::XLS_Type_EOF:	$this->_toNextRecord();	break 2;
				default:					$this->_toNextRecord();	break;
			}
		}

		foreach ($this->_aSheets as $sheet) {
			if ($sheet['sheetType'] != 0x00) {
				// 0x00: Worksheet, 0x02: Chart, 0x06: Visual Basic module
				continue;
			}

			$worksheetNames[] = $sheet['name'];
		}

		return $worksheetNames;
	}


    /**
     * Load workbook
     *
     * Loads the specified Excel5 file
     *
     * @param string $filename  The file that should be loaded
     * @return QExcel_Workbook  The loaded workbook
     * @throws Exception        Invalid file
     */
	public function load($filename)
	{
		// Read the OLE file
		$this->_loadOLE($filename);

		// total byte size of Excel data (workbook global substream + sheet substreams)
		$this->_dataSize = strlen($this->_data);

		// initialize
		$this->_pos					= 0;
		$this->_codepage			= 'CP1252';
        $this->_aSheets             = array();
		$this->_ref					= array();
		$this->_definedname			= array();
		$this->_sst					= array();
		$this->_xfIndex				= '';
		$this->_mapCellXfIndex		= array();
		$this->_mapCellStyleXfIndex	= array();

		// Parse Workbook Global Substream
		while ($this->_pos < $this->_dataSize) {
			$code = self::_GetInt2d($this->_data, $this->_pos);

			switch ($code) {
				case self::XLS_Type_BOF:			$this->_readBof();				break;
				case self::XLS_Type_FILEPASS:		$this->_readFilepass();			break;
				case self::XLS_Type_CODEPAGE:		$this->_readCodepage();			break;
				case self::XLS_Type_DATEMODE:		$this->_readDateMode();			break;
				case self::XLS_Type_FONT:			$this->_toNextRecord();				break;
				case self::XLS_Type_FORMAT:			$this->_toNextRecord();			break;
				case self::XLS_Type_XF:				$this->_toNextRecord();				break;
				case self::XLS_Type_XFEXT:			$this->_toNextRecord();			break;
				case self::XLS_Type_STYLE:			$this->_toNextRecord();			break;
				case self::XLS_Type_PALETTE:		$this->_toNextRecord();			break;
				case self::XLS_Type_SHEET:			$this->_readSheet();			break;
				case self::XLS_Type_EXTERNALBOOK:	$this->_toNextRecord();		break;
				case self::XLS_Type_EXTERNNAME:		$this->_toNextRecord();		break;
				case self::XLS_Type_EXTERNSHEET:	$this->_toNextRecord();		break;
				case self::XLS_Type_DEFINEDNAME:	$this->_toNextRecord();		break;
				case self::XLS_Type_MSODRAWINGGROUP:$this->_readMsoDrawingGroup();	break;
				case self::XLS_Type_SST:			$this->_readSst();				break;
				case self::XLS_Type_EOF:			$this->_toNextRecord();			break 2;
				default:							$this->_toNextRecord();			break;
			}
		}

		// Parse the individual sheets
		foreach ($this->_aSheets as $i => $sheet) {

            $this->_phpSheet = $this->_workbook->addSheet($sheet['name']);

			if ($sheet['sheetType'] != 0x00) {
				// 0x00: Worksheet, 0x02: Chart, 0x06: Visual Basic module
				continue;
			}

			// check if sheet should be skipped
			if ($this->getLoadSheetsOnly() && !in_array($sheet['name'], $this->getLoadSheetsOnly())) {
				continue;
			}

			$this->_pos = $sheet['offset'];

			// Initialize objs
			$this->_objs = array();

			// Initialize text objs
			$this->_textObjects = array();

			// Initialize cell annotations
			$this->textObjRef = -1;

			while ($this->_pos <= $this->_dataSize - 4) {
				$code = self::_GetInt2d($this->_data, $this->_pos);

				switch ($code) {
					case self::XLS_Type_BOF:					$this->_readBof();						break;
					case self::XLS_Type_PRINTGRIDLINES:			$this->_toNextRecord();			break;
					case self::XLS_Type_DEFAULTROWHEIGHT:		$this->_toNextRecord();			break;
					case self::XLS_Type_SHEETPR:				$this->_toNextRecord();					break;
					case self::XLS_Type_HORIZONTALPAGEBREAKS:	$this->_toNextRecord();		break;
					case self::XLS_Type_VERTICALPAGEBREAKS:		$this->_toNextRecord();		break;
					case self::XLS_Type_HEADER:					$this->_toNextRecord();					break;
					case self::XLS_Type_FOOTER:					$this->_toNextRecord();					break;
					case self::XLS_Type_HCENTER:				$this->_toNextRecord();					break;
					case self::XLS_Type_VCENTER:				$this->_toNextRecord();					break;
					case self::XLS_Type_LEFTMARGIN:				$this->_toNextRecord();				break;
					case self::XLS_Type_RIGHTMARGIN:			$this->_toNextRecord();				break;
					case self::XLS_Type_TOPMARGIN:				$this->_toNextRecord();				break;
					case self::XLS_Type_BOTTOMMARGIN:			$this->_toNextRecord();				break;
					case self::XLS_Type_PAGESETUP:				$this->_toNextRecord();				break;
					case self::XLS_Type_PROTECT:				$this->_toNextRecord();					break;
					case self::XLS_Type_SCENPROTECT:			$this->_toNextRecord();				break;
					case self::XLS_Type_OBJECTPROTECT:			$this->_toNextRecord();			break;
					case self::XLS_Type_PASSWORD:				$this->_toNextRecord();					break;
					case self::XLS_Type_DEFCOLWIDTH:			$this->_toNextRecord();				break;
					case self::XLS_Type_COLINFO:				$this->_toNextRecord();					break;
					case self::XLS_Type_DIMENSION:				$this->_toNextRecord();					break;
					case self::XLS_Type_ROW:					$this->_toNextRecord();						break;
					case self::XLS_Type_DBCELL:					$this->_toNextRecord();					break;
					case self::XLS_Type_RK:						$this->_readRk();						break;
					case self::XLS_Type_LABELSST:				$this->_readLabelSst();					break;
					case self::XLS_Type_MULRK:					$this->_readMulRk();					break;
					case self::XLS_Type_NUMBER:					$this->_readNumber();					break;
					case self::XLS_Type_FORMULA:				$this->_readFormula();					break;
					case self::XLS_Type_SHAREDFMLA:				$this->_toNextRecord();				break;
					case self::XLS_Type_BOOLERR:				$this->_readBoolErr();					break;
					case self::XLS_Type_MULBLANK:				$this->_toNextRecord();					break;
					case self::XLS_Type_LABEL:					$this->_readLabel();					break;
					case self::XLS_Type_BLANK:					$this->_toNextRecord();					break;
					case self::XLS_Type_MSODRAWING:				$this->_readMsoDrawing();				break;
					case self::XLS_Type_OBJ:					$this->_toNextRecord();						break;
					case self::XLS_Type_WINDOW2:				$this->_readWindow2();					break;
					case self::XLS_Type_SCL:					$this->_toNextRecord();						break;
					case self::XLS_Type_PANE:					$this->_toNextRecord();						break;
					case self::XLS_Type_SELECTION:				$this->_toNextRecord();				break;
					case self::XLS_Type_MERGEDCELLS:			$this->_toNextRecord();				break;
					case self::XLS_Type_HYPERLINK:				$this->_toNextRecord();				break;
					case self::XLS_Type_DATAVALIDATIONS:		$this->_toNextRecord();			break;
					case self::XLS_Type_DATAVALIDATION:			$this->_toNextRecord();			break;
					case self::XLS_Type_SHEETLAYOUT:			$this->_toNextRecord();				break;
					case self::XLS_Type_SHEETPROTECTION:		$this->_toNextRecord();			break;
					case self::XLS_Type_RANGEPROTECTION:		$this->_toNextRecord();			break;
					case self::XLS_Type_NOTE:					$this->_toNextRecord();						break;
					//case self::XLS_Type_IMDATA:				$this->_readImData();					break;
					case self::XLS_Type_TXO:					$this->_toNextRecord();				break;
					case self::XLS_Type_CONTINUE:				$this->_toNextRecord();					break;
					case self::XLS_Type_EOF:					$this->_toNextRecord();					break 2;
					default:									$this->_toNextRecord();					break;
				}

			}
		}

		return $this->_workbook;
	}


	/**
	 * Use OLE reader to extract the relevant data streams from the OLE file
	 *
	 * @param string $pFilename
	 */
	private function _loadOLE($pFilename)
	{
		// OLE reader
		$ole = new PHPExcel_Shared_OLERead();

		// get excel data,
		$res = $ole->read($pFilename);
		// Get workbook data: workbook stream + sheet streams
		$this->_data = $ole->getStream($ole->wrkbook);

		// Get summary information data
		//$this->_summaryInformation = $ole->getStream($ole->summaryInformation);

		// Get additional document summary information data
		//$this->_documentSummaryInformation = $ole->getStream($ole->documentSummaryInformation);

		// Get user-defined property data
//		$this->_userDefinedProperties = $ole->getUserDefinedProperties();
	}

    protected function _toNextRecord()
    {
        $length = self::_GetInt2d($this->_data, $this->_pos + 2);
        $this->_pos += 4 + $length;
    }


	/**
	 * Read BOF
	 */
	private function _readBof()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 2; size: 2; type of the following data
		$substreamType = self::_GetInt2d($recordData, 2);

		switch ($substreamType) {
			case self::XLS_WorkbookGlobals:
				$version = self::_GetInt2d($recordData, 0);
				if (($version != self::XLS_BIFF8) && ($version != self::XLS_BIFF7)) {
					throw new Exception('Cannot read this Excel file. Version is too old.');
				}
				$this->_version = $version;
				break;

			case self::XLS_Worksheet:
				// do not use this version information for anything
				// it is unreliable (OpenOffice doc, 5.8), use only version information from the global stream
				break;

			default:
				// substream, e.g. chart
				// just skip the entire substream
				do {
					$code = self::_GetInt2d($this->_data, $this->_pos);
					$this->_toNextRecord();
				} while ($code != self::XLS_Type_EOF && $this->_pos < $this->_dataSize);
				break;
		}
	}


	/**
	 * FILEPASS
	 *
	 * This record is part of the File Protection Block. It
	 * contains information about the read/write password of the
	 * file. All record contents following this record will be
	 * encrypted.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readFilepass()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
//		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		throw new Exception('Cannot read encrypted file');
	}


	/**
	 * CODEPAGE
	 *
	 * This record stores the text encoding used to write byte
	 * strings, stored as MS Windows code page identifier.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readCodepage()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; code page identifier
		$codepage = self::_GetInt2d($recordData, 0);

		$this->_codepage = PHPExcel_Shared_CodePage::NumberToName($codepage);
	}


	/**
	 * DATEMODE
	 *
	 * This record specifies the base date for displaying date
	 * values. All dates are stored as count of days past this
	 * base date. In BIFF2-BIFF4 this record is part of the
	 * Calculation Settings Block. In BIFF5-BIFF8 it is
	 * stored in the Workbook Globals Substream.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readDateMode()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; 0 = base 1900, 1 = base 1904
		PHPExcel_Shared_Date::setExcelCalendar(PHPExcel_Shared_Date::CALENDAR_WINDOWS_1900);
		if (ord($recordData{0}) == 1) {
			PHPExcel_Shared_Date::setExcelCalendar(PHPExcel_Shared_Date::CALENDAR_MAC_1904);
		}
	}

	/**
	 * SHEET
	 *
	 * This record is  located in the  Workbook Globals
	 * Substream  and represents a sheet inside the workbook.
	 * One SHEET record is written for each sheet. It stores the
	 * sheet name and a stream offset to the BOF record of the
	 * respective Sheet Substream within the Workbook Stream.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readSheet()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 4; absolute stream position of the BOF record of the sheet
		$rec_offset = self::_GetInt4d($recordData, 0);

		/*/ offset: 4; size: 1; sheet state
		switch (ord($recordData{4})) {
			case 0x00: $sheetState = PHPExcel_Worksheet::SHEETSTATE_VISIBLE;    break;
			case 0x01: $sheetState = PHPExcel_Worksheet::SHEETSTATE_HIDDEN;     break;
			case 0x02: $sheetState = PHPExcel_Worksheet::SHEETSTATE_VERYHIDDEN; break;
			default: $sheetState = PHPExcel_Worksheet::SHEETSTATE_VISIBLE;      break;
		}//*/

		// offset: 5; size: 1; sheet type
		$sheetType = ord($recordData{5});

		// offset: 6; size: var; sheet name
		if ($this->_version == self::XLS_BIFF8) {
			$string = self::_readUnicodeStringShort(substr($recordData, 6));
			$rec_name = $string['value'];
		} elseif ($this->_version == self::XLS_BIFF7) {
			$string = $this->_readByteStringShort(substr($recordData, 6));
			$rec_name = $string['value'];
		}

		$this->_aSheets[] = array(
			'name' => $rec_name,
			'offset' => $rec_offset,
			//'sheetState' => $sheetState,
			'sheetType' => $sheetType,
		);
	}


	/**
	 * Read MSODRAWINGGROUP record
	 */
	private function _readMsoDrawingGroup()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);

		// get spliced record data
		$splicedRecordData = $this->_getSplicedRecordData();
		$recordData = $splicedRecordData['recordData'];
	}


	/**
	 * SST - Shared String Table
	 *
	 * This record contains a list of all strings used anywhere
	 * in the workbook. Each string occurs only once. The
	 * workbook uses indexes into the list to reference the
	 * strings.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 **/
	private function _readSst()
	{
		// offset within (spliced) record data
		$pos = 0;

		// get spliced record data
		$splicedRecordData = $this->_getSplicedRecordData();

		$recordData = $splicedRecordData['recordData'];
		$spliceOffsets = $splicedRecordData['spliceOffsets'];

		// offset: 0; size: 4; total number of strings in the workbook
		$pos += 4;

		// offset: 4; size: 4; number of following strings ($nm)
		$nm = self::_GetInt4d($recordData, 4);
		$pos += 4;

		// loop through the Unicode strings (16-bit length)
		for ($i = 0; $i < $nm; ++$i) {

			// number of characters in the Unicode string
			$numChars = self::_GetInt2d($recordData, $pos);
			$pos += 2;

			// option flags
			$optionFlags = ord($recordData{$pos});
			++$pos;

			// bit: 0; mask: 0x01; 0 = compressed; 1 = uncompressed
			$isCompressed = (($optionFlags & 0x01) == 0) ;

			// bit: 2; mask: 0x02; 0 = ordinary; 1 = Asian phonetic
			$hasAsian = (($optionFlags & 0x04) != 0);

			// bit: 3; mask: 0x03; 0 = ordinary; 1 = Rich-Text
			$hasRichText = (($optionFlags & 0x08) != 0);

			if ($hasRichText) {
				// number of Rich-Text formatting runs
				$formattingRuns = self::_GetInt2d($recordData, $pos);
				$pos += 2;
			}

			if ($hasAsian) {
				// size of Asian phonetic setting
				$extendedRunLength = self::_GetInt4d($recordData, $pos);
				$pos += 4;
			}

			// expected byte length of character array if not split
			$len = ($isCompressed) ? $numChars : $numChars * 2;

			// look up limit position
			foreach ($spliceOffsets as $spliceOffset) {
				// it can happen that the string is empty, therefore we need
				// <= and not just <
				if ($pos <= $spliceOffset) {
					$limitpos = $spliceOffset;
					break;
				}
			}

			if ($pos + $len <= $limitpos) {
				// character array is not split between records

				$retstr = substr($recordData, $pos, $len);
				$pos += $len;

			} else {
				// character array is split between records

				// first part of character array
				$retstr = substr($recordData, $pos, $limitpos - $pos);

				$bytesRead = $limitpos - $pos;

				// remaining characters in Unicode string
				$charsLeft = $numChars - (($isCompressed) ? $bytesRead : ($bytesRead / 2));

				$pos = $limitpos;

				// keep reading the characters
				while ($charsLeft > 0) {

					// look up next limit position, in case the string span more than one continue record
					foreach ($spliceOffsets as $spliceOffset) {
						if ($pos < $spliceOffset) {
							$limitpos = $spliceOffset;
							break;
						}
					}

					// repeated option flags
					// OpenOffice.org documentation 5.21
					$option = ord($recordData{$pos});
					++$pos;

					if ($isCompressed && ($option == 0)) {
						// 1st fragment compressed
						// this fragment compressed
						$len = min($charsLeft, $limitpos - $pos);
						$retstr .= substr($recordData, $pos, $len);
						$charsLeft -= $len;
						$isCompressed = true;

					} elseif (!$isCompressed && ($option != 0)) {
						// 1st fragment uncompressed
						// this fragment uncompressed
						$len = min($charsLeft * 2, $limitpos - $pos);
						$retstr .= substr($recordData, $pos, $len);
						$charsLeft -= $len / 2;
						$isCompressed = false;

					} elseif (!$isCompressed && ($option == 0)) {
						// 1st fragment uncompressed
						// this fragment compressed
						$len = min($charsLeft, $limitpos - $pos);
						for ($j = 0; $j < $len; ++$j) {
							$retstr .= $recordData{$pos + $j} . chr(0);
						}
						$charsLeft -= $len;
						$isCompressed = false;

					} else {
						// 1st fragment compressed
						// this fragment uncompressed
						$newstr = '';
						for ($j = 0; $j < strlen($retstr); ++$j) {
							$newstr .= $retstr[$j] . chr(0);
						}
						$retstr = $newstr;
						$len = min($charsLeft * 2, $limitpos - $pos);
						$retstr .= substr($recordData, $pos, $len);
						$charsLeft -= $len / 2;
						$isCompressed = false;
					}

					$pos += $len;
				}
			}

			// convert to UTF-8
			$retstr = self::_encodeUTF16($retstr, $isCompressed);

			// read additional Rich-Text information, if any
			$fmtRuns = array();
			if ($hasRichText) {
				// list of formatting runs
				for ($j = 0; $j < $formattingRuns; ++$j) {
					// first formatted character; zero-based
					$charPos = self::_GetInt2d($recordData, $pos + $j * 4);

					// index to font record
					$fontIndex = self::_GetInt2d($recordData, $pos + 2 + $j * 4);

					$fmtRuns[] = array(
						'charPos' => $charPos,
						'fontIndex' => $fontIndex,
					);
				}
				$pos += 4 * $formattingRuns;
			}

			// read additional Asian phonetics information, if any
			if ($hasAsian) {
				// For Asian phonetic settings, we skip the extended string data
				$pos += $extendedRunLength;
			}

			// store the shared sting
			$this->_sst[] = array(
				'value' => $retstr,
				'fmtRuns' => $fmtRuns,
			);
		}

		// _getSplicedRecordData() takes care of moving current position in data stream
	}


	/**
	 * Read RK record
	 * This record represents a cell that contains an RK value
	 * (encoded integer or floating-point value). If a
	 * floating-point value cannot be encoded to an RK value,
	 * a NUMBER record will be written. This record replaces the
	 * record INTEGER written in BIFF2.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readRk()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; index to row
		$row = self::_GetInt2d($recordData, 0);

		// offset: 2; size: 2; index to column
		$column = self::_GetInt2d($recordData, 2);

        // offset: 4; size: 2; index to XF record
        $xfIndex = self::_GetInt2d($recordData, 4);

        // offset: 6; size: 4; RK value
        $rknum = self::_GetInt4d($recordData, 6);
        $numValue = self::_GetIEEE754($rknum);

        $this->_phpSheet->setCell(
            $row,
            $column,
            $this->getCellValue($numValue, PHPExcel_Cell_DataType::TYPE_NUMERIC)
        );
	}

    /**
     * Read LABELSST record
     * This record represents a cell that contains a string. It
     * replaces the LABEL record and RSTRING record used in
     * BIFF2-BIFF5.
     *
     * --	"OpenOffice.org's Documentation of the Microsoft
     * 		Excel File Format"
     */
    private function _readLabelSst()
    {
        $length = self::_GetInt2d($this->_data, $this->_pos + 2);
        $recordData = substr($this->_data, $this->_pos + 4, $length);

        // move stream pointer to next record
        $this->_pos += 4 + $length;

        // offset: 0; size: 2; index to row
        $row = self::_GetInt2d($recordData, 0);

        // offset: 2; size: 2; index to column
        $column = self::_GetInt2d($recordData, 2);

        // offset: 4; size: 2; index to XF record
        $xfIndex = self::_GetInt2d($recordData, 4);

        // offset: 6; size: 4; index to SST record
        $index = self::_GetInt4d($recordData, 6);

        $this->_phpSheet->setCell(
            $row,
            $column,
            $this->getCellValue($this->_sst[$index]['value'], PHPExcel_Cell_DataType::TYPE_STRING)
        );
    }


	/**
	 * Read MULRK record
	 * This record represents a cell range containing RK value
	 * cells. All cells are located in the same row.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readMulRk()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; index to row
		$row = self::_GetInt2d($recordData, 0);

		// offset: 2; size: 2; index to first column
		$colFirst = self::_GetInt2d($recordData, 2);

		// offset: var; size: 2; index to last column
		$colLast = self::_GetInt2d($recordData, $length - 2);
		$columns = $colLast - $colFirst + 1;

		// offset within record data
		$offset = 4;

		for ($i = 0; $i < $columns; ++$i) {
            $column = $colFirst + $i;

            // offset: var; size: 2; index to XF record
            $xfIndex = self::_GetInt2d($recordData, $offset);

            // offset: var; size: 4; RK value
            $numValue = self::_GetIEEE754(self::_GetInt4d($recordData, $offset + 2));

            $this->_phpSheet->setCell(
                $row,
                $column,
                $this->getCellValue($numValue, PHPExcel_Cell_DataType::TYPE_NUMERIC)
            );

			$offset += 6;
		}
	}

	/**
	 * Read NUMBER record
	 * This record represents a cell that contains a
	 * floating-point value.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readNumber()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; index to row
		$row = self::_GetInt2d($recordData, 0);

		// offset: 2; size 2; index to column
		$column = self::_GetInt2d($recordData, 2);

        $xfIndex = self::_GetInt2d($recordData, 4);

        $numValue = self::_extractNumber(substr($recordData, 6, 8));

        $this->_phpSheet->setCell(
            $row,
            $column,
            $this->getCellValue($numValue, PHPExcel_Cell_DataType::TYPE_NUMERIC)
        );
	}


	/**
	 * Read FORMULA record + perhaps a following STRING record if formula result is a string
	 * This record contains the token array and the result of a
	 * formula cell.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readFormula()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; row index
		$row = self::_GetInt2d($recordData, 0);

		// offset: 2; size: 2; col index
		$column = self::_GetInt2d($recordData, 2);

		// offset: 20: size: variable; formula structure
		$formulaStructure = substr($recordData, 20);

		// offset: 14: size: 2; option flags, recalculate always, recalculate on open etc.
		$options = self::_GetInt2d($recordData, 14);

        // offset: 16: size: 4; not used

        // offset: 4; size: 2; XF index
        $xfIndex = self::_GetInt2d($recordData, 4);

        // offset: 6; size: 8; result of the formula
        if ( (ord($recordData{6}) == 0)
            && (ord($recordData{12}) == 255)
            && (ord($recordData{13}) == 255) ) {

            // String formula. Result follows in appended STRING record
            $dataType = PHPExcel_Cell_DataType::TYPE_STRING;

            // read STRING record
            $value = $this->_readString();

        } elseif ((ord($recordData{6}) == 1)
            && (ord($recordData{12}) == 255)
            && (ord($recordData{13}) == 255)) {

            // Boolean formula. Result is in +2; 0=false, 1=true
            $dataType = PHPExcel_Cell_DataType::TYPE_BOOL;
            $value = (bool) ord($recordData{8});

        } elseif ((ord($recordData{6}) == 2)
            && (ord($recordData{12}) == 255)
            && (ord($recordData{13}) == 255)) {

            // Error formula. Error code is in +2
            $dataType = PHPExcel_Cell_DataType::TYPE_ERROR;
            $value = self::_mapErrorCode(ord($recordData{8}));

        } elseif ((ord($recordData{6}) == 3)
            && (ord($recordData{12}) == 255)
            && (ord($recordData{13}) == 255)) {

            // Formula result is a null string
            $dataType = PHPExcel_Cell_DataType::TYPE_NULL;
            $value = '';

        } else {

            // forumla result is a number, first 14 bytes like _NUMBER record
            $dataType = PHPExcel_Cell_DataType::TYPE_NUMERIC;
            $value = self::_extractNumber(substr($recordData, 6, 8));

        }

        $this->_phpSheet->setCell($row, $column, $this->getCellValue($value, $dataType));

        /*$cell = $this->_phpSheet->getCell($columnString . ($row + 1));

        // store the formula
        if (!$isPartOfSharedFormula) {
            // not part of shared formula
            // add cell value. If we can read formula, populate with formula, otherwise just used cached value
            try {
                if ($this->_version != self::XLS_BIFF8) {
                    throw new Exception('Not BIFF8. Can only read BIFF8 formulas');
                }
                $formula = $this->_getFormulaFromStructure($formulaStructure); // get formula in human language
                $this->_phpSheet->setCell($row, $column, $this->getCellValue('=' . $formula, PHPExcel_Cell_DataType::TYPE_FORMULA));

            } catch (Exception $e) {
                $this->_phpSheet->setCell($row, $column, $this->getCellValue($value, $dataType));
            }
        } else {
            if ($this->_version == self::XLS_BIFF8) {
                // do nothing at this point, formula id added later in the code
            } else {
                $this->_phpSheet->setCell($row, $column, $this->getCellValue($value, $dataType));
            }
        }

        // store the cached calculated value
        $cell->setCalculatedValue($value);*/
	}


	/**
	 * Read a STRING record from current stream position and advance the stream pointer to next record
	 * This record is used for storing result from FORMULA record when it is a string, and
	 * it occurs directly after the FORMULA record
	 *
	 * @return string The string contents as UTF-8
	 */
	private function _readString()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		if ($this->_version == self::XLS_BIFF8) {
			$string = self::_readUnicodeStringLong($recordData);
			$value = $string['value'];
		} else {
			$string = $this->_readByteStringLong($recordData);
			$value = $string['value'];
		}

		return $value;
	}


	/**
	 * Read BOOLERR record
	 * This record represents a Boolean value or error value
	 * cell.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readBoolErr()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; row index
		$row = self::_GetInt2d($recordData, 0);

		// offset: 2; size: 2; column index
		$column = self::_GetInt2d($recordData, 2);

        // offset: 4; size: 2; index to XF record
        $xfIndex = self::_GetInt2d($recordData, 4);

        // offset: 6; size: 1; the boolean value or error value
        $boolErr = ord($recordData{6});

        // offset: 7; size: 1; 0=boolean; 1=error
        $isError = ord($recordData{7});

        switch ($isError) {
            case 0: // boolean
                $value = (bool) $boolErr;
                $this->_phpSheet->setCell($row, $column, $this->getCellValue($value, PHPExcel_Cell_DataType::TYPE_BOOL));
                break;

            case 1: // error type
                $value = self::_mapErrorCode($boolErr);
                $this->_phpSheet->setCell($row, $column, $this->getCellValue($value, PHPExcel_Cell_DataType::TYPE_ERROR));
                break;
        }
	}

	/**
	 * Read LABEL record
	 * This record represents a cell that contains a string. In
	 * BIFF8 it is usually replaced by the LABELSST record.
	 * Excel still uses this record, if it copies unformatted
	 * text cells to the clipboard.
	 *
	 * --	"OpenOffice.org's Documentation of the Microsoft
	 * 		Excel File Format"
	 */
	private function _readLabel()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; index to row
		$row = self::_GetInt2d($recordData, 0);

		// offset: 2; size: 2; index to column
		$column = self::_GetInt2d($recordData, 2);

        // offset: 4; size: 2; XF index
        $xfIndex = self::_GetInt2d($recordData, 4);

        // add cell value
        // todo: what if string is very long? continue record
        if ($this->_version == self::XLS_BIFF8) {
            $string = self::_readUnicodeStringLong(substr($recordData, 6));
            $value = $string['value'];
        } else {
            $string = $this->_readByteStringLong(substr($recordData, 6));
            $value = $string['value'];
        }
        $this->_phpSheet->setCell($row, $column, $this->getCellValue($value, PHPExcel_Cell_DataType::TYPE_STRING));
	}


	/**
	 * Read MSODRAWING record
	 */
	private function _readMsoDrawing()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		// get spliced record data
		$splicedRecordData = $this->_getSplicedRecordData();
	}


	/**
	 * Read WINDOW2 record
	 */
	private function _readWindow2()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);
		$recordData = substr($this->_data, $this->_pos + 4, $length);

		// move stream pointer to next record
		$this->_pos += 4 + $length;

		// offset: 0; size: 2; option flags
		$options = self::_GetInt2d($recordData, 0);

		// bit: 10; mask: 0x0400; 0 = sheet not active, 1 = sheet active
		$isActive = (bool) ((0x0400 & $options) >> 10);
		if ($isActive) {
            $this->_phpSheet->active = true;
		}
	}


	private function _includeCellRangeFiltered($cellRangeAddress)
	{
		$includeCellRange = true;
		if ($this->getReadFilter() !== NULL) {
			$includeCellRange = false;
			$rangeBoundaries = PHPExcel_Cell::getRangeBoundaries($cellRangeAddress);
			$rangeBoundaries[1][0]++;
			for ($row = $rangeBoundaries[0][1]; $row <= $rangeBoundaries[1][1]; $row++) {
				for ($column = $rangeBoundaries[0][0]; $column != $rangeBoundaries[1][0]; $column++) {
					if ($this->getReadFilter()->readCell($column, $row, $this->_phpSheet->getTitle())) {
						$includeCellRange = true;
						break 2;
					}
				}
			}
		}
		return $includeCellRange;
	}


	/**
	 * Read IMDATA record
	 */
	private function _readImData()
	{
		$length = self::_GetInt2d($this->_data, $this->_pos + 2);

		// get spliced record data
		$splicedRecordData = $this->_getSplicedRecordData();
	}


	/**
	 * Reads a record from current position in data stream and continues reading data as long as CONTINUE
	 * records are found. Splices the record data pieces and returns the combined string as if record data
	 * is in one piece.
	 * Moves to next current position in data stream to start of next record different from a CONtINUE record
	 *
	 * @return array
	 */
	private function _getSplicedRecordData()
	{
		$data = '';
		$spliceOffsets = array();

		$i = 0;
		$spliceOffsets[0] = 0;

		do {
			++$i;

			// offset: 0; size: 2; identifier
			$identifier = self::_GetInt2d($this->_data, $this->_pos);
			// offset: 2; size: 2; length
			$length = self::_GetInt2d($this->_data, $this->_pos + 2);
			$data .= substr($this->_data, $this->_pos + 4, $length);

			$spliceOffsets[$i] = $spliceOffsets[$i - 1] + $length;

			$this->_pos += 4 + $length;
			$nextIdentifier = self::_GetInt2d($this->_data, $this->_pos);
		}
		while ($nextIdentifier == self::XLS_Type_CONTINUE);

		$splicedData = array(
			'recordData' => $data,
			'spliceOffsets' => $spliceOffsets,
		);

		return $splicedData;

	}

	/**
	 * Reads a cell address in BIFF8 e.g. 'A2' or '$A$2'
	 * section 3.3.4
	 *
	 * @param string $cellAddressStructure
	 * @return string
	 */
	private function _readBIFF8CellAddress($cellAddressStructure)
	{
		// offset: 0; size: 2; index to row (0... 65535) (or offset (-32768... 32767))
		$row = self::_GetInt2d($cellAddressStructure, 0) + 1;

		// offset: 2; size: 2; index to column or column offset + relative flags

			// bit: 7-0; mask 0x00FF; column index
			$column = PHPExcel_Cell::stringFromColumnIndex(0x00FF & self::_GetInt2d($cellAddressStructure, 2));

			// bit: 14; mask 0x4000; (1 = relative column index, 0 = absolute column index)
			if (!(0x4000 & self::_GetInt2d($cellAddressStructure, 2))) {
				$column = '$' . $column;
			}
			// bit: 15; mask 0x8000; (1 = relative row index, 0 = absolute row index)
			if (!(0x8000 & self::_GetInt2d($cellAddressStructure, 2))) {
				$row = '$' . $row;
			}

		return $column . $row;
	}


	/**
	 * Reads a cell address in BIFF8 for shared formulas. Uses positive and negative values for row and column
	 * to indicate offsets from a base cell
	 * section 3.3.4
	 *
	 * @param string $cellAddressStructure
	 * @param string $baseCell Base cell, only needed when formula contains tRefN tokens, e.g. with shared formulas
	 * @return string
	 */
	private function _readBIFF8CellAddressB($cellAddressStructure, $baseCell = 'A1')
	{
		list($baseCol, $baseRow) = PHPExcel_Cell::coordinateFromString($baseCell);
		$baseCol = PHPExcel_Cell::columnIndexFromString($baseCol) - 1;

		// offset: 0; size: 2; index to row (0... 65535) (or offset (-32768... 32767))
			$rowIndex = self::_GetInt2d($cellAddressStructure, 0);
			$row = self::_GetInt2d($cellAddressStructure, 0) + 1;

		// offset: 2; size: 2; index to column or column offset + relative flags

			// bit: 7-0; mask 0x00FF; column index
			$colIndex = 0x00FF & self::_GetInt2d($cellAddressStructure, 2);

			// bit: 14; mask 0x4000; (1 = relative column index, 0 = absolute column index)
			if (!(0x4000 & self::_GetInt2d($cellAddressStructure, 2))) {
				$column = PHPExcel_Cell::stringFromColumnIndex($colIndex);
				$column = '$' . $column;
			} else {
				$colIndex = ($colIndex <= 127) ? $colIndex : $colIndex - 256;
				$column = PHPExcel_Cell::stringFromColumnIndex($baseCol + $colIndex);
			}

			// bit: 15; mask 0x8000; (1 = relative row index, 0 = absolute row index)
			if (!(0x8000 & self::_GetInt2d($cellAddressStructure, 2))) {
				$row = '$' . $row;
			} else {
				$rowIndex = ($rowIndex <= 32767) ? $rowIndex : $rowIndex - 65536;
				$row = $baseRow + $rowIndex;
			}

		return $column . $row;
	}


	/**
	 * Reads a cell range address in BIFF5 e.g. 'A2:B6' or 'A1'
	 * always fixed range
	 * section 2.5.14
	 *
	 * @param string $subData
	 * @return string
	 * @throws Exception
	 */
	private function _readBIFF5CellRangeAddressFixed($subData)
	{
		// offset: 0; size: 2; index to first row
		$fr = self::_GetInt2d($subData, 0) + 1;

		// offset: 2; size: 2; index to last row
		$lr = self::_GetInt2d($subData, 2) + 1;

		// offset: 4; size: 1; index to first column
		$fc = ord($subData{4});

		// offset: 5; size: 1; index to last column
		$lc = ord($subData{5});

		// check values
		if ($fr > $lr || $fc > $lc) {
			throw new Exception('Not a cell range address');
		}

		// column index to letter
		$fc = PHPExcel_Cell::stringFromColumnIndex($fc);
		$lc = PHPExcel_Cell::stringFromColumnIndex($lc);

		if ($fr == $lr and $fc == $lc) {
			return "$fc$fr";
		}
		return "$fc$fr:$lc$lr";
	}


	/**
	 * Reads a cell range address in BIFF8 e.g. 'A2:B6' or 'A1'
	 * always fixed range
	 * section 2.5.14
	 *
	 * @param string $subData
	 * @return string
	 * @throws Exception
	 */
	private function _readBIFF8CellRangeAddressFixed($subData)
	{
		// offset: 0; size: 2; index to first row
		$fr = self::_GetInt2d($subData, 0) + 1;

		// offset: 2; size: 2; index to last row
		$lr = self::_GetInt2d($subData, 2) + 1;

		// offset: 4; size: 2; index to first column
		$fc = self::_GetInt2d($subData, 4);

		// offset: 6; size: 2; index to last column
		$lc = self::_GetInt2d($subData, 6);

		// check values
		if ($fr > $lr || $fc > $lc) {
			throw new Exception('Not a cell range address');
		}

		// column index to letter
		$fc = PHPExcel_Cell::stringFromColumnIndex($fc);
		$lc = PHPExcel_Cell::stringFromColumnIndex($lc);

		if ($fr == $lr and $fc == $lc) {
			return "$fc$fr";
		}
		return "$fc$fr:$lc$lr";
	}


	/**
	 * Reads a cell range address in BIFF8 e.g. 'A2:B6' or '$A$2:$B$6'
	 * there are flags indicating whether column/row index is relative
	 * section 3.3.4
	 *
	 * @param string $subData
	 * @return string
	 */
	private function _readBIFF8CellRangeAddress($subData)
	{
		// todo: if cell range is just a single cell, should this funciton
		// not just return e.g. 'A1' and not 'A1:A1' ?

		// offset: 0; size: 2; index to first row (0... 65535) (or offset (-32768... 32767))
			$fr = self::_GetInt2d($subData, 0) + 1;

		// offset: 2; size: 2; index to last row (0... 65535) (or offset (-32768... 32767))
			$lr = self::_GetInt2d($subData, 2) + 1;

		// offset: 4; size: 2; index to first column or column offset + relative flags

		// bit: 7-0; mask 0x00FF; column index
		$fc = PHPExcel_Cell::stringFromColumnIndex(0x00FF & self::_GetInt2d($subData, 4));

		// bit: 14; mask 0x4000; (1 = relative column index, 0 = absolute column index)
		if (!(0x4000 & self::_GetInt2d($subData, 4))) {
			$fc = '$' . $fc;
		}

		// bit: 15; mask 0x8000; (1 = relative row index, 0 = absolute row index)
		if (!(0x8000 & self::_GetInt2d($subData, 4))) {
			$fr = '$' . $fr;
		}

		// offset: 6; size: 2; index to last column or column offset + relative flags

		// bit: 7-0; mask 0x00FF; column index
		$lc = PHPExcel_Cell::stringFromColumnIndex(0x00FF & self::_GetInt2d($subData, 6));

		// bit: 14; mask 0x4000; (1 = relative column index, 0 = absolute column index)
		if (!(0x4000 & self::_GetInt2d($subData, 6))) {
			$lc = '$' . $lc;
		}

		// bit: 15; mask 0x8000; (1 = relative row index, 0 = absolute row index)
		if (!(0x8000 & self::_GetInt2d($subData, 6))) {
			$lr = '$' . $lr;
		}

		return "$fc$fr:$lc$lr";
	}


	/**
	 * Reads a cell range address in BIFF8 for shared formulas. Uses positive and negative values for row and column
	 * to indicate offsets from a base cell
	 * section 3.3.4
	 *
	 * @param string $subData
	 * @param string $baseCell Base cell
	 * @return string Cell range address
	 */
	private function _readBIFF8CellRangeAddressB($subData, $baseCell = 'A1')
	{
		list($baseCol, $baseRow) = PHPExcel_Cell::coordinateFromString($baseCell);
		$baseCol = PHPExcel_Cell::columnIndexFromString($baseCol) - 1;

		// TODO: if cell range is just a single cell, should this funciton
		// not just return e.g. 'A1' and not 'A1:A1' ?

		// offset: 0; size: 2; first row
		$frIndex = self::_GetInt2d($subData, 0); // adjust below

		// offset: 2; size: 2; relative index to first row (0... 65535) should be treated as offset (-32768... 32767)
		$lrIndex = self::_GetInt2d($subData, 2); // adjust below

		// offset: 4; size: 2; first column with relative/absolute flags

		// bit: 7-0; mask 0x00FF; column index
		$fcIndex = 0x00FF & self::_GetInt2d($subData, 4);

		// bit: 14; mask 0x4000; (1 = relative column index, 0 = absolute column index)
		if (!(0x4000 & self::_GetInt2d($subData, 4))) {
			// absolute column index
			$fc = PHPExcel_Cell::stringFromColumnIndex($fcIndex);
			$fc = '$' . $fc;
		} else {
			// column offset
			$fcIndex = ($fcIndex <= 127) ? $fcIndex : $fcIndex - 256;
			$fc = PHPExcel_Cell::stringFromColumnIndex($baseCol + $fcIndex);
		}

		// bit: 15; mask 0x8000; (1 = relative row index, 0 = absolute row index)
		if (!(0x8000 & self::_GetInt2d($subData, 4))) {
			// absolute row index
			$fr = $frIndex + 1;
			$fr = '$' . $fr;
		} else {
			// row offset
			$frIndex = ($frIndex <= 32767) ? $frIndex : $frIndex - 65536;
			$fr = $baseRow + $frIndex;
		}

		// offset: 6; size: 2; last column with relative/absolute flags

		// bit: 7-0; mask 0x00FF; column index
		$lcIndex = 0x00FF & self::_GetInt2d($subData, 6);
		$lcIndex = ($lcIndex <= 127) ? $lcIndex : $lcIndex - 256;
		$lc = PHPExcel_Cell::stringFromColumnIndex($baseCol + $lcIndex);

		// bit: 14; mask 0x4000; (1 = relative column index, 0 = absolute column index)
		if (!(0x4000 & self::_GetInt2d($subData, 6))) {
			// absolute column index
			$lc = PHPExcel_Cell::stringFromColumnIndex($lcIndex);
			$lc = '$' . $lc;
		} else {
			// column offset
			$lcIndex = ($lcIndex <= 127) ? $lcIndex : $lcIndex - 256;
			$lc = PHPExcel_Cell::stringFromColumnIndex($baseCol + $lcIndex);
		}

		// bit: 15; mask 0x8000; (1 = relative row index, 0 = absolute row index)
		if (!(0x8000 & self::_GetInt2d($subData, 6))) {
			// absolute row index
			$lr = $lrIndex + 1;
			$lr = '$' . $lr;
		} else {
			// row offset
			$lrIndex = ($lrIndex <= 32767) ? $lrIndex : $lrIndex - 65536;
			$lr = $baseRow + $lrIndex;
		}

		return "$fc$fr:$lc$lr";
	}


	/**
	 * Read BIFF8 cell range address list
	 * section 2.5.15
	 *
	 * @param string $subData
	 * @return array
	 */
	private function _readBIFF8CellRangeAddressList($subData)
	{
		$cellRangeAddresses = array();

		// offset: 0; size: 2; number of the following cell range addresses
		$nm = self::_GetInt2d($subData, 0);

		$offset = 2;
		// offset: 2; size: 8 * $nm; list of $nm (fixed) cell range addresses
		for ($i = 0; $i < $nm; ++$i) {
			$cellRangeAddresses[] = $this->_readBIFF8CellRangeAddressFixed(substr($subData, $offset, 8));
			$offset += 8;
		}

		return array(
			'size' => 2 + 8 * $nm,
			'cellRangeAddresses' => $cellRangeAddresses,
		);
	}


	/**
	 * Read BIFF5 cell range address list
	 * section 2.5.15
	 *
	 * @param string $subData
	 * @return array
	 */
	private function _readBIFF5CellRangeAddressList($subData)
	{
		$cellRangeAddresses = array();

		// offset: 0; size: 2; number of the following cell range addresses
		$nm = self::_GetInt2d($subData, 0);

		$offset = 2;
		// offset: 2; size: 6 * $nm; list of $nm (fixed) cell range addresses
		for ($i = 0; $i < $nm; ++$i) {
			$cellRangeAddresses[] = $this->_readBIFF5CellRangeAddressFixed(substr($subData, $offset, 6));
			$offset += 6;
		}

		return array(
			'size' => 2 + 6 * $nm,
			'cellRangeAddresses' => $cellRangeAddresses,
		);
	}


	/**
	 * Get a sheet range like Sheet1:Sheet3 from REF index
	 * Note: If there is only one sheet in the range, one gets e.g Sheet1
	 * It can also happen that the REF structure uses the -1 (FFFF) code to indicate deleted sheets,
	 * in which case an exception is thrown
	 *
	 * @param int $index
	 * @return string|false
	 * @throws Exception
	 */
	private function _readSheetRangeByRefIndex($index)
	{
		/*if (isset($this->_ref[$index])) {

			$type = $this->_externalBooks[$this->_ref[$index]['externalBookIndex']]['type'];

			switch ($type) {
				case 'internal':
					// check if we have a deleted 3d reference
					if ($this->_ref[$index]['firstSheetIndex'] == 0xFFFF or $this->_ref[$index]['lastSheetIndex'] == 0xFFFF) {
						throw new Exception('Deleted sheet reference');
					}

					// we have normal sheet range (collapsed or uncollapsed)
					$firstSheetName = $this->_aSheets[$this->_ref[$index]['firstSheetIndex']]['name'];
					$lastSheetName = $this->_aSheets[$this->_ref[$index]['lastSheetIndex']]['name'];

					if ($firstSheetName == $lastSheetName) {
						// collapsed sheet range
						$sheetRange = $firstSheetName;
					} else {
						$sheetRange = "$firstSheetName:$lastSheetName";
					}

					// escape the single-quotes
					$sheetRange = str_replace("'", "''", $sheetRange);

					// if there are special characters, we need to enclose the range in single-quotes
					// todo: check if we have identified the whole set of special characters
					// it seems that the following characters are not accepted for sheet names
					// and we may assume that they are not present: []*:/\?
					if (preg_match("/[ !\"@#£$%&{()}<>=+'|^,;-]/", $sheetRange)) {
						$sheetRange = "'$sheetRange'";
					}

					return $sheetRange;
					break;

				default:
					// TODO: external sheet support
					throw new Exception('Excel5 reader only supports internal sheets in fomulas');
					break;
			}
		}*/
		return false;
	}


	/**
	 * read BIFF8 constant value array from array data
	 * returns e.g. array('value' => '{1,2;3,4}', 'size' => 40}
	 * section 2.5.8
	 *
	 * @param string $arrayData
	 * @return array
	 */
	private static function _readBIFF8ConstantArray($arrayData)
	{
		// offset: 0; size: 1; number of columns decreased by 1
		$nc = ord($arrayData[0]);

		// offset: 1; size: 2; number of rows decreased by 1
		$nr = self::_GetInt2d($arrayData, 1);
		$size = 3; // initialize
		$arrayData = substr($arrayData, 3);

		// offset: 3; size: var; list of ($nc + 1) * ($nr + 1) constant values
		$matrixChunks = array();
		for ($r = 1; $r <= $nr + 1; ++$r) {
			$items = array();
			for ($c = 1; $c <= $nc + 1; ++$c) {
				$constant = self::_readBIFF8Constant($arrayData);
				$items[] = $constant['value'];
				$arrayData = substr($arrayData, $constant['size']);
				$size += $constant['size'];
			}
			$matrixChunks[] = implode(',', $items); // looks like e.g. '1,"hello"'
		}
		$matrix = '{' . implode(';', $matrixChunks) . '}';

		return array(
			'value' => $matrix,
			'size' => $size,
		);
	}


	/**
	 * read BIFF8 constant value which may be 'Empty Value', 'Number', 'String Value', 'Boolean Value', 'Error Value'
	 * section 2.5.7
	 * returns e.g. array('value' => '5', 'size' => 9)
	 *
	 * @param string $valueData
	 * @return array
	 */
	private static function _readBIFF8Constant($valueData)
	{
		// offset: 0; size: 1; identifier for type of constant
		$identifier = ord($valueData[0]);

		switch ($identifier) {
		case 0x00: // empty constant (what is this?)
			$value = '';
			$size = 9;
			break;
		case 0x01: // number
			// offset: 1; size: 8; IEEE 754 floating-point value
			$value = self::_extractNumber(substr($valueData, 1, 8));
			$size = 9;
			break;
		case 0x02: // string value
			// offset: 1; size: var; Unicode string, 16-bit string length
			$string = self::_readUnicodeStringLong(substr($valueData, 1));
			$value = '"' . $string['value'] . '"';
			$size = 1 + $string['size'];
			break;
		case 0x04: // boolean
			// offset: 1; size: 1; 0 = FALSE, 1 = TRUE
			if (ord($valueData[1])) {
				$value = 'TRUE';
			} else {
				$value = 'FALSE';
			}
			$size = 9;
			break;
		case 0x10: // error code
			// offset: 1; size: 1; error code
			$value = self::_mapErrorCode(ord($valueData[1]));
			$size = 9;
			break;
		}
		return array(
			'value' => $value,
			'size' => $size,
		);
	}

	/**
	 * Read byte string (8-bit string length)
	 * OpenOffice documentation: 2.5.2
	 *
	 * @param string $subData
	 * @return array
	 */
	private function _readByteStringShort($subData)
	{
		// offset: 0; size: 1; length of the string (character count)
		$ln = ord($subData[0]);

		// offset: 1: size: var; character array (8-bit characters)
		$value = $this->_decodeCodepage(substr($subData, 1, $ln));

		return array(
			'value' => $value,
			'size' => 1 + $ln, // size in bytes of data structure
		);
	}


	/**
	 * Read byte string (16-bit string length)
	 * OpenOffice documentation: 2.5.2
	 *
	 * @param string $subData
	 * @return array
	 */
	private function _readByteStringLong($subData)
	{
		// offset: 0; size: 2; length of the string (character count)
		$ln = self::_GetInt2d($subData, 0);

		// offset: 2: size: var; character array (8-bit characters)
		$value = $this->_decodeCodepage(substr($subData, 2));

		//return $string;
		return array(
			'value' => $value,
			'size' => 2 + $ln, // size in bytes of data structure
		);
	}


	/**
	 * Extracts an Excel Unicode short string (8-bit string length)
	 * OpenOffice documentation: 2.5.3
	 * function will automatically find out where the Unicode string ends.
	 *
	 * @param string $subData
	 * @return array
	 */
	private static function _readUnicodeStringShort($subData)
	{
		$value = '';

		// offset: 0: size: 1; length of the string (character count)
		$characterCount = ord($subData[0]);

		$string = self::_readUnicodeString(substr($subData, 1), $characterCount);

		// add 1 for the string length
		$string['size'] += 1;

		return $string;
	}


	/**
	 * Extracts an Excel Unicode long string (16-bit string length)
	 * OpenOffice documentation: 2.5.3
	 * this function is under construction, needs to support rich text, and Asian phonetic settings
	 *
	 * @param string $subData
	 * @return array
	 */
	private static function _readUnicodeStringLong($subData)
	{
		$value = '';

		// offset: 0: size: 2; length of the string (character count)
		$characterCount = self::_GetInt2d($subData, 0);

		$string = self::_readUnicodeString(substr($subData, 2), $characterCount);

		// add 2 for the string length
		$string['size'] += 2;

		return $string;
	}


	/**
	 * Read Unicode string with no string length field, but with known character count
	 * this function is under construction, needs to support rich text, and Asian phonetic settings
	 * OpenOffice.org's Documentation of the Microsoft Excel File Format, section 2.5.3
	 *
	 * @param string $subData
	 * @param int $characterCount
	 * @return array
	 */
	private static function _readUnicodeString($subData, $characterCount)
	{
		$value = '';

		// offset: 0: size: 1; option flags

			// bit: 0; mask: 0x01; character compression (0 = compressed 8-bit, 1 = uncompressed 16-bit)
			$isCompressed = !((0x01 & ord($subData[0])) >> 0);

			// bit: 2; mask: 0x04; Asian phonetic settings
			$hasAsian = (0x04) & ord($subData[0]) >> 2;

			// bit: 3; mask: 0x08; Rich-Text settings
			$hasRichText = (0x08) & ord($subData[0]) >> 3;

		// offset: 1: size: var; character array
		// this offset assumes richtext and Asian phonetic settings are off which is generally wrong
		// needs to be fixed
		$value = self::_encodeUTF16(substr($subData, 1, $isCompressed ? $characterCount : 2 * $characterCount), $isCompressed);

		return array(
			'value' => $value,
			'size' => $isCompressed ? 1 + $characterCount : 1 + 2 * $characterCount, // the size in bytes including the option flags
		);
	}


	/**
	 * Convert UTF-8 string to string surounded by double quotes. Used for explicit string tokens in formulas.
	 * Example:  hello"world  -->  "hello""world"
	 *
	 * @param string $value UTF-8 encoded string
	 * @return string
	 */
	private static function _UTF8toExcelDoubleQuoted($value)
	{
		return '"' . str_replace('"', '""', $value) . '"';
	}


	/**
	 * Reads first 8 bytes of a string and return IEEE 754 float
	 *
	 * @param string $data Binary string that is at least 8 bytes long
	 * @return float
	 */
	private static function _extractNumber($data)
	{
		$rknumhigh = self::_GetInt4d($data, 4);
		$rknumlow = self::_GetInt4d($data, 0);
		$sign = ($rknumhigh & 0x80000000) >> 31;
		$exp = (($rknumhigh & 0x7ff00000) >> 20) - 1023;
		$mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
		$mantissalow1 = ($rknumlow & 0x80000000) >> 31;
		$mantissalow2 = ($rknumlow & 0x7fffffff);
		$value = $mantissa / pow( 2 , (20 - $exp));

		if ($mantissalow1 != 0) {
			$value += 1 / pow (2 , (21 - $exp));
		}

		$value += $mantissalow2 / pow (2 , (52 - $exp));
		if ($sign) {
			$value *= -1;
		}

		return $value;
	}


	private static function _GetIEEE754($rknum)
	{
		if (($rknum & 0x02) != 0) {
			$value = $rknum >> 2;
		} else {
			// changes by mmp, info on IEEE754 encoding from
			// research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
			// The RK format calls for using only the most significant 30 bits
			// of the 64 bit floating point value. The other 34 bits are assumed
			// to be 0 so we use the upper 30 bits of $rknum as follows...
			$sign = ($rknum & 0x80000000) >> 31;
			$exp = ($rknum & 0x7ff00000) >> 20;
			$mantissa = (0x100000 | ($rknum & 0x000ffffc));
			$value = $mantissa / pow( 2 , (20- ($exp - 1023)));
			if ($sign) {
				$value = -1 * $value;
			}
			//end of changes by mmp
		}
		if (($rknum & 0x01) != 0) {
			$value /= 100;
		}
		return $value;
	}


	/**
	 * Get UTF-8 string from (compressed or uncompressed) UTF-16 string
	 *
	 * @param string $string
	 * @param bool $compressed
	 * @return string
	 */
	private static function _encodeUTF16($string, $compressed = '')
	{
		if ($compressed) {
			$string = self::_uncompressByteString($string);
 		}

		return PHPExcel_Shared_String::ConvertEncoding($string, 'UTF-8', 'UTF-16LE');
	}


	/**
	 * Convert UTF-16 string in compressed notation to uncompressed form. Only used for BIFF8.
	 *
	 * @param string $string
	 * @return string
	 */
	private static function _uncompressByteString($string)
	{
		$uncompressedString = '';
		$strLen = strlen($string);
		for ($i = 0; $i < $strLen; ++$i) {
			$uncompressedString .= $string[$i] . "\0";
		}

		return $uncompressedString;
	}


	/**
	 * Convert string to UTF-8. Only used for BIFF5.
	 *
	 * @param string $string
	 * @return string
	 */
	private function _decodeCodepage($string)
	{
		return PHPExcel_Shared_String::ConvertEncoding($string, 'UTF-8', $this->_codepage);
	}


	/**
	 * Read 16-bit unsigned integer
	 *
	 * @param string $data
	 * @param int $pos
	 * @return int
	 */
	public static function _GetInt2d($data, $pos)
	{
		return ord($data[$pos]) | (ord($data[$pos+1]) << 8);
	}


	/**
	 * Read 32-bit signed integer
	 *
	 * @param string $data
	 * @param int $pos
	 * @return int
	 */
	public static function _GetInt4d($data, $pos)
	{
		// FIX: represent numbers correctly on 64-bit system
		// http://sourceforge.net/tracker/index.php?func=detail&aid=1487372&group_id=99160&atid=623334
		// Hacked by Andreas Rehm 2006 to ensure correct result of the <<24 block on 32 and 64bit systems
		$_or_24 = ord($data[$pos + 3]);
		if ($_or_24 >= 128) {
			// negative number
			$_ord_24 = -abs((256 - $_or_24) << 24);
		} else {
			$_ord_24 = ($_or_24 & 127) << 24;
		}
		return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $_ord_24;
	}


	/**
	 * Map error code, e.g. '#N/A'
	 *
	 * @param int $subData
	 * @return string
	 */
	private static function _mapErrorCode($subData)
	{
		switch ($subData) {
			case 0x00: return '#NULL!';		break;
			case 0x07: return '#DIV/0!';	break;
			case 0x0F: return '#VALUE!';	break;
			case 0x17: return '#REF!';		break;
			case 0x1D: return '#NAME?';		break;
			case 0x24: return '#NUM!';		break;
			case 0x2A: return '#N/A';		break;
			default: return false;
		}
	}

	private function _parseRichText($is = '') {
		$value = new PHPExcel_RichText();

		$value->createText($is);

		return $value;
	}

}

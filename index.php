<php header('Content-Type: text/html; charset=utf-8'); ?>
<!doctype html>
<head>
    <title>QExcel test suite</title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <style type="text/css">
        html, body, table, form {
            font-family: Arial;
            font-size: 0.9em;
        }
        h2 {
            border-bottom: 1px solid black;
        }
    </style>
</head>
<body>
<?php

require_once('common.php');
require_once('library/QExcel/QExcel.php');

session_start();
$sessionOptions = isset($_SESSION['options']) ? $_SESSION['options'] : array();
if (isset($_POST['filter'])) {
    $sessionOptions = array_merge($sessionOptions, $_POST);
}
$_SESSION['options'] = $sessionOptions;



########################################################################################################################
## File form ###########################################################################################################
########################################################################################################################

echo '<h2>Excel file</h2>';

print formStart();

$tFiles = array(
    WEB_DIR . '/test/files/example1.xlsx',
    WEB_DIR . '/test/files/example2.xls',
    WEB_DIR . '/test/files/export.csv',
    WEB_DIR . '/test/files/example2.xml'
);
$files = array();
foreach ($tFiles as $file) {
    $files[$file] = basename($file);
}

$excelFile = isset($_POST['file']) ? $_POST['file'] : (isset($_SESSION['file']) ? $_SESSION['file'] : $tFiles[0]);
$_SESSION['file'] = $excelFile;
if (isset($_POST['file'])) {
    $_SESSION['options'] = $sessionOptions = array();
}

print formElement(
    formLabel('file', 'Load file'),
    formSelect('file', $excelFile, $files)
);

print formElement('', formSubmit('set-file', 'Change file'));

print formEnd();

########################################################################################################################
## Options form ########################################################################################################
########################################################################################################################

$reader = QExcel::createReaderForFile($excelFile);
$options = $reader->getDefaultOptions();
$loadSheet = 0;

echo '<h2>Options</h2>';

print formStart();

foreach ($options as $option => $defaultValue)
{
    switch ($option)
    {
        case 'encoding':
            // Some random encodings found on iconvlib
            $tEncodings = array('ASCII', 'ISO-8859-1', 'ISO-8859-2', 'ISO-8859-3', 'ISO-8859-4', 'ISO-8859-5', 'ISO-8859-7', 'ISO-8859-9',
                'ISO-8859-10', 'ISO-8859-13', 'ISO-8859-14', 'ISO-8859-15', 'ISO-8859-16', 'KOI8-R', 'KOI8-U', 'KOI8-RU', 'CP1250',
                'CP1251', 'CP1252', 'CP1253', 'CP1254', 'CP1257', 'CP850', 'CP866', 'CP1131', 'MacRoman', 'MacCentralEurope', 'MacIceland',
                'MacCroatian', 'MacRomania', 'MacCyrillic', 'MacUkraine', 'MacGreek', 'MacTurkish', 'Macintosh', 'UTF-8', 'UTF-16LE',
                'UTF-16BE', 'UTF-32LE');
            // We need a key value array for our formSelect 'view helper'
            $encodings = array();
            foreach ($tEncodings as $encoding) {
                $encodings[$encoding] = $encoding;
            }
            // print the option
            $optionValue = getOption('encoding', $defaultValue);
            print formElement(
                formLabel('encoding', 'File encoding'),
                formSelect('encoding', $optionValue, $encodings)
            );
            // Apply the option
            $reader->setOption('encoding', $optionValue);
            break;
        case 'delimiter':
            // Delimiter options
            $delimiters = array('' => 'Auto-detect');
            foreach ($reader->getOption('possibleDelimiters') as $delimiter) {
                $delimiters[$delimiter] = str_replace("\t", '\\t', $delimiter);
            }
            // Form element
            $optionValue = getOption('delimiter', $defaultValue);
            print formElement(
                formLabel('delimiter', 'CSV delimiter'),
                formSelect('delimiter', $optionValue, $delimiters)
            );
            // Apply delimiter
            $reader->setOption('delimiter', $optionValue);
            break;
        case 'enclosure':
            // Form element
            $optionValue = getOption('enclosure', $defaultValue);
            print formElement(
                formLabel('enclosure', 'CSV enclosure'),
                formSelect('enclosure', $optionValue, array(
                    '"' => 'Double quote (")',
                    "'" => "Single quote (')",
                ))
            );
            // Apply delimiter
            $reader->setOption('enclosure', $optionValue);
            break;
        case 'loadSheet':
            // Available sheets
            $tSheets = $reader->getSheetNames($excelFile);
            $sheets = array();
            foreach ($tSheets as $sheet) {
                $sheets[$sheet] = $sheet;
            }
            // Form element
            $optionValue = getOption('loadSheet', $tSheets[0]);
            print formElement(
                formLabel('loadSheet', 'Excel sheet'),
                formSelect('loadSheet', $optionValue, $sheets)
            );
            // Apply sheet
            $loadSheet = $optionValue;
            $reader->setOption('loadSheet', $optionValue);
            break;
    }
}

function getOption($name, $default = null)
{
    global $sessionOptions;
    return array_key_exists($name, $sessionOptions) ? $sessionOptions[$name] : $default;
}

print formElement('', formSubmit('filter', 'Save options'));
print formEnd();

########################################################################################################################
## Print Excel sheet ###################################################################################################
########################################################################################################################

echo '<h2>Preview</h2>';

$workbook = $reader->load($excelFile);
$sheet = $workbook->getSheet($loadSheet);

print '<table border="1" cellpadding="1" cellspacing="0">';

for ($row = 0; $row < min(10, $sheet->getNumRows()); $row++) {
    print '<tr>';
    for ($col = 0; $col < $sheet->getNumColumns(); $col++) {
        print '<td>'.$sheet->getCellValue($row, $col).'</td>';
    }
    print '</tr>';
}

print '</table>';

if ($sheet->getNumRows() > 10) {
    print '<p>'.($sheet->getNumRows()-10).' rows more in loaded file</p>';
}


########################################################################################################################
## Form View Helpers ###################################################################################################
########################################################################################################################

function formStart()
{
    return '<form action="" method="post"><table border="0">';
}

function formEnd()
{
    return '</table></form>';
}

function formElement($label, $input)
{
    return '<tr><td>'.$label.'</td><td>'.$input.'</td></tr>';
}

function formLabel($name, $value)
{
    return '<label for="'.htmlentities($name).'">'.htmlentities($value).'</label>';
}

function formSelect($name, $value, array $options)
{
    $html = '<select name="'.htmlentities($name).'" id="'.htmlentities($name).'">';
    foreach ($options as $val => $label) {
        $html .= '<option value="'.htmlentities($val).'"'.($val == $value ? ' selected="selected"' : '').'>'.htmlentities($label).'</option>';
    }
    $html .= '</select>';
    return $html;
}

function formSubmit($name, $value)
{
    return '<input type="submit" name="'.htmlentities($name).'" value="'.htmlentities($value).'" />';
}
?>
</body>
</html>
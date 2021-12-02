<?php

namespace QExcel;

use QExcel\Reader\AbstractReader;
use QExcel\Reader\CSV;
use QExcel\Reader\Excel2003XML;
use QExcel\Reader\Excel2007;
use QExcel\Reader\Excel5;

/**
 * QExcel
 *
 * The QExcel class acts as a Reader Factory for QExcel.
 * It provides methods to dynamically create readers or load entire workbooks.
 *
 * By default QExcel will only load Readers from its own library.
 * This can be changed however, by adding custom reader paths and/or types.
 *
 * You will probably always need a custom reader path if you want to add additional (or extending) readers.
 * You can add a path by defining its location, and the class prefix that is used.
 * For example if you have your readers stored in 'library/MyLib/Reader' and you follow the PEAR naming conventions,
 * you should add 'library/MyLib/Readers/' with prefix 'MyLib_Reader_'.
 *
 * Custom reader types (defaults are for example 'Excel5' and 'Excel2007') can be added, each with optional linked file
 * extensions. If you have a class 'MyLib_Reader_AwesomeReader' you should add 'AwesomeReader' with
 * extensions ['awe','som']
 * Remember that custom classes should always extend QExcel_Reader_ReaderAbstract
 *
 * Custom paths and readers will always be loaded before the default ones, following last in, first out.
 *
 * @package     QExcel
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-20 18:48
 * @author      ruud.seberechts
 */
class QExcel
{
    /**
     * Reader paths
     *
     * Reader paths are defined as [$path => $classPrefix]
     * The default path is set in the QExcel file
     *
     * @var array
     */
    protected static $readers = [
        'CSV'          => CSV::class,
        'Excel5'       => Excel5::class,
        'Excel2007'    => Excel2007::class,
        'Excel2003XML' => Excel2003XML::class,
    ];

    /**
     * Reader types
     *
     * Reader types are defined as [$readerType => $linkedExtensions[]]
     *
     * @var array
     */
    protected static $_readerTypes = array(
        'CSV'          => array('csv'),
        'Excel5'       => array('xls', 'xlsm'),
        'Excel2003XML' => array('xml'),
        'Excel2007'    => array('xlsx'),
    );

    /**
     * Add a reader path
     *
     * For more information about adding custom readers, please check the class documentation.
     * If a path already exists, this method will merge overwrite the existing class prefix.
     *
     * @static
     * @param string $readerType
     * @param string $className
     */
    public static function addReader($readerType, $className)
    {
        self::$readers[$readerType] = $$className;
    }

    /**
     * Add a reader type
     *
     * For more information about adding custom readers, please check the class documentation.
     * If a type already exists, this method will merge the linked extensions.
     *
     * @static
     * @param string $readerType The reader type (end of class name)
     * @param array  $linkedExtensions File extensions linked to the type (optional)
     */
    public static function addReaderType($readerType, array $linkedExtensions = array())
    {
        if (array_key_exists($readerType, self::$_readerTypes)) {
            $linkedExtensions = array_merge(self::$_readerTypes[$readerType], $linkedExtensions);
        }
        self::$_readerTypes[$readerType] = $linkedExtensions;
    }

    /**
     * Get all reader types
     *
     * Returns in reverse order to give priority to custom types and make sure the CSV reader is in last position
     *
     * Reader types are defined as [$readerType => $linkedExtensions[]]
     *
     * @static
     * @return array    Reader classes
     */
    public static function getReaderTypes()
    {
        return array_reverse(self::$_readerTypes, true);
    }

    /**
     * Get all registered readers
     *
     * Returns in reverse order to give priority to custom readers
     *
     * Readers are defined as [$readerType => $className]
     *
     * @static
     * @return array    Reader paths
     */
    public static function getReaders()
    {
        return array_reverse(self::$readers, true);
    }

    /**
     * Create a reader for a file
     *
     * The system will loop through all available readers (both default and custom) and try to find a fitting one.
     * First the reader is searched by file extension. And if that doesn't work out, it just tries them all.
     *
     * When the file exists, but doesn't really fit one of the profiles, the CSV Reader will be returned.
     *
     * @static
     * @param string $filename The file that we want to create a reader for
     * @return bool|AbstractReader    The matching reader, or FALSE
     */
    public static function createReaderForFile($filename)
    {
        // Check file exists

        if (!file_exists($filename)) {
            return false;
        }

        // Try to load file by looking at the extension
        $filenameParts = explode('.', $filename);
        $filenameExt = array_pop($filenameParts);
        $ReaderTypes = self::getReaderTypes();
        foreach ($ReaderTypes as $readerType => $extensions) {
            if ($readerType == 'CSV') continue;
            foreach ($extensions as $extension) {
                if ($filenameExt == $extension) {
                    $reader = self::createReader($readerType);
                    if ($reader && $reader->canRead($filename)) {
                        return $reader;
                    }
                }
            }
        }

        // Try to load file by looping through every possibility
        foreach ($ReaderTypes as $readerType => $extensions) {
            $reader = self::createReader($readerType);
            if ($reader && $reader->canRead($filename)) {
                return $reader;
            }
        }

        // We shouldn't get here because CSV should load everything, but anyway
        return false;
    }

    /**
     * Create a reader for a certain file type
     *
     * This method will first try to include the custom classes and paths
     *
     * @static
     * @param string $readerType The reader that should be created
     * @return ?AbstractReader
     */
    public static function createReader(string $readerType): ?AbstractReader
    {
        if (!array_key_exists($readerType, self::$_readerTypes)) {
            return null;
        }
        $readers = self::getReaders();
        if (isset($readers[$readerType])) {
            $className = $readers[$readerType];
            return new $className();
        }
        return null;
    }

    /**
     * Load a workbook
     *
     * If the reader type is known, it is advised to pass it as the second argument.
     * Otherwise the system will try to figure out which reader to use.
     *
     * @static
     * @param string $filename File that should be loaded
     * @param string $readerType Reader type (optional, but faster)
     * @return Workbook|bool     The loaded workbook, or FALSE
     */
    public static function loadWorkbook($filename, $readerType = null)
    {
        $reader = $readerType ? self::createReader($readerType) : self::createReaderForFile($filename);
        return $reader ? $reader->load($filename) : false;
    }
}
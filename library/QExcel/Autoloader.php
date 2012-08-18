<?php
/**
 * QExcel
 *
 * Original Autoloader by PHPExcel (http://www.codeplex.com/PHPExcel)
 *
 * @package     QExcel
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 */


QExcel_Autoloader::register();
/*PHPExcel_Shared_ZipStreamWrapper::register();
// check mbstring.func_overload
if (ini_get('mbstring.func_overload') & 2) {
    throw new Exception('Multibyte function overloading in PHP must be disabled for string functions (2).');
}
PHPExcel_Shared_String::buildCharacterSets();*/




/**
 * Autoloader
 *
 * @package     QExcel
 * @copyright   2012 Qronicle (http://www.qronicle.be)
 * @license     GNU LGPL (http://www.gnu.org/licenses/lgpl.txt)
 * @link        http://www.qronicle.be
 * @since       2012-08-18 15:12
 * @author      ruud.seberechts
 * @author      PHPExcel
 */
class QExcel_Autoloader
{
	/**
	 * Register the Autoloader with SPL
	 */
	public static function register()
    {
		if (function_exists('__autoload')) {
			//	Register any existing autoloader function with SPL, so we don't get any clashes
			spl_autoload_register('__autoload');
		}
		//	Register ourselves with SPL
		return spl_autoload_register(array('QExcel_Autoloader', 'load'));
	}


	/**
	 * Autoload a class identified by name
	 *
	 * @param string    $pClassName		Name of the object to load
     * @return bool|void                Returns false when the class could not be loaded
	 */
	public static function load($pClassName)
    {
		if ((class_exists($pClassName)) || (strpos($pClassName, 'QExcel') !== 0 && strpos($pClassName, 'PHPExcel') !== 0)) {
			//	Either already loaded, or no QExcel class request
			return false;
		}

		$pObjectFilePath = QEXCEL_ROOT . str_replace('_', DIRECTORY_SEPARATOR, str_replace('PHPExcel_', 'QExcel_', $pClassName)) . '.php';

		if ((file_exists($pObjectFilePath) === false) || (is_readable($pObjectFilePath) === false)) {
			//	Can't load
			return false;
		}

		require($pObjectFilePath);
	}

}
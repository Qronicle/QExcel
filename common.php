<?php

/**
 * Collection of constant definitions and utility functions
 * 
 * @package Common
 */

###############################################################################################################
# CONSTANTS DEFINITION ########################################################################################
###############################################################################################################

/**
 * Request scheme
 * 
 * The current request scheme, for example 'http' or 'https'.
 * 
 * @var string
 */
define('SCHEME', (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] == 'on') ? 'https://' : 'http://');

/** Domain name */
if(isset($_SERVER['HTTP_X_FORWARDED_SERVER'])) {
 $_SERVER['SERVER_NAME'] = $_SERVER['HTTP_X_FORWARDED_SERVER'];
}

/**
 * Hostname
 * 
 * The current request's host name, for example 'www.mywebsite.com'.
 * 
 * @var string
 */
define('HOST_NAME', $_SERVER['SERVER_NAME']);

$tmp = explode('.', $_SERVER['SERVER_NAME']);
if($tmp[0] == 'www') {
    array_shift($tmp);
}
define('DOMAIN_NAME', implode('.', $tmp));

/** 
 * Web directory
 * 
 * The public web directory, for example '/var/www/vhosts/my-site/www'.
 * 
 * @var string
 */
define('WEB_DIR', $_SERVER['DOCUMENT_ROOT']);

/**
 * Root directory
 * 
 * The hosts base directory, for example '/var/www/vhosts/my-site'.
 * 
 * @var string
 */
define('ROOT_DIR', dirname(__FILE__));

/**
 * Application directory
 * 
 * The directory with the Zend application, for example '/var/www/vhosts/my-site/application'.
 * 
 * @var string
 */
define('APPLICATION_DIR', ROOT_DIR.'/application');

/**
 * Data directory
 * 
 * The data directory, for example '/var/www/vhosts/my-site/data'.
 * 
 * @var string
 */
define('DATA_DIR', ROOT_DIR.'/data');

/**
 * Base URL
 * 
 * The current request hostname + server path
 * @todo Remove SERVER_PATH ?
 * 
 * @var unknown_type
 */
define('BASE_URL', $_SERVER['SERVER_NAME']);

/**
 * Image URL
 * 
 * The image URL, for example 'http://www.my-site.com/images'
 * 
 * @var string
 */
define('IMAGE_URL', SCHEME.BASE_URL.'/images');

/**
 * Module URL
 * 
 * The module public URL, for example 'http://www.my-site.com/modules'
 * 
 * @var string
 */
define('MODULE_URL', SCHEME.BASE_URL.'/modules');

/**
 * Data URL
 * 
 * The data URL, for example 'http://www.my-site.com/data'
 * 
 * @var string
 */
define('DATA_URL', SCHEME.BASE_URL.'/data');

/**
 * Stylesheets URL
 * 
 * The data URL, for example 'http://www.my-site.com/css'
 * 
 * @var string
 */
define('CSS_URL', SCHEME.BASE_URL.'/css');

/**
 * Scripts URL
 * 
 * The data URL, for example 'http://www.my-site.com/scripts'
 * 
 * @var string
 */
define('SCRIPTS_URL', SCHEME.BASE_URL.'/scripts');

/**
 * CGI URL
 * 
 * The CGI URL, for example 'http://www.my-site.com/cgi-bin'
 * 
 * @var string
 */
define('CGI_URL', SCHEME.BASE_URL.'/cgi-bin');

###############################################################################################################
# INCLUDE PATH ################################################################################################
###############################################################################################################
    
// Ensure library and application folders are in the include_path
set_include_path(implode(PATH_SEPARATOR, array(
    ROOT_DIR . '/library', '.', 
    ROOT_DIR . '/application',
)));

###############################################################################################################
# DUMP and all variations #####################################################################################
###############################################################################################################


/**
 * Dump browser-formatted variables.
 * 
 * Works a lot like var_dump.
 * 
 * Usage:
 * <code>dump($foo, $bar);</code>
 * 
 * @return void
 */
function dump ()
{
	$args = func_get_args();
	foreach ($args as $obj) {
		print '<pre>';
		print objToString($obj);
		print '</pre>';
	}
}

function tdump()
{
    $args = func_get_args();
    $title = array_shift($args);
    echo '<hr><b>'.$title.'</b><hr>';
    call_user_func_array('dump', $args);
}

/**
 * Converts all types to a string
 * 
 * Kindoff mimics var_dump, but only of one variable, and returns the result instead of printing it.
 * 
 * @param mixed $obj	The object that should be converted to a string
 * @return string
 */
function objToString($obj)
{
	if ($obj instanceof Zend_Db_Select) {
		return $obj->__toString();
    } elseif ($obj instanceof Exception) {
        return exceptionToString($obj);
	} elseif (is_bool($obj)) {
		return $obj ? 'bool(true)' : 'bool(false)';
	} elseif (is_string($obj)) {
		return 'string(' . strlen($obj) . ') "' . $obj . '"';
	} elseif (is_null($obj)) {
		return 'NULL';
	} elseif (is_array($obj) || is_object($obj)) {
		return print_r($obj, true);
	} else {
		return $obj;
	}
}

/**
 * Creates exception debug info string
 *
 * @param Exception $ex
 */
function exceptionToString(Exception $ex)
{
    return get_class($ex) . " Object\n{\n"
       . "    [message]  => " . $ex->getMessage() . "\n"
       . "    [code]     => " . $ex->getCode() . "\n"
       . "    [file]     => " . $ex->getFile() . "\n"
       . "    [line]     => " . $ex->getLine() . "\n"
       . "    [previous] => " . $ex->getPrevious() . "\n"
       . '    [trace]    => <div style="margin: 10px 0 10px 60px">' . $ex->getTraceAsString() . "</div>"
       . '}';
}

/**
 * Dump variables, then die.
 * 
 * Usage:
 * <code>ddump($foo, $bar);</code>
 * 
 * @return void
 */
function ddump ()
{
	$args = func_get_args();
    call_user_func_array('dump', $args);
    die;
}


function dtdump ()
{
    $args = func_get_args();
    call_user_func_array('tdump', $args);
    die;
}

/**
 * Get the dump of variables as a string
 * 
 * Usage:
 * <code>$log = sdump($foo, $bar);</code>
 * 
 * @return string
 */
function sdump()
{
	$args = func_get_args();
	$sdump = '';
	foreach ($args as $obj) {
		$sdump .= objToString($obj) . "\n\n";
	}
	return $sdump;
}

$__dump_enabled = null;
$__fdump_cleaned = array();
/**
 * Dump variables to a file in the web root to the 'f.dump' file.
 * 
 * Usage:
 * <code>fdump($foo, $bar, ..., $fileName, $clearPerRequest);</code>
 * 
 * This method can be used the same as the other dump functions, there are only two optional arguments you can end with:
 * <ul>
 *  <li><i>string</i> <b>$fileName</b>: Optional, a filename for the dump file (should end with .dump). Defaults to 'f.dump'.</li>
 *  <li><i>bool</i> <b>$clearPerRequest</b>: Optional, whether the file should be cleared when this is the first time this request this file is used to dump to. Defaults to 'true'.</li>
 * </ul>
 *
 * @return void
 */
function fdump()
{
	$fileName = 'f.dump';
	$clearPerRequest = true;
	
	$args = func_get_args();
	$numArgs = count($args);
	
	//check last argument is a string that ends in .dump
	if ($numArgs) {
		$last = end($args);
		if (is_string($last) && substr($last, -5) == '.dump') {
			$fileName = array_pop($args);
		}
		//check second to last argument is a string that ends in .dump and last argument is a boolean
		elseif ($numArgs > 1) {
			$last = end($args);
			$flast = $args[$numArgs-2];
			if (is_bool($last) && is_string($flast) && substr($flast, -5) == '.dump') {
				$clearPerRequest = array_pop($args); 
				$fileName = array_pop($args);
			}
		}
	}
	
	// Clear file when clearPerRequest is true, this file wasn't cleared yet, and exists
	global $__fdump_cleaned;
	if ($clearPerRequest && !isset($__fdump_cleaned[$fileName]) && file_exists($fileName)) {
		unlink($fileName);
	}
	$__fdump_cleaned[$fileName] = true;
	
	// Append args dump to file
	$file = fopen($fileName, 'a+');
	fwrite($file, '== ' . date('Y-m-d H:i:s') . " ====================================================\n\n");
	foreach ($args as $obj) {
		fwrite($file, objToString($obj) . "\n\n");
	}
	fwrite($file, "\n\n");
	
	fclose($file);
}

###############################################################################################################
# Recursive delete and move functions #########################################################################
###############################################################################################################

/**
 * Recursive remove directory
 * 
 * @param string $directory	The directory to remove
 * @param string $empty		Whether to only delete the directory contents and not the directory itself. Defaults to false.
 * @return bool				Whether the directory was deleted successfully
 */
function rmdir_r($directory, $empty=FALSE)
{
	// if the path has a slash at the end we remove it here
	if(substr($directory,-1) == '/') {
		$directory = substr($directory,0,-1);
	}

	// if the path is not valid or is not a directory ...
	if (!file_exists($directory) || !is_dir($directory)) {
		return FALSE;
	}
	// ... if the path is not readable
	elseif(!is_readable($directory)) {
		return FALSE;
	}
	// ... else if the path is readable
	else {
		// we open the directory
		$handle = opendir($directory);

		// and scan through the items inside
		while (FALSE !== ($item = readdir($handle))) {
			// if the filepointer is not the current directory
			// or the parent directory
			if ($item != '.' && $item != '..') {
				// we build the new path to delete
				$path = $directory.'/'.$item;
				// if the new path is a directory
				if (is_dir($path)) {
					// we call this function with the new path
					rmdir_r($path);
				}
				// if the new path is a file
				else {
					// we remove the file
					unlink($path);
				}
			}
		}
		// close the directory
		closedir($handle);

		// if the option to empty is not set to true
		if ($empty == FALSE) {
			// try to delete the now empty directory
			if (!rmdir($directory)) {
				// return false if not possible
				return FALSE;
			}
		}
		// return success
		return TRUE;
	}
}

/**
 * Moves all files and folders inside the source directory to the destination directory
 * 
 * @param string $srcDir		Source directory - ends with '/'
 * @param string $destDir		Destination directory - ends with '/'
 * @param bool $removeSrcDir	Whether to remove the source directory after moving it's files. Defaults to true.
 * @return void
 */
function move_content_r($srcDir, $destDir, $removeSrcDir = true)
{
	$dirHandle = opendir($srcDir);
	while ($file = readdir($dirHandle)) {
		if ($file != "." && $file != ".." && $file != "__MACOSX") {
			// Move directory
			if (is_dir($srcDir.$file)) {
				move_r($srcDir.$file, $destDir);
			}
			// Move file
			else {
				rename($srcDir.$file, $destDir.$file);
			}
		}
	}
	
	if ($removeSrcDir) {
		rmdir_r($srcDir);
	}
}

/**
 * 
 */
/**
 * Recursive move function
 * 
 * Recursive function to copy all subdirectories and contents
 * 
 * @param string $dirsource	The source directory
 * @param string $dirdest	The destination directory
 * @return void
 */
function move_r($dirsource, $dirdest)
{
	if(is_dir($dirsource))$dir_handle=opendir($dirsource);
	$dirname = substr($dirsource,strrpos($dirsource,"/")+1);

	mkdir($dirdest."/".$dirname, 0750);
	while($file=readdir($dir_handle))
	{
		if($file!="." && $file!="..")
		{
			if(!is_dir($dirsource."/".$file))
			{
				copy ($dirsource."/".$file, $dirdest."/".$dirname."/".$file);
				unlink($dirsource."/".$file);
			}
			else
			{
				$dirdest1 = $dirdest."/".$dirname;
				move_r($dirsource."/".$file, $dirdest1);
			}
		}
	}
	closedir($dir_handle);
	rmdir($dirsource);
}

/**
 * Recursive array merge function
 * 
 * @return array	Merged array
 */
function array_merge_r() 
{
    if (func_num_args() < 2) {
        trigger_error(__FUNCTION__ .' needs two or more array arguments', E_USER_WARNING);
        return;
    }
    $arrays = func_get_args();
    $merged = array();
    while ($arrays) {
        $array = array_shift($arrays);
        if (!is_array($array)) {
            trigger_error(__FUNCTION__ .' encountered a non array argument', E_USER_WARNING);
            return;
        }
        if (!$array) continue;
        foreach ($array as $key => $value)
            if (is_string($key)) {
                if (is_array($value) && array_key_exists($key, $merged) && is_array($merged[$key])) {
                    $merged[$key] = array_merge_r($merged[$key], $value);
                } else {
                    $merged[$key] = $value;
                }
            } else {
                $merged[] = $value;
            }
    }
    return $merged;
}

/**
 * Array all combinations
 * 
 * @param array $array
 * @return array
 */
function array_all_combinations($array) {
    // initialize by adding the empty set
    $results = array(array());

    foreach ($array as $key => $val) {
        foreach ($results as $combination) {
            array_push($results, array_merge(array($key => $val), $combination));
        }
    }

    return $results;
}
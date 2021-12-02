<?php

namespace QExcel\RichText;

/**
 * ITextElement
 *
 * @package    QExcel\RichText
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
interface ITextElement
{
	/**
	 * Get text
	 *
	 * @return string	Text
	 */
	public function getText();

	/**
	 * Set text
	 *
	 * @param 	$pText string	Text
	 * @return ITextElement
	 */
	public function setText($pText = '');

	/**
	 * Get font
	 *
	 * @return PHPExcel_Style_Font
	 */
	public function getFont();

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode();
}

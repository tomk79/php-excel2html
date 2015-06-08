<?php
/**
 * test for tomk79\php-excel2html
 */

class mainTest extends PHPUnit_Framework_TestCase{

	/**
	 * ファイルシステムユーティリティ
	 */
	private $fs;

	/**
	 * setup
	 */
	public function setup(){
		require_once(__DIR__.'/libs/simple_html_dom.php');
		$this->fs = new \tomk79\filesystem();
		mb_internal_encoding('utf-8');
	}

	/**
	 * *.xlsx to *.html
	 */
	public function testXlsx2HtmlConvert(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/default.xlsx'))->get_html();
		// var_dump($src);

		$html = str_get_html(
			$src ,
			true, // $lowercase
			true, // $forceTagsClosed
			DEFAULT_TARGET_CHARSET, // $target_charset
			false, // $stripRN
			DEFAULT_BR_TEXT, // $defaultBRText
			DEFAULT_SPAN_TEXT // $defaultSpanText
		);

		$this->assertEquals( 1, count($html->find('table')) );
		$this->assertEquals( 0, count($html->find('thead')) );
		$this->assertEquals( 18, count($html->find('tr')) );
		$this->assertEquals( 5, count($html->find('tr',0)->find('td')) );
		$this->assertEquals( 'E18', $html->find('tr',17)->childNodes(4)->innertext );
		$this->assertNull( $html->find('tr',18) );
		$this->assertNull( $html->find('tr',17)->childNodes(5) );



		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/default.xlsx'))->get_html(array(
			'header_row'=>2,
			'header_col'=>1,
			'strip_table_tag'=>true
		));
		// var_dump($src);

		$html = str_get_html(
			$src ,
			true, // $lowercase
			true, // $forceTagsClosed
			DEFAULT_TARGET_CHARSET, // $target_charset
			false, // $stripRN
			DEFAULT_BR_TEXT, // $defaultBRText
			DEFAULT_SPAN_TEXT // $defaultSpanText
		);

		$this->assertEquals( 0, count($html->find('table')) );
		$this->assertEquals( 1, count($html->find('thead')) );
		$this->assertEquals( 2, count($html->find('thead tr')) );
		$this->assertEquals( 18, count($html->find('tr')) );
		$this->assertEquals( 5, count($html->find('tr',0)->find('th')) );
		$this->assertEquals( 4, count($html->find('tr',2)->find('td')) );
		$this->assertEquals( 'E18', $html->find('tr',17)->childNodes(4)->innertext );
		$this->assertNull( $html->find('tr',18) );
		$this->assertNull( $html->find('tr',17)->childNodes(5) );

	}//testXlsx2HtmlConvert()

}

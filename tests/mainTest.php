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

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
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

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
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

	/**
	 * Cell Renderer
	 */
	public function testCellRenderer(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_renderer.xlsx'))->get_html(array('cell_renderer'=>'text'));
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'テキスト編集<br />改行コードを含む', $html->find('td',0)->innertext );
		$this->assertEquals( 'テキスト編集<br />&lt;strong style=&quot;font-weight:bold;&quot;&gt;HTMLタグを含む&lt;/strong&gt;', $html->find('td',1)->innertext );
		$this->assertEquals( 'テキスト編集<br /><br />- markdown を含む<br />- markdown の中に、&lt;code&gt;HTML code&lt;/code&gt; を含む。', $html->find('td',2)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_renderer.xlsx'))->get_html(array('cell_renderer'=>'html'));
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'テキスト編集'."\n".'改行コードを含む', $html->find('td',0)->innertext );
		$this->assertEquals( 'テキスト編集'."\n".'<strong style="font-weight:bold;">HTMLタグを含む</strong>', $html->find('td',1)->innertext );
		$this->assertEquals( 'テキスト編集'."\n"."\n".'- markdown を含む'."\n".'- markdown の中に、<code>HTML code</code> を含む。', $html->find('td',2)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_renderer.xlsx'))->get_html(array('cell_renderer'=>'markdown'));
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '<p>テキスト編集'."\n".'改行コードを含む</p>'."\n", $html->find('td',0)->innertext );
		$this->assertEquals( '<p>テキスト編集'."\n".'<strong style="font-weight:bold;">HTMLタグを含む</strong></p>'."\n", $html->find('td',1)->innertext );
		$this->assertEquals( '<p>テキスト編集</p>'."\n"."\n".'<ul>'."\n".'<li>markdown を含む</li>'."\n".'<li>markdown の中に、<code>HTML code</code> を含む。</li>'."\n".'</ul>'."\n", $html->find('td',2)->innertext );


	}//testCellRenderer()

	/**
	 * Merged Cells
	 */
	public function testMergedCells(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/merged_cells.xlsx'))->get_html();
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'B4', $html->find('tr',3)->find('td',1)->innertext );
		$this->assertEquals( '10', $html->find('tr',3)->find('td',1)->getAttribute('colspan') );
		$this->assertEquals( '3', $html->find('tr',3)->find('td',1)->getAttribute('rowspan') );

		$this->assertEquals( 'C16', $html->find('tr',15)->find('td',2)->innertext );
		$this->assertEquals( '2', $html->find('tr',15)->find('td',2)->getAttribute('colspan') );
		$this->assertEquals( '3', $html->find('tr',15)->find('td',2)->getAttribute('rowspan') );

		$this->assertEquals( 'E5', $html->find('tr',4)->find('td',1)->innertext );

		$this->assertEquals( 2, count($html->find('tr',16)->find('td')) );

	}//testMergedCells()

}

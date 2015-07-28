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
		@date_default_timezone_set('Asia/Tokyo');
	}

	/**
	 * *.xlsx to *.html
	 */
	public function testXlsx2HtmlConvert(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/default.xlsx'))->get_html(array(
			'renderer'=>'simplify'
		));
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
			'renderer'=>'simplify',
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
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_renderer.xlsx'))->get_html(array('renderer'=>'simplify','cell_renderer'=>'text'));
		// var_dump($src);
		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'テキスト編集<br />改行コードを含む', $html->find('td',0)->innertext );
		$this->assertEquals( 'テキスト編集<br />&lt;strong style=&quot;font-weight:bold;&quot;&gt;HTMLタグを含む&lt;/strong&gt;', $html->find('td',1)->innertext );
		$this->assertEquals( 'テキスト編集<br /><br />- markdown を含む<br />- markdown の中に、&lt;code&gt;HTML code&lt;/code&gt; を含む。', $html->find('td',2)->innertext );
		$this->assertEquals( '1,234,567,890', $html->find('td',3)->innertext );
		$this->assertEquals( '▲ 1,234,567,890', $html->find('td',4)->innertext );
		$this->assertEquals( '200', $html->find('td',5)->innertext );
		$this->assertEquals( '150', $html->find('td',6)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_renderer.xlsx'))->get_html(array('renderer'=>'simplify','cell_renderer'=>'html'));
		// var_dump($src);
		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'テキスト編集'."\n".'改行コードを含む', $html->find('td',0)->innertext );
		$this->assertEquals( 'テキスト編集'."\n".'<strong style="font-weight:bold;">HTMLタグを含む</strong>', $html->find('td',1)->innertext );
		$this->assertEquals( 'テキスト編集'."\n"."\n".'- markdown を含む'."\n".'- markdown の中に、<code>HTML code</code> を含む。', $html->find('td',2)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_renderer.xlsx'))->get_html(array('renderer'=>'simplify','cell_renderer'=>'markdown'));
		// var_dump($src);
		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '<p>テキスト編集'."\n".'改行コードを含む</p>'."\n", $html->find('td',0)->innertext );
		$this->assertEquals( '<p>テキスト編集'."\n".'<strong style="font-weight:bold;">HTMLタグを含む</strong></p>'."\n", $html->find('td',1)->innertext );
		$this->assertEquals( '<p>テキスト編集</p>'."\n"."\n".'<ul>'."\n".'<li>markdown を含む</li>'."\n".'<li>markdown の中に、<code>HTML code</code> を含む。</li>'."\n".'</ul>'."\n", $html->find('td',2)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/utf-8.csv'))->get_html(array('cell_renderer'=>'markdown'));
		// var_dump($src);
		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '<h2>markdown h2</h2>'."\n", $html->find('tr',6)->find('td',0)->innertext );
		$this->assertEquals( '<h2>html h2</h2>'."\n", $html->find('tr',7)->find('td',0)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/shift_jis.csv'))->get_html(array('cell_renderer'=>'text'));
		// var_dump($src);
		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '## markdown h2', $html->find('tr',6)->find('td',0)->innertext );
		$this->assertEquals( '&lt;h2&gt;html h2&lt;/h2&gt;', $html->find('tr',7)->find('td',0)->innertext );


		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/utf-8.csv'))->get_html(array('cell_renderer'=>'html'));
		// var_dump($src);
		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '## markdown h2', $html->find('tr',6)->find('td',0)->innertext );
		$this->assertEquals( '<h2>html h2</h2>', $html->find('tr',7)->find('td',0)->innertext );


	}//testCellRenderer()

	/**
	 * Merged Cells
	 */
	public function testMergedCells(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/merged_cells.xlsx'))->get_html(array('renderer'=>'simplify'));
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'B4', $html->find('tr',3)->find('td',1)->innertext );
		$this->assertEquals( '10', $html->find('tr',3)->find('td',1)->getAttribute('rowspan') );
		$this->assertEquals( '3', $html->find('tr',3)->find('td',1)->getAttribute('colspan') );

		$this->assertEquals( 'C16', $html->find('tr',15)->find('td',2)->innertext );
		$this->assertEquals( '2', $html->find('tr',15)->find('td',2)->getAttribute('rowspan') );
		$this->assertEquals( '3', $html->find('tr',15)->find('td',2)->getAttribute('colspan') );

		$this->assertEquals( 'E5', $html->find('tr',4)->find('td',1)->innertext );

		$this->assertEquals( 2, count($html->find('tr',16)->find('td')) );

	}//testMergedCells()


	/**
	 * Render by CSV in UTF-8
	 */
	public function testRenderByCSVinUtf8(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/utf-8.csv'))->get_html(array('renderer'=>'simplify'));
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '', $html->find('tr',0)->find('td',0)->innertext );
		$this->assertEquals( 'タイトル行1', $html->find('tr',0)->find('td',1)->innertext );
		$this->assertEquals( 'マルチバイト文字', $html->find('tr',1)->find('td',0)->innertext );
		$this->assertEquals( 'E9', $html->find('tr',8)->find('td',4)->innertext );
	}//testRenderByCSVinUtf8()


	/**
	 * Render by CSV in Shift_JIS
	 */
	public function testRenderByCSVinShift_JIS(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/shift_jis.csv'))->get_html(array('renderer'=>'simplify'));
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( '', $html->find('tr',0)->find('td',0)->innertext );
		$this->assertEquals( 'タイトル行1', $html->find('tr',0)->find('td',1)->innertext );
		$this->assertEquals( 'マルチバイト文字', $html->find('tr',1)->find('td',0)->innertext );
		$this->assertEquals( 'E9', $html->find('tr',8)->find('td',4)->innertext );
	}//testRenderByCSVinShift_JIS()


	/**
	 * cell styles
	 */
	public function testCellStyles(){
		$src = (new \tomk79\excel2html\main(__DIR__.'/sample/cell_styled.xlsx'))->get_html(array());
		// var_dump($src);

		$html = str_get_html( $src, true, true, DEFAULT_TARGET_CHARSET, false, DEFAULT_BR_TEXT, DEFAULT_SPAN_TEXT );
		$this->assertEquals( 'B4', $html->find('tr',3)->find('td',1)->innertext );

	}//testCellStyles()


	/**
	 * options
	 */
	public function testOptions(){
		$e2h = new \tomk79\excel2html\main(__DIR__.'/sample/cell_styled.xlsx');
		$src = $e2h->get_html(array());
		$options = $e2h->get_options();
		// var_dump($src);
		$this->assertEquals( 'strict', $options['renderer'] );


		$e2h = new \tomk79\excel2html\main(__DIR__.'/sample/cell_styled.xlsx');
		$src = $e2h->get_html(array('renderer'=>'simplify'));
		$options = $e2h->get_options();
		// var_dump($src);
		$this->assertEquals( 'simplify', $options['renderer'] );


		$e2h = new \tomk79\excel2html\main(__DIR__.'/sample/utf-8.csv');
		$src = $e2h->get_html(array('renderer'=>'strict'));
		$options = $e2h->get_options();
		// var_dump($src);
		$this->assertEquals( 'simplify', $options['renderer'] );
		$this->assertTrue( $options['render_cell_width'] );
		$this->assertFalse( $options['render_cell_borders'] );


		$e2h = new \tomk79\excel2html\main(__DIR__.'/sample/default.xlsx');
		$src = $e2h->get_html(array(
			'header_row' => 1 ,
			'header_col' => 1 ,
			'renderer' => 'simplify' ,
			'cell_renderer' => 'html' ,
			'render_cell_width' => true ,
			'strip_table_tag' => true
		));
		$options = $e2h->get_options();
		// var_dump($src);
		$this->assertEquals( 'simplify', $options['renderer'] );
		$this->assertTrue( $options['render_cell_width'] );
		$this->assertFalse( $options['render_cell_borders'] );

	}//testOptions()

}

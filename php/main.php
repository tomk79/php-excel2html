<?php
/**
 * excel2html - main.php
 */
namespace tomk79\excel2html;

/**
 * class main
 */
class main{

	/**
	 * filesystem utility
	 */
	private $fs;

	/**
	 * input xlsx filename
	 */
	private $filename;

	/**
	 * input files extension
	 */
	private $ext;

	/**
	 * render options
	 */
	private $options;

	/**
	 * constructor
	 */
	public function __construct( $filename ){
		$this->fs = new \tomk79\filesystem();
		$this->filename = $filename;
		$this->ext = strtolower($this->fs->get_extension( $this->filename ));
	}

	/**
	 * get converted html
	 * @param array $options オプション
	 * <dl>
	 *   <dt>string renderer</dt>
	 *     <dd>レンダリングモード。<code>simplify</code>(単純化)、または<code>strict</code>(そのまま表示) のいずれかを指定します。デフォルトは <code>strict</code> です。 CSVファイルの場合は設定に関わらず強制的に <code>simplify</code> が選択されます。</dd>
	 *   <dt>string cell_renderer</dt>
	 *     <dd>セルのレンダリングモード。<code>html</code>(HTMLコードとして処理)、<code>text</code>(プレーンテキストとして処理)、または<code>markdown</code>(Markdownとして処理) のいずれかを指定します。デフォルトは <code>text</code> です。</dd>
	 * 
	 *   <dt>bool render_cell_width</dt>
	 *     <dd>セル幅を再現する。</dd>
	 *   <dt>bool render_cell_height</dt>
	 *     <dd>セル高を再現する。</dd>
	 *   <dt>bool render_cell_background</dt>
	 *     <dd>セルの背景設定を再現する。</dd>
	 *   <dt>bool render_cell_font</dt>
	 *     <dd>セルの文字設定を再現する。</dd>
	 *   <dt>bool render_cell_borders</dt>
	 *     <dd>セルのボーダーを再現する。</dd>
	 *   <dt>bool render_cell_align</dt>
	 *     <dd>セルの左右位置揃えを再現する。</dd>
	 *   <dt>bool render_cell_vertical_align</dt>
	 *     <dd>セルの上下位置揃えを再現する。</dd>
	 * 
	 *   <dt>int header_row</dt>
	 *     <dd>ヘッダー行の番号。デフォルトは 0。</dd>
	 *   <dt>int header_col</dt>
	 *     <dd>ヘッダー列の番号。デフォルトは 0。</dd>
	 *   <dt>bool strip_table_tag</dt>
	 *     <dd>tableタグを削除するかどうか。true のとき、tableタグは削除した状態で出力されます。デフォルトは false です。</dd>
	 * </dl>
	 */
	public function get_html($options=array()){
		$options = $this->optimize_options( $options );
		$this->options = $options;

		if( $this->ext == 'csv' ){
			require_once( __DIR__.'/render_html_by_csv.php' );
			$renderer = new render_html_by_csv( $this->fs, $this->filename );
			$rtn = $renderer->render($options);
		}else{
			require_once( __DIR__.'/render_html_by_excel.php' );
			$renderer = new render_html_by_excel( $this->fs, $this->filename );
			$rtn = $renderer->render($options);
		}

		return $rtn;
	}// get_html()


	/**
	 * オプション情報を整理する
	 */
	private function optimize_options( $options ){
		$rtn = array();

		$rtn['renderer'] = @$options['renderer'].'';
		if(!strlen($rtn['renderer'])){
			$rtn['renderer'] = 'strict';
		}
		if( $this->ext == 'csv' ){
			// CSVが対象の場合、強制的に 単純化 する。
			$rtn['renderer'] = 'simplify';
		}
		$rtn['cell_renderer'] = @$options['cell_renderer'];
		if(!strlen($rtn['cell_renderer'])){
			$rtn['cell_renderer'] = 'text';
		}
		$rtn['header_row'] = @intval($options['header_row']);
		$rtn['header_col'] = @intval($options['header_col']);
		$rtn['strip_table_tag'] = @(bool) $options['strip_table_tag'];

		// レンダーオプションの初期値を設定
		$rtn['render_cell_width']          = true;
		$rtn['render_cell_height']         = false;
		$rtn['render_cell_background']     = false;
		$rtn['render_cell_font']           = false;
		$rtn['render_cell_borders']        = false;
		$rtn['render_cell_align']          = true;
		$rtn['render_cell_vertical_align'] = false;

		if( $rtn['renderer'] == 'strict' ){
			$rtn['render_cell_width']          = true;
			$rtn['render_cell_height']         = true;
			$rtn['render_cell_background']     = true;
			$rtn['render_cell_font']           = true;
			$rtn['render_cell_borders']        = true;
			$rtn['render_cell_align']          = true;
			$rtn['render_cell_vertical_align'] = true;
		}

		if( !is_null( @$options['render_cell_width'] ) ){
			$rtn['render_cell_width'] = @(bool) $options['render_cell_width'];
		}
		if( !is_null( @$options['render_cell_height'] ) ){
			$rtn['render_cell_height'] = @(bool) $options['render_cell_height'];
		}
		if( !is_null( @$options['render_cell_background'] ) ){
			$rtn['render_cell_background'] = @(bool) $options['render_cell_background'];
		}
		if( !is_null( @$options['render_cell_font'] ) ){
			$rtn['render_cell_font'] = @(bool) $options['render_cell_font'];
		}
		if( !is_null( @$options['render_cell_borders'] ) ){
			$rtn['render_cell_borders'] = @(bool) $options['render_cell_borders'];
		}
		if( !is_null( @$options['render_cell_align'] ) ){
			$rtn['render_cell_align'] = @(bool) $options['render_cell_align'];
		}
		if( !is_null( @$options['render_cell_vertical_align'] ) ){
			$rtn['render_cell_vertical_align'] = @(bool) $options['render_cell_vertical_align'];
		}

		return $rtn;
	}

	/**
	 * 算定されたオプションを取得する
	 */
	public function get_options(){
		return $this->options;
	}

}

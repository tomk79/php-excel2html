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
	 * input xlsx filename
	 */
	private $filename;

	/**
	 * PHPExcel Object
	 */
	private $objPHPExcel;

	/**
	 * constructor
	 */
	public function __construct( $filename ){
		$this->filename = $filename;
		$this->objPHPExcel = \PHPExcel_IOFactory::load( $this->filename );
	}

	/**
	 * get converted html
	 * @param array $options オプション
	 * <dl>
	 *   <dt>string renderer</dt>
	 *     <dd>レンダリングモード。<code>simplify</code>(単純化)、または<code>strict</code>(そのまま表示) のいずれかを指定します。デフォルトは <code>strict</code> です。</dd>
	 *   <dt>string cell_renderer</dt>
	 *     <dd>セルのレンダリングモード。<code>html</code>(HTMLコードとして処理)、<code>text</code>(プレーンテキストとして処理)、または<code>markdown</code>(Markdownとして処理) のいずれかを指定します。デフォルトは <code>text</code> です。</dd>
	 *   <dt>int header_row</dt>
	 *     <dd>ヘッダー行の番号。デフォルトは 0。</dd>
	 *   <dt>int header_col</dt>
	 *     <dd>ヘッダー列の番号。デフォルトは 0。</dd>
	 *   <dt>bool strip_table_tag</dt>
	 *     <dd>tableタグを削除するかどうか。true のとき、tableタグは削除した状態で出力されます。デフォルトは false です。</dd>
	 * </dl>
	 */
	public function get_html($options=array()){
		$options['renderer'] = @$options['renderer'].'';
		$options['cell_renderer'] = @$options['cell_renderer'];
		if(!strlen($options['cell_renderer'])){
			$options['cell_renderer'] = 'text';
		}
		$options['header_row'] = @intval($options['header_row']);
		$options['header_col'] = @intval($options['header_col']);

		$skipCell = array();

		$objWorksheet = $this->objPHPExcel->getActiveSheet();

		$rtn = '';
		if( !@$options['strip_table_tag'] ){
			$rtn .= '<table>'.PHP_EOL;
		}
		$thead = '';
		$tbody = '';
		foreach ($objWorksheet->getRowIterator() as $rowIdx=>$row) {
			// var_dump($rowIdx); //← $rowIdx は1から始まります

			$tmpRow = '';
			$tmpRow .= '<tr>'.PHP_EOL;
			$cellIterator = $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(true);
				// This loops through all cells,
				//    even if a cell value is not set.
				// By default, only cells that have a value 
				//    set will be iterated.
			foreach ($cellIterator as $colIdxName=>$cell) {
				$colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
				// var_dump($colIdx); //← $colIdx は1から始まります
				$cellTagName = 'td';
				if( $rowIdx <= $options['header_row'] || $colIdx <= $options['header_col'] ){
					$cellTagName = 'th';
				}
				$cellValue = $cell->getCalculatedValue();
				switch( $options['cell_renderer'] ){
					case 'text':
						$cellValue = htmlspecialchars($cellValue);
						$cellValue = preg_replace('/\r\n|\r|\n/', '<br />', $cellValue);
						break;
					case 'html':
						break;
					case 'markdown':
						$cellValue = \Michelf\MarkdownExtra::defaultTransform($cellValue);
						break;
				}
				$tmpRow .= '<'.$cellTagName.'>'.$cellValue.'</'.$cellTagName.'>'.PHP_EOL;
			}
			$tmpRow .= '</tr>'.PHP_EOL;

			if( $rowIdx <= $options['header_row'] ){
				$thead .= $tmpRow;
			}else{
				$tbody .= $tmpRow;
			}
		}

		if( strlen($thead) ){
			$rtn .= '<thead>'.PHP_EOL;
			$rtn .= $thead;
			$rtn .= '</thead>'.PHP_EOL;
		}
		$rtn .= '<tbody>'.PHP_EOL;
		$rtn .= $tbody;
		$rtn .= '</tbody>'.PHP_EOL;

		if( !@$options['strip_table_tag'] ){
			$rtn .= '</table>'.PHP_EOL;
		}

		return $rtn;
	}

}

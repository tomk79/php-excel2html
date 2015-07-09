<?php
/**
 * excel2html - render_html_by_csv.php
 */
namespace tomk79\excel2html;

/**
 * class render_html_by_csv
 */
class render_html_by_csv{

	/**
	 * filesystem utility
	 */
	private $fs;

	/**
	 * input xlsx filename
	 */
	private $filename;

	/**
	 * constructor
	 */
	public function __construct( $fs, $filename ){
		$this->fs = $fs;
		$this->filename = $filename;
	}


	/**
	 * CSVファイルからHTMLを描画する
	 */
	public function render( $options ){
		$csv = $this->fs->read_csv($this->filename);
		// var_dump($csv);

		$thead = '';
		$tbody = '';
		foreach ($csv as $rowIdx=>$row) {
			// var_dump($rowIdx); //← $rowIdx は0から始まります

			$tmpRow = '';
			$tmpRow .= '<tr>'.PHP_EOL;

			$html_colgroup = '';
			foreach ($row as $colIdx=>$cell) {
				$rowspan = 1;
				$colspan = 1;

				// var_dump($colIdx); //← $colIdx は0から始まります
				$cellTagName = 'td';
				if( $rowIdx < $options['header_row'] || $colIdx < $options['header_col'] ){
					$cellTagName = 'th';
				}
				$cellValue = $cell;
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

				// セルのスタイルを調べて、CSSを生成
				$styles = array();

				$tmpRow .= '<'.$cellTagName.($rowspan>1?' rowspan="'.$rowspan.'"':'').($colspan>1?' colspan="'.$colspan.'"':'').''.(count($styles)?' style="'.htmlspecialchars(implode(' ',$styles)).'"':'').'>';
				$tmpRow .= $cellValue;
				$tmpRow .= '</'.$cellTagName.'>'.PHP_EOL;
			}
			$tmpRow .= '</tr>'.PHP_EOL;

			if( $rowIdx < $options['header_row'] ){
				$thead .= $tmpRow;
			}else{
				$tbody .= $tmpRow;
			}
		}


		ob_start();
		if( !@$options['strip_table_tag'] ){
			print '<table>'.PHP_EOL;
		}
		if( strlen($html_colgroup) ){
			print '<colgroup>'.PHP_EOL;
			print $html_colgroup;
			print '</colgroup>'.PHP_EOL;
		}
		if( strlen($thead) ){
			print '<thead>'.PHP_EOL;
			print $thead;
			print '</thead>'.PHP_EOL;
		}
		print '<tbody>'.PHP_EOL;
		print $tbody;
		print '</tbody>'.PHP_EOL;

		if( !@$options['strip_table_tag'] ){
			print '</table>'.PHP_EOL;
		}
		$rtn = ob_get_clean();
		return $rtn;
	} // render_html_by_csv()

}

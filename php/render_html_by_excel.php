<?php
/**
 * excel2html - render_html_by_excel.php
 */
namespace tomk79\excel2html;

/**
 * class render_html_by_excel
 */
class render_html_by_excel{

	/**
	 * filesystem utility
	 */
	private $fs;

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
	public function __construct( $fs, $filename ){
		$this->fs = $fs;
		$this->filename = $filename;
		$this->objPHPExcel = \PHPExcel_IOFactory::load( $this->filename );
	}

	/**
	 * Excel ファイルからHTMLを描画する
	 */
	public function render( $options ){
		$skipCell = array();

		$objWorksheet = $this->objPHPExcel->getActiveSheet();

		$mergedCells = $objWorksheet->getMergeCells();
		// var_dump($mergedCells);

		// セル幅を記憶
		$col_widths = array();
		foreach ($objWorksheet->getRowIterator() as $rowIdx=>$row) {
			$cellIterator = $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(true);
			foreach ($cellIterator as $colIdxName=>$cell) {
				$colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
				$col_widths[$colIdx] = intval( $objWorksheet->getColumnDimension($colIdxName)->getWidth() );
			}
			break;
		}
		$col_width_sum = array_sum($col_widths);

		ob_start();
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
			$html_colgroup = '';
			foreach ($cellIterator as $colIdxName=>$cell) {
				$colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
				// var_dump($colIdx);
				$rowspan = 1;
				$colspan = 1;

				if( @$skipCell[$colIdxName.$rowIdx] ){
					continue;
				}
				foreach($mergedCells as $mergedCell){//連結セルの検索
					if( preg_match('/^'.preg_quote($colIdxName.$rowIdx).'\\:([a-zA-Z]+)([0-9]+)$/', $mergedCell, $matched) ){
						$maxIdxC = \PHPExcel_Cell::columnIndexFromString( $matched[1] );
						// var_dump($colIdx);
						// var_dump(\PHPExcel_Cell::stringFromColumnIndex($colIdx-1));
						// var_dump($maxIdxC);
						$maxIdxR = intval($matched[2]);
						for( $idxC=$colIdx; $idxC<=$maxIdxC; $idxC++ ){
							for( $idxR=$rowIdx; $idxR<=$maxIdxR; $idxR++ ){
								$skipCell[\PHPExcel_Cell::stringFromColumnIndex($idxC-1).$idxR] = \PHPExcel_Cell::stringFromColumnIndex($idxC-1).$idxR;
							}
						}
						$colspan = $maxIdxC-$colIdx+1;
						$rowspan = $maxIdxR-$rowIdx+1;
						break;
					}
				}

				// var_dump($colIdx); //← $colIdx は1から始まります
				$cellTagName = 'td';
				if( $rowIdx <= $options['header_row'] || $colIdx <= $options['header_col'] ){
					$cellTagName = 'th';
				}
				$cellValue = $cell->getFormattedValue();
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
				// var_dump( $cell->getNumberFormat() );

				// セルのスタイルを調べて、CSSを生成
				$styles = array();
				$cellStyle = $cell->getStyle();
				// print('<pre>');
				// // $cellStyle->getBorders()->getOutline();
				// var_dump($cellStyle->getBorders()->getLeft()->getColor()->getRGB());
				// print('</pre>');

				if( $options['render_cell_align'] ){
					if( $cellStyle->getAlignment()->getHorizontal() != 'general' ){
						// text-align は、単純化設定でも出力する
						array_push( $styles, 'text-align: '.strtolower($cellStyle->getAlignment()->getHorizontal()).';' );
					}
				}
				if( $options['render_cell_width'] ){
					$html_colgroup .= '<col style="width:'.floatval($col_widths[$colIdx]/$col_width_sum*100).'%;" />'.PHP_EOL;
				}
				if( $options['render_cell_height'] ){
					array_push( $styles, 'height: '.intval($objWorksheet->getRowDimension($rowIdx)->getRowHeight()).'px;' );
				}
				if( $options['render_cell_font'] ){
					array_push( $styles, 'color: #'.strtolower($cellStyle->getFont()->getColor()->getRGB()).';' );
					array_push( $styles, 'font-weight: '.($cellStyle->getFont()->getBold()?'bold':'normal').';' );
					array_push( $styles, 'font-size: '.intval($cellStyle->getFont()->getsize()/12*100).'%;' );
				}
				if( $options['render_cell_background'] ){
					array_push( $styles, 'background-color: #'.strtolower($cellStyle->getFill()->getStartColor()->getRGB()).';' );
				}
				if( $options['render_cell_vertical_align'] ){
					$verticalAlign = strtolower($cellStyle->getAlignment()->getVertical());
					array_push( $styles, 'vertical-align: '.($verticalAlign=='center'?'middle':$verticalAlign).';' );
				}
				if( $options['render_cell_borders'] ){
					array_push( $styles, 'border-top: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getTop()).';' );
					array_push( $styles, 'border-right: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getRight()).';' );
					array_push( $styles, 'border-bottom: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getBottom()).';' );
					array_push( $styles, 'border-left: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getLeft()).';' );
				}


				$tmpRow .= '<'.$cellTagName.($rowspan>1?' rowspan="'.$rowspan.'"':'').($colspan>1?' colspan="'.$colspan.'"':'').''.(count($styles)?' style="'.htmlspecialchars(implode(' ',$styles)).'"':'').'>';
				$tmpRow .= $cellValue;
				// $tmpRow .= $cellStyle->getFill()->getFillType();
				$tmpRow .= '</'.$cellTagName.'>'.PHP_EOL;
			}
			$tmpRow .= '</tr>'.PHP_EOL;

			if( $rowIdx <= $options['header_row'] ){
				$thead .= $tmpRow;
			}else{
				$tbody .= $tmpRow;
			}
		}
		// var_dump($skipCell);

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
	} // render()

	/**
	 * ボーダーオブジェクトからHTMLのborder-style名を得る
	 */
	private function get_borderstyle_by_border( $border ){
		$style = $border->getBorderStyle();
		$border_width = '1px';
		$border_style = 'solid';
		switch( $style ){
			case 'none':
				$border_width = '0';
				$border_style = 'none';
				break;
			case 'dashDot':
				$border_style = 'dashed';
				break;
			case 'dashDotDot':
				$border_style = 'dashed';
				break;
			case 'dashed':
				$border_style = 'dashed';
				break;
			case 'dotted':
				$border_style = 'dotted';
				break;
			case 'double':
				$border_width = '3px';
				$border_style = 'double';
				break;
			case 'hair':
				break;
			case 'medium':
				$border_width = '3px';
				break;
			case 'mediumDashDot':
				$border_width = '3px';
				$border_style = 'dashed';
				break;
			case 'mediumDashDotDot':
				$border_width = '3px';
				$border_style = 'dashed';
				break;
			case 'mediumDashed':
				$border_width = '3px';
				$border_style = 'dashed';
				break;
			case 'slantDashDot':
				$border_width = '3px';
				$border_style = 'solid';
				break;
			case 'thick':
				$border_width = '5px';
				$border_style = 'solid';
				break;
			case 'thin':
				$border_width = '1px';
				$border_style = 'solid';
				break;
		}
		$rtn = $border_width.' '.$border_style.' #'.strtolower($border->getColor()->getRGB()).'';
		return $rtn;
	} // get_borderstyle_by_border()


}

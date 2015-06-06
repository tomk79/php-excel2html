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
	 */
	public function get_html(){

		$objWorksheet = $this->objPHPExcel->getActiveSheet();

		$rtn = '';
		$rtn .= '<table>' . PHP_EOL;
		foreach ($objWorksheet->getRowIterator() as $row) {
			$rtn .= '<tr>' . PHP_EOL;
			$cellIterator = $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(true);
				// This loops through all cells,
				//    even if a cell value is not set.
				// By default, only cells that have a value 
				//    set will be iterated.
			foreach ($cellIterator as $cell) {
				$rtn .= '<td>' . $cell->getValue() . '</td>' . PHP_EOL;
			}
			$rtn .= '</tr>' . PHP_EOL;
		}
		$rtn .= '</table>' . PHP_EOL;

		return $rtn;
	}

}

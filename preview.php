<?php
require_once('./vendor/autoload.php');
?><!DOCTYPE html>
<html>
<head>
<title>php-excel2html</title>
<style>
	table{
		border-collapse: collapse;
	}
</style>
</head>
<body>

<h2>cell_styled.xls</h2>
<?php

$src = (new \tomk79\excel2html\main(__DIR__.'/tests/sample/cell_styled.xls'))->get_html(array('renderer'=>'strict'));
print($src);

?>


<h2>cell_styled.xlsx</h2>
<?php

$src = (new \tomk79\excel2html\main(__DIR__.'/tests/sample/cell_styled.xlsx'))->get_html(array('renderer'=>'strict'));
print($src);

?>

<h2>shift_jis.csv</h2>
<?php

$src = (new \tomk79\excel2html\main(__DIR__.'/tests/sample/shift_jis.csv'))->get_html(array('renderer'=>'simplify'));
print($src);

?>

<h2>utf-8.csv</h2>
<?php

$src = (new \tomk79\excel2html\main(__DIR__.'/tests/sample/utf-8.csv'))->get_html(array('renderer'=>'simplify'));
print($src);

?>

</body>
</html>
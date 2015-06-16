<!DOCTYPE html>
<html>
<head>
<title>php-excel2html</title>
<style>
	table{
		border-collapse: collapse;
		table-layout: fixed;
	}
</style>
</head>
<body>
<?php
require_once('./vendor/autoload.php');

$src = (new \tomk79\excel2html\main(__DIR__.'/tests/sample/cell_styled.xlsx'))->get_html(array('renderer'=>'strict'));
print($src);


?>
</body>
</html>
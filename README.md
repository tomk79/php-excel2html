# php-excel2html


Convert Excel(*.xlsx) to HTML Table. (with PHPExcel).

Excel 形式のファイルを、HTMLの tableタグに変換します。
([PHPExcel](https://github.com/PHPOffice/PHPExcel) を利用しています)

<table>
  <thead>
    <tr>
      <th></th>
      <th>Linux</th>
      <th>Windows</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>master</th>
      <td align="center">
        <a href="https://travis-ci.org/tomk79/php-excel2html"><img src="https://secure.travis-ci.org/tomk79/php-excel2html.svg?branch=master"></a>
      </td>
      <td align="center">
        <a href="https://ci.appveyor.com/project/tomk79/php-excel2html"><img src="https://ci.appveyor.com/api/projects/status/o2d8bo08weasyvlh/branch/master?svg=true"></a>
      </td>
    </tr>
    <tr>
      <th>develop</th>
      <td align="center">
        <a href="https://travis-ci.org/tomk79/php-excel2html"><img src="https://secure.travis-ci.org/tomk79/php-excel2html.svg?branch=develop"></a>
      </td>
      <td align="center">
        <a href="https://ci.appveyor.com/project/tomk79/php-excel2html"><img src="https://ci.appveyor.com/api/projects/status/o2d8bo08weasyvlh/branch/develop?svg=true"></a>
      </td>
    </tr>
  </tbody>
</table>


## Basic Usage - 使い方

```php
<?php
require_once( './vendor/autoload.php' );

$src = (new \tomk79\excel2html\main('path/to/your/excel.xlsx'))->get_html(array(
	'renderer'=>'simplify'
));

print $src;
```

## Options - オプション

<dl>
  <dt>string renderer</dt>
    <dd>レンダリングモード。<code>simplify</code>(単純化)、または<code>strict</code>(そのまま表示) のいずれかを指定します。デフォルトは <code>strict</code> です。 CSVファイルの場合は設定に関わらず強制的に <code>simplify</code> が選択されます。</dd>
  <dt>string cell_renderer</dt>
    <dd>セルのレンダリングモード。<code>html</code>(HTMLコードとして処理)、<code>text</code>(プレーンテキストとして処理)、または<code>markdown</code>(Markdownとして処理) のいずれかを指定します。デフォルトは <code>text</code> です。</dd>

  <dt>bool render_cell_width</dt>
    <dd>セル幅を再現する。</dd>
  <dt>bool render_cell_height</dt>
    <dd>セル高を再現する。</dd>
  <dt>bool render_cell_background</dt>
    <dd>セルの背景設定を再現する。</dd>
  <dt>bool render_cell_font</dt>
    <dd>セルの文字設定を再現する。</dd>
  <dt>bool render_cell_borders</dt>
    <dd>セルのボーダーを再現する。</dd>
  <dt>bool render_cell_align</dt>
    <dd>セルの左右位置揃えを再現する。</dd>
  <dt>bool render_cell_vertical_align</dt>
    <dd>セルの上下位置揃えを再現する。</dd>

  <dt>int header_row</dt>
    <dd>ヘッダー行の番号。デフォルトは 0。</dd>
  <dt>int header_col</dt>
    <dd>ヘッダー列の番号。デフォルトは 0。</dd>
  <dt>bool strip_table_tag</dt>
    <dd>tableタグを削除するかどうか。true のとき、tableタグは削除した状態で出力されます。デフォルトは false です。</dd>
</dl>


## ライセンス - License

MIT License


## 作者 - Author

- (C)Tomoya Koyanagi <tomk79@gmail.com>
- website: <http://www.pxt.jp/>
- Twitter: @tomk79 <http://twitter.com/tomk79/>






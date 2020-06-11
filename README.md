# tomk79/php-excel2html


Convert Excel(\*.xlsx) to HTML Table. (with [PHPExcel](https://github.com/PHPOffice/PHPExcel)).

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

`composer` コマンドを使用してプロジェクトに _tomk79/php-excel2html_ をインストールします。(`composer`について詳しくは[composerのドキュメント](https://getcomposer.org/doc/)をご覧ください)

```bash
$ composer require tomk79/php-excel2html
```

次のコードは実装例です。

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


## 更新履歴 - Change log

### tomk79/php-excel2html v0.0.8 (2020年6月11日)

- PHPExcelの特定の処理で異常終了する場合がある問題を修正。

### tomk79/php-excel2html v0.0.7 (2016年10月18日)

- PHPExcelの特定の処理で異常終了する場合がある問題を修正。

### tomk79/php-excel2html v0.0.6 (2016年10月17日)

- 最後の行が結合されている場合に、列幅指定が欠落する不具合を修正。

### tomk79/php-excel2html v0.0.5 (2016年10月4日)

- michelf/php-markdown を更新

### tomk79/php-excel2html v0.0.4 (2015年7月28日)

- PHP5.4系で起きていた不具合を修正。

### tomk79/php-excel2html v0.0.3 (2015年7月9日)

- CSVを入力した場合の処理を分離・調整、詳細なレンダリングオプションを追加。
- セルの値を、書式設定に従って表示するようになった。


### tomk79/php-excel2html v0.0.2 (2015年6月18日)

- セルの幅を `%` で計算するように修正。
- その他不具合の修正。

### tomk79/php-excel2html v0.0.1 (2015年6月9日)

- Initial Release.


## ライセンス - License

MIT License


## 作者 - Author

- (C)Tomoya Koyanagi <tomk79@gmail.com>
- website: <https://www.pxt.jp/>
- Twitter: @tomk79 <https://twitter.com/tomk79/>

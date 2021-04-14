config_writer_light.py
======================

Excelファイルからjinja2に渡すパラメータを読み込み、テンプレートからテキストファイルを出力する。

概要
----

- Excelファイルからjinja2に渡すパラメータを読み込み、テンプレートからテキストファイルを出力する。
- Excelファイルの1行目がjinja2の変数名となり、2行目以降を1行1ファイルのパラメータとして読み込む
- カレントディレクトリにoutput_yyyymmdd_hhmmssのディレクトリを作成し、ファイルを出力する
- 出力ファイル名は以下の通りとなる。
  - Excelのパラメータに`filename`がある場合、`{{ filename }}.txt`を作成する
  - 重複するファイル名が指定されていた場合は、`{{ filename }}_{{ n }}.txt`（nは2からインクリメント）を作成する
  - Excelのパラメータにfilenameがない場合、`output.txt`を作成する
  - filenameパラメータがあってもセルが空の場合、`None.txt`を作成する

必要なモジュール
----------------

- jinja2
- openpyxl

```
> pip install jinja2
> pip install openpyxl
>
```

使用方法
--------

- `config_writer_light.py` を実行し、パラメータシートとテンプレートを選択して実行する
- コマンドライン引数として、パラメータシートとテンプレートを指定して実行可能

注意事項
--------

- 数式の入ったセルの場合、最終の計算結果が利用される（openpyxlの仕様）
- 変数名が空欄の場合、'None'という変数名として定義される
- 同一変数名がある場合は、最も右側の列のパラメータが有効となる
- 使用したjinja2テンプレートのファイル名は、`template_filename`として渡される
- `filename`もパラメータとして渡される

実行例
------

`パラメータ.xlsx`

| filename | hostname | o1  | o2  | o3 | o4  | ipaddress     | mask | snmp                        | ntp                         |
|----------|----------|-----|-----|----|-----|---------------|------|-----------------------------|-----------------------------|
| Router   | Router   | 192 | 168 | 1  | 1   | 192.168.1.1   | 24   | 10.253.254.35,10.253.254.37 | 10.59.224.254,10.60.224.254 |
| Switch   | Switch   | 192 | 168 | 1  | 250 | 192.168.1.250 | 24   | 10.253.254.35,10.253.254.37 | 10.59.224.254,10.60.224.254 |
| Switch   | Switch   | 192 | 168 | 1  | 249 | 192.168.1.249 | 24   | 10.253.254.35,10.253.254.37 | 10.59.224.254,10.60.224.254 |

※ip address列は、`o1&"."&o2&"."&o3&"."&o4` で文字列結合したセル

`template.j2`

```jinja2
!
! {{ filename }}
! template:{{ template_filename }}
!
conf t
!
{{ hostname }}
!
{{ ipaddress }}/{{ mask }}
!
{% for ip in snmp.split(',') -%}
snmp-server {{ ip }}
{% endfor -%}
!
{% for ip in ntp.split(',') -%}
ntp-server {{ ip }}
{% endfor -%}
!
end
```

`Router.txt`

```txt
!
! Router
! template:test.j2
!
conf t
!
Router
!
192.168.1.1/24
!
snmp-server 10.253.254.35
snmp-server 10.253.254.37
!
ntp-server 10.59.224.254
ntp-server 10.60.224.254
!
end
```

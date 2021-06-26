Config Writer Light
===================

Excelファイルからjinja2に渡すパラメータを読み込み、テンプレートからテキストファイルを出力する。

概要
----

- ExcelファイルからJinjaに渡すパラメータを読み込み、テンプレートからテキストファイルを出力する
- Excelファイルの1行目がJinjaの変数名となり、2行目以降を1行1ファイルのパラメータとして読み込む
  - テンプレートファイル名は、変数`template_filename`に設定される
  - `filename`もパラメータとして渡される
  - Excelのセル設定によらず、すべての変数は文字列として渡される
- カレントディレクトリに`output_yyyymmdd_hhmmss`のディレクトリを作成し、ファイルを出力する
- 出力ファイル名は以下の通りとなる
  - Excelの変数名に`filename`がある場合、`{{ filename }}.txt`を作成する
  - 複数行でファイル名が重複する場合は、`{{ filename }}_n.txt`（nは2からインクリメント）を作成する
  - 変数名に`filename`がない場合または`filename`が空の場合、filenameには`None`が指定されたものと扱う

必要なモジュール
----------------

- jinja2
- openpyxl


使用方法
--------

- `config_writer_light.py` を実行し、パラメータシートとテンプレートを選択して実行する
- コマンドライン引数として、パラメータシートとテンプレートを指定して実行可能

注意事項
--------

- 数式の入ったセルの場合、最終の計算結果が利用される
- 変数名が空欄の場合、'None'という変数名として定義される
- セルが空欄の場合、'None' が設定される
- 同一変数名がある場合は、最も右側の列のパラメータが有効となる

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
{% if hostname == 'Router' -%}
ip routing
{% else -%}
mls qos
{% endif -%}
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
ip routing
!
snmp-server 10.253.254.35
snmp-server 10.253.254.37
!
ntp-server 10.59.224.254
ntp-server 10.60.224.254
!
end
```

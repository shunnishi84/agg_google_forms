# gen_html_form_google_forms.py
## 概要
* GoogleFormの集計結果をhtmlとして生成するスクリプトです。
  * 入力にはエクセルファイルが必要です。

## 使い方
```bash
pip install -r requirements.txt
python gen_html_from_google_forms.py input.xlsx > output.html
```
### グラフ内の日本語フォントが表示されない場合
* static/fonts.ini をOSにあわせて修正してください。
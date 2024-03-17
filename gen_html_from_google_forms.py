import sys
import re
import os
import platform
import math
import matplotlib.pyplot as plt
import matplotlib as mpl
import pandas as pd
import base64
from io import BytesIO

# 同種別の回答が1件であっても `その他` にしない設問をここに指定する
NOT_OTHERS = ['所属部署', '勤続年数']


# グラフの日本語フォントを指定
def get_graph_font(os_name):
    if os_name == 'Linux':
        font_name = 'Noto Sans JP'
    if os_name == 'Windows':
        font_name = 'Meiryo'
    # TODO: Macの場合どうなる？
    # if os_name == 'Darwin':
    #     font_name = ''
    return font_name


# グラフを描画してimgとして返す
def plot_to_base64(labels, sizes):
    plt.figure(figsize=(6, 6))

    plt.pie(sizes, autopct='%1.1f%%', shadow=True, startangle=90)
    plt.legend(loc="center left", bbox_to_anchor=(1, 0.5),
               labels=labels)  # 凡例を表示（かぶりを避ける）

    png_image = BytesIO()
    plt.savefig(png_image, format='png', bbox_inches='tight', dpi=100)
    plt.clf()  # プロットエリアをクリア
    encoded = base64.b64encode(png_image.getvalue()).decode('utf-8')
    img = f'<img src="data:image/png;base64,{encoded}">'
    return img


# html用スタイルシート
def stylesheet():
    css = """
    <style>
        table {
            padding: 20px;
            border-collapse: separate;
            border-spacing: 0;
        }
        li {
            padding-left: 8px;
            padding-top: 2px;
            padding-bottom: 2px;
        }
        th {
            padding: 10px;
        }
        th, td {
            padding: 8px;
            text-align: left;
            border-left: 2px solid #ddd; /* 左側に縦のボーダーを追加 */
        }
        .count {
            text-align: right;
        }
        th:first-child, td:first-child {
            border-left: none; /* 最初のセルの左側のボーダーを削除 */
        }
        thead tr {
            background-color: #333;
            color: white;
        }
        tbody tr:nth-child(odd) {
            background-color: #f2f2f2;
        }
        tbody tr:nth-child(even) {
            background-color: #ddd;
        }
    </style>"""
    return css


# 回答を dict 化する
def convert_from_answer_to_dict(fname):
    df = pd.read_excel(fname)

    ans = {col: df[col].value_counts().to_dict() for col in df.columns}

    for q, ans_dict in ans.items():
        # ループ内で使う一時利用のdict
        tmp = {}
        for a in ans_dict:
            line = str(a).split(", ")
            if (len(line)) == 1:
                continue
            # 複数回答可能な設問は1カラムに " ,"区切りで複数あるため分解して集計
            for v in line:
                if v not in tmp:
                    tmp.update({v: 1})
                else:
                    tmp[v] += 1
            # カンマ区切りのものを集計しなおした形で回答を上書き
            ans[q] = tmp

    return ans


# 回答のdictをhtml化する
def gen_html(data):
    # タグの先頭を出力
    html_base = f"""
<!DOCTYPE html>
<HTML><HEAD>{stylesheet()}</HEAD>"""

    print(html_base)
    q_cnt = 1

    for q, ans in data.items():
        if q == 'タイムスタンプ':
            continue
        # 設問が自由記述型がどうかのフラグ
        only_free_answer = False
        # 回答数が多い順にソートする
        ans = dict(sorted(ans.items(), key=lambda x: x[1], reverse=True))
        # その他回答の配列
        other_ans = []
        # 円グラフプロット用のdict
        plot_ans = {}

        title = f'<h2>Q{q_cnt}. {q}</h2>'
        table_header = "<table><thead><tr><th>回答</th><th>回答数</th>\
            <th>回答割合</th></tr></thead>"
        tables = ""

        # すべての回答が１種類しかないものは自由記述のものとして判断
        if (max(ans.values())) == 1:
            only_free_answer = True

        if not only_free_answer:
            for a, cnt in ans.items():
                # 同一回答が1件より多いものまたは何件あってもその他にいれない設問
                if cnt > 1 or q in NOT_OTHERS:
                    if a not in plot_ans:
                        # 円グラフプロット用の配列に追加
                        plot_ans[a] = cnt
                        tables += f'<tr><td>{a}</td><td class=count>{cnt}</td>\
                        <td class=count>{get_percentage(cnt, ans.values())}\
                        </td></tr>\n'
                else:
                    other_ans.append(a)
                    # 回答数が1個しかないものは円グラフプロット用の配列に`その他`で追加
                    if 'その他' not in plot_ans:
                        plot_ans['その他'] = 1
                    else:
                        plot_ans['その他'] += 1

        # html出力
        print(title)
        if not only_free_answer:
            print(plot_to_base64(plot_ans.keys(), plot_ans.values()))

            print(table_header)
            print(tables)
            print("</table>")

            if len(other_ans) != 0:
                print("<h3>その他には以下のような回答が寄せられています</h3>")
                for a in other_ans:
                    print(format_text(a))
        else:
            for key, value in ans.items():
                print(format_text(key))

        q_cnt += 1


# 回答率
def get_percentage(cnt, nums):
    percentage = f'{math.floor((cnt / sum(nums))* 10000)/100}%'
    return percentage


# フリーアンサー内で `・`とかmarkdownっぽい表記しているのを見やすくする
def format_text(intext):
    response = ""
    for a in intext.split(" ・"):
        a = re.sub(r'^・', '', a)
        a = re.sub(r'\- ', '<br>', re.sub(r'^\-', '', a))
        a = re.sub(r'## ', '<li>', a)
        a = re.sub(r'\<li\>\<li\>', '', a)
        if a != '':
            response += re.sub(r'<li><li>', '<li>', f'<li>{a}</li>\n')

    return response


# メイン処理
def main():
    if len(sys.argv) == 1:
        sys.exit(0)

    # ファイルを読み込む
    fname = sys.argv[1]

    if not os.path.exists(fname):
        print(f'{fname} not found.')
        sys.exit(1)

    # グラフのフォントを変える場合はここを書き換える
    os_name = platform.system()
    mpl.rcParams['font.family'] = get_graph_font(os_name=os_name)

    data = convert_from_answer_to_dict(fname)
    gen_html(data)


if __name__ == '__main__':
    main()

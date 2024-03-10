import sys, re, math
import matplotlib.pyplot as plt
import matplotlib as mpl
import base64
from collections import OrderedDict
from io import BytesIO

## 環境によってここのフォントを変える
mpl.rcParams['font.family'] = 'Noto Sans JP'

# グラフを描画してimgとして返す
def plot_to_base64(labels, sizes):
    plt.figure(figsize=(6,6))
    
    plt.pie(sizes, autopct='%1.1f%%', shadow=True, startangle=90)
    plt.legend(loc="center left", bbox_to_anchor=(1, 0.5), labels=labels)  # 凡例を表示（かぶりを避ける）

    png_image = BytesIO()
    plt.savefig(png_image, format='png', bbox_inches='tight', dpi=100)
    plt.clf()  # プロットエリアをクリア
    encoded = base64.b64encode(png_image.getvalue()).decode('utf-8')
    f = f'<img src="data:image/png;base64,{encoded}">'
    print(f)

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

# 回答のdictをhtml化する
def gen_html(question, data):
    # タグの先頭を出力
    html_base = f"""
<!DOCTYPE html>
<HTML><HEAD>{stylesheet()}</HEAD>"""
    
    print(html_base)
    q_cnt = 1

    for q in question:
        ans = data[q]

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
        table_header ="<table><thead><tr><th>回答</th><th>回答数</th><th>回答割合</th></tr></thead>"
        tables = ""

        # すべての回答が１種類しかないものは自由記述のものとして判断
        if(max(ans.values())) < 3:
            only_free_answer = True

        if not only_free_answer:
            for a,cnt in ans.items():
                if cnt > 1 or q in ['所属部署', '勤続年数']:
                    if a not in plot_ans:
                        # 円グラフプロット用の配列に追加
                        plot_ans[a] = cnt
                    tables += f'<tr><td>{a}</td><td class=count>{cnt}</td><td class=count>{get_percentage(cnt, ans.values())}</td></tr>\n'
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
            plot_to_base64(plot_ans.keys(), plot_ans.values())

            print(table_header)
            print(tables)
            print("</table>")

            if len(other_ans)!=0:
                print("<h3>その他には以下のような回答が寄せられています</h3>")
                for a in other_ans:
                    print(parse_text(a))
        else:
            for key, value in ans.items():
                print(parse_text(key))

        q_cnt += 1


def get_percentage(cnt, nums):
    percentage = f'{math.floor((cnt / sum(nums))* 10000)/100}%'
    return percentage

def parse_text(intext):
    response = ""
    for a in intext.split(" ・"):
        a = re.sub(r'^・', '', a)
        a = re.sub(r'\- ','<br>',re.sub(r'^\-','',a))
        a = re.sub(r'## ','<li>', a)
        a = re.sub(r'\<li\>\<li\>', '', a)
        if a != '':
            response += re.sub(r'<li><li>','<li>',f'<li>{a}</li>\n')

    return response

# メイン処理
if len(sys.argv) == 1:
    sys.exit(0)

# ファイルを読み込む
cont = open(sys.argv[1], 'r').read().split("\n")

# ヘッダ行は質問として取得
question = cont[0].split("\t")

# それ以降を回答と取得
cont = cont[1:]

ans = OrderedDict()

# print(type(ans))
for line in cont:
    for i, val in enumerate(line.split("\t")):
        # 設問とわず回答が空であるものはスキップ
        if val == '':
            continue

        # 複数回答可能なもののために `, `で区切って集計
        for v in val.split(", "):
            # 設問そのものが配列にない場合
            if question[i] not in ans:
                ans[question[i]] = {v : 1}
            else:
                # 設問はあるけど同様の回答がない場合
                if v not in ans[question[i]]:
                    ans[question[i]].update({v :1})
                else:
                    ans[question[i]][v] += 1

gen_html(question, ans)

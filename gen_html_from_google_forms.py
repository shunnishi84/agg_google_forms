"""Google Formsの集計結果をHTMLレポートに変換するスクリプト.

このモジュールは、Google Formsの回答データ（Excelファイル）を読み込み、
グラフ付きのHTMLレポートを生成します。
"""

import sys
import re
import os
import platform
import math
import configparser
from io import BytesIO
from typing import Dict, List, Tuple, Any
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib as mpl
import pandas as pd
import base64
import japanize_matplotlib


# 定数
class Constants:
    """アプリケーション全体で使用する定数."""

    NOT_OTHERS = ['所属部署', '勤続年数']
    CSS_FILE = './static/style.css'
    FONTS_CONFIG = './static/fonts.ini'
    SKIP_COLUMN = 'タイムスタンプ'

    # グラフ設定
    FIGURE_SIZE = (6, 6)
    DPI = 100

    # パーセンテージ計算の精度
    PERCENTAGE_PRECISION = 10000


class ConfigurationError(Exception):
    """設定関連のエラー."""
    pass


class DataProcessingError(Exception):
    """データ処理関連のエラー."""
    pass


class Config:
    """アプリケーションの設定を管理するクラス."""

    def __init__(self, css_file: str = Constants.CSS_FILE,
                 fonts_config: str = Constants.FONTS_CONFIG):
        """
        Configクラスのコンストラクタ.

        Args:
            css_file: CSSファイルのパス
            fonts_config: フォント設定ファイルのパス

        Raises:
            ConfigurationError: 設定ファイルの読み込みに失敗した場合
        """
        self.css_file = css_file
        self.fonts_config = fonts_config
        self._css_content: str = ""
        self._font_family: str = ""

    def load_css(self) -> str:
        """
        CSSファイルを読み込む.

        Returns:
            読み込んだCSS内容をstyleタグで囲んだ文字列

        Raises:
            ConfigurationError: CSSファイルの読み込みに失敗した場合
        """
        try:
            with open(self.css_file, 'r', encoding='utf-8') as f:
                plain_css = f.read()
            self._css_content = f"<style>\n{plain_css}\n</style>"
            return self._css_content
        except FileNotFoundError:
            raise ConfigurationError(
                f"CSSファイルが見つかりません: {self.css_file}"
            )
        except Exception as e:
            raise ConfigurationError(
                f"CSSファイルの読み込みに失敗しました: {e}"
            )

    def get_font_family(self) -> str:
        """
        OSに応じた日本語フォントを取得する.

        Returns:
            フォント名

        Raises:
            ConfigurationError: フォント設定の読み込みに失敗した場合
        """
        if self._font_family:
            return self._font_family

        try:
            config = configparser.ConfigParser()
            config.read(self.fonts_config)
            os_name = platform.system()

            config_dict = {
                section: dict(config.items(section))
                for section in config.sections()
            }

            if os_name not in config_dict:
                raise ConfigurationError(
                    f"OS '{os_name}' のフォント設定が見つかりません"
                )

            self._font_family = config_dict[os_name]['font']
            return self._font_family
        except Exception as e:
            raise ConfigurationError(
                f"フォント設定の読み込みに失敗しました: {e}"
            )

    @property
    def css(self) -> str:
        """CSS内容を取得する（遅延読み込み）."""
        if not self._css_content:
            self.load_css()
        return self._css_content


class ChartGenerator:
    """グラフ生成を担当するクラス."""

    @staticmethod
    def create_pie_chart(labels: List[str], sizes: List[int]) -> str:
        """
        円グラフを生成してbase64エンコードされたimgタグを返す.

        Args:
            labels: グラフのラベルリスト
            sizes: 各ラベルに対応する値のリスト

        Returns:
            base64エンコードされた画像を含むimgタグ
        """
        plt.figure(figsize=Constants.FIGURE_SIZE)
        japanize_matplotlib.japanize()

        plt.pie(sizes, autopct='%1.1f%%', shadow=True, startangle=90)
        plt.legend(loc="center left", bbox_to_anchor=(1, 0.5), labels=labels)

        png_image = BytesIO()
        plt.savefig(png_image, format='png', bbox_inches='tight',
                   dpi=Constants.DPI)
        plt.clf()
        plt.close()

        encoded = base64.b64encode(png_image.getvalue()).decode('utf-8')
        return f'<img src="data:image/png;base64,{encoded}">'


class DataProcessor:
    """データ処理を担当するクラス."""

    @staticmethod
    def load_excel_data(filename: str) -> Dict[str, Dict[str, int]]:
        """
        Excelファイルから回答データを読み込んでdictに変換する.

        Args:
            filename: 読み込むExcelファイルのパス

        Returns:
            質問をキー、回答の辞書を値とする辞書

        Raises:
            DataProcessingError: データの読み込みに失敗した場合
        """
        try:
            df = pd.read_excel(filename)
            return DataProcessor._convert_to_answer_dict(df)
        except FileNotFoundError:
            raise DataProcessingError(f"ファイルが見つかりません: {filename}")
        except Exception as e:
            raise DataProcessingError(
                f"Excelファイルの読み込みに失敗しました: {e}"
            )

    @staticmethod
    def _convert_to_answer_dict(df: pd.DataFrame) -> Dict[str, Dict[str, int]]:
        """
        DataFrameを回答の辞書に変換する.

        複数選択の回答（カンマ区切り）を個別にカウントします。

        Args:
            df: Pandas DataFrame

        Returns:
            質問をキー、回答の辞書を値とする辞書
        """
        ans = {col: df[col].value_counts().to_dict() for col in df.columns}

        # 複数選択の回答を展開
        for question, answer_dict in ans.items():
            tmp = {}
            for answer in answer_dict:
                line = str(answer).split(", ")
                if len(line) == 1:
                    continue
                for value in line:
                    tmp[value] = tmp.get(value, 0) + 1
                ans[question] = tmp

        return ans

    @staticmethod
    def calculate_percentage(count: int, total_counts: List[int]) -> str:
        """
        回答率を計算する.

        Args:
            count: 個別の回答数
            total_counts: 全回答数のリスト

        Returns:
            パーセンテージ文字列（例: "75.5%"）
        """
        total = sum(total_counts)
        if total == 0:
            return "0%"
        percentage = math.floor(
            (count / total) * Constants.PERCENTAGE_PRECISION
        ) / 100
        return f'{percentage}%'


class TextFormatter:
    """テキストフォーマット処理を担当するクラス."""

    @staticmethod
    def format_free_answer(text: str) -> str:
        """
        フリーアンサーをHTML形式にフォーマットする.

        Args:
            text: フォーマットするテキスト

        Returns:
            HTMLフォーマットされたテキスト
        """
        response = ""
        for answer in text.split(" ・"):
            answer = re.sub(r'^・', '', answer)
            answer = re.sub(r'\- ', '<br>', re.sub(r'^\-', '', answer))
            answer = re.sub(r'## ', '<li>', answer)
            answer = re.sub(r'\<li\>\<li\>', '', answer)
            if answer:
                response += f'<li>{answer}</li>\n'
        return response


class HTMLGenerator:
    """HTML生成を担当するクラス."""

    def __init__(self, config: Config):
        """
        HTMLGeneratorクラスのコンストラクタ.

        Args:
            config: Configインスタンス
        """
        self.config = config
        self.chart_generator = ChartGenerator()
        self.text_formatter = TextFormatter()

    def generate_html(self, data: Dict[str, Dict[str, int]]) -> str:
        """
        回答データからHTMLレポートを生成する.

        Args:
            data: 質問と回答の辞書

        Returns:
            生成されたHTML文字列
        """
        html_parts = [self._generate_header()]

        question_number = 1
        for question, answers in data.items():
            if question == Constants.SKIP_COLUMN:
                continue

            html_parts.append(
                self._generate_question_section(question, answers, question_number)
            )
            question_number += 1

        html_parts.append(self._generate_footer())
        return "\n".join(html_parts)

    def _generate_header(self) -> str:
        """HTML文書のヘッダーを生成する."""
        return f"""<!DOCTYPE html>
<HTML><HEAD>{self.config.css}</HEAD><BODY>"""

    def _generate_footer(self) -> str:
        """HTML文書のフッターを生成する."""
        return "</BODY></HTML>"

    def _generate_question_section(
        self,
        question: str,
        answers: Dict[str, int],
        question_number: int
    ) -> str:
        """
        個別の質問セクションを生成する.

        Args:
            question: 質問文
            answers: 回答の辞書
            question_number: 質問番号

        Returns:
            質問セクションのHTML文字列
        """
        is_free_answer = (max(answers.values()) == 1)

        if is_free_answer:
            return self._generate_free_answer_section(
                question, answers, question_number
            )
        else:
            return self._generate_choice_answer_section(
                question, answers, question_number
            )

    def _generate_free_answer_section(
        self,
        question: str,
        answers: Dict[str, int],
        question_number: int
    ) -> str:
        """
        フリーアンサー形式の質問セクションを生成する.

        Args:
            question: 質問文
            answers: 回答の辞書
            question_number: 質問番号

        Returns:
            フリーアンサーセクションのHTML文字列
        """
        parts = [f'<h2>Q{question_number}. {question}</h2>']
        for answer_key in answers.keys():
            parts.append(self.text_formatter.format_free_answer(answer_key))
        return "\n".join(parts)

    def _generate_choice_answer_section(
        self,
        question: str,
        answers: Dict[str, int],
        question_number: int
    ) -> str:
        """
        選択式の質問セクションを生成する.

        Args:
            question: 質問文
            answers: 回答の辞書
            question_number: 質問番号

        Returns:
            選択式セクションのHTML文字列
        """
        plot_data, table_rows, other_answers = self._process_choice_answers(
            question, answers
        )

        parts = [f'<h2>Q{question_number}. {question}</h2>']

        # グラフの追加
        if plot_data:
            chart = self.chart_generator.create_pie_chart(
                list(plot_data.keys()),
                list(plot_data.values())
            )
            parts.append(chart)

        # テーブルの追加
        if table_rows:
            table = self._generate_answer_table(table_rows)
            parts.append(table)

        # その他の回答の追加
        if other_answers:
            parts.append('<h3>その他には以下のような回答が寄せられています</h3>')
            for answer in other_answers:
                parts.append(self.text_formatter.format_free_answer(answer))

        return "\n".join(parts)

    def _process_choice_answers(
        self,
        question: str,
        answers: Dict[str, int]
    ) -> Tuple[Dict[str, int], List[str], List[str]]:
        """
        選択式回答を処理してグラフデータ、テーブル行、その他回答に分類する.

        Args:
            question: 質問文
            answers: 回答の辞書

        Returns:
            (グラフ用データ, テーブル行リスト, その他回答リスト)のタプル
        """
        plot_data = {}
        table_rows = []
        other_answers = []

        for answer, count in answers.items():
            if count > 1 or question in Constants.NOT_OTHERS:
                plot_data[answer] = count
                percentage = DataProcessor.calculate_percentage(
                    count, list(answers.values())
                )
                table_rows.append(
                    f'<tr><td>{answer}</td><td class=count>{count}</td>'
                    f'<td class=count>{percentage}</td></tr>'
                )
            else:
                other_answers.append(answer)
                plot_data['その他'] = plot_data.get('その他', 0) + 1

        return plot_data, table_rows, other_answers

    def _generate_answer_table(self, table_rows: List[str]) -> str:
        """
        回答テーブルを生成する.

        Args:
            table_rows: テーブル行のリスト

        Returns:
            テーブルのHTML文字列
        """
        header = """<table><thead><tr><th>回答</th>
                    <th>回答数</th><th>回答割合</th></tr></thead>"""
        return header + "\n".join(table_rows) + "</table>"


class Application:
    """アプリケーションのメインクラス."""

    def __init__(self):
        """Applicationクラスのコンストラクタ."""
        self.config = Config()
        self.data_processor = DataProcessor()
        self.html_generator = HTMLGenerator(self.config)

    def run(self, filename: str) -> str:
        """
        アプリケーションを実行する.

        Args:
            filename: 入力Excelファイルのパス

        Returns:
            生成されたHTML文字列

        Raises:
            ConfigurationError: 設定エラーの場合
            DataProcessingError: データ処理エラーの場合
        """
        # フォント設定
        font_family = self.config.get_font_family()
        mpl.rcParams['font.family'] = font_family

        # データ読み込みと処理
        data = self.data_processor.load_excel_data(filename)

        # HTML生成
        return self.html_generator.generate_html(data)


def main():
    """メイン関数."""
    if len(sys.argv) != 2:
        print("Usage: python gen_html_from_google_forms.py <filename>")
        sys.exit(1)

    filename = sys.argv[1]
    if not os.path.exists(filename):
        print(f'{filename} not found.')
        sys.exit(1)

    try:
        app = Application()
        html_output = app.run(filename)
        print(html_output)
    except (ConfigurationError, DataProcessingError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()

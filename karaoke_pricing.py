#!/usr/bin/env python3
"""カラオケ料金計算プログラム"""

import sys
import json
import re
from collections import deque


# 時間帯の境界 (17:50:00)
NIGHT_TIME_SECONDS = 17 * 3600 + 50 * 60  # 64200

# 30分ブロック(秒)
BLOCK_SECONDS = 30 * 60  # 1800
# 猶予時間(秒)
GRACE_SECONDS = 10 * 60  # 600

# 料金テーブル (plan -> (daytime_rate, nighttime_rate))
RATE_TABLE = {
    "one_drink": (100, 400),
    "free_refills": (200, 500),
    "alcohol_free_refills": (300, 650),
}


def time_to_seconds(time_str):
    """hh:mi:ss形式の時刻を秒に変換する"""
    match = re.fullmatch(r"(\d{2}):(\d{2}):(\d{2})", time_str)
    if not match:
        return None
    h, m, s = int(match.group(1)), int(match.group(2)), int(match.group(3))
    if h > 23 or m > 59 or s > 59:
        return None
    return h * 3600 + m * 60 + s


def is_nighttime(seconds):
    """ナイトタイムかどうか判定する"""
    return seconds >= NIGHT_TIME_SECONDS


def calculate_room_charge(enter_sec, leave_sec, plan):
    """一人あたりの室料を計算する"""
    total = 0
    charge_time = enter_sec  # 最初のチャージ時刻

    # 最初の30分は必ず発生
    day_rate, night_rate = RATE_TABLE[plan]
    rate = night_rate if is_nighttime(charge_time) else day_rate
    total += rate

    # 以降のブロック
    # 最初の30分が経過したあと10分経過すると次の30分の料金が発生
    next_threshold = enter_sec + BLOCK_SECONDS + GRACE_SECONDS  # 40分後
    while leave_sec > next_threshold:
        charge_time = next_threshold
        rate = night_rate if is_nighttime(charge_time) else day_rate
        total += rate
        next_threshold += BLOCK_SECONDS + GRACE_SECONDS  # さらに40分後

    return total


def output_result(code, total_amount=None):
    """結果をJSON出力する"""
    result = {"code": code}
    if total_amount is not None:
        result["total_amount"] = total_amount
    print(json.dumps(result, ensure_ascii=False))


def main():
    lines = sys.stdin.read().strip().split("\n")
    if not lines:
        output_result(999)
        return

    records = []
    for line in lines:
        parts = line.strip().split(" ")
        if len(parts) < 2:
            output_result(999)
            return
        records.append(parts)

    # ヘッダレコード解析
    if len(records) < 2:
        output_result(999)
        return

    header = records[0]
    if len(header) != 3 or header[1] != "header":
        output_result(999)
        return

    header_time = time_to_seconds(header[0])
    if header_time is None:
        output_result(999)
        return

    plan = header[2]
    if plan not in RATE_TABLE:
        output_result(999)
        return

    # フッタレコード確認
    footer = records[-1]
    if len(footer) != 2 or footer[1] != "footer":
        output_result(999)
        return

    footer_time = time_to_seconds(footer[0])
    if footer_time is None:
        output_result(999)
        return

    # 中間レコード解析
    enter_queue = deque()  # (enter_time_sec, count)
    total_entered = 0
    total_drink_count = 0
    total_amount = 0
    current_occupancy = 0
    prev_time = header_time

    for record in records[1:-1]:
        rec_time = time_to_seconds(record[0])
        if rec_time is None:
            output_result(999)
            return

        # 時刻が前のレコードより前でないことを確認
        if rec_time < prev_time:
            output_result(999)
            return
        prev_time = rec_time

        rec_type = record[1]

        if rec_type == "enter":
            if len(record) != 3:
                output_result(999)
                return
            try:
                n = int(record[2])
            except ValueError:
                output_result(999)
                return
            if n < 1 or n > 999:
                output_result(99)
                return
            enter_queue.append((rec_time, n))
            total_entered += n
            current_occupancy += n

        elif rec_type == "leave":
            if len(record) != 3:
                output_result(999)
                return
            try:
                n = int(record[2])
            except ValueError:
                output_result(999)
                return
            if n < 1 or n > 999:
                output_result(99)
                return
            if n > current_occupancy:
                output_result(99)
                return
            current_occupancy -= n

            # FIFO で退室処理し、室料を計算
            remaining = n
            while remaining > 0 and enter_queue:
                enter_time, count = enter_queue[0]
                if count <= remaining:
                    enter_queue.popleft()
                    charge = calculate_room_charge(enter_time, rec_time, plan)
                    total_amount += charge * count
                    remaining -= count
                else:
                    enter_queue[0] = (enter_time, count - remaining)
                    charge = calculate_room_charge(enter_time, rec_time, plan)
                    total_amount += charge * remaining
                    remaining = 0

        elif rec_type == "drink":
            if len(record) != 5:
                output_result(999)
                return
            drink_type = record[2]
            if drink_type not in ("alcohol", "soft_drink"):
                output_result(999)
                return
            try:
                unit_price = int(record[3])
                quantity = int(record[4])
            except ValueError:
                output_result(999)
                return
            if unit_price < 1 or unit_price > 9999:
                output_result(99)
                return
            if quantity < 1 or quantity > 99:
                output_result(99)
                return
            total_drink_count += quantity

            # 飲み放題プランの場合、対象ドリンクは無料
            if plan == "free_refills" and drink_type == "soft_drink":
                pass  # ソフトドリンク飲み放題: ソフトドリンクは無料
            elif plan == "alcohol_free_refills":
                pass  # アルコール飲み放題: 全ドリンク無料
            else:
                total_amount += unit_price * quantity

        elif rec_type == "food":
            if len(record) != 4:
                output_result(999)
                return
            try:
                unit_price = int(record[2])
                quantity = int(record[3])
            except ValueError:
                output_result(999)
                return
            if unit_price < 1 or unit_price > 9999:
                output_result(99)
                return
            if quantity < 1 or quantity > 99:
                output_result(99)
                return
            total_amount += unit_price * quantity

        else:
            output_result(999)
            return

    # フッタ時刻チェック
    if footer_time < prev_time:
        output_result(999)
        return

    # 全員退室確認: まだ残っている人がいればフッタ時刻で退室処理
    while enter_queue:
        enter_time, count = enter_queue.popleft()
        charge = calculate_room_charge(enter_time, footer_time, plan)
        total_amount += charge * count
        current_occupancy -= count

    if current_occupancy != 0:
        output_result(99)
        return

    # ワンドリンクチェック
    if plan == "one_drink" and total_drink_count < total_entered:
        output_result(1)
        return

    output_result(0, total_amount)


if __name__ == "__main__":
    main()

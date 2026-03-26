#!/usr/bin/env python3
"""カラオケ料金計算プログラム

SOLID原則に基づいた設計:
- S: 各クラスが単一の責務を持つ (Parser, Validator, PricingPlan, Calculator, Service)
- O: PricingPlan を継承して新プランを追加可能
- L: 全PricingPlanサブクラスが基底クラスと置換可能
- I: 各クラスが必要最小限のインターフェースを公開
- D: KaraokeService は具象クラスではなく抽象に依存 (DI)
"""

import json
import re
import sys
from abc import ABC, abstractmethod
from collections import deque
from dataclasses import dataclass
from enum import Enum


# ============================================================
# エラー定義
# ============================================================

class KaraokeError(Exception):
    def __init__(self, code: int, message: str = ""):
        self.code = code
        self.message = message
        super().__init__(message)


class InvalidInputError(KaraokeError):
    """不正な入力エラー (code: 999)"""
    def __init__(self, message: str = ""):
        super().__init__(code=999, message=message)


class TerminalOperationError(KaraokeError):
    """端末操作エラー (code: 99)"""
    def __init__(self, message: str = ""):
        super().__init__(code=99, message=message)


class OneDrinkError(KaraokeError):
    """ワンドリンクエラー (code: 1)"""
    def __init__(self, drink_count: int, message: str = ""):
        self.drink_count = drink_count
        super().__init__(code=1, message=message)


# ============================================================
# データモデル
# ============================================================

class PlanType(Enum):
    ONE_DRINK = "one_drink"
    FREE_REFILLS = "free_refills"
    ALCOHOL_FREE_REFILLS = "alcohol_free_refills"


class DrinkType(Enum):
    ALCOHOL = "alcohol"
    SOFT_DRINK = "soft_drink"


@dataclass(frozen=True)
class HeaderRecord:
    time_seconds: int
    plan: PlanType


@dataclass(frozen=True)
class EnterRecord:
    time_seconds: int
    count: int


@dataclass(frozen=True)
class LeaveRecord:
    time_seconds: int
    count: int


@dataclass(frozen=True)
class DrinkRecord:
    time_seconds: int
    drink_type: DrinkType
    unit_price: int
    quantity: int


@dataclass(frozen=True)
class FoodRecord:
    time_seconds: int
    unit_price: int
    quantity: int


@dataclass(frozen=True)
class FooterRecord:
    time_seconds: int


@dataclass(frozen=True)
class CalculationResult:
    code: int
    price: int | None = None
    drink: int | None = None

    def to_dict(self) -> dict:
        result = {"code": self.code}
        if self.price is not None:
            result["price"] = self.price
        if self.drink is not None:
            result["drink"] = self.drink
        return result


# ============================================================
# パーサー (Single Responsibility: 入力テキスト → レコード変換)
# ============================================================

_TIME_PATTERN = re.compile(r"^(\d{2}):(\d{2}):(\d{2})$")
_PLAN_MAP = {p.value: p for p in PlanType}
_DRINK_TYPE_MAP = {d.value: d for d in DrinkType}


class RecordParser:
    """入力テキストをレコードオブジェクトに変換する"""

    def parse_time(self, time_str: str) -> int:
        match = _TIME_PATTERN.match(time_str)
        if not match:
            raise InvalidInputError(f"invalid time format: {time_str}")
        h, m, s = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if h > 23 or m > 59 or s > 59:
            raise InvalidInputError(f"invalid time value: {time_str}")
        return h * 3600 + m * 60 + s

    def parse_line(self, line: str):
        parts = line.strip().split(" ")
        if len(parts) < 2:
            raise InvalidInputError("too few fields")

        time_sec = self.parse_time(parts[0])
        rec_type = parts[1]

        parsers = {
            "header": self._parse_header,
            "enter": self._parse_enter,
            "leave": self._parse_leave,
            "drink": self._parse_drink,
            "food": self._parse_food,
            "footer": self._parse_footer,
        }
        if rec_type not in parsers:
            raise InvalidInputError(f"unknown record type: {rec_type}")
        return parsers[rec_type](time_sec, parts)

    def parse(self, text: str) -> list:
        lines = text.strip().split("\n")
        if not lines or lines == [""]:
            raise InvalidInputError("empty input")
        return [self.parse_line(line) for line in lines]

    def _parse_header(self, time_sec: int, parts: list) -> HeaderRecord:
        if len(parts) != 3:
            raise InvalidInputError("header record requires 3 fields")
        if parts[2] not in _PLAN_MAP:
            raise InvalidInputError(f"unknown plan: {parts[2]}")
        return HeaderRecord(time_sec, _PLAN_MAP[parts[2]])

    def _parse_enter(self, time_sec: int, parts: list) -> EnterRecord:
        if len(parts) != 3:
            raise InvalidInputError("enter record requires 3 fields")
        try:
            return EnterRecord(time_sec, int(parts[2]))
        except ValueError:
            raise InvalidInputError(f"invalid count: {parts[2]}")

    def _parse_leave(self, time_sec: int, parts: list) -> LeaveRecord:
        if len(parts) != 3:
            raise InvalidInputError("leave record requires 3 fields")
        try:
            return LeaveRecord(time_sec, int(parts[2]))
        except ValueError:
            raise InvalidInputError(f"invalid count: {parts[2]}")

    def _parse_drink(self, time_sec: int, parts: list) -> DrinkRecord:
        if len(parts) != 5:
            raise InvalidInputError("drink record requires 5 fields")
        if parts[2] not in _DRINK_TYPE_MAP:
            raise InvalidInputError(f"unknown drink type: {parts[2]}")
        try:
            return DrinkRecord(time_sec, _DRINK_TYPE_MAP[parts[2]], int(parts[3]), int(parts[4]))
        except ValueError:
            raise InvalidInputError("invalid drink price or quantity")

    def _parse_food(self, time_sec: int, parts: list) -> FoodRecord:
        if len(parts) != 4:
            raise InvalidInputError("food record requires 4 fields")
        try:
            return FoodRecord(time_sec, int(parts[2]), int(parts[3]))
        except ValueError:
            raise InvalidInputError("invalid food price or quantity")

    def _parse_footer(self, time_sec: int, parts: list) -> FooterRecord:
        if len(parts) != 2:
            raise InvalidInputError("footer record requires 2 fields")
        return FooterRecord(time_sec)


# ============================================================
# バリデーター (Single Responsibility: レコード列の妥当性検証)
# ============================================================

class RecordValidator:
    """レコード列の構造・値を検証する"""

    def validate(self, records: list) -> None:
        self._validate_structure(records)
        self._validate_time_order(records)
        self._validate_values(records)

    def _validate_structure(self, records: list) -> None:
        if len(records) < 2:
            raise InvalidInputError("records must have header and footer")
        if not isinstance(records[0], HeaderRecord):
            raise InvalidInputError("first record must be header")
        if not isinstance(records[-1], FooterRecord):
            raise InvalidInputError("last record must be footer")

    def _validate_time_order(self, records: list) -> None:
        prev_time = records[0].time_seconds
        for record in records[1:]:
            if record.time_seconds < prev_time:
                raise InvalidInputError("time must not go backward")
            prev_time = record.time_seconds

    def _validate_values(self, records: list) -> None:
        current_occupancy = 0
        for record in records:
            if isinstance(record, EnterRecord):
                if record.count < 1 or record.count > 999:
                    raise TerminalOperationError(f"enter count out of range: {record.count}")
                current_occupancy += record.count
            elif isinstance(record, LeaveRecord):
                if record.count < 1 or record.count > 999:
                    raise TerminalOperationError(f"leave count out of range: {record.count}")
                if record.count > current_occupancy:
                    raise TerminalOperationError("leave count exceeds occupancy")
                current_occupancy -= record.count
            elif isinstance(record, DrinkRecord):
                if record.unit_price < 1 or record.unit_price > 9999:
                    raise TerminalOperationError(f"drink price out of range: {record.unit_price}")
                if record.quantity < 1 or record.quantity > 99:
                    raise TerminalOperationError(f"drink quantity out of range: {record.quantity}")
            elif isinstance(record, FoodRecord):
                if record.unit_price < 1 or record.unit_price > 9999:
                    raise TerminalOperationError(f"food price out of range: {record.unit_price}")
                if record.quantity < 1 or record.quantity > 99:
                    raise TerminalOperationError(f"food quantity out of range: {record.quantity}")


# ============================================================
# 料金プラン戦略 (Open/Closed, Liskov Substitution)
# ============================================================

NIGHT_TIME_SECONDS = 17 * 3600 + 50 * 60  # 17:50:00
BLOCK_SECONDS = 30 * 60
GRACE_SECONDS = 10 * 60


class PricingPlan(ABC):
    """料金プランの抽象基底クラス"""

    @abstractmethod
    def get_room_rate(self, time_seconds: int) -> int:
        """指定時刻の30分あたり室料を返す"""

    @abstractmethod
    def calculate_drink_charge(self, record: DrinkRecord) -> int:
        """ドリンク1レコード分の料金を返す"""

    @abstractmethod
    def validate_drink_order(self, total_entered: int, total_drink_count: int) -> None:
        """ドリンク注文数の妥当性を検証する"""

    def _is_nighttime(self, time_seconds: int) -> bool:
        return time_seconds >= NIGHT_TIME_SECONDS


class OneDrinkPlan(PricingPlan):
    DAYTIME_RATE = 100
    NIGHTTIME_RATE = 400

    def get_room_rate(self, time_seconds: int) -> int:
        return self.NIGHTTIME_RATE if self._is_nighttime(time_seconds) else self.DAYTIME_RATE

    def calculate_drink_charge(self, record: DrinkRecord) -> int:
        return record.unit_price * record.quantity

    def validate_drink_order(self, total_entered: int, total_drink_count: int) -> None:
        if total_drink_count < total_entered:
            raise OneDrinkError(
                drink_count=total_drink_count,
                message="drink count is less than entered count",
            )


class FreeRefillsPlan(PricingPlan):
    DAYTIME_RATE = 200
    NIGHTTIME_RATE = 500

    def get_room_rate(self, time_seconds: int) -> int:
        return self.NIGHTTIME_RATE if self._is_nighttime(time_seconds) else self.DAYTIME_RATE

    def calculate_drink_charge(self, record: DrinkRecord) -> int:
        if record.drink_type == DrinkType.SOFT_DRINK:
            return 0
        return record.unit_price * record.quantity

    def validate_drink_order(self, total_entered: int, total_drink_count: int) -> None:
        pass


class AlcoholFreeRefillsPlan(PricingPlan):
    DAYTIME_RATE = 300
    NIGHTTIME_RATE = 650

    def get_room_rate(self, time_seconds: int) -> int:
        return self.NIGHTTIME_RATE if self._is_nighttime(time_seconds) else self.DAYTIME_RATE

    def calculate_drink_charge(self, record: DrinkRecord) -> int:
        return 0

    def validate_drink_order(self, total_entered: int, total_drink_count: int) -> None:
        pass


def create_pricing_plan(plan_type: PlanType) -> PricingPlan:
    """プランタイプに対応する PricingPlan を生成するファクトリ"""
    plans = {
        PlanType.ONE_DRINK: OneDrinkPlan,
        PlanType.FREE_REFILLS: FreeRefillsPlan,
        PlanType.ALCOHOL_FREE_REFILLS: AlcoholFreeRefillsPlan,
    }
    return plans[plan_type]()


# ============================================================
# 室料計算 (Single Responsibility: 室料の計算ロジック)
# ============================================================

class RoomChargeCalculator:
    """入退室時刻から室料を計算する"""

    def __init__(self, pricing_plan: PricingPlan):
        self._plan = pricing_plan

    def calculate_per_person(self, enter_sec: int, leave_sec: int) -> int:
        total = 0
        charge_time = enter_sec
        total += self._plan.get_room_rate(charge_time)

        next_threshold = enter_sec + BLOCK_SECONDS + GRACE_SECONDS
        while leave_sec > next_threshold:
            charge_time = next_threshold
            total += self._plan.get_room_rate(charge_time)
            next_threshold += BLOCK_SECONDS + GRACE_SECONDS

        return total


# ============================================================
# サービス (依存性注入で各コンポーネントを組み合わせる)
# ============================================================

class KaraokeService:
    """カラオケ料金計算のオーケストレーター"""

    def __init__(
        self,
        parser: RecordParser,
        validator: RecordValidator,
        plan_factory=create_pricing_plan,
    ):
        self._parser = parser
        self._validator = validator
        self._plan_factory = plan_factory

    def calculate(self, input_text: str) -> CalculationResult:
        try:
            return self._do_calculate(input_text)
        except OneDrinkError as e:
            return CalculationResult(code=e.code, drink=e.drink_count)
        except KaraokeError as e:
            return CalculationResult(code=e.code)

    def _do_calculate(self, input_text: str) -> CalculationResult:
        records = self._parser.parse(input_text)
        self._validator.validate(records)

        header = records[0]
        plan = self._plan_factory(header.plan)
        room_calc = RoomChargeCalculator(plan)

        enter_queue: deque[list] = deque()
        total_entered = 0
        total_drink_count = 0
        total_price = 0

        for record in records[1:-1]:
            if isinstance(record, EnterRecord):
                enter_queue.append([record.time_seconds, record.count])
                total_entered += record.count

            elif isinstance(record, LeaveRecord):
                remaining = record.count
                while remaining > 0 and enter_queue:
                    entry = enter_queue[0]
                    enter_time, count = entry[0], entry[1]
                    if count <= remaining:
                        enter_queue.popleft()
                        total_price += room_calc.calculate_per_person(enter_time, record.time_seconds) * count
                        remaining -= count
                    else:
                        entry[1] -= remaining
                        total_price += room_calc.calculate_per_person(enter_time, record.time_seconds) * remaining
                        remaining = 0

            elif isinstance(record, DrinkRecord):
                total_drink_count += record.quantity
                total_price += plan.calculate_drink_charge(record)

            elif isinstance(record, FoodRecord):
                total_price += record.unit_price * record.quantity

        footer = records[-1]
        while enter_queue:
            entry = enter_queue.popleft()
            enter_time, count = entry[0], entry[1]
            total_price += room_calc.calculate_per_person(enter_time, footer.time_seconds) * count

        plan.validate_drink_order(total_entered, total_drink_count)

        return CalculationResult(code=0, price=total_price)


# ============================================================
# エントリーポイント
# ============================================================

def main():
    service = KaraokeService(
        parser=RecordParser(),
        validator=RecordValidator(),
    )
    input_text = sys.stdin.read()
    result = service.calculate(input_text)
    print(json.dumps(result.to_dict(), ensure_ascii=False))


if __name__ == "__main__":
    main()

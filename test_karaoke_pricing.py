#!/usr/bin/env python3
"""カラオケ料金計算プログラムのテスト (TDD)"""

import pytest
from karaoke_pricing import (
    # エラー
    InvalidInputError, TerminalOperationError, OneDrinkError,
    # モデル
    PlanType, DrinkType,
    HeaderRecord, EnterRecord, LeaveRecord,
    DrinkRecord, FoodRecord, FooterRecord,
    CalculationResult,
    # クラス
    RecordParser, RecordValidator,
    PricingPlan, OneDrinkPlan, FreeRefillsPlan, AlcoholFreeRefillsPlan,
    RoomChargeCalculator, KaraokeService,
    create_pricing_plan,
    # 定数
    NIGHT_TIME_SECONDS,
)


# ============================================================
# RecordParser テスト
# ============================================================

class TestParseTime:
    def test_midnight(self):
        assert RecordParser().parse_time("00:00:00") == 0

    def test_normal_time(self):
        assert RecordParser().parse_time("10:30:15") == 10 * 3600 + 30 * 60 + 15

    def test_max_time(self):
        assert RecordParser().parse_time("23:59:59") == 23 * 3600 + 59 * 60 + 59

    def test_invalid_format_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_time("1:00:00")

    def test_invalid_hour_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_time("25:00:00")

    def test_invalid_minute_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_time("10:60:00")

    def test_invalid_second_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_time("10:00:60")


class TestParseHeader:
    def test_one_drink(self):
        r = RecordParser().parse_line("10:00:00 header one_drink")
        assert r == HeaderRecord(36000, PlanType.ONE_DRINK)

    def test_free_refills(self):
        r = RecordParser().parse_line("10:00:00 header free_refills")
        assert r == HeaderRecord(36000, PlanType.FREE_REFILLS)

    def test_alcohol_free_refills(self):
        r = RecordParser().parse_line("10:00:00 header alcohol_free_refills")
        assert r == HeaderRecord(36000, PlanType.ALCOHOL_FREE_REFILLS)

    def test_invalid_plan_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 header unknown_plan")

    def test_missing_plan_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 header")


class TestParseEnter:
    def test_normal(self):
        r = RecordParser().parse_line("10:00:00 enter 3")
        assert r == EnterRecord(36000, 3)

    def test_missing_count_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 enter")

    def test_non_integer_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 enter abc")


class TestParseLeave:
    def test_normal(self):
        r = RecordParser().parse_line("11:00:00 leave 2")
        assert r == LeaveRecord(39600, 2)


class TestParseDrink:
    def test_alcohol(self):
        r = RecordParser().parse_line("10:00:00 drink alcohol 500 3")
        assert r == DrinkRecord(36000, DrinkType.ALCOHOL, 500, 3)

    def test_soft_drink(self):
        r = RecordParser().parse_line("10:00:00 drink soft_drink 300 2")
        assert r == DrinkRecord(36000, DrinkType.SOFT_DRINK, 300, 2)

    def test_invalid_type_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 drink beer 500 1")

    def test_missing_fields_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 drink alcohol 500")


class TestParseFood:
    def test_normal(self):
        r = RecordParser().parse_line("10:00:00 food 800 2")
        assert r == FoodRecord(36000, 800, 2)


class TestParseFooter:
    def test_normal(self):
        r = RecordParser().parse_line("11:00:00 footer")
        assert r == FooterRecord(39600)


class TestParseMultipleLines:
    def test_parse_all(self):
        text = "10:00:00 header one_drink\n10:00:00 enter 3\n10:30:00 footer"
        records = RecordParser().parse(text)
        assert len(records) == 3
        assert isinstance(records[0], HeaderRecord)
        assert isinstance(records[1], EnterRecord)
        assert isinstance(records[2], FooterRecord)

    def test_empty_input_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse("")

    def test_unknown_record_type_raises(self):
        with pytest.raises(InvalidInputError):
            RecordParser().parse_line("10:00:00 unknown 3")


# ============================================================
# RecordValidator テスト
# ============================================================

@pytest.fixture
def validator():
    return RecordValidator()


class TestValidatorStructure:
    def test_valid_structure(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 3),
            LeaveRecord(37800, 3),
            FooterRecord(37800),
        ]
        validator.validate(records)

    def test_empty_records_raises(self, validator):
        with pytest.raises(InvalidInputError):
            validator.validate([])

    def test_no_header_raises(self, validator):
        with pytest.raises(InvalidInputError):
            validator.validate([EnterRecord(36000, 3), FooterRecord(37800)])

    def test_no_footer_raises(self, validator):
        with pytest.raises(InvalidInputError):
            validator.validate([HeaderRecord(36000, PlanType.ONE_DRINK), EnterRecord(36000, 3)])

    def test_header_and_footer_only(self, validator):
        validator.validate([HeaderRecord(36000, PlanType.ONE_DRINK), FooterRecord(37800)])


class TestValidatorTimeOrder:
    def test_backward_time_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 3),
            LeaveRecord(35000, 3),
            FooterRecord(37800),
        ]
        with pytest.raises(InvalidInputError):
            validator.validate(records)

    def test_same_time_allowed(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 3),
            LeaveRecord(36000, 3),
            FooterRecord(36000),
        ]
        validator.validate(records)


class TestValidatorEnterCount:
    def test_zero_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 0),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_over_999_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1000),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_999_is_valid(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 999),
            FooterRecord(37800),
        ]
        validator.validate(records)


class TestValidatorLeaveCount:
    def test_leave_exceeds_occupancy_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 2),
            LeaveRecord(37800, 3),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_leave_zero_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 2),
            LeaveRecord(37800, 0),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)


class TestValidatorDrink:
    def test_price_zero_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1),
            DrinkRecord(36000, DrinkType.SOFT_DRINK, 0, 1),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_price_over_9999_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1),
            DrinkRecord(36000, DrinkType.SOFT_DRINK, 10000, 1),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_quantity_zero_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1),
            DrinkRecord(36000, DrinkType.SOFT_DRINK, 300, 0),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_quantity_over_99_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1),
            DrinkRecord(36000, DrinkType.SOFT_DRINK, 300, 100),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)


class TestValidatorFood:
    def test_price_zero_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1),
            FoodRecord(36000, 0, 1),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)

    def test_quantity_over_99_raises(self, validator):
        records = [
            HeaderRecord(36000, PlanType.ONE_DRINK),
            EnterRecord(36000, 1),
            FoodRecord(36000, 500, 100),
            FooterRecord(37800),
        ]
        with pytest.raises(TerminalOperationError):
            validator.validate(records)


# ============================================================
# PricingPlan テスト
# ============================================================

class TestOneDrinkPlan:
    def test_daytime_rate(self):
        plan = OneDrinkPlan()
        assert plan.get_room_rate(36000) == 100  # 10:00:00

    def test_nighttime_rate(self):
        plan = OneDrinkPlan()
        assert plan.get_room_rate(NIGHT_TIME_SECONDS) == 400  # 17:50:00

    def test_boundary_daytime(self):
        plan = OneDrinkPlan()
        assert plan.get_room_rate(NIGHT_TIME_SECONDS - 1) == 100  # 17:49:59

    def test_drink_charge(self):
        plan = OneDrinkPlan()
        record = DrinkRecord(36000, DrinkType.SOFT_DRINK, 300, 2)
        assert plan.calculate_drink_charge(record) == 600

    def test_validate_drink_order_ok(self):
        OneDrinkPlan().validate_drink_order(3, 3)

    def test_validate_drink_order_fail(self):
        with pytest.raises(OneDrinkError) as exc_info:
            OneDrinkPlan().validate_drink_order(3, 2)
        assert exc_info.value.drink_count == 2


class TestFreeRefillsPlan:
    def test_daytime_rate(self):
        assert FreeRefillsPlan().get_room_rate(36000) == 200

    def test_nighttime_rate(self):
        assert FreeRefillsPlan().get_room_rate(NIGHT_TIME_SECONDS) == 500

    def test_soft_drink_free(self):
        plan = FreeRefillsPlan()
        record = DrinkRecord(36000, DrinkType.SOFT_DRINK, 300, 5)
        assert plan.calculate_drink_charge(record) == 0

    def test_alcohol_charged(self):
        plan = FreeRefillsPlan()
        record = DrinkRecord(36000, DrinkType.ALCOHOL, 500, 2)
        assert plan.calculate_drink_charge(record) == 1000

    def test_validate_drink_order_always_passes(self):
        FreeRefillsPlan().validate_drink_order(3, 0)


class TestAlcoholFreeRefillsPlan:
    def test_daytime_rate(self):
        assert AlcoholFreeRefillsPlan().get_room_rate(36000) == 300

    def test_nighttime_rate(self):
        assert AlcoholFreeRefillsPlan().get_room_rate(NIGHT_TIME_SECONDS) == 650

    def test_all_drinks_free(self):
        plan = AlcoholFreeRefillsPlan()
        assert plan.calculate_drink_charge(DrinkRecord(36000, DrinkType.ALCOHOL, 500, 3)) == 0
        assert plan.calculate_drink_charge(DrinkRecord(36000, DrinkType.SOFT_DRINK, 300, 2)) == 0

    def test_validate_drink_order_always_passes(self):
        AlcoholFreeRefillsPlan().validate_drink_order(3, 0)


class TestCreatePricingPlan:
    def test_one_drink(self):
        assert isinstance(create_pricing_plan(PlanType.ONE_DRINK), OneDrinkPlan)

    def test_free_refills(self):
        assert isinstance(create_pricing_plan(PlanType.FREE_REFILLS), FreeRefillsPlan)

    def test_alcohol_free_refills(self):
        assert isinstance(create_pricing_plan(PlanType.ALCOHOL_FREE_REFILLS), AlcoholFreeRefillsPlan)


# ============================================================
# RoomChargeCalculator テスト
# ============================================================

class TestRoomChargeCalculator:
    def test_single_block_daytime(self):
        """30分以内 → 1ブロック"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        # 10:00 〜 10:30 (30分)
        assert calc.calculate_per_person(36000, 36000 + 1800) == 100

    def test_within_grace_period(self):
        """30分+10分以内(猶予内) → 1ブロック"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        # 10:00 〜 10:40 (40分ちょうど、猶予内)
        assert calc.calculate_per_person(36000, 36000 + 2400) == 100

    def test_exceeds_grace_period(self):
        """30分+10分超過 → 2ブロック"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        # 10:00 〜 10:40:01 (40分1秒)
        assert calc.calculate_per_person(36000, 36000 + 2401) == 200

    def test_three_blocks(self):
        """80分超過 → 3ブロック"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        # 10:00 〜 11:20:01 (80分1秒)
        assert calc.calculate_per_person(36000, 36000 + 4801) == 300

    def test_daytime_to_nighttime_transition(self):
        """デイタイムからナイトタイムにまたがるケース"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        # 17:40:00 入室 → 18:30:00 退室
        enter = 17 * 3600 + 40 * 60  # 17:40:00
        leave = 18 * 3600 + 30 * 60  # 18:30:00
        # ブロック1: 17:40 (daytime) → 100
        # ブロック2: 18:20 (nighttime) → 400
        assert calc.calculate_per_person(enter, leave) == 500

    def test_zero_duration(self):
        """入室と退室が同時 → 1ブロック"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        assert calc.calculate_per_person(36000, 36000) == 100

    def test_nighttime_only(self):
        """ナイトタイムのみ"""
        calc = RoomChargeCalculator(OneDrinkPlan())
        enter = 18 * 3600  # 18:00
        leave = 18 * 3600 + 50 * 60  # 18:50
        # ブロック1: 18:00 (night) → 400
        # ブロック2: 18:40 (night) → 400
        assert calc.calculate_per_person(enter, leave) == 800

    def test_free_refills_daytime(self):
        calc = RoomChargeCalculator(FreeRefillsPlan())
        assert calc.calculate_per_person(36000, 36000 + 1800) == 200

    def test_alcohol_free_refills_nighttime(self):
        calc = RoomChargeCalculator(AlcoholFreeRefillsPlan())
        assert calc.calculate_per_person(NIGHT_TIME_SECONDS, NIGHT_TIME_SECONDS + 1800) == 650


# ============================================================
# KaraokeService 統合テスト
# ============================================================

def make_service():
    return KaraokeService(
        parser=RecordParser(),
        validator=RecordValidator(),
    )


class TestServiceNormal:
    def test_simple_one_drink(self):
        """ワンドリンク、デイタイム、30分以内"""
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 3",
            "10:00:00 drink soft_drink 300 3",
            "10:00:00 food 500 2",
            "10:30:00 leave 3",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        # 室料: 3 × 100 = 300
        # ドリンク: 3 × 300 = 900
        # フード: 2 × 500 = 1000
        assert result == CalculationResult(code=0, price=2200)

    def test_free_refills_with_alcohol(self):
        """ソフトドリンク飲み放題、アルコール有料"""
        text = "\n".join([
            "17:00:00 header free_refills",
            "17:00:00 enter 2",
            "17:00:00 drink soft_drink 300 2",
            "17:30:00 drink alcohol 500 1",
            "18:00:00 leave 2",
            "18:00:00 footer",
        ])
        result = make_service().calculate(text)
        # 室料: 2 × (200 + 200) = 800 (17:00 daytime, 17:40 daytime)
        # ドリンク: soft_drink無料 + alcohol 500 = 500
        assert result == CalculationResult(code=0, price=1300)

    def test_alcohol_free_refills(self):
        """アルコール飲み放題、全ドリンク無料"""
        text = "\n".join([
            "10:00:00 header alcohol_free_refills",
            "10:00:00 enter 2",
            "10:00:00 drink alcohol 500 3",
            "10:00:00 drink soft_drink 200 2",
            "10:00:00 food 800 1",
            "10:05:00 enter 1",
            "10:30:00 leave 1",
            "11:00:00 leave 2",
            "11:00:00 footer",
        ])
        result = make_service().calculate(text)
        # FIFO退室:
        # 10:30 leave 1 → 10:00入室組から1名退室 (30分, 1ブロック: 300)
        # 11:00 leave 2 → 10:00入室組から1名 (60分, 2ブロック: 600)
        #                → 10:05入室組から1名 (55分, 2ブロック: 600)
        # 室料合計: 300 + 600 + 600 = 1500
        # フード: 800
        assert result == CalculationResult(code=0, price=2300)

    def test_daytime_to_nighttime_transition(self):
        """デイタイム→ナイトタイム跨ぎ"""
        text = "\n".join([
            "17:40:00 header one_drink",
            "17:40:00 enter 2",
            "17:40:00 drink soft_drink 300 2",
            "18:30:00 leave 2",
            "18:30:00 footer",
        ])
        result = make_service().calculate(text)
        # 室料: 2 × (100 + 400) = 1000 (17:40 day, 18:20 night)
        # ドリンク: 2 × 300 = 600
        assert result == CalculationResult(code=0, price=1600)

    def test_leave_at_footer(self):
        """フッタ時刻で自動退室処理"""
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 2",
            "10:00:00 drink soft_drink 300 2",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        # leaveなし → フッタ10:30で退室処理
        # 室料: 2 × 100 = 200
        # ドリンク: 2 × 300 = 600
        assert result == CalculationResult(code=0, price=800)

    def test_multiple_enter_leave(self):
        """複数回の入退室"""
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 2",
            "10:00:00 drink soft_drink 300 2",
            "10:20:00 leave 1",
            "10:30:00 enter 1",
            "10:30:00 drink soft_drink 300 1",
            "11:00:00 leave 2",
            "11:00:00 footer",
        ])
        result = make_service().calculate(text)
        # FIFO退室:
        # 10:20 leave 1 → 10:00組から1名 (20分, 1ブロック: 100)
        # 11:00 leave 2 → 10:00組から1名 (60分, 2ブロック: 200)
        #               → 10:30組から1名 (30分, 1ブロック: 100)
        # 室料: 100 + 200 + 100 = 400
        # ドリンク: 3 × 300 = 900
        assert result == CalculationResult(code=0, price=1300)


class TestServiceOneDrinkError:
    def test_drink_count_less_than_entered(self):
        """ワンドリンクエラー: ドリンク数 < 入室人数"""
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 3",
            "10:00:00 drink soft_drink 300 1",
            "10:30:00 leave 3",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result == CalculationResult(code=1, drink=1)

    def test_drink_count_zero(self):
        """ワンドリンクエラー: ドリンク注文なし"""
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 2",
            "10:30:00 leave 2",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result == CalculationResult(code=1, drink=0)

    def test_drink_count_equal_to_entered_is_ok(self):
        """ドリンク数 == 入室人数 → 正常"""
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 3",
            "10:00:00 drink soft_drink 300 3",
            "10:30:00 leave 3",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 0


class TestServiceInvalidInput:
    def test_empty_input(self):
        result = make_service().calculate("")
        assert result.code == 999

    def test_invalid_time_format(self):
        text = "1:00:00 header one_drink\n10:00:00 footer"
        result = make_service().calculate(text)
        assert result.code == 999

    def test_unknown_plan(self):
        text = "10:00:00 header unknown\n10:00:00 footer"
        result = make_service().calculate(text)
        assert result.code == 999

    def test_unknown_record_type(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 sing 3",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 999

    def test_time_goes_backward(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:30:00 enter 3",
            "10:00:00 leave 3",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 999

    def test_no_footer(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 3",
        ])
        result = make_service().calculate(text)
        assert result.code == 999


class TestServiceTerminalOperationError:
    def test_enter_count_over_999(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 1000",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 99

    def test_leave_exceeds_occupancy(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 2",
            "10:30:00 leave 3",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 99

    def test_drink_price_out_of_range(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 1",
            "10:00:00 drink soft_drink 10000 1",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 99

    def test_drink_quantity_over_99(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 1",
            "10:00:00 drink soft_drink 300 100",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 99

    def test_food_price_zero(self):
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 1",
            "10:00:00 food 0 1",
            "10:30:00 footer",
        ])
        result = make_service().calculate(text)
        assert result.code == 99


class TestCalculationResult:
    def test_success_to_dict(self):
        r = CalculationResult(code=0, price=1500)
        assert r.to_dict() == {"code": 0, "price": 1500}

    def test_one_drink_error_to_dict(self):
        r = CalculationResult(code=1, drink=2)
        assert r.to_dict() == {"code": 1, "drink": 2}

    def test_error_to_dict(self):
        r = CalculationResult(code=999)
        assert r.to_dict() == {"code": 999}


class TestServiceDependencyInjection:
    """依存性注入が正しく動作することを確認"""

    def test_custom_plan_factory(self):
        """カスタムのplan_factoryを注入できる"""
        class FixedRatePlan(PricingPlan):
            def get_room_rate(self, time_seconds: int) -> int:
                return 999

            def calculate_drink_charge(self, record: DrinkRecord) -> int:
                return 0

            def validate_drink_order(self, total_entered: int, total_drink_count: int) -> None:
                pass

        service = KaraokeService(
            parser=RecordParser(),
            validator=RecordValidator(),
            plan_factory=lambda _: FixedRatePlan(),
        )
        text = "\n".join([
            "10:00:00 header one_drink",
            "10:00:00 enter 1",
            "10:30:00 leave 1",
            "10:30:00 footer",
        ])
        result = service.calculate(text)
        assert result == CalculationResult(code=0, price=999)

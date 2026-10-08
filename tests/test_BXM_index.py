from datetime import date

import pandas as pd
import pandas_market_calendars as mcal
import pytest

from BXM_index import (
    CALENDAR_NAME,
    INPUT_FILE,
    calculate_midpoint,
    calculate_non_roll_return,
    calculate_roll_return,
    compute_index,
    compute_returns,
    export_excel,
    get_roll_date,
    load_inputs,
    next_month,
    populate_option_details,
    third_friday,
    validate_inputs,
)


def prepared_bxm_data():
    """
    Loading the original workbook and populating the active
    option strikes and expiration dates.
    """

    calendar = mcal.get_calendar(CALENDAR_NAME)
    df = load_inputs(INPUT_FILE)
    df = populate_option_details(df,calendar)
    return df, calendar

# ============================================================
# Financial Calculations
# ============================================================

def test_option_midpoint():
    result = calculate_midpoint(167.3,174.6)
    assert result == pytest.approx(170.95)


def test_august_24_ordinary_return():
    df, _ = prepared_bxm_data()
    previous = df.iloc[1]
    current = df.iloc[2]

    result = calculate_non_roll_return(
        current_spx=current["spx_close"],
        previous_spx=previous["spx_close"],
        current_call=current["mid"],
        previous_call=previous["mid"],
        dividend=current["div"]
    )

    assert result == pytest.approx(-0.0007164617399844087,abs=1e-10)


def test_august_21_roll_return():
    df, _ = prepared_bxm_data()
    previous = df.iloc[0]
    current = df.iloc[1]

    result = calculate_roll_return(
        previous_spx=previous["spx_close"],
        previous_call=previous["mid"],
        soq=current["soq"],
        dividend=current["div"],
        old_strike=previous["k"],
        spx_vwap=current["spx_vwap"],
        call_vwap=current["vwap"],
        current_spx=current["spx_close"],
        current_call=current["mid"]
    )

    assert result["settlement"] == pytest.approx(207.93017578125,abs=1e-6)
    assert result["ret"] == pytest.approx(0.0003765841655831714,abs=1e-10)


# ============================================================
# Option expiration and strike tracking
# ============================================================

def test_third_friday():
    assert third_friday(2026, 8) == date(2026, 8, 21)
    assert third_friday(2026, 9) == date(2026, 9, 18)


def test_holiday_adjusted_roll_date():
    calendar = mcal.get_calendar(CALENDAR_NAME)
    assert get_roll_date(2026, 6, calendar) == date(2026, 6, 18)


def test_december_to_january():
    assert next_month(2026, 12) == (2027, 1)


def test_option_strike_and_expiration_tracking():
    df, _ = prepared_bxm_data()

    assert df.iloc[0]["k"] == 7480.0
    assert df.iloc[1]["k"] == 7680.0
    assert df.iloc[-1]["k"] == 7680.0
    assert pd.Timestamp(df.iloc[-1]["expr_date"]).date() == date(2026, 9, 18)


# ============================================================
# Complete BXM replication
# ============================================================

def test_final_bxm_index_level():
    df, calendar = prepared_bxm_data()
    df = validate_inputs(df, calendar)
    df = compute_returns(df, calendar)
    df = compute_index(df, initial_level=100.0)

    assert len(df) == 8
    assert df["Index close"].iloc[0] == pytest.approx(100.0)
    assert df["Index close"].iloc[-1] == pytest.approx(
        100.487620873207, abs=1e-9
    )


# ============================================================
# Invalid Inputs
# ============================================================

def test_missing_trading_session():
    df, calendar = prepared_bxm_data()
    bad_df = df.drop(index=2).reset_index(drop=True)

    with pytest.raises(ValueError, match="Missing exchange trading dates"):
        validate_inputs(bad_df, calendar)


def test_incorrect_option_settlement():
    df, calendar = prepared_bxm_data()

    # Changing the August 21 settlement to check if validation rejects it.
    settlement_values = df["settle"].to_numpy(copy=True)
    settlement_values[1] = float(settlement_values[1]) + 1.0
    df["settle"] = settlement_values

    with pytest.raises(ValueError, match="Incorrect option settlement"):
        validate_inputs(df, calendar)


def test_bid_greater_than_ask():
    df, calendar = prepared_bxm_data()

    # Making the August 20 bid larger than the ask.
    bid_values = df["bid"].to_numpy(copy=True)
    ask_values = df["ask"].to_numpy()
    bid_values[0] = float(ask_values[0]) + 1.0
    df["bid"] = bid_values

    with pytest.raises(ValueError, match="Bid exceeds ask"):
        validate_inputs(df,calendar)


def test_incorrect_midpoint():
    df, calendar = prepared_bxm_data()

    # Changing the August 20 midpoint without changing its bid or ask.
    midpoint_values = df["mid"].to_numpy(copy=True)
    midpoint_values[0] = float(midpoint_values[0]) + 0.01
    df["mid"] = midpoint_values

    with pytest.raises(ValueError, match="Incorrect option midpoint"):
        validate_inputs(df, calendar)


# ============================================================
# Excel Export
# ============================================================

def test_bxm_excel_export(tmp_path):
    df, calendar = prepared_bxm_data()
    df = validate_inputs(df, calendar)
    df = compute_returns(df, calendar)
    df = compute_index(df, initial_level=100.0)

    test_output = tmp_path / "test_bxm_output.xlsx"
    export_excel(df,test_output)

    assert test_output.exists()

    exported_df = pd.read_excel(test_output, sheet_name="BXM")
    assert exported_df.shape == (8, 18)
    assert exported_df["Index close"].iloc[-1] == pytest.approx(
        100.487620873207,
        abs=1e-9
    )

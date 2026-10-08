import pytest

from market_holidays_calendars import get_market_holidays

# ============================================================
# Cboe holidays and early closes
# ============================================================


def test_cboe_juneteenth():
    result = get_market_holidays(
        exchange="CBOE_Index_Options",
        start_date="2026-06-15",
        end_date="2026-06-22",
    )

    assert "2026-06-19" in result["holidays"]


def test_cboe_christmas_eve_early_close():
    # The installed Cboe calendar needs the published December 24, 2026 early-close correction.
    result = get_market_holidays(
        exchange="CBOE_Index_Options",
        start_date="2026-12-21",
        end_date="2026-12-28",
    )

    assert "2026-12-24" in result["early_closes"]
    assert "2026-12-24" not in result["holidays"]
    assert "2026-12-25" in result["holidays"]

    without_early_closes = get_market_holidays(
        exchange="CBOE_Index_Options",
        start_date="2026-12-21",
        end_date="2026-12-28",
        include_early_closes=False,
    )

    assert without_early_closes["early_closes"] == []


# ============================================================
# NYSE holidays and early closes
# ============================================================


def test_nyse_thanksgiving():
    result = get_market_holidays(
        exchange="XNYS",
        start_date="2026-11-23",
        end_date="2026-11-30",
    )

    assert "2026-11-26" in result["holidays"]
    assert "2026-11-27" in result["early_closes"]

    # An early close is still a trading session, not a full-day holiday.
    assert "2026-11-27" not in result["holidays"]


# ============================================================
# Invalid inputs
# ============================================================


def test_invalid_date_range():
    with pytest.raises(ValueError, match="Start date cannot be after end date"):
        get_market_holidays(
            exchange="XNYS",
            start_date="2026-11-30",
            end_date="2026-11-23",
        )


def test_unsupported_exchange():
    with pytest.raises(ValueError, match="Unsupported exchange"):
        get_market_holidays(
            exchange="INVALID_EXCHANGE",
            start_date="2026-01-01",
            end_date="2026-12-31",
        )

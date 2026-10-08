import math

from datetime import date, timedelta
from pathlib import Path

from openpyxl.styles import Alignment, Font, PatternFill

import pandas as pd
import pandas_market_calendars as mcal

PROJECT_DIR = Path(__file__).resolve().parent

INPUT_FILE = PROJECT_DIR / "data" / "BXM.xlsx"

CALENDAR_NAME = "CBOE_Index_Options"

# =====================================================================
# 1. Read and prepare the input workbook
# =====================================================================

def load_inputs(file_path: Path) -> pd.DataFrame:
    """
    Reading and cleaning the original Cboe BXM data file.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"BXM data file not found: {file_path}")

    # Reading the Excel data file.
    df = pd.read_excel(
        file_path, 
        sheet_name=0,
        na_values=["NA"],
        engine="openpyxl"
    )

    # Columns containing market-data inputs.
    market_columns = [
        "spx_10am",
        "soq",
        "settle",
        "k",
        "vwap",
        "spx_vwap",
        "spx_close",
        "div",
        "bid",
        "ask",
        "mid"
    ]

    # Verifying that all required columns are present.
    required_columns = [
        "date", "expr_date", *market_columns,
        "retA", "retB", "retC", "ret", "Index close",
    ]

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing required columns in BXM data file: {missing_columns}")

    # Keeping rows containing a date or any market-data input.
    data_rows = (
        df["date"].notna() |
        df[market_columns].notna().any(axis=1)
    )

    df = df.loc[data_rows].copy()

    # Checking for rows with market data but no valid date.
    parsed_dates = pd.to_datetime(
        df["date"],
        errors="coerce")

    if parsed_dates.isna().any():

        invalid_rows = (
            df.index[parsed_dates.isna()] + 2
        ).tolist()

        raise ValueError(
            f"Missing or invalid trading dates"
            f" in BXM data file at rows: {invalid_rows}")

    df["date"] = parsed_dates

    # Converting the option expiration column to datetime.
    df["expr_date"] = pd.to_datetime(
        df["expr_date"],
        errors="raise")

    # Ensuring market-data inputs are numerics.
    for col in market_columns:
        df[col] = pd.to_numeric(
            df[col],
            errors="raise"
        )

    # Preparing the calculated columns for numerical output.
    calculated_columns = [
        "retA",
        "retB",
        "retC",
        "ret",
        "Index close"
    ]

    for col in calculated_columns:
        df[col] = pd.to_numeric(
            df[col],
            errors="raise"
        )

    # Sorting observations in date order.
    df = df.sort_values("date")

    # Resetting the DataFrame index.
    df = df.reset_index(drop=True)

    return df

# =====================================================================
# 2. Calculate daily portfolio returns
# =====================================================================

def calculate_midpoint(bid: float, ask: float) -> float:
    """
    Calculate the midpoint of the option's bid and ask prices.
    """
    mid = (bid + ask) / 2.0
    return mid

def calculate_non_roll_return(
    current_spx: float,
    previous_spx: float,
    current_call: float,
    previous_call: float,
    dividend: float
) -> float:
    """
    Calculating the BXM return on a non-roll trading days.
    """

    previous_portfolio = previous_spx - previous_call

    current_portfolio = (
        current_spx + dividend - current_call
    )

    if previous_portfolio <= 0:
        raise ValueError(
            "Previous portfolio value must be positive."
    )

    daily_return = (
        current_portfolio / previous_portfolio
     ) - 1.0

    return daily_return

def calculate_roll_return(
    previous_spx: float,
    previous_call: float,
    soq: float,
    dividend: float,
    old_strike: float,
    spx_vwap: float,
    call_vwap: float,
    current_spx: float,
    current_call: float
) -> dict:
    """
    Calculating the BXM return on a monthly option roll date.

    Returns the expiring option settlement, the three roll components,
    and the compound daily return.
    """
    
    # Calculating the expiring Option Settlement
    settlement = max(0.0, soq - old_strike)

    # Calculating the previous covered portfolio value.
    previous_portfolio = previous_spx - previous_call

    if previous_portfolio <= 0:
        raise ValueError(
            "Previous portfolio value must be positive."
        )

    # retA: from the previous close through old-call settlement
    ret_a = (
        (soq + dividend - settlement)
        / previous_portfolio
    ) - 1.0

    # retB: SPX movement before the new call is sold
    if soq <= 0:
        raise ValueError(
            "SOQ must be positive."
        )

    ret_b = (
        spx_vwap / soq
    ) - 1.0

    # Calculating the new covered portfolio value.
    new_portfolio = spx_vwap - call_vwap

    if new_portfolio <= 0:
        raise ValueError(
            "New portfolio value must be positive."
        )

    # Calculating the closing portfolio value.
    closing_portfolio = current_spx - current_call

    # retC: from selling the new call through the close.
    ret_c = (
        closing_portfolio / new_portfolio
    ) - 1.0

    # Compound the three return components.
    daily_return = (
        (1.0 + ret_a) *
        (1.0 + ret_b) *
        (1.0 + ret_c)
    ) - 1.0

    return {
        "settlement": settlement,
        "ret_a": ret_a,
        "ret_b": ret_b,
        "ret_c": ret_c,
        "ret": daily_return
    }

# =====================================================================
# 3. Find monthly option-roll dates
# =====================================================================

def third_friday(year: int, month: int) -> date:
    """
    Calculate the third Friday of a given month.
    """

    first_day = date(year, month, 1)

    # weekday(): Monday is 0, so Friday is 4.
    friday = 4

    # Number of days until the first Friday of the month.
    days_until_friday = (
        friday - first_day.weekday()
    ) % 7

    # Finding the first Friday.
    first_friday = first_day + timedelta(days=days_until_friday)

    # Finding the third Friday.
    return first_friday + timedelta(days=14)


def get_roll_date(
        year: int,
        month: int,
        calendar
) -> date:
    """
    Return the monthly BXM roll date.
    if the third Friday is a holiday, return the previous trading day.
    """

    scheduled_friday = third_friday(year, month)

    first_day = date(year, month, 1)

    # Retrieving the trading days up to the third Friday.
    sessions = calendar.valid_days(
        start_date=first_day,
        end_date=scheduled_friday
    )

    if sessions.empty:
        raise ValueError(
            f"No trading sessions found "
            f"for {year}-{month:02d}"
        )

    # The final available trading session is the adjusted roll date.
    return sessions[-1].date()


def next_month(year: int, month: int) -> tuple[int, int]:
    """
    Return the year and month immediately following
    the given year and month.
    """
    if month == 12:
        return year + 1, 1

    return year, month + 1

# =====================================================================
# 4. Track the active call option
# =====================================================================

def populate_option_details(
    df: pd.DataFrame,
    calendar
) -> pd.DataFrame:
    """
    Populating the expiration date and active strike of the
    BXM call option held on each trading date.

    The function tracks the active option position and
    updates it whenever the portfolio rolls into a new call.
    """

    df = df.copy()

    if df.empty:
        raise ValueError(
            "Cannot populate option details for an empty dataset."
        )

    # Converting the DataFrame into a list of row dictionaries.
    records = df.to_dict(orient="records")

    # Starting with the initial option position.
    first_row = records[0]

    first_date = pd.Timestamp(
        str(first_row["date"])
    ).date()

    if pd.isna(first_row["k"]):
        raise ValueError(
            "Initial option strike is missing."
        )

    current_strike: float = float(
        first_row["k"]
    )

    # Determining the roll date for the first observation.
    first_roll = get_roll_date(
        first_date.year,
        first_date.month,
        calendar
    )

    # Determining which monthly call is active at the start.
    if first_date < first_roll:

        current_expiry: date = first_roll

    else:

        year, month = next_month(
            first_date.year,
            first_date.month
        )

        current_expiry = get_roll_date(
            year,
            month,
            calendar
        )

    # Storing one expiration and strike for every input row.
    expiration_dates = []
    active_strikes = []

    previous_date = None

    # Updating the position only when a roll takes place.
    for i, row in enumerate(records):

        trading_date = pd.Timestamp(
            str(row["date"])
        ).date()

        # Verifying that trading observations are chronological.
        if previous_date is not None:

            if trading_date <= previous_date:
                raise ValueError(
                    f"Trading dates are not strictly increasing: "
                    f"{trading_date}"
                )

        # No prior row exists for the starting observation.
        if i > 0:

            # Checking whether this observation is the monthly roll.
            monthly_roll = get_roll_date(
                trading_date.year,
                trading_date.month,
                calendar
            )

            # Case A: Monthly option-roll date.
            if trading_date == monthly_roll:

                # The existing option must expire today.
                if current_expiry != trading_date:

                    raise ValueError(
                        f"Unexpected option expiration on "
                        f"{trading_date}"
                    )

                # The new option strike must be provided.
                if pd.isna(row["k"]):

                    raise ValueError(
                        f"New option strike missing on "
                        f"{trading_date}"
                    )

                # Updating to the new active option strike.
                current_strike = float(
                    row["k"]
                )

                # Determining the following month's expiration.
                year, month = next_month(
                    trading_date.year,
                    trading_date.month
                )

                current_expiry = get_roll_date(
                    year,
                    month,
                    calendar
                )

            # Case B: Ordinary trading day.
            else:

                # An existing option cannot remain active after its expiration date.
                if trading_date > current_expiry:

                    raise ValueError(
                        f"Option expired before {trading_date}; "
                        "a roll observation may be missing."
                    )

                supplied_strike = row["k"]

                # If the workbook provided a strike, it must match the strike of the active option.
                if pd.notna(supplied_strike):

                    if abs(
                        float(supplied_strike) - current_strike
                    ) > 1e-9:

                        raise ValueError(
                            f"Unexpected strike change on "
                            f"{trading_date}"
                        )

        # Validating any provided expiration date.
        supplied_expiry = row["expr_date"]

        if pd.notna(supplied_expiry):

            supplied_expiry_date = pd.Timestamp(
                str(supplied_expiry)
            ).date()

            if supplied_expiry_date != current_expiry:

                raise ValueError(
                    f"Incorrect expiration date on "
                    f"{trading_date}"
                )

        # Storing the active option information.
        expiration_dates.append(
            pd.Timestamp(current_expiry)
        )

        active_strikes.append(
            current_strike
        )

        previous_date = trading_date

    # Populating the completed DataFrame.
    df["expr_date"] = pd.to_datetime(
        expiration_dates
    )

    df["k"] = pd.Series(
        active_strikes,
        index=df.index,
        dtype="float64"
    )

    return df


# =====================================================================
# 5. Validate market data
# =====================================================================

def validate_inputs(
        df: pd.DataFrame,
        calendar
) -> pd.DataFrame:
    """
    Validating the input data for the BXM index replication.

    This function raises ValueError when critical input data is missing
    or inconsistent with the index methodology.
    """

    if df.empty:
        raise ValueError(
            "Input data is empty."
        )

    # Validating trading dates.
    if df["date"].isna().any():
        raise ValueError(
            "Missing trading dates in input data."
        )

    if df["date"].duplicated().any():
        raise ValueError(
            "Duplicate trading dates in input data."
        )

    if not df["date"].is_monotonic_increasing:
        raise ValueError(
            "Trading dates must be sorted in increasing order."
        )

    # Checking trading dates against the Cboe options calendar.
    start_date = df["date"].min().date()
    end_date = df["date"].max().date()

    sessions = calendar.valid_days(
        start_date=start_date,
        end_date=end_date
    )

    expected_dates = (
        pd.DatetimeIndex(sessions)
        .tz_localize(None)
        .normalize()
    )

    actual_dates = (
        pd.DatetimeIndex(df["date"])
        .normalize()
    )

    missing_dates = expected_dates.difference(actual_dates)

    unexpected_dates = actual_dates.difference(expected_dates)

    if len(missing_dates) > 0:
        raise ValueError(
            f"Missing exchange trading dates: "
            f"{missing_dates.date.tolist()}"
        )

    if len(unexpected_dates) > 0:
        raise ValueError(
            f"Unexpected non-trading dates: "
            f"{unexpected_dates.date.tolist()}"
        )
    
    # Validating daily market data.
    daily_columns = [
        "spx_close",
        "div",
        "bid",
        "ask",
        "mid",
        "k"
    ]

    for row in df.to_dict(orient="records"):

        trading_date = pd.Timestamp(
            row["date"]
        ).date()

        for column in daily_columns:

            value = row[column]

            if pd.isna(value):
                raise ValueError(
                    f"Missing {column} on {trading_date}"
                )

            if not math.isfinite(float(value)):
                raise ValueError(
                    f"Invalid {column} on {trading_date}"
                )

        # SPX and the active option strike must be positive.
        if row["spx_close"] <= 0:
            raise ValueError(
                f"Invalid SPX close on {trading_date}"
            )

        if row["k"] <= 0:
            raise ValueError(
                f"Invalid option strike on {trading_date}"
            )

        # Ordinary cash dividend contributions cannot be negative.
        if row["div"] < 0:
            raise ValueError(
                f"Negative dividend on {trading_date}"
            )

        # Checking the option bid/ask relationship. Quotes must be nonnegative, with bid no higher than ask.
        if row["bid"] < 0 or row["ask"] < 0:
            raise ValueError(
                f"Negative option quote on {trading_date}"
            )

        if row["bid"] > row["ask"]:
            raise ValueError(
                f"Bid exceeds ask on {trading_date}"
            )

        # Confirming that the provided midpoint matches bid and ask.
        expected_midpoint = calculate_midpoint(
            row["bid"],
            row["ask"]
        )

        if not math.isclose(
            row["mid"],
            expected_midpoint,
            rel_tol=0.0,
            abs_tol=1e-8
        ):
            raise ValueError(
                f"Incorrect option midpoint on "
                f"{trading_date}"
            )


    # Validating monthly roll observations. Roll days need additional prices and settlement inputs.
    records = df.to_dict(orient="records")

    for i in range(1, len(records)):

        row = records[i]
        previous = records[i - 1]

        trading_date = pd.Timestamp(
            row["date"]
        ).date()

        monthly_roll = get_roll_date(
            trading_date.year,
            trading_date.month,
            calendar
        )

        if trading_date != monthly_roll:
            continue

        roll_columns = [
            "spx_10am",
            "soq",
            "settle",
            "vwap",
            "spx_vwap"
        ]

        for column in roll_columns:

            value = row[column]

            if pd.isna(value):
                raise ValueError(
                    f"Missing roll input {column} "
                    f"on {trading_date}"
                )

            if not math.isfinite(float(value)):
                raise ValueError(
                    f"Invalid roll input {column} "
                    f"on {trading_date}"
                )

        if (
            row["spx_10am"] <= 0
            or row["soq"] <= 0
            or row["spx_vwap"] <= 0
        ):
            raise ValueError(
                f"Invalid SPX reference value on "
                f"{trading_date}"
            )

        if row["settle"] < 0 or row["vwap"] < 0:
            raise ValueError(
                f"Invalid option price on {trading_date}"
            )

        # Verifying that the old option expires on this date.
        old_expiry = pd.Timestamp(
            previous["expr_date"]
        ).date()

        if old_expiry != trading_date:
            raise ValueError(
                f"Old option expiration mismatch on "
                f"{trading_date}"
            )

        # Recalculating settlement using the expiring strike
        expected_settlement = max(
            0.0,
            row["soq"] - previous["k"]
        )

        if not math.isclose(
            row["settle"],
            expected_settlement,
            rel_tol=0.0,
            abs_tol=1e-6
        ):
            raise ValueError(
                f"Incorrect option settlement on "
                f"{trading_date}"
            )

        # The new strike must be at or above the supplied SPX strike-selection reference.
        if row["k"] < row["spx_10am"]:

            raise ValueError(
                f"New option strike is below the "
                f"SPX reference on {trading_date}"
            )

    return df

# =====================================================================
# 6. Calculate the return series
# =====================================================================

def compute_returns(
    df: pd.DataFrame,
    calendar
) -> pd.DataFrame:
    """
    Calculate daily BXM returns for every observation.

    Using the three-stage roll calculation on monthly roll
    dates and the ordinary covered-call return otherwise.
    """

    df = df.copy()

    records = df.to_dict(orient="records")

    # Storing calculated results separately.
    returns = []
    ret_a_values = []
    ret_b_values = []
    ret_c_values = []

    # Processing every trading observation.
    for i, row in enumerate(records):

        # The first row sets the base; its daily return is unknown.
        if i == 0:

            returns.append(float("nan"))

            ret_a_values.append(float("nan"))
            ret_b_values.append(float("nan"))
            ret_c_values.append(float("nan"))

            continue

        # Retrieving the previous trading observation.
        previous = records[i - 1]

        trading_date = pd.Timestamp(
            row["date"]
        ).date()

        monthly_roll = get_roll_date(
            trading_date.year,
            trading_date.month,
            calendar
        )

        # Case A: Monthly option-roll date.
        if trading_date == monthly_roll:

            roll_result = calculate_roll_return(
                previous_spx=previous["spx_close"],
                previous_call=previous["mid"],
                soq=row["soq"],
                dividend=row["div"],
                old_strike=previous["k"],
                spx_vwap=row["spx_vwap"],
                call_vwap=row["vwap"],
                current_spx=row["spx_close"],
                current_call=row["mid"]
            )

            returns.append(
                roll_result["ret"]
            )

            ret_a_values.append(
                roll_result["ret_a"]
            )

            ret_b_values.append(
                roll_result["ret_b"]
            )

            ret_c_values.append(
                roll_result["ret_c"]
            )

        # Case B: Ordinary trading day.
        else:

            daily_return = calculate_non_roll_return(
                current_spx=row["spx_close"],
                previous_spx=previous["spx_close"],
                current_call=row["mid"],
                previous_call=previous["mid"],
                dividend=row["div"]
            )

            returns.append(
                daily_return
            )

            # Roll components do not apply on ordinary days.
            ret_a_values.append(float("nan"))
            ret_b_values.append(float("nan"))
            ret_c_values.append(float("nan"))

    # Populating the original Excel calculation columns.
    df["retA"] = ret_a_values

    df["retB"] = ret_b_values

    df["retC"] = ret_c_values

    df["ret"] = returns

    return df

# =====================================================================
# 7. Build the normalized index series
# =====================================================================

def compute_index(
    df: pd.DataFrame,
    initial_level: float = 100.0
) -> pd.DataFrame:
    """
    Calculating the BXM closing index level recursively.

    Assumption: The initial level is normalized to 100 unless
    another positive starting value is provided.
    """

    df = df.copy()

    # Validating the initial index level.
    if not math.isfinite(initial_level):
        raise ValueError(
            "Initial index level must be finite."
        )

    if initial_level <= 0:
        raise ValueError(
            "Initial index level must be positive."
        )

    if df.empty:
        raise ValueError(
            "Cannot calculate index levels for an empty dataset."
        )

    # Initializing the index series.
    current_level = float(initial_level)

    index_levels = [current_level]

    
    # Compounding returns from the second observation onward.
    for daily_return in df["ret"].iloc[1:]:

        if pd.isna(daily_return):
            raise ValueError(
                "Missing daily return detected."
            )

        if not math.isfinite(float(daily_return)):
            raise ValueError(
                "Invalid daily return detected."
            )

        gross_return = 1.0 + daily_return

        if gross_return <= 0:
            raise ValueError(
                "Invalid gross return detected."
            )

        current_level = (
            current_level * gross_return
        )

        if not math.isfinite(current_level):
            raise ValueError(
                "Invalid calculated index level."
            )

        index_levels.append(
            current_level
        )

    # Populating the Index close column.
    df["Index close"] = index_levels

    return df

# =====================================================================
# 8. Export the completed workbook
# =====================================================================

def export_excel(
    df: pd.DataFrame,
    output_path: Path
) -> None:
    """
    Exporting the completed BXM calculations to a formatted
    Excel workbook without modifying the original input.
    """

    if df.empty:
        raise ValueError(
            "Cannot export an empty BXM DataFrame."
        )

    # Creating the output file if it does not exist.
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    # Writing the completed DataFrame to a new Excel workbook.
    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
        datetime_format="yyyy-mm-dd"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="BXM",
            index=False,
            na_rep=""
        )

        # Accessing the worksheet created by pandas.
        worksheet = writer.sheets["BXM"]

        # Keeping the column headers and first two columns visible.
        worksheet.freeze_panes = "C2"

        # Enabling filtering on the full output table.
        worksheet.auto_filter.ref = worksheet.dimensions

        # Hiding default gridlines for a cleaner presentation.
        worksheet.sheet_view.showGridLines = False

        # Formatting the header row.
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="17365D"
        )

        header_font = Font(
            bold=True,
            color="FFFFFF"
        )

        for cell in worksheet[1]:

            cell.fill = header_fill
            cell.font = header_font

            cell.alignment = Alignment(
                horizontal="center"
            )


        # Setting readable column widths.
        worksheet.column_dimensions["A"].width = 16
        worksheet.column_dimensions["B"].width = 16

        for column in "CDEFGHIJKLMNOPQR":

            worksheet.column_dimensions[column].width = 20

        # Formatting the data rows.
        for row in worksheet.iter_rows(
            min_row=2,
            max_row=worksheet.max_row
        ):

            # dates: Columns A–B
            row[0].number_format = "yyyy-mm-dd"
            row[1].number_format = "yyyy-mm-dd"

            # Market inputs and option prices: columns C–M.
            for cell in row[2:13]:

                cell.number_format = "0.00000000"

            # Return components and daily return: N–Q.
            for cell in row[13:17]:

                cell.number_format = "0.000000000000"

            # Normalized BXM index close: column R.
            row[17].number_format = "0.000000000000"

    print(
        f"\nCompleted Excel workbook saved to: "
        f"{output_path}"
    )

# =====================================================================
# 9. Run the full calculation
# =====================================================================

def main() -> None:
    """
    Running the complete Cboe BXM Index replication.

    Reads the supplied Excel workbook, calculates the
    missing option details and index values, and exports
    the completed workbook.
    """

    print("=" * 60)
    print("CBOE BXM INDEX REPLICATION")
    print("=" * 60)

    
    # Loading the original Excel workbook.
    print(f"\nReading input: {INPUT_FILE}")

    df = load_inputs(INPUT_FILE)

    print(
        f"Loaded {len(df)} trading observations."
    )

   
    # Initializing the exchange calendar.
    calendar = mcal.get_calendar(
        CALENDAR_NAME
    )

    # Populating active option strikes and expirations.
    df = populate_option_details(
        df,
        calendar
    )

    print("Option strikes and expirations populated.")

    
    # Validating the prepared market data.

    df = validate_inputs(
        df,
        calendar
    )

    print("Market data validation passed.")

    # Calculating all daily BXM returns.
    df = compute_returns(
        df,
        calendar
    )

    print("Daily BXM returns calculated.")

    # Calculating the normalized BXM index levels.
    initial_level = 100.0

    df = compute_index(
        df,
        initial_level=initial_level
    )

    print("BXM index levels calculated.")

    # Exporting the completed Excel workbook.
    output_path = (
        Path(__file__).resolve().parent
        / "output"
        / "BXM_index_output.xlsx"
    )

    export_excel(
        df,
        output_path
    )

    # Displaying a calculation summary.
    first_date = pd.Timestamp(
        df["date"].iloc[0]
    ).date()

    last_date = pd.Timestamp(
        df["date"].iloc[-1]
    ).date()

    final_level = float(
        df["Index close"].iloc[-1]
    )

    print("\n" + "=" * 60)
    print("REPLICATION SUMMARY")
    print("=" * 60)

    print(f"First trading date: {first_date}")
    print(f"Last trading date:  {last_date}")
    print(f"Trading observations: {len(df)}")

    print(
        f"Initial normalized level: "
        f"{initial_level:.12f}"
    )

    print(
        f"Final normalized level:   "
        f"{final_level:.12f}"
    )

    print(f"\nOutput workbook: {output_path}")

    print("\nBXM replication completed successfully.")



if __name__ == "__main__":
    main()

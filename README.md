# Cboe BXM replication and all exchanges market calendars

This project covers the two Python tasks in the Cboe assessment. For the first task, I used the provided Excel data to reproduce the S&P 500 BuyWrite Index (BXM) calculations and fill in the missing values. For the second, I built a function that finds market holidays and early closes for a chosen exchange and date range. The holiday script can also generate calendars for all exchanges supported by the installed library.

I kept the tasks in separate scripts. Each one can be run on its own and saves its results in an Excel workbook.

## Project files

```text
Cboe_Quant_Assessment/
├── BXM_index.py
├── market_holidays_calendars.py
├── README.md
├── requirements.txt
├── pytest.ini
├── data/
│   └── BXM.xlsx
├── output/
│   ├── BXM_index_output.xlsx
│   ├── market_holidays_output.xlsx
│   └── market_holidays_all_2026.xlsx
└── tests/
    ├── test_BXM_index.py
    └── test_market_holidays_calendars.py
```

## Setup

I developed and tested the project with Python 3.13. From the project folder, install the packages with:

```bash
python -m pip install -r requirements.txt
```

The main packages I used are `pandas` for working with the data, `openpyxl` for Excel files, `pandas_market_calendars` for exchange schedules, and `pytest` for testing. The `requirements.txt` file lists the dependencies and version constraints used for the project.

## Assignment 1: BXM index replication

I read `data/BXM.xlsx` using pandas, prepared the eight trading observations, and filled in the missing option expiration dates, strikes, returns, and index levels. The original Excel file stays unchanged; the script writes a new completed workbook.

I separated the calculations into ordinary trading days and option-roll days because the BXM methodology handles them differently.

On an ordinary day, I value the call at the midpoint of its closing bid and ask:

```text
mid = (bid + ask) / 2
```

The daily return uses the current and previous SPX closes, the two option midpoints, and that day's dividend contribution.

On a monthly roll day, the expiring call is settled first and a new call is sold. I calculated the three return components (`retA`, `retB`, and `retC`) using the supplied settlement and VWAP data, then compounded them to get the day's total return. The settlement calculation uses the *old* strike; the closing option valuation uses the *new* call.

To keep track of the position, I calculated each month's third Friday and moved the roll to the preceding exchange trading day when needed. The active strike and expiration are carried forward until the next roll. I used the new strike supplied in the workbook rather than trying to reconstruct an option chain that was not provided.

Finally, I calculated each index level from the previous level and the daily return:

```text
Index[t] = Index[t-1] × (1 + return[t])
```

### Starting value and result

The input starts on August 20, 2026, but does not give the published BXM level for that date. I therefore set the first index level to **100** and calculated the remaining levels from there. The first row has no daily return because there is no preceding observation in the supplied data.

The final **normalized** index level on August 31, 2026, is:

```text
100.487620873207
```

This is the result for the provided sample starting at 100, not the published historical BXM index level.

### Run it

```bash
python BXM_index.py
```

The completed workbook is saved to `output/BXM_index_output.xlsx`, on the `BXM` sheet. It keeps the original 18 columns and fills in the calculated values. The roll-return component cells are intentionally blank on non-roll days.

I added checks for missing or duplicate dates, missing market inputs, invalid quotes or midpoints, inconsistent option settlements, and unexpected strike changes. The script stops with an error if it cannot calculate a result reliably.

## Assignment 2: Market holiday calendar

### What I did

I wrote `get_market_holidays()` so I could pass in an exchange name, start date, and end date instead of maintaining a separate hardcoded holiday list for each exchange.

I used `pandas_market_calendars` to get the exchange's trading schedule. My function compares the days an exchange would normally trade with its actual trading sessions to identify full-day closures. It reports early-close sessions separately, since those are still trading days. Ordinary weekends are not included in the holiday list.

The script accepts one or more exchange names supported by the installed calendar library. It also has an `--all` option to generate calendars for all distinct exchanges available in the library without entering their names individually.

Some calendars have multiple registered names, so I remove duplicate aliases when building the full exchange list. If a calendar cannot be processed during an `--all` run, the script records its name and error in a separate worksheet and continues with the remaining exchanges.

I tested the individual-exchange option with `XNYS` and `CBOE_Index_Options`.

**Calendar-data note:** The installed Cboe index-options calendar omitted December 24, 2026, as an early close, even though it appears in Cboe's published U.S. options schedule. I added a correction limited to that exchange and date. For production use, I would check the calendar provider against official exchange notices, especially for unusual closures or future schedules.

Cboe trading-hours reference: https://www.cboe.com/about/hours/us-options

### Run it

For NYSE and Cboe over the full 2026 calendar year:

```bash
python market_holidays_calendars.py --exchange XNYS CBOE_Index_Options --start 2026-01-01 --end 2026-12-31
```

This saves the results to `output/market_holidays_output.xlsx`.

For a shorter period and one exchange:

```bash
python market_holidays_calendars.py --exchange XNYS --start 2026-11-23 --end 2026-11-30
```

To generate calendars for all supported exchanges over the full 2026 calendar year:

```bash
python market_holidays_calendars.py --all --start 2026-01-01 --end 2026-12-31 --output output/market_holidays_all_2026.xlsx
```

In my final all-exchange run, the library provided 106 distinct calendars. The script successfully processed 105 and recorded one failure (`XSGO`) in the workbook. The other exchanges were processed without stopping the run.

To leave early closes out of the results, add `--no-early-closes`. To choose a different output file, add `--output path/to/file.xlsx`.

The output workbooks contain three sheets:

- **Summary:** date range and counts of full-day holidays and early closes for each successfully processed exchange.
- **Calendar Events:** exchange, event date, and event type.
- **Failed Calendars:** exchange names and error messages for any calendars that could not be processed. This sheet is empty when there are no failures.

The date range is inclusive, and command-line dates use `YYYY-MM-DD`.

## Tests

I kept the tests separate for the two assignments:

```text
tests/test_BXM_index.py
tests/test_market_holidays_calendars.py
```

The `pytest.ini` file sets the project folder as the Python import path for the tests.

To run the tests from the project folder:

```bash
python -m pytest -q
```

The tests check the BXM formulas and final index value, option expiration logic, selected invalid inputs, Excel export, holidays, early closes, and invalid calendar requests.

**All 18 tests passed in my final run.**

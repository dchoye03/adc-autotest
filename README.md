# GW9251 ADC Characterization Automation

[한국어](README.md)

Automates the manual GW9251 (24-bit ADC) characterization workflow — DAC
voltage setup → register configuration → 1024-sample capture → paste into the
Excel report — previously done by hand in Tera Term.

- `gw9251_autotest.py` — CLI. Auto-identifies the boards, sets the DAC,
  writes/verifies registers, captures samples, saves a CSV backup and writes
  the Excel report; one full DUT cycle per run
- `gw9251_gui.pyw` — Tkinter GUI front-end (imports the module above).
  DAC codes / sample count / DUT number / Excel file selection, live log,
  result summary (ENOB · STD · accuracy), auto-creates DUT blocks in the
  report when needed
- `test_gw9251.py` — regression tests, no hardware required
  (`python test_gw9251.py`)

## Hardware setup

- Windows PC ↔ two STM32 Nucleo boards over USB (ST-Link VCP, VID 0x0483)
- **DAC board**: drives the differential input. Boots into a chip-select menu
  (`Usage: gw9121 | gw9241 | gw9251`) → enter `gw9241`, then `dac init`,
  `dac set 1 <val>`, `dac set 2 <val>`
- **Test board**: controls the GW9251EVM (DUT). Boots directly into the
  `GW9251:` prompt
- Power is removed between DUTs (chip swap), so USB re-enumerates and DAC
  state is lost every cycle — the script handles both automatically

## Install / run

```
pip install pyserial openpyxl

# CLI: write DUT #6 into test_copy.xlsx
python gw9251_autotest.py --excel test_copy.xlsx --dut 6

# GUI: double-click, or
pythonw gw9251_gui.pyw
```

- Close Tera Term (or any serial terminal) first — COM ports are exclusive
- Keep the Excel report and the Python files in the same folder
- Don't keep the report open in Excel during a test (a retry dialog appears
  on save if you do)

## What one DUT cycle does

1. Wait for both boards on USB → identify DAC vs test board by probing their
   firmware prompts (no config file needed)
2. Confirm the test board shows the `GW9251:` prompt
3. DAC board: `exit` → `gw9241` → `dac init` → `dac set 1/2`
4. Internal Short mode: `wr 0x03 0x02` → verify via `rr` → `wr 0x04 0x60` →
   verify → `rd <n>` capture (default 1024 samples at ~11 samples/s ≈ 95 s)
5. Channel A mode: repeat with `wr 0x04 0x00`
6. Save the CSV backup + write the samples into the report's DUT block
7. Power off → swap chip → power on → next DUT

## Firmware quirks (do not regress)

1. **Output interleaves NUL bytes and ANSI escapes** (including the
   non-standard `ESC[-1D`). All received data must pass through `clean()`
   before parsing
2. **Line ending is CR only** (`LINE_END = "\r"`). Sending CRLF leaves a
   stray `\n` that corrupts the next command (firmware prints usage instead
   of executing)
3. **Submenu state survives serial reconnects** (only a power cycle resets
   it). Send `exit` before `gw9241` on the DAC board
4. **Never verify chip select by chip name**: a rejected select prints the
   usage line containing every chip name, and a successful select prints a
   submenu that doesn't contain the name. Check submenu markers instead
   (`dac` / `9251`); `gw9121` in a response means still at the top menu
5. **The `rd` stream can pause mid-capture for seconds**: capture end is
   decided by an idle timeout (20 s of silence) plus a 300 s hard cap, and
   every capture reports why it ended, elapsed time, and the longest
   mid-stream gap — check those numbers first when something looks wrong
6. **`rr` reply format**: echo, then `RREG:A+D: <addr> <val>` (hex). Parsed
   exactly; note the `A`/`D` decorations are valid hex digits, so don't
   loosen the parser back to token soup

## Excel report layout

- Row 5: merged `DUT#n` headers every 2 columns
- Row 10: mode headers (`Internal Short (0x04 0x60)` / `Channel A (0x04 0x00)`)
- Rows 11–20: ENOB / STD / accuracy formulas (never touched by the script)
- Rows 22–1045: the 1024 samples
- The target block is located dynamically by regexing the row-5 headers and
  cross-checking row 10 before writing. With `--dut` omitted, the first empty
  block is used
- If the requested DUT block doesn't exist, the GUI clones the last block's
  template (formulas re-pointed, styles, merges) into new columns — asking
  before clearing any leftover values

## Regression tests

```
python test_gw9251.py
```

61 checks covering the serial parser, register-verify matching, GUI input
validation, result math, and Excel block cloning. No hardware needed. Run
after every change.

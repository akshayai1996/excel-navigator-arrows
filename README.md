# excel-navigator-arrows
Excel VBA focus-mode navigator for filtered rows
# Excel Navigator Arrows (Focus Mode)

High-performance Excel VBA add-in that lets you navigate filtered rows like a playlist — one row at a time — with instant focus mode.

## Features
- Works with existing Excel filters
- Single-row "Focus Mode" navigation
- Left / Right arrow navigation
- Safe for 20k+ rows
- Excel 2010 → Microsoft 365 compatible
- State-safe (no hidden row traps)
- Designed for XLAM add-ins

## How It Works
1. Apply a filter on your data
2. Run `NavigatorArrows_Apply`
3. Use arrow buttons to move next / previous
4. Run `NavigatorArrows_Remove` to restore view

## Public Macros
- `NavigatorArrows_Apply`
- `NavigatorArrows_Remove`

## Installation
1. Open Excel → ALT + F11
2. Import `modNavigatorArrows.bas`
3. (Optional) Save as `.xlam` for add-in usage

## License
MIT

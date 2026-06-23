# Monthly EIS Calculator Script Spec

## Goal
Create or maintain a Python script that reads one EIS monthly Excel file and one or more employee WorkingNote Excel files, then automatically fills the EIS project columns by summing task quantities from WorkingNote column E.

## Input files
1. `EIS.xlsx`
   - Sheet: auto-detected by looking for a `Summary` header in row 1. If multiple sheets exist, the first sheet with a `Summary` header is used. Can be overridden with `--eis-sheet`.
   - Row 1 contains project names from column C onward; the rightmost project column is followed by `Summary`.
   - Column A contains official EIS work item names.
   - Optional: `Project_Alias` sheet (columns: `Alias`, `Real_Project`) — maps WorkingNote project names to canonical EIS project names.
   - Optional: `Item_Alias` sheet (columns: `Alias`, `EIS_Item`) — maps WorkingNote task descriptions to canonical EIS item names.
2. One or more `WorkingNote.xlsx`
   - Column E header contains `工作項目`.
   - Each cell contains one work record per line in dash format (see below).

## WorkingNote format
Each line in the `工作項目` cell follows this format:

```text
ProjectName - Task Description -xN
```

- `ProjectName` — the project or special category name; everything before the first `-`.
- `Task Description` — must match an EIS column A item as closely as possible.
- `-xN` — quantity; `N` is a number (integer or decimal). Can also be written as `xN` without a leading `-`. If omitted, defaults to `1`.
- Status tags like `[Complete]`, `[In Progress]` may appear anywhere and are automatically stripped.
- Lines that start with a space/indent are ignored.
- Lines containing only a bracket label (e.g. `[05/16]`) are ignored.

### Example
```text
Aymara - Release MSA Report -x1
Aymara - Review Logs (Fail / Incosistant / Mixed Logs) -x3
XYZ_Project - Provide Measurement Distribution Chart -x2
XYZ_Project - Review Logs (Fail / Incosistant / Mixed Logs)
Kore - SWQA internal process review -x1
Other - Study new testing framework -x1
RFI_RFQ - Review RFQ document -x2
```

## Parsing rules
- Each line is parsed by `parse_dash_line`: split on the first `-` to get project and task, then strip the trailing quantity pattern (`-xN` or `xN`).
- Status tags `[...]` anywhere in the task description are removed before matching.
- If no quantity suffix is present, quantity defaults to `1`.
- Lines starting with whitespace are skipped entirely.
- Lines that are only a bracket label (e.g. `[05/16]`) are skipped.

## EIS writing rules
- Match the parsed work item to EIS column A using `resolve_task_row`:
  1. `Item_Alias` sheet lookup (exact match on normalised key).
  2. Exact normalised match.
  3. Containment match (EIS item is contained in task text, or vice versa).
  4. Token subset match (≥ 4 tokens, one set is a subset of the other).
  5. Conservative fuzzy match (cutoff 0.88).
- Match the parsed project name to EIS row 1 project names using `project_col_by_norm` (normalised key); apply `Project_Alias` if defined.
- Write/sum the quantity into the intersecting cell. If the cell already has a number, add to it.
- If multiple employees have the same project and task, their quantities are accumulated.
- If the project does not exist in row 1, a new column is inserted immediately before the `Summary` column. The `Summary` column **must** exist; the script raises an error if it is missing. The new column copies formatting from the adjacent project column. Summary formulas are rebuilt to include the new column.

## Special category rules
The following project names are not written into EIS project columns; they are collected separately:
- `Kore`
- `Other` / `Others`
- `RFI_RFQ`

Their task quantities are written to `Special_Category_Summary` sheet.

## Output sheets written to the output workbook

| Sheet | Content |
|-------|---------|
| `Special_Category_Summary` | Category, Task, Quantity for Kore / Other / RFI_RFQ entries |
| `Unknown_Project_Summary` | Project, Task, Quantity for records where a new project column was created |
| `Unknown_Item_Summary` | Project, Task, Quantity for records whose task could not be matched to EIS column A |
| `Unmatched_WorkingNote_Detail` | Full detail rows (Project, Task, Quantity, Source File, Source Sheet, Source Row, Reason) for unmatched items |
| `Process_Log` | Full processing log for every parsed record (status, matched EIS row, matched EIS project column) |

## Command line usage
```bash
python monthly_eis_calculator.py \
  --eis EIS.xlsx \
  --working-notes WorkingNote_A.xlsx WorkingNote_B.xlsx \
  --output EIS_calculated.xlsx \
  [--eis-sheet <sheet_name>]
```

| Argument | Required | Description |
|----------|----------|-------------|
| `--eis` | Yes | EIS template / monthly file |
| `--working-notes` | Yes | One or more WorkingNote files (space-separated) |
| `--output` | Yes | Output xlsx path |
| `--eis-sheet` | No | EIS sheet name; auto-detected if omitted |

## Important implementation notes
- Alias mappings are stored in optional sheets (`Project_Alias`, `Item_Alias`) inside EIS.xlsx — not hardcoded in the script.
- Do not overwrite the original EIS file; always save to a new output file.
- Preserve EIS formatting when adding a new project column (copy style from the adjacent column).
- When new project columns are inserted, Summary formulas are explicitly rebuilt to include them.
- The EIS active sheet is not blindly used as the main sheet because saving from Excel with an alias/helper sheet selected may change the active sheet. Auto-detection looks for a `Summary` header in row 1.
- Keep the script compatible with VS Code terminal execution.

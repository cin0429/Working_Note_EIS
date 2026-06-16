# Monthly EIS Calculator Script Spec

## Goal
Create or maintain a Python script that reads one EIS monthly Excel file and one or more employee WorkingNote Excel files, then automatically fills the EIS project columns by summing task quantities from WorkingNote column E.

## Input files
1. `EIS.xlsx`
   - Sheet: usually one monthly sheet, for example `EIS_6`.
   - Row 1 contains project names from column C onward.
   - Column A contains official EIS work item names.
   - Column B contains hours per item, but the script currently fills quantity counts, not total hours.
2. One or more `WorkingNote.xlsx`
   - Column E header contains `工作項目`.
   - Each cell may contain multiple project sections and tasks.

## WorkingNote format example
```text
[05/16]
[Midwest]
1. Release MSA GRR report *2 [Complete]
2. Provide measurement distribution chart * 1 [Complete]
[Maera]
1. Provide measurement distribution chart * 1 [Complete]
[5/18]
[Midwest]
1. Release MSA GRR report *2 [Complete]
```

## Parsing rules
- Text inside `[]` is treated as either a project/category or a date.
- `[05/16]`, `[5/18]`, etc. are dates and ignored.
- Other bracket labels are project names unless they are special categories.
- Numbered lines such as `1. xxx *2 [Complete]` are parsed as work items.
- Quantity is read from `*2`, `* 1`, etc.
- If no `* number` exists, default quantity is `1`.
- `[Complete]`, `[In Progress]`, etc. are removed from matching text.
- Continuation lines that are not numbered are ignored.

## EIS writing rules
- Match the parsed work item to EIS column A.
- Match the parsed project to the project name in EIS row 1.
- Write/sum the quantity into the intersecting cell.
- If multiple employees have the same project and task, add the quantities together.
- If the project does not exist in row 1, insert a new column before `Summary` if present; otherwise append at the far right.

## Special category rules
The following bracket labels are not written into project columns:
- `[Kore]`
- `[Other]`
- `[RFI_RFQ]`

Instead, summarize their task quantities in a new sheet called `Special_Category_Summary`.

## Unmatched handling
If a WorkingNote task cannot be matched to EIS column A, do not guess silently. Write it to `Unmatched_WorkingNote_Items` with:
- project/category
- task text
- quantity
- source file
- source sheet
- reason

## Command line usage
```bash
python monthly_eis_calculator.py --eis EIS.xlsx --working-notes WorkingNote_A.xlsx WorkingNote_B.xlsx --output EIS_calculated.xlsx
```

## Login / web link phase
If WorkingNote files are provided as web links and login is required, use Playwright later. The recommended first stable version should process downloaded Excel files first. After the Excel logic is verified, add a second layer:
1. User inputs one or more URLs.
2. Script opens browser by Playwright.
3. If login page appears, user manually logs in.
4. Script downloads each Excel file.
5. Script runs the same local Excel calculation logic.

## Important implementation notes
- Keep task alias mapping editable at the top of the script.
- Do not overwrite the original EIS file; always save to a new output file.
- Preserve EIS formatting when adding a new project column.
- Keep the script compatible with VS Code terminal execution.

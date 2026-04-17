# Whole School Import Template Guide

Template file:
`whole_school_results_template.xlsx`

Use this workbook with:
`Marks` -> `Import Whole School Results`

## What Is Inside

- `SUMMARY - Instructions`
  This sheet is only for guidance and will be skipped during import.
- `PP1`, `PP2`
- `Grade 1` to `Grade 9`
- if a class already has streams in the system, the workbook creates one sheet per stream
  Example: `Grade 4 Blue`, `Grade 4 Red`

Each class has its own sheet, or one sheet per stream where streams exist.

## Required Layout

Keep the class sheets in this format:

- Row 4 contains the import header
- Column `A`: `No`
- Column `B`: `Admission No`
- Column `C`: `Learner Name`
- From column `D` onward: subject columns

Do not remove the `Learner Name` column.

## How To Fill It

- Enter one learner per row starting on row 5.
- Enter marks as numbers from `0` to `100`.
- Leave cells blank where a learner has no mark.
- `Admission No` is optional.
- You can leave unused subject columns blank for the whole sheet.
- Subject order follows the school template bands:
  `PP1-PP2`
  `LANG | MATH | ENVI | CRE | CREATIVE`
  `Grade 1-3`
  `LANG | MATH | KIS | ENVI | CRE | CREATIVE | FRENCH`
  `Grade 4-6`
  `ENG | MATH | KIS | SCI | AGRI | SST | CRE | CREATIVE | FRENCH`
  `Grade 7-9`
  `MATH | ENG | KIS | INT. SCI | AGRI | SST | CRE | C/A | PRE-TECH | FRENCH`

## Sheet Names

Keep the class sheet names unchanged where possible:

- `PP1`
- `PP2`
- `Grade 1` to `Grade 9`
- `Grade 4 Blue`
- `PP1 Red`

If you use streams, names like `Grade 4 Blue` can still work because the importer detects the class name from the sheet title.
The importer now also detects pre-primary stream sheets like `PP1 Blue`.

## Import Steps

1. Open the app.
2. Go to `Enter Marks`.
3. Click `Import Whole School Results`.
4. Choose the correct `Term`.
5. Choose the correct `Exam`.
6. Select `whole_school_results_template.xlsx` after filling it.

## Notes

- Sheets with words like `SUMMARY`, `ANALYSIS`, or `OVERALL` are skipped automatically.
- Blank learner rows are ignored.
- Blank mark cells are ignored.
- New classes and subjects can be added automatically by the import flow if needed.

# ExcelDiff
Excel Macro to check for changes in excel files

This macro is designed to help check changes in excel files.

Provide the names of the old and new files.
Both files have to be open.

A copy of the new file is created first.
The all shapes (Pictures etc.) are removed.
All conditional formatting is removed.
All text is turned black.
All cell backgrounds are reset.
All hidden sheets are made visible.
The purpose of this is to make all content visible.

The macro then tries to match all sheets to sheets in the old file.
If sheet names don't match you will be asked to associate a sheet name.
If the sheet is new, this can be specified and it will be ignored in the comparison.
Deleted sheets will be automatically ignored.

The comparison is done on Formulas. That means:
Changes in formulas will be highlighted.
Changes in entered values will be highlighted.
If the calculated value in a cell changes due to different input values, it is not highlighted.
The background of the cells with changes gets set to red.

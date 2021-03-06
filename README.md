# excel-VBA
Attempt to export a VBA module, which moves active cell down one square in Excel

## To use
Copy modules from VBA.txt into Excel Macros

Then use keyboard shortcut Ctrl + d to run macro to move active cell down by one

Recommended shortcuts:
    - go into Macros
    - Options
    - set shortcuts:
        - up = Ctrl + r
        - down = Ctrl + f
        - left = Ctrl + d
        - right = Ctrl + g

if this fails click on macro / module named down() to run

## Bugs 
### Debug
was only working on single digit rows
    - use ActiveCell.Row
        instead of:
        Right(live_cell, 1)

only works on single letter columns currently

## Development
active cell
    - down.txt has module to move active cell down by 1
    - up.txt has module to move active cell up by 1
        - also has comments

## References
### Export VBA
[How can I export VBA code to text from de modules?](https://stackoverflow.com/questions/58490045/how-can-i-export-vba-code-to-text-from-de-modules#58490363)

### Move Active Cell Down one in Excel
[vba active cell](https://www.wallstreetmojo.com/vba-active-cell/)
[increment cell value by one](https://stackoverflow.com/questions/51521576/increment-cell-values-in-a-range-by-1-with-vba-excel)
[increment string character by one](https://www.mrexcel.com/board/threads/increment-each-of-the-character-in-a-string-by-1-vba.1024767/)
[string manipulation](https://www.excel-easy.com/vba/string-manipulation.html)
[string to number](https://www.automateexcel.com/vba/convert-text-string-to-number/)
[concatenate](https://www.educba.com/vba-concatenate-strings/)
[integer to string](https://stackoverflow.com/questions/11595226/how-do-i-convert-an-integer-to-a-string-in-excel-vba)
[vba active cell](https://www.educba.com/vba-active-cell/)
[get selected rows](https://www.excelvbasolutions.com/2021/03/get-selected-rows-using-vba-macro.html)
[VBA – Get the Active Cell’s Column or Row](https://www.automateexcel.com/vba/activecell-row-column/)
[How To Automatically Increase Letter By One To Get The Next Letter In Excel?](https://www.extendoffice.com/documents/excel/4027-excel-increase-letter-by-one.html)
[Excel VBA - increment column reference in range](https://stackoverflow.com/questions/44280296/excel-vba-increment-column-reference-in-range)
[VBA – Get the Active Cell’s Column or Row](https://www.automateexcel.com/vba/activecell-row-column/)
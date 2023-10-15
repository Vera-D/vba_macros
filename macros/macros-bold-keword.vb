Sub bold_keywords()
' Created by Vera
' This vba script colorizes keywords in a worksheet labeled
' condensed list using the words in the 1st column
' of a tab labeled keywords

Dim lRow As Long
Dim lCol As Long
    'sheet_id = "MDD Annex I"
    'ref_row = 1
    'ref_col = 1
    
    sheet_id = "Annex I"
    ref_row = 2
    ref_col = 2
    
    'sheet_id = "HS"
    'ref_row = 2
    'ref_col = 2
    
    'Find the last non-blank cell in column of column condensed list tab
    lRow = Sheets(sheet_id).Cells(Rows.Count, ref_col).End(xlUp).row
    
    'Find the last non-blank cell in row 1
    lCol = Sheets(sheet_id).Cells(ref_row, Columns.Count).End(xlToLeft).Column
    
    ' set a range of cells that will be formatted
    Dim rng As Range
    Dim rng_col As Variant
    
    ' loop through the columns you will colorize
    'For c = 0 To 1
    For c = 1 To 4
        Set rng = Worksheets(sheet_id).Range(Sheets(sheet_id).Cells(ref_row + 2, ref_col + c), _
        Sheets(sheet_id).Cells(lRow, ref_col + c))
        
        rng.Select
        Size = rng.Count
        Call colorize(rng)
    Next c
    Debug.Print ("ran bold_keywords")
End Sub

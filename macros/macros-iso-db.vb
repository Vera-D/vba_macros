Sub iso_db()
' Created by Vera
' This vba script colorizes keywords in a worksheet labeled
' condensed list using the words in the 1st column
' of a tab labeled keywords

Dim lRow As Long
Dim lCol As Long
    
    sheet_id = "temp"
    ref_row = 1
    ref_col = 1
    
    'Find the last non-blank cell in column of column condensed list tab
    lRow = Sheets(sheet_id).Cells(Rows.Count, ref_col).End(xlUp).row
    
    'Find the last non-blank cell in row
    lCol = Sheets(sheet_id).Cells(ref_row, Columns.Count).End(xlToLeft).Column
    
    ' set a range of cells that will be formatted
    Dim rng As Range
    Dim rng_col As Variant
    Dim arr(1) As Variant
    ' loop through the columns and create a newlist
    Dim n_r As Integer
    n_r = 13
    For r = 24 To lRow
        temp = Sheets(sheet_id).Cells(r, ref_col) & "|" & Sheets(sheet_id).Cells(r + 1, ref_col) _
                & "|" & Sheets(sheet_id).Cells(r, ref_col + 1) & "|" & Sheets(sheet_id).Cells(r, ref_col + 2) _
                & "," & Sheets(sheet_id).Cells(r + 1, ref_col + 2)
                
        Sheets(sheet_id).Cells(n_r, ref_col + 5) = temp
        
        r = r + 1
        n_r = n_r + 1
    Next r
    
    Debug.Print ("ran bold_keywords")
    
End Sub
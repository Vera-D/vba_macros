Sub get_stds()
' This function finds the unique values in the harmonized standards column.
    Dim lRow As Long
    Dim lCol As Long
    Dim g_arr() As Variant
    
    Dim varIn As Variant
    Dim varUnique As Variant
    Dim iInCol As Long
    Dim iInRow As Long
    Dim iUnique As Long
    Dim nUnique As Long
    Dim isUnique As Boolean
    
    sheet_id = "Annex I"
    ref_row = 2
    ref_col = 8
    
    'Find the last non-blank cell in column of column condensed list tab
    lRow = Sheets(sheet_id).Cells(Rows.Count, ref_col).End(xlUp).row
    'Find the last non-blank cell in row 1
    lCol = Sheets(sheet_id).Cells(ref_row, Columns.Count).End(xlToLeft).Column
        
    first = 4
    last = lRow
    nUnique = 0
    ReDim Preserve g_arr(nUnique)
    
    For r = first To last
        
        'Debug.Print ("row:" & " ;" & r & " ;" & c & " ;" & g_arr(c))
        
        temp_str = Worksheets(sheet_id).Cells(r, ref_col)
        cell_std_arr = Split(temp_str, ",")
        
        isUnique = True
        
        For i = LBound(cell_std_arr) To UBound(cell_std_arr)
            'save the value as string to a temp variable
            If (i = 0) Then
                If (r = first) Then
                    txt = cell_std_arr(i)
                    g_arr(nUnique) = txt
                    nUnique = nUnique + 1
                    ReDim Preserve g_arr(nUnique)
                Else
                    'Debug.Print (cell_std_arr(i))
                    txt = cell_std_arr(i)
                    
                    For iUnique = 0 To UBound(g_arr)
                        If txt = g_arr(iUnique) Then
                            isUnique = False
                            Exit For
                        End If
                    Next iUnique
                    
                    If isUnique = True Then
                        
                        g_arr(nUnique) = txt
                        nUnique = nUnique + 1
                        ReDim Preserve g_arr(nUnique)
                    End If
            End If
            Else
                'Debug.Print (Mid(cell_std_arr(i), 2, 10))
                txt = Mid(cell_std_arr(i), 2, 10)
                
                For iUnique = 1 To UBound(g_arr)
                    If txt = g_arr(iUnique) Then
                        isUnique = False
                        Exit For
                    End If
                Next iUnique
                
                If isUnique = True Then
                    
                    g_arr(nUnique) = txt
                    nUnique = nUnique + 1
                    ReDim Preserve g_arr(nUnique)
                    
                End If
                
            
            End If
            

        Next i
        
        If r = lRow Then
            For j = 0 To UBound(g_arr)
                
                Debug.Print ("row:" & " ;" & r & " ;" & j & " ;" & g_arr(j))
            
            Next j
        End If        
    Next r

End Sub





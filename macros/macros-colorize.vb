Sub colorize(rng As Range)
' Color code words in cells that offer clues to the standard titles

'array of keywords
Dim key_words() As Variant
Dim i As Integer

'terms = Worksheets("keywords").Range("A1:A4")
lRowTerms = Worksheets("keywords").Cells(Rows.Count, 1).End(xlUp).row

ReDim key_words(lRowTerms)

'store the keywords in an array
For i = 1 To lRowTerms
    word = Worksheets("keywords").Cells(i, 1)
    key_words(i) = word
Next

Dim rCell As Range, sToFind As String, iSeek As Long

' Colorize the cell
    For Each st In rng
        wd = st.Value
        wd = LCase(wd)
        
        For j = 1 To UBound(key_words)
                sToFind = key_words(j)
                iSeek = InStr(1, wd, sToFind)
            Do While iSeek > 0
                st.Characters(iSeek, Len(sToFind)).Font.Bold = True
                If sToFind = "annex xvi" Or sToFind = "part 3 of annex vi" Then
                    st.Characters(iSeek, Len(sToFind)).Font.Color = vbRed
                ElseIf sToFind = "active" Or sToFind = "non-implantable" Or _
                       sToFind = "implant" Or sToFind = "report" Or sToFind = "risk" Or _
                       sToFind = "invasive" Or sToFind = "sterile" Or sToFind = "implantable" _
                       Then
                    st.Characters(iSeek, Len(sToFind)).Font.Color = vbMagenta
                Else
                    st.Characters(iSeek, Len(sToFind)).Font.Color = vbBlue
                End If
                iSeek = InStr(iSeek + 1, st.Value, sToFind)
                        
            Loop
        Next j
    Next st

End Sub


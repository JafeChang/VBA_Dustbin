Sub MyScript()

'ActiveDocument.Tables(1).Cell(1, 3).Range.InsertAfter (ActiveDocument.Tables(2).Cell(1, 1).Range.Text)
row1count = ActiveDocument.Tables(1).Rows.Count
col1count = ActiveDocument.Tables(1).Columns.Count
row2count = ActiveDocument.Tables(2).Rows.Count
col2count = ActiveDocument.Tables(2).Columns.Count
row1 = 1
row2 = 3
col1 = 1
col2 = 1
enterStr = Chr(13) & Chr(7)
For row2 = 4 To row2count
    insertText = enterStr
    cellText = ActiveDocument.Tables(2).Cell(row2, 1).Range.Text
    If Right(cellText, 3) = (b & enterStr) Then
    Else
        barP = InStr(1, cellText, "-")
        strL = Left(cellText, barP - 1)
        strR = Mid(cellText, barP + 1, 1)
        nl = CInt(strL)
        nr = CInt(strR)
        If nr <> 1 Then
            energy = ActiveDocument.Tables(2).Cell(row2, 2).Range.Text
            energy = Left(energy, Len(energy) - 2)
            insertText = insertText & "¦¤E=" & energy & " eV"
        End If
        symmetry = Trim(Left(Right(ActiveDocument.Tables(2).Cell(row2, 3).Range.Text, 5), 3))
        insertText = insertText & " " & symmetry
        ActiveDocument.Tables(1).Cell(nl + 1, nr + 2).Range.InsertAfter (insertText)
        'cellText
    End If
   
Next
i = ActiveDocument.Tables(2).Rows.Count
 'MsgBox (i)
t = ActiveDocument.Tables(2).Cell(2, 1).Range.Text
'enter = 13 7
x = CInt(Left(t, 1)) + 1
MsgBox (x)

End Sub



Sub myScript()
enterStr = Chr(13) & Chr(7)
For i = 3 To 22
    strText = ActiveDocument.Tables(2).Cell(i, 1).Range.Text
    strText = Left(strText, Len(strText) - 2)
    strText = enterStr & strText
    ActiveDocument.Tables(1).Cell(i, 2).Range.InsertAfter (strText)
Next
MsgBox ("OK")
End Sub

Attribute VB_Name = "FireCopiaESepara"
Sub CopiaeSepara()
      Range("A3").Select
      Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(0, 27).Select
        If (Not IsEmpty(ActiveCell)) Then
            
            ActiveCell.EntireRow.Copy
            ActiveCell.Offset(1).EntireRow.Insert shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
            ActiveCell.Offset(1).EntireRow.PasteSpecial xlPasteFormats
            
            ActiveSheet.Range(Cells(ActiveCell.Row, 28), Cells(ActiveCell.Row, 46)).Copy ActiveSheet.Range(Cells(ActiveCell.Row, 8), Cells(ActiveCell.Row, 26))

         End If
         ActiveCell.Offset(1, -ActiveCell.Column + 1).Select
      Loop
      Columns("AB:AT").EntireColumn.Delete
      Range("A3").Select
End Sub

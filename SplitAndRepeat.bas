Attribute VB_Name = "SplitAndRepeat"
Sub SplitAndRepeat()
      Dim MaxLentgh As Integer
      Dim Value As Variant
      Dim Column As Integer
      Dim c As Integer
      MaxLentgh = 2
      Range("A1").Select
      Do Until IsEmpty(ActiveCell)
         Do Until IsEmpty(ActiveCell)
            If (ActiveCell.Column > 1 And (ActiveCell.Column - 1) Mod MaxLentgh = 0) Then
                
                Column = ActiveCell.Column
                
                ActiveCell.EntireRow.Copy
                ActiveCell.Offset(1).EntireRow.Insert shift:=xlDown, CopyOrigin:=xlFormatFromRightOrAbove
                ActiveCell.Offset(1).EntireRow.PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
                ActiveCell.Offset(-1, Column - 1).Select
                Do Until IsEmpty(ActiveCell)
                    ActiveCell.Clear
                    ActiveCell.Offset(0, 1).Select
                Loop
                ActiveCell.Offset(1, -ActiveCell.Column + 1).Select
                For c = 1 To MaxLentgh
                    ActiveCell.Delete shift:=xlToLeft
                Next
            End If
            ActiveCell.Offset(0, 1).Select
         Loop
         ActiveSheet.Cells(ActiveCell.row, 1).Select
         ActiveCell.Offset(1, 0).Select
      Loop
End Sub

Option Explicit

' Save and Backup ticket data

Sub cmdBackUpStats()

  ActiveWorkbook.Save
  Sheets("Cumulative Stats").Select
  ActiveSheet.Range("A1:H43").Select
  ActiveWorkbook.EnvelopeVisible = True
  With ActiveSheet.mailEnvelope
    .Item.to = 
    .Item.Subject = "Ticket Stats Backup"
    
    ' Use .Item.display below if you want to edit before sending
    '.Item.display
    .Item.send
  End With
  
  'Copy over daily log before saving and quitting
  ' The following few lines copy the data we want
  Sheets("Cumulative Stats").Select
  Range("G2:G10").select
  Selection.Copy
  
  ' Go to the sheet we want to paste the data to
  Sheets("Daily Log").Select
  Range("B1").Select
  
  ' Find the first empty cell
  Selection.End(xlDown).Select
  ActiveCell.Offset(1, 0).Select
  
  ' Transpose the previously copied data and get rid of formatting
  Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
  Application.CutCopyMode = False
  With Selection.Interiior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  
  ' Clear boarders and formatting
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  Selection.Borders(xlEdgeLeft).LineStyle = xlNone
  Selection.Borders(xlEdgeTop).LineStyle = xlNone
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  Selection.Borders(xlEdgeRight).LineStyle = xlNone
  Selection.Borders(xlInsideVertical).LineStyle = xlNone
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  Selection.Font.Bold = False
  Selection.Font.Italic = False
  
  ' Add date and time stamp to column A
  ActiveCell.Offset(0, -1).Select
  ActiveCell.Value = Now
  
  ' Go to the Cumulative Stats Page
  Sheets("Cumulative Stats").Select
  Range("I13").Select
  
  ' Re-save the workbook before closing out
  ActiveWorkbook.Save
  
  ' Close out excel
  Application.Quit
  
End sub

' Refresh data on pivot table

Sub cmdRefreshData()
  
  ActiveWorkbook.Save
  ActiveWorkbook.RefreshAll
  ActiveWorkbook.Save
  
End Sub

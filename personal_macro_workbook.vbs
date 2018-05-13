Sub HighlightYellow()
'
' HighlightYellow Macro
'
' Keyboard Shortcut: Ctrl+h
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub HighlightNone()
'
' HighlightNone Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
'
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub SetupWorksheet()
'
' To setup a worksheet in my style.
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    With ActiveSheet
        .Cells.Font.Size = 10
        .Cells.Font.Name = "Arial"
        .Cells.VerticalAlignment = xlCenter
        .Rows.RowHeight = 12.5
    End With

    With ActiveSheet.Columns(1)
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 3
    End With

    With Range("A1")
        .Font.Size = 30
        .Font.Name = "Roboto"
        .Font.Bold = True
        .Rows.RowHeight = 38
        .Value = "TABLE NAME"
        .HorizontalAlignment = xlLeft
    End With

    With Range("A2")
        .Font.Name = "Roboto"
        .Rows.RowHeight = 13
        .Value = "Information about the table goes here."
        .HorizontalAlignment = xlLeft
    End With


    Range("A4").Value = "SN"
    Range("B4").Value = "KEY"
    Range("C4").Value = "VALUE"


    With ActiveSheet.Rows(4)
        .Interior.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    Range("A5:A14").Formula = "=row()-4"


End Sub

Sub SetupWorksheetAsTreatise()
'
' To setup a worksheet as a treatise.
'
' Keyboard Shortcut: Ctrl+Shift+O
'
    With ActiveSheet
        .Cells.Font.Size = 10
        .Cells.Font.Name = "Arial"
        .Cells.VerticalAlignment = xlCenter
        .Rows.RowHeight = 12.5
    End With

    With ActiveSheet.Columns(1)
        .HorizontalAlignment = xlCenter
    End With

    With ActiveSheet.Range(Columns(1), Columns(5))
        .ColumnWidth = 5
    End With

    With Range("A1")
        .Font.Size = 30
        .Font.Name = "Roboto"
        .Font.Bold = True
        .Rows.RowHeight = 38
        .Value = "TREATISE NAME"
        .HorizontalAlignment = xlLeft
    End With

    With Range("A2")
        .Font.Name = "Roboto"
        .Rows.RowHeight = 13
        .Value = "Information about the treatise goes here."
        .HorizontalAlignment = xlLeft
    End With


    Range("A4").Value = "SN"
    Range("B4").Value = "POINT"


    With ActiveSheet.Rows(4)
        .Interior.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    Range("A5:A14").Formula = "=row()-4"


End Sub
Sub GoToNote()
    If Not ActiveCell.Find("^", LookIn:=xlValues) Is Nothing Then
        If ActiveCell.Column() = 1 Then
            ActiveSheet.Cells.Find("^" & Right(ActiveCell, _
            Len(ActiveCell.Value) - InStr(1, ActiveCell.Value, "^")), _
            LookIn:=xlValues).Select
        Else
            ActiveSheet.Columns(1).Find("^" & Right(ActiveCell, _
            Len(ActiveCell.Value) - InStr(1, ActiveCell.Value, "^")), _
            LookIn:=xlValues).Select
        End If
    End If
End Sub
'
Sub SelectAbove()
'
' SelectAbove Macro
'
    toprow = Selection.Rows(1).row - 1
    leftcol = Selection.Columns(1).Column
    rightcol = Selection.Columns(Selection.Columns.Count).Column
    Range(Cells(1, leftcol).Address(), Cells(toprow, rightcol).Address()).Select
End Sub

Sub FillRandomly()

    If Not (TypeOf Application.Selection Is Range) Then
        MsgBox "The current selection is not a range."
        Exit Sub
    End If

    fill_randomly.Show

End Sub

Sub ReplaceUnreadable()
'
' ReplaceUnreadable Macro
' to fix formatting after webscrapping

'
    Cells.Replace What:="â€“", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="Â", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="Ã¡", Replacement:="á", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="€™", Replacement:="'", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="Ã©", Replacement:="é", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub


Sub PasteValues()
'
' PasteValues Macro
'
' Keyboard Shortcut: Ctrl+Shift+V
'
    Selection.Value = Selection.Value

End Sub



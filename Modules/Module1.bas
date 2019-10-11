Attribute VB_Name = "Module1"
Option Explicit

Sub CopyColumnInsertRight()

    Dim lString As String
    
    lString = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Address
    
    Range(lString).EntireColumn.Select
    
    Application.CutCopyMode = False
    
    Selection.Copy
    
    Selection.Insert Shift:=xlToRight
    ActiveCell.Offset(1, 1).Select
    
    Application.CutCopyMode = False
    
    ActiveSheet.UsedRange
    
End Sub

Sub NewRow(inInputMessage)
    Dim lShape As Shape
    Dim lRow As Integer
    Dim lNewRow As Integer
    Dim lColumn As Integer
    Dim lNewColumn As Integer
    Dim lCellHeight As Integer
    Dim lCellTop As Integer
    Dim lNewDataCell As Range
    
    'Empty clipboard
    Application.CutCopyMode = False

    'Find position of button pressed and destination cells
    Set lShape = ActiveSheet.Shapes(Application.Caller)
    lRow = lShape.TopLeftCell.Row
    lNewRow = lRow + 1
    lColumn = lShape.TopLeftCell.Column
    lNewColumn = lColumn - 1
    
    'Insert new row
    Rows(lNewRow).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Set lNewDataCell = Cells(lNewRow, lNewColumn)
    lCellHeight = lNewDataCell.Height
    lCellTop = lNewDataCell.Top
    lShape.Height = lCellHeight * 0.9
    lShape.Top = lNewDataCell.Top + 0.05 * lCellHeight

    'Copy validation to new cell from cell above
    Cells(lRow, lNewColumn).Copy
    lNewDataCell.Select
    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Change inputMessage
    lNewDataCell.Validation.InputMessage = inInputMessage
    
    ActiveSheet.UsedRange
End Sub

Sub NewStationRow()
    Const cInputMessage As String = "Her kan du legge inn ytterligere en stasjon hvis øvrig" _
        & " informasjon er felles med stasjonen(e) over."
    Call NewRow(cInputMessage)
End Sub

Sub NewDate()
    Const cInputMessage As String = "Her kan du legge inn ytterligere en prøvetakingsdato" _
        & " hvis øvrig informasjon er felles med datoen over."
    Call NewRow(cInputMessage)
End Sub

Sub NewDepth()
    Const cInputMessage As String = "Her kan du legge inn ytterligere et prøvetakingsdyp" _
    & " eller intervall hvis øvrig informasjon er felles med dypet over."
    Call NewRow(cInputMessage)
End Sub

Sub NewCore()
    Const cInputMessage As String = "Her kan du legge inn ytterligere en kjerne" _
    & " hvis øvrig informasjon er felles med kjernen over."
    Call NewRow(cInputMessage)
End Sub

Sub NewSlice()
    Const cInputMessage As String = "Her kan du legge inn ytterligere et snitt" _
    & " hvis øvrig informasjon er felles med snittet over."
    Call NewRow(cInputMessage)
End Sub

Sub NewSpecimen()
    Const cInputMessage As String = "Her kan du legge inn ytterligere et individ" _
    & " hvis øvrig informasjon er felles med individet over."
    Call NewRow(cInputMessage)
End Sub

Sub NewAnalys()
    Dim lShape As Shape
    Dim lRow As Integer
    Dim lNewRow As Integer
    Dim lCellHeight As Integer
    Dim lCellTop As Integer
    Dim lNewDataCell As Range
    Dim lCheckboxLeft As Integer
    Dim lCheckboxTop As Integer
    Dim lCheckboxWidth As Integer
    Dim lCheckboxHeight As Integer
    Dim lColumn As Integer
    Dim lNewColumn As Integer
    
    'Empty clipboard
    Application.CutCopyMode = False
    
    'Find position of button pressed and destination cells
    Set lShape = ActiveSheet.Shapes(Application.Caller)
    lRow = lShape.TopLeftCell.Row
    lNewRow = lRow + 1
    lColumn = lShape.TopLeftCell.Column
    lNewColumn = lColumn - 1
    
    'Insert new empty row
    Rows(lNewRow).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Set lNewDataCell = Cells(lNewRow, lNewColumn)

    'Copy validation to new cell from cell above
    Cells(lRow, lNewColumn).Copy
    lNewDataCell.Select
    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Change inputMessage
    lNewDataCell.Validation.InputMessage = "Her kan du legge inn ytterligere analyse" _
        & " hvis øvrig informasjon er felles med analysen over."
    
    'Create check-box
    lCellHeight = lNewDataCell.Height
    lCellTop = lNewDataCell.Top
    lShape.Height = lCellHeight * 0.9
    lShape.Top = lNewDataCell.Top + 0.05 * lCellHeight
    lCheckboxWidth = 24
    lCheckboxLeft = lNewDataCell.Left + (lNewDataCell.Offset.Width - lCheckboxWidth) / 2
    lCheckboxHeight = 20
    lCheckboxTop = lNewDataCell.Top + (lNewDataCell.Height - lCheckboxHeight) / 2
    ActiveSheet.CheckBoxes.Add(lCheckboxLeft, lCheckboxTop, lCheckboxWidth, lCheckboxHeight).Select
    Selection.Characters.Text = ""
    
    Cells(lNewDataCell.Row, 1).Select
    
    Application.CutCopyMode = False
    ActiveSheet.UsedRange
  
End Sub

Sub listButtons()

    Dim LCurrentWorksheet As Worksheet
    Dim LCurrentShape As Shape
    
    Set LCurrentWorksheet = ActiveSheet
    
    For Each LCurrentShape In LCurrentWorksheet.Shapes
    If LCurrentShape.Type = 8 Then
        Debug.Print "Name: " & LCurrentShape.Name
        Debug.Print "Row: " & LCurrentShape.TopLeftCell.Row
        Debug.Print "Type: " & LCurrentShape.Type
        Debug.Print "Connector: " & LCurrentShape.ID
        Debug.Print "ID: " & LCurrentShape.ID
        Debug.Print "Alternative text: " & LCurrentShape.AlternativeText
        
        End If
    
    Next LCurrentShape

End Sub


Sub resizeAlignButtons()
'Brukes for å midtstille knapper i cellen, og lage størrelsen 80% av cellen de er i
'Brukes kun i tomme ark uten ekstra-kolonner

    Dim LCurrentWorksheet As Worksheet
    Dim LCurrentShape As Shape
    Dim lCellHeight As Integer
    Dim lCellTop As Integer
    Dim lCell As Range
    Dim lTopRow As Integer
    Dim lLeftCol As Integer
    Dim lBottomRow As Integer
    Dim lRightCol As Integer
    Dim lCol As Integer
    Dim lRow As Integer
    Dim lCellWidth As Integer
    Dim lCellLeft As Integer
    
    Set LCurrentWorksheet = ActiveSheet
    
    For Each LCurrentShape In LCurrentWorksheet.Shapes
        If LCurrentShape.Type = 8 Then
        
            lTopRow = LCurrentShape.TopLeftCell.Row
            lBottomRow = LCurrentShape.BottomRightCell.Row
            lLeftCol = LCurrentShape.TopLeftCell.Column
            lRightCol = LCurrentShape.BottomRightCell.Column
            
            lCol = (lLeftCol + lRightCol) \ 2
            lRow = (lTopRow + lBottomRow) \ 2
            
            Set lCell = Cells(lRow, lCol)
            
            lCellLeft = lCell.Left
            lCellWidth = lCell.Width
            lCellTop = lCell.Top
            lCellHeight = lCell.Height
            
            If LCurrentShape.FormControlType = 0 Then
                LCurrentShape.Height = lCellHeight * 0.8
                LCurrentShape.Width = lCellWidth * 0.8
            End If
            
            LCurrentShape.Top = lCellTop + (lCellHeight - LCurrentShape.Height) / 2
            LCurrentShape.Left = lCellLeft + (lCellWidth - LCurrentShape.Width) / 2
            
        End If
    
    Next LCurrentShape

End Sub


Public Function getValidationErrorMessage(inCell As Range) As String
    getValidationErrorMessage = inCell.Validation.ErrorMessage
End Function


Public Function getValidationInputMessage(inCell As Range)
    getValidationInputMessage = inCell.Validation.InputMessage
End Function



Attribute VB_Name = "Module1"
Option Explicit

Function findCell(inWorksheet As Worksheet, inValue As Variant) As Range
    ' Returns range object of first cell with value = inValue
    Dim lCurrentCell
    
    For Each lCurrentCell In inWorksheet.UsedRange.Cells
        If lCurrentCell.Value = inValue Then
            Set findCell = lCurrentCell
            Exit Function
        End If
    Next
End Function


Sub unhideAllRowsAndFormControls()
    
    Dim LCurrentWorksheet As Worksheet
    Dim lCurrentShape As Shape
    
    Application.ScreenUpdating = False
    ActiveSheet.UsedRange
    
    Set LCurrentWorksheet = ActiveSheet
    
    ' Unhide all rows
    Cells.Select
    Selection.EntireRow.Hidden = False
    
    Range("A1").Select
    
    ' Unhide form controls
    For Each lCurrentShape In LCurrentWorksheet.Shapes
      
      lCurrentShape.Visible = msoTrue
    
    Next lCurrentShape
    
  Application.ScreenUpdating = True
  
End Sub


Sub hideUncheckedAnalyses()
    
    Dim LCurrentWorksheet As Worksheet
    Dim lCurrentShape As Shape
    Dim lStartRow As Integer
    Dim lCurrentRow As Integer
    Dim lCurrentColumn As Integer
    Dim lStartColumn As Integer
    Dim lLastRow
    Dim lLastColumn
    Dim lHideRow As Boolean
    
    Application.ScreenUpdating = False
    ActiveSheet.UsedRange
    
    Set LCurrentWorksheet = ActiveSheet
    
    lStartRow = findCell(Worksheets("Analyserekvisisjon ferskvann"), "Ønskede analyser listes nedenfor").Row + 1
    lStartColumn = lStartRow = findCell(Worksheets("Analyserekvisisjon ferskvann"), "Ønskede analyser listes nedenfor").Column + 1
    
    
    lLastRow = LCurrentWorksheet.UsedRange.Rows.Count - 2
    lLastColumn = LCurrentWorksheet.UsedRange.Columns.Count
    
    'Loop through all rows with analysismethods
    For lCurrentRow = lStartRow To lLastRow
    
        ' Loop through all cells/columns in each row. If one or more analyses are checked,
        ' the row should not get hidden
        
        lHideRow = True
        
        For lCurrentColumn = lStartColumn To lLastColumn + 1
        
            If isChecked(LCurrentWorksheet, lCurrentRow, lCurrentColumn) = xlOn Then
                lHideRow = False
                Exit For
            End If
            
        Next lCurrentColumn
        
        If lHideRow = True Then
            ' Hide checkboxes, they are not hidden with the row, when the row is hidden
            hideCheckBoxesInRow LCurrentWorksheet, lCurrentRow
            
            ' Hide entire cell row
            Cells(lCurrentRow, lCurrentColumn).EntireRow.Hidden = True
        End If
        
    Next lCurrentRow
    
  Application.ScreenUpdating = True
  
End Sub


Private Function isChecked(inWorksheet As Worksheet, inRow As Integer, inColumn As Integer)

    Dim lCurrentShape As Shape
    Dim lTopRow As Integer
    Dim lLeftCol As Integer
    Dim lBottomRow As Integer
    Dim lRightCol As Integer
    Dim lCol As Integer
    Dim lRow As Integer

    For Each lCurrentShape In inWorksheet.Shapes

        If lCurrentShape.Type = msoFormControl Then
            
            If lCurrentShape.FormControlType = xlCheckBox Then
            
              ' Find position of current checkbox
              lTopRow = lCurrentShape.TopLeftCell.Row
              lBottomRow = lCurrentShape.BottomRightCell.Row
              lLeftCol = lCurrentShape.TopLeftCell.Column
              lRightCol = lCurrentShape.BottomRightCell.Column
              
              ' Find middle reference of checkbox
              lCol = (lLeftCol + lRightCol) \ 2
              lRow = (lTopRow + lBottomRow) \ 2
              
              ' Test if current checkbox has same reference as input
              If lRow = inRow And lCol = inColumn Then
                
                isChecked = lCurrentShape.ControlFormat.Value
                
                Exit Function
                
              End If
              
            End If
            
        End If
        
    Next lCurrentShape
    
    isChecked = xlOff
    
End Function


Private Sub hideCheckBoxesInRow(inWorksheet As Worksheet, inRow As Integer)
    
    Dim lCurrentShape As Shape
    Dim lTopRow As Integer
    Dim lBottomRow As Integer
    Dim lRightCol As Integer
    Dim lRow As Integer

    For Each lCurrentShape In inWorksheet.Shapes

        If lCurrentShape.Type = msoFormControl Then
            
            If lCurrentShape.FormControlType = xlCheckBox Then
              
              ' Find reference for center of current checkbox
              lTopRow = lCurrentShape.TopLeftCell.Row
              lBottomRow = lCurrentShape.BottomRightCell.Row

              lRow = (lTopRow + lBottomRow) \ 2
              
              If lRow = inRow Then
                
                lCurrentShape.Visible = msoFalse
          
              End If
              
            End If
        
        End If
    
    Next lCurrentShape
    
End Sub


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
    Range(Cells(lRow, 1), Cells(lRow, lNewColumn)).Copy
    lNewDataCell.Select
    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    'Copy format to new row from cell above
    Range(Cells(lNewRow, 1), lNewDataCell).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    'Change inputMessage
    Cells(lNewRow, 1).Validation.InputMessage = "Her kan du legge inn ytterligere analyse" _
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
    Selection.Value = True
    
    Cells(lNewDataCell.Row, 1).Select
    
    Application.CutCopyMode = False
    ActiveSheet.UsedRange
  
End Sub

Sub listButtons()

    Dim LCurrentWorksheet As Worksheet
    Dim lCurrentShape As Shape
    
    Set LCurrentWorksheet = ActiveSheet
    
    For Each lCurrentShape In LCurrentWorksheet.Shapes
    If lCurrentShape.Type = msoFormControl Then
        Debug.Print "Name: " & lCurrentShape.name
        Debug.Print "Row: " & lCurrentShape.TopLeftCell.Row
        Debug.Print "Type: " & lCurrentShape.Type
        Debug.Print "Connector: " & lCurrentShape.ID
        Debug.Print "ID: " & lCurrentShape.ID
        Debug.Print "Alternative text: " & lCurrentShape.AlternativeText
        
        End If
    
    Next lCurrentShape

End Sub


Sub resizeAlignButtons()
'Brukes for å midtstille form controls i cellene, og lage størrelsen 80% av cellen de er i
'Brukes kun i tomme ark uten ekstra-kolonner

    Dim LCurrentWorksheet As Worksheet
    Dim lCurrentShape As Shape
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
    
    ActiveSheet.UsedRange
    Set LCurrentWorksheet = ActiveSheet
    
    For Each lCurrentShape In LCurrentWorksheet.Shapes
        If lCurrentShape.Type = msoFormControl Then
                     Debug.Print (lCurrentShape.Type & ", " & lCurrentShape.name & ", " & lCurrentShape.TopLeftCell.Row & ", " & lCurrentShape.BottomRightCell.Row)
              
          If lCurrentShape.FormControlType = xlCheckBox Or lCurrentShape.FormControlType = xlButtonControl Then
  
              lTopRow = lCurrentShape.TopLeftCell.Row
              lBottomRow = lCurrentShape.BottomRightCell.Row
              lLeftCol = lCurrentShape.TopLeftCell.Column
              lRightCol = lCurrentShape.BottomRightCell.Column
              
              lCol = (lLeftCol + lRightCol) \ 2
              lRow = (lTopRow + lBottomRow) \ 2
              
              Set lCell = Cells(lRow, lCol)
              
              lCellLeft = lCell.Left
              lCellWidth = lCell.Width
              lCellTop = lCell.Top
              lCellHeight = lCell.Height
              
              If lCurrentShape.FormControlType = xlButtonControl Then
                  lCurrentShape.Height = lCellHeight * 0.8
                  lCurrentShape.Width = lCellWidth * 0.8
              End If
              
              lCurrentShape.Top = lCellTop + (lCellHeight - lCurrentShape.Height) / 2
              lCurrentShape.Left = lCellLeft + (lCellWidth - lCurrentShape.Width) / 2
              
             'Debug.Print (lCurrentShape.Type & ", " & lCurrentShape.name & ", " & lCurrentShape.TopLeftCell.Row & ", " & lCurrentShape.BottomRightCell.Row)
              
          End If
        End If
    Next lCurrentShape

End Sub


Sub resizeRightAlignButtons()
'Brukes for å midtstille fom controls i cellene, og lage størrelsen 80% av cellen de er i
'Brukes i prosjektinfo-arket

    Dim LCurrentWorksheet As Worksheet
    Dim lCurrentShape As Shape
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
    
    ActiveSheet.UsedRange
    Set LCurrentWorksheet = ActiveSheet
    
    For Each lCurrentShape In LCurrentWorksheet.Shapes
        If lCurrentShape.Type = msoFormControl Then
          If lCurrentShape.FormControlType = xlCheckBox Or lCurrentShape.FormControlType = xlButtonControl Then
              lTopRow = lCurrentShape.TopLeftCell.Row
              lBottomRow = lCurrentShape.BottomRightCell.Row
              lLeftCol = lCurrentShape.TopLeftCell.Column
              lRightCol = lCurrentShape.BottomRightCell.Column
              
              lCol = (lLeftCol + lRightCol) \ 2
              lRow = (lTopRow + lBottomRow) \ 2
              
              Set lCell = Cells(lRow, lCol)
              
              lCellLeft = lCell.Left
              lCellWidth = lCell.Width
              lCellTop = lCell.Top
              lCellHeight = lCell.Height
              
              If lCurrentShape.FormControlType = xlButtonControl Then
                  lCurrentShape.Height = lCellHeight * 0.8
                  lCurrentShape.Width = lCellWidth * 0.8
              End If
              
              lCurrentShape.Top = lCellTop + (lCellHeight - lCurrentShape.Height) / 2
              'LCurrentShape.Left = lCellLeft + (lCellWidth - LCurrentShape.Width) / 2
              lCurrentShape.Left = lCellLeft + lCellWidth * 0.9 - lCurrentShape.Width / 2
              
          End If
        End If
    Next lCurrentShape

End Sub


Public Function getValidationErrorMessage(inCell As Range) As String
    getValidationErrorMessage = inCell.Validation.ErrorMessage
End Function


Public Function getValidationInputMessage(inCell As Range)
    getValidationInputMessage = inCell.Validation.InputMessage
End Function





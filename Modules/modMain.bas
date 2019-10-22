Attribute VB_Name = "modMain"
Option Explicit


Private Function findCell(inWorksheet As Worksheet, inValue As Variant) As Range
    ' Returns range of first cell with value = inValue
    
    Dim lCurrentCell
    
    For Each lCurrentCell In inWorksheet.UsedRange.Cells
        If lCurrentCell.Value = inValue Then
            Set findCell = lCurrentCell
            Exit Function
        End If
    Next
End Function


Sub unhideAllRowsAndFormControls()
    
    Dim lCurrentWorksheet As Worksheet
    Dim lCurrentShape As Shape
    
    Application.ScreenUpdating = False
    ActiveSheet.UsedRange
    
    Set lCurrentWorksheet = ActiveSheet
    
    ' Unhide all rows
    Cells.Select
    Selection.EntireRow.Hidden = False
    
    Range("A1").Select
    
    ' Unhide form controls
    For Each lCurrentShape In lCurrentWorksheet.Shapes
      
      lCurrentShape.Visible = msoTrue
    
    Next lCurrentShape
    
  Application.ScreenUpdating = True
  
End Sub


Sub hideUncheckedAnalyses()
    
    Dim lCurrentWorksheet As Worksheet
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
    
    Set lCurrentWorksheet = ActiveSheet
    
    
    lStartRow = findCell(Worksheets("Analyserekvisisjon ferskvann"), "Analyser:").Row + 1
    lStartColumn = lStartRow = findCell(Worksheets("Analyserekvisisjon ferskvann"), "Analyser:").Column + 1
    
    
    lLastRow = lCurrentWorksheet.UsedRange.Rows.Count - 2
    lLastColumn = lCurrentWorksheet.UsedRange.Columns.Count
    
    'Loop through all rows with analysismethods
    For lCurrentRow = lStartRow To lLastRow
    
        ' Loop through all cells/columns in each row. If one or more analyses are checked,
        ' the row should not get hidden
        
        lHideRow = True
        
        For lCurrentColumn = lStartColumn To lLastColumn + 1
        
            If isChecked(lCurrentWorksheet, lCurrentRow, lCurrentColumn) = xlOn Then
                lHideRow = False
                Exit For
            End If
            
        Next lCurrentColumn
        
        If lHideRow = True Then
            ' Hide checkboxes, they are not hidden with the row, when the row is hidden
            hideCheckBoxesInRow lCurrentWorksheet, lCurrentRow
            
            ' Hide entire cell row
            Cells(lCurrentRow, lCurrentColumn).EntireRow.Hidden = True
        End If
        
    Next lCurrentRow
    
  Application.ScreenUpdating = True
  
End Sub


Function isChecked(inWorksheet As Worksheet, inRow As Integer, inColumn As Integer)

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


Private Sub NewRow(inCellContainingValidation As Range)
    
    Dim lShape As Shape
    Dim lRow As Integer
    Dim lNewRow As Integer
    Dim lColumn As Integer
    Dim lNewColumn As Integer
    Dim lCellHeight As Integer
    Dim lCellTop As Integer
    Dim lNewDataCell As Range
    
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
    
    'Copy Data Validation from DataValidations sheet
    inCellContainingValidation.Copy
    
    lNewDataCell.Select
    
    Selection.PasteSpecial Paste:=xlPasteValidation
    
    ActiveSheet.UsedRange
End Sub


Sub NewStationRow()
    Dim lCellCopyValidation As Range
    
    Set lCellCopyValidation = findCell(Worksheets("DataValidations"), "Stasjonskode og navn").Offset(columnoffset:=1)

    Call NewRow(lCellCopyValidation)
End Sub


Sub NewDate()
    Dim lCellCopyValidation As Range
    
    Set lCellCopyValidation = findCell(Worksheets("DataValidations"), "Prøvetakingsdato").Offset(columnoffset:=1)

    Call NewRow(lCellCopyValidation)
End Sub


Sub NewDepth()
    Dim lCellCopyValidation As Range
    
    Set lCellCopyValidation = findCell(Worksheets("DataValidations"), "Prøvetakingsdyp").Offset(columnoffset:=1)

    Call NewRow(lCellCopyValidation)
End Sub


Sub NewCore()
    Dim lCellCopyValidation As Range
    
    Set lCellCopyValidation = findCell(Worksheets("DataValidations"), "Kjerneidentifikasjon/grabb").Offset(columnoffset:=1)

    Call NewRow(lCellCopyValidation)
End Sub


Sub NewSlice()
    Dim lCellCopyValidation As Range
    
    Set lCellCopyValidation = findCell(Worksheets("DataValidations"), "Snitt").Offset(columnoffset:=1)

    Call NewRow(lCellCopyValidation)
End Sub


Sub NewSpecimen()
    Dim lCellCopyValidation As Range
    
    Set lCellCopyValidation = findCell(Worksheets("DataValidations"), "Individnummer/prøvenr").Offset(columnoffset:=1)

    Call NewRow(lCellCopyValidation)
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
    
    'Copy validation to new row from row above
    Range(Cells(lRow, 1), Cells(lRow, lNewColumn)).Copy
    Range(Cells(lNewRow, 1), Cells(lNewRow, lNewColumn)).Select
    Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    'Copy format to new row from row above
    Range(Cells(lNewRow, 1), Cells(lNewRow, lNewColumn)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
           
    ' Insert check boxes for entire new row
    createCheckBoxes Range(Cells(lNewRow, 2), Cells(lNewRow, lNewColumn))
    
    ' Move New Analysis button down
    lCellHeight = lNewDataCell.Height
    lShape.Height = lCellHeight * 0.9
    lShape.Top = lNewDataCell.Top + 0.05 * lCellHeight
    
    lNewDataCell.Select
    
    Application.CutCopyMode = False
    ActiveSheet.UsedRange
  
End Sub


Private Sub createCheckbox(inCell As Range)
    ' Create check-box in cell and check it
    ' Used when creating new row in anaylses sheets
    
    Dim lShape As Shape
    Dim lRow As Integer
    Dim lNewRow As Integer
    Dim lCellHeight As Integer
    Dim lCellTop As Integer
    Dim lCheckboxLeft As Integer
    Dim lCheckboxTop As Integer
    Dim lCheckboxWidth As Integer
    Dim lCheckboxHeight As Integer

    lCellHeight = inCell.Height
    lCellTop = inCell.Top
    lCheckboxWidth = 24
    lCheckboxLeft = inCell.Left + (inCell.Offset.Width - lCheckboxWidth) / 2
    lCheckboxHeight = 20
    lCheckboxTop = inCell.Top + (inCell.Height - lCheckboxHeight) / 2
    
    ' Check if checkbox already exists
    If Not isChecked(inCell.Worksheet, inCell.Row, inCell.Column) = vbEmpty Then
        Exit Sub
    End If
    
    ' Create check box
    ActiveSheet.CheckBoxes.Add(lCheckboxLeft, lCheckboxTop, lCheckboxWidth, lCheckboxHeight).Select
    
    Selection.Characters.Text = ""
    
    ' Precheck chekbox
    Selection.Value = True
End Sub


Private Sub createCheckBoxes(inRow As Range)
    ' Used when creating new row in anaylses sheets
    Dim lCurrentCell As Range
    
    For Each lCurrentCell In inRow
        createCheckbox lCurrentCell
    Next

End Sub


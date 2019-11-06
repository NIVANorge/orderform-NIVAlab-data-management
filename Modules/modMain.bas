Attribute VB_Name = "modMain"
Option Explicit


Private Function findCell(inWorksheet As Worksheet, inValue As Variant) As Range
  ' Returns range of first cell with value = inValue in inWorksheet
  
  Dim lCurrCell
  
  For Each lCurrCell In inWorksheet.UsedRange.Cells
      
    If lCurrCell.Value = inValue Then
        
      Set findCell = lCurrCell
      
      Exit Function
      
    End If
  
  Next
  
End Function


Sub unhideAllRowsAndFormControls()
  ' This procedure shows all rows, checkboxes and other shapes
  ' when pressing "Vis alle" in Analyser-section of active
  ' worksheet
  
  Dim lCurrSheet As Worksheet
  Dim lCurrShape As Shape
  Dim initialScreenUpdating As Boolean
  
  initialScreenUpdating = Application.ScreenUpdating
  
  Application.ScreenUpdating = False
  
  ' Reset protection
  protect_wb_and_ws
  
  ActiveSheet.UsedRange
  
  Set lCurrSheet = ActiveSheet
  
  ' Unhide all rows
  Cells.Select
  Selection.EntireRow.Hidden = False
  
  lCurrSheet.Range("B2").Select
  
  ' Unhide form controls
  For Each lCurrShape In lCurrSheet.Shapes
    
    lCurrShape.Visible = msoTrue
  
  Next lCurrShape
  
  ActiveSheet.GroupObjects("OptionButtonsShowAllVsSelected").ShapeRange.GroupItems("optnBtnShowAll").ControlFormat.Value = xlOn
  
  Application.ScreenUpdating = initialScreenUpdating
  
End Sub


Sub hideUncheckedAnalyses()
    ' This procedure hides all rows and corresponding checkboxes
    ' when pressing "Vis kun avkryssede" in Analyser-section of active
    ' worksheet.
    
    Dim lCurrSheet As Worksheet
    Dim lStartRow As Integer
    Dim lCurrRow As Integer
    Dim lCurrColumn As Integer
    Dim lStartColumn As Integer
    Dim lLastRow As Integer
    Dim lLastColumn As Integer
    Dim lHideRow As Boolean
    Dim lCurrCell As Range
    
    Dim initialScreenUpdating As Boolean
  
    initialScreenUpdating = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    
    ' Reset protection
    protect_wb_and_ws
    
    ActiveSheet.UsedRange
    
    Set lCurrSheet = ActiveSheet
    
    ' Find "Analyser"-cell, only the rows beneath are going to be hidden
    lStartRow = findCell(ActiveSheet, "Analyser:").Row + 1
    lStartColumn = findCell(ActiveSheet, "Analyser:").Column + 1
    
    lLastRow = lCurrSheet.UsedRange.Rows.Count - 2
    lLastColumn = lCurrSheet.UsedRange.Columns.Count
    
    'Loop through all rows with analysismethods
    For lCurrRow = lStartRow To lLastRow
    
        ' Loop through all cells/columns in each row. If one or more analyses are checked,
        ' the row should not become hidden
        
        lHideRow = True
        
        For lCurrColumn = lStartColumn To lLastColumn
        
          Set lCurrCell = lCurrSheet.Cells(lCurrRow, lCurrColumn)
    
          If lCurrCell = True Then
          
            lHideRow = False
      
            Exit For
          
          End If
                        
        Next lCurrColumn
        
        If lHideRow Then
            ' Hide checkboxes, they are not hidden with the row, when the row is hidden
            hideCheckBoxesInRow lCurrSheet, lCurrRow
            
            ' Hide entire cell row
            Cells(lCurrRow, lCurrColumn).EntireRow.Hidden = True
        End If
        
    Next lCurrRow
    
    ActiveSheet.GroupObjects("OptionButtonsShowAllVsSelected").ShapeRange.GroupItems("optnBtnShowCheckedOnly").ControlFormat.Value = xlOn
    
  Application.ScreenUpdating = initialScreenUpdating
  
End Sub


Private Function linkedCheckBoxExists(inWorksheet As Worksheet, inAddress As String) As Boolean

  Dim lCurrCheckBox As CheckBox

  For Each lCurrCheckBox In inWorksheet.CheckBoxes

    If lCurrCheckBox.LinkedCell = inAddress Then
      
      linkedCheckBoxExists = True
      
      Exit Function
    
    End If

  Next lCurrCheckBox
  
  linkedCheckBoxExists = False
  
End Function


Private Sub hideCheckBoxesInRow(inWorksheet As Worksheet, inRow As Integer)
    
  Dim lCurrCheckBox As CheckBox
  Dim lCurrCheckBoxRow As Integer

  For Each lCurrCheckBox In inWorksheet.CheckBoxes
    lCurrCheckBoxRow = Split(lCurrCheckBox.LinkedCell, "$")(2)

    If lCurrCheckBoxRow = inRow Then
      
      lCurrCheckBox.Visible = False

    End If
               
  Next lCurrCheckBox
 
End Sub


Sub newAnalysColumn()

  Dim lFromCol As Range
  Dim lFromCheckBox As CheckBox
  Dim lNewCheckBox As CheckBox
  Dim lCopyButton As Shape
  Dim lFromColNo As Integer
  Dim lNewColNo As Integer
  Dim lNewCheckBoxLinkedCell As String
  Dim lShowCheckedOnlyInitialState As Integer

  Application.ScreenUpdating = False
    
  ' Reset protection
  protect_wb_and_ws
    
  ActiveSheet.UsedRange

  ' Unhide rows, the procedure wont work if cell rows are hidden
  ' First, check if rows are hidden, and keep track of state
  
  lShowCheckedOnlyInitialState = _
    ActiveSheet.GroupObjects("OptionButtonsShowAllVsSelected").ShapeRange.GroupItems("optnBtnShowCheckedOnly").ControlFormat.Value
    
  If Not lShowCheckedOnlyInitialState = xlOff Then
    unhideAllRowsAndFormControls
  End If
  
  Application.ScreenUpdating = False

  ' Insert new column to the right and copy all cell values and properties _
    from old column, except shape objects ("Kopier..."-button and checboxes)

  Set lCopyButton = ActiveSheet.Shapes(Application.Caller)
  Set lFromCol = lCopyButton.TopLeftCell.EntireColumn

  lFromCol.Offset(0, 1).EntireColumn.Select

  Application.CutCopyMode = False

  Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

  lFromCol.Select

  Selection.Copy

  lFromCol.Offset(0, 1).EntireColumn.Select

  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

  Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
      xlNone, SkipBlanks:=False, Transpose:=False

  Application.CutCopyMode = False

  ' Insert check boxes
  lFromColNo = lFromCol.Column

  lNewColNo = lFromColNo + 1

  For Each lFromCheckBox In ActiveSheet.CheckBoxes

    If Range(lFromCheckBox.LinkedCell).Column = lFromColNo Then

     lNewCheckBoxLinkedCell = Range(lFromCheckBox.LinkedCell).Offset(0, 1).Address

     createCheckbox (ActiveSheet.Range(lNewCheckBoxLinkedCell))

    End If

  Next lFromCheckBox

  ' Move "Kopier..."-button
  lCopyButton.Left = lCopyButton.Left + lFromCol.Width

  ActiveSheet.UsedRange

  If lShowCheckedOnlyInitialState = xlOn Then
    hideUncheckedAnalyses
  End If
  
  lCopyButton.TopLeftCell.Select
  
  Application.ScreenUpdating = True

End Sub


Private Sub NewRow(inCellContainingValidation As Range)
    
  Dim lshape As Shape
  Dim lRow As Integer
  Dim lNewRow As Integer
  Dim lColumn As Integer
  Dim lNewColumn As Integer
  Dim lCellHeight As Integer
  Dim lCellTop As Integer
  Dim lNewDataCell As Range

  Application.ScreenUpdating = False
    
  ' Reset protection
  protect_wb_and_ws
    
  ActiveSheet.UsedRange
    
  'Find position of button pressed and destination cells
  Set lshape = ActiveSheet.Shapes(Application.Caller)
  lRow = lshape.TopLeftCell.Row
  lNewRow = lRow + 1
  lColumn = lshape.TopLeftCell.Column
  lNewColumn = lColumn - 1
  
  'Insert new row
  Rows(lNewRow).Select
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  
  Set lNewDataCell = Cells(lNewRow, lNewColumn)
  lCellHeight = lNewDataCell.Height
  lCellTop = lNewDataCell.Top
  lshape.Height = lCellHeight * 0.9
  lshape.Top = lNewDataCell.Top + 0.05 * lCellHeight
  
  'Copy Data Validation from DataValidations sheet
  inCellContainingValidation.Copy
  
  lNewDataCell.Select
  
  Selection.PasteSpecial Paste:=xlPasteValidation

  Application.ScreenUpdating = True
  
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


Sub NewAnalysRow()
  Dim lshape As Shape
  Dim lRow As Integer
  Dim lNewRow As Integer
  Dim lCellHeight As Integer
  Dim lNewDataCell As Range
  Dim lColumn As Integer
  Dim lNewColumn As Integer
    
  Application.ScreenUpdating = False
    
  ' Reset protection
  protect_wb_and_ws
    
  ActiveSheet.UsedRange

  'Empty clipboard
  Application.CutCopyMode = False
  
  'Find position of button pressed and destination cells
  Set lshape = ActiveSheet.Shapes(Application.Caller)
  lRow = lshape.TopLeftCell.Row
  lNewRow = lRow + 1
  lColumn = lshape.TopLeftCell.Column
  lNewColumn = lColumn - 1
  
  'Insert new empty row
  Rows(lNewRow).Select
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  
  Set lNewDataCell = Cells(lNewRow, lNewColumn)
  
  'Copy validation to new row from row above
  Application.CutCopyMode = False
  
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
  
  ' Set new datacell = true - checkbox will be cheked
  lNewDataCell = True
  
  ' Move New Analysis button down
  lCellHeight = lNewDataCell.Height
  lshape.Height = lCellHeight * 0.9
  lshape.Top = lNewDataCell.Top + 0.05 * lCellHeight
  
  ActiveSheet.Cells(lNewDataCell.Row, 1).Select
  
  Application.CutCopyMode = False
  
  ActiveSheet.UsedRange
  
  Application.ScreenUpdating = True
    
End Sub


Private Sub createCheckbox(inCell As Range)
  ' Create check-box in cell and check it
  ' Used when creating new row in anaylses sheets
  
  Dim lcheckbox As CheckBox
  Const cCheckboxSize As Double = 12.75
  Const cTopOffset = 1.5
  
  ' Check if checkbox already exists
  If linkedCheckBoxExists(inCell.Worksheet, inCell.Address) = True Then
     
    Exit Sub
  
  End If
  
  ' Create check box
  Set lcheckbox = inCell.Worksheet.CheckBoxes.Add( _
  Top:=inCell.Top + cTopOffset, Left:=inCell.Left, _
  Height:=cCheckboxSize, Width:=cCheckboxSize)

  With lcheckbox
    
    If inCell.Height = 0 Then
      .Visible = False
    End If
    
    .Height = cCheckboxSize
    .Width = cCheckboxSize
    .Value = inCell.Value
    .LinkedCell = inCell.Address
    .Caption = ""
    .Top = inCell.Top + cTopOffset
    .Left = inCell.Left + (inCell.Width - .Width) / 2
    
  End With
    
End Sub


Private Sub createCheckBoxes(inRow As Range)
  ' Used when creating new row in anaylses sheets
  Dim lCurrCell As Range
  
  For Each lCurrCell In inRow
      createCheckbox lCurrCell
  Next

End Sub


Sub protect_wb_and_ws()
  ' Activate or reset protection of workbook and worksheets
  ' The setting UserInterfaceOnly:=True seems not to persist after reopening workbook,
  ' therefore it should be reset before any manipulating of worksheet with vba
  
  Dim lCurrentWorksheet As Worksheet

  ActiveWorkbook.Protect "encrypted"
  
  For Each lCurrentWorksheet In ActiveWorkbook.Worksheets

    lCurrentWorksheet.Protect "encrypted", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True

    lCurrentWorksheet.EnableSelection = xlNoRestrictions

  Next lCurrentWorksheet

  ActiveWorkbook.Protect "encrypted"
    
End Sub


Sub unprotect_wb_and_ws()
    
  ' Remove all protection of workbook and worksheet
  ' Used before maiplualting worksheets in gui
    
  Dim lCurrentWorksheet As Worksheet
  
  For Each lCurrentWorksheet In ActiveWorkbook.Worksheets
  
    lCurrentWorksheet.Unprotect "encrypted"
      
  Next lCurrentWorksheet
  
  ActiveWorkbook.Unprotect "encrypted"
    
End Sub

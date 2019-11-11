Attribute VB_Name = "modDevTools"
Option Explicit

' Procedures and functions for developer use

Private Sub resizeAlignButtons()
'Brukes for å midtstille form controls i cellene, og lage størrelsen 80% av cellen de er i
'Brukes kun i tomme ark uten ekstra-kolonner

    Dim lCurrentWorksheet As Worksheet
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
    Const cTopOffset = 1.5
    
    ActiveSheet.UsedRange
    Set lCurrentWorksheet = ActiveSheet
    
    For Each lCurrentShape In lCurrentWorksheet.Shapes
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
              
              If lCurrentShape.FormControlType = xlCheckBox Then
                lCurrentShape.Top = lCellTop + cTopOffset
              Else
                lCurrentShape.Top = lCellTop + (lCellHeight - lCurrentShape.Height) / 2
              End If
              
              lCurrentShape.Left = lCellLeft + (lCellWidth - lCurrentShape.Width) / 2
              
          End If
        End If
    Next lCurrentShape

End Sub

Private Sub repopulateAnalysCheckboxesActiveSheet()
  Application.ScreenUpdating = False

  repopulateAnalysCheckboxes ActiveSheet
  
  Application.ScreenUpdating = True
End Sub


Private Sub repopulateAnalysCheckboxes(inSheet As Worksheet)
' Delete all checkboxes and put on new ones
' New ones are generated for cells where validation input message starts with "Kryss av"
    
    Dim lCurrCell As Range
    Dim lCurrCheckBox As CheckBox
    Const cCheckboxSize As Double = 12.75
  
    inSheet.Activate
    ActiveSheet.UsedRange
  
    ' First, delete all checkboxes:
    inSheet.CheckBoxes.Delete

    ' Insert new checkboxes
    ' Insert checkbox for each cell where validation text starts with "Kryss av"
    For Each lCurrCell In inSheet.Cells.SpecialCells(xlCellTypeAllValidation)
      
      If lCurrCell.Validation.InputMessage Like "Kryss av*" Then
          
        Set lCurrCheckBox = inSheet.CheckBoxes.Add(Top:=lCurrCell.Top _
          , Left:=lCurrCell.Left, Height:=cCheckboxSize, Width:=cCheckboxSize)

        lCurrCheckBox.Value = False
        lCurrCheckBox.LinkedCell = lCurrCell.Address
        lCurrCheckBox.Caption = ""
        
        ' also set fontcolor of cell same as background
        ' We dont want to see the False/true value in the underlying cell
        
        lCurrCell.Font.ColorIndex = lCurrCell.Interior.ColorIndex
        
      End If
  
    Next lCurrCell
    
    centerCheckBoxes inSheet
    
End Sub

Private Sub DeleteOptnBtnGrp()

    Dim lGroup As Object
    Dim lOptnBtn As OptionButton
    
    For Each lGroup In ActiveSheet.GroupObjects
      lGroup.Delete
    Next lGroup
    
    For Each lOptnBtn In ActiveSheet.OptionButtons
      lOptnBtn.Delete
    Next lOptnBtn
    
End Sub
    
    
Private Sub ReCreateOptnBtnGrp()

    Dim lCell As Range
    Dim lGroup As Object
    Dim lOptnShowChckd As OptionButton
    Dim lOptnShowAll As OptionButton
    Dim lOptnShowChckdName As String
    Dim lOptnShowAllName As String

    DeleteOptnBtnGrp
    
    Set lCell = findCell(ActiveSheet, "Analyser:")
    
    lOptnShowChckdName = "optnBtnShowCheckedOnly" '& ActiveSheet.Index
    lOptnShowAllName = "optnBtnShowAll" '& ActiveSheet.Index
    
    Set lOptnShowChckd = ActiveSheet.OptionButtons.Add(60, lCell.Top, 55, 10)
    lOptnShowChckd.Characters.Text = "Skjul uavkryssede"
    lOptnShowChckd.name = lOptnShowChckdName
    
    Set lOptnShowAll = ActiveSheet.OptionButtons.Add(120, lCell.Top, 30, 10)
    lOptnShowAll.Characters.Text = "Vis alle"

    lOptnShowAll.name = lOptnShowAllName
    
    ActiveSheet.Shapes.Range(Array(lOptnShowChckdName, lOptnShowAllName)).Select
    
      Selection.ShapeRange.Group.Select
      
      Set lGroup = Selection
      lGroup.name = "OptionButtonsShowAllVsSelected" '& ActiveSheet.Index
      lGroup.Top = lCell.Top + 16
      lGroup.Height = 11
      lGroup.Width = 136
      lGroup.Left = 58
      
      Range("B2").Select
      
      lGroup.ShapeRange.GroupItems(lOptnShowAllName).ControlFormat.Value = 1
      
      
End Sub
  

Private Sub resizeRightAlignButtons()
'Brukes for å midtstille form controls i cellene, og lage størrelsen 80% av cellen de er i
'Brukes i prosjektinfo-arket

    Dim lCurrentWorksheet As Worksheet
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
    Set lCurrentWorksheet = ActiveSheet
    
    For Each lCurrentShape In lCurrentWorksheet.Shapes
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
              lCurrentShape.Left = lCellLeft + lCellWidth * 0.9 - lCurrentShape.Width / 2
              
          End If
        End If
    Next lCurrentShape

End Sub


Private Sub resizeCheckBoxes()

  Dim lCurrentSheet As Worksheet
  Dim lCurrentCheckbox As CheckBox
  Const cCheckboxSize As Double = 12.75
  
  Set lCurrentSheet = ActiveSheet
  
  For Each lCurrentSheet In Worksheets
    For Each lCurrentCheckbox In lCurrentSheet.CheckBoxes
      lCurrentCheckbox.Width = cCheckboxSize
      lCurrentCheckbox.Height = cCheckboxSize
    Next
  Next

End Sub


Private Sub repopulateDatatypeCheckboxes(inSheet As Worksheet)
' Delete all checkboxes and put on new ones
' New ones are generated for cells where validation input message starts with "Kryss av"
    
    Dim lCurrCell As Range
    Dim lCurrCheckBox As CheckBox
    Const cCheckboxSize As Double = 12.75

    inSheet.Activate
    ActiveSheet.UsedRange
  
    ' First, delete all checkboxes:
    inSheet.CheckBoxes.Delete

    ' Insert new checkboxes
    ' Insert checkbox for each cell where text starts with "' - "
    For Each lCurrCell In inSheet.UsedRange
      
      If lCurrCell.Value Like "* - *" Then
          
        Set lCurrCheckBox = inSheet.CheckBoxes.Add(Top:=lCurrCell.Top _
          , Left:=lCurrCell.Left, Height:=cCheckboxSize, Width:=cCheckboxSize)

        With lCurrCheckBox
        
          .name = lCurrCell.Address
          .Caption = ""
          
        End With
        
      End If
    
    Next lCurrCell
    
    rightAlignCheckBoxes inSheet
    
End Sub


Private Sub center()
  centerCheckBoxes
End Sub


Private Sub centerCheckBoxes(Optional inSheet As Worksheet)

    Dim lCurrCell As Range
    Dim lCurrCheckBox As CheckBox
    Const cCheckboxSize As Double = 12.75
    Const cTopOffset = 1.5

    If inSheet Is Nothing Then

      Set inSheet = ActiveSheet

    End If

    For Each lCurrCheckBox In inSheet.CheckBoxes
      Set lCurrCell = inSheet.Range(lCurrCheckBox.LinkedCell)
      
      With lCurrCheckBox
    
        .Height = cCheckboxSize
        .Width = cCheckboxSize
        .Top = lCurrCell.Top + cTopOffset
        .Left = lCurrCell.Left + (lCurrCell.Width - .Width) / 2
        
      End With
  
    Next lCurrCheckBox
    
End Sub


Private Sub rightAlignCheckBoxes(inSheet As Worksheet)

    Dim lCurrCell As Range
    Dim lCurrCheckBox As CheckBox
    Const cCheckboxSize As Double = 12.75
    Const cTopOffset = 1.5

    For Each lCurrCheckBox In inSheet.CheckBoxes
      Set lCurrCell = inSheet.Range(lCurrCheckBox.LinkedCell)
      
      With lCurrCheckBox
    
        .Height = cCheckboxSize
        .Width = cCheckboxSize
        .Top = lCurrCell.Top + cTopOffset
        .Left = lCurrCell.Left + lCurrCell.Width - .Width * 2
        
      End With
  
    Next lCurrCheckBox
    
End Sub


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



Private Sub test_repopulateDatatypeCheckboxes()
  Dim lCurrSheet As Worksheet
  
  Set lCurrSheet = Worksheets(1)
  lCurrSheet.Activate
  repopulateDatatypeCheckboxes lCurrSheet
  
End Sub

Private Sub test_repopulateAnalysCheckboxes()
  Dim lCurrSheet As Worksheet

  Set lCurrSheet = Worksheets(2)
  lCurrSheet.Activate
  repopulateAnalysCheckboxes lCurrSheet
  
End Sub


Private Sub test_centerCheckBoxes()
  Dim lCurrSheet As Worksheet

  Set lCurrSheet = Worksheets(5)
  lCurrSheet.Activate
  centerCheckBoxes lCurrSheet
  
End Sub


Private Sub test_rightAlignCheckBoxes()
  Dim lCurrSheet As Worksheet
  
  Set lCurrSheet = Worksheets(1)
  lCurrSheet.Activate
  rightAlignCheckBoxes lCurrSheet
  
End Sub


Private Sub test_unhideAllRowsAndFormControls()
  Dim lCurrSheet As Worksheet
  
  Set lCurrSheet = Worksheets(2)
  lCurrSheet.Activate
  unhideAllRowsAndFormControls

End Sub


Private Sub test_hideUncheckedAnalyses()
  Dim lCurrSheet As Worksheet
  On Error Resume Next

  Set lCurrSheet = Worksheets(2)
  lCurrSheet.Activate
  hideUncheckedAnalyses
  
End Sub


Private Sub check()

  Dim ws As Worksheet
  Dim sh As Shape
  Dim obj As Object
  
    Set ws = ActiveSheet
    
      For Each sh In ws.Shapes
        
        If sh.Type <> msoFormControl Then
        
          Debug.Print sh.name
          Debug.Print sh.ID
          Debug.Print sh.Type
     
     End If
    
    Next sh
    
End Sub
    
    
  

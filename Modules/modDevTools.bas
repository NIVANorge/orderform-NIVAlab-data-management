Attribute VB_Name = "modDevTools"
Option Explicit

' Procedures and functions for developer use

Private Sub protect_wb_and_ws()
    
    Dim lCurrentWorksheet As Worksheet
    
    For Each lCurrentWorksheet In ActiveWorkbook.Worksheets
    
        lCurrentWorksheet.Protect "encrypted", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
        lCurrentWorksheet.EnableSelection = xlUnlockedCells
    
    Next lCurrentWorksheet
    
    ActiveWorkbook.Protect "encrypted"
    
End Sub


Private Sub unprotect_wb_and_ws()
    
    Dim lCurrentWorksheet As Worksheet
    
    For Each lCurrentWorksheet In ActiveWorkbook.Worksheets
    
        lCurrentWorksheet.Unprotect "encrypted"
        
    Next lCurrentWorksheet
    
    ActiveWorkbook.Unprotect "encrypted"
    
End Sub

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
              lCurrentShape.Left = lCellLeft + (lCellWidth - lCurrentShape.Width) / 2
              
             'Debug.Print (lCurrentShape.Type & ", " & lCurrentShape.name & ", " & lCurrentShape.TopLeftCell.Row & ", " & lCurrentShape.BottomRightCell.Row)
              
          End If
        End If
    Next lCurrentShape

End Sub


Private Sub resizeRightAlignButtons()
'Brukes for å midtstille fom controls i cellene, og lage størrelsen 80% av cellen de er i
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
              'LCurrentShape.Left = lCellLeft + (lCellWidth - LCurrentShape.Width) / 2
              lCurrentShape.Left = lCellLeft + lCellWidth * 0.9 - lCurrentShape.Width / 2
              
          End If
        End If
    Next lCurrentShape

End Sub


Private Function getValidationErrorMessage(inCell As Range) As String
    getValidationErrorMessage = inCell.Validation.ErrorMessage
End Function


Private Function getValidationInputMessage(inCell As Range)
    getValidationInputMessage = inCell.Validation.InputMessage
End Function


Private Sub PrepWorkBook()
    ResetUsedRng
    DeleteEmptySheets
End Sub

Private Sub DeleteEmptySheets()
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=396

    Dim CurrentSheet As Worksheet

    For Each CurrentSheet In Worksheets
        If Not IsChart(CurrentSheet) Then
            If Application.WorksheetFunction.CountA(CurrentSheet.Cells) = 0 Then
                Application.DisplayAlerts = False
                CurrentSheet.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next

End Sub

Sub ResetUsedRng()
    'Resets used range for all worksheets
    'Convenient before using Ctrl-End, Ctrl-down arrow etc.
    ' JVE
    Dim InitialActiveSheet As Worksheet
    Dim lWorksheet As Worksheet
    
    Set InitialActiveSheet = ActiveSheet
    Application.ScreenUpdating = False
    
    For Each lWorksheet In Worksheets
        lWorksheet.Activate
        ActiveSheet.UsedRange
    Next
    
    InitialActiveSheet.Activate
    Application.ScreenUpdating = True
    
End Sub

Private Function IsChart(InSheet As Worksheet) As Boolean
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=396

    Dim tmpChart As Chart
    On Error Resume Next
    Set tmpChart = Charts(InSheet.name)
    IsChart = IIf(tmpChart Is Nothing, False, True)
End Function

Private Sub DeleteSheet(inName As String, OutResult As Long)
    Dim CurrentSheet As Worksheet
    Dim MsBoxReturnValue As Long
    
    For Each CurrentSheet In ActiveWorkbook.Worksheets
        If CurrentSheet.name = inName Then
            MsBoxReturnValue = MsgBox("Sheet " & inName & " will be deleted", vbOKCancel)
            
            If MsBoxReturnValue = 1 Then
                Application.DisplayAlerts = False
                CurrentSheet.Delete
                Application.DisplayAlerts = True
                OutResult = 1
                Exit For
            Else
                OutResult = 2
            End If
            Exit For
        End If
    Next
End Sub

Private Sub ListAllObjectsActiveSheet()
    Dim NewSheet As Worksheet
    Dim MySheet As Worksheet
    Dim MyShape As Shape
    Dim MySheetName As String
    
    Dim i As Long

    Set MySheet = ActiveSheet
    MySheetName = Replace(ActiveSheet.name, "Analyserekv ", "")
    Set NewSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    NewSheet.name = MySheetName & "_shapes"

    With NewSheet
        .Range("A1").Value = "Name"
        .Range("B1").Value = "Visible(-1) or Not Visible(0)"
        .Range("C1").Value = "Shape type"
        .Range("D1").Value = "Width"
        .Range("E1").Value = "Height"
        .Range("F1").Value = "Left"
        .Range("G1").Value = "Top"
        .Range("H1").Value = "Alternative Text"
        .Range("I1").Value = "Id"
        
        i = 2

        For Each MyShape In MySheet.Shapes
            .Cells(i, 1).Value = MyShape.name
            .Cells(i, 2).Value = MyShape.Visible
            .Cells(i, 3).Value = MyShape.Type
            .Cells(i, 4).Value = MyShape.Width
            .Cells(i, 5).Value = MyShape.Height
            .Cells(i, 6).Value = MyShape.Left
            .Cells(i, 7).Value = MyShape.Top
            .Cells(i, 8).Value = MyShape.AlternativeText
            .Cells(i, 9).Value = MyShape.ID

            i = i + 1
        Next MyShape

        .Range("A1:I1").Font.Bold = True
        .Columns.AutoFit

    End With

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub ListAllObjectsActiveSheet()
    Dim NewSheet As Worksheet
    Dim MySheet As Worksheet
    Dim MyShape As Shape
    Dim MySheetName As String
    
    Dim i As Long

    Set MySheet = ActiveSheet
    MySheetName = Replace(ActiveSheet.Name, "Analyserekvisisjon ", "")
    Set NewSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    NewSheet.Name = MySheetName & "_shapes"

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
            .Cells(i, 1).Value = MyShape.Name
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

























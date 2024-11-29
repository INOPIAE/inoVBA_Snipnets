Attribute VB_Name = "mdl_ExportShapeData"
Option Explicit

Private ExApp As Excel.Application
Private wks As Worksheet
Private wkb As Workbook

Sub ExportCurrentShapeData()
     On Error GoTo Fehler
    Excel_Open
    Set wkb = ExApp.Workbooks.Add
    Set wks = wkb.Worksheets(1)
    wks.Activate
    
    Dim TheShp As Visio.Shape
    Set TheShp = ActiveWindow.Selection.Item(1)
    
    Dim Zelle As Long
    Zelle = 3
    
    Zelle = ExportShapeTransform(Zelle, TheShp) + 2
    Zelle = ExportAction(Zelle, TheShp) + 2
    Zelle = ExportUser(Zelle, TheShp) + 2
    Exit Sub
Fehler:
    Select Case Err.Number
       Case -2032465753
            MsgBox "Es wurde kein Shape ausgewählt"
            Err.Clear
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Debug.Print Err.Number
    End Select
End Sub

Private Function ExportUser(Zelle As Long, TheShp As Visio.Shape) As Long
    On Error GoTo Fehler
    Dim TheSec As Visio.Section
    Dim TheCell As Visio.Cell
    Set TheSec = TheShp.Section(visSectionUser)
    Dim RowNum As Long
    Dim ColNum As Long
    Dim ZelleI As Long
    Dim Header As String
    ZelleI = Zelle
    For RowNum = 0 To TheSec.Count - 1
        For ColNum = 0 To 1
            Set TheCell = TheShp.CellsSRC(visSectionUser, RowNum, ColNum)
            If ZelleI = Zelle Then
                If ColNum = 0 Then
                
                    Header = TheCell.Name
                    wks.Cells(ZelleI - 1, ColNum + 1).Value = "User defined Cells"
                Else
                    wks.Cells(ZelleI - 1, ColNum + 1).Value = Replace(TheCell.Name, Header & ".", "")
                End If
                
            End If
            wks.Cells(ZelleI, ColNum + 1).Value = TheCell.Formula
        Next
        ZelleI = ZelleI + 1
    Next
    
    ExportUser = ZelleI
    Exit Function
Fehler:
    Select Case Err.Number
        Case -2032465751
            Err.Clear
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Debug.Print Err.Number
    End Select
    ExportUser = Zelle
End Function

Private Function ExportAction(Zelle As Long, TheShp As Visio.Shape) As Long
    Dim TheSec As Visio.Section
    Dim TheCell As Visio.Cell
    Set TheSec = TheShp.Section(visSectionAction)
    Dim RowNum As Long
    Dim ColNum As Long
    Dim ZelleI As Long
    Dim Header As String
    ZelleI = Zelle
    For RowNum = 0 To TheSec.Count - 1
        For ColNum = 0 To 10
            Set TheCell = TheShp.CellsSRC(visSectionAction, RowNum, ColNum)
            If ZelleI = Zelle Then
                If ColNum = 0 Then
                
                    Header = TheCell.Name
                    wks.Cells(ZelleI - 1, ColNum + 1).Value = "Actions"
                Else
                    wks.Cells(ZelleI - 1, ColNum + 1).Value = Replace(TheCell.Name, Header & ".", "")
                End If
                
            End If
            wks.Cells(ZelleI, ColNum + 1).Value = TheCell.Formula
        Next
        ZelleI = ZelleI + 1
    Next
    
    ExportAction = ZelleI
    Exit Function
Fehler:
    Select Case Err.Number
        Case -2032465751
            Err.Clear
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Debug.Print Err.Number
    End Select
    ExportAction = Zelle
End Function

Private Sub Excel_Open()
    On Error GoTo Fehler
        Set ExApp = GetObject(Class:="Excel.Application")
        ExApp.Visible = True
    
    Exit Sub
Fehler:
    Select Case Err.Number
        Case 429
            Set ExApp = CreateObject(Class:="Excel.Application")
            Resume Next
        Case Else
            VBA.MsgBox Err.Number & " " & Err.Description
    End Select
End Sub

Private Function ExportShapeTransform(Zelle As Long, TheShp As Visio.Shape) As Long
    On Error GoTo Fehler
    Dim TheSec As Visio.Section
    Dim TheCell As Visio.Cell
'    Set TheSec = TheShp.Section(visSectionUser)
    Dim RowNum As Long
    Dim ColNum As Long
    Dim ZelleI As Long
    Dim Header As String
    
    ZelleI = Zelle
        
    wks.Cells(ZelleI, 1).Value = "Shape Transform"
    ZelleI = ZelleI + 1
    wks.Cells(ZelleI, 1).Value = "Width"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("Width")
    ZelleI = ZelleI + 1
        wks.Cells(ZelleI, 1).Value = "Height"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("Height")
    ZelleI = ZelleI + 1
        wks.Cells(ZelleI, 1).Value = "Angle"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("Angle")
    ZelleI = ZelleI + 1
        wks.Cells(ZelleI, 1).Value = "PinX"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("PinX")
    ZelleI = ZelleI + 1
        wks.Cells(ZelleI, 1).Value = "PinY"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("PinY")
    ZelleI = ZelleI + 1
        wks.Cells(ZelleI, 1).Value = "LocPinX"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("LocPinX")
    ZelleI = ZelleI + 1
        wks.Cells(ZelleI, 1).Value = "LocPinY"
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("LocPinY")
    ZelleI = ZelleI + 1
    
    ExportShapeTransform = ZelleI
    Exit Function
Fehler:
    Select Case Err.Number
        Case -2032465751
            Err.Clear
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Debug.Print Err.Number
    End Select
    ExportShapeTransform = Zelle
End Function


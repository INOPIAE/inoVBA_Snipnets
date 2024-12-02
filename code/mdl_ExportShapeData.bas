Attribute VB_Name = "mdl_ExportShapeData"
Option Explicit

Private ExApp As Excel.Application
Private wks As Worksheet
Private wkb As Workbook

Sub ExportCurrentShapeData()
     On Error GoTo Fehler
    
    
    Dim TheShp As Visio.Shape
    Dim shp As Visio.Shape
    Set TheShp = ActiveWindow.Selection.Item(1)
    
    Set MsgDict = CreateObject("Scripting.Dictionary")
    Call LoadDictionary

    
    Excel_Open
    Set wkb = ExApp.Workbooks.Add
    Set wks = wkb.Worksheets(1)
    wks.Activate
    
    Dim Zelle As Long
    Zelle = 3
    
    
    Zelle = ShapesAuswerten(TheShp, Zelle)

    For Each shp In TheShp.Shapes
        Zelle = ShapesAuswerten(shp, Zelle)
    Next


    Exit Sub
Fehler:
    Select Case Err.Number
       Case -2032465753
            MsgBox "Es wurde kein Shape ausgewählt"
            Err.Clear
            On Error GoTo 0
        Case -2032465751
            Err.Clear
            Resume Next
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Debug.Print Err.Number
    End Select

End Sub

Private Function ShapesAuswerten(ByVal TheShp As Visio.Shape, ByVal Zelle As Long) As Long
    
    With wks.Cells(Zelle, 1)
        .Value = TheShp.Name
        .Font.FontStyle = "Fett"
        .Font.Size = 14
    End With
    Zelle = Zelle + 1
    Zelle = ExportShapeTransform(Zelle, TheShp)
    Zelle = ExportSection(Zelle, TheShp, visSectionUser, "User defined cell", "User", "Value", 2)
    Zelle = ExportSection(Zelle, TheShp, visSectionProp, "Shape Data", "Prop", "Label", 8)
    Zelle = ExportSection(Zelle, TheShp, visSectionHyperlink, "Hyperlinks", "Hyperlink", "Description", 9)
    Zelle = ExportSection(Zelle, TheShp, visSectionConnectionPts, "Connection Points", "Connections", "X1", 5)
    Zelle = ExportSection(Zelle, TheShp, visSectionAction, "Actions", "Actions", "Menue", 9)
    
    Zelle = ExportSection(Zelle, TheShp, visSectionCharacter, "Character", "Character", "Font", 20)
    Zelle = ExportSection(Zelle, TheShp, visSectionParagraph, "Paragraph", "Paragraph", "IndFirst", 20)
    
    Zelle = ExportGeometry(Zelle, TheShp, visSectionParagraph, "Geometry", "Paragraph", "IndFirst", 20)
    

    
    ShapesAuswerten = Zelle + 2
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

    Dim RowNum As Long
    Dim ColNum As Long
    Dim ZelleI As Long
    Dim Header As String
    
    ZelleI = Zelle
        
    With wks.Cells(ZelleI, 1)
        .Value = "Shape Transform"
        .Font.FontStyle = "Fett"
    End With
    
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "Name"
    wks.Cells(ZelleI, 2).Value = "Formula"
    wks.Cells(ZelleI, 3).Value = "Value"
    
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "Width"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("Width")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("Width").FormulaU
    
    ZelleI = ZelleI + 1
    wks.Cells(ZelleI, 1).Value = "Height"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("Height")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("Height").FormulaU
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "Angle"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("Angle")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("Angle").FormulaU
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "PinX"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("PinX")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("PinX").FormulaU
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "PinY"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("PinY")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("PinY").FormulaU
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "LocPinX"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("LocPinX")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("LocPinX").FormulaU
    ZelleI = ZelleI + 1
    
    wks.Cells(ZelleI, 1).Value = "LocPinY"
    wks.Cells(ZelleI, 3).Value = TheShp.Cells("LocPinY")
    wks.Cells(ZelleI, 2).Value = TheShp.Cells("LocPinY").FormulaU
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

Private Function ExportSection(ByVal Zelle As Long, ByVal TheShp As Visio.Shape, ByVal visSection As Long, _
        ByVal SectionName As String, ByVal SectionReplace As String, ByVal SectionFirstColumn As String, _
        ByVal ColMax As Long) As Long
        
    On Error GoTo Fehler
    Dim TheSec As Visio.Section
    Dim TheCell As Visio.Cell
    If TheShp.SectionExists(visSection, visExistsLocally) = False Then
        ExportSection = Zelle
        Exit Function
    End If
    Zelle = Zelle + 2
    Set TheSec = TheShp.Section(visSection)
    Dim RowNum As Long
    Dim ColNum As Long
    Dim ZelleI As Long
    Dim Header As String
    Dim VersatzSpalte As Long
    VersatzSpalte = 2
    ZelleI = Zelle + 1
    For RowNum = 0 To TheSec.Count - 1
        For ColNum = 0 To ColMax
            Set TheCell = TheShp.CellsSRC(visSection, RowNum, ColNum)
            If ZelleI = Zelle + 1 Then
                If ColNum = 0 Then
                
                    Header = TheCell.Name
                    With wks.Cells(ZelleI - 2, ColNum + 1)
                        .Value = SectionName
                        .Font.FontStyle = "Fett"
                    End With
                    wks.Cells(ZelleI - 1, ColNum + 1).Value = "Name"
                    wks.Cells(ZelleI - 1, ColNum + VersatzSpalte).Value = SectionFirstColumn
                Else
                    wks.Cells(ZelleI - 1, ColNum + VersatzSpalte).Value = Replace(TheCell.Name, Header & ".", "")
                End If
                
            End If
            If ColNum = 0 Then
                wks.Cells(ZelleI, ColNum + 1).Value = Replace(TheCell.Name, SectionReplace & ".", "")
            End If
            wks.Cells(ZelleI, ColNum + VersatzSpalte).Value = TheCell.FormulaU
        Next
        ZelleI = ZelleI + 1
    Next
    
    ExportSection = ZelleI
    Exit Function
Fehler:
    Select Case Err.Number
        Case -2032465751
            Err.Clear
            On Error GoTo 0
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Debug.Print Err.Number
    End Select
    ExportSection = Zelle
End Function

Private Function ExportGeometry(ByVal Zelle As Long, ByVal TheShp As Visio.Shape, ByVal visSection As Long, _
        ByVal SectionName As String, ByVal SectionReplace As String, ByVal SectionFirstColumn As String, _
        ByVal ColMax As Long) As Long
        
    Dim TheSec As Visio.Section
    Dim TheCell As Visio.Cell
    Dim intCurrentGeometrySection As Integer
    Dim intCurrentGeometrySectionIndex As Integer
    Dim intRows As Integer
    Dim intCells As Integer
    Dim intCurrentRow As Integer
    Dim intCurrentCell As Integer
    Dim intSections As Integer
 
    Zelle = Zelle + 2
    
    Dim RowNum As Long
    Dim ColNum As Long
    Dim ZelleI As Long
    Dim Header As String
    Dim VersatzSpalte As Long
    VersatzSpalte = 2
    ZelleI = Zelle + 1
    'Get the count of Geometry sections in the shape.
    '(If the shape is a group, this will be 0.)
    
    intSections = TheShp.GeometryCount
    
    'Iterate through all Geometry sections for the shape.
    'Because we are adding the current Geometry section index to
    'the constant visSectionFirstComponent, we must start with 0.
    For intCurrentGeometrySectionIndex = 0 To intSections - 1
        With wks.Cells(ZelleI - 2, 1)
            Header = SectionName & intCurrentGeometrySectionIndex + 1
            .Value = SectionName & intCurrentGeometrySectionIndex + 1
            .Font.FontStyle = "Fett"
        End With
         'Set a variable to use when accessing the current
         'Geometry section.
         intCurrentGeometrySection = visSectionFirstComponent + intCurrentGeometrySectionIndex
         
         'Get the count of rows in the current Geometry section.
         intRows = TheShp.RowCount(intCurrentGeometrySection)

         'Loop through the rows. The count is zero-based.
         For intCurrentRow = 0 To (intRows - 1)
         
         'Get the count of cells in the current row.
             intCells = TheShp.RowsCellCount(intCurrentGeometrySection, intCurrentRow)
             
             'Loop through the cells. Again, this is zero-based.
             For intCurrentCell = 0 To (intCells - 1)
                
                If intCurrentRow = 0 Then
                    wks.Cells(ZelleI - 1, intCurrentCell + VersatzSpalte).Value = Replace(TheShp.CellsSRC(intCurrentGeometrySection, intCurrentRow, _
                    intCurrentCell).LocalName, Header & ".", "")
                End If
                If intCurrentRow = 1 Then
                    wks.Cells(ZelleI - 1, 1).Value = "Name"
                    wks.Cells(ZelleI - 1, 2).Value = "X"
                    wks.Cells(ZelleI - 1, 3).Value = "Y"
                    wks.Cells(ZelleI - 1, 4).Value = "A"
                    wks.Cells(ZelleI - 1, 5).Value = "B"
                    wks.Cells(ZelleI - 1, 6).Value = "C"
                    wks.Cells(ZelleI - 1, 7).Value = "D"
                    wks.Cells(ZelleI - 1, 8).Value = "E"
                    
                End If
                 
                 wks.Cells(ZelleI, intCurrentCell + VersatzSpalte).Value = TheShp.CellsSRC(intCurrentGeometrySection, intCurrentRow, _
                 intCurrentCell).FormulaU
                 
                 
             Next intCurrentCell
             If intCurrentRow = 0 Then
                ZelleI = ZelleI + 2
             Else
                wks.Cells(ZelleI, 1).Value = ResolveEnum("G", TheShp.RowType(intCurrentGeometrySection, intCurrentRow))
                ZelleI = ZelleI + 1
             End If
         Next intCurrentRow
         ZelleI = ZelleI + 3
     Next intCurrentGeometrySectionIndex
     

    
    ExportGeometry = ZelleI
End Function


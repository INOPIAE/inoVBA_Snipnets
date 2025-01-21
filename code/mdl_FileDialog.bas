Attribute VB_Name = "mdl_OpenFileDialog"
Option Explicit

'// Module: OpenExcelFile
'//
'// This is code that uses the Windows API to invoke the Open File
'// common dialog. It is used by users to choose an Excel file that
'// contains organizational data.

' taken from https://visguy.com/vgforum/index.php?topic=738.0

Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (OFN As OPENFILENAME) As Boolean

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Public Sub FindExcelFile(ByRef filepath As String, _
                        ByRef cancelled As Boolean)

   Dim OpenFile As OPENFILENAME
   Dim lReturn As Long
   Dim sFilter As String
   
   ' On Error GoTo errTrap
   
   OpenFile.lStructSize = LenB(OpenFile)

   '// Sample filter:
   '// "Text Files (*.txt)" & Chr$(0) & "*.sky" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
   sFilter = "Excel Files (*.xl*)" & Chr(0) & "*.xl*"
   
   OpenFile.lpstrFilter = sFilter
   OpenFile.nFilterIndex = 1
   OpenFile.lpstrFile = String(257, 0)
   OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
   OpenFile.lpstrFileTitle = OpenFile.lpstrFile
   OpenFile.nMaxFileTitle = OpenFile.nMaxFile
   OpenFile.lpstrInitialDir = ThisDocument.path
   
   OpenFile.lpstrTitle = "Find Excel Data Source"
   OpenFile.Flags = 0
   lReturn = GetOpenFileName(OpenFile)
   
   If lReturn = 0 Then
      cancelled = True
      filepath = vbNullString
   Else
     cancelled = False
     filepath = Trim(OpenFile.lpstrFile)
     filepath = Replace(filepath, Chr(0), vbNullString)
   End If

   Exit Sub
   
errTrap:
   Exit Sub
   Resume

End Sub

Public Function OpenFileDialog() As String

    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim sFilter As String
    Dim filepath As String
    
    ' On Error GoTo errTrap
    
    OpenFile.lStructSize = LenB(OpenFile)
    
    '// Sample filter:
    '// "Text Files (*.txt)" & Chr$(0) & "*.sky" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    sFilter = "Excel Files (*.xl*)" & Chr(0) & "*.xl*" & Chr$(0) & "PDF Files (*.pdf)" & Chr(0) & "*.pdf" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = ThisDocument.path
    
    OpenFile.lpstrTitle = "Dateiauswahl (Öffnen zum Bestätigen nutzen)"
    OpenFile.Flags = 0
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
       filepath = vbNullString
    Else
      filepath = Trim(OpenFile.lpstrFile)
      filepath = Replace(filepath, Chr(0), vbNullString)
    End If
    
    OpenFileDialog = filepath
    
    Exit Function
   
errTrap:
       Exit Function
       Resume

End Function

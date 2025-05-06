Attribute VB_Name = "mdl_ExportSQL"
Option Compare Database
Option Explicit

Sub ExportTableAsSQL(ByVal tableName As String, ByVal filePath As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim sqlLine As String
    Dim exportSQL As String
    Dim tableName As String
    Dim filePath As String
    Dim f As Integer
    Dim i As Integer

    Set db = CurrentDb
    Set rs = db.OpenRecordset(tableName, dbOpenSnapshot)

    If rs.EOF Then
        MsgBox "Tabelle leer."
        Exit Sub
    End If

    f = FreeFile()
    Open filePath For Output As #f
    
    Do While Not rs.EOF
        sqlLine = "INSERT INTO " & tableName & " ("
        
        ' Spaltennamen
        For i = 0 To rs.Fields.Count - 1
            sqlLine = sqlLine & rs.Fields(i).Name & ", "
        Next i
        sqlLine = Left(sqlLine, Len(sqlLine) - 2) ' letztes Komma entfernen
        
        sqlLine = sqlLine & ") VALUES ("
        
        ' Werte
        For i = 0 To rs.Fields.Count - 1
            If IsNull(rs.Fields(i).Value) Then
                sqlLine = sqlLine & "NULL, "
            ElseIf rs.Fields(i).Type = dbText Or rs.Fields(i).Type = dbMemo Then
                sqlLine = sqlLine & "'" & Replace(rs.Fields(i).Value, "'", "''") & "', "
            ElseIf rs.Fields(i).Type = dbDate Then
                sqlLine = sqlLine & "#" & Format(rs.Fields(i).Value, "yyyy-mm-dd hh:nn:ss") & "#, "
            ElseIf rs.Fields(i).Type = dbBoolean Then
                sqlLine = sqlLine & IIf(rs.Fields(i).Value, "True", "False") & ", "
            Else
                sqlLine = sqlLine & rs.Fields(i).Value & ", "
            End If
        Next i
        
        sqlLine = Left(sqlLine, Len(sqlLine) - 2) ' letztes Komma entfernen
        
        sqlLine = sqlLine & ");"
        
        Print #f, sqlLine
        
        rs.MoveNext
    Loop

    Close #f
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "SQL Export abgeschlossen!" & vbCrLf & "Datei: " & filePath
End Sub

Function ExportCreateTableSQL(tableName As String) As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim sql As String
    Dim fieldLine As String
    Dim i As Integer

    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)

    sql = "CREATE TABLE [" & tableName & "] (" & vbCrLf

    For i = 0 To tdf.Fields.Count - 1
        Set fld = tdf.Fields(i)
        fieldLine = "  [" & fld.Name & "] " & GetSQLType(fld)

        If fld.Required Then fieldLine = fieldLine & " NOT NULL"

        If i < tdf.Fields.Count - 1 Then
            fieldLine = fieldLine & ","
        End If

        sql = sql & fieldLine & vbCrLf
    Next i

    sql = sql & ");"

    ExportCreateTableSQL = sql
End Function

Function GetSQLType(fld As DAO.Field) As String
    Select Case fld.Type
        Case dbText
            GetSQLType = "TEXT(" & fld.Size & ")"
        Case dbMemo
            GetSQLType = "MEMO"
        Case dbByte
            GetSQLType = "BYTE"
        Case dbInteger
            GetSQLType = "SMALLINT"
        Case dbLong
            If (fld.Attributes And dbAutoIncrField) <> 0 Then
                GetSQLType = "COUNTER"
            Else
                GetSQLType = "LONG"
            End If
        Case dbSingle
            GetSQLType = "SINGLE"
        Case dbDouble
            GetSQLType = "DOUBLE"
        Case dbCurrency
            GetSQLType = "CURRENCY"
        Case dbDate
            GetSQLType = "DATETIME"
        Case dbBoolean
            GetSQLType = "YESNO"
        Case Else
            GetSQLType = "UNKNOWN"
    End Select
End Function

Sub ExportCreateSQLToFile(tableName As String, filePath As String)
    Dim sql As String
    sql = ExportCreateTableSQL(tableName)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, sql
    Close #fileNum
    
    MsgBox "SQL-Definition gespeichert unter: " & filePath
End Sub

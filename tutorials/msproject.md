# MS Project

Der aufgezeigt VBA-Code wird jeweils auch am Ende eines Abschnittes als Link zur Verfügung zu einer Datei gestellt, die einfach über den VBA-Editor in die MS Project Datei importiertet werden kann.

## Ribbon Erstellung für MS Project Dateien.

Für eigene Ribbons in MS Project kann nicht der Custom UI Editor verwendetet werden.

Allerdings kann dieser Code genutzt werden.

Zur Erstellung eines Ribbon kann dieser Code genutzt werden:

```
Sub AddCustomUI()
    Dim customUiXml As String
 
    customUiXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" _
        & "<mso:ribbon><mso:tabs><mso:tab id=""inoTabReporting"" label=""Reporting"" " _
        & "insertBeforeQ=""mso:TabView"">" _
        & "<mso:group id=""inoGrpZeit"" label=""Zeitskalen"">" _
        & "<mso:button id=""inoBtnJMW"" label=""Jahr Monat Woche"" size=""normal"" " _
        & " onAction=""ZeitleisteJahrMonatWoche"" />" _
        & "<mso:button id=""inoBtnWTH"" label=""Woche Tage Stunden"" size=""normal"" " _
        & " onAction=""ZeitleisteWocheTageStunden"" />" _
        & "</mso:group></mso:tab></mso:tabs></mso:ribbon></mso:customUI>"
        
    ActiveProject.SetCustomUI (customUiXml)
End Sub

```

In dem Codebeispiel wird eine neue Registerkarte mit dem Namen Reporting vor der Registerkarte Ansicht eingefügt.

![Screenshot Ribbon](/sources/screenshoot_project_ribbon.png)

Diese enthält die Gruppe Zeitskalen mit zwei Schlatflächen "Jahr Monat Woche" und "Woche Tage Stunden".

Mit dem OnAction-Tag wird die entsprechende Prozedure im Code angesprochen.

Der Unterschied zu den Prozeduren für die OnAction Tags der anderen Office-Producte wird hier kein Parameter vom Typ iRibbonControl übergeben.

Damit das Ribbon in der Datei angezeigt wird, muss die Prozedur AddCustomUI in der Project_Open Routine aufgerufen werden.

```
Private Sub Project_Open(ByVal pj As Project)
    mdl_ribbon.AddCustomUI
End Sub
```



Zum Aufräumen kann dieser Code verwendet werden:

```
Sub RemoveCustomUI()
    Dim customUiXml As String
 
    customUiXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" _
        & "<mso:ribbon></mso:ribbon></mso:customUI>"
 
    ActiveProject.SetCustomUI (customUiXml)
End Sub
```

Dieser wird in der Project_BeforeClose Routine aufgerufen.

```
Private Sub Project_BeforeClose(ByVal pj As Project)
    mdl_ribbon.RemoveCustomUI
End Sub
```

Eine angepasste Version des Ribbons mit 3 Bereichen:
* Anpassung Zeitachse im Ganttdiagramm
* Export aller Vorgänge in den eigenen Kalender
* Export aller Besprechungen als Meeting nach Outlook

```
Sub AddCustomUI()
    Dim customUiXml As String
    Dim customUiXml1 As String
    Dim customUiXml2 As String
    Dim customUiXml3 As String
    
    customUiXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" _
        & "<mso:ribbon><mso:tabs><mso:tab id=""inoTabReporting"" label=""Reporting"" " _
        & "insertBeforeQ=""mso:TabView"">"
    'Group Timescale
    customUiXml1 = "<mso:group id=""inoGrpZeit"" label=""Zeitskalen"">" _
        & "<mso:button id=""inoBtnJMW"" label=""Jahr Monat Woche"" size=""normal"" " _
        & " onAction=""ZeitleisteJahrMonatWoche"" />" _
        & "<mso:button id=""inoBtnWTH"" label=""Woche Tage Stunden"" size=""normal"" " _
        & " onAction=""ZeitleisteWocheTageStunden"" />" _
        & "</mso:group>"
    'Group Export tasks
    customUiXml2 = "<mso:group id=""inoGrpExport"" label=""Export nach Outlook"">" _
        & "<mso:button id=""inoBtnOExport"" label=""Alles"" size=""normal"" " _
        & " onAction=""ExportTasksToOutlook"" />" _
        & "<mso:button id=""inoBtnOExportM"" label=""Meilensteine"" size=""normal"" " _
        & " onAction=""ExportMilestonesToOutlook"" />" _
        & "<mso:button id=""inoBtnOExportS"" label=""Sammelvorgänge"" size=""normal"" " _
        & " onAction=""ExportSummaryToOutlook"" />" _
        & "</mso:group>"
    'Group Export Meetings
    customUiXml3 = "<mso:group id=""inoGrpExportMeeting"" label=""Meeting Export nach Outlook"">" _
        & "<mso:button id=""inoBtnOExportMeeting"" label=""Meetings"" size=""normal"" " _
        & " onAction=""ExportMeetingsToOutlook"" />" _
        & "<mso:button id=""inoBtnOExportTeamsMeeting"" label=""Teams Meetings"" size=""normal"" " _
        & " onAction=""ExportTeamsMeetingsToOutlook"" />" _
        & "</mso:group>"
    
    customUiXml = customUiXml & customUiXml1 & customUiXml2 & customUiXml3 & "</mso:tab></mso:tabs></mso:ribbon></mso:customUI>"
        
    ActiveProject.SetCustomUI (customUiXml)
End Sub
```
Die Anzeige der einzelnen Bereiche kann durch auskommentieren des Blocks beeinflusst werden.

[Beispieldatei project/mdl_Ribbon.bas](/project/mdl_ribbon.bas)

## Einstellen der Zeitskala für das Gantt-Diagramm in MS Project

```
Sub ZeitleisteJahrMonatWoche()

    TimescaleEdit TierCount:=3, _
    TopUnits:=PjTimescaleUnit.pjTimescaleYears, TopLabel:=PjDateLabel.pjYear_yyyy, TopCount:=1, _
    MajorUnits:=PjTimescaleUnit.pjTimescaleMonths, MajorLabel:=PjMonthLabel.pjMonthLabelMonth_mmmm, MajorCount:=1, _
    MinorUnits:=PjTimescaleUnit.pjTimescaleWeeks, MinorLabel:=PjWeekLabel.pjWeekLabelWeekNumber_ww, MinorCount:=1, _
    Separator:=True
    
    GotoTaskDates

End Sub
```

```
Sub ZeitleisteWocheTageStunden()

    TimescaleEdit TierCount:=3, _
    TopUnits:=PjTimescaleUnit.pjTimescaleWeeks, TopLabel:=PjDateLabel.pjWeekNumber_ww, TopCount:=1, _
    MajorUnits:=PjTimescaleUnit.pjTimescaleDays, MajorLabel:=PjDateLabel.pjDay_ddi_mm_dd, MajorCount:=1, _
    MinorUnits:=PjTimescaleUnit.pjTimescaleHours, MinorLabel:=PjDateLabel.pjHour_hh, MinorCount:=6, _
    Separator:=True
    
    GotoTaskDates
    
End Sub
```

[Beispieldatei project/mdl_Zeitleiste.bas](/project/mdl_Zeitleiste.bas)

## Export der Vorgänge in den eigenen Outlook-Kalender

Um Daten nach Outlook zu exportieren muss im VBA Editor unter Extras-Verweis der Verweis auf die Mircrosoft Outlook xx Bibliothek gesetzt werden.


```
Option Explicit

Private olApp As Outlook.Application

Public Sub ExportTasksToOutlook()
    ExportToOutlook "A"
End Sub

Public Sub ExportMilestonesToOutlook()
    ExportToOutlook "M"
End Sub

Public Sub ExportSummaryToOutlook()
    ExportToOutlook "S"
End Sub

Public Sub ExportToOutlook(ByVal strType, Optional strFlag As String = "")
    Dim t As Task
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim StartDiff As Integer
    Dim EndDiff As Integer
    
    Set olApp = GetObject(, "Outlook.Application")

   
    StartDiff = DateDiff("d", Date, ActiveProject.ProjectStart)
    EndDiff = DateDiff("d", Date, ActiveProject.ProjectFinish)
    
    dtStart = Date - 1 + StartDiff
    dtEnd = Date + 30 + EndDiff
    
    For Each t In ActiveProject.Tasks
        Dim dtFinish As Date
        If t.Milestone = True Then
            dtFinish = DateAdd("n", 15, t.Finish)
        Else
            dtFinish = t.Finish
        End If
        Select Case strType
            
            Case "M"
                If t.Milestone = True Then
                    ExportAppointment dtStart, dtEnd, t.start, dtFinish, t.Name
                End If
            Case "S"
                If t.Summary = True Then
                    ExportAppointment dtStart, dtEnd, t.start, dtFinish, t.Name
                End If
            Case Else
                ExportAppointment dtStart, dtEnd, t.start, dtFinish, t.Name
        End Select
    Next
End Sub

Public Sub ExportAppointment(ByVal dtPStart As Date, ByVal dtPEnd As Date, ByVal dtStart As Date, ByVal dtEnd As Date, ByVal strSubject As String)

    Dim olAppoint As Outlook.AppointmentItem
    
    Set olAppoint = GetAppointmentInRange(dtPStart, dtPEnd, strSubject)
    
    If (Not (olAppoint Is Nothing)) Then
    
    Else
        Set olAppoint = olApp.CreateItem(olAppointmentItem)
    End If
    
    With olAppoint
        .start = dtStart
        .End = dtEnd
        .subject = strSubject
        .ReminderSet = False
        .AllDayEvent = False
        .Save
    End With
   
End Sub

Function GetAppointmentInRange(ByVal dtStart As Date, ByVal dtEnd As Date, ByVal strSubject As String) As Outlook.AppointmentItem

    Dim oCalendar As Folder
    
    Dim objItems As Items
    Dim objRestrictedItems As Items
    
    Dim filterRange As String
    
    Dim oItem As AppointmentItem
    
    Dim iIt As Long
    Dim nItFilter As Long
    Dim nIt As Long
    
    Set oCalendar = olApp.Session.GetDefaultFolder(olFolderCalendar)
       
    Set objItems = oCalendar.Items
    objItems.IncludeRecurrences = True
    objItems.Sort "[Start]"
                  
    filterRange = "[Start] >= " & Chr(34) & Format(dtStart, "yyyy-mm-dd hh:mm AM/PM") & Chr(34) & " AND " & _
                  "[End] <= " & Chr(34) & Format(dtEnd, "yyyy-mm-dd hh:mm AM/PM") & Chr(34)
        
    Set objRestrictedItems = objItems.Restrict(filterRange)
    
    nItFilter = objRestrictedItems.Count

    nIt = 0
    
    For Each oItem In objRestrictedItems
        If (Not (oItem Is Nothing)) Then
            nIt = nIt + 1
            
            If strSubject = oItem.subject Then
                Set GetAppointmentInRange = oItem
                Exit Function
            End If
            
        End If
    Next oItem
    
End Function
```

Die drei Funktionen `ExportTasksToOutlook`, `ExportMilestonesToOutlook` und `ExportSummaryToOutlook` werden im Ribbon genutzt.

[Beispieldatei project/mdl_OutlookExport.bas](/project/mdl_OutlookExport.bas)

## Export der Besprechungen als Meeting oder Teams-Meeting

Um Daten nach Outlook zu exportieren muss im VBA Editor unter Extras-Verweis der Verweis auf die Mircrosoft Outlook xx Bibliothek gesetzt werden.

Die Funktionalität nutzt das benutzerdefinierte Feld  `Text30` umdort Informationen zum exportierten Termin abzulegen. Das Feld wird auf `MeetingCheck` umbenannt.

```
Option Explicit

Private olApp As Outlook.Application

Public Sub ExportMeetingsToOutlook()
    ExportMeetingToOutlook
End Sub

Public Sub ExportTeamsMeetingsToOutlook()
    ExportMeetingToOutlook True
End Sub

Public Sub ExportMeetingToOutlook(Optional blnTeams As Boolean = False)
    Dim t As Task
    
    Set olApp = GetObject(, "Outlook.Application")

    RenameCustomColumn "Text30", "MeetingCheck"
    
    For Each t In ActiveProject.Tasks
    
        If t.Recurring And Not t.Summary And t.Resources.Count > 0 Then
            If t.Text30 <> t.Name & "|" & t.start & "|" & t.Finish Then
                ExportMeeting t
                t.Text30 = t.Name & "|" & t.start & "|" & t.Finish
            End If
        End If

    Next
 
End Sub

Public Sub ExportMeeting(ByVal t As Task, Optional blnTeams As Boolean = False)

    Dim olAppoint As Outlook.AppointmentItem
    Dim myRequiredAttendee As Outlook.Recipient
    Dim pr As Resource
    Dim EMail As String
            
    Set olAppoint = olApp.CreateItem(olAppointmentItem)

    With olAppoint
    
        .start = t.start
        .End = t.Finish
        .subject = t.Name
        .ReminderSet = False
        .AllDayEvent = False
        .MeetingStatus = olMeeting
    
        For Each pr In t.Resources

            If pr.EMailAddress <> "" Then
                EMail = pr.EMailAddress
            Else
                EMail = pr.Name
            End If
            
            Set myRequiredAttendee = olAppoint.Recipients.Add(EMail)
            myRequiredAttendee.Type = olRequired
            
        Next
        .Display
        If blnTeams Then
            SendKeys "&H", True
            SendKeys "TM", True
            appilation.wait (Now + TimeValue("00:00:01"))
        End If
    End With
   
End Sub

Public Sub RenameCustomColumn(ByVal InternalName, ByVal NewName As String, Optional FieldType As Long = pjTask)
    Dim c As Long
  
    c = FieldNameToFieldConstant(InternalName, FieldType) ' get constant of custom field by name
    
    If CustomFieldGetName(c) <> NewName Then
        CustomFieldRename FieldID:=c, NewName:=NewName  'Rename/set custom field title
    End If
End Sub
```

Die beiden Funktionen `ExportMeetingsToOutlook` und `ExportTeamsMeetingsToOutlook` werden im Ribbon genutzt.

[Beispieldatei project/mdl_OutlookExport.bas](/project/mdl_OutlookExport.bas)
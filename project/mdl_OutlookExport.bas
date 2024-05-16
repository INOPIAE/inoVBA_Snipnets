Attribute VB_Name = "mdl_OutlookExport"
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




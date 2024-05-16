Attribute VB_Name = "mdl_OutlookExportMeeting"
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

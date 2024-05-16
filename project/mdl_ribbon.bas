Attribute VB_Name = "mdl_ribbon"
Option Explicit

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


Sub RemoveCustomUI()
    Dim customUiXml As String
 
    customUiXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" _
        & "<mso:ribbon></mso:ribbon></mso:customUI>"
 
    ActiveProject.SetCustomUI (customUiXml)
End Sub


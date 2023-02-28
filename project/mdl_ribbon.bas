Attribute VB_Name = "mdl_ribbon"
Option Explicit

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

Sub RemoveCustomUI()
    Dim customUiXml As String
 
    customUiXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" _
        & "<mso:ribbon></mso:ribbon></mso:customUI>"
 
    ActiveProject.SetCustomUI (customUiXml)
End Sub



# inoVBA Snipnets

Hier finden sich einige VBA Code-Schniptzel und Beispieldateien für VBA-Lösungen in MS Office Produkten.


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

[project/mdl_Ribbon.bas](/project/mdl_ribbon.bas)

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

[project/mdl_Zeitleiste.bas](/project/mdl_Zeitleiste.bas)
Attribute VB_Name = "mdl_Zeitleiste"
Option Explicit

Sub ZeitleisteJahrMonatWoche()

    TimescaleEdit TierCount:=3, _
    TopUnits:=PjTimescaleUnit.pjTimescaleYears, TopLabel:=PjDateLabel.pjYear_yyyy, TopCount:=1, _
    MajorUnits:=PjTimescaleUnit.pjTimescaleMonths, MajorLabel:=PjMonthLabel.pjMonthLabelMonth_mmmm, MajorCount:=1, _
    MinorUnits:=PjTimescaleUnit.pjTimescaleWeeks, MinorLabel:=PjWeekLabel.pjWeekLabelWeekNumber_ww, MinorCount:=1, _
    Separator:=True
    
    GotoTaskDates
    
End Sub


Sub ZeitleisteWocheTageStunden()

    TimescaleEdit TierCount:=3, _
    TopUnits:=PjTimescaleUnit.pjTimescaleWeeks, TopLabel:=PjDateLabel.pjWeekNumber_ww, TopCount:=1, _
    MajorUnits:=PjTimescaleUnit.pjTimescaleDays, MajorLabel:=PjDateLabel.pjDay_ddi_mm_dd, MajorCount:=1, _
    MinorUnits:=PjTimescaleUnit.pjTimescaleHours, MinorLabel:=PjDateLabel.pjHour_hh, MinorCount:=6, _
    Separator:=True
    
    GotoTaskDates
    
End Sub


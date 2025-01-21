# Dateiauswahldialog für Visio VBA

Leider gibt es in Visio keinen funktionierenden Dateiauswahldialog. 

Wenn die Datei [mdl_FileDialog.bas](/code/mdl_FileDialog.bas) in das aktuelle VBA-Projekt eingebunden wird, steht ein entsprechender Dialog zur Verfügung.

Eine Datei kann dann so ausgewählt werden:

```
Sub Test
    Dim Dateiname as String
    Dateiname = OpenFileDialog
    If Dateiname = vbNullString Then
        MsgBox "Keine Datei angegeben."
        Exit Sub
    End If
End Sub
```

Es muss immer die Öffnen Schaltfläche genutzt werden auch beim Speichern.

![Screenshot Dateidialog](/sources/VisioDateiDialog.png)

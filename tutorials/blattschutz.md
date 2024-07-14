# Blattschutz beim Öffnen einer Excel-Datei abhängig vom Nutzer aufheben

Es gibt die Situation, dass eine Excel-Datei von den meisten Benutzer mit einem Blattschutz geöffnet werden soll. Der primäre Bearbeiter soll aber in der ungeschützten Datei arbeiten ohne jeweils das Kennwort eingeben zu müssen. Dazu wird die Excel Datei wie folgt angepasst.

Zuerst werden alle Bereiche, in der die Werte eingeben können, entsperrt. 

Mit `Zellen formatieren` wird auf der Registerkarte Schutz der Haken aus `Gesperrt` entfernt.

![Screenshot Zellschutz setzen](/sources/schutz_zellschutz.png)

Nun erfolgen die Anpassungen in VBA. Zunächst wird der VBA-Editor geöffnet. (Am schnellsten mit `Alt + F11`)

![Screenshot VBAProject](/sources/schutz_vba_project.png)

Im Projekt-Explorer wird der VBA Bereich von DieseArbeitsmappe mit Doppelklick geöffnet. (Falls der Project-Explorer nicht sichtbar ist, mit `Strg + R` öffnen.)

Der folgende Code wird in den Bereich hineinkopiert.

```
Option Explicit

Private Const pw As String = "MeinKennwort"

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim user As String
    user = Environ("Username")
    Select Case user
        Case "MeinBenutzername"
            If MsgBox("Soll die Datei geschützt werden?", vbYesNo) = vbYes Then
                DateiSperren
                ActiveWorkbook.Save
            End If
    End Select
End Sub

Private Sub Workbook_Open()
    Dim user As String
    user = Environ("Username")
    Select Case user
        Case "MeinBenutzername"
            DateiEntsperren
    End Select
End Sub

Sub DateiEntsperren()
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Unprotect pw
    Next
End Sub

Sub DateiSperren()
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=pw
    Next
End Sub
```

Nun werden folgende Anpassungen gemacht:

- "MeinKennwort" wird durch das eigene Kennwort ersetzt.
- "MeinBenutzername" wird durch den Windowsname des Anwenders ersetzt, der in der offenen Version arbeiten soll. Für mehrere Benutzer werden einfach die Benutzernamen hintereinander mit Komma getrennt aufgelistet.( `Case "Benutzername1", "Benutzername2"` )

Zum Schluss wird die Datei noch als Excel-Arbeitsmappe mit Makro (*.xlsm) gespeichert und steht zur Benutzung bereit.

Wird die Datei in Excel-Online (Sharepoint, MS Teams) geöffnet, ist sie für alle Benutzer mit dem Blattschutz gesperrt.

Wird die Datei in Excel geöffnet ergeben sich folgende Varianten:
- die Datei wird ohne Makros geöffnet. (Makros sind deaktiviert) Hier ist der Blattschutz für alle Nutzer gesetzt.

- die Datei wird von einem aufgelisteten Benutzer mit aktivierten Makros geöffnet. Der Blattschutz ist entfernt. Beim Schließen der Datei wird der Benutzer gefragt, ob er die Datei wieder schützen möchte.

- die Datei wird von andern Benutzer mit aktivierten Makros geöffnet. Der Blattschutz ist gesetzt.

Der Code im einzelnen:

Die nachfolgende Zeile sorgt dafür, das alle Variablen deklariert sein müssen.
```
Option Explicit
```

Hier wird das Blattschutz-Kennwort festgelegt.
```
Private Const pw As String = "MeinKennwort"
```

Die nachfolgende Routine wird beim Schließen der Datei durchlaufen.
```
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim user As String
    user = Environ("Username")
    Select Case user
        Case "MeinBenutzername"
            If MsgBox("Soll die Datei geschützt werden?", vbYesNo) = vbYes Then
                DateiSperren
                ActiveWorkbook.Save
            End If
    End Select
End Sub
```
 Hier wird auf den aktuellen Benutzer geprüft und ggf. die Prozedur `DateiSperren` ausgeführt, die den Blattschutz aktiviert.

Die nachfolgende Routine wird beim Öffnen der Datei durchlaufen.
```
Private Sub Workbook_Open()
    Dim user As String
    user = Environ("Username")
    Select Case user
        Case "MeinBenutzername"
            DateiEntsperren
    End Select
End Sub
```
 Hier wird auf den aktuellen Benutzer geprüft und ggf. die Prozedur `DateiEntsperren` ausgeführt, die den Blattschutz aufhebt.

 Die Prozedur `DateiEntsperren` hebt für alle Tabellenblätter den Blattschutz auf.
```
Sub DateiEntsperren()
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Unprotect pw
    Next
End Sub
```

 Die Prozedur `DateiSperren` setzt für alle Tabellenblätter den Blattschutz mit den Optionen Auswählen aller Zellen, Verändern der "offenen" Zellen und Nutzen des Autofilters.
```
Sub DateiSperren()
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, Password:=pw
    Next
End Sub
```

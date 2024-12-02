# Export von ShapeSheet Daten nach Excel

Um die ShapeSheet-Daten eines Visos-Shape formatiert nach Excel zu exportieren kann entweder die Datei [ExportShapeSheetData.vsdm](/samples/ExportShapeSheetData.vsdm)  genutzt werden oder man erstellt eine eigene Datei in der die beiden VBA-Modulen [mdl_ExportShapeData.bas](/code/mdl_ExportShapeData.bas) und [mdl_VisTools.bas](/code/mdl_VisTools.bas) eingebunden werden.

## Vorarbeiten

Zum VBA-Editor wechseln.

Dort `Extras - Verweise` auswählen.

![Screenshot Verweise öffnen](/sources/verweise1.png)

Die Verweise zu `Microsoft Excel 16.0 Object Library` und `Microsoft Scripting Runtime` auswählen. 

![Screenshot Verweise setzen](/sources/verweise2.png)


## ShapeSheet-Daten exportieren

Die Datei `ExportShapeSheetData.vsdm` öffnen.

Das Shape auswählen, dessen Shapedaten exportiert werden solle.

Auf der Registerkarte `Entwicklertools  - Code - Makros' öffnen (Alt+F8)

![Screenshot Ribbon](/sources/VisioMakros.png)

Dann bei `Makros in` die Datei `ExportShapeSheetData.vsdm` auswählen und das Makro ausführen.

![Screenshot Makrowahl](/sources/VisoMakroWahl.png)

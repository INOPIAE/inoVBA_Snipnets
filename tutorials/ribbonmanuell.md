# Das Ribbon manuell anpassen

Wenn das Ribbon nicht mit [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) erstellt bzw. angepasst werden kann.

Zuerst wird die Datei-Endung im Windows Explorer von auf zip geändert.

Anschließend wird die Zip-Datei geöffnet.

Kopieren Sie die Datei _rels/.rels auf den Desktop.

Die Datei wird angepasst in dem der folgende Code in Gruppe Relationships eingefügt.

```
<Relationship Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI14.xml" Id="Rac7cd3273f444e97" />
```

Der Pfad im Target muss auf die folgende Datei verweisen.

Anschließend wird die geänderte Datei in die Zip-Datei zurückkopiert.

Als nächste wird ein Ordner customUI angelegt.

In diesem Ordner wird eine XML-Datei mit dem Namen `customUI14.xml` erstellt.

## Beispielinhalt

```
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
	<ribbon>
		<tabs>
			<tab id="customTab" label="Mein Beispiel" insertAfterMso="TabHome">
				<group id="customGroup" label="Meine Tools">
					<button id="customButton1" label="Version"  onAction="rbVersion" />
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>
```

Anschließend wird der Ordner in die ZIP-Datei kopiert.

Mehr Infos zur Ausgestaltung von RibbonX findet sich hier [https://www.rholtz-office.de/ribbonx/sichtbarkeit-der-elemente](https://www.rholtz-office.de/ribbonx/sichtbarkeit-der-elemente).

Eine Beispiel Datei findet sich hier [RibbonDemo.xlsm](../samples/RibbonDemo.xlsm).
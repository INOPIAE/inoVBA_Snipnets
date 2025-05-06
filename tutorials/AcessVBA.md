# Access VBA

Hier geht es um kleine Hilfsfunktionen zu Access VBA.

Mit dem Module [mdl_ExportSQL.bas](/code/mdl_ExportSQL.bas) lassen sich aus einer Access Datenbank die Tabellenstruktur und der Inhalt einer Tabelle als SQL_String extrahieren.

Zum Export der Tabelledefinition wird die Funktion `ExportCreateSQLToFile` aufgerufen.
```
ExportCreateSQLToFile "TabellenName", "PfadZuEXportDatei"
```

Zum Export der Tabelledefinition wird die Funktion `ExportTableAsSQL` aufgerufen.
```
ExportTableAsSQL "TabellenName", "PfadZuEXportDatei"
```

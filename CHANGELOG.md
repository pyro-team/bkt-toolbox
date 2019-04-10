# Unreleased

## ADDED
  * BKT: Event-Handling, bspw. zum Ausführen von Aktionen bei Start oder Wechsel eines Fensters
  * PPT-Toolbar: Shape-Format austauschen
  * PPT-Toolbar: Breite/Höhe setzen auf Durchschnitt aller gewählten Shapes
  * PPT-Toolbar: Anfasser-Werte aller Shapes angleichen (bspw. gleicher Kurvenradius)
  * PPT-Toolbar: Vereinfachte UI für Funktionen des Erweiterten Anordnens
  * PPT-Toolbar: Shapes ganz einfach in eigene Favoriten-Shape-Library hinzufügen
  * PPT-Toolbar: Flag Library mit gleichem Seitenverhältnis hinzugefügt
  * PPT-Toolbar: Shape-Abstand von verschiedenen Kanten einstellen (im Spinner-Menü)
  * PPT-Toolbar: Agenda-Tab (wird bei Agenda-Slides sichtbar)
  * PPT-Toolbar: Verknüpfte-Shapes-Tab (wird bei Auswahl von verknüpftem Shape sichtbar)
  * PPT-Toolbar: Auswahlmenü für Hintergrund- und Linien-Transparenz
  * PPT-Toolbar: Widescreen-Version der Toolbar als Feature-Modul aktivierbar
  * PPT-Toolbar: Kleines Popup-Fenster für Wechselshapes hinzugefügt
  * PPT-Toolbar: Funktion zum Übertragen von Bildausschnitt/-zuschnitt sowie Tabellengrößen
  * PPT-Toolbar: Diverse Erweiterungen für verknüpfte Shapes (Eigenschaften für Suche, Toleranz, Anzahl Folien, etc.)
  * PPT-Toolbar: [BETA] Werkzeugleiste für schnelles Farbbearbeitung und -auswahl (QuickEdit-Feature)
  * PPT-Toolbar: [BETA] Modul für Definition von bis zu 5 benutzerdefinierten Formaten
  * PPT-Toolbar: [EXPERIMENTELL] Mini-Popup markiert Master-Shape (zuerst bzw. zuletzt markiertes Shape)
  * PPT-Toolbar: [EXPERIMENTELL] Implementierung von Popup-Fenstern für spezielle Shapes

## CHANGED
  * BKT: Update auf aktuelle IronPython-Version 2.7.9
  * BKT: Migration vieler Dialoge auf WPF
  * BKT: Diverse Aufräumarbeiten und Restrukturierungen, um die Entwicklung neuer Features zu ermöglichen (bspw. WPF-Dialoge)
  * BKT: Diverse Performance-Optimierungen, bspw. Callback in C#, Cache für Image-Resources
  * PPT-Toolbar: Kleinere Verbesserung bei Shape-Table Funktionen
  * PPT-Toolbar: Funktion zur Shape-Selektion funktioniert nun innerhalb gruppierter Shapes
  * PPT-Toolbar: Auswahl erstes/letztes Shape als Mastershape (bei Erweitertes Anordnen) wird nun gespeichert
  * PPT-Toolbar: Flag Library aktualisiert
  * PPT-Toolbar: Position und Größen Spinner (Seite 2) sind nun leer bei unterschiedlichen Werten (analog Standard-Spinner auf Seite 1)
  * PPT-Toolbar: Swap Funktion mit Einstellung der Shape-Ecke zum tauschen
  * PPT-Toolbar: Harvey-Moon Tab nun nicht mehr in Kontext-Tabs da nicht zuverlässig
  * PPT-Toolbar: UI für Anordnung auf Paragraphen/Tabellen/Shapes verbessert
  * PPT-Toolbar: Menüstruktur zum Einfügen "spezieller" Shapes aufgeräumt und einige Features umgezogen
  * PPT-Toolbar: Veränderte Logik für Wechselshapes, diese müssen nun erst in ein solches konvertiert werden

## FIXED
  * PPT-Toolbar: Kleinere Fehler und Performance-Verbesserungen für Shape-Selektion anhand mehrerer Attribute
  * PPT-Toolbar: Z-Order wurde bei "Ersetzen und Größe erhalten" nicht immer richtig gesetzt
  * PPT-Toolbar: Größen-Sprinner auf Seite 2 funktioniernen nun auch mit rotierten Shapes
  * PPT-Toolbar: Chartlib Thumbnails werden nun korrekt generiert und Menü schließt sich nicht mehr


# Release 2018-03-29

## ADDED
  * PPT-Toolbar: Wechsel-Shapes Funktion ermöglicht schnellen Wechsel von Status-Shapes, wie Ampeln oder Skalen
  * PPT-Toolbar: Einfügen einer Likert-Scale als Wechsel-Shape
  * PPT-Toolbar: Funktion um Bilder transparent zu machen (indem diese als Füllung für Shapes gesetzt werden)
  * PPT-Toolbar: Funktion um Zwischenablage auf mehrere Folien einzufügen
  * PPT-Toolbar: Horizonal/Vertikal stapeln (Button neben Spinner), Rotation zurücksetzen (Button neben Spinner)
  * Visio-Toolbar: Diverse neue Funktionen und Erweiterungen

## CHANGED
  * PPT-Toolbar: Folien-Thumbnails für Inhalt der Folie erstellen (ohne Platzhalter wie Kopf- und Fußzeilen)
  * PPT-Toolbar: Notizen-Farbe zurücksetzen
  * PPT-Toolbar: Spinner für rechten Einzug und Checkboxen für Word-Wrap und Auto-Size (auf Seite 2)
  * PPT-Toolbar: Senacor-Shape-Library nutzt nun Wechsel-Shapes für Ampeln und Likert-Scales
  * PPT-Toolbar: Überarbeitete Flaggen-Library (kein Vektor-Format wegen schlechtem Rendering in PPT)
  * PPT-Toolbar: UX vom Circlify-Feature angeglichen an Shape-Table-Feature

## FIXED
  * BKT: Kompatibilität mit Office 64-bit
  * PPT-Toolbar: Fix für sich schließene Chart Library beim ersten Öffnen
  * Visio-Toolbar: Fix für Sichtbarkeit der Ribbon-Tabs nach Start
  * Visio-Toolbar: Fix für nicht mehr funktionierende Rückgängig-Funktion
  * XLS-Minirechner: Formular Layout bei hohen DPI-Werten korrigiert


# Release 2017-12-19

## ADDED
  * BKT: Richtiger Installer, der alle Installationsschwierigkeiten lösen sollte
  * PPT-Toolbar: Verknüpfte Shapes: Shapes folienübergreifend verknüpfen und nachträglich ausrichten oder angleichen
  * PPT-Toolbar: Unicode-Symbole einfügen, insb. geschütztes Trennzeichen
  * PPT-Toolbar: Markierung umkehren
  * PPT-Toolbar: Standardpositionen für Repositionierung und Erweitertes Anordnen
  * PPT-Toolbar: Aufzählungszeichen korrigieren, Symbol und Farbe ändern
  * PPT-Toolbar: Entferne doppelte Leerzeichen und Kommentare
  * PPT-Toolbar: Selektion anhand diverser Shape-Eigenschaften wie Schriftfarbe, Größe, etc.
  * PPT-Toolbar: Harvey-Ball Tab um einige Funktionen ergänzt
  * PPT-Toolbar: Referenz von Folien-Thumbnails öffnen und Dateireferenz ersetzen
  * PPT-Toolbar: Agenda-Funktion kann Abschnitte je Agendapunkt einfügen
  * PPT-Toolbar: "Fine-Tuning" funktioniert nun auch für gebogene/gewinkelte Verbinder und andere Shapes
  * PPT-Toolbar: Diagramm-Dimensionen auf andere Diagramme übertragen
  * PPT-Toolbar: Spinner für Hintergrund- und Linien-Transparenz sowie Liniendicke

## CHANGED
  * PPT-Toolbar: Mehrere Folien gleichzeitig als aktualisieren Thumbnail kopieren und einfügen
  * PPT-Toolbar: Beim Aktualisieren von Thumbnails werden fehlerhafte Thumbnails markiert
  * PPT-Toolbar: Einheitlicher benutzerdefinierter Bereich für Positionierung, Erweitertes Anordnen und Shape-Tabellen
  * PPT-Toolbar: Im Tab "Harvey Balls" kann der Füllstand relativ zueinander angepasst werden im Spinner mit Alt-Taste
  * PPT-Toolbar: Große Schritte bei Spinner-Elementen über Shift-Taste
  * PPT-Toolbar: Text-Funktionen in Text-Menü zusammengefasst
  * PPT-Toolbar: Texte löschen oder durch Platzhalter (Lorem Ipsum) ersetzen geht nun auch für Tabellen
  * PPT-Toolbar: Folien-Thumbnails verwenden nun relative Pfade
  * PPT-Toolbar: Tracker auf Folien verteilen
  * PPT-Toolbar: Text/Platzhalter-Funktionen in Menü "Slidedeck aufräumen" sowie "Textoperationen" verschoben
  * PPT-Toolbar: UI-Verbesserungen von Shape-Tabellen Funktionen
  * Visio-Toolbar: Kleinere Funktionserweiterungen
  * BKT: Kleinere Performance-Verbesserungen

## FIXED
  * XLS-Toolbar: Diverse kleinere Fehler behoben
  * PPT-Toolbar: Kleinere Fehler beim Anordnen von Shapes auf Paragraphen behoben
  * PPT-Toolbar: Funktion Agenda-Update ordnet bestehende Folien besser zu
  * PPT-Toolbar: Diverse Spinner funktionieren nun besser innerhalb von Tabellenzellen


# Release 2017-10-11

## ADDED
  * PPT-Toolbar: Thumbnail-Folienreferenz ersetzen
  * PPT-Toolbar: Tracker auf allen Folien ausrichten
  * PPT-Toolbar: Kreissegmente erzeugen
  * PPT-Toolbar: Shapes kreisförmig anordnen
  * PPT-Toolbar: Shapes vervielfachen
  * PPT-Toolbar: Schnellwahl-Buttons bei "Erweitertes Anordnen"
  * PPT-Toolbar: Shape-Position und -Größe tauschen
  * PPT-Toolbar: Pfeilrichtung umkehren (im Kontextmenü für Pfeile/Verbinder)
  * XLS-Toolbar: Liste aller Kommentare, bedingten Formatierung und Dokumenteneigenschaften erstellen
  * XLS-Toolbar: Konvertieren von Text in echtes Datum
  * XLS-Toolbar: Sortieren von Blättern
  * XLS-Toolbar: Zwischenspeichern einer Selektion

## CHANGED
  * PPT-Toolbar: Shift/Strg nutzen um "Gleiche Höhe/Breite" und "Shapes tauschen" auf anderen Modus umzuschalten
  * PPT-Toolbar: Kontextmenü für Folien-Thumbnails
  * PPT-Toolbar: UK Englisch zu Sprachauswahl hinzugefügt
  * PPT-Toolbar: Optionen bei "Gleiche Höhe/Breite"
  * PPT-Toolbar: Diverse Verbesserung bei "Shapes als Tabelle anordnen"
  * PPT-Toolbar: Diverse Verbesserungen der Toolbox-UI
  * XLS-Toolbar: Vorschau-Funktion (bspw. bei Formel anwenden) verbessert
  * XLS-Toolbar: Platzhalter bei Text voranstellen/anhängen
  * XLS-Toolbar: "Markierung umkehren" verbessert

## FIXED
  * BKT: Fehler bei Leerzeichen im BKT-Pfad behoben
  * PPT-Toolbar: Bugfix für "Nummerierung hinzufügen"
  * PPT-Toolbar: Agenda berücksichtigt Highlight-Farbe
  * XLS-Toolbar: Diverse kleinere Probleme behoben


# Release 2017-09-07

## ADDED
 * PPT-Toolbar: Chart-Library
 * PPT-Toolbar: Shape-Library
 * PPT-Toolbar: Harvey Balls
 * PPT-Toolbar: Tracker
 * PPT-Toolbar: Prozesspfeile mit Header
 * PPT-Toolbar: Shapes horizontal/vertikal teilen
 * PPT-Toolbar: Sharing von individuellen Libraries und Funktionen zwischen CSTs
 * PPT-Toolbar: Thumbnail-Generator


# Release 2017-03-13
Beta-Release


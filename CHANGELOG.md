# Release 2024-01-10 (3.0.1)

  * Font Awesome updated to 6.5.1
  * Update definition for Material Icons and added newer Material Symbols
  * Added function to exclude icon fonts


# Release 2023-08-31 (3.0.0)

  * Switch to IronPython 3.4.1 (avoiding issues with virus scanners)
  * Moved most logic into dedicated models to increase lazy loading
  * Fixed updating of multiple selected thumbnails via popup
  * Excel: Added button to unfreeze UI in case a previous function cased an exception


# Release 2023-08-30 (2.7.5)

  * Added folder icons for chart and shape libraries
  * Added feature to make PPT Thumbnails clickable in presentation mode (hyperlink to slide)
  * Added feature to replace all PPT Thumbnails with regular picture (remove all meta data)
  * Added feature to toggle and remove PPT Notes from all selected slides
  * Added new options to style harvey moons
  * Font icons now support Segoe Fluent Icon Font
  * Font Awesome updated to 6.4.2
  * Fixed icons mixed up in arrage easy group (thanks to robertzb)
  * Fixed adding multiple shapes to a shape library didn't work
  * Fixed umlauts in install path / config paths didn't work
  * Fixed handling of animations in group helper functions and state-shapes


# Release 2022-09-03 (2.7.4)

  * Restructuring of context menus to load various function dynamically based on selected shape
  * Added support to edit parameters of process shapes and circular segments
  * Added support for Font Awesome 6 font icons
  * Added toggle to disable conversion of shape adjustment values
  * Added function to add current slide/shape to any chart/shape library
  * Improved "copy-paste dimensions of table/pictures" to work with multiple shapes at once
  * Improved various icons to be better visible in dark Office themes
  * Fixed update of multiple thumbnails (e.g. all thumbnails in the presentation)
  * Fixed some errors when updating an agenda (e.g. when deleting textbox or file is read-only)
  * Fixed PowerPoint crash when table is selected
  * Fixed some issues in functions using the clipboard (e.g. split paragraphs)
  * Updated IronPython to 2.7.12
  * Excel: Added function to list all data validations
  * Excel: Added function to consolidate sheets by column (extension of row-consolidation function)
  * Excel: Added small dialog to show hidden and very hidden sheets
  * Excel: Fixed Excel workbook and sheet consolidation


# Release 2021-01-29 (2.7.3)

  * Added option switch for auto-com-release (potential workaround for PPT restart bug)
  * Added function in Excel toolbar to remove all hidden columns or rows
  * Improved circula segments dialog, now radius can be defined
  * Improved usability of shape scale dialog
  * Improved "Paste and Distribute" by ordering selected shapes top-left to bottom-right
  * Improved Excel "Apply formula" function with better preview and text function support (e.g. UPPER)
  * Fixed bug in ChartLib creating new libraries
  * Fixed FontAwesome gallery not loading (split gallery as more than 1000 gallery items not supported by PPT)
  * Fixed error inserting shape tables
  * Fixed top position when using "Text out of shape" function
  * Fixed crash in Excel when whole row or column is selected and certain functions are applied
  * Known Issue: Random PPT restart after closing


# Release 2020-10-29 (2.7.2)

  * Rename "Master" to "Reference" or "Main" across the whole UI
  * Added functions to linked shapes: Select reference shape, skip between linked shapes, set link ID for each individual selected shape
  * Added function to select shapes with tags
  * Added font-icon dropdown to search group
  * Added option to remove empty sections when sending slides
  * Added functions to sort Z-Order (and therefore selection) by position top-down or left-right
  * Added compact font/paragraph group as new default also for 3-page layout
  * Added slide synchronization function to equalize multiple slides at once (e.g. template slides)
  * Added feature to copy shape texts and to paste and distribute paragraphs on all selected shapes
  * Improved handlings of chart textframes
  * Improved shape names of various BKT shapes
  * Improved linked shapes functionalities now support multiple selected linked shapes at once
  * Improved statistics: now shown in dialog window with additional data
  * Improved UI settings editor with scroll-viewer
  * Improved traffic light to recognize color changes
  * Improved likert scale is now easier adoptable (not backwards compatible)
  * Improved layout of ppt_notes group
  * Fixed errors for certain chart types when changing language or showing paragraph spinners
  * Fixed split by section (in consolidation-split feature)
  * Fixed chartlib creating thumbnails for potx files that were not required
  * Fixed QuickEdit first-start popup showing up every time
  * Update FontAwesome5 definition to v5.15.0
  * Known Issue: Random PPT restart after closing


# Release 2020-07-30 (2.7.1)

  * Added quickedit dark theme
  * Added function to open slideshow in window or fullscreen
  * Added option to remove hidden slides when sending slides
  * Added checkbox stateshape
  * Added sticker with custom text
  * Added support for transparent background in chartlib thumbnails
  * Added function to open in explorer
  * Added function to toggle master shape visibility
  * Added multiple traffic light styles
  * Added shape scaling function
  * Added dialog to apply multiple slide cleanup functions at once
  * Added function to apply theme from file
  * Added support for template files (*.potx) in chartlib
  * Added format sync via context menu and format painter button if multiple shapes are selected
  * Added more shape types to shape adjustments
  * Added "Format Syncer" that allows to sync format for all selected shapes (master can be the one right clicked)
  * Added minified font and text group (default in 2-page layout)
  * Added function to reset adjustment values
  * Added function to choose aspect ratio from picklist or reset to default ratio
  * Added rows support to process chevrons
  * Excel: Added function to translate english formula to local and vice versa
  * Excel: Added function to apply theme from other file and import/export font schema
  * Improved shape statistics: supertip show details about found numbers, result can be copied to clipboard, correct processing of decimals for various countries
  * Improved text replacement with additional placeholders and proper template engine
  * Improved arrange behavior if master/reference shape is deleted
  * Improved various WPF dialogs, i.e. shape selection dialog
  * Improved chartlib may paste on multiple selected slides
  * Improved chartlib insertion of slides
  * Improved shape connector support for chevron, pentagon, hexagon and oval
  * Fixed shape statistics did not consider negative numbers
  * Fixed logging was processed even if logging was turned on
  * Fixed error in settings shelve auto repair
  * Fixed use of slide index vs slide number in multiple places
  * Fixed maintaining lock aspect ratio for all stateshapes (incl. likert scale)
  * Fixed sending slides of unsaved presentation
  * Fixed popup positioning with multiple monitors (especially negative coordinates)
  * Fixed charlib refresh not working due to premature comrelease
  * Fixed getting correct content area size for slide master
  * Fixed multiple issues related to getting slide in slide master view
  * Fixed charlib not updating all libraries
  * Fixed iteration of non-contiguous selection of cells in Excel
  * Fixed error message if ribbon id is unknown (e.g. contact popup in Word)


# Release 2020-05-17 (2.7.0)

  * Added option to link shapes by its shape names
  * Added function to merge rows or columns in excel
  * Added support to change styling (fill, line, text, height) of agenda selector incl. support of theme colors
  * Added agenda popup dialog
  * Added function to edit languages in languages setter out of several languages
  * Added font icon search
  * Added double click for harvey moons, thumbnails, state shapes and traffic light
  * Added search for new release (incl. scheduled automatic search)
  * Added icon font IcoMoon-Free
  * Added slide export as PDF and PNG (before only as PPTX)
  * Added function to copy slide as high quality PNG into clipboard
  * Replaced dev module with devkit feature folder (+many new dev features)
  * Slide deletion when sending selected slides works much faster now
  * Fixed calculation of shaded theme colors in all color galleries and quickedit
  * Fixed harvey ball outline not always working
  * Fixed harvey ball without background is still selectable in fill area
  * Fixed color gallery not updating when color theme is changed (e.g on switch between presentations)
  * Fixed "select by color" in Excel not working properly
  * Fixed thumbnails function when path contains special characters
  * Fixed update of headered pentagon/chevron for flipped and rotated shapes
  * Fixed restoring of master agenda slide based on regular agenda slide when master is deleted
  * Fixed problems with ampersand in paths or chartlib filenames
  * Fixed python import to use absolute import as in python 3
  * Fixed deletion of temporary images files created for various functions
  * Fixed use of FileDialogs by using WinForms dialogs instead of internal office dialogs
  * Fixed local documents folder is now correctly identified if moved by user or OneDrive
  * Support for legacy annotations syntax is disabled by default
  * Updated to IronPython 2.7.10
  * Updated and fixed the unit tests (still more tests to add)
  * Excel: Tool to apply (count, split, replace) cells by regular expression (regex)


# Release 2019-12-13 (2.6.1)

  * Fixed gradients with theme colors in custom formats
  * Fixed removal of sections when sending selected slides
  * Fixed text selection not always working when adding notes (feature ppt_notes)
  * Fixed harvey balls inverted if pie shape is flipped (done by other toolbars)
  * Fixed duplication of font awesome 5 category menu on each opening
  * Added font awesome 4/5 category menu cache
  * Harvey balls will now use theme colors so that they are working with dark designs
  * Harvey ball background color can now be changed
  * In harvey ball tab setting value <0 or >100 is possible using modulo 100
  * Added harvey ball popup
  * Added modern harvey ball style (toggle in harvey tab)
  * Added function to remove unused designs
  * Added checkbox to remove unused slide layouts when sending selected slides


# Release 2019-11-07 (2.6.0)

  * Custom formats: Support for font gradients, font glow, font reflection, font shadow
  * Consolidation-Split-Feature now has photos2slides function and progress bar when exporting slides
  * Added toggle to use visual size and position to according spinners
  * Improved linked shapes features: apply certain actions and sync single properties such as center left point, sync text formatting
  * Improved text-to-shape function for shapes with background color
  * Fixes for inserting icons from Font Awesome 5
  * Make PPT-Toolbox UI customizable via settings dialog
  * Excel Toolbar: Refactoring of folder and file list generator with many improvements


# Release 2019-08-23 (2.5.3)

  * Popup for process chevrons shapes (add/remove chevrons)
  * Split paragraphs into shapes feature is ignoring empty paragraphs and uses better positioning of new shapes
  * Sticker and Underlined textbox support dark themes
  * Multiple sticker texts and sticker settings can be adjusted
  * Slide export and language settings are now in thumbnail menu
  * Edge autofixer can be configured to some extend
  * Feature to swap Z-Order and feature to place shape just behind/before shape
  * Refactoring and improvement of tracker-generator
  * Refactoring and improvement of shape-connector between 2 shapes (now supporting update/re-connect and rotated shapes)
  * Refactoring of circular arrangement to avoid small movements and allow to define angle of first shape and center shape
  * Fixes and configuration options for autofix edges feature
  * Process chevrons, stateshapes and ppt-thumbnails now properly support rotation
  * Small helper functions for swapping position and size
  * Allow to set anchor point for shape rotation
  * Added various text operations and language settings to shape context menus
  * Added function to add shape into group and recursively ungroup shapes
  * CustomFormats now properly supports connectors
  * Various small fixes


# Release 2019-07-05 (2.5.2)

  * Custom formats: Complete refactoring with gallery as selection and function for pickup-apply of individual formats
  * QuickEdit: 3 catalogs for own colors selectable (with mouse wheel); improvement with shades of the theme colors (use of shade index instead of brightness)
  * Symbol menu: Last used symbols, dynamic icon font menu, insert as shape or image, unicode font selectable
  * PPT-Thumbnails: Paste now also possible via "normal" copying (from OLE project, but only for one slide)
  * Extended (euclidian) distance: Configurable behavior when selecting more than 2 shapes
  * New feature for statistics (count, sum) of selected shapes 'ppt_statistics' (must be activated)
  * New feature to adjust the edges 'Edge autofixer' in the position menu
  * Shape distance to the left or top now works correctly
  * Arrangement in tables/paragraphs/shapes now significantly improved
  * Shape-Tables: Negative distances possible, function to distribute width/height equally
  * Storage of various settings in user settings (e.g. Global LocPin, ShapeDistance mode, Placeholder text, Recent symbols, etc.)
  * Function to remove unused master layouts, as well as remove all external links
  * Focus problem with popups fixed
  * Avoid error message by invalid return of get_selected_item_index
  * ColorGallery is now also available in Excel


# Release 2019-05-13 (2.5.1)

  * QuickEdit: UI enhancements
  * Chartlib: Library refresh is using threads, progress bar and cache
  * Unicode support for log files
  * Error handling for popups added
  * Improvement of the import cache incl. conflict handling


# Release 2019-04-18 (2.5)

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
  * PPT-Toolbar: Kleines Popup-Fenster für Wechselshapes, verknüpfte Shapes und Thumbnails hinzugefügt
  * PPT-Toolbar: Funktion zum Übertragen von Bildausschnitt/-zuschnitt sowie Tabellengrößen
  * PPT-Toolbar: Diverse Erweiterungen für verknüpfte Shapes (Eigenschaften für Suche, Toleranz, Anzahl Folien, etc.)
  * PPT-Toolbar: [BETA] Werkzeugleiste für schnelles Farbbearbeitung und -auswahl (QuickEdit-Feature)
  * PPT-Toolbar: [BETA] Modul für Definition von bis zu 5 benutzerdefinierten Formaten
  * PPT-Toolbar: [EXPERIMENTELL] Mini-Popup markiert Master-Shape (zuerst bzw. zuletzt markiertes Shape)

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


# Konfigurationsdatei

Hier ist einer Übersicht der Konfigurationsmöglichkeiten in der `config.txt`. Falls der Key nicht vorhanden ist, kann er einfach angelegt werden.

Key 					| Values/**Default**	| Explanation
--- 					| --- 					| ---
async_startup 			| True, **False** 		| Asynchroner Start: Lädt die UI verzögert, wodurch der PowerPoint-Start beschleunigt wird. [BETA-Funktion]
log_level				| CRITICAL, ERROR, **WARNING**, INFO, DEBUG | Mindestlevel für Logging.
log_write_file			| True, **False**		| Log-Datei `bkt-debug.log` und `bkt-debug-py.log` schreiben an/aus
log_show_msgbox			| True, **False**			| Log-Einträge als Messagebox anzeigen.
show_exception			| **True**, False		| Kritische Fehler als Messagebox anzeigen.
local_fav_path			| \~\Documents\BKT-Favoriten\		| Pfad zur Speicherung von BKT-Favoriten, bspw. Custom Formats, Farbleiste, Chartlib.
local_cache_path		| {INSTALLDIR}\resources\cache\		| Pfad zur Anlage von Cache-Dateien.
local_settings_path		| {INSTALLDIR}\resources\settings\	| Pfad zur Speicherung der Einstellungsdatenbank.
task_panes				| True, **False**	| Task Panes (Seitenleiste) de-/aktivieren. [BETA-Funktion]
use_keymouse_hooks		| **True**, False 	| Maus- und Tastaturevents verwenden, bspw. für Contextdialogs.
ppt_use_contextdialogs	| **True**, False 	| PowerPoint-Contextdialogs ein-/ausschalten.
ppt_hide_format_tab		| True, **False** 	| PowerPoint Format-Tab ein-/ausblenden, um die Wechsel zu dem Tab bei neuen Shapes zu verhindern.
excel_ignore_warnings	| True, **False** 	| Rückgängig-Warnmeldung in Excel nicht mehr anzeigen.

<dl>
  <dt>{INSTALLDIR}</dt>
  <dd>BKT Installationsverzeichnis. Üblicherweise `C:\Users\USERNAME\AppData\Local\BKT-Toolbox`</dd>
  <dt>~</dt>
  <dd>Das Benutzerverzeichnis. Üblicherweise `C:\Users\USERNAME`</dd>
</dl>
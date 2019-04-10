<img src="documentation/screenshot.png">

## Einführung

Die BKT besteht aus 2 Teilen, dem BKT-Framework und der BKT-Toolbox. Das Framework bietet eine einfach Möglichkeit, um Office-Addins für PowerPoint, Excel, Word, Outlook oder Visio in Python zu schreiben. Die BKT-Toolbox gibt es aktuell für PowerPoint, Excel und Visio.

Die PowerPoint-Toolbox fügt mehrere Tabs hinzu, die einen strukturierten Zugriff auf alle Standard-Funktionen von PowerPoint bieten, ergänzt um viele bisher fehlende Funktionen.

Die BKT wird von uns in der Freizeit entwickelt, daher können wir keinen Support anbieten oder auf spezielle Wünsche eingehen.

## Systemvoraussetzungen

Die BKT läuft unter Windows ab Office 2010 in allen aktuellen Office-Versionen. Eine Mac-Version ist nicht verfügbar, da die entsprechende Office-Schnittstelle (COM-Addin) im Mac-Office nicht von Microsoft angeboten wird.

## Bekannte Probleme

 * Da die IronPython-Umgebung erst hochfahren muss, verzögert sich der PowerPoint-Start mit aktiviertem Addin. Wir arbeiten daran, diesen Effekt etwas zu minimieren.
 * Wenn im Hintergrund eine PowerPoint-Präsentation in der geschützten Ansicht geöffnet ist, reagiert PowerPoint mit Toolbox unvorhersehbar, bspw. werden Selektionen nicht mehr richtig angenommen. Dies scheint ein PPT-Bug zu sein, den wir leider nicht fixen können.
 * Manchmal gibt es bei längerer Nutzung Performance-Probleme (Texteingaben werden verschluckt, Shape-Auswahl funktioniert nicht mehr richtig, ...) wenn man auf einem BKT-Tab ist. Leider hilft aktuell nur der Wechsel auf einen Standard-Tab (bspw. "Start"), oder ein PowerPoint Neustart. Wir untersuchen dieses Problem noch.

## Installation

Am einfachsten geht die Installation über das [Setup](https://github.com/mrflory/bkt-toolbox/releases/latest).

Alternativ kann man das Repository klonen und die Datei `installer\install.bat` ausführen. Nach einem Update muss die Datei ggf. neu ausgeführt werden.

***Hinweise:***

 * Es gibt ein separates Setup for Office 2010. Beim Klonen des Repositories muss vor der Installation die Datei `dotnet\build2010.bat`     ausgeführt werden, damit das Addin für Office 2010 kompiliert wird.
 * Die Business Kasper Toolbox ist nach Installation standardmäßig nur in PowerPoint aktiv, 
    jedoch auch in Excel, Outlook, Word und Visio verfügbar. Dort lässt sich die BKT über den
    Addin-Dialog aktivieren (Datei > Optionen > Add-Ins)
 * Über den Addin-Dialog lässt sich ferner das BKT-Dev-Plugin aktivieren. Dieses erlaubt
    Laden und Entladen des Addins zur Laufzeit der Office-Applikation

## Anwenderhandbuch

*In Arbeit*

## Entwicklerdokumentation

*In Arbeit*

[Übersicht der Einstellungen in der Konfigurationsdatei](config.md)

## Contributions

 * [IronPython](https://github.com/IronLanguages/ironpython2)
 * [Fluent.Ribbon](https://github.com/fluentribbon/Fluent.Ribbon)
 * [ControlzEx](https://github.com/ControlzEx/ControlzEx)
 * [MahApps.Metro](https://github.com/MahApps/MahApps.Metro)
 * [MouseKeyHooks](https://github.com/gmamaladze/globalmousekeyhook)
 * [InnoSetup](http://www.jrsoftware.org/isinfo.php)

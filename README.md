# BKT - Business Kasper Toolbox

<img src="documentation/screenshot.png">

## Introduction

The BKT consists of 2 parts, the BKT Framework and the BKT Toolbox. The framework offers an easy way to write office add-ins for PowerPoint, Excel, Word, Outlook or Visio in Python. The BKT Toolbox is currently available for PowerPoint, Excel and Visio.

The PowerPoint Toolbox adds multiple tabs that provide structured access to all of PowerPoint's standard features, plus many missing features.

The BKT is developed by us in our spare time, so we cannot offer support or respond to special requests.

### Language

Historically, the BKT was developed in German and unfortunately, we do not have the time to translate the whole toolbox. We hope that most functions are self-explanatory. If you have experience in multi-language python projects, feel free to support us.

### Documentation

Currently, we only have the [GitHub Wiki](https://github.com/mrflory/bkt-toolbox/wiki) with some English documentation for adventurous users and developers.

## System requirements

The BKT runs under Windows from Office 2010 in all current Office versions. A Mac version is not available because Microsoft does not offer the corresponding Office interface (COM add-in) in the Mac Office.

## Installation

The easiest way to install is via the [Setup](https://github.com/mrflory/bkt-toolbox/releases/latest) (only for Office 2013+).

Alternatively, you can clone the repository and run the `installer\install.bat` file. After an update, the file may need to be re-run.

***Notes:***

 * There is a separate setup for Office 2010. When cloning the repository, the file `dotnet\build2010.bat` must be executed before installation to compile the addin for Office 2010.
 * The Business Kasper Toolbox is only active in PowerPoint by default after installation, but also available in Excel, Outlook, Word and Visio. There, the BKT can be accessed via the Activate Addin dialog (File > Options > Add-Ins)
 * The Addin dialog can also be used to activate the BKT Dev Plugin. This allows loading and unloading the addin at runtime of the office application.

## Contributions

 * [IronPython](https://github.com/IronLanguages/ironpython2)
 * [Fluent.Ribbon](https://github.com/fluentribbon/Fluent.Ribbon)
 * [ControlzEx](https://github.com/ControlzEx/ControlzEx)
 * [MahApps.Metro](https://github.com/MahApps/MahApps.Metro)
 * [MouseKeyHooks](https://github.com/gmamaladze/globalmousekeyhook)
 * [InnoSetup](http://www.jrsoftware.org/isinfo.php)
 * [Google Material Icons](https://material.io/tools/icons/) & [Material Design Icons](https://materialdesignicons.com/)

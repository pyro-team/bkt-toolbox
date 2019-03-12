BKT Code Documentation
======================


Interaction of dotnet- and python-addin
***************************************

TODO



Overview modules
****************

============  ==================================================================  =================
Modul         Description                                                         Dependencies
============  ==================================================================  =================
**Implementation of addin-interface**
---------------------------------------------------------------------------------------------------
addin         Basic module of addin and interface to .Net-world                   *all*
boostrap      Helper-modul to bootstrap addin-object                              addin
\
                                                                                  
**Functionality for representation of controls and callbacks**
---------------------------------------------------------------------------------------------------
apps          Office-app specific objects and logic: app-callbacks etc.
callbacks     Definition of callback-types from MSOfficeXML and callback objects     
ribbon        Objects for MSOfficeXML-types (buttons, tabs, ...)                  callbacks, xml
\
                                                                                  
**Functionality for annotation-based definition of BKT-features**
---------------------------------------------------------------------------------------------------
annotation    Abstract classes for annotation commands                            callbacks, helpers
decorators    Definition of BKT-decorators for methods and classes                annotation, callbacks
factory       ControlFactory for creation of controls from annotated classes
\
                                                                                  
**Basic features**                                                                 
---------------------------------------------------------------------------------------------------
install       Install-module for BKT                                              helpers
ui            UI-capabilities of BKT: console, message box, etc.                  
\
                                                                                  
**Basic modules and helper-methods (BKT-independent)**                              
---------------------------------------------------------------------------------------------------
dotnet        Encapsulation of CLR references
helpers       Helper-modul for logging etc.
xml           XML-Factory-Wrapper for Linq                                        dotnet
============  ==================================================================  =================





Implementation of addin-interface
*********************************

addin
-----
Basic module of addin and interface to .Net-world

.. automodule:: bkt.addin
   :members:


bootstrap
---------
Helper-modul to bootstrap addin-object
.. automodule:: bkt.bootstrap
   :members:


Functionality for representation of controls and callbacks
***********************************************************

apps
----
Office-app specific objects and logic (for PowerPoint, Visio, ...): e.g. app-callbacks

.. automodule:: bkt.apps
   :members:


callbacks
----------
Definition of callback-types from MSOfficeXML and callback objects

.. automodule:: bkt.callbacks
   :members:


ribbon
------
Objects for MSOfficeXML-types (buttons, tabs, ...)
Definition of special classes, e.g. spinner box, color galery

.. automodule:: bkt.ribbon
   :members:



Functionality for annotation-based definition of BKT-features
*************************************************************

annotation
----------
Abstract classes for annotation commands

.. automodule:: bkt.annotation
   :members:

    
decorators
----------
Definition of BKT-decorators for methods and classes

.. .. automodule:: bkt.decorators
..    :members:

    
factory
-------
ControlFactory for creation of controls from annotated classes

.. .. automodule:: bkt.factory
..    :members:


Basic features
**************

install
-------
Install-module for BKT (initialization config, customize registry).

.. automodule:: bkt.install
   :members:


ui
---
UI-capabilities of BKT: console, message box, etc.

.. automodule:: bkt.ui
   :members:



Basic modules and helper-methods (BKT-independent)
**************************************************


dotnet
------
Encapsulation of CLR references

.. .. automodule:: bkt.dotnet
..    :members:

helpers
-------
Helper-modul for logging etc.

.. .. automodule:: bkt.helpers
..    :members:

xml
----
XML-Factory-Wrapper for Linq

.. .. automodule:: bkt.xml
..    :members:
   



Build .Net-Addin
****************

Add the following folder to the PATH vaiable: %SystemRoot%\Microsoft.Net\Framework\v4.0.30319

Replace the version number to whatever .Net framework version you are using.
   

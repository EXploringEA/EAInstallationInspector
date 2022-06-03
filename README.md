# EAInstallationInspector

Source code for the [EXploringEA.com's](http://EXploringEA.com "EXploringEA") EA Installation Inspector 

This is a windows application which displays registry entries associated with Sparxsystems EA Addins.
For more information review the document **eaInstallationInspectorInformationV5.pdf** located under the resources directory.

Current release is V6 - the current development version is V7 which has the functionality of previous versions with support for the 64-bit version of EA
As of May 2022 - this is an alpha release as there are known issues I need to address when both 32-bit and 64-bit versions of the same AddIn are installed.  I'm working on these and of course more testing is required across a range of platforms.  Main issue is that I have to sort out the detailed for the use case when both 32-bit/64-bit versions are installed.  Previously I looked across many hives to find details for a specific addin, not least as sometimes information gets placed in the wrong place!  For the 32/64-bit case I may need to revert to much stricter compliance with what the installer should do but need to think about the situations that may not be coevred.  Any input and issues found please let me know.

A pre-built .exe and help file are provided under the folders V6 or V7 (dev version!) just below the root level.

NOTE: This code is released under the terms of the GPL3 licence agreement.

For more information about the EA Installation Inspector and AddIn Installation errors check out - http://tools.exploringea.co.uk/index.php?n=EaInspector.Overview

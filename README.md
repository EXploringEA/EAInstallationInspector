# EAInstallationInspector

Source code for the [EXploringEA's EA Installation Inspector](https://exploringea.co.uk/tools/index.php/EAInspector/EAInspector.html) 

This is a windows application which displays registry entries associated with Sparxsystems EA Addins.

For more information review the document **eaInstallationInspectorInformationV7007.pdf** located under the resources directory.

Current release is V7.0.0.9 297SEP2022 - this release adds a green back color on the EA version and location to indicate which version is providing the COM server to the EA API.
NOTE: Only a single version of EA can be registered for COM 

The aim has been to provide a version that includes basic features and then update as more refinements are identified in relation to making it useful for both 32-bit and 64-bit work.

I continue to tinker with the code over the next few weeks as it is tested more and if you find any issues or have suggestions on potential improvements please let me know.

Please NOTE: that the colours for indicating the condition of an AddIn has changed.  See https://exploringea.co.uk/tools/index.php/EaInspector/EaInspector.html for details - I'll try and keep the site updated - the documents are lagging behind at the moment,

A pre-built .exe and help file are provided under in folders below the root.  

* V7007 contains the current 32-bit/64-bit (note you now need the additional dll's in the same directory as the exe file if you wish to look at dll meta data)

NOTE: This code is released under the terms of the GPL3 licence agreement and is used at yoru own risk.

For more information about the EA Installation Inspector and AddIn Installation errors check out - https://exploringea.co.uk/index.php/EaInspector/

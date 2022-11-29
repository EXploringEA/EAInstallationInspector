' Copyright (C) 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
'
'Imports eaInstallationInspector.RegistryContents
Imports Microsoft.Win32
'Imports eaInstallationInspector.ClassInformation
Public Class AddInInformation


    ' Sparx subkey folders - which exist in both HKCU and HKLM
    ''' <summary>
    ''' The sparx keys - 32-bit  
    ''' </summary>
    Friend Const SparxKeys32 As String = "Software\Sparx Systems\EAAddins"
    ' Private SparxKeys32 As String = "SOFTWARE\Sparx Systems\EAAddins"

    ''' <summary>
    ''' The sparx keys - 32-bit addin for 64-bit OS
    ''' </summary>
    Friend Const SparxKeysWOW32 As String = "Software\Wow6432Node\Sparx Systems\EAAddins"


    ' 64-bit AddIns

    ''' <summary>
    ''' The sparx keys - x64
    ''' </summary>
    Friend Const SparxKeys64 As String = "Software\Sparx Systems\EAAddins64"

    ' Exact registry locations for Sparx keys
    ' 32-bit AddIns

    ''' <summary>
    ''' HKCU Keys
    ''' </summary>
    Friend Const eaHKCU32AddInKeys As String = "HKEY_CURRENT_USER\" & SparxKeys32

    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const eaHKLM32AddInKey64 As String = "HKEY_LOCAL_MACHINE\" & SparxKeysWOW32

    Friend Const eaHKLM32AddInKeys As String = "HKEY_LOCAL_MACHINE\" & SparxKeys32
    ' 64-bit AddIns
    ''' <summary>
    ''' HKCU Keys
    ''' </summary>6
    Friend Const eaHKCU64AddInKeys As String = "HKEY_CURRENT_USER\" & SparxKeys64
    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const eaHKLM64AddInKeys As String = "HKEY_LOCAL_MACHINE\" & SparxKeys64

#Region "Orignal approach"


    '''' <summary>
    '''' Gets a list of addin entries from Windows Registry
    '''' - start with HKCU as they have precendence over the HKLM
    '''' Assume that the registry keys in the following locations identify 32-bit AddIns
    ''''	"HKCU\SOFTWARE\Sparx Systems\EAAddins" - HKCU
    ''''	"HKLM\ SOFTWARE\Sparx Systems\EAAddins" - HKLM, although if the addin Is running on a 64-bit operating system this would be "HKLM\ SOFTWARE\WOW6432NODE\Sparx Systems\EAAddins"  HKLMWow
    ''''	 Assume that the registry keys in the following locations identify 64-bit AddIns
    ''''	 "HKCU\SOFTWARE\Sparx Systems\EAAddins64" - HKCU64
    ''''	 "HKLM\ SOFTWARE\Sparx Systems\EAAddins64" – HKLM64
    '''' </summary>
    '''' <returns>List of AddIn entries</returns>
    'Function getListOfEAAddinEntries() As ArrayList

    '    ' for all of the keys which are not nothing we add them to an array
    '    Dim myAddInEntries As New ArrayList
    '    For Each entry As AddInEntry In getListof32BitHKCUAddinEntries()
    '        myAddInEntries.Add(entry)
    '    Next

    '    For Each entry As AddInEntry In getListof32BitHKLMAddinEntries()
    '        myAddInEntries.Add(entry)
    '    Next

    '    For Each entry As AddInEntry In getListof64BitHKCUAddinEntries()
    '        myAddInEntries.Add(entry)
    '    Next
    '    For Each entry As AddInEntry In getListof64BitHKLMAddinEntries()
    '        myAddInEntries.Add(entry)
    '    Next

    '    Return myAddInEntries

    'End Function


    ''' <summary>
    ''' 32 bit AddIns when install for current users only
    ''' Sparx keys (32-bit apps) will be in HKCU\Software\SparxSystems\EAAddIns
    ''' </summary>
    ''' <returns>List of 32-bit AddIns for current user</returns>
    Function getListof32BitHKCUAddinEntries() As ArrayList
        Dim my32HKCUAddInEntries As New ArrayList
        Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys32) ' get 32-bit AddIn for the current user
        If myCUKey IsNot Nothing Then
            For Each pEntryKey In myCUKey.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey
                AddInInfo.ClassName = Registry.GetValue(eaHKCU32AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
                AddInInfo.SparxAddinLocation = AddInEntry.cHKCU32
                my32HKCUAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my32HKCUAddInEntries
    End Function
    ''' <summary>
    ''' 32 bit AddIns when installed local machine (for all users)
    ''' Sparx keys (32-bit apps) can be in HKLM\Software\SparxSystems\EAAddIns
    ''' BUT NOTE for 64-bit operating systems the keys will be located in HKLM\Software\Wow6432Node\SparxSystems\EAAddIns
    ''' </summary>
    ''' <returns>List of 32-bit Addins for local machine</returns>
    Function getListof32BitHKLMAddinEntries() As ArrayList

        Dim my32HKLMAddInEntries As New ArrayList
        Dim myLMKey As RegistryKey = IIf(Environment.Is64BitOperatingSystem, Registry.LocalMachine.OpenSubKey(SparxKeysWOW32), Registry.LocalMachine.OpenSubKey(SparxKeys32))
        If myLMKey IsNot Nothing Then
            For Each pEntryKey In myLMKey.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey ' registry location for addin name 
                AddInInfo.ClassName = Registry.GetValue(eaHKLM32AddInKey64 & cBackSlash & pEntryKey, "", cNotSet)
                AddInInfo.SparxAddinLocation = AddInEntry.cHKLM32 ' If(Environment.Is64BitOperatingSystem, AddInEntry.cHKLM32Wow, AddInEntry.cHKLM32)
                my32HKLMAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my32HKLMAddInEntries
    End Function
    ''' <summary>
    ''' 64 bit AddIns when install for current users only
    ''' HKCU\Software\SparxSystems\EAAddins64
    ''' </summary>
    ''' <returns>List of current user 64-bit addin entries</returns>
    Function getListof64BitHKCUAddinEntries() As ArrayList
        Dim my64HKUCAddInEntries As New ArrayList
        Dim myCUKey64 As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys64)
        If myCUKey64 IsNot Nothing Then
            For Each pEntryKey In myCUKey64.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey
                AddInInfo.ClassName = Registry.GetValue(eaHKCU64AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
                AddInInfo.SparxAddinLocation = AddInEntry.cHKCU64
                my64HKUCAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my64HKUCAddInEntries
    End Function
    ''' <summary>
    ''' 64 bit AddIns when installed for all users
    ''' HKLM\Software\Wow6432Node\SparxSystems\EAAddIns64
    ''' </summary>
    ''' <returns>List of 64-bit addins for Local Machine (all users)</returns>
    Function getListof64BitHKLMAddinEntries() As ArrayList
        Dim my64KLMAddInEntries As New ArrayList
        Dim myLMKey64 As RegistryKey = Registry.LocalMachine.OpenSubKey(SparxKeys64) '"SOFTWARE\Sparx Systems\EAAddins64") 'SparxKeys)
        If myLMKey64 IsNot Nothing Then
            For Each pEntryKey In myLMKey64.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey ' registry location for addin name 
                AddInInfo.ClassName = Registry.GetValue(eaHKLM64AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name CLSID name
                AddInInfo.SparxAddinLocation = AddInEntry.cHKLM64
                my64KLMAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my64KLMAddInEntries
    End Function
#End Region
    '#Region "Sparx Entries"


    '    ' SPARX entries

    '    ''' <summary>
    '    ''' 32 bit AddIns when install for current users only
    '    ''' Sparx keys (32-bit apps) will be in HKCU\Software\SparxSystems\EAAddIns
    '    ''' </summary>
    '    ''' <returns>List of 32-bit AddIns for current user</returns>
    '    Friend Function getListof32BitHKCUSparxEntries() As List(Of SparxEntry)
    '        Dim my32HKCUAddInEntries As New List(Of SparxEntry)
    '        Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys32) ' get 32-bit AddIn for the current user
    '        If myCUKey IsNot Nothing Then
    '            For Each pEntryKey In myCUKey.GetSubKeyNames
    '                Dim AddInInfo As New SparxEntry
    '                AddInInfo.Name = pEntryKey
    '                AddInInfo.ClassName = Registry.GetValue(eaHKCU32AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
    '                AddInInfo.HIVE = AddInEntry.cHKCU32
    '                my32HKCUAddInEntries.Add(AddInInfo)
    '            Next
    '        End If
    '        Return my32HKCUAddInEntries
    '    End Function

    '    ''' <summary>
    '    ''' 32 bit AddIns when installed local machine (for all users)
    '    ''' Sparx keys (32-bit apps) can be in HKLM\Software\SparxSystems\EAAddIns
    '    ''' BUT NOTE for 64-bit operating systems the keys will be located in HKLM\Software\Wow6432Node\SparxSystems\EAAddIns
    '    ''' </summary>
    '    ''' <returns>List of 32-bit Addins for local machine</returns>
    '    Friend Function getListof32BitHKLMSparxEntries() As List(Of SparxEntry)

    '        Dim my32HKLMAddInEntries As New List(Of SparxEntry)
    '        Dim myLMKey As RegistryKey = IIf(Environment.Is64BitOperatingSystem, Registry.LocalMachine.OpenSubKey(SparxKeysWOW32), Registry.LocalMachine.OpenSubKey(SparxKeys32))
    '        If myLMKey IsNot Nothing Then
    '            For Each pEntryKey In myLMKey.GetSubKeyNames
    '                Dim AddInInfo As New SparxEntry
    '                AddInInfo.Name = pEntryKey
    '                AddInInfo.ClassName = Registry.GetValue(eaHKLM32AddInKey64 & cBackSlash & pEntryKey, "", cNotSet)
    '                AddInInfo.HIVE = AddInEntry.cHKLM32 ' If(Environment.Is64BitOperatingSystem, AddInEntry.cHKLM32Wow, AddInEntry.cHKLM32)
    '                my32HKLMAddInEntries.Add(AddInInfo)
    '            Next
    '        End If
    '        Return my32HKLMAddInEntries
    '    End Function
    '    ''' <summary>
    '    ''' 64 bit AddIns when install for current users only
    '    ''' HKCU\Software\SparxSystems\EAAddins64
    '    ''' </summary>
    '    ''' <returns>List of current user 64-bit addin entries</returns>
    '    Friend Function getListof64BitHKCUSparxEntries() As List(Of SparxEntry)
    '        Dim my64HKUCAddInEntries As New List(Of SparxEntry)
    '        Dim myCUKey64 As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys64)
    '        If myCUKey64 IsNot Nothing Then
    '            For Each pEntryKey In myCUKey64.GetSubKeyNames
    '                Dim AddInInfo As New SparxEntry
    '                AddInInfo.Name = pEntryKey
    '                AddInInfo.ClassName = Registry.GetValue(eaHKCU64AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
    '                AddInInfo.HIVE = AddInEntry.cHKCU64
    '                my64HKUCAddInEntries.Add(AddInInfo)
    '            Next
    '        End If
    '        Return my64HKUCAddInEntries
    '    End Function
    '    ''' <summary>
    '    ''' 64 bit AddIns when installed for all users
    '    ''' HKLM\Software\Wow6432Node\SparxSystems\EAAddIns64
    '    ''' </summary>
    '    ''' <returns>List of 64-bit addins for Local Machine (all users)</returns>
    '    Friend Function getListof64BitHKLMSparxEntries() As List(Of SparxEntry)
    '        Dim my64KLMAddInEntries As New List(Of SparxEntry)
    '        Dim myLMKey64 As RegistryKey = Registry.LocalMachine.OpenSubKey(SparxKeys64) '"SOFTWARE\Sparx Systems\EAAddins64") 'SparxKeys)
    '        If myLMKey64 IsNot Nothing Then
    '            For Each pEntryKey In myLMKey64.GetSubKeyNames
    '                Dim AddInInfo As New SparxEntry
    '                AddInInfo.Name = pEntryKey
    '                AddInInfo.ClassName = Registry.GetValue(eaHKLM64AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name CLSID name
    '                AddInInfo.HIVE = AddInEntry.cHKLM64
    '                my64KLMAddInEntries.Add(AddInInfo)
    '            Next
    '        End If
    '        Return my64KLMAddInEntries
    '    End Function
    '#End Region


    'Friend Shared Function getClassIDs(pSparxEntries As List(Of SparxEntry)) As List(Of ClassIDEntry)

    '    Dim ClassIDs As New List(Of ClassIDEntry)

    '    Try
    '        For Each classentry In pSparxEntries
    '            Dim _classname As String = classentry.ClassName
    '            Dim _hive As String = classentry.HIVE
    '            Dim ClassID As String = ""
    '            Dim ClassIDExists As Boolean = False
    '            Dim Classsource As String = ""
    '            Dim ClassIDEntry As New ClassIDEntry

    '            Select Case _hive
    '                Case AddInEntry.cHKCU32
    '                    ' check if we have a HKCU32 entry - expected
    '                    ClassID = Registry.GetValue(cHKCR_ClassesRoot & _classname & cBackSlash & cCLSID, "", cNotFound)

    '                    If ClassID IsNot Nothing And ClassID <> "" Then ' we know we have a copy lets checkif it exists in the required HIVE
    '                        ClassIDEntry.ClassName = _classname
    '                        ClassIDEntry.HIVE = AddInEntry.cHKCR
    '                        ClassIDEntry.ClassID = ClassID

    '                        ' now check if in HKCU if
    '                        Dim CLSID = Registry.GetValue(cHKCU_Classes & cBackSlash & _classname & cBackSlash & cCLSID, "", cNotFound)
    '                        If (CLSID IsNot Nothing And CLSID <> "") Then
    '                            ClassIDEntry.HIVE = AddInEntry.cHKCU32
    '                            ClassIDs.Add(ClassIDEntry)
    '                            Continue For
    '                        End If

    '                        CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & _classname & cBackSlash & cCLSID, "", cNotFound)
    '                        If (CLSID IsNot Nothing And CLSID <> "") Then
    '                            ClassIDEntry.HIVE = AddInEntry.cHKLM32
    '                            ClassIDs.Add(ClassIDEntry)
    '                            Continue For
    '                        End If
    '                        ClassIDs.Add(ClassIDEntry)

    '                    End If

    '                Case AddInEntry.cHKLM32
    '                    ' check if we have a HKCU32 entry - expected
    '                    ClassID = Registry.GetValue(cHKCR_ClassesRoot & _classname & cBackSlash & cCLSID, "", cNotFound)

    '                    If ClassID IsNot Nothing And ClassID <> "" Then ' we know we have a copy lets checkif it exists in the required HIVE
    '                        ClassIDEntry.ClassName = _classname
    '                        ClassIDEntry.HIVE = AddInEntry.cHKCR
    '                        ClassIDEntry.ClassID = ClassID

    '                        ' now check if in HKCU if
    '                        Dim CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & _classname & cBackSlash & cCLSID, "", cNotFound)
    '                        If (CLSID IsNot Nothing And CLSID <> "") Then
    '                            ClassIDEntry.HIVE = AddInEntry.cHKLM32
    '                            ClassIDs.Add(ClassIDEntry)
    '                            Continue For
    '                        End If

    '                        CLSID = Registry.GetValue(cHKCU_Classes & cBackSlash & _classname & cBackSlash & cCLSID, "", cNotFound)
    '                        If (CLSID IsNot Nothing And CLSID <> "") Then
    '                            ClassIDEntry.HIVE = AddInEntry.cHKCU32
    '                            ClassIDs.Add(ClassIDEntry)
    '                            Continue For
    '                        End If
    '                        ClassIDs.Add(ClassIDEntry)

    '                    End If
    '            End Select




    '        Next

    '        Return ClassIDs

    '        'ClassSource = ""
    '        'ClassID = ""
    '        '' for 64-bit OS we would expect the class information to be in the relevant hive - we could get from HKCR
    '        '' so use the Spaxr AddIn location PLUS the OS to identify where we would expect to find or not - get the information from 
    '        '' get all information from HKCR
    '        '' 64-bit OS
    '        'ClassSource = cNotFound
    '        'Debug.Print(AddInName)
    '        'Select Case AddInSource
    '        '   '32-bit add in
    '        '    Case AddInEntry.cHKCU32
    '        '        ' check if we have a HKCU32 entry - expected

    '        '        ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '        If ClassID <> "" Then
    '        '            ClassIDExists = True
    '        '            ' need to check if there value in HKCU (as the ClassID may be present due to HKLM)
    '        '            Dim CLID = Registry.GetValue(cHKCU_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '            ClassSource = IIf(CLID Is Nothing Or CLID = "", AddInEntry.cHKCR, AddInEntry.cHKCU32)
    '        '            CheckHKCU32(True)
    '        '            ' also check whether there is a HKLM entry
    '        '            Dim c As String = CheckHKLM32(False) : If c <> "" Then ClassSource += "," & c
    '        '        End If

    '        '        'Case AddInEntry.cHKLM32
    '        '    ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '    If ClassID <> "" Then
    '        '        ClassIDExists = True
    '        '        Dim CLID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '        ClassSource = IIf(CLID Is Nothing Or CLID = "", AddInEntry.cHKCR, AddInEntry.cHKLM32)
    '        '        ' also check whether there is a HKCR entry
    '        '        CheckHKLM32(True)
    '        '        Dim c As String = CheckHKCU32(False) : If c <> "" Then ClassSource += "," & c
    '        '    End If
    '        '    '64-bit addins
    '        'Case AddInEntry.cHKCU64
    '        '    ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '    If ClassID <> "" Then
    '        '        ClassIDExists = True
    '        '        Dim CLID = Registry.GetValue(cHKCU_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '        ClassSource = IIf(CLID Is Nothing Or CLID = "", AddInEntry.cHKCR, AddInEntry.cHKCU64)
    '        '        CheckHKCU(True)
    '        '        ' also check whether there is a HKLM entry
    '        '        Dim c As String = CheckHKLM(False) : If c <> "" Then ClassSource += "," & c

    '        '    End If
    '        'Case AddInEntry.cHKLM64
    '        '    ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '    If ClassID <> "" Then
    '        '        ClassIDExists = True
    '        '        Dim CLID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
    '        '        ClassSource = IIf(CLID Is Nothing Or CLID = "", AddInEntry.cHKCR, AddInEntry.cHKLM64)
    '        '        CheckHKLM(True)
    '        '        ' also check whether there is a HKCR entry
    '        '        Dim c As String = CheckHKCU(False) : If c <> "" Then ClassSource += "," & c
    '        '    End If
    '        'Case Else
    '        '        MsgBox("Unable to get class ID for " & AddInName & " Source " & AddInSource)
    '        'End Select
    '        'Return
    '        'End Sub
    '    Catch ex As Exception

    '    End Try
    'End Function


    '' ----- CLASSID check for DLL file functions

    '' checkHKCU32
    '' checkHKLM32
    '' checkHKCU
    '' checkHKLM

    '' each function gets the filename from the expected location derived from the classID entry and compares with the 
    '' get location if present
    '' * check for filename of dll
    '' * if dll filename exists then get dll details
    '' * then also check if there is a other location DLL HLCU32 <> HKLM32 or HKCU <> HKLM
    '' * 
    '' routines return location found or "" if not found
    '''' <summary>
    '''' Check if there is an HKCU entry the current classID
    '''' </summary>
    '''' <param name="pPrimaryCheck">used to flag that this is a secondary check - so we may already have an entry</param>
    '''' <returns>HKCR if yes else ""</returns>
    'Private Function CheckHKCU32(pPrimaryCheck As Boolean) As String
    '    Dim _location As String = IIf(OS64Bit, cHKCUWOW_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
    '                           cHKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
    '    If _location <> "" Then
    '        Dim prefix As String = ""
    '        ' get name of DLL based on location
    '        Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
    '        If _filename = "" Then
    '            _filename = Registry.GetValue(_location, "", "")
    '            If _filename <> "" Then prefix = "Default key: "
    '        End If
    '        '1st pass we check we have a filename 
    '        If pPrimaryCheck Then
    '            If (_filename <> "" And _filename IsNot Nothing) Then
    '                Filename = _filename ' this is the primary filename
    '                DLLSource = AddInEntry.cHKCU32
    '                getDLLAssembly()
    '                RegistryLocation = _location
    '                RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
    '                ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
    '                AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
    '            Else
    '                DLLSource = "No HKCU32 filename"
    '            End If
    '        Else ' secondary check
    '            ' we do not have a filename have another entry
    '            If ((_filename <> "" And _filename IsNot Nothing) And _filename = Filename) Then ' wee have a match so this is a possible entry
    '                Debug.Print("CheckHKCU32() same filename")
    '                DLLSource += " : " & AddInEntry.cHKCU32
    '            End If
    '        End If
    '    End If
    '    Return ""
    'End Function
    '''' <summary>
    '''' Check if there are any HKLM entries for the current classID
    '''' </summary>
    '''' <returns>HKLM if yes else ""</returns>
    'Private Function CheckHKLM32(pPrimaryCheck As Boolean) As String
    '    ' there are 2 locations for WOW6432 in HKLM - it is believed they mirror each other but not sure so check both
    '    Dim _location As String = IIf(OS64Bit, cHKLMWow1_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
    '                   cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
    '    If _location = "" Then ' try other location
    '        _location = IIf(OS64Bit, cHKLMWow2_Classes & cBackSlash & ClassID & cBackSlash & cInprocServer32,
    '                   cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
    '    End If
    '    If _location <> "" Then
    '        Dim prefix As String = ""

    '        ' get name of DLL based on location
    '        Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
    '        If _filename = "" Then
    '            _filename = Registry.GetValue(_location, "", "")
    '            If _filename <> "" Then prefix = "Default key: "
    '        End If
    '        '1st pass we check we have a filename 
    '        If pPrimaryCheck Then
    '            If (_filename <> "" And _filename IsNot Nothing) Then
    '                Filename = _filename ' this is the primary filename
    '                DLLSource = prefix & AddInEntry.cHKLM32
    '                getDLLAssembly()
    '                RegistryLocation = _location
    '                RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
    '                ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
    '                AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
    '                Return DLLSource
    '            Else
    '                DLLSource = "No HKLM filename"
    '                Return DLLSource
    '            End If
    '        Else ' secondary check
    '            ' we do not have a filename have another entry
    '            If ((_filename <> "" And _filename IsNot Nothing) And _filename = Filename) Then ' wee have a match so this is a possible entry
    '                Debug.Print("CheckHKLM32() same filename")
    '                DLLSource += " : " & AddInEntry.cHKLM32
    '                Return DLLSource
    '            End If
    '        End If

    '    End If
    '    Return ""
    'End Function
End Class

' used with original
Friend Class AddInEntry

    'Registry related strings defines the source of the Sparx AddIn key and OS which indicates where the class information should be located

    Friend Const cHKCR32 As String = "HKCR32"
    Friend Const cHKCR64 As String = "HKCR64"
    Friend Const cHKCR As String = "HKCR" ' set when not clear which 

    ' 32-bit AddIns on 32-bit EA
    Friend Const cHKCU32 As String = "HKCU32" ' 32-bit Addins on 32-bit OS 
    Friend Const cHKLM32 As String = "HKLM32" ' 32-bit Addins on 32-bit OS

    ' 32-bit AddIns on 64-bit OS
    Friend Const cHKLM32Wow As String = "HKLM32Wow" ' 32-bit AddIn on 64-bit OS ?? NT SURE WE NEED THIS as the Sparx key is detached from the class registration

    ' 64-bit AddIns
    Friend Const cHKCU64 As String = "HKCU64" ' 64-bit AddIn on 64-bit OS
    Friend Const cHKLM64 As String = "HKLM64" ' 64-bit AddIn on 64-bit OS

    '' AddIn Name | Class | Source | CLSID | Source | DLL
    ''' <summary>
    ''' Gets or sets the name of the add in.
    ''' </summary>
    ''' <value>
    ''' The AddIn Name
    ''' </value>
    Property AddInName As String = ""
    ''' <summary>
    ''' Gets or sets the class definition.
    ''' </summary>
    ''' <value>
    ''' The class name i.e. Assembly.Class
    ''' </value>
    Property ClassName As String = "" ' Assembly.Class
    Property SparxAddinLocation As String = "" ' Values are below


    ''' <summary>
    ''' Gets or sets the sparx entry.
    ''' </summary>
    ''' <value>
    ''' The location of the SparxEntry in registry - e.g. HKCU or HKLM
    ''' </value>
    Property SparxEntry As String
    ''' <summary>
    ''' Gets or sets the class source.
    ''' </summary>
    ''' <value>
    ''' Location of class in registry  - e.g. HKCU or HKLM
    ''' </value>
    Property ClassSource
    ''' <summary>
    ''' Gets or sets the CLSID.
    ''' </summary>
    ''' <value>
    ''' Class ID (GUID)
    ''' </value>
    Property CLSID
    ''' <summary>
    ''' Gets or sets the CLSID source.
    ''' </summary>
    ''' <value>
    ''' Class ID Source in registry  - e.g. HKCU or HKLM
    ''' </value>
    Property CLSIDSource
    ''' <summary>
    ''' Gets or sets the DLL.
    ''' </summary>
    ''' <value>
    ''' The DLL full file pathname
    ''' </value>
    Property DLL


End Class
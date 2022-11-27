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
Imports Microsoft.Win32

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

    ''' <summary>
    ''' Gets a list of addin entries from Windows Registry
    ''' - start with HKCU as they have precendence over the HKLM
    ''' Assume that the registry keys in the following locations identify 32-bit AddIns
    '''	"HKCU\SOFTWARE\Sparx Systems\EAAddins" - HKCU
    '''	"HKLM\ SOFTWARE\Sparx Systems\EAAddins" - HKLM, although if the addin Is running on a 64-bit operating system this would be "HKLM\ SOFTWARE\WOW6432NODE\Sparx Systems\EAAddins"  HKLMWow
    '''	 Assume that the registry keys in the following locations identify 64-bit AddIns
    '''	 "HKCU\SOFTWARE\Sparx Systems\EAAddins64" - HKCU64
    '''	 "HKLM\ SOFTWARE\Sparx Systems\EAAddins64" – HKLM64
    ''' </summary>
    ''' <returns>List of AddIn entries</returns>
    Function getListOfEAAddinEntries() As ArrayList

        ' for all of the keys which are not nothing we add them to an array
        Dim myAddInEntries As New ArrayList
        For Each entry As AddInEntry In getListof32BitHKCUAddinEntries()
            myAddInEntries.Add(entry)
        Next

        For Each entry As AddInEntry In getListof32BitHKLMAddinEntries()
            myAddInEntries.Add(entry)
        Next

        For Each entry As AddInEntry In getListof64BitHKCUAddinEntries()
            myAddInEntries.Add(entry)
        Next
        For Each entry As AddInEntry In getListof64BitHKLMAddinEntries()
            myAddInEntries.Add(entry)
        Next

        Return myAddInEntries

    End Function


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

End Class

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
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

#Region "Functions to get AddIn Entries"
    ''' <summary>
    ''' 32 bit AddIns when install for current users only
    ''' Sparx keys (32-bit apps) will be in HKCU\Software\SparxSystems\EAAddIns
    ''' </summary>
    ''' <returns>List of 32-bit AddIns for current user</returns>
    Function getListof32BitHKCUAddinEntries() As ArrayList
        Dim my32HKCUAddInEntries As New ArrayList
        Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cSparxKeys32) ' get 32-bit AddIn for the current user
        If myCUKey IsNot Nothing Then
            For Each pEntryKey In myCUKey.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey
                AddInInfo.ClassName = Registry.GetValue(ceaHKCU32AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
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
        Dim myLMKey As RegistryKey = IIf(Environment.Is64BitOperatingSystem, Registry.LocalMachine.OpenSubKey(cSparxKeysWOW32), Registry.LocalMachine.OpenSubKey(cSparxKeys32))
        If myLMKey IsNot Nothing Then
            For Each pEntryKey In myLMKey.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey ' registry location for addin name 
                AddInInfo.ClassName = Registry.GetValue(ceaHKLM32AddInKey64 & cBackSlash & pEntryKey, "", cNotSet)
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
        Dim myCUKey64 As RegistryKey = Registry.CurrentUser.OpenSubKey(cSparxKeys64)
        If myCUKey64 IsNot Nothing Then
            For Each pEntryKey In myCUKey64.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey
                AddInInfo.ClassName = Registry.GetValue(ceaHKCU64AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
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
        Dim myLMKey64 As RegistryKey = Registry.LocalMachine.OpenSubKey(cSparxKeys64) '"SOFTWARE\Sparx Systems\EAAddins64") 'SparxKeys)
        If myLMKey64 IsNot Nothing Then
            For Each pEntryKey In myLMKey64.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey ' registry location for addin name 
                AddInInfo.ClassName = Registry.GetValue(ceaHKLM64AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name CLSID name
                AddInInfo.SparxAddinLocation = AddInEntry.cHKLM64
                my64KLMAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my64KLMAddInEntries
    End Function
#End Region

End Class


﻿' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports Microsoft.Win32

Module SupportFunctions

    ''' <summary>
    ''' Versions the string.
    ''' </summary>
    ''' <returns>EA Installation inspector version</returns>
    Friend Function versionString() As String
        Dim myVersion As String = "NOT SET"
        Try
            myVersion = "EAII V" & My.Application.Info.Version.Major.ToString & "." & _
               My.Application.Info.Version.Minor.ToString & "." & _
               My.Application.Info.Version.Build.ToString & "." & _
                My.Application.Info.Version.Revision.ToString
        Catch ex As Exception

        End Try
        Return myVersion
    End Function
    ''' <summary>
    ''' Dls the lexists.
    ''' </summary>
    ''' <param name="pFilePath">DLL Filename path.</param>
    ''' <returns>True if exists else false</returns>
    Friend Function DLLexists(pFilePath As String) As Boolean
        Try
            ' remove file from front of string
            If Strings.InStr(pFilePath, "file:///") Then
                Dim myNewFN As String = Strings.Right$(pFilePath, Len(pFilePath) - 8)
                If System.IO.File.Exists(myNewFN) Then Return True
            End If

        Catch ex As Exception

        End Try
        Return False
    End Function




    ' set the relative widths of the columns
    Friend Const AddInNameWidth As Integer = 2
    Friend Const Classwidth As Integer = 3
    Friend Const srcWidth As Integer = 1
    Friend Const CLSIDWidth As Integer = 5
    Friend Const DLLWidth As Integer = 10

  
    ''' <summary>
    ''' init the list view headers
    ''' add headers
    ''' set width
    ''' </summary>
    ''' <param name="plv">Listview</param>
    Friend Sub init_lv(plv As ListView)
        Try
            plv.Columns.Clear()
            plv.Items.Clear()
            plv.Visible = True
            plv.View = View.Details
            plv.GridLines = True
            Dim width As Integer = plv.Width
            width = width / 40
            Dim colHeading = plv.Columns.Add("AddIn Name")
            colHeading.Width = width * AddInNameWidth
            colHeading = plv.Columns.Add("Sparx key")
            colHeading.Width = width * Classwidth
            colHeading = plv.Columns.Add("Class (Assembly name)")
            colHeading.Width = width * Classwidth
            colHeading = plv.Columns.Add("Source")
            colHeading.Width = width * srcWidth
            colHeading = plv.Columns.Add("CLSID")
            colHeading.Width = width * CLSIDWidth
            colHeading = plv.Columns.Add("Source")
            colHeading.Width = width * srcWidth
            colHeading = plv.Columns.Add("DLL")
            colHeading.Width = width * DLLWidth

        Catch ex As Exception

        End Try
    End Sub
    ''' <summary>
    ''' Scales the widths of Listview based based on either width provided or existing width
    ''' </summary>
    ''' <param name="plv">Listview</param>
    ''' <param name="w">optional listview width, otherwise uses existing width </param>
    Friend Sub setWidths(plv As ListView, Optional w As Integer = 0)
        Try
            Dim mylv As ListView = plv
            Dim width As Integer = 1
            If w < 1 Then
                width = mylv.Width
            Else
                width = w
            End If
            width = width / (AddInNameWidth + Classwidth + srcWidth + CLSIDWidth + srcWidth + DLLWidth)
            If mylv.Columns.Count < 6 Then Return
            mylv.Columns.Item(0).Width = width * AddInNameWidth
            mylv.Columns.Item(1).Width = width * srcWidth
            mylv.Columns.Item(2).Width = width * Classwidth
            mylv.Columns.Item(3).Width = width * srcWidth
            mylv.Columns.Item(4).Width = width * CLSIDWidth
            mylv.Columns.Item(5).Width = width * srcWidth
            mylv.Columns.Item(6).Width = width * DLLWidth

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Gets a list of addin entries from Windows Registry
    ''' </summary>
    ''' <returns>List of AddIn entries</returns>
    Function getListOfEAAddinEntries() As ArrayList

        ' keys can be in
        '? HKCU\Software\SparxSystems\EAAddIns
        ' HKCU\Software\Wow6432Node\SparxSystems\EAAddIns
        '? HKLM\Software\SparxSystems\EAAddIns
        ' HKLM\Software\Wow6432Node\SparxSystems\EAAddIns

        Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys) ' get the sparx keys listing the addins for the current user
        Dim myLMKey As RegistryKey = Registry.LocalMachine.OpenSubKey(SparxKeys)

        Dim myAddInEntries As New ArrayList ' to store AddInEntries 
        Dim myClassKey As String
        ' for all of the keys which are not nothing we add them to an array
        Dim myEntries As New ArrayList
        If myCUKey IsNot Nothing Then
            For Each pEntryKey In myCUKey.GetSubKeyNames
                Dim newKey As New AddInEntry
                newKey.AddInName = pEntryKey ' registry location for addin name ' this only works for top level not the 6432
                myClassKey = HKCUfullKey & "\" & pEntryKey ' create the key for the Class name
                newKey.ClassDefinition = Registry.GetValue(myClassKey, "", "NOT SET") ' look for the class name CLSID name
                newKey.SparxEntry = "HKCU"
                myEntries.Add(newKey)
            Next
        End If

        If myLMKey IsNot Nothing Then
            For Each pEntryKey In myLMKey.GetSubKeyNames
                Dim newKey As New AddInEntry
                newKey.AddInName = pEntryKey ' registry location for addin name ' this only works for top level not the 6432
                myClassKey = HKLMfullKey & "\" & pEntryKey ' create the key for the Class name
                newKey.ClassDefinition = Registry.GetValue(myClassKey, "", "NOT SET") ' look for the class name CLSID name
                newKey.SparxEntry = "HKLM"
                myEntries.Add(newKey)
            Next
        End If
        Return myEntries

    End Function

    ''' <summary>
    ''' Routine to:
    ''' * Look for the class related to addin entries
    ''' * Populate the listview with addin details
    ''' </summary>
    ''' <param name="plv">The PLV.</param>
    Sub GetAddInClassDetailsAndPopulateListview(plv As ListView)

        For Each AddIn As AddInEntry In getListOfEAAddinEntries()
            ' we have the name, classname and sparx ref
            Try
                Dim RowItem As ListViewItem = plv.Items.Add(AddIn.AddInName) ' add the addin name to list
                RowItem.SubItems.Add(AddIn.SparxEntry) ' sparx ref
                Dim myClassKeyString As String = AddIn.ClassDefinition
                RowItem.SubItems.Add(myClassKeyString) ' class string

                ' find class DLL ID
                ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{CLSID}
                ' EA addins are registered 
                ' More information - http://www.codeproject.com/Articles/1265/COM-IDs-Registry-keys-in-a-nutshell

                ' For each string there should be a keys either under HKCU or HKLM
                ' 1st look in HKCU 

                Dim CLSIDsrc As String = "HKCU" ' we are now looking in 
                Dim cs1 As String = HKCRClasses & "\" & myClassKeyString & "\" & "CLSID" ' this should be class e.g. eaForms.eaForms - but only for CU
                Dim myCSCLSID As String = Registry.GetValue(cs1, "", cNotSet) ' get the CLSID
                If myCSCLSID = "" Then ' if it doesn't exist then need to look in HKLM
                    CLSIDsrc = "HKLM"
                    Dim hs1 As String = HKLMClasses & "\" & myClassKeyString & "\" & "CLSID" ' this should be class e.g. eaForms.eaForms
                    myCSCLSID = Registry.GetValue(hs1, "", cNotSet) ' get CLSID in HKLM
                    If myCSCLSID Is Nothing Then
                        CLSIDsrc = "Not found"
                    ElseIf myCSCLSID = cNotSet Then

                    End If
                End If
                RowItem.SubItems.Add(CLSIDsrc) ' location from which the CLSID is found
                RowItem.SubItems.Add(myCSCLSID) ' add the CLSID

                If myCSCLSID = "" Then ' if we find no class found then the DLL cannot be found
                    RowItem.BackColor = Color.Magenta
                Else
                    ' Now the DLL may be referenced by the CLSID in either HKCU or HKLM 
                    ' expect that it will be the same location as that in which the CLSID has been found
                    Dim DLLsrc As String = "HKCU" ' keep a track of the src
                    Dim csv As String = HKCRClasses & "\CLSID\" & myCSCLSID & "\InprocServer32"

                    If Environment.Is64BitOperatingSystem Then
                        csv = HKCRClasses & "\Wow6432Node\CLSID\" & myCSCLSID & "\InprocServer32"
                    End If

                    Dim myCS1V As String = Registry.GetValue(csv, "CodeBase", cNotSet) 'using the class try to find the DLL path
                    If myCS1V = "" Then ' if it isn't found in HKCU then look in HKLM
                        DLLsrc = "HKLM"
                        csv = HKLMClasses & "\CLSID\" & myCSCLSID & "\InprocServer32"
                        If Environment.Is64BitOperatingSystem Then
                            csv = HKLMClasses & "\Wow6432Node\CLSID\" & myCSCLSID & "\InprocServer32"
                        End If
                        myCS1V = Registry.GetValue(csv, "CodeBase", cNotSet)
                        ' This was added in V3
                        If myCS1V = cNotSet Then
                            DLLsrc = "HKCR"
                            Dim CheckHKCR As String = "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\" & myCSCLSID & "\InProcServer32"
                            myCS1V = Registry.GetValue(CheckHKCR, "", cNotSet)
                        End If
                    End If

                    RowItem.SubItems.Add(DLLsrc) ' add the source
                    RowItem.SubItems.Add(myCS1V) ' add the DLL pathname

                    ' now depending on where items were found flag accordingly
                    If CLSIDsrc <> DLLsrc Then
                        RowItem.BackColor = Color.Red
                    Else
                        RowItem.BackColor = Color.LightGreen
                        ' check if the DLL file exists
                        If Not DLLexists(myCS1V) Then
                            RowItem.BackColor = Color.Cyan
                        End If

                    End If
                    ' and if not set then we know we have an issue
                    If myCS1V = cNotSet Then
                        RowItem.BackColor = Color.Yellow
                    End If
                End If

            Catch ex As Exception
                MsgBox("Init the registry list exception - " & ex.ToString)
            End Try
        Next

    End Sub
End Module

''' <summary>
''' AddIn entry summary
''' </summary>
Friend Class AddInEntry

    '' AddIn Name | Class | Source | CLSID | Source | DLL
    Property AddInName As String
    Property ClassDefinition As String ' Assembly.Class
    Property SparxEntry As String


End Class

''' <summary>
''' AddIn entry details
''' </summary>
Public Class AddInDetail

    '' AddIn Name | Class | Source | CLSID | Source | DLL
    ''' <summary>
    ''' Gets or sets the name of the add in.
    ''' </summary>
    ''' <value>
    ''' The AddIn Name
    ''' </value>
    Property AddInName As String
    ''' <summary>
    ''' Gets or sets the class definition.
    ''' </summary>
    ''' <value>
    ''' The class name i.e. Assembly.Class
    ''' </value>
    Property ClassDefinition As String
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
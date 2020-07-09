' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports Microsoft.Win32
Imports System.Reflection

Module SupportFunctions

    ''' <summary>
    ''' Versions the string.
    ''' </summary>
    ''' <returns>EA Installation inspector version</returns>
    Friend Function versionString() As String
        Dim myVersion As String = cNotSet
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
            ElseIf Strings.InStr(pFilePath, ":\") Or Strings.InStr(pFilePath, ":/") Then
                If System.IO.File.Exists(pFilePath) Then Return True
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
    Friend Const DLLVersionWidth As Integer = 2
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
            colHeading = plv.Columns.Add(cCLSID)
            colHeading.Width = width * CLSIDWidth
            colHeading = plv.Columns.Add("Source")
            colHeading.Width = width * srcWidth
            colHeading = plv.Columns.Add("DLL version")
            colHeading.Width = width * DLLVersionWidth
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
            width = width / (AddInNameWidth + Classwidth + srcWidth + CLSIDWidth + srcWidth + DLLVersionWidth + DLLWidth)
            If mylv.Columns.Count < 7 Then Return
            mylv.Columns.Item(0).Width = width * AddInNameWidth
            mylv.Columns.Item(1).Width = width * srcWidth
            mylv.Columns.Item(2).Width = width * Classwidth
            mylv.Columns.Item(3).Width = width * srcWidth
            mylv.Columns.Item(4).Width = width * CLSIDWidth
            mylv.Columns.Item(5).Width = width * srcWidth
            mylv.Columns.Item(6).Width = width * DLLVersionWidth
            mylv.Columns.Item(7).Width = width * DLLWidth

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Gets a list of addin entries from Windows Registry
    ''' - start with HKCU as they have precendence over the HKLM
    ''' </summary>
    ''' <returns>List of AddIn entries</returns>
    Function getListOfEAAddinEntries() As ArrayList

        ' Sparx keys (32-bit apps) can be in
        ' HKCU\Software\SparxSystems\EAAddIns
        ' HKLM\Software\Wow6432Node\SparxSystems\EAAddIns
        'Computer\HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins
        'Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins

        Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cAddins) '"Software\Sparx Systems\EAAddins") 'SparxKeys) ' get the sparx keys listing the addins for the current user
        Dim myLMKey As RegistryKey = Registry.LocalMachine.OpenSubKey(cAddins) '"SOFTWARE\Sparx Systems\EAAddins") 'SparxKeys)
        If Environment.Is64BitOperatingSystem Then myLMKey = Registry.LocalMachine.OpenSubKey(cWowAddins) '"SOFTWARE\WOW6432Node\Sparx Systems\EAAddins") 'SparxKeys)

        Dim myAddInEntries As New ArrayList ' to store AddInEntries 
        Dim myClassKey As String

        ' for all of the keys which are not nothing we add them to an array
        Dim myEntries As New ArrayList
        If myCUKey IsNot Nothing Then
            For Each pEntryKey In myCUKey.GetSubKeyNames
                Dim newKey As New AddInEntry
                newKey.AddInName = pEntryKey ' registry location for addin name ' this only works for top level not the 6432
                myClassKey = HKCUfullKey & cBackSlash & pEntryKey ' create the key for the Class name
                newKey.ClassDefinition = Registry.GetValue(myClassKey, "", cNotSet) ' look for the class name CLSID name
                newKey.SparxAddinLocation = cHKCU
                myEntries.Add(newKey)
            Next
        End If

        If myLMKey IsNot Nothing Then
            For Each pEntryKey In myLMKey.GetSubKeyNames
                Dim newKey As New AddInEntry
                newKey.AddInName = pEntryKey ' registry location for addin name ' this only works for top level not the 6432
                myClassKey = HKLMfullKey & cBackSlash & pEntryKey ' create the key for the Class name
                newKey.ClassDefinition = Registry.GetValue(myClassKey, "", cNotSet) ' look for the class name CLSID name
                newKey.SparxAddinLocation = If(Environment.Is64BitOperatingSystem, cHKLMWow, cHKLM)
                myEntries.Add(newKey)
            Next
        End If
        Return myEntries

    End Function




    ' Colour coding for List View
    ' * Green – OK  - the AddIn DLL has been found and the keys exist in the same hive; we assume that AddIn will be found by EA.
    ' * Cyan – indicates that all the keys look fine but the DLL file does not exist at the specified location
    ' * Magenta – means that no Class ID has been found for the specified AddIn classname, hence the DLL cannot be indentified
    ' * Red – indicates that CLSID and DLL are specified in different registry Key Sections
    ' * Yellow – means that the DLL path is not set so cannot be found


    ''' <summary>
    ''' Routine to:
    ''' * Look for the class related to addin entries
    ''' * Populate the listview with addin details
    ''' </summary>
    ''' <param name="plv">The PLV.</param>
    Sub GetAddInClassDetailsAndPopulateListview(plv As ListView)

        For Each AddIn As AddInEntry In getListOfEAAddinEntries()


            Try
                ' 1. Output the AddIn name, classname and sparx location
                Dim _RowItem As ListViewItem = plv.Items.Add(AddIn.AddInName) ' add the addin name to list
                _RowItem.SubItems.Add(AddIn.SparxAddinLocation) ' sparx ref
                Dim _ClassDefinition As String = AddIn.ClassDefinition ' get the name of the class that we looking for
                _RowItem.SubItems.Add(_ClassDefinition) ' class string

                ' 2. Find the CLSID and source from the Classname
                ' There are 3 location that we can check for the class
                ' \HKEY_CURRENT_USER\Software\Classes\ - this is the 1st place to look as it has precendence for current user
                ' \HKEY_LOCAL_MACHINE\SOFTWARE\Classes
                ' \HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node - this is pretty unlikely unless 

                ' find class DLL ID
                ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{CLSID}
                ' EA addins are registered 
                ' More information - http://www.codeproject.com/Articles/1265/COM-IDs-Registry-keys-in-a-nutshell

                ' For each string there should be a keys either under HKCU or HKLM
                ' 1st look in HKCU this has precendence for current user

                ' HKCU
                Dim _ClassSource As String = cHKCU ' we are now looking in 
                Dim _CLSID As String = Registry.GetValue(HKCUClasses & cBackSlash & _ClassDefinition & cBackSlash & cCLSID, "", cNotFound) ' get the CLSID
                
                'HKLM
                If _CLSID = "" Then ' if it doesn't exist then need to look in HKLM
                    _ClassSource = cHKLM 'AddIn.CLSIDSrc
                    _CLSID = Registry.GetValue(HKLMClasses & cBackSlash & _ClassDefinition & cBackSlash & cCLSID, "", cNotFound) ' get CLSID in HKLM
                End If

                'HKLMWow
                If _CLSID = "" Then ' if it doesn't exist then need to look in HKLM
                    _ClassSource = cHKLMWow 'AddIn.CLSIDSrc
                    _CLSID = Registry.GetValue(HKLMWowClasses & cBackSlash & _ClassDefinition & cBackSlash & cCLSID, "", cNotFound) ' get CLSID in HKLM
                End If

                If _CLSID = "" Then _ClassSource = cNotFound
                _RowItem.SubItems.Add(_ClassSource) ' location from which the CLSID is found
                _RowItem.SubItems.Add(_CLSID) ' add the CLSID

                '3. Look for DLL file

                If _CLSID = "" Then ' if we find no class found then the DLL cannot be found
                    ' class not found
                    _RowItem.BackColor = Color.Magenta
                Else
                    ' CLSIDsrc indicates where the class entry is defined, hence we would expect to find the details within the same HIVE; note this may not always be the case
                    ' Also depending on whether it is a 32-bit or 64-bit machine will determine where the DLL detail is located
                    ' However, we have seen cases where the library is put in the "wrong place"
                    Dim _ClassInfo As DLLFileInfo
                    _ClassInfo = getClassDLLFilename(_ClassSource, _CLSID)

                    ' The DLL src will indicate where the information has been found
                    _RowItem.SubItems.Add(_ClassInfo.DLLSource) ' add the source of the class information

                    ' get details of DLL assembly
                    Dim _filename As String = _ClassInfo.Filename
                    If _filename IsNot Nothing Then
                        If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                        Try
                            Dim ass As Assembly = Assembly.LoadFile(_filename)
                            _RowItem.SubItems.Add(ass.GetName().Version.ToString)
                            _RowItem.SubItems.Add(_filename) ' add the DLL pathname
                        Catch ex As Exception
                            _RowItem.SubItems.Add("Unable to determine")
                            _RowItem.SubItems.Add(_filename) ' add the DLL pathname
                        End Try
                    End If

                    ' now depending on where items were found flag accordingly
                    If _ClassSource <> _ClassInfo.DLLSource Then
                        _RowItem.BackColor = Color.Red
                    ElseIf _ClassInfo.Filename = cNotSet Then
                        _RowItem.BackColor = Color.Yellow
                    Else
                        _RowItem.BackColor = If(DLLexists(_ClassInfo.Filename), Color.LightGreen, Color.Cyan)  ' File does not exist
                    End If
                End If

            Catch ex As Exception
                MsgBox("Init the registry list exception - " & ex.ToString)
            End Try
        Next

    End Sub

    ' Get the DLL filename which we assume is in the relevant place for HKCU / HKLM - the hive is provided by the caller
    ' TODO - if it fails to find then check other locations

    Private Function getClassDLLFilename(pHIVE As String, pClassID As String) As DLLFileInfo
        Dim _result As New DLLFileInfo
        Try

            ' Now the DLL may be referenced by the CLSID in either HKCU or HKLM 
            ' expect that it will be the same location as that in which the CLSID has been found
            'Dim DLLsrc As String = "HKCU" ' keep a track of the src

            '   Dim root As String = If(pHIVE = cHKCU, HKCUClasses, HKLMClasses)
            Dim _location As String = ""
            Select Case pHIVE
                Case cHKCU
                    _location = HKCUClasses & cBackSlash & cCLSID & cBackSlash & pClassID & cBackSlash & cInprocServer32
                    _result.DLLSource = cHKCU

                Case cHKLM, cHKLMWow
                    _result.DLLSource = cHKLM
                    _location = If(Environment.Is64BitOperatingSystem, HKLMClasses & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & pClassID & cBackSlash & cInprocServer32, _
                                   HKLMClasses & cBackSlash & cCLSID & cBackSlash & pClassID & cBackSlash & cInprocServer32)

                Case Else
                    ' ERROR
            End Select

            _result.Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
            _result.Location = _location

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return _result
    End Function

    ' function to get class information
    Friend Function getClassInformation(pHIVE As String, pID As String) As ClassRegistryInformation
        Dim myClassInfo As New ClassRegistryInformation
        Try

            Dim _HIVE As String = If(pHIVE = cHKCU, HKCUClasses, HKLMClasses) ' get the appropriate HIVE
            Dim KeyLocation32bitRoot As String = _HIVE & cBackSlash & cCLSID & cBackSlash & pID
            Dim KeyLocation64bitRoot As String = _HIVE & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & pID
            Dim _keylocation As String = If(Environment.Is64BitOperatingSystem, KeyLocation64bitRoot, KeyLocation32bitRoot)


            myClassInfo.HIVE = _HIVE
            myClassInfo.CodeBase = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cCodeBase, cNotSet)
            myClassInfo.Assembly = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cAssembly, cNotSet)
            myClassInfo.ClassName = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cClass, cNotSet)
            myClassInfo.RunTimeVersion = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cRuntimeVersion, cNotSet)
            myClassInfo.ProgID = Registry.GetValue(_keylocation & cBackSlash & cProgId, "", cNotSet)

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return myClassInfo

    End Function

    ' execute shell command - SHOULD be done started in a background thread 
    Friend Function ExecuteCommand(pCommand As String) As String
        Dim output As String = ""
        Try
            ' probably need to do this in a different thread so that if it carries on too long we can about

            Dim ProcessInfo As ProcessStartInfo
            Dim Process As New Process
            ProcessInfo = New ProcessStartInfo("cmd.exe", "/C " + pCommand)
            ProcessInfo.CreateNoWindow = True
            ProcessInfo.UseShellExecute = False
            ProcessInfo.WindowStyle = ProcessWindowStyle.Hidden
            Dim p1 As Process = Process.Start(ProcessInfo)
            output = "Process ID = " & p1.Id
            ' MsgBox("Process ID = " & p1.Id)
            p1.WaitForExit()
            Dim exitCode = p1.ExitCode
            p1.Close()

        Catch ex As Exception
            Dim a = ex
        End Try

        Return output


    End Function

    Friend Sub init_lvquery(plv As ListView)
        Try
            plv.Columns.Clear()
            plv.Items.Clear()
            plv.Visible = True
            plv.View = View.Details
            plv.GridLines = True
            Dim width As Integer = plv.Width
            width = width / 50
            Dim colHeading = plv.Columns.Add("Row")
            colHeading.Width = width * 2
            colHeading = plv.Columns.Add("Message")
            colHeading.Width = width * 48


        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


End Module

Friend Class DLLFileInfo
    Property Filename As String ' full filename of the DLL
    Property Location As String ' full location in registry
    Property DLLSource As String ' hive / Wow
End Class

Friend Class ClassRegistryInformation
    Friend HIVE As String = ""
    Friend ClassName As String = ""
    Friend Assembly As String = ""
    Friend CodeBase As String = ""
    Friend RunTimeVersion As String = ""
    Friend ProgID As String = ""
End Class
''' <summary>
''' AddIn entry summary
''' </summary>
Friend Class AddInEntry

    '' AddIn Name | Class | Source | CLSID | Source | DLL
    Property AddInName As String = ""
    Property ClassDefinition As String = "" ' Assembly.Class
    Property SparxAddinLocation As String = ""


End Class

''' <summary>
''' AddIn entry details
''' </summary>
Friend Class AddInDetail

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
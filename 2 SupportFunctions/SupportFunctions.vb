' Copyright (C) 2015 - 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
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

    ' List view 
    ' set the relative widths of the columns
    Friend Const AddInNameWidth As Integer = 4
    Friend Const Classwidth As Integer = 3
    Friend Const srcWidth As Integer = 2
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





    ' Colour coding for List View
    ' * Green – OK  - the AddIn DLL has been found and the keys exist in the same hive; we assume that AddIn will be found by EA.
    ' * Cyan – indicates that all the keys look fine but the DLL file does not exist at the specified location
    ' * Magenta – means that no Class ID has been found for the specified AddIn classname, hence the DLL cannot be indentified
    ' * Red – indicates that CLSID and DLL are specified in different registry Key Sections
    ' * Yellow – means that the DLL path is not set so cannot be found

    ''' <summary>
    ''' Routine to:
    ''' * Look for the class related to addin entries for both 32 and 64 bit classes
    ''' * Populate the listview with addin details
    ''' </summary>
    ''' <param name="plv">The PLV.</param>
    ''' <remarks>
    ''' 32-bit addin keys are located in
    ''' * HKCU
    ''' * HLKM WOW6432Node
    ''' 64-bit addin keys are located in
    ''' * HKCU
    ''' * HKLM
    ''' </remarks>
    Sub Get3264AddInClassDetailsAndPopulateListview(plv As ListView)

        ' Get the entries for each area in the registry containing Sparx AddIn entries 32/64-bit

        Dim AI As New AddInInformation

        For Each CurrentAddInEntry In AI.getListof32BitHKCUAddinEntries() ' 32-bit EA entries for current user
            Dim _RowItem As ListViewItem = plv.Items.Add(CurrentAddInEntry.AddInName) ' add the addin name to list
            _RowItem.SubItems.Add(CurrentAddInEntry.SparxAddinLocation) ' sparx ref
            _RowItem.SubItems.Add(CurrentAddInEntry.ClassName) ' class string
            Dim _C As New ClassInformation
            _C.GetClassInformation(CurrentAddInEntry.ClassName, {AddInEntry.cHKCU32, AddInEntry.cHKLM32})
            Try
                ' Add information to table
                _RowItem.SubItems.Add(_C.ClassSource) ' location from which the CLSID is found
                _RowItem.SubItems.Add(_C.ClassID) ' add the CLSID
                _RowItem.SubItems.Add(_C.DLLSource) ' DLLSource indicates where the src add the source of the class information
                _RowItem.SubItems.Add(_C.Version)
                _RowItem.SubItems.Add(_C.Filename)
                _RowItem.BackColor = _C.Colour
            Catch ex As Exception
                MsgBox("Init the registry list exception - " & ex.ToString)
            End Try
        Next


        For Each CurrentAddInEntry In AI.getListof32BitHKLMAddinEntries() ' 32-bit EA entries for local machine
            ' 1. Output the AddInInformation - name, classname and sparx location
            Dim _RowItem As ListViewItem = plv.Items.Add(CurrentAddInEntry.AddInName) ' add the addin name to list
            _RowItem.SubItems.Add(CurrentAddInEntry.SparxAddinLocation) ' sparx ref
            _RowItem.SubItems.Add(CurrentAddInEntry.ClassName) ' class string
            Dim _C As New ClassInformation
            If Environment.Is64BitOperatingSystem Then
                _C.GetClassInformation(CurrentAddInEntry.ClassName, {AddInEntry.cHKLM32Wow, AddInEntry.cHKLM32})
            Else
                _C.GetClassInformation(CurrentAddInEntry.ClassName, {AddInEntry.cHKLM32, AddInEntry.cHKCU32})
            End If
            Try
                ' Add information to table
                _RowItem.SubItems.Add(_C.ClassSource) ' location from which the CLSID is found
                _RowItem.SubItems.Add(_C.ClassID) ' add the CLSID
                _RowItem.SubItems.Add(_C.DLLSource) ' DLLSource indicates where the src add the source of the class information
                _RowItem.SubItems.Add(_C.Version)
                _RowItem.SubItems.Add(_C.Filename)
                _RowItem.BackColor = _C.Colour
            Catch ex As Exception
                MsgBox("Init the registry list exception - " & ex.ToString)
            End Try
        Next

        ' 64-bit HKCU
        For Each CurrentAddInEntry In AI.getListof64BitHKCUAddinEntries() ' 64-bit EA entries for current user
            ' 1. Output the AddInInformation - name, classname and sparx location
            Dim _RowItem As ListViewItem = plv.Items.Add(CurrentAddInEntry.AddInName) ' add the addin name to list
            _RowItem.SubItems.Add(CurrentAddInEntry.SparxAddinLocation) ' sparx ref
            _RowItem.SubItems.Add(CurrentAddInEntry.ClassName) ' class string
            Dim _C As New ClassInformation
            _C.GetClassInformation(CurrentAddInEntry.ClassName, {AddInEntry.cHKCU64, AddInEntry.cHKLM64})
            Try
                ' Add information to table
                _RowItem.SubItems.Add(_C.ClassSource) ' location from which the CLSID is found
                _RowItem.SubItems.Add(_C.ClassID) ' add the CLSID
                _RowItem.SubItems.Add(_C.DLLSource) ' DLLSource indicates where the src add the source of the class information
                _RowItem.SubItems.Add(_C.Version)
                _RowItem.SubItems.Add(_C.Filename)
                _RowItem.BackColor = _C.Colour
            Catch ex As Exception
                MsgBox("Init the registry list exception - " & ex.ToString)
            End Try
        Next

        ' 64-bit HKLM
        For Each CurrentAddInEntry In AI.getListof64BitHKLMAddinEntries() ' 64-bit EA entries for Local machine
            ' 1. Output the AddInInformation - name, classname and sparx location
            Dim _RowItem As ListViewItem = plv.Items.Add(CurrentAddInEntry.AddInName) ' add the addin name to list
            _RowItem.SubItems.Add(CurrentAddInEntry.SparxAddinLocation) ' sparx ref
            _RowItem.SubItems.Add(CurrentAddInEntry.ClassName) ' class string
            Dim _C As New ClassInformation
            _C.GetClassInformation(CurrentAddInEntry.ClassName, {AddInEntry.cHKLM64, AddInEntry.cHKCU64})
            Try
                ' Add information to table
                _RowItem.SubItems.Add(_C.ClassSource) ' location from which the CLSID is found
                _RowItem.SubItems.Add(_C.ClassID) ' add the CLSID
                _RowItem.SubItems.Add(_C.DLLSource) ' DLLSource indicates where the src add the source of the class information
                _RowItem.SubItems.Add(_C.Version)
                _RowItem.SubItems.Add(_C.Filename)
                _RowItem.BackColor = _C.Colour
            Catch ex As Exception
                MsgBox("Init the registry list exception - " & ex.ToString)
            End Try
        Next


    End Sub




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


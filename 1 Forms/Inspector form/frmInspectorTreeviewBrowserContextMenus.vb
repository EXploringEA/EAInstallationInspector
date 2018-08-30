' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports System.IO

Partial Friend Class frmInspector


    Private Sub ProjectBrowser_MouseClick(sender As System.Object, ByVal e As TreeNodeMouseClickEventArgs) Handles Browser.NodeMouseClick

        Try

            ' HKLM : CLSID get GUID
            ' HKLM : Classname - get class name A.B
            ' HKLM : Filename - get filename
            ' HKCU: CLSID get GUID
            ' HKCU : Classname - get class name A.B
            ' HKCU : Filename - get filename

            If (e.Button = MouseButtons.Right) Then

                Dim _MyTreeview As TreeView = sender
                Dim _SelectedNode As TreeNode = _MyTreeview.SelectedNode
                Dim _info As NodeInfo = TryCast(_SelectedNode.Tag, NodeInfo)
                Dim _NodeName As String = _SelectedNode.Tag.name 'TryCast(mySS1.SubItems.Item(1).Text, String) ' Text from message

                ' create a new context menu
                Dim mnuDGContextMenu As New ContextMenu()

                ' filename
                If _NodeName.Contains(cHKCUFilename) Or _NodeName.Contains(cHKLMFilename) Or _NodeName.Contains(cHKLMWowFilename) Then
                    _filename = _info.Filename
                    _FileDirectory = Path.GetDirectoryName(_info.FilePathName)

                    Dim mnuItemOpenDLLLocation As New MenuItem("Open file location in explorer: " & _FileDirectory)
                    mnuDGContextMenu.MenuItems.Add(mnuItemOpenDLLLocation)
                    AddHandler mnuItemOpenDLLLocation.Click, AddressOf OpenFullFilenameWithExplorer
                    Dim mnuFilenameInfo As New MenuItem("File info : " & _filename)
                    mnuDGContextMenu.MenuItems.Add(mnuFilenameInfo)
                    AddHandler mnuFilenameInfo.Click, AddressOf DisplayFileInformation
                    Dim mnuFilenameCheck As New MenuItem("Filename query : " & _filename)
                    AddHandler mnuFilenameCheck.Click, AddressOf RunQueryForFilename
                    mnuDGContextMenu.MenuItems.Add(mnuFilenameCheck)
                End If

                'classname
                If _NodeName.Contains(cHKCUClassname) Or _NodeName.Contains(cHKLMClassname) Or _NodeName.Contains(cHKLMWowClassname) Then
                    _Classname = _info.ClassName
                    'Dim mnuClassnameInfo As New MenuItem("Class info : " & _classname)
                    'mnuDGContextMenu.MenuItems.Add(mnuClassnameInfo)
                    'AddHandler mnuClassnameInfo.Click, AddressOf DisplayClassnameInformation
                    Dim mnuClassnameCheck As New MenuItem("Classname query : " & _Classname)
                    AddHandler mnuClassnameCheck.Click, AddressOf RunQueryForClassname
                    mnuDGContextMenu.MenuItems.Add(mnuClassnameCheck)
                End If

                'Class ID
                If _NodeName.Contains(cHKCUCLSID) Or _NodeName.Contains(cHKLMCLSID) Or _NodeName.Contains(cHKLMWowCLSID) Then
                    _classID = _info.CLSID
                    'Dim mnuCLSIDInfo As New MenuItem("Class info : " & _classID)
                    'mnuDGContextMenu.MenuItems.Add(mnuCLSIDInfo)
                    'AddHandler mnuCLSIDInfo.Click, AddressOf DisplayClassIDInformation
                    Dim mnuCLSIDCheck As New MenuItem("Class ID query : " & _classID)
                    AddHandler mnuCLSIDCheck.Click, AddressOf RunQueryForClassID
                    mnuDGContextMenu.MenuItems.Add(mnuCLSIDCheck)
                End If

                'Prog ID
                If _NodeName.Contains(cHKCUProgID) Or _NodeName.Contains(cHKLMProgID) Or _NodeName.Contains(cHKLMWowProgID) Then
                    _ProgID = _info.ProgID
                    Dim mnuProgIDCheck As New MenuItem("Prog ID query : " & _ProgID)
                    AddHandler mnuProgIDCheck.Click, AddressOf RunQueryForProgID
                    mnuDGContextMenu.MenuItems.Add(mnuProgIDCheck)
                End If

                ' display context menu
                mnuDGContextMenu.Show(sender, New Point(e.Location.X, e.Location.Y))

            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub



End Class

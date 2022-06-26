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

    Private currentFilename As String = ""
    Private currentFilePath As String = ""
    Private currentCLSIDLocation As String = ""
    Private Sub ProjectBrowser_MouseClick(sender As System.Object, ByVal e As TreeNodeMouseClickEventArgs) Handles Browser.NodeMouseClick

        Try

            If (e.Button = MouseButtons.Right) Then

                Dim _MyTreeview As TreeView = sender
                Dim _SelectedNode As TreeNode = _MyTreeview.SelectedNode
                Dim _Node As NodeInfo = TryCast(_SelectedNode.Tag, NodeInfo)


                currentCLSIDLocation = _Node.CLSIDLocation
                currentFilename = _Node.Filename
                _Classname = _Node.ClassName
                _classID = _Node.CLSID

                ' create a new context menu
                Dim mnuDGContextMenu As New ContextMenu()
                ' filename
                If _Node.Name.Contains(cNodeDLLFilename) Then '
                    currentFilePath = Path.GetDirectoryName(_Node.FilePathName)
                    currentFilename = Path.GetFileName(_Node.FilePathName)

                    Dim mnuItemOpenDLLLocation As New MenuItem("Open file location in explorer: " & currentFilePath)
                    mnuDGContextMenu.MenuItems.Add(mnuItemOpenDLLLocation)
                    AddHandler mnuItemOpenDLLLocation.Click, AddressOf OpenFullFilenameWithExplorer
                    Dim mnuFilenameInfo As New MenuItem("File info : " & currentFilename)
                    mnuDGContextMenu.MenuItems.Add(mnuFilenameInfo)
                    AddHandler mnuFilenameInfo.Click, AddressOf DisplayFileInformation
                    Dim mnuFilenameCheck As New MenuItem("Filename query : " & currentFilename)
                    AddHandler mnuFilenameCheck.Click, AddressOf RunQueryForFilename
                    mnuDGContextMenu.MenuItems.Add(mnuFilenameCheck)
                End If

                'classname
                If _Node.Name.Contains(cNodeClassname) Then '
                    _Classname = _Node.ClassName
                    Dim mnuClassnameCheck As New MenuItem("Classname query : " & _Classname)
                    AddHandler mnuClassnameCheck.Click, AddressOf RunQueryForClassname
                    mnuDGContextMenu.MenuItems.Add(mnuClassnameCheck)
                End If

                'Class ID
                If _Node.Name.Contains(cNodeCLSID) Then '
                    _classID = _Node.CLSID
                    Dim mnuCLSIDCheck As New MenuItem("Class ID query : " & _classID)
                    AddHandler mnuCLSIDCheck.Click, AddressOf RunQueryForClassID
                    mnuDGContextMenu.MenuItems.Add(mnuCLSIDCheck)
                End If

                'Prog ID
                If _Node.Name.Contains(cNodeProgID) Then '
                    _ProgID = _Node.ProgID
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

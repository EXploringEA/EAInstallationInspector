' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Partial Friend Class frmInspector
    ' this handles the click event, specifically we are interested in right click where we present a context menu based on the content of the line
    ' * Codebase / file prefix we assume a file - check references of this file
    ' * Codebase / file prefix we assume a file - open DLL location in windows explorer
    Private Sub lvQuery_MouseClick(sender As Object, e As MouseEventArgs) Handles lvQuery.MouseClick
        Try
            If (e.Button = MouseButtons.Right) Then

                If TypeOf sender Is ListView Then
                    Dim mySS As ListView = sender
                    Dim mySS1 As ListViewItem = mySS.SelectedItems(0)
                    _lvQuerySelectedLine = TryCast(mySS1.SubItems.Item(1).Text, String) ' Text from message

                    ' create a new context menu
                    Dim mnuDGContextMenu As New ContextMenu()
                    ' now decide on options based on what is read
                    If _lvQuerySelectedLine.Contains(cCodeBase) AndAlso _lvQuerySelectedLine.Contains(cFilePrefix) Then


                        _filename = ExtractFilenameFromCodebaseEntry(_lvQuerySelectedLine)

                        ' get the filename and include in the context menu
                        Dim mnuItemInfo As New MenuItem("Get info on file: " & _filename)
                        mnuDGContextMenu.MenuItems.Add(mnuItemInfo)
                        AddHandler mnuItemInfo.Click, AddressOf DisplayFileInformation

                        Dim mnuItemCheckDLL As New MenuItem("Check for other references to same DLL filename: " & _filename)
                        mnuDGContextMenu.MenuItems.Add(mnuItemCheckDLL)
                        AddHandler mnuItemCheckDLL.Click, AddressOf RunQueryForFilename

                        Dim mnuItemOpenDLLLocation As New MenuItem("Open DLL location in explorer: " & _filename)
                        mnuDGContextMenu.MenuItems.Add(mnuItemOpenDLLLocation)
                        AddHandler mnuItemOpenDLLLocation.Click, AddressOf OpenWithExplorer

                    ElseIf _lvQuerySelectedLine.Contains("HKEY") Then
                        Dim mnuItemDelete As New MenuItem("Delete key - NOT IMPLEMENTED YET")
                        mnuDGContextMenu.MenuItems.Add(mnuItemDelete)

                    ElseIf _lvQuerySelectedLine.Contains(cLogFile) AndAlso _lvQuerySelectedLine.Contains(".log") Then
                        Dim mnuItemOpenLogFile As New MenuItem("Open log file")
                        _logfilename = ExtractFilenameFromLogFileEntry(_lvQuerySelectedLine)
                        mnuDGContextMenu.MenuItems.Add(mnuItemOpenLogFile)
                        AddHandler mnuItemOpenLogFile.Click, AddressOf OpenLogFileWithDefaultProgram

                    End If

                    Dim mnuItemClearQueryWindow As New MenuItem("Clear query window")
                    mnuDGContextMenu.MenuItems.Add(mnuItemClearQueryWindow)
                    AddHandler mnuItemClearQueryWindow.Click, AddressOf ClearQueryLogWindow


                    ' display context menu
                    mnuDGContextMenu.Show(sender, New Point(e.Location.X, e.Location.Y))

                End If

            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
End Class

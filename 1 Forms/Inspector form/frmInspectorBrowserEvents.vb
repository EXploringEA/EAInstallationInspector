' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports System.ComponentModel
Imports System.Threading

Partial Friend Class frmInspector


#Region "Treeview events"
    ' after select (click)
    ' double click

    Private doneEvent As New AutoResetEvent(False)


    ''' <summary>
    ''' Display on click 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProjectBrowser_AfterSelect(sender As System.Object, e As System.Windows.Forms.TreeViewEventArgs) Handles Browser.AfterSelect
        Try
            Dim myTreeView As TreeView = sender
            Dim myNode As TreeNode = myTreeView.SelectedNode
            Dim myI As String = myNode.Name
            Dim myTag As NodeInfo = myNode.Tag

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


    ''' <summary>
    ''' Function that responds to selected item in treeview and
    ''' displays the element information in the tab
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProjectBrowser_DoubleClick(sender As Object, e As System.EventArgs) Handles Browser.DoubleClick
        Try
            Dim myTreeView As TreeView = sender
            Dim myNode As TreeNode = myTreeView.SelectedNode
            Dim myI As String = myNode.Name
            Dim myTag As NodeInfo = myNode.Tag

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

#End Region


End Class

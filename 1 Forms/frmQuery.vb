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

Public Class frmQuery

    Friend filename As String = ""
    Friend HKLM As Boolean = False
    Friend HKCU As Boolean = False
    Friend Command As String = ""

    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        '  DialogResult = Windows.Forms.DialogResult.Ignore - don't set the result otherwise form closes!!
        Try
            ' open file browser
            Dim _FileBrowser As New OpenFileDialog
            _FileBrowser.Title = "Select DLL file to check"
            _FileBrowser.Filter = "DLL files (*.dll)|*.dll|All files (*.*)|*.*"
            _FileBrowser.FilterIndex = 1
            If _FileBrowser.ShowDialog() = Windows.Forms.DialogResult.OK Then
                filename = _FileBrowser.FileName
                If File.Exists(filename) Then
                    ' DialogResult = Windows.Forms.DialogResult.OK
                    tbQuery.Text = "Query file:" & filename
                Else
                    MsgBox("No file selected please try again", MsgBoxStyle.OkOnly, "Error selecting file")
                End If
            Else
                MsgBox("File dialog failed", MsgBoxStyle.OkOnly, "Error selecting file")
                '  DialogResult = Windows.Forms.DialogResult.Abort
            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub

    Private Sub btnRunQuery_Click(sender As Object, e As EventArgs) Handles btnRunQuery.Click
        DialogResult = Windows.Forms.DialogResult.OK

        HKCU = rbHKCU.Checked
        HKLM = rbHKLM.Checked
        Command = tbQuery.Text
    End Sub
End Class
' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
Imports System.IO
Imports System.Text
''' <summary>
''' Present help form
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Public Class frmHelp

    ''' <summary>
    ''' Loads the help file in form
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
    Private Sub Information_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            Dim ms As New MemoryStream
            Dim buffer As Byte() = Encoding.UTF8.GetBytes(My.Resources.eaInstallationInspectorInformationV3)

            ms.Write(buffer, 0, buffer.Length)
            ms.Seek(0, SeekOrigin.Begin)
            RichTextBox1.LoadFile(ms, RichTextBoxStreamType.RichText)
        Catch ex As Exception
            ' MsgBox("Ex as " & ex.ToString)
        End Try

    End Sub
End Class
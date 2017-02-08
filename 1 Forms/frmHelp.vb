' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
' You may use, distribute and modify this code under the terms of the 3-Clause BSD License
'
' You should have received a copy of the 3-Clause BSD License with this file. 
' If not, please email: eaForms@EXploringEA.co.uk 
'=====================================================================================
Imports System.IO
Imports System.Text
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
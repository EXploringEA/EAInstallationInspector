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
#Region "Handlers"

    Private Sub DisplayFileInformation()
        Try
            ' pop up window of filename information
            Dim _fileInfo = My.Computer.FileSystem.GetFileInfo(_filename)
            Dim s As String = "Fileinfo : " & _filename & vbCrLf & vbCrLf
            s += "Full name: " & _fileInfo.FullName & vbCrLf & vbCrLf
            s += "Last accessed: " & _fileInfo.LastAccessTime.ToLongDateString & vbCrLf & vbCrLf
            '  s += "Length " & _fileInfo.Length & vbCrLf
            s += "Directory: " & _fileInfo.DirectoryName & vbCrLf & vbCrLf
            '   MsgBox(s, MsgBoxStyle.OkOnly, "File information " & _filename)
            MessageBox.Show(s, "File information " & _filename, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub


    ' queries should run in background
    Private Sub RunQueryForFilename()
        Try
            RunQuery(_filename)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub

    Private Sub RunQueryForClassID()
        Try
            RunQuery(_classID)


        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub

    Private Sub RunQueryForClassname()
        Try
            '   RunQuery(_Classname)
            ' create queries - then schedule
            Dim cmd As String = "reg query HKLM\SOFTWARE\Classes /reg:32 /s /f " & _Classname
            _Query.addQuery(cmd)
            cmd = "reg query HKCU\SOFTWARE\Classes /reg:32 /s /f " & _Classname
            _Query.addQuery(cmd)


        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub

    Private Sub RunQueryForProgID()
        Try
            RunQuery(_ProgID)

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub

    Private Sub ClearQueryLogWindow()
        Try
            lvQuery.Items.Clear()
            Query.ResetQueryStatus()
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
#End Region
End Class

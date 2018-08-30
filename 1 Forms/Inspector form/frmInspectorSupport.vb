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
#Region "Context menu support functions"
    Private Sub RunQuery(_pram As String)
        Try
            ' create queries - then schedule
            Dim cmd As String = "reg query HKLM\SOFTWARE\CLASSES /reg:32 /s /f " & _pram
            _Query.addQuery(cmd)
            cmd = "reg query HKCU\SOFTWARE\CLASSES /reg:32 /s /f " & _pram
            _Query.addQuery(cmd)


        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub




    ''' <summary>
    ''' Open file using windows associated program
    ''' We assume that the filename has been set up AND Exists
    ''' </summary>
    Private Sub OpenLogFileWithDefaultProgram()
        Try
            openGenFileFile(_logfilename)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If

        End Try

    End Sub
    ''' <summary>
    ''' Function to support openning an app using the filename provided using the current windows associated application
    ''' </summary>
    ''' <param name="filename">The filename.</param>
    ''' <returns>True if successful else False</returns>
    ''' <remarks>Checks that the file exists otherwise will fail</remarks>
    Private Function openGenFileFile(ByVal filename As String)
        Try
            If File.Exists(filename) Then
                Dim psi As New ProcessStartInfo()
                With psi
                    .FileName = filename
                    .UseShellExecute = True
                End With
                Process.Start(psi)
            End If
        Catch ex As Exception

#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
            Return False
        End Try
        Return True
    End Function

    Private Sub OpenFullFilenameWithExplorer()
        Try
            If Directory.Exists(_FileDirectory) Then
                Dim _argument As String = "/e, " & _FileDirectory
                Process.Start("Explorer.exe", _argument)
            Else

                MessageBox.Show("Directory doesn't exist", "Failed to launch explorer", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
    ' function to open Explorer with specified name
    Private Sub OpenWithExplorer()
        Dim _filename As String = ""
        Try
            ' get the selected line and extract the filename
            If _lvQuerySelectedLine.Contains(cCodeBase) AndAlso _lvQuerySelectedLine.Contains(cFilePrefix) Then
                ' find start of filename
                Dim start As Integer = Strings.InStr(_lvQuerySelectedLine, cFilePrefix)
                Dim _r As String = Strings.Right(_lvQuerySelectedLine, Len(_lvQuerySelectedLine) - fileprefixlength - start + 1)
                ' remove all after last \
                Dim _end As Integer = Strings.InStrRev(_r, cForwardSlash)
                Dim _path As String = Strings.Left(_r, _end)

                _filename = Strings.RTrim(_r)
                _path = Strings.Replace(_path, cForwardSlash, cBackSlash)
                Dim _argument As String = "/e, " & _path

                If File.Exists(_filename) Then
                    Process.Start("Explorer.exe", _argument)
                Else

                    MessageBox.Show("File doesn't exist", "Failed to launch explorer", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                End If
            Else
                MessageBox.Show("No file in selected line", "No file to process", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub

#End Region


    ' function to extract a filename from Code base query report line
    Private Function ExtractFilenameFromCodebaseEntry(pSelectedLine As String) As String
        '    Dim _filename As String = ""
        Try
            ' find start of filename
            Dim start As Integer = Strings.InStr(pSelectedLine, cFilePrefix)
            Dim _r As String = Strings.Right(pSelectedLine, Len(pSelectedLine) - fileprefixlength - start + 1)
            ' remove all after last \
            Dim _end As Integer = Strings.InStrRev(_r, "/")
            Dim _path As String = Strings.Left(_r, _end)
            _filename = Strings.RTrim(_r)
            ' need to extract the file name
            _filename = Path.GetFileName(_filename)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return _filename
    End Function



    ' function to extract a filename from Code base query report line
    Private Function ExtractFilenameFromLogFileEntry(pSelectedLine As String) As String
        Try
            _logfilename = Strings.Right(pSelectedLine, Len(pSelectedLine) - cLogFile.Length)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return _logfilename
    End Function

End Class

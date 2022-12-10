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
Imports System.Reflection

Partial Friend Class frmInspector
#Region "Handlers"

    Private Sub HandleDisplayFileInformation()
        Try
            ' pop up window of filename information
            Dim s As String = "File does not exist: " & currentFilename
            Dim _fileInfo = My.Computer.FileSystem.GetFileInfo(currentFilePath & cBackSlash & currentFilename)

            If Not File.Exists(_fileInfo.FullName) Then Return
            s = "Fileinfo : " & currentFilename & vbCrLf & vbCrLf
            Try
                Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_fileInfo.FullName)
                s += "Version: " & ass.Version.ToString & vbCrLf
                s += "Processor architecture: " & ass.ProcessorArchitecture.ToString & vbCrLf & vbCrLf
            Catch ex As Exception
                s += "Failed to get assembly information for: " & currentFilename & vbCrLf

            End Try

            s += "Full name: " & _fileInfo.FullName & vbCrLf & vbCrLf
            s += "Created: " & _fileInfo.CreationTime.ToLongDateString & vbCrLf
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


#End Region
#Region "ListView context menu handlers"


    'routine to get the keys associated with a specific AddIn



    Private Sub HandleGetRegistryValues()
        Try
            Dim myListViewSubItems As ListViewItem.ListViewSubItemCollection = lvListOfAddIns.SelectedItems.Item(0).SubItems
            Dim myEntryDetail As New AddInEntry
            myEntryDetail.AddInName = myListViewSubItems(0).Text
            myEntryDetail.SparxEntry = myListViewSubItems(1).Text
            myEntryDetail.ClassName = myListViewSubItems(2).Text
            myEntryDetail.ClassSource = myListViewSubItems(4).Text ' Check the HIVE in which the assembly is defined
            myEntryDetail.CLSID = myListViewSubItems(3).Text
            myEntryDetail.CLSIDSource = myListViewSubItems(5).Text
            GetRegistryValuesHandler(myEntryDetail)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub HandleExportCreateKeys()

        Try

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub HandleExportDeleteKeys()

    End Sub
#End Region

End Class

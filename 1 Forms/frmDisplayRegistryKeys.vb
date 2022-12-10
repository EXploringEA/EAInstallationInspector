Public Class frmDisplayRegistryKeys

    Private _Keys As ArrayList
    Private _CreateKeys As ArrayList
    Private _DeleteKeys As ArrayList
    Private _classname As String
    Private _hive As String

    Friend Sub init(pKeys As ArrayList, pCreateKeys As ArrayList, pDeleteKeys As ArrayList, pClassname As String, pHive As String)

        Try
            _Keys = pKeys
            _CreateKeys = pCreateKeys
            _DeleteKeys = pDeleteKeys
            _classname = pClassname
            _hive = pHive
            Me.Text = "Registry keys for " & _hive & ":" & _classname
            PopulateList(_Keys)

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub PopulateList(pArray As ArrayList)
        Try

            lbRegKeys.Items.Clear()

            For Each s In pArray
                lbRegKeys.Items.Add(s)
            Next

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub btCreate_Click(sender As Object, e As EventArgs) Handles btCreate.Click
        Try
            PopulateList(_CreateKeys)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub btDelete_Click(sender As Object, e As EventArgs) Handles btDelete.Click
        Try
            PopulateList(_DeleteKeys)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


    Private Sub btExportFiles_Click_1(sender As Object, e As EventArgs) Handles btExportFiles.Click
        Try
            Dim _createRegfile As New RegFileOutput("AddRegfile_" & _hive & "_" & _classname & "_")
            For Each s As String In _CreateKeys
                _createRegfile.output(s)
            Next
            _createRegfile.close()
            Dim _deleteRegFile As New RegFileOutput("DeleteRegFile_" & _hive & "_" & _classname & "_")
            For Each s As String In _DeleteKeys
                _deleteRegFile.output(s)
            Next
            _deleteRegFile.close()

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


End Class
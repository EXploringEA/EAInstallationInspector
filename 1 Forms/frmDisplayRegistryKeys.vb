Public Class frmDisplayRegistryKeys

    Private _Keys As List(Of RegKeyItem)
    Private _CreateKeys As List(Of RegKeyItem)
    Private _DeleteKeys As List(Of RegKeyItem)
    Private _classname As String
    Private _hive As String

    Friend Sub init(pKeys As List(Of RegKeyItem), pCreateKeys As List(Of RegKeyItem), pDeleteKeys As List(Of RegKeyItem), pClassname As String, pHive As String)

        Try
            _Keys = pKeys
            _CreateKeys = pCreateKeys
            _DeleteKeys = pDeleteKeys
            _classname = pClassname
            _hive = pHive
            Me.Text = "Registry keys for " & _hive & ":" & _classname
            PopulateList(_Keys)
            _lastAction = cCreateKeys
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub PopulateList(pArray As List(Of RegKeyItem))
        Try

            lbRegKeys.Items.Clear()
            Dim addKey As Boolean = False

            For Each s As RegKeyItem In pArray
                If s.Exposed = RegKeyItem.KeyType.Standard Or (s.Exposed = RegKeyItem.KeyType.ExposedAssembly And cbIncludeAssemblies.Checked) Then
                    lbRegKeys.Items.Add(s.KeyName)
                End If
            Next

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private _lastAction As String = ""
    Private Const cCreateKeys As String = "CREATE"
    Private Const cDeleteKeys As String = "DELETE"
    Private Sub btCreateAddKeys_Click(sender As Object, e As EventArgs) Handles btCreate.Click
        Try
            PopulateList(_CreateKeys)
            _lastAction = cCreateKeys
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub btCreateDeleteKeys_Click(sender As Object, e As EventArgs) Handles btDelete.Click
        Try
            PopulateList(_DeleteKeys)
            _lastAction = cDeleteKeys
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


    Private Sub btExportFiles_Click_1(sender As Object, e As EventArgs) Handles btExportFiles.Click
        Try
            Dim _createRegfile As New RegFileOutput("AddRegfile_" & _hive & "_" & _classname & "_")
            For Each s As RegKeyItem In _CreateKeys
                If s.Exposed = RegKeyItem.KeyType.Standard Or (s.Exposed = RegKeyItem.KeyType.ExposedAssembly And cbIncludeAssemblies.Checked) Then
                    _createRegfile.output(s.KeyName)
                End If
            Next
            _createRegfile.close()
            Dim _deleteRegFile As New RegFileOutput("DeleteRegFile_" & _hive & "_" & _classname & "_")
            For Each s As RegKeyItem In _DeleteKeys
                If s.Exposed = RegKeyItem.KeyType.Standard Or (s.Exposed = RegKeyItem.KeyType.ExposedAssembly And cbIncludeAssemblies.Checked) Then
                    _deleteRegFile.output(s.KeyName)
                End If
            Next
            _deleteRegFile.close()

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub cbIncludeAssemblies_CheckedChanged(sender As Object, e As EventArgs) Handles cbIncludeAssemblies.CheckedChanged
        Try
            ' if here is a checked change then we force a refresh
            Select Case _lastAction
                Case cCreateKeys
                    PopulateList(_CreateKeys)
                Case cDeleteKeys
                    PopulateList(_DeleteKeys)

            End Select
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
End Class
Friend Class RegKeyItem
    Enum KeyType
        Notset = 0
        Standard = 1
        ExposedAssembly = 2
    End Enum

    Public KeyName As String
    Public Exposed As KeyType
End Class
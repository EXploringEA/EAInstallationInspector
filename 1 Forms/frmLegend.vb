Public Class frmLegend


    Private Sub frmLegend_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            pbLegend.Image = My.Resources.Legend
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub
End Class
Public Class ContextMenuClick
    Inherits EventArgs

    Public Property queryparam As String
    Public Sub New(pQueryParam As String)
        MyBase.New()
        Me.queryparam = pQueryParam
    End Sub

End Class


Imports System.IO
Imports System.Text

Public Class RegFileOutput

    Private _Streamwriter As StreamWriter
    Friend Sub New(Optional pFilename As String = "EAII_RegFile")
        Dim myfilename As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & cBackSlash & pFilename & String.Format("{0:yyyyMMdd_HHmmss}.reg", DateTime.Now)
        _Streamwriter = New StreamWriter(myfilename, True, Encoding.UTF8)
        _Streamwriter.AutoFlush = True
    End Sub
    Friend Sub New(Optional ps As String = "", Optional pFilename As String = "EAII_RegFile")
        Me.New(pFilename)
        If ps <> "" Then Me.output(ps)
    End Sub
    Public Sub output(ps As String)
        _Streamwriter.WriteLine(ps)
    End Sub


    Friend Sub close()
        _Streamwriter.Flush()
        _Streamwriter.Close()
        _Streamwriter = Nothing
    End Sub


End Class

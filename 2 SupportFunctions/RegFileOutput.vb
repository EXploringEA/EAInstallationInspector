
Imports System.IO
Imports System.Text

Public Class RegFileOutput

    Private _Streamwriter As StreamWriter
    Private _filename As String

    Friend Sub New(Optional pFilename As String = "EAII_RegFile")
        _filename = My.Computer.FileSystem.SpecialDirectories.MyDocuments & cBackSlash & pFilename & String.Format("{0:yyyyMMdd_HHmmss}.reg", DateTime.Now)
        _Streamwriter = New StreamWriter(_filename, True, Encoding.UTF8)
        _Streamwriter.AutoFlush = True
    End Sub
    Friend Sub New(Optional ps As String = "", Optional pFilename As String = "EAII_RegFile")
        Me.New(pFilename)
        If ps <> "" Then Me.output(ps)
    End Sub
    Public Sub output(ps As String)
        _Streamwriter.WriteLine(ps)
    End Sub

    Friend Function filename() As String
        Return _filename
    End Function
    Friend Sub close()
        _Streamwriter.Flush()
        _Streamwriter.Close()
        _Streamwriter = Nothing
    End Sub


End Class

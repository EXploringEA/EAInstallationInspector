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
Imports System.Text

Public Class logger


#If DEBUG Then

    Public Shared logger As logger ' used to output stuff in debug mode
    Private _lasttime As Date = DateAndTime.Now


    Private _Streamwriter As StreamWriter
    Friend Sub New(Optional pFilename As String = "SparxExportElementLog")
        Dim myfilename As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & cBackSlash & pFilename & String.Format("{0:yyyyMMdd_HHmmss}.txt", DateTime.Now)
        _Streamwriter = New StreamWriter(myfilename, True, Encoding.UTF8)
        _Streamwriter.AutoFlush = True
        _Streamwriter.WriteLine("Starting log file @ " & DateAndTime.Now.ToString("hh:mm:ss.fff"))
        _lasttime = DateAndTime.Now
    End Sub
    Friend Sub New(Optional ps As String = "", Optional pFilename As String = "SparxExportElementLog")
        Me.New(pFilename)
        If ps <> "" Then Me.log(ps)
    End Sub
    Public Sub log(ps As String)
        _Streamwriter.WriteLine(DateAndTime.Now.ToString("hh:mm:ss.fff") & "::" & ps)
        _lasttime = DateAndTime.Now
    End Sub
    Public Sub logLastTimeInterval(ps As String)
        Dim _timenow As Date = DateAndTime.Now
        Dim _timeinterval As TimeSpan = _timenow - _lasttime
        Dim _tmilliseconds = _timeinterval.Milliseconds
        Dim _tSecs = _timeinterval.Seconds
        Dim _tmins = _timeinterval.Minutes
        Dim _thours = _timeinterval.Hours
        Dim _totaltime As Double = _thours * 3600 + _tmins * 60 + _tSecs + _tmilliseconds / 1000
        Dim _tt = _totaltime.ToString
        _Streamwriter.WriteLine(DateAndTime.Now.ToString("hh:mm:ss.fff") & " :: " & ps & ":: Time taken = " & _tt & " seconds")
        _lasttime = DateAndTime.Now
    End Sub

    Friend Sub write(ps As String)
        _Streamwriter.WriteLine(ps)
    End Sub

    Friend Sub close()
        log("END ========================= END")
        _Streamwriter.Flush()
        _Streamwriter.Close()
        _Streamwriter = Nothing
    End Sub

#End If
End Class

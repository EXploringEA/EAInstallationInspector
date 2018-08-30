' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports System.Threading
Imports System.ComponentModel
Imports System.IO

' This class manages registry queries
' As they can take some time they are scheduled to operate in the background and then when complete signal to the main form 
' need to have a single query class to which we add queries which are then processed in the background

Public Class Query

    Private _CmdQueue As New Queue ' the command queue
    Private _logfilename As String = "" ' name of single log file
    Private _lvResults As ListView = Nothing ' results listview 
    Private _tabControl As TabControl = Nothing ' tab control containing query tab
    Private _tbQueryMessage As TextBox = Nothing ' query messages
    Private _tbQueryActive As TextBox = Nothing ' query active message
    '    Private _btnQueryActive As Button = Nothing

    Private Shared _theQueryInstance As Query = Nothing ' used 
    Private Shared _classLocker As New Object 'The class locker flag 

    Private _queryCmd As String = "" ' current query
    Private _queryActive As Boolean = False ' flag to indicate that query is active
    Private Shared _count As Integer = 0 ' query listview row counter - displayed in column 0

#Region "Interface"

    ' interface function to start the query instance
    Friend Shared Function InitQuery(pTabControl As TabControl, pLV As ListView, pQueryMessage As TextBox, pQueryActive As TextBox, pBtnQueryActive As Button) As Query
        Try
            If _theQueryInstance IsNot Nothing Then _theQueryInstance = Nothing
            If _theQueryInstance Is Nothing Then
                SyncLock (_classLocker)
                    If _theQueryInstance Is Nothing Then _theQueryInstance = New Query(pTabControl, pLV, pQueryMessage, pQueryActive, pBtnQueryActive)
                End SyncLock
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return _theQueryInstance
    End Function

    Friend Sub resetCount()
        _count = 0
    End Sub

    ' routine to reset the query items - clears results window, clears cmd queue and resets flags
    Friend Shared Sub ResetQueryStatus()
        Try
            _count = 0
            _theQueryInstance._queryCmd = ""
            _theQueryInstance._queryActive = False
            _theQueryInstance._tbQueryMessage.Text = "Query reset"
            _theQueryInstance._tbQueryActive.Visible = False
            '    _theQueryInstance._btnQueryActive.Visible = False
            _theQueryInstance._CmdQueue.Clear()
            _theQueryInstance._logfilename = Path.GetTempPath() & "EAInspector_" & String.Format("{0:yyyyMMdd_HHmmss}.log", DateTime.Now) ' single log file used for all queries
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub
#End Region

    ' constructor - private as only single instance created
    Private Sub New(pTabControl As TabControl, pLV As ListView, pQueryMessage As TextBox, pQueryActive As TextBox, pButtonQueryActive As Button)
        _tabControl = pTabControl
        _lvResults = pLV
        _tbQueryMessage = pQueryMessage
        _tbQueryActive = pQueryActive
        '   _btnQueryActive = pButtonQueryActive
        'If pButtonQueryActive IsNot Nothing Then
        '    AddHandler pButtonQueryActive.Click, AddressOf StopBackgroundWorker

        'End If
        ' the idea is that we have a single logfile, although we do need a file for each results set
        _logfilename = Path.GetTempPath() & "EAInspector_" & String.Format("{0:yyyyMMdd_HHmmss}.log", DateTime.Now) ' single log file used for all queries

    End Sub



    ' this receives the queries and adds to the end of a list which is processed
    ' if query list is 1 at the ends then this makes a call to query
    Friend Sub addQuery(pQuery As String)
        Try
            SyncLock _CmdQueue.SyncRoot
                _CmdQueue.Enqueue(pQuery)
            End SyncLock
#If DEBUG Then
            Debug.Print("New query request : " & pQuery)
            RunQuery()
#End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


    Private bw As BackgroundWorker = Nothing

    ''' <summary>
    ''' This function executes a registry query placing the outputs into the specific listview
    ''' </summary>
    Private Function RunQuery() As Boolean
        Dim _result As Boolean = False
        Try
            If Not _queryActive Then

                SyncLock _CmdQueue.SyncRoot
                    _queryCmd = _CmdQueue.Dequeue
                End SyncLock

                If _queryCmd <> "" Then

                    _tabControl.SelectedIndex = 2 ' witch to query tab
                    addrow("Active query : " & _queryCmd, cQueryRowColorNew) ' : " & cmd)) ' add the addin name to list
                    _tbQueryMessage.Text = "Query started: " & _queryCmd
                    _tbQueryMessage.BackColor = Color.Yellow

                    bw = New BackgroundWorker()
                    AddHandler bw.DoWork, AddressOf bw_DoWork
                    AddHandler bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted

                    Dim _Args As Object() = New Object() {_queryCmd, "Some information..."}
                    _queryActive = True
                    _tbQueryActive.Visible = True
                    ' _theQueryInstance._btnQueryActive.Visible = True
                    bw.WorkerSupportsCancellation = True
               
                    bw.RunWorkerAsync(_Args)
                    _result = True ' indicates that all started
                End If
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return _result
    End Function

    Dim _ProcID
    ''' <summary>
    ''' Runs a registry query its own thread
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub bw_DoWork(sender As Object, e As DoWorkEventArgs)
        Try

            Dim _Args As Object() = TryCast(e.Argument, Object())

            '    bw.WorkerSupportsCancellation = True
            '    If Not bw.CancellationPending Then
            ' create full query
            Dim _queryCommand As String = _Args(0)
            ' need to create a results fill
            Dim _ResultsFile As String = Path.GetTempPath() & "EAInspector_" & String.Format("{0:yyyyMMdd_HHmmss}.log", DateTime.Now) ' single log file used for all queries
            '' Dim _redirectLogFile As String = If(File.Exists(_logfilename), " >> ", " > ") ' set as output or append to logfile
            Dim _redirectLogFile As String = " > "
            Dim cmd As String = _queryCommand & _redirectLogFile & _ResultsFile

#If DEBUG Then
            Debug.Print("New query actioned : " & cmd)
#End If
            _ProcID = ExecuteCommand(cmd) ' returns process id
            ' Do the work and then return the results which are handled in RunWorkerCompleted
            e.Result = New Object() {_ResultsFile}
            'Else

            'e.Cancel = True
            'End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


    '    Private Sub StopBackgroundWorker()
    '        Try

    '            If bw IsNot Nothing And bw.IsBusy Then
    '                If bw.WorkerSupportsCancellation Then bw.CancelAsync()
    '            End If
    '        Catch ex As Exception
    '#If DEBUG Then
    '            Debug.Print(ex.ToString)
    '#End If
    '        End Try

    '    End Sub
    ''' <summary>
    ''' Handle the completion from the background worker which has found the node children
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub bw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        ' Get the log file of results
        Try
            Dim oResult As Object() = TryCast(e.Result, Object())
            Dim _fn As String = TryCast(oResult(0), String)
            addrow(cLogFile & _fn, cQueryRowColorDefault)

            Dim _something As Boolean = False
            ' we have the results so need to parse and update listview
            ' open file and read each line and add to listview
            ' does file exist


            Using sw As StreamWriter = File.AppendText(_logfilename)
                sw.WriteLine("New query: " & _queryCmd)
                If File.Exists(_fn) Then
                    Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(_fn)
                    Dim _line As String = ""
                    Do
                        _line = reader.ReadLine
                        _something = True
                        addrow(_line, cQueryRowColorDefault)
                        sw.WriteLine(_line)
                    Loop Until _line Is Nothing
                    reader.Close()
                End If

                sw.WriteLine("================= End of query results =================")
            End Using

            If _something Then
                _tbQueryMessage.Text = "Results complete - view in query window"
                _tbQueryMessage.BackColor = Color.SpringGreen
                addrow("Results completed", cQueryRowColorComplete)
            Else
                addrow("NO results", cQueryRowColorComplete)
                addrow("Query compeleted with NO Results", cQueryRowWarning)
                _tbQueryMessage.Text = "No result from query"
                _tbQueryMessage.BackColor = Color.SpringGreen
            End If

            _queryActive = False
            _tbQueryActive.Visible = False
            '        _theQueryInstance._btnQueryActive.Visible = False
            If _CmdQueue.Count > 0 Then RunQuery() ' initiate the query if items in queue


        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ' output row to query results 
    Sub addrow(pText As String, pColor As Color)
        _count += 1
        Dim RowItem As ListViewItem = _lvResults.Items.Add(_count) '
        RowItem.SubItems.Add(pText)
        RowItem.BackColor = pColor
    End Sub



End Class

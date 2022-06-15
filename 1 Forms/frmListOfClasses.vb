' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
''' <summary>
''' Form to display a list of public classes and methods for a selected AddIn entry (DLL)
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Public Class frmListOfClasses


    ''' <summary>
    ''' List of public methods exposed by DLL
    ''' </summary>
    Private ListOfClasses As New ArrayList
    ''' <summary>
    ''' Initializes a new instance of the <see cref="frmListOfClasses"/> class.
    ''' </summary>
    ''' <param name="pListOfClasses">list of methods.</param>
    Public Sub New(ByRef pListOfClasses As ArrayList)
        InitializeComponent()
        ListOfClasses = pListOfClasses
    End Sub



    ''' <summary>
    ''' Handles the Click event of the btClose control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btClose_Click(sender As Object, e As EventArgs) Handles btClose.Click
        Me.Close()
    End Sub


    ''' <summary>
    ''' Handles the Load event of the frmListOfMethods control - populate the datagrid with the list of public methods provided
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub frmListOfClasses_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Initdg()
            Dim rowcounter As Integer = 0
            Dim columnItems(3) As Object
            For Each entry In ListOfClasses
                rowcounter += 1
                columnItems(0) = rowcounter
                columnItems(1) = entry.ToString
                Dim s As String = "No Methods"
                Dim firstEntry As Boolean = True
                For Each m In entry.DeclaredMethods
                    Dim k As String = m.ToString
                    If firstEntry Then
                        s = ""
                        firstEntry = False
                    End If
                    s += k & vbCrLf & vbCrLf
                Next
                columnItems(2) = s
                dgView.Rows.Add(columnItems)
            Next
        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Initialise the datagrid
    ''' </summary>
    Private Sub Initdg()
        Try

            dgView.ColumnHeadersDefaultCellStyle.Font = New Font(dgView.Font, FontStyle.Bold)

            Dim width As Integer = dgView.Width
            width = width / 40
            Dim myNewcolumn As New DataGridViewTextBoxColumn
            myNewcolumn.HeaderText = "Index"
            myNewcolumn.Width = width * 2


            dgView.Columns.Add(myNewcolumn)
            myNewcolumn = New DataGridViewTextBoxColumn
            myNewcolumn.HeaderText = "Public class"
            myNewcolumn.Width = width * 19
            myNewcolumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            dgView.Columns.Add(myNewcolumn)
            myNewcolumn = New DataGridViewTextBoxColumn
            myNewcolumn.HeaderText = "Public methods"
            myNewcolumn.Width = width * 20

            dgView.Columns.Add(myNewcolumn)

            dgView.AutoGenerateColumns = False
            dgView.RowHeadersVisible = False
            dgView.MultiSelect = True
            dgView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgView.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        Catch ex As Exception

        End Try
    End Sub



    ''' <summary>
    ''' Handles the ColumnWidthChanged event of the dgView control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="DataGridViewColumnEventArgs"/> instance containing the event data.</param>
    Private Sub dgView_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dgView.ColumnWidthChanged
        DGresize()
    End Sub
    ''' <summary>
    ''' Handles the Resize event of the dgView control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub dgView_Resize(sender As Object, e As EventArgs) Handles dgView.Resize
        DGresize()

    End Sub

    ''' <summary>
    ''' Resize datagrid
    ''' </summary>
    Private Sub DGresize()
        Try
            Dim numColumns As Integer = dgView.Columns.Count
            If dgView.Columns.Count < 2 Then Return

            Dim twidth As Integer = 0
            Dim hwidth As Integer = dgView.Columns.GetColumnsWidth(DataGridViewElementStates.Displayed)
            Dim dgwidth As Integer = dgView.Width
            For i = 1 To numColumns - 2
                twidth += dgView.Columns(i).Width
            Next
            dgView.Columns(numColumns - 1).Width = dgView.Width - twidth '+ 5
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Handles the Click event of the btSelectAll to select all data and copy to clipboard
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btSelectAll_Click(sender As Object, e As EventArgs) Handles btSelectAll.Click

        dgView.MultiSelect = True
        dgView.SelectAll()
        dgView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Dim dataObj As DataObject = dgView.GetClipboardContent()
        Clipboard.SetDataObject(dataObj, True)
    End Sub
End Class
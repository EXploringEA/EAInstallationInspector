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

''' <summary>
''' Form to present the detail for a single entry
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Friend Class frmEntryDetail

    Private _DLLFilename As String = ""
    'Private Const fileprefix As String = "file:///"
    'Private Const fileprefixlength As Integer = 8
    ''' <summary>
    ''' Initializes a new instance of the <see cref="frmEntryDetail"/> class.
    ''' Populate with the detail for the selected addin row
    ''' </summary>
    ''' <param name="pEntryDetail">The p entry detail.</param>
    Protected Friend Sub New(pEntryDetail As AddInEntry)
        Try
            InitializeComponent()
            tbAddInName.Text = pEntryDetail.AddInName
            tbSparxRef.Text = pEntryDetail.SparxEntry
            tbAssemblyName.Text = pEntryDetail.ClassName
            tbClassSource.Text = pEntryDetail.ClassSource
            tbCLSID.Text = pEntryDetail.CLSID
            tbCLSIDSRC.Text = pEntryDetail.CLSIDSource

            tbDLL.Text = pEntryDetail.DLL
            _DLLFilename = pEntryDetail.DLL
            Dim _filename As String = Path.GetFullPath(_DLLFilename)

            ' Get the DLL file details
            If File.Exists(_filename) Then
                ' either fileName or fileName2 works for me.
                Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(_filename)
                ' now this fvi has all the properties for the FileVersion information.
                tbDLLVersion.Text = fvi.FileVersion ' but other useful properties exist too.
                Dim _DLLDate As DateTime = File.GetLastWriteTime(_filename)
                tbDLLDate.Text = _DLLDate.ToString
            Else
                tbDLLVersion.Text = "FILE DOES NOT EXIST"
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
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
    ''' Handles copy button which capture the details screen and copy to the windows clipboard
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btCopyDetailToClipboard_Click(sender As Object, e As EventArgs) Handles btCopyDetailToClipboard.Click
        Try
            Dim gfx As Graphics = Me.CreateGraphics()
            Dim bmp As Bitmap = New Bitmap(Me.Width, Me.Height)
            Me.DrawToBitmap(bmp, New Rectangle(0, 0, Me.Width, Me.Height))
            My.Computer.Clipboard.SetImage(bmp)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


    Private Sub btDLLDetail_Click(sender As Object, e As EventArgs) Handles btDLLDetail.Click
        Try
            Dim _filename As String = Path.GetFullPath(_DLLFilename)

            'Dim filename As String = _DLLFilename.Replace("file:///", "")
            'filename = filename.Replace("/", "\")
            'filename = _f
            If File.Exists(_filename) Then
                Dim assembly As Assembly = assembly.LoadFrom(_filename)
                Dim types As Type() = assembly.GetTypes()
                '   Dim s As String = " List of types " & vbCrLf
                Dim _ListOfTypes As New ArrayList
                For Each t As Type In types
                    _ListOfTypes.Add(t)
                Next
                '  MsgBox(s, MsgBoxStyle.OkOnly, "List of public methods for DLL")
                Dim frmMethods As New frmListOfClasses(_ListOfTypes)
                frmMethods.ShowDialog()
            Else
                MessageBox.Show("File " & _filename & " does not exist")
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
            MessageBox.Show("Error accessing information " & vbCrLf & "(see detail below for some, but not all, the explanation): - " & vbCrLf & "====================================" _
                            & vbCrLf & ex.ToString)
        End Try
    End Sub
End Class
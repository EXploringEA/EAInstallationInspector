' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
' You may use, distribute and modify this code under the terms of the 3-Clause BSD License
'
' You should have received a copy of the 3-Clause BSD License with this file. 
' If not, please email: adrian@EXploringEA.co.uk 
'=====================================================================================

Imports System.IO
Imports System.Reflection

''' <summary>
''' Form to present the detail for a single entry
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Public Class frmEntryDetail

    Private _DLLFilename As String = ""
    ''' <summary>
    ''' Initializes a new instance of the <see cref="frmEntryDetail"/> class.
    ''' Populate with the detail for the selected addin row
    ''' </summary>
    ''' <param name="pEntryDetail">The p entry detail.</param>
    Protected Friend Sub New(pEntryDetail As AddInDetail)
        Try
            InitializeComponent()
            tbAddInName.Text = pEntryDetail.AddInName
            tbSparxRef.Text = pEntryDetail.SparxEntry
            tbAssemblyName.Text = pEntryDetail.ClassDefinition
            tbClassSource.Text = pEntryDetail.ClassSource
            tbCLSID.Text = pEntryDetail.CLSID
            tbCLSIDSRC.Text = pEntryDetail.CLSIDSource

            tbDLL.Text = pEntryDetail.DLL
            _DLLFilename = pEntryDetail.DLL

        Catch ex As Exception

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

        End Try
    End Sub


    Private Sub btDLLDetail_Click(sender As Object, e As EventArgs) Handles btDLLDetail.Click
        Try
            Dim filename As String = _DLLFilename.Replace("file:///", "")
            filename = filename.Replace("/", "\")
            If File.Exists(filename) Then
                Dim assembly As Assembly = assembly.LoadFrom(_DLLFilename)
                Dim types As Type() = assembly.GetTypes()
                Dim s As String = " List of types " & vbCrLf
                Dim a As New ArrayList

                For Each t As Type In types
                    '   s += t.ToString & vbCrLf
                    '
                    a.Add(t)
                Next
                '  MsgBox(s, MsgBoxStyle.OkOnly, "List of public methods for DLL")
                Dim frmMethods As New frmListOfClasses(a)
                frmMethods.ShowDialog()
            End If


        Catch ex As Exception

        End Try
    End Sub
End Class
' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports Microsoft.Win32

''' <summary>
''' Main form - presents the results of checking for EA AddIn's and related classes
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Public Class frmInspector

    ''' <summary>
    ''' Handles the Load event of the frmInspector2 control - which retrieves and presents the information for EA AddIns
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub frmInspector2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            LinkLabel1.Links.Add(0, 15, "http://Exploringea.com")
            init_lv(lv2) ' set up the list view
            setWidths(lv2) ' set list view column widths
            GetAddInClassDetailsAndPopulateListview(lv2)
            ToolTip1.SetToolTip(btHelp, versionString)
            ' Check OS and set text
            tbOS.Text = If(Environment.Is64BitOperatingSystem, "64-bit detected", "32-bit detected")
            tbLocation.Text = Registry.GetValue(EA, "Install Path", "")
            tbVersion.Text = Registry.GetValue(EA, "Version", "")

        Catch ex As Exception

        End Try
    End Sub



    ''' <summary>
    ''' Handles the SizeChanged for listview
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub lv2_SizeChanged(sender As Object, e As EventArgs) Handles lv2.SizeChanged
        Try
            If TypeOf sender Is ListView Then
                Dim lv As ListView = sender
                ' we want to add up width of cols 0 to 6 and set the DLL as rest
                Dim col05width As Integer = 0
                For i = 0 To 5
                    col05width += lv.Columns(i).Width
                Next
                lv.Columns(6).Width = lv.Width - col05width

            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Handles the refresh button
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btRefresh_Click(sender As Object, e As EventArgs) Handles btRefresh.Click
        Try
            lv2.Items.Clear()
            GetAddInClassDetailsAndPopulateListview(lv2)
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Handles the help button to present the help information
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btHelp_Click(sender As Object, e As EventArgs) Handles btHelp.Click
        Dim myInfo As New frmHelp
        frmHelp.ShowDialog()
    End Sub

    ''' <summary>
    ''' Handles copy button which capture the screen and copy to the windows clipboard
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btCopy_Click(sender As Object, e As EventArgs) Handles btCopy.Click
        Try
            Dim gfx As Graphics = Me.CreateGraphics()
            Dim bmp As Bitmap = New Bitmap(Me.Width, Me.Height)
            Me.DrawToBitmap(bmp, New Rectangle(0, 0, Me.Width, Me.Height))
            My.Computer.Clipboard.SetImage(bmp)
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Handles the LinkClicked event of the LinkLabel1 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="LinkLabelLinkClickedEventArgs"/> instance containing the event data.</param>
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString())
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' Handles the DoubleClick event for listview entry - used to present the addin detail form for the selected row
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub lv2_DoubleClick(sender As Object, e As EventArgs) Handles lv2.DoubleClick
        Try
            Dim myS As Object = sender
            If TypeOf sender Is ListView Then
                Dim mySS As ListView = sender
                ' need to get information from entry
                Dim mySS1 As ListViewItem = mySS.SelectedItems(0)
                Dim myListViewSubItems As ListViewItem.ListViewSubItemCollection = mySS1.SubItems
                ' launch entry detail form
                Dim myEntryDetail As New AddInDetail
                myEntryDetail.AddInName = myListViewSubItems(0).Text
                myEntryDetail.SparxEntry = myListViewSubItems(1).Text
                myEntryDetail.ClassDefinition = myListViewSubItems(2).Text
                myEntryDetail.ClassSource = myListViewSubItems(3).Text
                myEntryDetail.CLSID = myListViewSubItems(4).Text
                myEntryDetail.CLSIDSource = myListViewSubItems(5).Text
                myEntryDetail.DLL = myListViewSubItems(6).Text
                Dim myDetail As New frmEntryDetail(myEntryDetail)
                myDetail.ShowDialog()
            End If

        Catch ex As Exception

        End Try
    End Sub


End Class
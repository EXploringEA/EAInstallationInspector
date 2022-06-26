' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports Microsoft.Win32
Imports System.IO
Imports System.Reflection

''' <summary>
''' Main form - presents the results of checking for EA AddIn's and related classes
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Partial Friend Class frmInspector

#Region "Main form events"


    Private _lvQuerySelectedLine As String = ""
    Private _filename As String = ""
    Private _Classname As String = ""
    Private _classID As String = ""
    Private _ProgID As String = ""
    Private _logfilename As String = ""
    Private _FileDirectory As String = ""

    Private _installPath32 As String = ""
    Private _installPath64 As String = ""

    Const cLogFile As String = "log file "
    Const cLogFileLength As Integer = 9

    Friend _Query As Query = Nothing

    ''' <summary>
    ''' Handles the Load event of the frmInspector control - which retrieves and presents the information for EA AddIns
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub frmInspector2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
#If DEBUG Then
            logger.logger = New logger("EAInspectorLog")
#End If

            Me.Text = "EA Installation Inspector Version " & My.Application.Info.Version.ToString

            LinkLabel1.Links.Add(0, 15, "http://Exploringea.com")
            init_lv(lvListOfAddIns) ' set up the list view
            setWidths(lvListOfAddIns) ' set list view column widths


            ToolTip1.SetToolTip(btHelp, versionString()) ' sets the app version as tool tip to help
            ' Check OS and set text in 
            tbOS.Text = If(Environment.Is64BitOperatingSystem, "64-bit detected", "32-bit detected")

            ' need to check if the key exists
            _installPath32 = Registry.GetValue(EAHKCU32, "Install Path", "") ' assume installed for current user
            If _installPath32 Is Nothing Then _installPath32 = Registry.GetValue(EAHKLM32, "Install Path", "") ' check if install for all users
            _installPath32 += cBackSlash & "EA.exe"

            '            _installPath32 = Path.GetFullPath(_installPath32)
            _installPath32 = Path.GetDirectoryName(_installPath32)

            If _installPath32 Is Nothing Then _installPath32 = "Not installed"
            tbLocation32.Text = _installPath32 ' Registry.GetValue(EA, "Install Path", "")
            Dim _EAVersion32 As String = Registry.GetValue(EAHKCU32, "Version", "")
            If _EAVersion32 = "" Then _EAVersion32 = Registry.GetValue(EAHKLM32, "Version", "")
            tbVersion32.Text = _EAVersion32

            _installPath64 = Registry.GetValue(EAHKCU64, "Install Path", "")
            If _installPath64 Is Nothing Then _installPath64 = Registry.GetValue(EAHKLM64, "Install Path", "")
            If _installPath64 Is Nothing Then _installPath64 = "Not installed"
            tbLocation64.Text = _installPath64 ' Registry.GetValue(EA, "Install Path", "")
            Dim _EAVersion64 As String = Registry.GetValue(EAHKCU64, "Version", "")
            If _EAVersion64 = "" Then _EAVersion64 = Registry.GetValue(EAHKLM64, "Version", "")
            tbVersion64.Text = _EAVersion64


            '
            checkDebugFrameworkConfig()

            Get3264AddInClassDetailsAndPopulateListview(lvListOfAddIns)

            ' initialise the registry tree and create node for SPARX Addin
            Browser.Nodes.Clear()
            Dim ModelNode As NodeInfo = New NodeInfo(NodeType.SparxRoot, 2)
            ModelNode.Name = "SPARX AddIns"
            startTree(ModelNode)

            _Query = Query.InitQuery(tabControl, lvQuery, tbQueryMessage, tbQueryActive, btStopQueryActive)
            ' initialise the query tab results list view
            init_lvquery(lvQuery)


#If DEBUG Then
            If logger.logger IsNot Nothing Then logger.logger.log("Form loaded")
#End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ' FORM closing - used to cloise the logger file
    Private Sub frmInspector_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
#If DEBUG Then
        If logger.logger IsNot Nothing Then logger.logger.close()
#End If
    End Sub

#End Region

#Region "Tab control events"

    Private Sub tabControl_TabIndexChanged(sender As Object, e As EventArgs) Handles tabControl.TabIndexChanged
        If TypeOf sender Is TabControl Then
            ' based on tab page we get different buttons
            Dim ti = tabControl.TabIndex
        End If
    End Sub

    Private Const cListOfAddinsTab As Integer = 0
    Private Const cTreeViewListOfAddinsTab As Integer = 1
    Private Const cQueryTab As Integer = 2

    Private Sub tabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabControl.SelectedIndexChanged
        If TypeOf sender Is TabControl Then
            ' based on tab page we get different buttons
            Dim ti = tabControl.SelectedIndex

            BtnQuery.Visible = False
            btnClearQueryWindow.Visible = False
            btnCheckDLL.Visible = False
            lbQuery.Visible = False
            tbQueryMessage.Visible = False


            Select Case tabControl.SelectedIndex
                Case cListOfAddinsTab
                Case cTreeViewListOfAddinsTab
                Case cQueryTab
                    BtnQuery.Visible = True
                    btnClearQueryWindow.Visible = True
                    ' btnCheckDLL.Visible = True
                    lbQuery.Visible = True
                    tbQueryMessage.Visible = True

            End Select
        End If
    End Sub
#End Region

#Region "Main form buttons"


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
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
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
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try

    End Sub



    Private Sub btnRegisterDLL_Click(sender As Object, e As EventArgs) Handles btnRegisterDLL.Click
        Try
            ' select DLL open file browser
            Dim _FileBrowser As New OpenFileDialog
            _FileBrowser.Title = "Select DLL file to check"
            _FileBrowser.Filter = "DLL files (*.dll)|*.dll"
            _FileBrowser.FilterIndex = 1
            If _FileBrowser.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                MsgBox("No file selected please try again", MsgBoxStyle.OkOnly, "Error selecting file")
                Return
            End If


            Dim _filename As String = _FileBrowser.FileName
            If Not _filename.Contains(".dll") Then
                MsgBox("You can only register DLL")
                Return
            End If

            ' check the file exists
            If File.Exists(_filename) Then
                ' get DLL name without path
                Dim _DLLname As String = Path.GetFileName(_filename)


                Dim ass As Assembly = Assembly.LoadFile(_filename)
                Dim s As String = "Please confirm you want to register the following DLL" & vbCrLf
                s += _filename & vbCrLf
                s += ass.ToString
                If MsgBox(s, MsgBoxStyle.YesNo, "Register DLL for current user") = MsgBoxResult.Yes Then
                    ' will need to determine  if 32-bit or 64-bit and then use the appropriate version of RegAsm
                    ' Also the run time framework - so need to get this from the DLL assembly
                    ' And then run the command with admin rights

                    ' register DLL Current User
                    '       Dim batchcontent As String = "C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" ".\bin\Debug\MyAddIn.dll" /codebase

                End If
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
#Region "Test button"

    ' Check a DLL and then if it is registered
    Private Sub btCheckDLL_Click(sender As Object, e As EventArgs) Handles btnCheckDLL.Click
        Try
            ' open file browser
            Dim _FileBrowser As New OpenFileDialog
            _FileBrowser.Title = "Select DLL file to check"
            _FileBrowser.Filter = "DLL files (*.dll)|*.dll|All files (*.*)|*.*"
            _FileBrowser.FilterIndex = 1
            If _FileBrowser.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                MsgBox("No file selected please try again", MsgBoxStyle.OkOnly, "Error selecting file")
                Return
            End If

            Dim _filename As String = _FileBrowser.FileName
            ' check the file exists
            If File.Exists(_filename) Then
                ' get DLL name without path
                Dim _DLLname As String = Path.GetFileName(_filename)

                Dim tempPath As String = Path.GetTempPath()
                Dim _logFileName As String = tempPath & "EAInspector_" & String.Format("{0:yyyyMMdd_HHmmss}.log", DateTime.Now)

                ' need to check if it exists within the registry
                ' we know that if it is an AddIn DLL it must have a public class
                '_DLLname = "{9ABF49BC-9ACA-3841-B3C1-C2C8064BC36B}" ' LM
                '_DLLname = "{1CE08978-F879-375D-A0B1-0AC63207639D}" 'CU
                '_DLLname = "eaFormsCE.eaForms"

                Dim cmd As String = "reg query HKCU\SOFTWARE\CLASSES /reg:32 /s /f " & _DLLname & " > " & _logFileName '& " && exit" HK..\SOFTWARE\Classes
                Dim s = ExecuteCommand(cmd)
                MsgBox(cLogFile & _logFileName)
                ' want to get a list of the classID's from these files and then get the entries for each 
                ' note that even though indicated in 64bit they are in WOW6432Node - so need to look at each
                ' collect registry entries
                ' get proper registry path
                ' CLSID
                ' Class name  (this we can then check against known EA class names
                ' Filename
                ' Version
                ' RT version (.NET)



            End If



        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

#End Region

#End Region

#Region "List of AddIns listview events"
    ''' <summary>
    ''' Handles the refresh list button
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btRefresh_Click(sender As Object, e As EventArgs) Handles btRefresh.Click
        Try
            checkDebugFrameworkConfig()


            lvListOfAddIns.Items.Clear()
            '            GetAddInClassDetailsAndPopulateListview(lvListOfAddIns)
            Get3264AddInClassDetailsAndPopulateListview(lvListOfAddIns)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ''' <summary>
    ''' Handles the SizeChanged for listview
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub lvListOfAddIns_SizeChanged(sender As Object, e As EventArgs) Handles lvListOfAddIns.SizeChanged
        Try
            If TypeOf sender Is ListView Then
                Dim lv As ListView = sender
                If lv.Columns.Count < 1 Then Return ' prior to initialisation don't need to do anything 
                ' we want to add up width of cols 0 to 6 and set the DLL as rest
                Dim col05width As Integer = 0
                For i = 0 To 6
                    col05width += lv.Columns(i).Width
                Next
                lv.Columns(7).Width = lv.Width - col05width

            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ''' <summary>
    ''' Handles the DoubleClick event for listview entry - used to present the addin detail form for the selected row
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub lvListOfAddIns_DoubleClick(sender As Object, e As EventArgs) Handles lvListOfAddIns.DoubleClick
        Try
            Dim myS As Object = sender
            If TypeOf sender Is ListView Then
                Dim mySS As ListView = sender
                ' need to get information from entry
                Dim mySS1 As ListViewItem = mySS.SelectedItems(0)
                Dim myListViewSubItems As ListViewItem.ListViewSubItemCollection = mySS1.SubItems
                ' launch entry detail form
                Dim myEntryDetail As New AddInEntry
                myEntryDetail.AddInName = myListViewSubItems(0).Text
                myEntryDetail.SparxEntry = myListViewSubItems(1).Text
                myEntryDetail.ClassName = myListViewSubItems(2).Text
                myEntryDetail.ClassSource = myListViewSubItems(4).Text ' Check the HIVE in which tyhe assembly is defined
                myEntryDetail.CLSID = myListViewSubItems(3).Text
                myEntryDetail.CLSIDSource = myListViewSubItems(5).Text ' Class ID source normally HKCR
                myEntryDetail.DLL = myListViewSubItems(7).Text
                Dim myDetail As New frmEntryDetail(myEntryDetail)
                myDetail.ShowDialog()
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

#End Region

#Region "Treeview tab"

    ''' <summary>
    ''' Start a tree from the given node
    ''' </summary>
    ''' <param name="pNode"></param>
    ''' <remarks></remarks>
    Private Sub startTree(pNode As NodeInfo)

        Try
            ' Want to start with the name - 
            Dim NewTreeNode As TreeNode = Browser.Nodes.Add(pNode.Name)
            NewTreeNode.Tag = pNode
            NewTreeNode.Name = pNode.Name
            AddVirtualNode(NewTreeNode)

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
#End Region

#Region "Query tab"

    ''' <summary>
    ''' Handles the SizeChanged for listview
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub lvQuery_SizeChanged(sender As Object, e As EventArgs) Handles lvQuery.SizeChanged
        Try
            If TypeOf sender Is ListView Then
                Dim lv As ListView = sender
                If lv.Columns.Count < 2 Then Return ' prior to initialisation don't need to do anything 
                Dim col05width As Integer = 0
                col05width += lv.Columns(0).Width
                lv.Columns(1).Width = lv.Width - col05width

            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub


#Region "Run query - only visible when query tab visible"

    Private Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles BtnQuery.Click

        Try
            Dim _q As New frmQuery
            If _q.ShowDialog() = Windows.Forms.DialogResult.OK Then

                Dim _hklmCmd As String = "reg query HKLM\SOFTWARE\CLASSES /reg:32 /s /f "
                Dim _hkcuCmd As String = "reg query HKCU\SOFTWARE\CLASSES /reg:32 /s /f "
                Dim cmd As String = ""
                Dim _filename As String = _q.filename
                ' if filename is set then this 
                ' check the file exists
                If File.Exists(_filename) Then
                    ' get DLL name without path
                    Dim _DLLname As String = Path.GetFileName(_filename)
                    ' request to run a query
                    Cursor = Cursors.WaitCursor

                    If _q.HKLM Then
                        cmd = "reg query HKLM\SOFTWARE\CLASSES  /s /f " & _DLLname
                        _Query.addQuery(cmd)
                    End If
                    If _q.HKCU Then
                        cmd = "reg query HKCU\SOFTWARE\CLASSES  /s /f " & _DLLname
                        _Query.addQuery(cmd)
                    End If

                ElseIf _q.Command <> "" Then
                    If _q.HKLM Then
                        cmd = "reg query HKLM\SOFTWARE\CLASSES  /s /f " & _q.Command
                        _Query.addQuery(cmd)
                    End If
                    If _q.HKCU Then
                        cmd = "reg query HKCU\SOFTWARE\CLASSES  /s /f " & _q.Command
                        _Query.addQuery(cmd)
                    End If
                End If
            End If


        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Cursor = Cursors.Default
    End Sub

#End Region

    Private Sub btnClearQueryWindow_Click(sender As Object, e As EventArgs) Handles btnClearQueryWindow.Click
        lvQuery.Items.Clear()
        If _Query IsNot Nothing Then _Query.resetCount()
    End Sub


    ' method to open the query logfile

    Private Sub lvQuery_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lvQuery.MouseDoubleClick
        Try
            If TypeOf sender Is ListView Then
                Dim lv As ListView = sender
                Dim _logfile = lv.SelectedItems(0).SubItems(1).Text
                ' remove "log file " from start of line
                If cLogFile.Contains(cLogFile) Then
                    Dim _filename As String = Strings.Right(_logfile, _logfile.Length - cLogFileLength)
                    If File.Exists(_filename) Then
                        Dim psi As New ProcessStartInfo()
                        With psi
                            .FileName = _filename
                            .UseShellExecute = True
                        End With
                        Process.Start(psi)
                    End If
                End If
            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
#End Region

    ' check on the debug framework
    Private Sub btDebugFramework_Click(sender As Object, e As EventArgs) Handles btDebugFramework.Click
        Try
            '            MsgBox("EA install location " & _installPath)

            Dim _DebugConfigFilename As String = _installPath32 & "\EA.Exe.config"
            Dim _msg As String = "No framework config file found"
            If _DebugConfigFilename <> _installPath32 AndAlso File.Exists(_DebugConfigFilename) Then
                btDebugFramework.BackColor = Color.SpringGreen
                btDebugFramework.Enabled = True
                _msg = "Config file found" & vbCrLf
                Dim _s As String = File.ReadAllText(_DebugConfigFilename)
                _msg += _s
            Else
                btDebugFramework.BackColor = Color.Gray
                btDebugFramework.Enabled = False
            End If
            MsgBox(_msg, MsgBoxStyle.Information)

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Sub checkDebugFrameworkConfig()

        Dim _DebugConfigFilename As String = _installPath32 & "\EA.Exe.config"
        If _DebugConfigFilename <> _installPath32 AndAlso File.Exists(_DebugConfigFilename) Then
            btDebugFramework.BackColor = Color.SpringGreen
            btDebugFramework.Enabled = True
            btDebugFramework.Visible = True
        Else
            btDebugFramework.Visible = False
            btDebugFramework.BackColor = Color.Gray
            btDebugFramework.Enabled = False
        End If
    End Sub


End Class
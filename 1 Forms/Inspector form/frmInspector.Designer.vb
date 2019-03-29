' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInspector
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInspector))
        Me.lvListOfAddIns = New System.Windows.Forms.ListView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbOS = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbVersion = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbLocation = New System.Windows.Forms.TextBox()
        Me.btHelp = New System.Windows.Forms.Button()
        Me.btRefresh = New System.Windows.Forms.Button()
        Me.btCopy = New System.Windows.Forms.Button()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnCheckDLL = New System.Windows.Forms.Button()
        Me.tabControl = New System.Windows.Forms.TabControl()
        Me.tabListOfAddins = New System.Windows.Forms.TabPage()
        Me.tabRegistryTree = New System.Windows.Forms.TabPage()
        Me.Browser = New System.Windows.Forms.TreeView()
        Me.tabQueryRegistry = New System.Windows.Forms.TabPage()
        Me.lvQuery = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.BtnQuery = New System.Windows.Forms.Button()
        Me.btnClearQueryWindow = New System.Windows.Forms.Button()
        Me.btnRegisterDLL = New System.Windows.Forms.Button()
        Me.lbQuery = New System.Windows.Forms.Label()
        Me.tbQueryMessage = New System.Windows.Forms.TextBox()
        Me.tbQueryActive = New System.Windows.Forms.TextBox()
        Me.btStopQueryActive = New System.Windows.Forms.Button()
        Me.btDebugFramework = New System.Windows.Forms.Button()
        Me.tabControl.SuspendLayout()
        Me.tabListOfAddins.SuspendLayout()
        Me.tabRegistryTree.SuspendLayout()
        Me.tabQueryRegistry.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvListOfAddIns
        '
        Me.lvListOfAddIns.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvListOfAddIns.FullRowSelect = True
        Me.lvListOfAddIns.Location = New System.Drawing.Point(2, 2)
        Me.lvListOfAddIns.MultiSelect = False
        Me.lvListOfAddIns.Name = "lvListOfAddIns"
        Me.lvListOfAddIns.Size = New System.Drawing.Size(1109, 293)
        Me.lvListOfAddIns.TabIndex = 0
        Me.lvListOfAddIns.UseCompatibleStateImageBehavior = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "OS"
        '
        'tbOS
        '
        Me.tbOS.Location = New System.Drawing.Point(40, 6)
        Me.tbOS.Name = "tbOS"
        Me.tbOS.ReadOnly = True
        Me.tbOS.Size = New System.Drawing.Size(100, 20)
        Me.tbOS.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(146, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "EA Version"
        '
        'tbVersion
        '
        Me.tbVersion.Location = New System.Drawing.Point(211, 6)
        Me.tbVersion.Name = "tbVersion"
        Me.tbVersion.ReadOnly = True
        Me.tbVersion.Size = New System.Drawing.Size(100, 20)
        Me.tbVersion.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(317, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "EA Install location "
        '
        'tbLocation
        '
        Me.tbLocation.Location = New System.Drawing.Point(417, 6)
        Me.tbLocation.Name = "tbLocation"
        Me.tbLocation.ReadOnly = True
        Me.tbLocation.Size = New System.Drawing.Size(610, 20)
        Me.tbLocation.TabIndex = 6
        '
        'btHelp
        '
        Me.btHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btHelp.BackColor = System.Drawing.Color.Red
        Me.btHelp.Location = New System.Drawing.Point(12, 383)
        Me.btHelp.Name = "btHelp"
        Me.btHelp.Size = New System.Drawing.Size(75, 23)
        Me.btHelp.TabIndex = 7
        Me.btHelp.Text = "Help"
        Me.btHelp.UseVisualStyleBackColor = False
        '
        'btRefresh
        '
        Me.btRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btRefresh.BackColor = System.Drawing.Color.SpringGreen
        Me.btRefresh.Location = New System.Drawing.Point(316, 383)
        Me.btRefresh.Name = "btRefresh"
        Me.btRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btRefresh.TabIndex = 8
        Me.btRefresh.Text = "Refresh Addins list"
        Me.ToolTip1.SetToolTip(Me.btRefresh, "Refresh list of addins")
        Me.btRefresh.UseVisualStyleBackColor = False
        '
        'btCopy
        '
        Me.btCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btCopy.BackColor = System.Drawing.Color.Gold
        Me.btCopy.Location = New System.Drawing.Point(93, 383)
        Me.btCopy.Name = "btCopy"
        Me.btCopy.Size = New System.Drawing.Size(217, 23)
        Me.btCopy.TabIndex = 9
        Me.btCopy.Text = "Copy current list image to clipboard"
        Me.btCopy.UseVisualStyleBackColor = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel1.Location = New System.Drawing.Point(1000, 382)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(149, 20)
        Me.LinkLabel1.TabIndex = 10
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "EXploringEA.com"
        '
        'btnCheckDLL
        '
        Me.btnCheckDLL.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCheckDLL.BackColor = System.Drawing.Color.Yellow
        Me.btnCheckDLL.Location = New System.Drawing.Point(658, 382)
        Me.btnCheckDLL.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCheckDLL.Name = "btnCheckDLL"
        Me.btnCheckDLL.Size = New System.Drawing.Size(101, 23)
        Me.btnCheckDLL.TabIndex = 11
        Me.btnCheckDLL.Text = "Check a DLL"
        Me.ToolTip1.SetToolTip(Me.btnCheckDLL, "check if DLL registered")
        Me.btnCheckDLL.UseVisualStyleBackColor = False
        Me.btnCheckDLL.Visible = False
        '
        'tabControl
        '
        Me.tabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabControl.Controls.Add(Me.tabListOfAddins)
        Me.tabControl.Controls.Add(Me.tabRegistryTree)
        Me.tabControl.Controls.Add(Me.tabQueryRegistry)
        Me.tabControl.Location = New System.Drawing.Point(14, 49)
        Me.tabControl.Margin = New System.Windows.Forms.Padding(2)
        Me.tabControl.Name = "tabControl"
        Me.tabControl.SelectedIndex = 0
        Me.tabControl.Size = New System.Drawing.Size(1121, 323)
        Me.tabControl.TabIndex = 12
        '
        'tabListOfAddins
        '
        Me.tabListOfAddins.Controls.Add(Me.lvListOfAddIns)
        Me.tabListOfAddins.Location = New System.Drawing.Point(4, 22)
        Me.tabListOfAddins.Margin = New System.Windows.Forms.Padding(2)
        Me.tabListOfAddins.Name = "tabListOfAddins"
        Me.tabListOfAddins.Padding = New System.Windows.Forms.Padding(2)
        Me.tabListOfAddins.Size = New System.Drawing.Size(1113, 297)
        Me.tabListOfAddins.TabIndex = 0
        Me.tabListOfAddins.Text = "List of AddIns"
        Me.tabListOfAddins.UseVisualStyleBackColor = True
        '
        'tabRegistryTree
        '
        Me.tabRegistryTree.Controls.Add(Me.Browser)
        Me.tabRegistryTree.Location = New System.Drawing.Point(4, 22)
        Me.tabRegistryTree.Margin = New System.Windows.Forms.Padding(2)
        Me.tabRegistryTree.Name = "tabRegistryTree"
        Me.tabRegistryTree.Padding = New System.Windows.Forms.Padding(2)
        Me.tabRegistryTree.Size = New System.Drawing.Size(1113, 297)
        Me.tabRegistryTree.TabIndex = 1
        Me.tabRegistryTree.Text = "Registry tree view"
        Me.tabRegistryTree.UseVisualStyleBackColor = True
        '
        'Browser
        '
        Me.Browser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Browser.Location = New System.Drawing.Point(2, 2)
        Me.Browser.Margin = New System.Windows.Forms.Padding(2)
        Me.Browser.Name = "Browser"
        Me.Browser.Size = New System.Drawing.Size(1109, 293)
        Me.Browser.TabIndex = 0
        '
        'tabQueryRegistry
        '
        Me.tabQueryRegistry.Controls.Add(Me.lvQuery)
        Me.tabQueryRegistry.Location = New System.Drawing.Point(4, 22)
        Me.tabQueryRegistry.Margin = New System.Windows.Forms.Padding(2)
        Me.tabQueryRegistry.Name = "tabQueryRegistry"
        Me.tabQueryRegistry.Size = New System.Drawing.Size(1113, 297)
        Me.tabQueryRegistry.TabIndex = 2
        Me.tabQueryRegistry.Text = "Query results"
        Me.tabQueryRegistry.UseVisualStyleBackColor = True
        '
        'lvQuery
        '
        Me.lvQuery.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.lvQuery.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvQuery.FullRowSelect = True
        Me.lvQuery.GridLines = True
        Me.lvQuery.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvQuery.Location = New System.Drawing.Point(0, 0)
        Me.lvQuery.Margin = New System.Windows.Forms.Padding(2)
        Me.lvQuery.Name = "lvQuery"
        Me.lvQuery.Size = New System.Drawing.Size(1113, 297)
        Me.lvQuery.TabIndex = 0
        Me.lvQuery.UseCompatibleStateImageBehavior = False
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Width = 2000
        '
        'BtnQuery
        '
        Me.BtnQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnQuery.Location = New System.Drawing.Point(763, 380)
        Me.BtnQuery.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnQuery.Name = "BtnQuery"
        Me.BtnQuery.Size = New System.Drawing.Size(74, 23)
        Me.BtnQuery.TabIndex = 13
        Me.BtnQuery.Text = "Run query"
        Me.BtnQuery.UseVisualStyleBackColor = True
        Me.BtnQuery.Visible = False
        '
        'btnClearQueryWindow
        '
        Me.btnClearQueryWindow.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClearQueryWindow.Location = New System.Drawing.Point(841, 380)
        Me.btnClearQueryWindow.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClearQueryWindow.Name = "btnClearQueryWindow"
        Me.btnClearQueryWindow.Size = New System.Drawing.Size(145, 23)
        Me.btnClearQueryWindow.TabIndex = 14
        Me.btnClearQueryWindow.Text = "Clear results window"
        Me.btnClearQueryWindow.UseVisualStyleBackColor = True
        Me.btnClearQueryWindow.Visible = False
        '
        'btnRegisterDLL
        '
        Me.btnRegisterDLL.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRegisterDLL.Location = New System.Drawing.Point(538, 382)
        Me.btnRegisterDLL.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRegisterDLL.Name = "btnRegisterDLL"
        Me.btnRegisterDLL.Size = New System.Drawing.Size(116, 23)
        Me.btnRegisterDLL.TabIndex = 15
        Me.btnRegisterDLL.Text = "Register DLL"
        Me.btnRegisterDLL.UseVisualStyleBackColor = True
        Me.btnRegisterDLL.Visible = False
        '
        'lbQuery
        '
        Me.lbQuery.AutoSize = True
        Me.lbQuery.Location = New System.Drawing.Point(317, 37)
        Me.lbQuery.Name = "lbQuery"
        Me.lbQuery.Size = New System.Drawing.Size(55, 13)
        Me.lbQuery.TabIndex = 1
        Me.lbQuery.Text = "Query info"
        Me.lbQuery.Visible = False
        '
        'tbQueryMessage
        '
        Me.tbQueryMessage.Location = New System.Drawing.Point(417, 34)
        Me.tbQueryMessage.Name = "tbQueryMessage"
        Me.tbQueryMessage.Size = New System.Drawing.Size(610, 20)
        Me.tbQueryMessage.TabIndex = 16
        Me.tbQueryMessage.Visible = False
        '
        'tbQueryActive
        '
        Me.tbQueryActive.BackColor = System.Drawing.Color.Yellow
        Me.tbQueryActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbQueryActive.ForeColor = System.Drawing.Color.Red
        Me.tbQueryActive.Location = New System.Drawing.Point(1032, 37)
        Me.tbQueryActive.Name = "tbQueryActive"
        Me.tbQueryActive.ReadOnly = True
        Me.tbQueryActive.Size = New System.Drawing.Size(102, 22)
        Me.tbQueryActive.TabIndex = 17
        Me.tbQueryActive.Text = "Query active"
        Me.tbQueryActive.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.tbQueryActive.Visible = False
        '
        'btStopQueryActive
        '
        Me.btStopQueryActive.BackColor = System.Drawing.Color.Yellow
        Me.btStopQueryActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btStopQueryActive.ForeColor = System.Drawing.Color.Red
        Me.btStopQueryActive.Location = New System.Drawing.Point(1033, 6)
        Me.btStopQueryActive.Name = "btStopQueryActive"
        Me.btStopQueryActive.Size = New System.Drawing.Size(101, 29)
        Me.btStopQueryActive.TabIndex = 18
        Me.btStopQueryActive.Text = "Query active"
        Me.btStopQueryActive.UseVisualStyleBackColor = False
        Me.btStopQueryActive.Visible = False
        '
        'btDebugFramework
        '
        Me.btDebugFramework.Location = New System.Drawing.Point(391, 383)
        Me.btDebugFramework.Name = "btDebugFramework"
        Me.btDebugFramework.Size = New System.Drawing.Size(142, 22)
        Me.btDebugFramework.TabIndex = 19
        Me.btDebugFramework.Text = "EA debug config"
        Me.ToolTip1.SetToolTip(Me.btDebugFramework, "EA.exe.config present if green," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Press to display contents of file" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "to view order" & _
        " of selection")
        Me.btDebugFramework.UseVisualStyleBackColor = True
        '
        'frmInspector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(1161, 416)
        Me.Controls.Add(Me.btDebugFramework)
        Me.Controls.Add(Me.btStopQueryActive)
        Me.Controls.Add(Me.tbQueryActive)
        Me.Controls.Add(Me.tbQueryMessage)
        Me.Controls.Add(Me.lbQuery)
        Me.Controls.Add(Me.btnRegisterDLL)
        Me.Controls.Add(Me.btnClearQueryWindow)
        Me.Controls.Add(Me.BtnQuery)
        Me.Controls.Add(Me.tabControl)
        Me.Controls.Add(Me.btnCheckDLL)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.btCopy)
        Me.Controls.Add(Me.btRefresh)
        Me.Controls.Add(Me.btHelp)
        Me.Controls.Add(Me.tbLocation)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbVersion)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbOS)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInspector"
        Me.Text = "EA Installation Inspector Version 6 in development"
        Me.tabControl.ResumeLayout(False)
        Me.tabListOfAddins.ResumeLayout(False)
        Me.tabRegistryTree.ResumeLayout(False)
        Me.tabQueryRegistry.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvListOfAddIns As System.Windows.Forms.ListView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbOS As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbVersion As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbLocation As System.Windows.Forms.TextBox
    Friend WithEvents btHelp As System.Windows.Forms.Button
    Friend WithEvents btRefresh As System.Windows.Forms.Button
    Friend WithEvents btCopy As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnCheckDLL As System.Windows.Forms.Button
    Friend WithEvents tabControl As System.Windows.Forms.TabControl
    Friend WithEvents tabListOfAddins As System.Windows.Forms.TabPage
    Friend WithEvents tabRegistryTree As System.Windows.Forms.TabPage
    Friend WithEvents Browser As System.Windows.Forms.TreeView
    Friend WithEvents BtnQuery As System.Windows.Forms.Button
    Friend WithEvents tabQueryRegistry As System.Windows.Forms.TabPage
    Friend WithEvents lvQuery As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnClearQueryWindow As System.Windows.Forms.Button
    Friend WithEvents btnRegisterDLL As System.Windows.Forms.Button
    Friend WithEvents lbQuery As System.Windows.Forms.Label
    Friend WithEvents tbQueryMessage As System.Windows.Forms.TextBox
    Friend WithEvents tbQueryActive As System.Windows.Forms.TextBox
    Friend WithEvents btStopQueryActive As System.Windows.Forms.Button
    Friend WithEvents btDebugFramework As System.Windows.Forms.Button
End Class

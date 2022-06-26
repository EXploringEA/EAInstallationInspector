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
Partial Class frmEntryDetail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEntryDetail))
        Me.lbAddInName = New System.Windows.Forms.Label()
        Me.lbAssemblyname = New System.Windows.Forms.Label()
        Me.tbAddInName = New System.Windows.Forms.TextBox()
        Me.lbDLL = New System.Windows.Forms.Label()
        Me.tbAssemblyName = New System.Windows.Forms.TextBox()
        Me.tbDLL = New System.Windows.Forms.TextBox()
        Me.btClose = New System.Windows.Forms.Button()
        Me.lbSparxRef = New System.Windows.Forms.Label()
        Me.tbSparxRef = New System.Windows.Forms.TextBox()
        Me.lbClasssSrc = New System.Windows.Forms.Label()
        Me.tbClassSource = New System.Windows.Forms.TextBox()
        Me.lbCLSID = New System.Windows.Forms.Label()
        Me.tbCLSID = New System.Windows.Forms.TextBox()
        Me.lbCLSIDSRC = New System.Windows.Forms.Label()
        Me.tbCLSIDSRC = New System.Windows.Forms.TextBox()
        Me.btCopyDetailToClipboard = New System.Windows.Forms.Button()
        Me.btDLLDetail = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbDLLVersion = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbDLLDate = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'lbAddInName
        '
        Me.lbAddInName.AutoSize = True
        Me.lbAddInName.Location = New System.Drawing.Point(11, 9)
        Me.lbAddInName.Name = "lbAddInName"
        Me.lbAddInName.Size = New System.Drawing.Size(94, 13)
        Me.lbAddInName.TabIndex = 0
        Me.lbAddInName.Text = "Sparx AddIn name"
        Me.ToolTip1.SetToolTip(Me.lbAddInName, "Name of key where the AddIn is define")
        '
        'lbAssemblyname
        '
        Me.lbAssemblyname.AutoSize = True
        Me.lbAssemblyname.Location = New System.Drawing.Point(14, 67)
        Me.lbAssemblyname.Name = "lbAssemblyname"
        Me.lbAssemblyname.Size = New System.Drawing.Size(61, 13)
        Me.lbAssemblyname.TabIndex = 1
        Me.lbAssemblyname.Text = "Class name"
        Me.ToolTip1.SetToolTip(Me.lbAssemblyname, "Class name as defined in Spaxr AddIn key")
        '
        'tbAddInName
        '
        Me.tbAddInName.Location = New System.Drawing.Point(120, 6)
        Me.tbAddInName.Name = "tbAddInName"
        Me.tbAddInName.ReadOnly = True
        Me.tbAddInName.Size = New System.Drawing.Size(502, 20)
        Me.tbAddInName.TabIndex = 2
        '
        'lbDLL
        '
        Me.lbDLL.AutoSize = True
        Me.lbDLL.Location = New System.Drawing.Point(14, 173)
        Me.lbDLL.Name = "lbDLL"
        Me.lbDLL.Size = New System.Drawing.Size(73, 13)
        Me.lbDLL.TabIndex = 3
        Me.lbDLL.Text = "DLL - full path"
        Me.ToolTip1.SetToolTip(Me.lbDLL, "Filename as defined in HKCR entry")
        '
        'tbAssemblyName
        '
        Me.tbAssemblyName.Location = New System.Drawing.Point(120, 64)
        Me.tbAssemblyName.Name = "tbAssemblyName"
        Me.tbAssemblyName.ReadOnly = True
        Me.tbAssemblyName.Size = New System.Drawing.Size(502, 20)
        Me.tbAssemblyName.TabIndex = 4
        '
        'tbDLL
        '
        Me.tbDLL.Location = New System.Drawing.Point(120, 151)
        Me.tbDLL.Multiline = True
        Me.tbDLL.Name = "tbDLL"
        Me.tbDLL.ReadOnly = True
        Me.tbDLL.Size = New System.Drawing.Size(502, 57)
        Me.tbDLL.TabIndex = 5
        '
        'btClose
        '
        Me.btClose.BackColor = System.Drawing.Color.SpringGreen
        Me.btClose.Location = New System.Drawing.Point(540, 248)
        Me.btClose.Name = "btClose"
        Me.btClose.Size = New System.Drawing.Size(75, 26)
        Me.btClose.TabIndex = 6
        Me.btClose.Text = "Close"
        Me.btClose.UseVisualStyleBackColor = False
        '
        'lbSparxRef
        '
        Me.lbSparxRef.AutoSize = True
        Me.lbSparxRef.Location = New System.Drawing.Point(11, 35)
        Me.lbSparxRef.Name = "lbSparxRef"
        Me.lbSparxRef.Size = New System.Drawing.Size(100, 13)
        Me.lbSparxRef.TabIndex = 7
        Me.lbSparxRef.Text = "Sparx entry location"
        Me.ToolTip1.SetToolTip(Me.lbSparxRef, "Location of Sparx AddIn entry")
        '
        'tbSparxRef
        '
        Me.tbSparxRef.Location = New System.Drawing.Point(120, 35)
        Me.tbSparxRef.Name = "tbSparxRef"
        Me.tbSparxRef.ReadOnly = True
        Me.tbSparxRef.Size = New System.Drawing.Size(502, 20)
        Me.tbSparxRef.TabIndex = 8
        '
        'lbClasssSrc
        '
        Me.lbClasssSrc.AutoSize = True
        Me.lbClasssSrc.Location = New System.Drawing.Point(211, 96)
        Me.lbClasssSrc.Name = "lbClasssSrc"
        Me.lbClasssSrc.Size = New System.Drawing.Size(81, 13)
        Me.lbClasssSrc.TabIndex = 9
        Me.lbClasssSrc.Text = "Class defined in"
        Me.ToolTip1.SetToolTip(Me.lbClasssSrc, resources.GetString("lbClasssSrc.ToolTip"))
        '
        'tbClassSource
        '
        Me.tbClassSource.Location = New System.Drawing.Point(298, 93)
        Me.tbClassSource.Name = "tbClassSource"
        Me.tbClassSource.ReadOnly = True
        Me.tbClassSource.Size = New System.Drawing.Size(325, 20)
        Me.tbClassSource.TabIndex = 10
        '
        'lbCLSID
        '
        Me.lbCLSID.AutoSize = True
        Me.lbCLSID.Location = New System.Drawing.Point(14, 125)
        Me.lbCLSID.Name = "lbCLSID"
        Me.lbCLSID.Size = New System.Drawing.Size(46, 13)
        Me.lbCLSID.TabIndex = 11
        Me.lbCLSID.Text = "Class ID"
        '
        'tbCLSID
        '
        Me.tbCLSID.Location = New System.Drawing.Point(120, 122)
        Me.tbCLSID.Name = "tbCLSID"
        Me.tbCLSID.ReadOnly = True
        Me.tbCLSID.Size = New System.Drawing.Size(502, 20)
        Me.tbCLSID.TabIndex = 12
        '
        'lbCLSIDSRC
        '
        Me.lbCLSIDSRC.AutoSize = True
        Me.lbCLSIDSRC.Location = New System.Drawing.Point(14, 96)
        Me.lbCLSIDSRC.Name = "lbCLSIDSRC"
        Me.lbCLSIDSRC.Size = New System.Drawing.Size(81, 13)
        Me.lbCLSIDSRC.TabIndex = 13
        Me.lbCLSIDSRC.Text = "Class ID source"
        '
        'tbCLSIDSRC
        '
        Me.tbCLSIDSRC.Location = New System.Drawing.Point(120, 93)
        Me.tbCLSIDSRC.Name = "tbCLSIDSRC"
        Me.tbCLSIDSRC.ReadOnly = True
        Me.tbCLSIDSRC.Size = New System.Drawing.Size(85, 20)
        Me.tbCLSIDSRC.TabIndex = 14
        '
        'btCopyDetailToClipboard
        '
        Me.btCopyDetailToClipboard.BackColor = System.Drawing.Color.Gold
        Me.btCopyDetailToClipboard.Location = New System.Drawing.Point(89, 248)
        Me.btCopyDetailToClipboard.Name = "btCopyDetailToClipboard"
        Me.btCopyDetailToClipboard.Size = New System.Drawing.Size(116, 26)
        Me.btCopyDetailToClipboard.TabIndex = 15
        Me.btCopyDetailToClipboard.Text = "Copy to clipboard"
        Me.btCopyDetailToClipboard.UseVisualStyleBackColor = False
        '
        'btDLLDetail
        '
        Me.btDLLDetail.BackColor = System.Drawing.Color.Yellow
        Me.btDLLDetail.Location = New System.Drawing.Point(211, 248)
        Me.btDLLDetail.Name = "btDLLDetail"
        Me.btDLLDetail.Size = New System.Drawing.Size(189, 26)
        Me.btDLLDetail.TabIndex = 16
        Me.btDLLDetail.Text = "Get list of classes and methods"
        Me.btDLLDetail.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 221)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "DLL version"
        '
        'tbDLLVersion
        '
        Me.tbDLLVersion.Location = New System.Drawing.Point(120, 221)
        Me.tbDLLVersion.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.tbDLLVersion.Name = "tbDLLVersion"
        Me.tbDLLVersion.ReadOnly = True
        Me.tbDLLVersion.Size = New System.Drawing.Size(138, 20)
        Me.tbDLLVersion.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(277, 225)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "File datetime"
        '
        'tbDLLDate
        '
        Me.tbDLLDate.Location = New System.Drawing.Point(391, 221)
        Me.tbDLLDate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.tbDLLDate.Name = "tbDLLDate"
        Me.tbDLLDate.ReadOnly = True
        Me.tbDLLDate.Size = New System.Drawing.Size(231, 20)
        Me.tbDLLDate.TabIndex = 20
        '
        'frmEntryDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(634, 282)
        Me.Controls.Add(Me.tbDLLDate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbDLLVersion)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btDLLDetail)
        Me.Controls.Add(Me.btCopyDetailToClipboard)
        Me.Controls.Add(Me.tbCLSIDSRC)
        Me.Controls.Add(Me.lbCLSIDSRC)
        Me.Controls.Add(Me.tbCLSID)
        Me.Controls.Add(Me.lbCLSID)
        Me.Controls.Add(Me.tbClassSource)
        Me.Controls.Add(Me.lbClasssSrc)
        Me.Controls.Add(Me.tbSparxRef)
        Me.Controls.Add(Me.lbSparxRef)
        Me.Controls.Add(Me.btClose)
        Me.Controls.Add(Me.tbDLL)
        Me.Controls.Add(Me.tbAssemblyName)
        Me.Controls.Add(Me.lbDLL)
        Me.Controls.Add(Me.tbAddInName)
        Me.Controls.Add(Me.lbAssemblyname)
        Me.Controls.Add(Me.lbAddInName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(654, 325)
        Me.MinimumSize = New System.Drawing.Size(654, 325)
        Me.Name = "frmEntryDetail"
        Me.Text = "EA AddIn Entry & DLL Details"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbAssemblyname As System.Windows.Forms.Label
    Friend WithEvents tbAddInName As System.Windows.Forms.TextBox
    Friend WithEvents lbDLL As System.Windows.Forms.Label
    Friend WithEvents tbAssemblyName As System.Windows.Forms.TextBox
    Friend WithEvents tbDLL As System.Windows.Forms.TextBox
    Friend WithEvents btClose As System.Windows.Forms.Button
    Friend WithEvents lbSparxRef As System.Windows.Forms.Label
    Friend WithEvents tbSparxRef As System.Windows.Forms.TextBox
    Friend WithEvents lbClasssSrc As System.Windows.Forms.Label
    Friend WithEvents tbClassSource As System.Windows.Forms.TextBox
    Friend WithEvents lbCLSID As System.Windows.Forms.Label
    Friend WithEvents tbCLSID As System.Windows.Forms.TextBox
    Friend WithEvents lbCLSIDSRC As System.Windows.Forms.Label
    Friend WithEvents tbCLSIDSRC As System.Windows.Forms.TextBox
    Friend WithEvents btCopyDetailToClipboard As System.Windows.Forms.Button
    Friend WithEvents btDLLDetail As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbDLLVersion As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbDLLDate As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As ToolTip
    Private WithEvents lbAddInName As Label
End Class

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
        Me.SuspendLayout()
        '
        'lbAddInName
        '
        Me.lbAddInName.AutoSize = True
        Me.lbAddInName.Location = New System.Drawing.Point(16, 23)
        Me.lbAddInName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbAddInName.Name = "lbAddInName"
        Me.lbAddInName.Size = New System.Drawing.Size(83, 17)
        Me.lbAddInName.TabIndex = 0
        Me.lbAddInName.Text = "AddIn name"
        '
        'lbAssemblyname
        '
        Me.lbAssemblyname.AutoSize = True
        Me.lbAssemblyname.Location = New System.Drawing.Point(269, 57)
        Me.lbAssemblyname.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbAssemblyname.Name = "lbAssemblyname"
        Me.lbAssemblyname.Size = New System.Drawing.Size(42, 17)
        Me.lbAssemblyname.TabIndex = 1
        Me.lbAssemblyname.Text = "Class"
        '
        'tbAddInName
        '
        Me.tbAddInName.Location = New System.Drawing.Point(116, 20)
        Me.tbAddInName.Margin = New System.Windows.Forms.Padding(4)
        Me.tbAddInName.Name = "tbAddInName"
        Me.tbAddInName.ReadOnly = True
        Me.tbAddInName.Size = New System.Drawing.Size(323, 22)
        Me.tbAddInName.TabIndex = 2
        '
        'lbDLL
        '
        Me.lbDLL.AutoSize = True
        Me.lbDLL.Location = New System.Drawing.Point(16, 160)
        Me.lbDLL.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbDLL.Name = "lbDLL"
        Me.lbDLL.Size = New System.Drawing.Size(97, 17)
        Me.lbDLL.TabIndex = 3
        Me.lbDLL.Text = "DLL - full path"
        '
        'tbAssemblyName
        '
        Me.tbAssemblyName.Location = New System.Drawing.Point(269, 76)
        Me.tbAssemblyName.Margin = New System.Windows.Forms.Padding(4)
        Me.tbAssemblyName.Name = "tbAssemblyName"
        Me.tbAssemblyName.ReadOnly = True
        Me.tbAssemblyName.Size = New System.Drawing.Size(549, 22)
        Me.tbAssemblyName.TabIndex = 4
        '
        'tbDLL
        '
        Me.tbDLL.Location = New System.Drawing.Point(116, 160)
        Me.tbDLL.Margin = New System.Windows.Forms.Padding(4)
        Me.tbDLL.Multiline = True
        Me.tbDLL.Name = "tbDLL"
        Me.tbDLL.ReadOnly = True
        Me.tbDLL.Size = New System.Drawing.Size(703, 69)
        Me.tbDLL.TabIndex = 5
        '
        'btClose
        '
        Me.btClose.BackColor = System.Drawing.Color.SpringGreen
        Me.btClose.Location = New System.Drawing.Point(718, 275)
        Me.btClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btClose.Name = "btClose"
        Me.btClose.Size = New System.Drawing.Size(100, 32)
        Me.btClose.TabIndex = 6
        Me.btClose.Text = "Close"
        Me.btClose.UseVisualStyleBackColor = False
        '
        'lbSparxRef
        '
        Me.lbSparxRef.AutoSize = True
        Me.lbSparxRef.Location = New System.Drawing.Point(451, 23)
        Me.lbSparxRef.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbSparxRef.Name = "lbSparxRef"
        Me.lbSparxRef.Size = New System.Drawing.Size(80, 17)
        Me.lbSparxRef.TabIndex = 7
        Me.lbSparxRef.Text = "Sparx entry"
        '
        'tbSparxRef
        '
        Me.tbSparxRef.Location = New System.Drawing.Point(539, 20)
        Me.tbSparxRef.Margin = New System.Windows.Forms.Padding(4)
        Me.tbSparxRef.Name = "tbSparxRef"
        Me.tbSparxRef.ReadOnly = True
        Me.tbSparxRef.Size = New System.Drawing.Size(280, 22)
        Me.tbSparxRef.TabIndex = 8
        '
        'lbClasssSrc
        '
        Me.lbClasssSrc.AutoSize = True
        Me.lbClasssSrc.Location = New System.Drawing.Point(116, 55)
        Me.lbClasssSrc.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbClasssSrc.Name = "lbClasssSrc"
        Me.lbClasssSrc.Size = New System.Drawing.Size(89, 17)
        Me.lbClasssSrc.TabIndex = 9
        Me.lbClasssSrc.Text = "Class source"
        '
        'tbClassSource
        '
        Me.tbClassSource.Location = New System.Drawing.Point(116, 76)
        Me.tbClassSource.Margin = New System.Windows.Forms.Padding(4)
        Me.tbClassSource.Name = "tbClassSource"
        Me.tbClassSource.ReadOnly = True
        Me.tbClassSource.Size = New System.Drawing.Size(112, 22)
        Me.tbClassSource.TabIndex = 10
        '
        'lbCLSID
        '
        Me.lbCLSID.AutoSize = True
        Me.lbCLSID.Location = New System.Drawing.Point(269, 108)
        Me.lbCLSID.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbCLSID.Name = "lbCLSID"
        Me.lbCLSID.Size = New System.Drawing.Size(59, 17)
        Me.lbCLSID.TabIndex = 11
        Me.lbCLSID.Text = "Class ID"
        '
        'tbCLSID
        '
        Me.tbCLSID.Location = New System.Drawing.Point(269, 128)
        Me.tbCLSID.Margin = New System.Windows.Forms.Padding(4)
        Me.tbCLSID.Name = "tbCLSID"
        Me.tbCLSID.ReadOnly = True
        Me.tbCLSID.Size = New System.Drawing.Size(549, 22)
        Me.tbCLSID.TabIndex = 12
        '
        'lbCLSIDSRC
        '
        Me.lbCLSIDSRC.AutoSize = True
        Me.lbCLSIDSRC.Location = New System.Drawing.Point(116, 108)
        Me.lbCLSIDSRC.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbCLSIDSRC.Name = "lbCLSIDSRC"
        Me.lbCLSIDSRC.Size = New System.Drawing.Size(106, 17)
        Me.lbCLSIDSRC.TabIndex = 13
        Me.lbCLSIDSRC.Text = "Class ID source"
        '
        'tbCLSIDSRC
        '
        Me.tbCLSIDSRC.Location = New System.Drawing.Point(116, 128)
        Me.tbCLSIDSRC.Margin = New System.Windows.Forms.Padding(4)
        Me.tbCLSIDSRC.Name = "tbCLSIDSRC"
        Me.tbCLSIDSRC.ReadOnly = True
        Me.tbCLSIDSRC.Size = New System.Drawing.Size(112, 22)
        Me.tbCLSIDSRC.TabIndex = 14
        '
        'btCopyDetailToClipboard
        '
        Me.btCopyDetailToClipboard.BackColor = System.Drawing.Color.Gold
        Me.btCopyDetailToClipboard.Location = New System.Drawing.Point(116, 275)
        Me.btCopyDetailToClipboard.Margin = New System.Windows.Forms.Padding(4)
        Me.btCopyDetailToClipboard.Name = "btCopyDetailToClipboard"
        Me.btCopyDetailToClipboard.Size = New System.Drawing.Size(155, 32)
        Me.btCopyDetailToClipboard.TabIndex = 15
        Me.btCopyDetailToClipboard.Text = "Copy to clipboard"
        Me.btCopyDetailToClipboard.UseVisualStyleBackColor = False
        '
        'btDLLDetail
        '
        Me.btDLLDetail.BackColor = System.Drawing.Color.Yellow
        Me.btDLLDetail.Location = New System.Drawing.Point(279, 275)
        Me.btDLLDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.btDLLDetail.Name = "btDLLDetail"
        Me.btDLLDetail.Size = New System.Drawing.Size(252, 32)
        Me.btDLLDetail.TabIndex = 16
        Me.btDLLDetail.Text = "Get list of classes and methods"
        Me.btDLLDetail.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 241)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 17)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "DLL version"
        '
        'tbDLLVersion
        '
        Me.tbDLLVersion.Location = New System.Drawing.Point(116, 236)
        Me.tbDLLVersion.Name = "tbDLLVersion"
        Me.tbDLLVersion.ReadOnly = True
        Me.tbDLLVersion.Size = New System.Drawing.Size(236, 22)
        Me.tbDLLVersion.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(370, 241)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 17)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "File datetime"
        '
        'tbDLLDate
        '
        Me.tbDLLDate.Location = New System.Drawing.Point(464, 241)
        Me.tbDLLDate.Name = "tbDLLDate"
        Me.tbDLLDate.ReadOnly = True
        Me.tbDLLDate.Size = New System.Drawing.Size(354, 22)
        Me.tbDLLDate.TabIndex = 20
        '
        'frmEntryDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(843, 339)
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
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximumSize = New System.Drawing.Size(865, 390)
        Me.MinimumSize = New System.Drawing.Size(865, 390)
        Me.Name = "frmEntryDetail"
        Me.Text = "EA AddIn Entry & DLL Details"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbAddInName As System.Windows.Forms.Label
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
End Class

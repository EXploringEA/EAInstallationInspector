' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
' You may use, distribute and modify this code under the terms of the 3-Clause BSD License
'
' You should have received a copy of the 3-Clause BSD License with this file. 
' If not, please email: eaForms@EXploringEA.co.uk 
'=====================================================================================
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInspector3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInspector3))
        Me.lv2 = New System.Windows.Forms.ListView()
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
        Me.SuspendLayout()
        '
        'lv2
        '
        Me.lv2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lv2.FullRowSelect = True
        Me.lv2.Location = New System.Drawing.Point(12, 32)
        Me.lv2.MultiSelect = False
        Me.lv2.Name = "lv2"
        Me.lv2.Size = New System.Drawing.Size(1137, 340)
        Me.lv2.TabIndex = 0
        Me.lv2.UseCompatibleStateImageBehavior = False
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
        Me.btHelp.Location = New System.Drawing.Point(12, 379)
        Me.btHelp.Name = "btHelp"
        Me.btHelp.Size = New System.Drawing.Size(75, 23)
        Me.btHelp.TabIndex = 7
        Me.btHelp.Text = "Help"
        Me.btHelp.UseVisualStyleBackColor = True
        '
        'btRefresh
        '
        Me.btRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btRefresh.Location = New System.Drawing.Point(94, 378)
        Me.btRefresh.Name = "btRefresh"
        Me.btRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btRefresh.TabIndex = 8
        Me.btRefresh.Text = "Refresh"
        Me.btRefresh.UseVisualStyleBackColor = True
        '
        'btCopy
        '
        Me.btCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btCopy.Location = New System.Drawing.Point(176, 378)
        Me.btCopy.Name = "btCopy"
        Me.btCopy.Size = New System.Drawing.Size(217, 23)
        Me.btCopy.TabIndex = 9
        Me.btCopy.Text = "Copy current list image to clipboard"
        Me.btCopy.UseVisualStyleBackColor = True
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
        'frmInspector3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1161, 416)
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
        Me.Controls.Add(Me.lv2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInspector3"
        Me.Text = "EA Installation Inspector Version 3"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lv2 As System.Windows.Forms.ListView
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
End Class

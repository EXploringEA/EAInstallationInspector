' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
' You may use, distribute and modify this code under the terms of the 3-Clause BSD License
'
' You should have received a copy of the 3-Clause BSD License with this file. 
' If not, please email: adrian@EXploringEA.co.uk 
'=====================================================================================
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmListOfClasses
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmListOfClasses))
        Me.btClose = New System.Windows.Forms.Button()
        Me.dgView = New System.Windows.Forms.DataGridView()
        Me.btSelectAll = New System.Windows.Forms.Button()
        CType(Me.dgView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btClose
        '
        Me.btClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btClose.Location = New System.Drawing.Point(884, 460)
        Me.btClose.Name = "btClose"
        Me.btClose.Size = New System.Drawing.Size(75, 23)
        Me.btClose.TabIndex = 0
        Me.btClose.Text = "Close"
        Me.btClose.UseVisualStyleBackColor = True
        '
        'dgView
        '
        Me.dgView.AllowUserToAddRows = False
        Me.dgView.AllowUserToDeleteRows = False
        Me.dgView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgView.Location = New System.Drawing.Point(12, 12)
        Me.dgView.Name = "dgView"
        Me.dgView.ReadOnly = True
        Me.dgView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgView.Size = New System.Drawing.Size(947, 442)
        Me.dgView.TabIndex = 2
        '
        'btSelectAll
        '
        Me.btSelectAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btSelectAll.Location = New System.Drawing.Point(13, 460)
        Me.btSelectAll.Name = "btSelectAll"
        Me.btSelectAll.Size = New System.Drawing.Size(202, 23)
        Me.btSelectAll.TabIndex = 3
        Me.btSelectAll.Text = "Select all and copy to clipboard"
        Me.btSelectAll.UseVisualStyleBackColor = True
        '
        'frmListOfMethods
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(971, 495)
        Me.Controls.Add(Me.btSelectAll)
        Me.Controls.Add(Me.dgView)
        Me.Controls.Add(Me.btClose)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmListOfMethods"
        Me.Text = "List of AddIn public classes / methods"
        CType(Me.dgView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btClose As System.Windows.Forms.Button
    Friend WithEvents dgView As System.Windows.Forms.DataGridView
    Friend WithEvents btSelectAll As System.Windows.Forms.Button
End Class

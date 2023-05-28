<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDisplayRegistryKeys
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDisplayRegistryKeys))
        Me.lbRegKeys = New System.Windows.Forms.ListBox()
        Me.btCreate = New System.Windows.Forms.Button()
        Me.btDelete = New System.Windows.Forms.Button()
        Me.btExportFiles = New System.Windows.Forms.Button()
        Me.cbIncludeAssemblies = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'lbRegKeys
        '
        Me.lbRegKeys.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbRegKeys.FormattingEnabled = True
        Me.lbRegKeys.Location = New System.Drawing.Point(0, 0)
        Me.lbRegKeys.Name = "lbRegKeys"
        Me.lbRegKeys.Size = New System.Drawing.Size(779, 563)
        Me.lbRegKeys.TabIndex = 0
        '
        'btCreate
        '
        Me.btCreate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btCreate.Location = New System.Drawing.Point(12, 572)
        Me.btCreate.Name = "btCreate"
        Me.btCreate.Size = New System.Drawing.Size(75, 23)
        Me.btCreate.TabIndex = 2
        Me.btCreate.Text = "Add keys"
        Me.btCreate.UseVisualStyleBackColor = True
        '
        'btDelete
        '
        Me.btDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btDelete.Location = New System.Drawing.Point(134, 572)
        Me.btDelete.Name = "btDelete"
        Me.btDelete.Size = New System.Drawing.Size(75, 23)
        Me.btDelete.TabIndex = 3
        Me.btDelete.Text = "Delete keys"
        Me.btDelete.UseVisualStyleBackColor = True
        '
        'btExportFiles
        '
        Me.btExportFiles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btExportFiles.Location = New System.Drawing.Point(256, 574)
        Me.btExportFiles.Name = "btExportFiles"
        Me.btExportFiles.Size = New System.Drawing.Size(75, 23)
        Me.btExportFiles.TabIndex = 4
        Me.btExportFiles.Text = "Export files"
        Me.btExportFiles.UseVisualStyleBackColor = True
        '
        'cbIncludeAssemblies
        '
        Me.cbIncludeAssemblies.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbIncludeAssemblies.AutoSize = True
        Me.cbIncludeAssemblies.Checked = True
        Me.cbIncludeAssemblies.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbIncludeAssemblies.Location = New System.Drawing.Point(664, 578)
        Me.cbIncludeAssemblies.Name = "cbIncludeAssemblies"
        Me.cbIncludeAssemblies.Size = New System.Drawing.Size(115, 17)
        Me.cbIncludeAssemblies.TabIndex = 5
        Me.cbIncludeAssemblies.Text = "Include assemblies"
        Me.cbIncludeAssemblies.UseVisualStyleBackColor = True
        '
        'frmDisplayRegistryKeys
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(780, 606)
        Me.Controls.Add(Me.cbIncludeAssemblies)
        Me.Controls.Add(Me.btExportFiles)
        Me.Controls.Add(Me.btDelete)
        Me.Controls.Add(Me.btCreate)
        Me.Controls.Add(Me.lbRegKeys)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDisplayRegistryKeys"
        Me.Text = "Registry keys for selected item"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbRegKeys As ListBox


    Friend WithEvents btCreate As Button
    Friend WithEvents btDelete As Button
    Friend WithEvents btExportFiles As Button
    Friend WithEvents cbIncludeAssemblies As CheckBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQuery
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQuery))
        Me.rbHKCU = New System.Windows.Forms.RadioButton()
        Me.rbHKLM = New System.Windows.Forms.RadioButton()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnRunQuery = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.tbQuery = New System.Windows.Forms.TextBox()
        Me.btnSelectFile = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'rbHKCU
        '
        Me.rbHKCU.AutoSize = True
        Me.rbHKCU.Checked = True
        Me.rbHKCU.Location = New System.Drawing.Point(4, 4)
        Me.rbHKCU.Name = "rbHKCU"
        Me.rbHKCU.Size = New System.Drawing.Size(55, 17)
        Me.rbHKCU.TabIndex = 0
        Me.rbHKCU.TabStop = True
        Me.rbHKCU.Text = "HKCU"
        Me.rbHKCU.UseVisualStyleBackColor = True
        '
        'rbHKLM
        '
        Me.rbHKLM.AutoSize = True
        Me.rbHKLM.Checked = True
        Me.rbHKLM.Location = New System.Drawing.Point(4, 5)
        Me.rbHKLM.Name = "rbHKLM"
        Me.rbHKLM.Size = New System.Drawing.Size(55, 17)
        Me.rbHKLM.TabIndex = 1
        Me.rbHKLM.TabStop = True
        Me.rbHKLM.Text = "HKLM"
        Me.rbHKLM.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnRunQuery
        '
        Me.btnRunQuery.Location = New System.Drawing.Point(105, 116)
        Me.btnRunQuery.Name = "btnRunQuery"
        Me.btnRunQuery.Size = New System.Drawing.Size(75, 23)
        Me.btnRunQuery.TabIndex = 2
        Me.btnRunQuery.Text = "Run query"
        Me.btnRunQuery.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(12, 116)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'tbQuery
        '
        Me.tbQuery.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbQuery.Location = New System.Drawing.Point(12, 81)
        Me.tbQuery.Name = "tbQuery"
        Me.tbQuery.Size = New System.Drawing.Size(354, 20)
        Me.tbQuery.TabIndex = 4
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(12, 12)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(75, 23)
        Me.btnSelectFile.TabIndex = 5
        Me.btnSelectFile.Text = "Select file"
        Me.btnSelectFile.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rbHKLM)
        Me.Panel1.Location = New System.Drawing.Point(218, 10)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(75, 25)
        Me.Panel1.TabIndex = 6
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.rbHKCU)
        Me.Panel2.Location = New System.Drawing.Point(121, 10)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(75, 25)
        Me.Panel2.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(217, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Selected filename or enter query search term"
        '
        'frmQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(378, 161)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnSelectFile)
        Me.Controls.Add(Me.tbQuery)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnRunQuery)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(1000, 200)
        Me.MinimumSize = New System.Drawing.Size(0, 200)
        Me.Name = "frmQuery"
        Me.Text = "Custom query"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rbHKCU As System.Windows.Forms.RadioButton
    Friend WithEvents rbHKLM As System.Windows.Forms.RadioButton
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnRunQuery As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents tbQuery As System.Windows.Forms.TextBox
    Friend WithEvents btnSelectFile As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class

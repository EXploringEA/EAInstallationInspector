<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLegend
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
        Me.pbLegend = New System.Windows.Forms.PictureBox()
        CType(Me.pbLegend, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbLegend
        '
        Me.pbLegend.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pbLegend.Location = New System.Drawing.Point(0, 0)
        Me.pbLegend.Name = "pbLegend"
        Me.pbLegend.Size = New System.Drawing.Size(617, 152)
        Me.pbLegend.TabIndex = 0
        Me.pbLegend.TabStop = False
        '
        'frmLegend
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(616, 151)
        Me.Controls.Add(Me.pbLegend)
        Me.MaximumSize = New System.Drawing.Size(632, 190)
        Me.MinimumSize = New System.Drawing.Size(632, 145)
        Me.Name = "frmLegend"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Meaning of colours for row entries"
        Me.TopMost = True
        CType(Me.pbLegend, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pbLegend As PictureBox
End Class

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
        Me.SuspendLayout()
        '
        'lbAddInName
        '
        Me.lbAddInName.AutoSize = True
        Me.lbAddInName.Location = New System.Drawing.Point(12, 19)
        Me.lbAddInName.Name = "lbAddInName"
        Me.lbAddInName.Size = New System.Drawing.Size(64, 13)
        Me.lbAddInName.TabIndex = 0
        Me.lbAddInName.Text = "AddIn name"
        '
        'lbAssemblyname
        '
        Me.lbAssemblyname.AutoSize = True
        Me.lbAssemblyname.Location = New System.Drawing.Point(202, 46)
        Me.lbAssemblyname.Name = "lbAssemblyname"
        Me.lbAssemblyname.Size = New System.Drawing.Size(32, 13)
        Me.lbAssemblyname.TabIndex = 1
        Me.lbAssemblyname.Text = "Class"
        '
        'tbAddInName
        '
        Me.tbAddInName.Location = New System.Drawing.Point(87, 16)
        Me.tbAddInName.Name = "tbAddInName"
        Me.tbAddInName.ReadOnly = True
        Me.tbAddInName.Size = New System.Drawing.Size(243, 20)
        Me.tbAddInName.TabIndex = 2
        '
        'lbDLL
        '
        Me.lbDLL.AutoSize = True
        Me.lbDLL.Location = New System.Drawing.Point(12, 130)
        Me.lbDLL.Name = "lbDLL"
        Me.lbDLL.Size = New System.Drawing.Size(27, 13)
        Me.lbDLL.TabIndex = 3
        Me.lbDLL.Text = "DLL"
        '
        'tbAssemblyName
        '
        Me.tbAssemblyName.Location = New System.Drawing.Point(202, 62)
        Me.tbAssemblyName.Name = "tbAssemblyName"
        Me.tbAssemblyName.ReadOnly = True
        Me.tbAssemblyName.Size = New System.Drawing.Size(413, 20)
        Me.tbAssemblyName.TabIndex = 4
        '
        'tbDLL
        '
        Me.tbDLL.Location = New System.Drawing.Point(87, 130)
        Me.tbDLL.Multiline = True
        Me.tbDLL.Name = "tbDLL"
        Me.tbDLL.ReadOnly = True
        Me.tbDLL.Size = New System.Drawing.Size(528, 68)
        Me.tbDLL.TabIndex = 5
        '
        'btClose
        '
        Me.btClose.Location = New System.Drawing.Point(540, 213)
        Me.btClose.Name = "btClose"
        Me.btClose.Size = New System.Drawing.Size(75, 23)
        Me.btClose.TabIndex = 6
        Me.btClose.Text = "Close"
        Me.btClose.UseVisualStyleBackColor = True
        '
        'lbSparxRef
        '
        Me.lbSparxRef.AutoSize = True
        Me.lbSparxRef.Location = New System.Drawing.Point(338, 19)
        Me.lbSparxRef.Name = "lbSparxRef"
        Me.lbSparxRef.Size = New System.Drawing.Size(60, 13)
        Me.lbSparxRef.TabIndex = 7
        Me.lbSparxRef.Text = "Sparx entry"
        '
        'tbSparxRef
        '
        Me.tbSparxRef.Location = New System.Drawing.Point(404, 16)
        Me.tbSparxRef.Name = "tbSparxRef"
        Me.tbSparxRef.ReadOnly = True
        Me.tbSparxRef.Size = New System.Drawing.Size(211, 20)
        Me.tbSparxRef.TabIndex = 8
        '
        'lbClasssSrc
        '
        Me.lbClasssSrc.AutoSize = True
        Me.lbClasssSrc.Location = New System.Drawing.Point(87, 46)
        Me.lbClasssSrc.Name = "lbClasssSrc"
        Me.lbClasssSrc.Size = New System.Drawing.Size(67, 13)
        Me.lbClasssSrc.TabIndex = 9
        Me.lbClasssSrc.Text = "Class source"
        '
        'tbClassSource
        '
        Me.tbClassSource.Location = New System.Drawing.Point(87, 62)
        Me.tbClassSource.Name = "tbClassSource"
        Me.tbClassSource.ReadOnly = True
        Me.tbClassSource.Size = New System.Drawing.Size(85, 20)
        Me.tbClassSource.TabIndex = 10
        '
        'lbCLSID
        '
        Me.lbCLSID.AutoSize = True
        Me.lbCLSID.Location = New System.Drawing.Point(202, 88)
        Me.lbCLSID.Name = "lbCLSID"
        Me.lbCLSID.Size = New System.Drawing.Size(46, 13)
        Me.lbCLSID.TabIndex = 11
        Me.lbCLSID.Text = "Class ID"
        '
        'tbCLSID
        '
        Me.tbCLSID.Location = New System.Drawing.Point(202, 104)
        Me.tbCLSID.Name = "tbCLSID"
        Me.tbCLSID.ReadOnly = True
        Me.tbCLSID.Size = New System.Drawing.Size(413, 20)
        Me.tbCLSID.TabIndex = 12
        '
        'lbCLSIDSRC
        '
        Me.lbCLSIDSRC.AutoSize = True
        Me.lbCLSIDSRC.Location = New System.Drawing.Point(87, 88)
        Me.lbCLSIDSRC.Name = "lbCLSIDSRC"
        Me.lbCLSIDSRC.Size = New System.Drawing.Size(81, 13)
        Me.lbCLSIDSRC.TabIndex = 13
        Me.lbCLSIDSRC.Text = "Class ID source"
        '
        'tbCLSIDSRC
        '
        Me.tbCLSIDSRC.Location = New System.Drawing.Point(87, 104)
        Me.tbCLSIDSRC.Name = "tbCLSIDSRC"
        Me.tbCLSIDSRC.ReadOnly = True
        Me.tbCLSIDSRC.Size = New System.Drawing.Size(85, 20)
        Me.tbCLSIDSRC.TabIndex = 14
        '
        'frmEntryDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(632, 248)
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
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEntryDetail"
        Me.Text = "EA AddIn Entry"
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
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OPTIONS
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
        Me.cbAutostart = New System.Windows.Forms.CheckBox()
        Me.cbNotifications = New System.Windows.Forms.CheckBox()
        Me.cbMinimize = New System.Windows.Forms.CheckBox()
        Me.cbAutoActivateExcelMacro = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'cbAutostart
        '
        Me.cbAutostart.AutoSize = True
        Me.cbAutostart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAutostart.Location = New System.Drawing.Point(15, 48)
        Me.cbAutostart.Name = "cbAutostart"
        Me.cbAutostart.Size = New System.Drawing.Size(241, 20)
        Me.cbAutostart.TabIndex = 0
        Me.cbAutostart.Text = "Démarrage à l'ouverture de session"
        Me.cbAutostart.UseVisualStyleBackColor = True
        '
        'cbNotifications
        '
        Me.cbNotifications.AutoSize = True
        Me.cbNotifications.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbNotifications.Location = New System.Drawing.Point(15, 12)
        Me.cbNotifications.Name = "cbNotifications"
        Me.cbNotifications.Size = New System.Drawing.Size(237, 20)
        Me.cbNotifications.TabIndex = 1
        Me.cbNotifications.Text = "Alertes de mise à jour du catalogue"
        Me.cbNotifications.UseVisualStyleBackColor = True
        '
        'cbMinimize
        '
        Me.cbMinimize.AutoSize = True
        Me.cbMinimize.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbMinimize.Location = New System.Drawing.Point(15, 84)
        Me.cbMinimize.Name = "cbMinimize"
        Me.cbMinimize.Size = New System.Drawing.Size(246, 20)
        Me.cbMinimize.TabIndex = 2
        Me.cbMinimize.Text = "Réduire dans la zone de notifications"
        Me.cbMinimize.UseVisualStyleBackColor = True
        '
        'cbAutoActivateExcelMacro
        '
        Me.cbAutoActivateExcelMacro.AutoSize = True
        Me.cbAutoActivateExcelMacro.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAutoActivateExcelMacro.Location = New System.Drawing.Point(15, 122)
        Me.cbAutoActivateExcelMacro.Name = "cbAutoActivateExcelMacro"
        Me.cbAutoActivateExcelMacro.Size = New System.Drawing.Size(272, 20)
        Me.cbAutoActivateExcelMacro.TabIndex = 3
        Me.cbAutoActivateExcelMacro.Text = "Activation automatique de la macro Excel"
        Me.cbAutoActivateExcelMacro.UseVisualStyleBackColor = True
        '
        'OPTIONS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(292, 152)
        Me.Controls.Add(Me.cbAutoActivateExcelMacro)
        Me.Controls.Add(Me.cbMinimize)
        Me.Controls.Add(Me.cbNotifications)
        Me.Controls.Add(Me.cbAutostart)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OPTIONS"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Options"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbAutostart As System.Windows.Forms.CheckBox
    Friend WithEvents cbNotifications As System.Windows.Forms.CheckBox
    Friend WithEvents cbMinimize As System.Windows.Forms.CheckBox
    Friend WithEvents cbAutoActivateExcelMacro As System.Windows.Forms.CheckBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WAIT
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
        Me.labelLoading = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'labelLoading
        '
        Me.labelLoading.BackColor = System.Drawing.Color.Transparent
        Me.labelLoading.Image = Global.CATALOGUE.My.Resources.Resources.loading_large
        Me.labelLoading.Location = New System.Drawing.Point(5, 12)
        Me.labelLoading.Name = "labelLoading"
        Me.labelLoading.Size = New System.Drawing.Size(149, 54)
        Me.labelLoading.TabIndex = 0
        '
        'WAIT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(157, 74)
        Me.ControlBox = False
        Me.Controls.Add(Me.labelLoading)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WAIT"
        Me.Opacity = 0.8R
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chargement..."
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents labelLoading As System.Windows.Forms.Label
End Class

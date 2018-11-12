<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SEARCH
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
        Me.lvSearchResults = New System.Windows.Forms.ListView()
        Me.SuspendLayout()
        '
        'lvSearchResults
        '
        Me.lvSearchResults.FullRowSelect = True
        Me.lvSearchResults.GridLines = True
        Me.lvSearchResults.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvSearchResults.LabelWrap = False
        Me.lvSearchResults.Location = New System.Drawing.Point(12, 12)
        Me.lvSearchResults.MultiSelect = False
        Me.lvSearchResults.Name = "lvSearchResults"
        Me.lvSearchResults.Size = New System.Drawing.Size(636, 272)
        Me.lvSearchResults.TabIndex = 0
        Me.lvSearchResults.UseCompatibleStateImageBehavior = False
        Me.lvSearchResults.View = System.Windows.Forms.View.Details
        '
        'SEARCH
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(660, 296)
        Me.Controls.Add(Me.lvSearchResults)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SEARCH"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Résultat(s) de la recherche"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lvSearchResults As System.Windows.Forms.ListView
End Class

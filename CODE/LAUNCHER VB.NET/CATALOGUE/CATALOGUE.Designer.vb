<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CATALOGUE
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CATALOGUE))
        Me.cbAddSheet = New System.Windows.Forms.Button()
        Me.cbHTMLCatalogue = New System.Windows.Forms.Button()
        Me.cbXLSCatalogue = New System.Windows.Forms.Button()
        Me.panelInfo = New System.Windows.Forms.Panel()
        Me.labelInfo = New System.Windows.Forms.Label()
        Me.labelWaitHtml = New System.Windows.Forms.Label()
        Me.labelWaitXsl = New System.Windows.Forms.Label()
        Me.labelWaitAddSheet = New System.Windows.Forms.Label()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.labelWaitSearch = New System.Windows.Forms.Label()
        Me.cbSearch = New System.Windows.Forms.Button()
        Me.tbSearch = New System.Windows.Forms.TextBox()
        Me.panelInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbAddSheet
        '
        Me.cbAddSheet.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cbAddSheet.FlatAppearance.CheckedBackColor = System.Drawing.Color.White
        Me.cbAddSheet.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
        Me.cbAddSheet.FlatAppearance.MouseOverBackColor = System.Drawing.Color.WhiteSmoke
        Me.cbAddSheet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbAddSheet.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAddSheet.Image = Global.CATALOGUE.My.Resources.Resources.plus_circle_small
        Me.cbAddSheet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cbAddSheet.Location = New System.Drawing.Point(11, 88)
        Me.cbAddSheet.Name = "cbAddSheet"
        Me.cbAddSheet.Padding = New System.Windows.Forms.Padding(17, 0, 0, 0)
        Me.cbAddSheet.Size = New System.Drawing.Size(231, 32)
        Me.cbAddSheet.TabIndex = 3
        Me.cbAddSheet.Text = "Créer une nouvelle fiche       "
        Me.cbAddSheet.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbAddSheet.UseVisualStyleBackColor = True
        '
        'cbHTMLCatalogue
        '
        Me.cbHTMLCatalogue.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cbHTMLCatalogue.FlatAppearance.CheckedBackColor = System.Drawing.Color.White
        Me.cbHTMLCatalogue.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
        Me.cbHTMLCatalogue.FlatAppearance.MouseOverBackColor = System.Drawing.Color.WhiteSmoke
        Me.cbHTMLCatalogue.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbHTMLCatalogue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbHTMLCatalogue.Image = Global.CATALOGUE.My.Resources.Resources.book_open_page_variant_small
        Me.cbHTMLCatalogue.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cbHTMLCatalogue.Location = New System.Drawing.Point(11, 12)
        Me.cbHTMLCatalogue.Name = "cbHTMLCatalogue"
        Me.cbHTMLCatalogue.Padding = New System.Windows.Forms.Padding(16, 0, 0, 0)
        Me.cbHTMLCatalogue.Size = New System.Drawing.Size(231, 32)
        Me.cbHTMLCatalogue.TabIndex = 1
        Me.cbHTMLCatalogue.Text = "Ouvrir le catalogue                 "
        Me.cbHTMLCatalogue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbHTMLCatalogue.UseVisualStyleBackColor = True
        '
        'cbXLSCatalogue
        '
        Me.cbXLSCatalogue.BackColor = System.Drawing.Color.White
        Me.cbXLSCatalogue.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cbXLSCatalogue.FlatAppearance.CheckedBackColor = System.Drawing.Color.White
        Me.cbXLSCatalogue.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
        Me.cbXLSCatalogue.FlatAppearance.MouseOverBackColor = System.Drawing.Color.WhiteSmoke
        Me.cbXLSCatalogue.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbXLSCatalogue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbXLSCatalogue.Image = Global.CATALOGUE.My.Resources.Resources.database_small
        Me.cbXLSCatalogue.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cbXLSCatalogue.Location = New System.Drawing.Point(11, 50)
        Me.cbXLSCatalogue.Name = "cbXLSCatalogue"
        Me.cbXLSCatalogue.Padding = New System.Windows.Forms.Padding(18, 0, 0, 0)
        Me.cbXLSCatalogue.Size = New System.Drawing.Size(231, 32)
        Me.cbXLSCatalogue.TabIndex = 2
        Me.cbXLSCatalogue.Text = "Editer le catalogue                 "
        Me.cbXLSCatalogue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbXLSCatalogue.UseVisualStyleBackColor = False
        '
        'panelInfo
        '
        Me.panelInfo.BackColor = System.Drawing.Color.Black
        Me.panelInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panelInfo.Controls.Add(Me.labelInfo)
        Me.panelInfo.Location = New System.Drawing.Point(-16, 167)
        Me.panelInfo.Name = "panelInfo"
        Me.panelInfo.Size = New System.Drawing.Size(280, 39)
        Me.panelInfo.TabIndex = 4
        '
        'labelInfo
        '
        Me.labelInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelInfo.ForeColor = System.Drawing.Color.White
        Me.labelInfo.Location = New System.Drawing.Point(24, 3)
        Me.labelInfo.Name = "labelInfo"
        Me.labelInfo.Size = New System.Drawing.Size(251, 23)
        Me.labelInfo.TabIndex = 0
        '
        'labelWaitHtml
        '
        Me.labelWaitHtml.BackColor = System.Drawing.Color.WhiteSmoke
        Me.labelWaitHtml.Image = Global.CATALOGUE.My.Resources.Resources.loading
        Me.labelWaitHtml.Location = New System.Drawing.Point(30, 15)
        Me.labelWaitHtml.Name = "labelWaitHtml"
        Me.labelWaitHtml.Size = New System.Drawing.Size(31, 26)
        Me.labelWaitHtml.TabIndex = 5
        Me.labelWaitHtml.Visible = False
        '
        'labelWaitXsl
        '
        Me.labelWaitXsl.BackColor = System.Drawing.Color.WhiteSmoke
        Me.labelWaitXsl.CausesValidation = False
        Me.labelWaitXsl.ForeColor = System.Drawing.Color.Transparent
        Me.labelWaitXsl.Image = Global.CATALOGUE.My.Resources.Resources.loading
        Me.labelWaitXsl.Location = New System.Drawing.Point(30, 53)
        Me.labelWaitXsl.Name = "labelWaitXsl"
        Me.labelWaitXsl.Size = New System.Drawing.Size(31, 26)
        Me.labelWaitXsl.TabIndex = 6
        Me.labelWaitXsl.Visible = False
        '
        'labelWaitAddSheet
        '
        Me.labelWaitAddSheet.BackColor = System.Drawing.Color.WhiteSmoke
        Me.labelWaitAddSheet.CausesValidation = False
        Me.labelWaitAddSheet.ForeColor = System.Drawing.Color.Transparent
        Me.labelWaitAddSheet.Image = Global.CATALOGUE.My.Resources.Resources.loading
        Me.labelWaitAddSheet.Location = New System.Drawing.Point(30, 91)
        Me.labelWaitAddSheet.Name = "labelWaitAddSheet"
        Me.labelWaitAddSheet.Size = New System.Drawing.Size(31, 26)
        Me.labelWaitAddSheet.TabIndex = 7
        Me.labelWaitAddSheet.Visible = False
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "CATALOGUE"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'labelWaitSearch
        '
        Me.labelWaitSearch.BackColor = System.Drawing.Color.WhiteSmoke
        Me.labelWaitSearch.CausesValidation = False
        Me.labelWaitSearch.ForeColor = System.Drawing.Color.Transparent
        Me.labelWaitSearch.Image = Global.CATALOGUE.My.Resources.Resources.loading
        Me.labelWaitSearch.Location = New System.Drawing.Point(30, 129)
        Me.labelWaitSearch.Name = "labelWaitSearch"
        Me.labelWaitSearch.Size = New System.Drawing.Size(31, 26)
        Me.labelWaitSearch.TabIndex = 9
        Me.labelWaitSearch.Visible = False
        '
        'cbSearch
        '
        Me.cbSearch.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cbSearch.FlatAppearance.CheckedBackColor = System.Drawing.Color.White
        Me.cbSearch.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
        Me.cbSearch.FlatAppearance.MouseOverBackColor = System.Drawing.Color.WhiteSmoke
        Me.cbSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbSearch.Image = Global.CATALOGUE.My.Resources.Resources.search_small
        Me.cbSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cbSearch.Location = New System.Drawing.Point(11, 126)
        Me.cbSearch.Name = "cbSearch"
        Me.cbSearch.Padding = New System.Windows.Forms.Padding(19, 0, 0, 0)
        Me.cbSearch.Size = New System.Drawing.Size(231, 32)
        Me.cbSearch.TabIndex = 4
        Me.cbSearch.Text = "Rechercher dans les fiches  "
        Me.cbSearch.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbSearch.UseVisualStyleBackColor = True
        '
        'tbSearch
        '
        Me.tbSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbSearch.Location = New System.Drawing.Point(78, 132)
        Me.tbSearch.Name = "tbSearch"
        Me.tbSearch.Size = New System.Drawing.Size(159, 21)
        Me.tbSearch.TabIndex = 5
        Me.tbSearch.Visible = False
        '
        'CATALOGUE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(254, 191)
        Me.Controls.Add(Me.tbSearch)
        Me.Controls.Add(Me.labelWaitSearch)
        Me.Controls.Add(Me.cbSearch)
        Me.Controls.Add(Me.labelWaitAddSheet)
        Me.Controls.Add(Me.labelWaitXsl)
        Me.Controls.Add(Me.labelWaitHtml)
        Me.Controls.Add(Me.panelInfo)
        Me.Controls.Add(Me.cbAddSheet)
        Me.Controls.Add(Me.cbHTMLCatalogue)
        Me.Controls.Add(Me.cbXLSCatalogue)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(260, 220)
        Me.MinimumSize = New System.Drawing.Size(260, 220)
        Me.Name = "CATALOGUE"
        Me.Text = "CATALOGUE"
        Me.panelInfo.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbXLSCatalogue As System.Windows.Forms.Button
    Friend WithEvents cbHTMLCatalogue As System.Windows.Forms.Button
    Friend WithEvents cbAddSheet As System.Windows.Forms.Button
    Friend WithEvents panelInfo As System.Windows.Forms.Panel
    Friend WithEvents labelInfo As System.Windows.Forms.Label
    Friend WithEvents labelWaitHtml As System.Windows.Forms.Label
    Friend WithEvents labelWaitXsl As System.Windows.Forms.Label
    Friend WithEvents labelWaitAddSheet As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents labelWaitSearch As System.Windows.Forms.Label
    Friend WithEvents cbSearch As System.Windows.Forms.Button
    Friend WithEvents tbSearch As System.Windows.Forms.TextBox

End Class

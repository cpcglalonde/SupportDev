<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class About
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(About))
        Me.lblLink = New System.Windows.Forms.LinkLabel()
        Me.lblBugs = New System.Windows.Forms.Label()
        Me.lvlVersion = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblLink
        '
        Me.lblLink.AutoSize = True
        Me.lblLink.Location = New System.Drawing.Point(121, 75)
        Me.lblLink.Name = "lblLink"
        Me.lblLink.Size = New System.Drawing.Size(149, 13)
        Me.lblLink.TabIndex = 0
        Me.lblLink.TabStop = True
        Me.lblLink.Text = "greglalonde@conservative.ca"
        '
        'lblBugs
        '
        Me.lblBugs.AutoSize = True
        Me.lblBugs.Location = New System.Drawing.Point(29, 75)
        Me.lblBugs.Name = "lblBugs"
        Me.lblBugs.Size = New System.Drawing.Size(71, 13)
        Me.lblBugs.TabIndex = 1
        Me.lblBugs.Text = "Email Bugs to"
        '
        'lvlVersion
        '
        Me.lvlVersion.AutoSize = True
        Me.lvlVersion.Location = New System.Drawing.Point(29, 43)
        Me.lvlVersion.Name = "lvlVersion"
        Me.lvlVersion.Size = New System.Drawing.Size(60, 13)
        Me.lvlVersion.TabIndex = 2
        Me.lvlVersion.Text = "Version 3.0"
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(80, 9)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(146, 13)
        Me.lblName.TabIndex = 3
        Me.lblName.Text = "CIMS / AD Account Simplifier"
        '
        'About
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(297, 111)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lvlVersion)
        Me.Controls.Add(Me.lblBugs)
        Me.Controls.Add(Me.lblLink)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "About"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblLink As System.Windows.Forms.LinkLabel
    Friend WithEvents lblBugs As System.Windows.Forms.Label
    Friend WithEvents lvlVersion As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
End Class

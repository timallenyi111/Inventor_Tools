<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenu))
        AssemblyCopyToolButton = New Button()
        SuspendLayout()
        ' 
        ' AssemblyCopyToolButton
        ' 
        AssemblyCopyToolButton.BackColor = SystemColors.ControlDark
        AssemblyCopyToolButton.FlatStyle = FlatStyle.Flat
        AssemblyCopyToolButton.ForeColor = SystemColors.ControlText
        AssemblyCopyToolButton.Location = New Point(200, 100)
        AssemblyCopyToolButton.Margin = New Padding(0)
        AssemblyCopyToolButton.Name = "AssemblyCopyToolButton"
        AssemblyCopyToolButton.Size = New Size(200, 50)
        AssemblyCopyToolButton.TabIndex = 0
        AssemblyCopyToolButton.Text = "Load Model"
        AssemblyCopyToolButton.UseVisualStyleBackColor = False
        ' 
        ' MainMenu
        ' 
        AutoScaleDimensions = New SizeF(10F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.LightGray
        ClientSize = New Size(578, 244)
        Controls.Add(AssemblyCopyToolButton)
        ForeColor = SystemColors.ControlDarkDark
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "MainMenu"
        Text = "MainMenu"
        ResumeLayout(False)
    End Sub

    Friend WithEvents AssemblyCopyToolButton As Button
End Class

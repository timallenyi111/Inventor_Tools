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
        AssemblyCopyToolButton = New Button()
        SuspendLayout()
        ' 
        ' AssemblyCopyToolButton
        ' 
        AssemblyCopyToolButton.Location = New Point(292, 187)
        AssemblyCopyToolButton.Name = "AssemblyCopyToolButton"
        AssemblyCopyToolButton.Size = New Size(232, 48)
        AssemblyCopyToolButton.TabIndex = 0
        AssemblyCopyToolButton.Text = "Assembly Copy Tool"
        AssemblyCopyToolButton.UseVisualStyleBackColor = True
        ' 
        ' MainMenu
        ' 
        AutoScaleDimensions = New SizeF(11F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(AssemblyCopyToolButton)
        Name = "MainMenu"
        Text = "MainMenu"
        ResumeLayout(False)
    End Sub

    Friend WithEvents AssemblyCopyToolButton As Button
End Class

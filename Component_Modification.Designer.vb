<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Component_Modification
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
        components = New ComponentModel.Container()
        Button_CompModAccept = New Button()
        TB_ModCompName = New TextBox()
        CB_DoNotCopy = New CheckBox()
        ToolTip1 = New ToolTip(components)
        SuspendLayout()
        ' 
        ' Button_CompModAccept
        ' 
        Button_CompModAccept.Dock = DockStyle.Bottom
        Button_CompModAccept.Location = New Point(0, 201)
        Button_CompModAccept.Margin = New Padding(4, 3, 4, 3)
        Button_CompModAccept.Name = "Button_CompModAccept"
        Button_CompModAccept.Size = New Size(578, 43)
        Button_CompModAccept.TabIndex = 0
        Button_CompModAccept.Text = "Accept"
        Button_CompModAccept.UseVisualStyleBackColor = True
        ' 
        ' TB_ModCompName
        ' 
        TB_ModCompName.Location = New Point(50, 50)
        TB_ModCompName.Name = "TB_ModCompName"
        TB_ModCompName.Size = New Size(500, 37)
        TB_ModCompName.TabIndex = 1
        ' 
        ' CB_DoNotCopy
        ' 
        CB_DoNotCopy.AutoSize = True
        CB_DoNotCopy.Location = New Point(215, 120)
        CB_DoNotCopy.Name = "CB_DoNotCopy"
        CB_DoNotCopy.Size = New Size(148, 33)
        CB_DoNotCopy.TabIndex = 2
        CB_DoNotCopy.Text = "Don't Copy"
        ToolTip1.SetToolTip(CB_DoNotCopy, "Don't copy or replace this component in the ""copied"" assembly." & vbCrLf & "The ""copied"" assembly will reference back to this component's" & vbCrLf & "in the current file location.")
        CB_DoNotCopy.UseVisualStyleBackColor = True
        ' 
        ' Component_Modification
        ' 
        AutoScaleDimensions = New SizeF(12F, 29F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(578, 244)
        Controls.Add(CB_DoNotCopy)
        Controls.Add(TB_ModCompName)
        Controls.Add(Button_CompModAccept)
        Font = New Font("Calibri", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        FormBorderStyle = FormBorderStyle.FixedSingle
        Margin = New Padding(4, 3, 4, 3)
        MaximizeBox = False
        MinimizeBox = False
        Name = "Component_Modification"
        Text = "Modify Component"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Button_CompModAccept As Button
    Friend WithEvents TB_ModCompName As TextBox
    Friend WithEvents CB_DoNotCopy As CheckBox
    Friend WithEvents ToolTip1 As ToolTip
End Class

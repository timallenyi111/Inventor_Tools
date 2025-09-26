<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AssemblyCopyToolForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Label1 = New Label()
        Label2 = New Label()
        ProjDirTB = New TextBox()
        FileNameTB = New TextBox()
        oComponent_TV = New TreeView()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(25, 26)
        Label1.Name = "Label1"
        Label1.Size = New Size(184, 29)
        Label1.TabIndex = 0
        Label1.Text = "Project Directory:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(90, 79)
        Label2.Name = "Label2"
        Label2.Size = New Size(118, 29)
        Label2.TabIndex = 1
        Label2.Text = "File Name:"
        ' 
        ' ProjDirTB
        ' 
        ProjDirTB.Location = New Point(210, 18)
        ProjDirTB.Margin = New Padding(2, 3, 2, 3)
        ProjDirTB.Name = "ProjDirTB"
        ProjDirTB.ReadOnly = True
        ProjDirTB.Size = New Size(987, 37)
        ProjDirTB.TabIndex = 4
        ' 
        ' FileNameTB
        ' 
        FileNameTB.Location = New Point(210, 71)
        FileNameTB.Margin = New Padding(2, 3, 2, 3)
        FileNameTB.Name = "FileNameTB"
        FileNameTB.ReadOnly = True
        FileNameTB.Size = New Size(987, 37)
        FileNameTB.TabIndex = 5
        ' 
        ' oComponent_TV
        ' 
        oComponent_TV.Location = New Point(104, 192)
        oComponent_TV.Name = "oComponent_TV"
        oComponent_TV.Size = New Size(372, 691)
        oComponent_TV.TabIndex = 6
        ' 
        ' AssemblyCopyToolForm
        ' 
        AutoScaleDimensions = New SizeF(12F, 29F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1238, 1291)
        Controls.Add(oComponent_TV)
        Controls.Add(FileNameTB)
        Controls.Add(ProjDirTB)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Font = New Font("Calibri", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Margin = New Padding(3, 4, 3, 4)
        Name = "AssemblyCopyToolForm"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents ProjDirTB As TextBox
    Friend WithEvents FileNameTB As TextBox
    Friend WithEvents oComponent_TV As TreeView

End Class

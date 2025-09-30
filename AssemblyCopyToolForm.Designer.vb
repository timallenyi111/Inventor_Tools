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
        nComponent_TV = New TreeView()
        copyButton = New Button()
        Label3 = New Label()
        newDirTB = New TextBox()
        NewDirectoryFolderBrowser = New FolderBrowserDialog()
        newDirButton = New Button()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(26, 25)
        Label1.Name = "Label1"
        Label1.Size = New Size(184, 29)
        Label1.TabIndex = 20
        Label1.Text = "Project Directory:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(90, 80)
        Label2.Name = "Label2"
        Label2.Size = New Size(118, 29)
        Label2.TabIndex = 19
        Label2.Text = "File Name:"
        ' 
        ' ProjDirTB
        ' 
        ProjDirTB.Location = New Point(210, 18)
        ProjDirTB.Margin = New Padding(2, 3, 2, 3)
        ProjDirTB.Name = "ProjDirTB"
        ProjDirTB.ReadOnly = True
        ProjDirTB.Size = New Size(987, 37)
        ProjDirTB.TabIndex = 13
        ProjDirTB.TabStop = False
        ' 
        ' FileNameTB
        ' 
        FileNameTB.Location = New Point(210, 71)
        FileNameTB.Margin = New Padding(2, 3, 2, 3)
        FileNameTB.Name = "FileNameTB"
        FileNameTB.ReadOnly = True
        FileNameTB.Size = New Size(987, 37)
        FileNameTB.TabIndex = 12
        FileNameTB.TabStop = False
        FileNameTB.Tag = "t"
        ' 
        ' oComponent_TV
        ' 
        oComponent_TV.BorderStyle = BorderStyle.FixedSingle
        oComponent_TV.Location = New Point(25, 150)
        oComponent_TV.Name = "oComponent_TV"
        oComponent_TV.Size = New Size(662, 700)
        oComponent_TV.TabIndex = 11
        oComponent_TV.TabStop = False
        ' 
        ' nComponent_TV
        ' 
        nComponent_TV.BorderStyle = BorderStyle.FixedSingle
        nComponent_TV.Location = New Point(712, 150)
        nComponent_TV.Name = "nComponent_TV"
        nComponent_TV.Size = New Size(662, 700)
        nComponent_TV.TabIndex = 10
        nComponent_TV.TabStop = False
        ' 
        ' copyButton
        ' 
        copyButton.Font = New Font("Calibri", 14.0F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        copyButton.Location = New Point(615, 1300)
        copyButton.Name = "copyButton"
        copyButton.Size = New Size(170, 70)
        copyButton.TabIndex = 8
        copyButton.Text = "Copy"
        copyButton.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(50, 875)
        Label3.Name = "Label3"
        Label3.Size = New Size(160, 29)
        Label3.TabIndex = 18
        Label3.Text = "New Directory:"
        ' 
        ' newDirTB
        ' 
        newDirTB.Location = New Point(215, 875)
        newDirTB.Margin = New Padding(2, 3, 2, 3)
        newDirTB.Name = "newDirTB"
        newDirTB.Size = New Size(1031, 37)
        newDirTB.TabIndex = 17
        newDirTB.TabStop = False
        newDirTB.Tag = "t"
        ' 
        ' newDirButton
        ' 
        newDirButton.Location = New Point(1262, 878)
        newDirButton.Name = "newDirButton"
        newDirButton.Size = New Size(112, 34)
        newDirButton.TabIndex = 21
        newDirButton.Text = "Browse"
        newDirButton.UseVisualStyleBackColor = True
        ' 
        ' AssemblyCopyToolForm
        ' 
        AutoScaleDimensions = New SizeF(12.0F, 29.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1400, 1400)
        Controls.Add(newDirButton)
        Controls.Add(newDirTB)
        Controls.Add(Label3)
        Controls.Add(copyButton)
        Controls.Add(nComponent_TV)
        Controls.Add(oComponent_TV)
        Controls.Add(FileNameTB)
        Controls.Add(ProjDirTB)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Font = New Font("Calibri", 12.0F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
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
    Friend WithEvents nComponent_TV As TreeView
    Friend WithEvents copyButton As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents newDirTB As TextBox
    Friend WithEvents NewDirectoryFolderBrowser As FolderBrowserDialog
    Friend WithEvents newDirButton As Button

End Class

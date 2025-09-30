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
        TB_ProjDir = New TextBox()
        TB_FileName = New TextBox()
        TV_oComponent = New TreeView()
        TV_nComponent = New TreeView()
        CopyButton = New Button()
        Label3 = New Label()
        TB_newDir = New TextBox()
        NewDirectoryFolderBrowser = New FolderBrowserDialog()
        newDirButton = New Button()
        Label4 = New Label()
        TB_NewAssemblyName = New TextBox()
        TB_Suffix = New TextBox()
        TB_Prefix = New TextBox()
        Label_NewAssmName = New Label()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(66, 26)
        Label1.Name = "Label1"
        Label1.Size = New Size(184, 29)
        Label1.TabIndex = 20
        Label1.Text = "Project Directory:"
        Label1.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(132, 79)
        Label2.Name = "Label2"
        Label2.Size = New Size(118, 29)
        Label2.TabIndex = 19
        Label2.Text = "File Name:"
        Label2.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' TB_ProjDir
        ' 
        TB_ProjDir.Location = New Point(263, 18)
        TB_ProjDir.Margin = New Padding(2, 3, 2, 3)
        TB_ProjDir.Name = "TB_ProjDir"
        TB_ProjDir.ReadOnly = True
        TB_ProjDir.Size = New Size(934, 37)
        TB_ProjDir.TabIndex = 13
        TB_ProjDir.TabStop = False
        ' 
        ' TB_FileName
        ' 
        TB_FileName.Location = New Point(263, 71)
        TB_FileName.Margin = New Padding(2, 3, 2, 3)
        TB_FileName.Name = "TB_FileName"
        TB_FileName.ReadOnly = True
        TB_FileName.Size = New Size(934, 37)
        TB_FileName.TabIndex = 12
        TB_FileName.TabStop = False
        TB_FileName.Tag = "t"
        ' 
        ' TV_oComponent
        ' 
        TV_oComponent.BorderStyle = BorderStyle.FixedSingle
        TV_oComponent.Location = New Point(25, 150)
        TV_oComponent.Name = "TV_oComponent"
        TV_oComponent.Size = New Size(662, 700)
        TV_oComponent.TabIndex = 11
        TV_oComponent.TabStop = False
        ' 
        ' TV_nComponent
        ' 
        TV_nComponent.BorderStyle = BorderStyle.FixedSingle
        TV_nComponent.Location = New Point(712, 150)
        TV_nComponent.Name = "TV_nComponent"
        TV_nComponent.Size = New Size(662, 700)
        TV_nComponent.TabIndex = 10
        TV_nComponent.TabStop = False
        ' 
        ' CopyButton
        ' 
        CopyButton.Font = New Font("Calibri", 14.0F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        CopyButton.Location = New Point(615, 1300)
        CopyButton.Name = "CopyButton"
        CopyButton.Size = New Size(170, 70)
        CopyButton.TabIndex = 8
        CopyButton.Text = "Copy"
        CopyButton.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(90, 875)
        Label3.Name = "Label3"
        Label3.Size = New Size(160, 29)
        Label3.TabIndex = 18
        Label3.Text = "New Directory:"
        Label3.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' TB_newDir
        ' 
        TB_newDir.Location = New Point(247, 867)
        TB_newDir.Margin = New Padding(2, 3, 2, 3)
        TB_newDir.Name = "TB_newDir"
        TB_newDir.Size = New Size(983, 37)
        TB_newDir.TabIndex = 1
        TB_newDir.Tag = "t"
        ' 
        ' newDirButton
        ' 
        newDirButton.Location = New Point(1235, 872)
        newDirButton.Name = "newDirButton"
        newDirButton.Size = New Size(112, 34)
        newDirButton.TabIndex = 4
        newDirButton.TabStop = False
        newDirButton.Text = "Browse"
        newDirButton.UseVisualStyleBackColor = True
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(25, 926)
        Label4.Name = "Label4"
        Label4.Size = New Size(231, 29)
        Label4.TabIndex = 3
        Label4.Text = "New Assembly Name: "
        Label4.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' TB_NewAssemblyName
        ' 
        TB_NewAssemblyName.Location = New Point(391, 918)
        TB_NewAssemblyName.Margin = New Padding(2, 3, 2, 3)
        TB_NewAssemblyName.Name = "TB_NewAssemblyName"
        TB_NewAssemblyName.Size = New Size(440, 37)
        TB_NewAssemblyName.TabIndex = 2
        TB_NewAssemblyName.Tag = "t"
        ' 
        ' TB_Suffix
        ' 
        TB_Suffix.Location = New Point(835, 918)
        TB_Suffix.Margin = New Padding(2, 3, 2, 3)
        TB_Suffix.Name = "TB_Suffix"
        TB_Suffix.Size = New Size(131, 37)
        TB_Suffix.TabIndex = 21
        TB_Suffix.Tag = "t"
        ' 
        ' TB_Prefix
        ' 
        TB_Prefix.Location = New Point(247, 918)
        TB_Prefix.Margin = New Padding(2, 3, 2, 3)
        TB_Prefix.Name = "TB_Prefix"
        TB_Prefix.Size = New Size(140, 37)
        TB_Prefix.TabIndex = 22
        TB_Prefix.Tag = "t"
        TB_Prefix.TextAlign = HorizontalAlignment.Right
        ' 
        ' Label_NewAssmName
        ' 
        Label_NewAssmName.AutoSize = True
        Label_NewAssmName.Location = New Point(971, 921)
        Label_NewAssmName.Name = "Label_NewAssmName"
        Label_NewAssmName.Size = New Size(225, 29)
        Label_NewAssmName.TabIndex = 23
        Label_NewAssmName.Text = "  = 1_Default Name_2"
        Label_NewAssmName.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' AssemblyCopyToolForm
        ' 
        AutoScaleDimensions = New SizeF(12.0F, 29.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1400, 1400)
        Controls.Add(Label_NewAssmName)
        Controls.Add(TB_Prefix)
        Controls.Add(TB_Suffix)
        Controls.Add(TB_NewAssemblyName)
        Controls.Add(Label4)
        Controls.Add(newDirButton)
        Controls.Add(TB_newDir)
        Controls.Add(Label3)
        Controls.Add(CopyButton)
        Controls.Add(TV_nComponent)
        Controls.Add(TV_oComponent)
        Controls.Add(TB_FileName)
        Controls.Add(TB_ProjDir)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Font = New Font("Calibri", 12.0F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Margin = New Padding(3, 4, 3, 4)
        Name = "AssemblyCopyToolForm"
        Text = "Assembly Copy Tool"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TB_ProjDir As TextBox
    Friend WithEvents TB_FileName As TextBox
    Friend WithEvents TV_oComponent As TreeView
    Friend WithEvents TV_nComponent As TreeView
    Friend WithEvents CopyButton As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents TB_newDir As TextBox
    Friend WithEvents NewDirectoryFolderBrowser As FolderBrowserDialog
    Friend WithEvents newDirButton As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents TB_NewAssemblyName As TextBox
    Friend WithEvents TB_Suffix As TextBox
    Friend WithEvents TB_Prefix As TextBox
    Friend WithEvents Label_NewAssmName As Label

End Class

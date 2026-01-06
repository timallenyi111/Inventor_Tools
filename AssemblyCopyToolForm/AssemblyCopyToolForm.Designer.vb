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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssemblyCopyToolForm))
        LB_ProjectDirectory = New Label()
        LB_FileName = New Label()
        TB_ProjDir = New TextBox()
        TB_FileName = New TextBox()
        TV_nComponent = New TreeView()
        CopyButton = New Button()
        LB_NewDirectory = New Label()
        TB_newDir = New TextBox()
        NewDirectoryFolderBrowser = New FolderBrowserDialog()
        newDirButton = New Button()
        LB_Prefix = New Label()
        TB_Suffix = New TextBox()
        TB_Prefix = New TextBox()
        TestButton = New Button()
        LB_CopyComplete = New Label()
        LB_Suffix = New Label()
        GB_PreSuffix = New GroupBox()
        BT_PreSuffix = New Button()
        LB_TestLabel = New Label()
        GB_PreSuffix.SuspendLayout()
        SuspendLayout()
        ' 
        ' LB_ProjectDirectory
        ' 
        LB_ProjectDirectory.AutoSize = True
        LB_ProjectDirectory.ForeColor = SystemColors.ControlText
        LB_ProjectDirectory.Location = New Point(66, 26)
        LB_ProjectDirectory.Name = "LB_ProjectDirectory"
        LB_ProjectDirectory.Size = New Size(184, 29)
        LB_ProjectDirectory.TabIndex = 20
        LB_ProjectDirectory.Text = "Project Directory:"
        LB_ProjectDirectory.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' LB_FileName
        ' 
        LB_FileName.AutoSize = True
        LB_FileName.ForeColor = SystemColors.ControlText
        LB_FileName.Location = New Point(132, 79)
        LB_FileName.Name = "LB_FileName"
        LB_FileName.Size = New Size(118, 29)
        LB_FileName.TabIndex = 19
        LB_FileName.Text = "File Name:"
        LB_FileName.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' TB_ProjDir
        ' 
        TB_ProjDir.BackColor = Color.Silver
        TB_ProjDir.BorderStyle = BorderStyle.FixedSingle
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
        TB_FileName.BackColor = Color.Silver
        TB_FileName.BorderStyle = BorderStyle.FixedSingle
        TB_FileName.Location = New Point(263, 71)
        TB_FileName.Margin = New Padding(2, 3, 2, 3)
        TB_FileName.Name = "TB_FileName"
        TB_FileName.ReadOnly = True
        TB_FileName.Size = New Size(934, 37)
        TB_FileName.TabIndex = 12
        TB_FileName.TabStop = False
        TB_FileName.Tag = "t"
        ' 
        ' TV_nComponent
        ' 
        TV_nComponent.BackColor = Color.Silver
        TV_nComponent.BorderStyle = BorderStyle.FixedSingle
        TV_nComponent.FullRowSelect = True
        TV_nComponent.HideSelection = False
        TV_nComponent.LabelEdit = True
        TV_nComponent.Location = New Point(27, 195)
        TV_nComponent.Name = "TV_nComponent"
        TV_nComponent.Size = New Size(800, 800)
        TV_nComponent.TabIndex = 10
        TV_nComponent.TabStop = False
        ' 
        ' CopyButton
        ' 
        CopyButton.BackColor = SystemColors.ControlDark
        CopyButton.FlatStyle = FlatStyle.Flat
        CopyButton.Font = New Font("Calibri", 14F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        CopyButton.ForeColor = SystemColors.ControlText
        CopyButton.Location = New Point(458, 1043)
        CopyButton.Name = "CopyButton"
        CopyButton.Size = New Size(170, 70)
        CopyButton.TabIndex = 8
        CopyButton.Text = "Copy"
        CopyButton.UseVisualStyleBackColor = False
        ' 
        ' LB_NewDirectory
        ' 
        LB_NewDirectory.AutoSize = True
        LB_NewDirectory.ForeColor = SystemColors.ControlText
        LB_NewDirectory.Location = New Point(98, 135)
        LB_NewDirectory.Name = "LB_NewDirectory"
        LB_NewDirectory.Size = New Size(160, 29)
        LB_NewDirectory.TabIndex = 18
        LB_NewDirectory.Text = "New Directory:"
        LB_NewDirectory.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' TB_newDir
        ' 
        TB_newDir.BackColor = Color.Silver
        TB_newDir.BorderStyle = BorderStyle.FixedSingle
        TB_newDir.Location = New Point(263, 132)
        TB_newDir.Margin = New Padding(2, 3, 2, 3)
        TB_newDir.Name = "TB_newDir"
        TB_newDir.Size = New Size(644, 37)
        TB_newDir.TabIndex = 1
        TB_newDir.Tag = "t"
        ' 
        ' newDirButton
        ' 
        newDirButton.BackColor = SystemColors.ControlDark
        newDirButton.FlatStyle = FlatStyle.Flat
        newDirButton.Font = New Font("Calibri", 12F, FontStyle.Bold)
        newDirButton.ForeColor = SystemColors.ControlText
        newDirButton.Location = New Point(1038, 135)
        newDirButton.Name = "newDirButton"
        newDirButton.Size = New Size(112, 45)
        newDirButton.TabIndex = 4
        newDirButton.TabStop = False
        newDirButton.Text = "Browse"
        newDirButton.UseVisualStyleBackColor = False
        ' 
        ' LB_Prefix
        ' 
        LB_Prefix.AutoSize = True
        LB_Prefix.ForeColor = SystemColors.ControlText
        LB_Prefix.Location = New Point(-3, 35)
        LB_Prefix.Name = "LB_Prefix"
        LB_Prefix.Size = New Size(74, 29)
        LB_Prefix.TabIndex = 3
        LB_Prefix.Text = "Prefix:"
        LB_Prefix.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' TB_Suffix
        ' 
        TB_Suffix.BackColor = Color.Silver
        TB_Suffix.BorderStyle = BorderStyle.FixedSingle
        TB_Suffix.Location = New Point(104, 107)
        TB_Suffix.Margin = New Padding(2, 3, 2, 3)
        TB_Suffix.Name = "TB_Suffix"
        TB_Suffix.Size = New Size(131, 37)
        TB_Suffix.TabIndex = 21
        TB_Suffix.Tag = "t"
        ' 
        ' TB_Prefix
        ' 
        TB_Prefix.BackColor = Color.Silver
        TB_Prefix.BorderStyle = BorderStyle.FixedSingle
        TB_Prefix.Location = New Point(83, 35)
        TB_Prefix.Margin = New Padding(2, 3, 2, 3)
        TB_Prefix.Name = "TB_Prefix"
        TB_Prefix.Size = New Size(165, 37)
        TB_Prefix.TabIndex = 22
        TB_Prefix.Tag = ""
        ' 
        ' TestButton
        ' 
        TestButton.BackColor = SystemColors.ControlDark
        TestButton.FlatStyle = FlatStyle.Flat
        TestButton.Font = New Font("Calibri", 14F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        TestButton.ForeColor = SystemColors.ControlText
        TestButton.Location = New Point(662, 1043)
        TestButton.Name = "TestButton"
        TestButton.Size = New Size(80, 70)
        TestButton.TabIndex = 24
        TestButton.Text = "Test"
        TestButton.UseVisualStyleBackColor = False
        ' 
        ' LB_CopyComplete
        ' 
        LB_CopyComplete.AutoSize = True
        LB_CopyComplete.Font = New Font("Calibri", 8F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LB_CopyComplete.Location = New Point(27, 998)
        LB_CopyComplete.Name = "LB_CopyComplete"
        LB_CopyComplete.Size = New Size(121, 19)
        LB_CopyComplete.TabIndex = 25
        LB_CopyComplete.Text = "COPY COMPLETE"
        LB_CopyComplete.Visible = False
        ' 
        ' LB_Suffix
        ' 
        LB_Suffix.AutoSize = True
        LB_Suffix.ForeColor = SystemColors.ControlText
        LB_Suffix.Location = New Point(3, 104)
        LB_Suffix.Name = "LB_Suffix"
        LB_Suffix.Size = New Size(72, 29)
        LB_Suffix.TabIndex = 26
        LB_Suffix.Text = "Suffix:"
        LB_Suffix.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' GB_PreSuffix
        ' 
        GB_PreSuffix.Controls.Add(BT_PreSuffix)
        GB_PreSuffix.Controls.Add(TB_Prefix)
        GB_PreSuffix.Controls.Add(LB_Suffix)
        GB_PreSuffix.Controls.Add(TB_Suffix)
        GB_PreSuffix.Controls.Add(LB_Prefix)
        GB_PreSuffix.FlatStyle = FlatStyle.System
        GB_PreSuffix.ForeColor = SystemColors.ControlDarkDark
        GB_PreSuffix.Location = New Point(843, 195)
        GB_PreSuffix.Margin = New Padding(0)
        GB_PreSuffix.Name = "GB_PreSuffix"
        GB_PreSuffix.Padding = New Padding(0)
        GB_PreSuffix.Size = New Size(323, 253)
        GB_PreSuffix.TabIndex = 27
        GB_PreSuffix.TabStop = False
        ' 
        ' BT_PreSuffix
        ' 
        BT_PreSuffix.BackColor = SystemColors.ControlDark
        BT_PreSuffix.FlatStyle = FlatStyle.Flat
        BT_PreSuffix.Font = New Font("Calibri", 12F, FontStyle.Bold)
        BT_PreSuffix.ForeColor = SystemColors.ControlText
        BT_PreSuffix.Location = New Point(104, 176)
        BT_PreSuffix.Margin = New Padding(0)
        BT_PreSuffix.Name = "BT_PreSuffix"
        BT_PreSuffix.Size = New Size(100, 45)
        BT_PreSuffix.TabIndex = 27
        BT_PreSuffix.Text = "Apply"
        BT_PreSuffix.UseVisualStyleBackColor = False
        ' 
        ' LB_TestLabel
        ' 
        LB_TestLabel.AutoSize = True
        LB_TestLabel.ForeColor = SystemColors.ControlText
        LB_TestLabel.Location = New Point(870, 891)
        LB_TestLabel.Name = "LB_TestLabel"
        LB_TestLabel.Size = New Size(52, 29)
        LB_TestLabel.TabIndex = 28
        LB_TestLabel.Text = "Test"
        LB_TestLabel.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' AssemblyCopyToolForm
        ' 
        AutoScaleDimensions = New SizeF(12F, 29F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.LightGray
        ClientSize = New Size(1178, 1144)
        Controls.Add(LB_TestLabel)
        Controls.Add(GB_PreSuffix)
        Controls.Add(LB_CopyComplete)
        Controls.Add(TestButton)
        Controls.Add(newDirButton)
        Controls.Add(TB_newDir)
        Controls.Add(LB_NewDirectory)
        Controls.Add(CopyButton)
        Controls.Add(TV_nComponent)
        Controls.Add(TB_FileName)
        Controls.Add(TB_ProjDir)
        Controls.Add(LB_FileName)
        Controls.Add(LB_ProjectDirectory)
        Font = New Font("Calibri", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ForeColor = SystemColors.ControlDarkDark
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Margin = New Padding(3, 4, 3, 4)
        MaximizeBox = False
        Name = "AssemblyCopyToolForm"
        SizeGripStyle = SizeGripStyle.Hide
        Text = "Assembly Copy Tool"
        GB_PreSuffix.ResumeLayout(False)
        GB_PreSuffix.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents LB_ProjectDirectory As Label
    Friend WithEvents LB_FileName As Label
    Friend WithEvents TB_ProjDir As TextBox
    Friend WithEvents TB_FileName As TextBox
    Friend WithEvents TV_nComponent As TreeView
    Friend WithEvents CopyButton As Button
    Friend WithEvents LB_NewDirectory As Label
    Friend WithEvents TB_newDir As TextBox
    Friend WithEvents NewDirectoryFolderBrowser As FolderBrowserDialog
    Friend WithEvents newDirButton As Button
    Friend WithEvents LB_Prefix As Label
    Friend WithEvents TB_Suffix As TextBox
    Friend WithEvents TB_Prefix As TextBox
    Friend WithEvents TestButton As Button
    Friend WithEvents LB_CopyComplete As Label
    Friend WithEvents LB_Suffix As Label
    Friend WithEvents GB_PreSuffix As GroupBox
    Friend WithEvents BT_PreSuffix As Button
    Friend WithEvents LB_TestLabel As Label

End Class

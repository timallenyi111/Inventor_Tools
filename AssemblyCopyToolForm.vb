Imports System
Imports System.Activator
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Type
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Inventor


Friend Class AssemblyCopyToolForm
    Inherits System.Windows.Forms.Form
    Public _invApp As Inventor.Application
    Public oAsmDoc As Inventor.AssemblyDocument
    Dim newRootDirectory As String = ""
    Dim mainAsmObj As New InvtAssemblyObj
    Dim projectDir As String
    Dim oAsmFileName As String
    Dim oAsmCompDef As AssemblyComponentDefinition
    Dim newDirectory As String

    Dim defaultSuffix As String = "_2"
    Dim medium_gap As Integer = 10
    Dim labelRightEdge As Integer = 164

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next

        'get an active session of Inventor
        _invApp = Win32.GetActiveObject("Inventor.Application")
        If Err.Number Then
            MsgBox("Inventor must be running.")
        End If

        On Error GoTo 0

        oAsmDoc = _invApp.ActiveDocument
        If Err.Number Then
            MsgBox("A document must be open in Inventor.")
            End
        End If

        On Error GoTo 0

        oAsmFileName = GetAssemblyFileName(oAsmDoc)
        TB_FileName.Text = oAsmFileName

        ' by default the new assembly body is the original assembly name
        TB_NewAssemblyName.Text = GetComponentName(oAsmDoc)
        ' put the default suffix in the TB and then make the new assembly name
        TB_Suffix.Text = defaultSuffix
        Dim newAssemblyName As String = AddPrefixSuffix(TB_NewAssemblyName.Text, TB_Prefix.Text, TB_Suffix.Text)
        Label_NewAssmName.Text = "  :  " & newAssemblyName

        'setup the initial new directory and filename based on the current project directory
        'and default suffix
        projectDir = GetProjectDirectory(_invApp)
        newRootDirectory = projectDir & newAssemblyName & "\"
        LongTextboxWrite(TB_ProjDir, projectDir)
        LongTextboxWrite(TB_newDir, newRootDirectory)

        mainAsmObj.NewFilePath = newRootDirectory
        mainAsmObj.NewName = newAssemblyName

        'setup the form layout after assigning values
        FormLayoutSetup(True)

        oAsmCompDef = oAsmDoc.ComponentDefinition


        mainAsmObj = AssemblyObjSetup(oAsmDoc, mainAsmObj)

        ' Dim oTreeView As TreeView
        Dim oTreeView = TV_oComponent
        SetupTreeView(oTreeView, mainAsmObj)

        Dim newAsmObj As New InvtAssemblyObj
        newAsmObj = AssemblyObjSetup(oAsmDoc, newAsmObj)

        Dim nTreeView = TV_nComponent
        SetupTreeView(nTreeView, newAsmObj)


    End Sub

    Private Sub AssemblyCopyToolForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        FormLayoutSetup(False)
        ResetCarets()
    End Sub

    Private Sub FormLayoutSetup(ByRef initialLayout As Boolean)
        Dim clientWidth As Integer = Me.ClientSize.Width
        Dim clientHeight As Integer = Me.ClientSize.Height
        'Dim medium_gap = clientWidth * 0.01
        Dim aboveTV_space As Integer = 100
        Dim belowTV_space As Integer = 370
        Dim standardHeight As Integer = 25

        'stuff you don't want to scale with resize
        If initialLayout Then
            DB("Client Width: " & clientWidth.ToString)
            DB("Client Height: " & clientHeight.ToString)

            Label1.Height = standardHeight
            Label1.Left = CInt(labelRightEdge - Label1.Width)
            Label1.Top = 17

            Dim label1_2Spacing As Integer = 25
            Label2.Height = standardHeight
            Label2.Left = CInt(labelRightEdge - Label2.Width)
            Label2.Top = CInt(Label1.Top + Label1.Height + label1_2Spacing)

            TB_ProjDir.Height = standardHeight
            TB_ProjDir.Left = labelRightEdge
            TB_ProjDir.Top = Label1.Top

            TB_FileName.Height = standardHeight
            TB_FileName.Left = labelRightEdge
            TB_FileName.Top = Label2.Top

            Label3.Height = standardHeight
            Label3.Left = labelRightEdge - Label3.Width

            newDirButton.Width = 70
            newDirButton.Height = TB_newDir.Height

            TB_newDir.Left = labelRightEdge
            TB_newDir.Height = standardHeight

            Label4.Left = labelRightEdge - Label4.Width
            Label4.Height = standardHeight

            TB_Prefix.Left = labelRightEdge
            TB_Prefix.Height = standardHeight

            TB_NewAssemblyName.Height = standardHeight

            TB_Suffix.Height = standardHeight

        End If

        TB_ProjDir.Width = CInt(clientWidth - TB_ProjDir.Left - medium_gap)
        TB_FileName.Width = CInt(clientWidth - TB_FileName.Left - medium_gap)

        Dim tv_width = (clientWidth - (medium_gap * 3)) / 2
        TV_oComponent.Width = CInt(tv_width)
        TV_nComponent.Width = CInt(tv_width)

        Dim tv_height As Integer = Me.ClientSize.Height - aboveTV_space - belowTV_space
        TV_oComponent.Height = CInt(tv_height)
        TV_nComponent.Height = CInt(tv_height)

        TV_oComponent.Left = CInt(medium_gap)
        TV_nComponent.Left = CInt(tv_width + medium_gap * 2)

        CopyButton.Left = CInt(clientWidth / 2) - CopyButton.Width / 2
        CopyButton.Top = CInt(clientHeight / 14) * 13

        Dim newDirLabelTop As Integer = TV_oComponent.Top + TV_oComponent.Height + medium_gap * 1.5
        Label3.Top = newDirLabelTop

        newDirButton.Top = newDirLabelTop
        newDirButton.Left = medium_gap * 2 + tv_width * 2 - newDirButton.Width

        TB_newDir.Top = newDirLabelTop
        TB_newDir.Width = CInt(clientWidth - TB_newDir.Left - medium_gap - newDirButton.Width)

        Label4.Top = newDirLabelTop + TB_newDir.Height + medium_gap

        ResizeAssemblyNameLayout()
        ResetCarets()

    End Sub

    Private Sub NewDirButton_Click(sender As Object, e As EventArgs) Handles newDirButton.Click
        Using NewDirectoryFolderBrowser As New FolderBrowserDialog()
            NewDirectoryFolderBrowser.SelectedPath = newDirectory
            If NewDirectoryFolderBrowser.ShowDialog() = DialogResult.OK Then
                newDirectory = NewDirectoryFolderBrowser.SelectedPath & "\"
                TB_newDir.Text = newDirectory
            End If
        End Using
    End Sub

    Private Sub ChangeNewAssemblyNameLabel(ByRef assmName As String)
        Label_NewAssmName.Text = "  :  " & assmName
        ResizeAssemblyNameLayout()
    End Sub

    Private Sub ResizeAssemblyNameLayout()
        Dim clientWidth As Integer = Me.ClientSize.Width
        Dim newAssemblyNameArea As Integer = clientWidth - labelRightEdge - Label_NewAssmName.Width - medium_gap

        TB_Prefix.Top = Label4.Top
        TB_Prefix.Left = labelRightEdge
        TB_Prefix.Width = newAssemblyNameArea * 0.1

        TB_NewAssemblyName.Top = Label4.Top
        TB_NewAssemblyName.Left = TB_Prefix.Left + TB_Prefix.Width
        TB_NewAssemblyName.Width = newAssemblyNameArea * 0.8

        TB_Suffix.Top = Label4.Top
        TB_Suffix.Left = TB_NewAssemblyName.Left + TB_NewAssemblyName.Width
        TB_Suffix.Width = newAssemblyNameArea * 0.1

        Label_NewAssmName.Left = clientWidth - medium_gap - Label_NewAssmName.Width
    End Sub

    Private Sub PrefixTB_TextChanged(sender As Object, e As EventArgs) Handles TB_Prefix.TextChanged
        UpdateNewAssemblyName()
    End Sub

    Private Sub TB_NewAssemblyName_TextChanged(sender As Object, e As EventArgs) Handles TB_NewAssemblyName.TextChanged
        UpdateNewAssemblyName()
    End Sub

    Private Sub TB_Suffix_TextChanged(sender As Object, e As EventArgs) Handles TB_Suffix.TextChanged
        UpdateNewAssemblyName()
    End Sub

    Private Sub UpdateNewAssemblyName()
        If newRootDirectory = "" Then
            'this happens on initial load
        Else
            Dim newName As String = TB_Prefix.Text & TB_NewAssemblyName.Text & TB_Suffix.Text
            Dim dirString As String = newRootDirectory
            'you have to do it twice to get rid of the old file name
            dirString = dirString.Substring(0, dirString.LastIndexOf("\"))
            dirString = dirString.Substring(0, dirString.LastIndexOf("\") + 1) & newName & "\"
            newRootDirectory = dirString
            TB_newDir.Text = newRootDirectory
            mainAsmObj.NewName = newName
            mainAsmObj.NewFilePath = newRootDirectory
            ResetCarets()
        End If

    End Sub

    Private Sub CopyButton_Click(sender As Object, e As EventArgs) Handles CopyButton.Click
        SaveAssemblyFile(mainAsmObj)
    End Sub

    Sub SetupTreeView(ByRef treeView As System.Windows.Forms.TreeView, ByRef asmObj As InvtAssemblyObj)
        treeView.Nodes.Clear()
        Dim rootNode As TreeNode = treeView.Nodes.Add(asmObj.OriginalName)
        AddSubNodes(rootNode, asmObj)
        treeView.ExpandAll()
    End Sub

    Sub AddSubNodes(ByRef parentNode As TreeNode, ByRef asmObj As InvtAssemblyObj)
        Dim newNode As TreeNode
        For Each comp As InvtComponentObj In asmObj.AssemblyComponents
            newNode = parentNode.Nodes.Add(comp.Name & " | qty: " & comp.Quantity.ToString)
            If comp.Type = "Assembly" Then
                AddSubNodes(newNode, comp.AssemblyObject)
            End If
        Next
    End Sub

    Private Sub ResetCarets()
        MoveCaret(TB_ProjDir)
        MoveCaret(TB_newDir)
    End Sub


End Class

Public Class Win32
	<DllImport("ole32")>
	Private Shared Function CLSIDFromProgIDEx(
		<MarshalAs(UnmanagedType.LPWStr)> ByVal lpszProgID As String, <Out> ByRef lpclsid As Guid) As Integer
	End Function
	<DllImport("ole32")>
	Private Shared Function CLSIDFromProgID(
		<MarshalAs(UnmanagedType.LPWStr)> ByVal lpszProgID As String, <Out> ByRef lpclsid As Guid) As Integer
	End Function
	<DllImport("oleaut32")>
	Private Shared Function GetActiveObject(
		<MarshalAs(UnmanagedType.LPStruct)> ByVal rclsid As Guid, ByVal pvReserved As IntPtr, <Out>
		<MarshalAs(UnmanagedType.IUnknown)> ByRef ppunk As Object) As Integer
	End Function

	Public Shared Function GetActiveObject(ByVal progID As String) As Object
		Dim obj As Object = Nothing
		Dim clsid As Guid

		' Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
		' CLSIDFromProgIDEx doesn't exist.
		Try
			CLSIDFromProgIDEx(progID, clsid)
		Catch ex As Exception
			CLSIDFromProgID(progID, clsid)
		End Try

		Dim hr = GetActiveObject(clsid, IntPtr.Zero, obj)
		If hr < 0 Then
			Err.Raise(0)
		End If
		Return obj
	End Function
End Class

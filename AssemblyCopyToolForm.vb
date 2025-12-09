Imports System
Imports System.Activator
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Type
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Inventor
Imports System.IO


Friend Class AssemblyCopyToolForm
    Inherits System.Windows.Forms.Form
    Public _invApp As Inventor.Application
    Public oAsmDoc As Inventor.AssemblyDocument
    Public _stream As FileStream
    Public _writer As StreamWriter
    Dim logPath As String = "C:\Users\Tim\source\repos\Inventor_Tools\LogFiles\"
    Dim EnableLog As Boolean = True
    Dim oAsmCompDef As AssemblyComponentDefinition
    Dim newDirectory As String
    Dim rootAssemblyObject As AssemblyCopyObject
    Dim defaultSuffix As String = "_2"
    Dim defaultPrefix As String = ""
    Dim medium_gap As Integer = 10
    Dim labelRightEdge As Integer = 164
    Dim highlightSet As Inventor.HighlightSet

    Private Sub AssemblyCopyFormLoad(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next

        'get an active session of Inventor
        _invApp = Win32.GetActiveObject("Inventor.Application")
        If Err.Number Then
            MsgBox("Inventor must be running.")
        End If

        On Error GoTo 0

        oAsmDoc = _invApp.ActiveDocument
        highlightSet = _invApp.ActiveDocument.CreateHighlightSet()
        If Err.Number Then
            MsgBox("A document must be open in Inventor.")
            End
        End If

        On Error GoTo 0

        ' setup Log file
        If EnableLog Then
            Dim timestamp As String = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
            Dim fileName As String = $"Log_{timestamp}.txt"
            _stream = New FileStream(logPath & fileName, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)
            _writer = New StreamWriter(_stream)
            _writer.AutoFlush = True
        End If


        ' we have to set the default prefix and suffix textboxes before assembly object setup
        ' because the assembly object setup references these values
        TB_Prefix.Text = defaultPrefix
        TB_Suffix.Text = defaultSuffix

        'create and setup the AssemblyCopyObject
        rootAssemblyObject = New AssemblyCopyObject(Me, _invApp)
        rootAssemblyObject.InitialSetup()
        If EnableLog Then
            rootAssemblyObject.GenerateSetupLog()
        End If


        TB_FileName.Text = rootAssemblyObject.OriginalName & ".iam"
        ' by default the new assembly name is the same as the original
        ' this is just the middle of the name not including prefix and suffix
        TB_NewAssemblyName.Text = rootAssemblyObject.OriginalName

        Label_NewAssmName.Text = "  :  " & rootAssemblyObject.NewName
        LongTextboxWrite(TB_ProjDir, rootAssemblyObject.GetProjectDirectory(_invApp))
        LongTextboxWrite(TB_newDir, rootAssemblyObject.NewRootDirectory)

        ' setup the form layout after assigning values
        FormLayoutSetup(True)

        ' setup tree views
        TV_oComponent.Nodes.Add(rootAssemblyObject.OriginalTreeNode)
        TV_nComponent.Nodes.Add(rootAssemblyObject.NewTreeNode)

    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        highlightSet.Clear()
        MyBase.OnFormClosing(e)
        _stream.Close()
        ' _writer.Close()
    End Sub

    Public Sub Log(message As String, Optional ByRef numTabs As Integer = 0, Optional ByRef numLines As Integer = 0)
        If EnableLog Then
            Dim x As Integer = 0
            While x < numTabs
                message = vbTab & message
                x += 1
            End While

            _writer.WriteLine(message)

            x = 0
            While x < numLines
                _writer.WriteLine("")
                x += 1
            End While
        End If

    End Sub





#Region "Form Text Controls"

    ''' <summary>
    ''' Updates the new assembly name label and then calls for the resizing of the UI in that row
    ''' </summary>
    ''' <param name="assmName"></param>
    Private Sub ChangeNewAssemblyNameLabel(ByRef assmName As String)
        Label_NewAssmName.Text = "  :  " & assmName
        ResizeAssemblyNameLayout()
    End Sub

    Private Sub PrefixTB_TextChanged(sender As Object, e As EventArgs) Handles TB_Prefix.TextChanged
        UpdateNewFileName()
    End Sub

    Private Sub TB_NewAssemblyName_TextChanged(sender As Object, e As EventArgs) Handles TB_NewAssemblyName.TextChanged
        UpdateNewFileName()
    End Sub

    Private Sub TB_Suffix_TextChanged(sender As Object, e As EventArgs) Handles TB_Suffix.TextChanged
        UpdateNewFileName()
    End Sub

    Private Sub UpdateNewFileName()

        If rootAssemblyObject Is Nothing Then
            ' this happens on the initial load
        Else
            Dim asmName As String = TB_Prefix.Text & TB_NewAssemblyName.Text & TB_Suffix.Text
            Label_NewAssmName.Text = "  :  " & asmName
            ResizeAssemblyNameLayout()
            rootAssemblyObject.NewName = asmName
            TB_newDir.Text = rootAssemblyObject.NewRootDirectory
            ResetCarets()
        End If


    End Sub

#End Region


#Region "Button Clicks"

    Private Sub CopyButton_Click(sender As Object, e As EventArgs) Handles CopyButton.Click
        rootAssemblyObject.UpdateNewProperties()
        rootAssemblyObject.CreateNewFiles(dryrun:=False)
        rootAssemblyObject.ReplaceOccurences()
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

    Private Sub TestButton_Click(sender As Object, e As EventArgs) Handles TestButton.Click
        rootAssemblyObject.AssignNodeTags()
    End Sub

#End Region

#Region "Form Control Functions"
    Private Sub FormLayoutSetup(ByRef initialLayout As Boolean)
        Dim clientWidth As Integer = Me.ClientSize.Width
        Dim clientHeight As Integer = Me.ClientSize.Height
        'Dim medium_gap = clientWidth * 0.01
        Dim aboveTV_space As Integer = 100
        Dim belowTV_space As Integer = 370
        Dim standardHeight As Integer = 25

        'stuff you don't want to scale with resize
        If initialLayout Then

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

        TestButton.Left = CopyButton.Left + CopyButton.Width + medium_gap
        TestButton.Top = CopyButton.Top

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
    Private Sub ResetCarets()
        MoveCaret(TB_ProjDir)
        MoveCaret(TB_newDir)
    End Sub

    Private Sub AssemblyCopyToolForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        FormLayoutSetup(False)
        ResetCarets()
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

    ''' <summary>
    ''' Writes text in a textbox and scrolls to the end
    ''' </summary>
    ''' <param name="textBox"></param>
    ''' <param name="msg"></param>
    Sub LongTextboxWrite(ByRef textBox As System.Windows.Forms.TextBox, ByRef msg As String)
        textBox.Text = msg
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub
    Sub MoveCaret(ByRef textBox As System.Windows.Forms.TextBox)
        textBox.SelectionStart = 0
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub
    

    ' When a node is clicked, select it and give the treeview focus so the highlight is visible
    Private Sub TV_nComponent_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TV_nComponent.NodeMouseClick
        highlightSet.Clear()
        TV_nComponent.SelectedNode = e.Node
        TV_nComponent.Focus()

        Dim Ena As Boolean = True

        If Ena Then
            If TV_nComponent.SelectedNode.Parent Is Nothing Then
                'this is a root assembly occurence

            ElseIf TV_nComponent.SelectedNode.Parent Is rootAssemblyObject.NewTreeNode Then
                'this is a component of the root assembly
                For Each occ As ComponentOccurrence In TV_nComponent.SelectedNode.Tag
                    highlightSet.AddItem(occ)
                Next
            Else
                For Each occProx As ComponentOccurrenceProxy In TV_nComponent.SelectedNode.Tag
                    highlightSet.AddItem(occProx)
                Next
            End If
        End If


        'rootAssemblyObject.HighlightOccurenceByNode(TV_nComponent.SelectedNode)
    End Sub

#End Region


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

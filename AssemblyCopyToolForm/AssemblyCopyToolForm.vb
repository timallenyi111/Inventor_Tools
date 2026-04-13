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
    Public doubleClickNode As TreeNode

    Private selectedNode As TreeNode
    Private oAsmCompDef As AssemblyComponentDefinition
    Private newDirectory As String
    'Private rootAssemblyObject As AssemblyCopyObject
    Private rootAssemblyObject As InvtAssembly
    Dim defaultSuffix As String = "_2"
    Dim defaultPrefix As String = ""
    Dim highlightSet As Inventor.HighlightSet


    '****************************
#Region "Program Settings"
    Dim logPath As String = "C:\Users\TimAllen\source\repos\timallenyi111\Inventor_Tools\LogFiles\"
    Dim EnableLog As Boolean = True
    Dim EnableNodeHighlighting As Boolean = False

#End Region

    Private Sub AssemblyCopyFormLoad(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next

        Dim sw As Stopwatch = Stopwatch.StartNew()

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

        'Debug.WriteLine(_invApp.DesignProjectManager.ActiveDesignProject.ContentCenterPath)
        Dim documents As Inventor.Documents = _invApp.Documents
        For Each doc As Inventor.Document In documents
            Debug.WriteLine(doc.FullFileName)
        Next

        'testing getting a component count for a loading bar
        Dim curDoc As AssemblyDocument = _invApp.ActiveDocument
        Dim componentCount As Integer = GetTotalNumberOfComponents(curDoc.ComponentDefinition.Occurrences)
        LB_TestLabel.Text = componentCount

        ' we have to set the default prefix and suffix textboxes before assembly object setup
        ' because the assembly object setup references these values
        TB_Prefix.Text = defaultPrefix
        TB_Suffix.Text = defaultSuffix

        'create the root assembly object
        rootAssemblyObject = InitialSetup(_invApp, Me)

        If rootAssemblyObject Is Nothing Then
            'there wasn't an assembly document open so the program will end after the user clicks ok on the message box
            Me.Close()
        End If

        TB_FileName.Text = rootAssemblyObject.OriginalName & ".iam"

        Dim actProj As Inventor.DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
        Dim projectDir As String = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\") + 1)
        LongTextboxWrite(TB_ProjDir, projectDir)
        LongTextboxWrite(TB_newDir, rootAssemblyObject.NewRootDirectory)

        ' setup the form layout after assigning values
        FormLayoutSetup(True)

        ' setup tree view
        TV_nComponent.Nodes.Add(rootAssemblyObject.TreeNode)



        sw.Stop()
        Debug.WriteLine("Form Load Time: " & sw.ElapsedMilliseconds & " ms")
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        highlightSet.Clear()
        MyBase.OnFormClosing(e)
        If EnableLog Then
            Log("Closing log file.")
            _stream.Close()
        End If

        ' _writer.Close()
    End Sub

    ''' <summary>
    ''' The Number of Lines is the number of lines that will be added after the message.
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="numTabs"></param>
    ''' <param name="numLinesAfter"></param>
    ''' <param name="debugWrite"></param>
    Public Sub Log(message As String, Optional ByRef numTabs As Integer = 0, Optional ByRef numLinesBefore As Integer = 0, Optional ByRef numLinesAfter As Integer = 0, Optional ByRef debugWrite As Boolean = True)
        Dim x As Integer = 0
        'Write to Debug console
        If debugWrite Then

            While x < numTabs
                message = vbTab & message
                x += 1
            End While

            x = 0
            While x < numLinesBefore
                Debug.WriteLine("")
                x += 1
            End While

            Debug.WriteLine(message)

            x = 0
            While x < numLinesAfter
                Debug.WriteLine("")
                x += 1
            End While
        End If

        'write to log file
        If EnableLog Then
            x = 0
            While x < numTabs
                message = vbTab & message
                x += 1
            End While
            x = 0
            While x < numLinesBefore
                _writer.WriteLine("")
                x += 1
            End While
            _writer.WriteLine(message)
            x = 0
            While x < numLinesAfter
                _writer.WriteLine("")
                x += 1
            End While
        End If

    End Sub

    Private Function GetTotalNumberOfComponents(ByRef asmOccs As ComponentOccurrences) As Integer
        Dim occCount As Integer = asmOccs.Count

        For Each occ As ComponentOccurrence In asmOccs
            If occ.Type = ObjectTypeEnum.kAssemblyComponentDefinitionObject Then
                Dim occDoc As AssemblyDocument = occ.Definition.Document
                occCount += GetTotalNumberOfComponents(occDoc.ComponentDefinition.Occurrences)
            End If
        Next

        Return occCount
    End Function

#Region "Form Text Controls"

    ''' <summary>
    ''' Updates the new assembly name label and then calls for the resizing of the UI in that row
    ''' </summary>
    ''' <param name="assmName"></param>
    'Private Sub ChangeNewAssemblyNameLabel(ByRef assmName As String)
    '    ResizeAssemblyNameLayout()
    'End Sub

    Private Sub PrefixTB_TextChanged(sender As Object, e As EventArgs) Handles TB_Prefix.TextChanged
        'UpdateNewFileName()
    End Sub

    Private Sub TB_NewAssemblyName_TextChanged(sender As Object, e As EventArgs)
        'UpdateNewFileName()
    End Sub

    Private Sub TB_Suffix_TextChanged(sender As Object, e As EventArgs) Handles TB_Suffix.TextChanged
        'UpdateNewFileName()
    End Sub

    Private Sub UpdateNewFileName(ByVal newAsmName As String)



        'If rootAssemblyObject Is Nothing Then
        '    ' this happens on the initial load
        'Else
        '    Dim asmName As String = TB_Prefix.Text & TB_NewAssemblyName.Text & TB_Suffix.Text
        '    ResizeAssemblyNameLayout()
        '    rootAssemblyObject.NewName = asmName
        '    TB_newDir.Text = rootAssemblyObject.NewRootDirectory
        '    ResetCarets()
        'End If


    End Sub

#End Region

#Region "Button Clicks"

    Private Sub CopyButton_Click(sender As Object, e As EventArgs) Handles CopyButton.Click
        CopyButtonHandler(Me, rootAssemblyObject, _invApp, sender, e)
    End Sub

    Private Sub NewDirButton_Click(sender As Object, e As EventArgs) Handles newDirButton.Click
        'Using NewDirectoryFolderBrowser As New FolderBrowserDialog()
        '    NewDirectoryFolderBrowser.SelectedPath = newDirectory
        '    If NewDirectoryFolderBrowser.ShowDialog() = DialogResult.OK Then
        '        newDirectory = NewDirectoryFolderBrowser.SelectedPath & "\"
        '        TB_newDir.Text = newDirectory
        '    End If
        'End Using
        NewDirectoryButtonHandler(Me, sender, e)
    End Sub

    Private Sub TestButton_Click(sender As Object, e As EventArgs) Handles TestButton.Click
        TestButtonClickHandler(sender, e, invApp:=_invApp)
    End Sub

    Private Sub BT_PreSuffix_Click(sender As Object, e As EventArgs) Handles BT_PreSuffix.Click
        'Dim node As TreeNode = TV_nComponent.SelectedNode
        'Dim oNodeText As String = node.Text
        'node.Text = TB_Prefix.Text & oNodeText & TB_Suffix.Text
        BT_PrefixSuffixHandler(Me, sender, e)
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
        Dim medium_gap As Integer = 10

        'stuff you don't want to scale with resize
        If initialLayout Then

            LB_ProjectDirectory.Height = standardHeight
            LB_ProjectDirectory.Left = medium_gap
            LB_ProjectDirectory.Top = 17

            TB_ProjDir.Height = standardHeight
            TB_ProjDir.Left = LB_ProjectDirectory.Left + LB_ProjectDirectory.Width
            TB_ProjDir.Top = LB_ProjectDirectory.Top
            TB_ProjDir.Width = CInt(clientWidth - TB_ProjDir.Left - medium_gap)

            LB_FileName.Height = standardHeight
            LB_FileName.Left = CInt(LB_ProjectDirectory.Left + LB_ProjectDirectory.Width - LB_FileName.Width)
            LB_FileName.Top = CInt(LB_ProjectDirectory.Top + LB_ProjectDirectory.Height + standardHeight)

            TB_FileName.Height = standardHeight
            TB_FileName.Left = TB_ProjDir.Left
            TB_FileName.Top = LB_FileName.Top
            TB_FileName.Width = CInt(clientWidth - TB_FileName.Left - medium_gap)

            LB_NewDirectory.Height = standardHeight
            LB_NewDirectory.Left = CInt(LB_ProjectDirectory.Left + LB_ProjectDirectory.Width - LB_NewDirectory.Width)
            LB_NewDirectory.Top = CInt(LB_FileName.Top + LB_FileName.Height + standardHeight)

            newDirButton.Width = 70
            newDirButton.Height = TB_newDir.Height
            newDirButton.Top = LB_NewDirectory.Top
            newDirButton.Left = CInt(clientWidth - newDirButton.Width - medium_gap)

            TB_newDir.Left = LB_NewDirectory.Left + LB_NewDirectory.Width
            TB_newDir.Height = standardHeight
            TB_newDir.Top = LB_NewDirectory.Top
            TB_newDir.Width = CInt(clientWidth - TB_newDir.Left - (clientWidth - newDirButton.Left) - medium_gap)

            CopyButton.Left = CInt(clientWidth / 2 - CopyButton.Width / 2)
            CopyButton.Top = CInt(clientHeight - medium_gap - CopyButton.Height)

            TestButton.Left = CopyButton.Left + CopyButton.Width + medium_gap
            TestButton.Top = CopyButton.Top

            LB_CopyComplete.Left = medium_gap
            LB_CopyComplete.Top = CopyButton.Top - LB_CopyComplete.Height - medium_gap

            TV_nComponent.Top = LB_NewDirectory.Top + LB_NewDirectory.Height + standardHeight
            TV_nComponent.Height = CInt(clientHeight - (TV_nComponent.Top) - (clientHeight - LB_CopyComplete.Top))
            TV_nComponent.Left = CInt(medium_gap)
            TV_nComponent.Width = CInt(clientWidth * 0.75 - medium_gap)


            GB_PreSuffix.Top = TV_nComponent.Top - medium_gap
            GB_PreSuffix.Left = TV_nComponent.Left + TV_nComponent.Width + medium_gap / 2
            GB_PreSuffix.Width = CInt(clientWidth - GB_PreSuffix.Left - medium_gap / 2)
            GB_PreSuffix.Height = CInt(standardHeight * 6)


            LB_Prefix.BringToFront()
            LB_Prefix.Left = medium_gap / 2
            LB_Prefix.Height = standardHeight
            LB_Prefix.Top = CInt(standardHeight)

            TB_Prefix.Left = LB_Prefix.Left + LB_Prefix.Width
            TB_Prefix.Height = standardHeight
            TB_Prefix.Top = LB_Prefix.Top
            TB_Prefix.Width = CInt(GB_PreSuffix.Width - TB_Prefix.Left - medium_gap / 2)

            LB_Suffix.Left = LB_Prefix.Left
            LB_Suffix.Height = standardHeight
            LB_Suffix.Top = LB_Prefix.Top + LB_Prefix.Height + standardHeight

            TB_Suffix.Left = TB_Prefix.Left
            TB_Suffix.Height = standardHeight
            TB_Suffix.Top = LB_Suffix.Top
            TB_Suffix.Width = TB_Prefix.Width

            'BT_PreSuffix.Left = CInt(LB_Prefix.Left + (LB_Prefix.Width + TB_Prefix.Width) / 2 - BT_PreSuffix.Width / 2)
            BT_PreSuffix.Left = GB_PreSuffix.Width / 2 - BT_PreSuffix.Width / 2
            BT_PreSuffix.Top = CInt(TB_Suffix.Top + TB_Suffix.Height + standardHeight / 2)

            GB_PreSuffix.SendToBack()

            Debug.WriteLine("Initial Layout Complete")

        End If


        'ResizeAssemblyNameLayout()
        ResetCarets()

    End Sub

    Private Sub ResetCarets()
        MoveCaret(TB_ProjDir)
        MoveCaret(TB_newDir)
    End Sub

    ''' <summary>
    ''' Writes text in a textbox and scrolls to the end
    ''' </summary>
    ''' <param name="textBox"></param>
    ''' <param name="msg"></param>
    Private Sub LongTextboxWrite(ByRef textBox As System.Windows.Forms.TextBox, ByRef msg As String)
        textBox.Text = msg
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub

    Private Sub MoveCaret(ByRef textBox As System.Windows.Forms.TextBox)
        textBox.SelectionStart = 0
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub


#End Region

#Region "Tree View Events"
    ' When a node is clicked, select it and give the treeview focus so the highlight is visible
    Private Sub TV_nComponent_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TV_nComponent.NodeMouseClick
        highlightSet.Clear()
        TV_nComponent.SelectedNode = e.Node
        TV_nComponent.Focus()
        selectedNode = e.Node

        Dim Ena As Boolean = True

        If Ena Then
            If TV_nComponent.SelectedNode.Parent Is Nothing Then
                'this is a root assembly occurence

            ElseIf TV_nComponent.SelectedNode.Parent Is rootAssemblyObject.TreeNode Then
                'this is a component of the root assembly
                For Each occ As ComponentOccurrence In TV_nComponent.SelectedNode.Tag
                    If occ IsNot Nothing Then highlightSet.AddItem(occ)
                Next
            Else
                For Each occProx As ComponentOccurrenceProxy In TV_nComponent.SelectedNode.Tag
                    highlightSet.AddItem(occProx)
                Next
            End If
        End If

        'rootAssemblyObject.HighlightOccurenceByNode(TV_nComponent.SelectedNode)
    End Sub

    Private Sub TV_nComponent_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TV_nComponent.NodeMouseDoubleClick
        doubleClickNode = e.Node
        Component_Modification.Show()
    End Sub

    ''' <summary>
    ''' Marks components that should not be copied by changing their background color to Gray
    ''' </summary>
    Public Sub DontCopyNode()
        doubleClickNode.ForeColor = System.Drawing.Color.Red
        If doubleClickNode.Nodes.Count > 0 Then
            For Each childNode As TreeNode In doubleClickNode.Nodes
                childNode.ForeColor = System.Drawing.Color.Red
            Next
        End If
    End Sub

    Sub AdjustRootDirectory()
        If doubleClickNode.Text = rootAssemblyObject.NewName Then
            'nothing needs to be done
        Else
            rootAssemblyObject.NewName = doubleClickNode.Text
            TB_newDir.Text = rootAssemblyObject.NewRootDirectory
            MoveCaret(TB_newDir)
        End If
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

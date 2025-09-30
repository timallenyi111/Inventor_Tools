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
    Dim projectDir As String
    Dim oAsmFileName As String
    Dim oAsmCompDef As AssemblyComponentDefinition
    Dim newDirectory As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next

        FormLayoutSetup(True)

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

        'setup the initial new directory based on the current project directory
        projectDir = GetProjectDirectory(_invApp)
        newDirectory = projectDir
        ProjDirTB.Text = projectDir
        NewDirectoryFolderBrowser.SelectedPath = projectDir
        newDirTB.Text = projectDir

        DB("Project Directory: " & projectDir)


        oAsmFileName = GetAssemblyFileName(oAsmDoc)
        FileNameTB.Text = oAsmFileName

        oAsmCompDef = oAsmDoc.ComponentDefinition

        DB(oAsmDoc.DisplayName)

        Dim mainAsmObj As New InvtAssemblyObj
        mainAsmObj = AssemblyObjSetup(oAsmDoc, mainAsmObj)

        ' Dim oTreeView As TreeView
        Dim oTreeView = oComponent_TV
        SetupTreeView(oTreeView, mainAsmObj)

        Dim newAsmObj As New InvtAssemblyObj
        newAsmObj = AssemblyObjSetup(oAsmDoc, newAsmObj)

        Dim nTreeView = nComponent_TV
        SetupTreeView(nTreeView, newAsmObj)


    End Sub

    Private Sub AssemblyCopyToolForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        FormLayoutSetup(False)
    End Sub

    Private Sub FormLayoutSetup(ByRef initialLayout As Boolean)
        Dim clientWidth As Integer = Me.ClientSize.Width
        Dim clientHeight As Integer = Me.ClientSize.Height
        'Dim tv_gap = clientWidth * 0.01
        Dim tv_gap = 10
        Dim tv_width = (clientWidth - (tv_gap * 3)) / 2


        'stuff you don't want to scale with resize
        If initialLayout Then
            DB("Client Width: " & clientWidth.ToString)
            DB("Client Height: " & clientHeight.ToString)
            Dim label1_2RightEdge As Integer = 140
            Dim label1_2Height As Integer = 25
            Label1.Height = label1_2Height
            Label1.Left = CInt(label1_2RightEdge - Label1.Width)
            Label1.Top = 17

            Dim label1_2Spacing As Integer = 25
            Label2.Height = label1_2Height
            Label2.Left = CInt(label1_2RightEdge - Label2.Width)
            Label2.Top = CInt(Label1.Top + Label1.Height + label1_2Spacing)

            ProjDirTB.Height = label1_2Height
            ProjDirTB.Left = label1_2RightEdge
            ProjDirTB.Top = Label1.Top

            FileNameTB.Height = label1_2Height
            FileNameTB.Left = CInt(label1_2RightEdge)
            FileNameTB.Top = Label2.Top

            Label3.Height = label1_2Height
            Label3.Left = CInt(label1_2RightEdge - Label3.Width)

            newDirButton.Width = 70
            newDirButton.Height = newDirTB.Height


            newDirTB.Left = CInt(label1_2RightEdge)
            newDirTB.Height = label1_2Height

        End If

        ProjDirTB.Width = CInt(clientWidth - ProjDirTB.Left - tv_gap)
        FileNameTB.Width = CInt(clientWidth - FileNameTB.Left - tv_gap)

        oComponent_TV.Width = CInt(tv_width)
        nComponent_TV.Width = CInt(tv_width)

        Dim tv_height As Integer = Me.ClientSize.Height * 0.5
        oComponent_TV.Height = CInt(tv_height)
        nComponent_TV.Height = CInt(tv_height)

        oComponent_TV.Left = CInt(tv_gap)
        nComponent_TV.Left = CInt(tv_width + tv_gap * 2)

        copyButton.Left = CInt(clientWidth / 2) - copyButton.Width / 2
        copyButton.Top = CInt(clientHeight / 14) * 13

        Dim newDirLabelTop As Integer = oComponent_TV.Top + oComponent_TV.Height + tv_gap * 1.5
        Label3.Top = newDirLabelTop

        newDirButton.Top = newDirLabelTop
        newDirButton.Left = tv_gap * 2 + tv_width * 2 - newDirButton.Width

        newDirTB.Top = newDirLabelTop
        newDirTB.Width = CInt(clientWidth - newDirTB.Left - tv_gap - newDirButton.Width)


    End Sub

    Private Sub NewDirButton_Click(sender As Object, e As EventArgs) Handles newDirButton.Click
        Using NewDirectoryFolderBrowser As New FolderBrowserDialog()
            NewDirectoryFolderBrowser.SelectedPath = newDirectory
            If NewDirectoryFolderBrowser.ShowDialog() = DialogResult.OK Then
                newDirectory = NewDirectoryFolderBrowser.SelectedPath & "\"
                newDirTB.Text = newDirectory
            End If
        End Using
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

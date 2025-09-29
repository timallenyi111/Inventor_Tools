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

	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		On Error Resume Next

		FormLayoutSetup()

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

		projectDir = GetProjectDirectory(_invApp)
		ProjDirTB.Text = projectDir

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
        FormLayoutSetup()
    End Sub

    Private Sub FormLayoutSetup()
		Dim tv_gap As Integer = Me.ClientSize.Width * 0.01
		Dim tv_width As Integer = (Me.ClientSize.Width - (tv_gap * 3)) / 2
		oComponent_TV.Width = CInt(tv_width)
        nComponent_TV.Width = CInt(tv_width)

		Dim tv_height As Integer = Me.ClientSize.Height * 0.5
		oComponent_TV.Height = CInt(tv_height)
        nComponent_TV.Height = CInt(tv_height)

        oComponent_TV.Left = CInt(tv_gap)
        nComponent_TV.Left = CInt(tv_width + tv_gap * 2)

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

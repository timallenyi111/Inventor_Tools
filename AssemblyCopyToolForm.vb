Imports System
Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor
Imports System.Diagnostics


Friend Class AssemblyCopyToolForm
	Inherits System.Windows.Forms.Form
	Public _invApp As Inventor.Application
	Public oAsmDoc As Inventor.AssemblyDocument
	Dim projectDir As String
	Dim oAsmFileName As String
	Dim oAsmCompDef As AssemblyComponentDefinition

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

		projectDir = GetProjectDirectory(_invApp)
        ProjDirTB.Text = projectDir

		oAsmFileName = GetAssemblyFileName(oAsmDoc)
		FileNameTB.Text = oAsmFileName

        oAsmCompDef = oAsmDoc.ComponentDefinition

		DB(oAsmDoc.DisplayName)

		Dim mainAsmObj As New InvtAssemblyObj
        mainAsmObj = SetupAssemblyObj(oAsmDoc, mainAsmObj)

		Dim oTreeView As New TreeView
		oTreeView = oComponent_TV
		SetupTreeView(oTreeView, mainAsmObj)


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

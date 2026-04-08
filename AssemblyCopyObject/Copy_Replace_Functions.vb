Imports Inventor
Imports frm = Inventor_Tools.AssemblyCopyToolForm
Module Copy_Replace_Functions

    Private Sub Log(message As String, Optional numLines As Integer = 0)
        AssemblyCopyToolForm.Log(message, numLines)
    End Sub

    ''' <summary>
    ''' Updates the iProperties "Part Number"
    ''' </summary>
    ''' <param name="curOcc"></param>
    ''' <param name="component"></param>
    ''' <param name="_invApp"></param>
    Public Sub UpdatePartNumber(curOcc As ComponentOccurrence, component As Object, _invApp As Inventor.Application)
        If TypeOf component Is InvtPartObj Then
            Dim part As InvtPartObj = CType(component, InvtPartObj)
            'Dim replacedPartDoc As PartDocument = _invApp.Documents.ItemByName(part.NewFullFileName)
            Dim replacedPartDoc As PartDocument = curOcc.Definition.Document
            replacedPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = part.NewName
            curOcc.Name = part.NewName
            Log("Updated Part Number for Part: " & part.NewName, numLines:=1)
        ElseIf TypeOf component Is AssemblyCopyObject Then
            Dim subAsy As AssemblyCopyObject = CType(component, AssemblyCopyObject)
            Dim replacedAsyDoc As AssemblyDocument = curOcc.Definition.Document
            'Dim replacedAsyDoc As AssemblyDocument = _invApp.Documents.ite(subAsy.NewFullFileName)
            replacedAsyDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = subAsy.NewName
            curOcc.Name = subAsy.NewName
            Log("Updated Part Number for Assembly: " & subAsy.NewName, numLines:=1)
        Else
            ' do nothing for other types
            Log("Component is neither InvtPartObj nor AssemblyCopyObject. No Part Number update performed.", numLines:=1)
        End If
    End Sub


End Module

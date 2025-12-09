Imports Inventor

Module TestFunctions



    Sub LoadOccurrenceProxy(ByVal _invApp As Inventor.Application, ByRef rootAssembly As AssemblyCopyObject, ByVal form As AssemblyCopyToolForm)
        Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
        Dim highlightSet As Inventor.HighlightSet = _invApp.ActiveDocument.CreateHighlightSet()

        For Each occurrence As Inventor.ComponentOccurrence In activeDoc.ComponentDefinition.Occurrences
            If occurrence.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                'Debug.WriteLine("Part Occurrence: " & occurrence.Name)
            ElseIf occurrence.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                'Debug.WriteLine("Assembly Occurrence: " & occurrence.Name)
                LoadSubOccurrence(occurrence.SubOccurrences, "    ", highlightSet)
            End If
        Next

    End Sub

    Sub LoadSubOccurrence(ByRef occurrences As Inventor.ComponentOccurrences, ByVal indent As String, ByVal highlightSet As Inventor.HighlightSet)
        Dim index As Integer = 1
        Dim searchName As String = "1HP Electric Motor:1"

        While index <= occurrences.Count
            Dim occ As Inventor.ComponentOccurrenceProxy = occurrences.Item(index)
            Debug.WriteLine(indent & "Sub-Occurrence " & index & ": " & occ.Name)
            If occ.Name = searchName Then
                Debug.WriteLine(">>> Found matching occurrence: " & occ.Name)
                highlightSet.AddItem(occ)
            End If
            index += 1
        End While
    End Sub


End Module

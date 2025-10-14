Module TestFunctions
    Sub ReadFrameAttribute(asmDoc As Inventor.AssemblyDocument)
        'Dim selectedCompOcc As Inventor.ComponentOccurrence = asmDoc.GetSelectedObject()
        Dim doc As Inventor.Document = asmDoc
        Dim selSet As Inventor.SelectSet = doc.SelectSet()
        Dim selectedOcc As Inventor.ComponentOccurrence = asmDoc.SelectSet.Item(1)
        Dim selectedDef As Inventor.AssemblyComponentDefinition = selectedOcc.Definition
        Debug.WriteLine(selectedOcc.Name)
        'Dim selectedOcc As Inventor.ComponentOccurrence = asmDoc.SelectSet.Item(1)
        'Dim attSet As Inventor.AttributeSet = selectedOcc.Definition.AttributeSets
        For Each attSet As Inventor.AttributeSet In selectedDef.AttributeSets
            Debug.WriteLine("att set name: " & attSet.Name)
            For Each atri As Inventor.Attribute In attSet
                Debug.WriteLine("att name: " & atri.Name)
                Debug.WriteLine("att value: " & atri.Value)
            Next
        Next

    End Sub
End Module

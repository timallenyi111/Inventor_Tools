Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.IO
Imports System.IO.IsolatedStorage
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization
Imports System.Security.Permissions
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

    Sub ReadSelectSet(ByVal _invApp As Inventor.Application)
        Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
        Dim selectSet As Inventor.SelectSet = activeDoc.SelectSet
        For i As Integer = 1 To selectSet.Count
            Dim selectedObj As Object = selectSet.Item(i)
            Dim type As Inventor.ObjectTypeEnum = selectSet.Item(i).Type
            If TypeOf selectedObj Is Inventor.ComponentOccurrence Then
                Dim occ As Inventor.ComponentOccurrence = selectedObj
                Debug.WriteLine("Selected Occurrence: " & occ.Name)
                ReadOccurrenceAttributes(occ)
            ElseIf type = Inventor.ObjectTypeEnum.kSketch3DProxyObject Then
                Dim sketchProxy As Inventor.Sketch3DProxy = selectedObj
                Debug.WriteLine("3D Sketch Proxy")
                For Each attSet As Inventor.AttributeSet In sketchProxy.AttributeSets
                    Debug.WriteLine("Attribute Set: " & attSet.Name)
                    For Each att As Inventor.Attribute In attSet
                        Debug.WriteLine("  Attribute Name: " & att.Name & ", Value: " & att.Value)
                    Next
                Next
            Else
                Debug.WriteLine("Selected Object is: " & type.ToString)
            End If
        Next
    End Sub

    Sub ReadOccurrenceAttributes(ByRef occurrence As Inventor.ComponentOccurrence)

        If occurrence.AttributeSets.Count > 0 Then
            Debug.WriteLine(occurrence.Name)
            For Each attSet As Inventor.AttributeSet In occurrence.AttributeSets
                Debug.WriteLine("Attribute Set: " & attSet.Name)
                For Each att As Inventor.Attribute In attSet
                    Debug.WriteLine("  Attribute Name: " & att.Name & ", Value: " & att.Value)
                Next
                Debug.WriteLine("-----")
            Next
        End If

    End Sub

    'Sub ReadOccurrenceDefinitionAttributes(ByRef _invApp As Inventor.Application)
    '    Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
    '    Dim selectSet As Inventor.SelectSet = activeDoc.SelectSet
    '    For i As Integer = 1 To selectSet.Count
    '        Dim selectedObj As Object = selectSet.Item(i)
    '        Dim type As Inventor.ObjectTypeEnum = selectSet.Item(i).Type
    '        If TypeOf selectedObj Is Inventor.ComponentOccurrence Then
    '            Dim occ As Inventor.ComponentOccurrence = selectedObj
    '            Debug.WriteLine("Selected Occurrence: " & occ.Name)
    '            ReadDefinitionAttributes(occ)
    '        Else
    '            Debug.WriteLine("Selected Object is: " & type.ToString)
    '        End If
    '        Dim occDef As Inventor.ComponentDefinition = occurrence.Definition
    '        If occDef.AttributeSets.Count > 0 Then
    '            Debug.WriteLine(occurrence.Name)
    '            For Each attSet As Inventor.AttributeSet In occDef.AttributeSets
    '                Debug.WriteLine("Attribute Set: " & attSet.Name)
    '                For Each att As Inventor.Attribute In attSet
    '                    Debug.WriteLine("  Attribute Name: " & att.Name & ", Value: " & att.Value.ToString)
    '                Next
    '                Debug.WriteLine("-----")
    '            Next
    '        End If
    '    Next

    'End Sub

    Sub ReadSelectionAttributes(ByRef _invApp As Inventor.Application)
        Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
        Dim selectSet As Inventor.SelectSet = activeDoc.SelectSet
        For i As Integer = 1 To selectSet.Count
            Dim selectedObj As Object = selectSet.Item(i)
            Dim type As Inventor.ObjectTypeEnum = selectSet.Item(i).Type
            If TypeOf selectedObj Is Inventor.ComponentOccurrence Then
                Dim occ As Inventor.ComponentOccurrence = selectedObj
                Debug.WriteLine("Selected Component: " & occ.Name)
                Debug.WriteLine("Occurrence Attributes:")
                For Each attSet As Inventor.AttributeSet In occ.AttributeSets
                    Debug.WriteLine("Attribute Set: " & attSet.Name)
                    For Each att As Inventor.Attribute In attSet
                        Debug.WriteLine(vbCrLf)
                        Debug.WriteLine("  Attribute Name: " & att.Name & ", Value: " & att.Value)
                        Debug.WriteLine("  Attribute Name: " & att.Name)
                        Debug.WriteLine("Value:")
                        Debug.WriteLine(att.Value)
                        Debug.WriteLine(vbCrLf)
                    Next
                    Debug.WriteLine("-----")
                Next
                Debug.WriteLine(vbCrLf & "-----------------------" & vbCrLf)

                'read component definition attributes
                Debug.WriteLine("Component Definition Attributes:")
                Dim compDef As Inventor.ComponentDefinition = occ.Definition
                For Each attSet As Inventor.AttributeSet In compDef.AttributeSets
                    Debug.WriteLine("Attribute Set: " & attSet.Name)
                    For Each att As Inventor.Attribute In attSet
                        Debug.WriteLine(vbCrLf)
                        Debug.WriteLine("  Attribute Name: " & att.Name)
                        Debug.WriteLine("Value:")
                        Debug.WriteLine(att.Value)
                        Debug.WriteLine(vbCrLf)
                    Next
                    Debug.WriteLine("-----")
                Next

                Debug.WriteLine(vbCrLf & "-----------------------" & vbCrLf)

                Dim doc As Inventor.AssemblyDocument = selectedObj.definition.Document
                Debug.WriteLine("Selected Document: " & doc.DisplayName)
                For Each attSet As Inventor.AttributeSet In doc.AttributeSets
                    Debug.WriteLine("Attribute Set: " & attSet.Name)
                    For Each att As Inventor.Attribute In attSet
                        Debug.WriteLine(vbCrLf)
                        Debug.WriteLine("  Attribute Name: " & att.Name)
                        Debug.WriteLine("Value:")
                        Debug.WriteLine(att.Value)
                        Debug.WriteLine(vbCrLf)
                    Next
                    Debug.WriteLine("-----")
                Next
            Else
                Debug.WriteLine("Selected Object is: " & type.ToString)
            End If
        Next

    End Sub

    Sub CreateAttributeLog(ByVal _invApp As Inventor.Application)
        Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
        Dim selectSet As Inventor.SelectSet = activeDoc.SelectSet

        For i As Integer = 1 To selectSet.Count
            Dim selectedObj As Object = selectSet.Item(i)
            Dim type As Inventor.ObjectTypeEnum = selectSet.Item(i).Type
            Debug.WriteLine("Selected Object Type: " & type.ToString)
            If TypeOf selectedObj Is Inventor.ComponentOccurrence Then
                Debug.WriteLine("Selected Component is a Component Occurence")

                Dim occurrence As Inventor.ComponentOccurrence = selectedObj
                Dim occDef As Inventor.ComponentDefinition = occurrence.Definition
                Debug.WriteLine("Occurence Attribute Count: " & occurrence.AttributeSets.Count)
                Debug.WriteLine("Occurrence Definition Attribute Count: " & occDef.AttributeSets.Count)

                If occDef.AttributeSets.Count > 0 Or occurrence.AttributeSets.Count > 0 Then
                    Debug.WriteLine("Creating Attribute Log")
                    Dim attriLogPath As String = "C:\Users\TimAllen\source\repos\timallenyi111\Inventor_Tools\LogFiles\"
                    Dim timestamp As String = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
                    Dim fileName As String = $"AttibuteLog_{timestamp}.txt"
                    Dim _stream As FileStream
                    Dim _writer As StreamWriter
                    _stream = New FileStream(attriLogPath & fileName, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)
                    _writer = New StreamWriter(_stream)
                    _writer.AutoFlush = True
                    _writer.WriteLine(occurrence.Name)
                    _writer.WriteLine("-----Occurrence Attributes-----")

                    For Each attSet As Inventor.AttributeSet In occurrence.AttributeSets
                        _writer.WriteLine("Attribute Set: " & attSet.Name)
                        For Each att As Inventor.Attribute In attSet
                            _writer.WriteLine("  Attribute Name: " & att.Name)
                            _writer.WriteLine("  Value: ")
                            For Each line In MultiLineAttributeValue(att.Value.ToString)
                                _writer.WriteLine(vbTab & vbTab & line)
                            Next
                            _writer.WriteLine("")
                            _writer.WriteLine("-------------")
                            _writer.WriteLine("")
                        Next
                    Next

                    _writer.WriteLine("-----Occurrence Definition Attributes-----")
                    For Each attSet As Inventor.AttributeSet In occDef.AttributeSets
                        _writer.WriteLine("Attribute Set: " & attSet.Name)
                        For Each att As Inventor.Attribute In attSet
                            _writer.WriteLine("  Attribute Name: " & att.Name)
                            _writer.WriteLine("  Value: ")
                            For Each line In MultiLineAttributeValue(att.Value.ToString)
                                _writer.WriteLine(vbTab & vbTab & line)
                            Next
                            _writer.WriteLine("")
                            _writer.WriteLine("-------------")
                            _writer.WriteLine("")
                        Next
                    Next
                    _stream.Close()
                    Debug.WriteLine("Attribute File Complete")
                End If

            End If

        Next


    End Sub

    Function MultiLineAttributeValue(ByRef attrValue As String) As List(Of String)
        Dim valueList As New List(Of String)
        Dim subString As String = ""
        While attrValue.Length > 0
            If attrValue.IndexOf(">") = -1 Then
                valueList.Add(attrValue)
                Exit While
            End If
            subString = attrValue.Substring(0, attrValue.IndexOf(">") + 1)
            valueList.Add(subString)
            attrValue = attrValue.Remove(0, subString.Length)
            Debug.WriteLine("Remaining Length: " & attrValue.Length)
        End While
        Return valueList
    End Function

    Sub GetRootAssemblyAttributes(_invApp As Inventor.Application)
        Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
        'Dim rootOcc As Inventor.ComponentOccurrence = activeDoc.ComponentDefinition.Occurrences.Item(1)
        Dim activeDocDef As Inventor.AssemblyComponentDefinition = activeDoc.ComponentDefinition
        Debug.WriteLine("Root Assembly Attribute Sets Count: " & activeDoc.AttributeSets.Count)
        Debug.WriteLine("Root Assembly Definition Atrribute Sets Count" & activeDocDef.AttributeSets.Count)
        Debug.WriteLine("Root Assembly Document Descriptor Count: " & activeDoc.ReferencedDocumentDescriptors.Count)
        For i As Integer = 1 To activeDoc.ReferencedDocumentDescriptors.Count
            Dim docDesc As Inventor.DocumentDescriptor = activeDoc.ReferencedDocumentDescriptors.Item(i)
            Debug.WriteLine(docDesc.DisplayName)
            Dim objType As Inventor.DocumentTypeEnum = docDesc.ReferencedDocumentType
            Debug.WriteLine(objType.ToString)
        Next



    End Sub

    Sub NameFirstOccurrence(_invApp As Inventor.Application)
        Dim activeDoc As Inventor.AssemblyDocument = _invApp.ActiveDocument
        Dim selectSet As Inventor.SelectSet = activeDoc.SelectSet
        Dim selectedOcc As Inventor.ComponentOccurrence = selectSet.Item(1)
        Dim asmDef As Inventor.AssemblyComponentDefinition = selectedOcc.Definition
        Dim firstOcc As Inventor.ComponentOccurrence = asmDef.Occurrences.Item(1)
        Debug.WriteLine("First Occurrence Name: " & firstOcc.Name)
    End Sub

End Module

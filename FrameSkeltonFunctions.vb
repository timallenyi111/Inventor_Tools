Imports Inventor

Module FrameSkeltonFunctions

    Sub ReadFrameAttribute(asmDoc As Inventor.AssemblyDocument)
        'Dim selectedCompOcc As Inventor.ComponentOccurrence = asmDoc.GetSelectedObject()
        Dim doc As Inventor.Document = asmDoc
        Dim selSet As Inventor.SelectSet = doc.SelectSet()
        Dim selectedOcc As Inventor.ComponentOccurrence = asmDoc.SelectSet.Item(1)
        If selectedOcc.DefinitionDocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim selectedDef As Inventor.AssemblyComponentDefinition = selectedOcc.Definition
            For Each attSet As Inventor.AttributeSet In selectedDef.AttributeSets
                'Debug.WriteLine("att set name: " & attSet.Name)
                For Each atri As Inventor.Attribute In attSet
                    'Debug.WriteLine("att name: " & atri.Name)
                    'Debug.WriteLine("att value: " & atri.Value)
                    If atri.Name = "Frame.Skeletons" Then
                        Dim oAtriVal As String = atri.Value
                        Dim skelIdStart As Integer = GetSkelIdStart(oAtriVal)
                        Dim skelIdEnd As Integer = GetSkelIdEnd(oAtriVal, skelIdStart)

                        Dim oSkelId As String = oAtriVal.Substring(skelIdStart, skelIdEnd - skelIdStart)

                        Dim nSkelId As String = GenerateNewSkelId(oSkelId)

                        Debug.WriteLine("New Skeleton ID: " & nSkelId)

                        Dim nAtriVal As String = oAtriVal.Substring(0, skelIdStart) & nSkelId &
                        oAtriVal.Substring(skelIdEnd)

                        Debug.WriteLine(vbCritical & "New Attribute Value: " & vbCrLf & nAtriVal)

                        'ReadSkeletonAttributes(selectedDef.Occurrences)

                        Dim skeletonComp As ComponentOccurrence = GetSkeletonOcc(selectedDef.Occurrences)
                        Debug.WriteLine(vbCrLf & "Skeleton Component Name : " & vbCrLf & skeletonComp.Name)
                        ReadSkeletonAttributes(selectedDef.Occurrences)
                    End If
                Next
            Next

            Dim frmCompOccs As Inventor.ComponentOccurrences = selectedDef.Occurrences

        ElseIf selectedOcc.DefinitionDocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
            Dim selectedDef As Inventor.PartComponentDefinition = selectedOcc.Definition
            For Each attSet As Inventor.AttributeSet In selectedOcc.AttributeSets
                Debug.WriteLine(vbCrLf & "part attribute set name: " & attSet.Name)
                For Each att As Inventor.Attribute In attSet
                    Debug.WriteLine(vbTab & "attribute Name" & att.Name)
                Next
            Next

            For Each attSet As Inventor.AttributeSet In selectedDef.AttributeSets
                'Debug.WriteLine("att set name: " & attSet.Name)
                For Each atri As Inventor.Attribute In attSet
                    'Debug.WriteLine("att name: " & atri.Name)
                    'Debug.WriteLine("att value: " & atri.Value)                    
                Next
            Next
        Else
            Debug.WriteLine("invalid selection")
        End If


        'Debug.WriteLine(selectedOcc.Name)
        'Dim selectedOcc As Inventor.ComponentOccurrence = asmDoc.SelectSet.Item(1)
        'Dim attSet As Inventor.AttributeSet = selectedOcc.Definition.AttributeSets


    End Sub

    Function GSkelID(ByVal atri As String) As String
        'remove everything before SkeletonId
        Dim skelIdStart As Integer = GetSkelIdStart(atri)
        Dim skelIdEnd As Integer = GetSkelIdEnd(atri, skelIdStart)
        Dim skelId As String = atri.Substring(skelIdStart, skelIdEnd - skelIdStart)

        'skelId = skelId.Substring(InStr(skelId, """"), InStr(skelId, ">") - 2 - InStr(skelId, """"))
        Debug.WriteLine(vbCrLf & vbCrLf & "Original Skeleton ID: " & skelId)
        Dim newSkelId As String = GenerateNewSkelId(skelId)
        Debug.WriteLine("New Skeleton ID: " & newSkelId)
        Return skelId
    End Function

    Function GetSkelIdStart(ByVal atri As String) As Integer
        Dim skelIDStart = InStr(atri, "SkeletonID")
        Dim skelId As String = atri.Substring(skelIDStart)
        skelIDStart = skelIDStart + InStr(skelId, """")
        Return skelIDStart
    End Function

    Function GetSkelIdEnd(ByVal atri As String, ByVal skelIdStart As Integer) As Integer
        Dim skelId As String = atri.Substring(skelIdStart)
        Dim skelIdEnd As Integer = skelIdStart + InStr(skelId, """") - 1
        Return skelIdEnd
    End Function

    Function GenerateNewSkelId(ByVal oSkelId As String) As String
        Dim newSkelIdEnd As String = oSkelId.Substring(oSkelId.LastIndexOf("-") + 1)
        Debug.WriteLine("SkeletonID End: " & newSkelIdEnd)
        Dim i As Integer = 0
        Dim rnd As New Random
        While i < newSkelIdEnd.Length
            Dim newInt As Integer = rnd.Next(0, 9)
            Dim newChar As String = newInt.ToString
            newSkelIdEnd = newSkelIdEnd.Substring(0, i) & newChar & newSkelIdEnd.Substring(i + 1)
            i += 1
        End While
        Dim newSkelId As String = oSkelId.Substring(0, oSkelId.LastIndexOf("-") + 1) & newSkelIdEnd
        Return newSkelId
    End Function

    Sub ReadSkeletonAttributes(ByRef frmOccs As Inventor.ComponentOccurrences)
        Dim idVal As String = Nothing
        For Each occ As Inventor.ComponentOccurrence In frmOccs
            Debug.WriteLine(occ.Name)
            For Each attSet As AttributeSet In occ.AttributeSets
                Debug.WriteLine(vbCrLf & "Attribute Set Name: " & attSet.Name)
                For Each atri As Attribute In attSet
                    If atri.Name = "ID" Then
                        idVal = atri.Value
                    End If
                    If atri.Name = "Type" Then
                        If atri.Value = "SkeletonType" Then
                            'this is the skeleton occurence
                            Debug.WriteLine(vbCrLf & "Skeleton Component ID: " & idVal)
                        End If
                    End If
                    'Debug.WriteLine("Attribute Name: " & atri.Name)
                    'Debug.WriteLine("Attribute Value: " & vbCrLf & atri.Value)
                Next
            Next
        Next
    End Sub

    Function GetSkeletonOcc(ByVal frmOccs As ComponentOccurrences) As ComponentOccurrence
        Dim skeletonOcc As ComponentOccurrence = Nothing
        For Each occ As ComponentOccurrence In frmOccs
            For Each attSet In occ.AttributeSets
                For Each ati As Attribute In attSet
                    Debug.WriteLine(ati.Name)
                    If ati.Name = "Type" Then
                        If ati.Value = "SkeletonType" Then
                            skeletonOcc = occ
                            Return skeletonOcc
                        End If
                    End If
                Next
            Next
        Next
        Return skeletonOcc
    End Function

End Module

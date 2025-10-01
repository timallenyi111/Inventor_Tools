Imports Inventor
Module InventorFunctions

    Function AssemblyObjSetup(ByRef asmDoc As Inventor.AssemblyDocument, invtAsmObj As InvtAssemblyObj) As InvtAssemblyObj
        invtAsmObj.OriginalFullFileName = asmDoc.FullFileName
        DB("Original File Name: " & invtAsmObj.OriginalFullFileName)
        invtAsmObj.OriginalFilePath = GetFilePath(asmDoc)
        invtAsmObj.OriginalName = GetComponentName(asmDoc)
        invtAsmObj.NewName = invtAsmObj.OriginalName
        invtAsmObj.OriginalAssemblyDocument = asmDoc

        Dim asmCompDef As Inventor.AssemblyComponentDefinition = GetAssemblyComponentDefinition(asmDoc)
        'invtAsmObj.AssemblyComponentDefinition = asmCompDef

        For Each occ As Inventor.ComponentOccurrence In asmCompDef.Occurrences
            Dim occDoc As Inventor.Document = occ.Definition.Document

            If invtAsmObj.CheckIfComponentExists(occDoc.FullFileName, 1) Then
                ' do nothing because we already added to the quantity

            Else
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    ' curCompObj.Type = "Assembly"

                    Dim subAsmDoc As Inventor.AssemblyDocument = occ.Definition.Document
                    Dim subAsmObj As InvtAssemblyObj = AssemblyObjSetup(subAsmDoc, New InvtAssemblyObj)

                    Dim curCompObj As New InvtComponentObj With {
                        .Name = GetComponentName(occ.Definition.Document),
                        .OriginalFileName = occDoc.FullFileName,
                        .Type = "Assembly",
                        .AssemblyObject = subAsmObj
                    }

                    invtAsmObj.AssemblyComponents.Add(curCompObj)

                ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then

                    Dim partDoc As PartDocument = occ.Definition.Document
                    Dim partObj As InvtPartObj = SetupPartObject(partDoc)
                    Dim curCompObj As New InvtComponentObj With {
                        .Name = GetComponentName(occ.Definition.Document),
                        .OriginalFileName = occDoc.FullFileName,
                        .Type = "Part",
                        .PartObject = partObj
                    }

                    invtAsmObj.AssemblyComponents.Add(curCompObj)

                Else
                    DB("Unknown document type: " & occ.DefinitionDocumentType.ToString)
                    Continue For
                End If
            End If
        Next

        Return invtAsmObj
    End Function

    Function SetupPartObject(ByRef partDoc As Inventor.PartDocument) As InvtPartObj
        Dim invtPartObj As New InvtPartObj
        invtPartObj.OriginalName = GetComponentName(partDoc)
        invtPartObj.OriginalFullFileName = partDoc.FullFileName
        invtPartObj.OriginalFilePath = GetFilePath(partDoc)
        invtPartObj.OriginalPartDocument = partDoc

        invtPartObj.NewName = invtPartObj.OriginalName

        Return invtPartObj
    End Function

    Function GetProjectDirectory(invApp As Inventor.Application) As String
        Dim actProj As Inventor.DesignProject = invApp.DesignProjectManager.ActiveDesignProject
        Dim projectDir As String = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\") + 1)
        Return projectDir
    End Function

    ''' <summary>
    ''' Returns the assembly file name only, without the path
    ''' </summary>
    Function GetAssemblyFileName(asmDoc As Inventor.AssemblyDocument) As String
        Dim asmFileName As String = asmDoc.FullFileName.Substring(asmDoc.FullFileName.LastIndexOf("\") + 1)
        Return asmFileName
    End Function

    Function GetAssemblyComponentDefinition(asmDoc As Inventor.AssemblyDocument) As Inventor.ComponentDefinition
        Dim asmCompDef As Inventor.ComponentDefinition = asmDoc.ComponentDefinition
        Return asmCompDef
    End Function

    Function GetListOfComponentNames(ByRef asmCompDef As Inventor.ComponentDefinition) As List(Of String)
        Dim compNames As New List(Of String)
        Dim oOccs As Inventor.ComponentOccurrences = asmCompDef.Occurrences
        Dim oOcc As Inventor.ComponentOccurrence
        For Each oOcc In oOccs
            compNames.Add(oOcc.Name)
        Next
        Return compNames
    End Function

    Sub DB_List(list As List(Of String))
        ''' Debug Print the list of strings
        Dim result As String = ""
        For Each item As String In list
            Debug.Print(item)
        Next
    End Sub

    Sub DB(ByRef msg As String)
        ''' Debug Print the message
        Debug.Print(msg)
    End Sub

    Function GetListOfAssmblyOccurrences(ByRef asmCompDef As Inventor.ComponentDefinition) As List(Of Inventor.ComponentOccurrence)
        ''' Returns a list of assembly occurrences in the given assembly component definition (deletes duplicates)
        Dim asmOccs As New List(Of Inventor.ComponentOccurrence)
        Dim oOccs As Inventor.ComponentOccurrences = asmCompDef.Occurrences
        Dim oOcc As Inventor.ComponentOccurrence
        For Each oOcc In oOccs
            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                asmOccs.Add(oOcc)
            End If
        Next
        Return asmOccs
    End Function

    Function AddPrefixSuffix(ByRef fileName As String, Optional prefix As String = "", Optional ByRef suffix As String = "") As String
        ''' Adds a suffix to the file name before the file extension
        Dim newFileName As String = prefix & fileName
        If fileName.Contains(".") Then
            newFileName = fileName.Substring(0, fileName.LastIndexOf(".")) & suffix & fileName.Substring(fileName.LastIndexOf("."))
        Else
            newFileName = fileName & suffix
        End If
        Return newFileName
    End Function

    Function GetListOfPartOccurrences(ByRef asmCompDef As Inventor.ComponentDefinition) As List(Of Inventor.ComponentOccurrence)
        ''' Returns a list of part occurrences in the given assembly component definition (deletes duplicates)
        Dim partOccs As New List(Of Inventor.ComponentOccurrence)
        Dim oOccs As Inventor.ComponentOccurrences = asmCompDef.Occurrences
        Dim oOcc As Inventor.ComponentOccurrence
        For Each oOcc In oOccs
            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                partOccs.Add(oOcc)
            End If
        Next
        Return partOccs
    End Function

    Function GetListOfDisplayNames(ByRef occs As List(Of Inventor.ComponentOccurrence)) As List(Of String)
        Dim displayNames As New List(Of String)
        Dim oOcc As Inventor.ComponentOccurrence
        For Each oOcc In occs
            displayNames.Add(oOcc.Name)
        Next
        Return displayNames
    End Function

    Function GetFilePath(ByRef occ As Inventor.Document) As String
        Dim filePath As String = occ.FullFileName
        filePath = filePath.Substring(0, filePath.LastIndexOf("\") + 1)
        Return filePath
    End Function

    ''' <summary>
    ''' returns the component name without the file extension
    ''' </summary>
    ''' <param name="occ"></param>
    ''' <returns></returns>
    Function GetComponentName(ByRef occ As Inventor.Document) As String
        Dim compName As String = occ.DisplayName
        ' DB(compName)
        ' DB(occ.SubType.ToString)
        If compName.Contains(".") Then
            compName = compName.Substring(0, compName.LastIndexOf("."))
        End If
        Return compName
    End Function

    Function CheckIfComponentIsImported(ByRef occ As Inventor.ComponentOccurrence) As Boolean
        Dim isImported As Boolean = False
        If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As Inventor.PartDocument = occ.Definition.Document

            DB(partDoc.SubType.ToString)
            If partDoc.SubType = "Imported" Then
                isImported = True
            End If

        ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim asmDoc As Inventor.AssemblyDocument = occ.Definition.Document
            Dim count As Long = asmDoc.ReferencedDocumentDescriptors.Count()
            Dim countInt As Integer = CInt(count)
            DB(GetComponentName(asmDoc))
            Debug.Print("Referenced Assembly Count: " & countInt)
            DB("----")
        End If
        Return isImported
    End Function

    Function CheckIfDuplicateComponent(ByRef occ As Inventor.ComponentOccurrence, ByRef compFileNames As List(Of String)) As Boolean
        Dim isDuplicate As Boolean = False
        Dim curDoc As Inventor.Document = occ.Definition.Document
        Dim curFileName As String = curDoc.FullFileName
        If compFileNames.Contains(curFileName) Then
            isDuplicate = True
        End If
        Return isDuplicate
    End Function

    Function CheckIfFrame(ByRef asmDoc As AssemblyDocument) As Boolean
        Dim isFrame As Boolean = False
        Dim occ As Inventor.ComponentOccurrence
        For Each occ In asmDoc.ComponentDefinition.Occurrences
            If GetComponentName(occ.Definition.Document).ToLower.Contains("skeleton") Then
                isFrame = True
                DB("Frame found: " & GetComponentName(asmDoc))
                Dim skeletonDoc As Inventor.PartDocument = occ.Definition.Document
                DB("Referenced Document Count: " & asmDoc.ReferencedDocumentDescriptors.Count)
                Dim refDoc As Inventor.Document
                Dim fileName As String
                For Each refDoc In asmDoc.ReferencingDocuments
                    fileName = refDoc.FullFileName
                    fileName = fileName.Substring(fileName.LastIndexOf("\") + 1)
                    DB("   Ref: " & fileName)
                Next

            End If
        Next

        Return isFrame
    End Function



End Module

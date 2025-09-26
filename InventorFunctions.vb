Imports Inventor
Module InventorFunctions

    Function GetProjectDirectory(invApp As Inventor.Application) As String
        Dim actProj As Inventor.DesignProject = invApp.DesignProjectManager.ActiveDesignProject
        Dim projectDir As String = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\"))
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

    Function GetComponentName(ByRef occ As Inventor.Document) As String
        Dim compName As String = occ.DisplayName
        compName = compName.Substring(0, compName.LastIndexOf("."))
        Return compName
    End Function

    Function SetupAssemblyObj(ByRef asmDoc As Inventor.AssemblyDocument, invtAsmFile As InvtAssemblyObj) As InvtAssemblyObj
        invtAsmFile.FileName = GetAssemblyFileName(asmDoc)
        invtAsmFile.Name = GetComponentName(asmDoc)
        Dim asmCompDef As Inventor.AssemblyComponentDefinition = GetAssemblyComponentDefinition(asmDoc)
        invtAsmFile.AssemblyComponentDefinition = asmCompDef
        invtAsmFile.ComponentNames = GetListOfComponentNames(asmCompDef)
        For Each occ As Inventor.ComponentOccurrence In GetListOfAssmblyOccurrences(asmCompDef)
            If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                If invtAsmFile.AssemblyNames.Contains(GetComponentName(occ.Definition.Document)) Then
                    'do nothing, already in the list
                Else
                    Dim subAsmFile As New InvtAssemblyObj
                    Dim subAsmDoc As Inventor.AssemblyDocument = occ.Definition.Document
                    subAsmFile = SetupAssemblyObj(subAsmDoc, subAsmFile)
                    invtAsmFile.SubAssemblyObjectList.Add(subAsmFile)
                    invtAsmFile.AssemblyNames.Add(subAsmFile.Name)
                End If
            ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                If invtAsmFile.PartNames.Contains(GetComponentName(occ.Definition.Document)) Then
                    'do nothing, already in the list
                Else
                    Dim partObj As New InvtPartObj
                    partObj = SetupPartObject(occ.Definition.Document)
                    invtAsmFile.PartNames.Add(partObj.Name)
                End If
            Else
                Continue For
            End If

        Next
        Return invtAsmFile
    End Function

    Function SetupPartObject(ByRef partDoc As Inventor.PartDocument) As InvtPartObj
        Dim invtPartObj As New InvtPartObj
        invtPartObj.Name = partDoc.ComponentDefinition.Name
        Return invtPartObj
    End Function

    Sub SetupTreeView(ByRef treeView As TreeView, ByRef asmObj As InvtAssemblyObj)
        treeView.Nodes.Clear()
        Dim rootNode As TreeNode = treeView.Nodes.Add(asmObj.Name)
        AddSubNodes(rootNode, asmObj)
        treeView.ExpandAll()
    End Sub

    Sub AddSubNodes(ByRef parentNode As TreeNode, ByRef asmObj As InvtAssemblyObj)
        Dim asmNode As TreeNode
        Dim partNode As TreeNode
        For Each subAsm As InvtAssemblyObj In asmObj.SubAssemblyObjectList
            asmNode = parentNode.Nodes.Add(subAsm.Name)
            AddSubNodes(asmNode, subAsm)
        Next
        For Each partName As String In asmObj.PartNames
            partNode = parentNode.Nodes.Add(partName)
        Next
    End Sub

End Module

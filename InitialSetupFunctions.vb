
Imports System.Windows
Imports Inventor


Module InitialSetupFunctions
    Private _ContentCenterPath As String = String.Empty
    Private _NewRootDirectory As String = String.Empty
    Private _ProjectDirectory As String = String.Empty
    Private _form As AssemblyCopyToolForm
    Function InitialSetup(ByRef inventorApp As Inventor.Application, ByRef form As AssemblyCopyToolForm) As InvtAssembly
        'store the form in as a global variable
        _form = form
        Dim rootAssemblyDoc As Inventor.AssemblyDocument

        'try setting the root assembly document
        Try
            rootAssemblyDoc = inventorApp.ActiveDocument
        Catch ex As Exception
            MessageBox.Show("Please open an assembly document before running the copy process.", "No Assembly Document Found")
            Return Nothing
        End Try

        '*******SET THE INITIAL SETUP GLOBAL VARIABLES********
        'get the project directory to use as a base path for the copied assembly
        Dim actProj As Inventor.DesignProject = inventorApp.DesignProjectManager.ActiveDesignProject
        _ProjectDirectory = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\") + 1)
        _ContentCenterPath = inventorApp.DesignProjectManager.ActiveDesignProject.ContentCenterPath
        Dim rootAssemblyFullFileName As String = rootAssemblyDoc.FullFileName
        _NewRootDirectory = SetNewRootDirectory(_ProjectDirectory, rootAssemblyFullFileName)
        '*****************************************************

        'create the root assembly object (the occurrence index for the root is 0 (it doesn't really matter))      
        Dim rootAssemblyObject As New InvtAssembly(rootAssemblyDoc, 0, nRootDirectory:=_NewRootDirectory)
        'setup the new properties in the root assembly object        
        rootAssemblyObject.NewName = _form.TB_Prefix.Text & rootAssemblyObject.OriginalName & _form.TB_Suffix.Text
        rootAssemblyObject.TreeNode = New TreeNode(rootAssemblyObject.NewName)

        'add all of the components to the assembly object
        rootAssemblyObject = AddSubComponents(rootAssemblyObject)

        'clear the parts that are highlighted on startup by clearing the active document selectset
        rootAssemblyDoc.SelectSet.Clear()

        'assign the necessary tags to the tree nodes for highlighting in the Inventor model
        AssignNodeTags(rootAssemblyObject, parentIsRoot:=True)

        'log the results of the initial setup
        LogAssembly(rootAssemblyObject, isRoot:=True, startingMessage:="*****INITIAL SETUP*****")

        Return rootAssemblyObject
    End Function

    ''' <summary>
    ''' Adds all components to the part and assembly list of the input assembly object. This function is recursive so it will continue to add subcomponents until there are no more subassemblies in the structure.
    ''' </summary>
    ''' <param name="parentAsmObject"></param>
    ''' <param name="parentIsRoot">Indicates whether the parent assembly is the root assembly.</param>
    ''' <returns></returns>
    Private Function AddSubComponents(ByRef parentAsmObject As InvtAssembly) As InvtAssembly
        Dim compOccs As Inventor.ComponentOccurrences = parentAsmObject.OriginalAsmDocument.ComponentDefinition.Occurrences
        Dim curOccIndex As Integer = 1
        'Setup all components in the root assembly object
        'seperate the parts, assemblies, and frame assemblies
        For Each curOcc As Inventor.ComponentOccurrence In compOccs

            If curOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject And parentAsmObject.CheckForDuplicatePart(curOcc, curOccIndex) = False Then
                'This is a part occurrence and not a duplicate
                Dim newPart As InvtPart = SetupInventorPart(curOcc, parentAsmObject.TreeNode, curOccIndex)
                'add new part to the part list
                parentAsmObject.AddPartToList(newPart)

            ElseIf curOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject And parentAsmObject.CheckForDuplicateAssembly(curOcc, curOccIndex) = False Then
                'this is an assembly occurrence and not a duplicate
                'now we need to check if this is a frame or a regular assembly.
                If CheckIfOccurenceIsFrame(curOcc) Then
                    'this is a frame assembly                    
                    'create an InvtAssembly object for setting up a new InvtFrame Object
                    Dim newFrameAssemblyObject As InvtAssembly = SetupFrameCoreAssembly(curOcc, parentAsmObject.TreeNode, curOccIndex)
                    'now create the new frame object using the new frame assembly object
                    Dim newFrameObject As New InvtFrame(newFrameAssemblyObject, parentAsmObject.NewName)

                    'now add the frame object to the SubFrameList
                    parentAsmObject.AddSubFrameToList(newFrameObject)
                Else
                    'this is a regular assembly
                    Dim newSubAsy As InvtAssembly = SetupSubAssembly(curOcc, parentAsmObject.TreeNode, curOccIndex)

                    'add new subassembly to the subassembly list
                    parentAsmObject.AddSubAssemblyToList(newSubAsy)
                End If
            Else
                '_form.Log(curOcc._DisplayName & " was a duplicate", numTabs:=_logTab)
            End If
            curOccIndex += 1
        Next

        Return parentAsmObject
    End Function

    Private Function SetupInventorPart(ByRef partOcc As Inventor.ComponentOccurrence, ByRef ParentAssemblyNode As TreeNode,
                                        ByRef curOccIndex As Integer) As InvtPart

        'the part document can be accessed from the occurrence definition
        Dim partDoc As Inventor.PartDocument = partOcc.Definition.Document

        If partDoc.FullFileName.ToLower.StartsWith(_ContentCenterPath.ToLower) Then
            'This is a content center part
            Dim newPartObject As New InvtPart(partDoc, partOcc, curOccIndex, _NewRootDirectory, isContentCenter:=True)
            newPartObject.NewName = newPartObject.OriginalName ' For now.... we aren't changing the new names by default            
            newPartObject.TreeNode = ParentAssemblyNode.Nodes.Add(newPartObject.NewName)
            Return newPartObject
        Else
            Dim newPartObject As New InvtPart(partDoc, partOcc, curOccIndex, _NewRootDirectory)
            newPartObject.NewName = newPartObject.OriginalName ' For now.... we aren't changing the new names by default
            newPartObject.TreeNode = ParentAssemblyNode.Nodes.Add(newPartObject.NewName)
            Return newPartObject
        End If

    End Function

    Private Function SetupSubAssembly(ByRef assemblyOcc As Inventor.ComponentOccurrence, ByRef ParentAssemblyNode As TreeNode, ByRef occIndex As Integer) As InvtAssembly
        Dim assemblyDoc As Inventor.AssemblyDocument = assemblyOcc.Definition.Document

        Dim newAssemblyObject As New InvtAssembly(assemblyDoc, occIndex, nRootDirectory:=_NewRootDirectory, AsyOcc:=assemblyOcc)

        newAssemblyObject.NewName = newAssemblyObject.OriginalName ' For now.... we aren't changing the new names by default
        newAssemblyObject.TreeNode = ParentAssemblyNode.Nodes.Add(newAssemblyObject.NewName)

        'setup all components in the subassembly
        'this is recursive so it will continue to go down the structure until there are no more subassemblies
        newAssemblyObject = AddSubComponents(newAssemblyObject)

        Return newAssemblyObject
    End Function

    ''' <summary>
    ''' Sets up an InvtAssembly Object to be used for initializing a new InvtFrameObject
    ''' </summary>
    ''' <param name="frameOcc"></param>
    ''' <param name="ParentAssemblyNode"></param>
    ''' <returns></returns>
    Private Function SetupFrameCoreAssembly(ByRef frameOcc As Inventor.ComponentOccurrence, ByRef ParentAssemblyNode As TreeNode, ByRef occIndex As Integer) As InvtAssembly
        Dim frameDoc As Inventor.AssemblyDocument = frameOcc.Definition.Document
        Dim newAssemblyObject As New InvtAssembly(frameDoc, occIndex, nRootDirectory:=_NewRootDirectory, AsyOcc:=frameOcc)
        newAssemblyObject.TreeNode = ParentAssemblyNode.Nodes.Add(newAssemblyObject.NewName)

        'setup all components in the frame assembly
        'this is recursive so it will continue to go down the structure until there are no more subassemblies
        newAssemblyObject = AddSubComponents(newAssemblyObject)

        Return newAssemblyObject
    End Function

    Private Function CheckIfOccurenceIsFrame(ByRef compOcc As ComponentOccurrence) As Boolean
        Dim isFrame As Boolean = False
        If compOcc.AttributeSets.Count > 0 Then
            'Debug.WriteLine("Assembly: " & compOcc.Name)
        End If

        For Each attSet As AttributeSet In compOcc.AttributeSets
            For Each atri As Inventor.Attribute In attSet
                'Debug.WriteLine("Attribute Value: " & atri.Value.ToString)
                'Debug.WriteLine("atribute value type " & atri.Value.GetType.ToString)
                If VarType(atri.Value) = vbString Then
                    If atri.Value = "MasterFrameOcc" Then
                        isFrame = True
                        'Debug.WriteLine(compOcc.Name & " is a frame assembly")
                        'Debug.WriteLine(atri.Value)
                    End If
                End If
            Next

            For i As Integer = 0 To 3
                'Debug.WriteLine("")
            Next

        Next

        'Debug.WriteLine("*****")
        Return isFrame
    End Function

    ''' <summary>
    ''' Assigns the necessary tag to the component tree node for highlighting components (along with their duplicates) in the Inventor model.  
    ''' | ROOT ASSEMBLY COMPONENTS -> component occurrence
    ''' | SUB-ASSEMBLY COMPONENTS -> component occurrence proxy
    ''' </summary>
    ''' <param name="parentAssemblyObject"></param>
    ''' <param name="parentIsRoot"></param>
    Private Sub AssignNodeTags(ByRef parentAssemblyObject As InvtAssembly, ByRef parentIsRoot As Boolean)
        If parentIsRoot Then
            parentAssemblyObject.TreeNode.Tag = parentAssemblyObject

            For Each part In parentAssemblyObject.PartList
                Dim occList As List(Of Inventor.ComponentOccurrence) = part.DuplicateOccurrences
                occList.Add(part.OriginalComponentOccurrence)
                part.TreeNode.Tag = occList
            Next
            For Each subAsm In parentAssemblyObject.AssemblyList
                Dim occList As List(Of Inventor.ComponentOccurrence) = subAsm.DuplicateOccurrences
                occList.Add(subAsm.OriginalComponentOccurrence)
                subAsm.TreeNode.Tag = occList
                AssignNodeTags(subAsm, parentIsRoot:=False)
            Next
            For Each subFrame In parentAssemblyObject.FrameList
                Dim occList As List(Of Inventor.ComponentOccurrence) = subFrame.DuplicateOccurrences
                occList.Add(subFrame.OriginalComponentOccurrence)
                subFrame.TreeNode.Tag = occList
                AssignNodeTags(subFrame.CoreAssemblyObject, parentIsRoot:=False)
            Next

        Else
            Dim parentAsmOccurrences As Inventor.ComponentOccurrences = parentAssemblyObject.OriginalComponentOccurrence.SubOccurrences
            For Each part In parentAssemblyObject.PartList
                Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
                Dim occProxy As Inventor.ComponentOccurrenceProxy = parentAsmOccurrences.Item(part.OccurrenceIndex)
                occProxyList.Add(occProxy)
                For Each occIndex In part.DuplicateOccurrenceIndexList
                    occProxyList.Add(parentAsmOccurrences.Item(occIndex))
                Next
                part.TreeNode.Tag = occProxyList
            Next

            For Each subAsm In parentAssemblyObject.AssemblyList
                Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
                Dim occProxy As Inventor.ComponentOccurrenceProxy = parentAsmOccurrences.Item(subAsm.OccurrenceIndex)
                occProxyList.Add(occProxy)
                For Each occIndex In subAsm.DuplicateOccurrenceIndexList
                    occProxyList.Add(parentAsmOccurrences.Item(occIndex))
                Next
                subAsm.TreeNode.Tag = occProxyList
                AssignNodeTags(subAsm, parentIsRoot:=False)
            Next

            For Each subFrame In parentAssemblyObject.FrameList
                Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
                Dim occProxy As Inventor.ComponentOccurrenceProxy = parentAsmOccurrences.Item(subFrame.OccurrenceIndex)
                occProxyList.Add(occProxy)
                For Each occIndex In subFrame.DuplicateOccurrenceIndexList
                    occProxyList.Add(parentAsmOccurrences.Item(occIndex))
                Next
                subFrame.TreeNode.Tag = occProxyList

                AssignNodeTags(subFrame.CoreAssemblyObject, parentIsRoot:=False)
            Next
        End If

    End Sub

    ''' <summary>
    ''' Creates a new root directory for the copied assembly based on the project directory and the name of the root assembly. The new root directory will be used as the base file path for all copied components in the structure, so it is important that this is set correctly before any names or file paths are changed. The new root directory is set to be a folder with the same name as the root assembly in the project directory. For example, if the project directory is "C:\Projects\MyProject\" and the root assembly name is "MyAssembly.iam" then the new root directory will be set to "C:\Projects\MyProject\MyAssembly\".
    ''' </summary>
    ''' <param name="projectPath"></param>
    ''' <param name="rootAssemblyName"></param>
    ''' <returns></returns>
    Private Function SetNewRootDirectory(ByRef projectPath As String, ByRef rootAssemblyFullFileName As String) As String
        Dim rootAssemblyName As String = rootAssemblyFullFileName.Substring(rootAssemblyFullFileName.LastIndexOf("\") + 1)
        'now remove the .iam
        rootAssemblyName = rootAssemblyName.Substring(0, rootAssemblyName.Length - 4)
        If _ProjectDirectory = String.Empty Then
            Throw New Exception("Project directory is not set. Cannot set new root directory.")
        End If

        Dim NewRootDirectory = _ProjectDirectory & _form.TB_Prefix.Text & rootAssemblyName & _form.TB_Suffix.Text & "\"
        Return NewRootDirectory
    End Function

#Region "Logging Functions"

    Public Sub LogAssembly(ByRef parentAsmObj As InvtAssembly, Optional ByRef tabIndex As Integer = 0, Optional ByRef isRoot As Boolean = False,
                            Optional startingMessage As String = "")
        If isRoot Then
            'this is the root assembly
            _form.Log(startingMessage, numLinesBefore:=3, debugWrite:=False)
            _form.Log("New Root Directory: " & _NewRootDirectory, debugWrite:=False)
            _form.Log("Original Root Assembly Name: " & parentAsmObj.OriginalName, debugWrite:=False)
            _form.Log("New Root Assembly Name: " & parentAsmObj.NewName, debugWrite:=False)
            _form.Log("Original Root Assembly Full File Name: " & RemoveRootDirectory(parentAsmObj.OriginalFullFileName), debugWrite:=False)
            _form.Log("New Root Assembly Full File Name: " & RemoveRootDirectory(parentAsmObj.NewFullFileName), debugWrite:=False)
            'increase the tab index for subcomponents
            tabIndex += 1
        Else
            _form.Log("*ASSEMBLY*", numLinesBefore:=1, numTabs:=tabIndex, debugWrite:=False)
            _form.Log("New Assembly Name: " & parentAsmObj.NewName, numTabs:=tabIndex, debugWrite:=False)
            _form.Log("Original Assembly Full File Name: " & RemoveRootDirectory(parentAsmObj.OriginalFullFileName), numTabs:=tabIndex, debugWrite:=False)
            _form.Log("New Full File Name: " & RemoveRootDirectory(parentAsmObj.NewFullFileName), numTabs:=tabIndex, debugWrite:=False)
            _form.Log("Number of Duplicate Occurrences: " & parentAsmObj.DuplicateOccurrences.Count, numTabs:=tabIndex + 1, debugWrite:=False)
            _form.Log("Copy Enabled: " & parentAsmObj.CopyEnabled.ToString(), numTabs:=tabIndex + 1, debugWrite:=False)

            'increase the tab index for subcomponents
            tabIndex += 1
        End If

        For Each part In parentAsmObj.PartList
            LogPart(part, tabIndex)
        Next
        For Each subAsm In parentAsmObj.AssemblyList
            LogAssembly(subAsm, tabIndex)
        Next
        For Each subFrame In parentAsmObj.FrameList
            LogFrameAssembly(subFrame, tabIndex)
        Next

        _form.Log("****Initial Setup Complete!****", numLinesAfter:=3, numLinesBefore:=1, debugWrite:=False)

    End Sub

    Private Sub LogPart(ByRef part As InvtPart, ByRef tabIndex As Integer)
        If part.IsContentCenter Then
            _form.Log("*CONTENT CENTER PART*", numTabs:=tabIndex, numLinesBefore:=1, debugWrite:=False)
            _form.Log("Part Name: " & part.OriginalName, numTabs:=tabIndex, debugWrite:=False)
            _form.Log("Number of Duplicate Occurrences: " & part.DuplicateOccurrences.Count, numTabs:=tabIndex + 1, debugWrite:=False)
        Else
            _form.Log("*PART*", numTabs:=tabIndex, numLinesBefore:=1, debugWrite:=False)
            _form.Log("New Part Name: " & part.NewName, numTabs:=tabIndex, debugWrite:=False)
            _form.Log("Original Part Full File Name: " & RemoveRootDirectory(part.OriginalFullFileName), numTabs:=tabIndex, debugWrite:=False)
            _form.Log("New Part Full File Name: " & RemoveRootDirectory(part.NewFullFileName), numTabs:=tabIndex, debugWrite:=False)
            _form.Log("Number of Duplicate Occurrences: " & part.DuplicateOccurrences.Count, numTabs:=tabIndex + 1, debugWrite:=False)
            _form.Log("Copy Enabled: " & part.CopyEnabled.ToString(), numTabs:=tabIndex + 1, debugWrite:=False)
        End If
    End Sub

    Private Sub LogFrameAssembly(ByRef parentFrameObj As InvtFrame, ByRef tabIndex As Integer)
        _form.Log("*FRAME ASSEMBLY*", numTabs:=tabIndex, numLinesBefore:=1, debugWrite:=False)
        _form.Log("Original Frame Name: " & parentFrameObj.OriginalName, numTabs:=tabIndex, debugWrite:=False)
        _form.Log("New Frame Name: " & parentFrameObj.NewName, numTabs:=tabIndex, debugWrite:=False)
        _form.Log("Original Frame Assembly Full File Name: " & RemoveRootDirectory(parentFrameObj.OriginalFullFileName), numTabs:=tabIndex, debugWrite:=False)
        _form.Log("New Frame Assembly Full File Name: " & RemoveRootDirectory(parentFrameObj.NewFullFileName), numTabs:=tabIndex, debugWrite:=False)
        _form.Log("Original Frame Skeleton ID: " & parentFrameObj.OriginalSkeletonID, numTabs:=tabIndex, debugWrite:=False)
        _form.Log("New Frame Skeleton ID: " & parentFrameObj.NewSkeletonID, numTabs:=tabIndex, debugWrite:=False)
        _form.Log("Number of Duplicate Occurrences: " & parentFrameObj.DuplicateOccurrences.Count, numTabs:=tabIndex + 1, debugWrite:=False)
        _form.Log("Copy Enabled: " & parentFrameObj.CopyEnabled.ToString(), numTabs:=tabIndex + 1, debugWrite:=False)
        _form.Log("***Old Skeleton Attribute Value: ", numTabs:=tabIndex, debugWrite:=False)
        _form.Log(parentFrameObj.OriginalSkeletonAttributeValue, numTabs:=tabIndex, debugWrite:=False)
        _form.Log("***New Skeleton Attribute Value: ", numTabs:=tabIndex, debugWrite:=False)
        _form.Log(parentFrameObj.NewSkeletonAttributeValue, numTabs:=tabIndex, debugWrite:=False)

        'increase the tab index for subcomponents
        tabIndex += 1

        For Each part In parentFrameObj.PartList
            LogPart(part, tabIndex)
        Next
        For Each subAsm In parentFrameObj.AssemblyList
            LogAssembly(subAsm, tabIndex)
        Next
        For Each subFrame In parentFrameObj.FrameList
            LogFrameAssembly(subFrame, tabIndex)
        Next
    End Sub

    ''' <summary>
    ''' Removes the _NewRootDirectory from component paths to make printing and logging clearer
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    Private Function RemoveRootDirectory(ByRef path As String) As String
        Dim newPath As String = "{PROJECT DIRECTORY}\" & path.Substring(_NewRootDirectory.Length)
        Return newPath
    End Function

#End Region

End Module

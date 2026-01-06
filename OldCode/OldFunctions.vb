Module OldFunctions
    Private Sub UpdateMainAssemblyNewName()
        Dim newRootDirectory As String = TB_newDir.Text
        If newRootDirectory = "" Then
            'this happens on initial load
        Else
            Dim newName As String = TB_Prefix.Text & TB_NewAssemblyName.Text & TB_Suffix.Text
            Dim dirString As String = newRootDirectory
            'you have to do it twice to get rid of the old file name
            dirString = dirString.Substring(0, dirString.LastIndexOf("\"))
            dirString = dirString.Substring(0, dirString.LastIndexOf("\") + 1) & newName & "\"
            newRootDirectory = dirString
            TB_newDir.Text = newRootDirectory
            mainAsmObj.NewName = newName
            mainAsmObj.NewTreeNode.Text = newName
            mainAsmObj.NewFilePath = newRootDirectory
            ResetCarets()
        End If

    End Sub

    Sub SetupTreeView(ByRef treeView As System.Windows.Forms.TreeView, ByRef asmObj As InvtAssemblyObj, ByRef newAsm As Boolean)
        treeView.Nodes.Clear()
        Dim rootNode As TreeNode
        If newAsm Then
            rootNode = treeView.Nodes.Add(asmObj.NewName)
            asmObj.NewTreeNode = rootNode
        Else
            rootNode = treeView.Nodes.Add(asmObj.OriginalName)
        End If
        AddSubNodes(rootNode, asmObj, newAsm)
        treeView.ExpandAll()
    End Sub

    Sub AddSubNodes(ByRef parentNode As TreeNode, ByRef asmObj As InvtAssemblyObj, ByRef newAsm As Boolean)
        Dim newNode As TreeNode
        For Each comp As InvtComponentObj In asmObj.AssemblyComponents
            'newNode = parentNode.Nodes.Add(comp.Name)
            If comp.Type = "Assembly" Then
                If newAsm Then
                    newNode = parentNode.Nodes.Add(comp.AssemblyObject.NewName)
                    comp.AssemblyObject.NewTreeNode = newNode
                Else
                    newNode = parentNode.Nodes.Add(comp.AssemblyObject.OriginalName)
                End If

                AddSubNodes(newNode, comp.AssemblyObject, newAsm)

            ElseIf comp.Type = "Part" Then
                If newAsm Then
                    newNode = parentNode.Nodes.Add(comp.PartObject.NewName)
                    comp.PartObject.NewTreeNode = newNode
                Else
                    newNode = parentNode.Nodes.Add(comp.PartObject.OriginalName)
                End If

            End If
        Next
    End Sub

    Private Sub ReplaceFrame(ByRef frmOcc As ComponentOccurrence)
        Debug.WriteLine("Replacing Frame Assembly: " & frmOcc.Name)

        'replace the old skeleton id with a new one
        Dim nSkelId As String = Nothing

        'replace the skelton id in the frame assembly attributes
        For Each attSet As AttributeSet In frmOcc.Definition.AttributeSets
            For Each atri As Attribute In attSet
                If atri.Name = "Frame.Skeletons" Then
                    Dim oAtriVal As String = atri.Value
                    Dim skelIdStart As Integer = GetSkelIdStartInt(oAtriVal)
                    Dim skelIdEnd As Integer = GetSkelIdEndInt(oAtriVal, skelIdStart)

                    Dim oSkelId As String = oAtriVal.Substring(skelIdStart, skelIdEnd - skelIdStart)

                    nSkelId = GenerateNewSkelId(oSkelId)

                    Dim nAtriVal As String = oAtriVal.Substring(0, skelIdStart) & nSkelId &
                        oAtriVal.Substring(skelIdEnd)

                    atri.Value = nAtriVal
                End If
            Next
        Next

        Dim subAsyOccs As ComponentOccurrences = frmOcc.Definition.Occurrences
        'replace the parts in frame assembly
        If prtList.Count > 0 Then
            For Each part As InvtPartObj In prtList
                If part.SubType IsNot "Content Center Part" Then
                    'This replaces all occurrences so no need to replace duplicates separately"
                    'Debug.WriteLine("Replacing Frame part: " & part.OriginalComponentOccurence.Name & " with " & part.NewFullFileName)
                    _form.LB_CopyComplete.Text = "Replacing: " & part.OriginalComponentOccurrence.Name & " with " & part.NewFullFileName
                    Dim curOcc As ComponentOccurrence
                    Try
                        curOcc = subAsyOccs.ItemByName(part.OriginalComponentOccurrence.Name)
                    Catch ex As Exception
                        Debug.WriteLine("Could not find occurrence by name: " & part.OriginalComponentOccurrence.Name & "; trying by document full name.")
                        _form.Log("Could not find occurrence by name: " & part.OriginalComponentOccurrence.Name & "; trying by document full name.")
                        curOcc = FindOccurrenceByDocumentFullName(subAsyOccs, part.OriginalFullFileName)
                    End Try
                    'curOcc.Replace(part.NewFullFileName, True)
                    ComponentReplace(curOcc, part.NewFullFileName)

                    'we need to update the part number in the iProperties of components that have a new component name
                    If part.OriginalName IsNot part.NewName Then
                        Dim replacedPartDoc As PartDocument = _invApp.Documents.ItemByName(part.NewFullFileName)
                        replacedPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = part.NewName
                        curOcc.Name = part.NewName
                    End If
                Else
                    Debug.WriteLine("Skipping Content Center Part: " & part.OriginalName)
                End If

            Next
        End If

        'replace frame sub-assemblies
        If subAsyList.Count > 0 Then
            For Each subAsy As AssemblyCopyObject In subAsyList
                'get the occurence of the current subAsy by searching for it by name using the original occurence name


                'recall this sub by getting the occurence of the component to be replaced by 
                _form.LB_CopyComplete.Text = "Replacing: " & subAsy.OriginalName & " with " & subAsy.NewFullFileName
                Dim curOcc As ComponentOccurrence
                Try
                    curOcc = subAsyOccs.ItemByName(subAsy.OriginalComponentOccurrence.Name)
                Catch ex As Exception
                    Debug.WriteLine("Could not find occurrence by name: " & subAsy.OriginalComponentOccurrence.Name & "; trying by document full name.")
                    _form.Log("Could not find occurrence by name: " & subAsy.OriginalComponentOccurrence.Name & "; trying by document full name.")
                    curOcc = FindOccurrenceByDocumentFullName(subAsyOccs, subAsy.OriginalFullFileName)
                End Try
                'curOcc.Replace(subAsy.NewFullFileName, True)
                ComponentReplace(curOcc, subAsy.NewFullFileName)

                'we need to update the part number in the iProperties of components that have a new component name
                If subAsy.OriginalName IsNot subAsy.NewName Then
                    Dim replacedAsyDoc As AssemblyDocument = _invApp.Documents.ItemByName(subAsy.NewFullFileName)
                    replacedAsyDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = subAsy.NewName
                    curOcc.Name = subAsy.NewName
                End If

                If subAsy.SubType = "Frame" Then
                    subAsy.ReplaceFrame(curOcc)
                End If
                subAsy.ReplaceOccurences(curOcc)
            Next
        End If

        'find the skeleton occurence so we can replace the id
        Dim skelOcc As ComponentOccurrence = GetSkeletonOcc(frmOcc.Definition.Occurrences)
        For Each attSet As AttributeSet In skelOcc.AttributeSets
            For Each att As Attribute In attSet
                ' replace the old skeleton id with the new
                If att.Name = "ID" Then
                    att.Value = nSkelId
                End If
            Next
        Next

    End Sub
End Module

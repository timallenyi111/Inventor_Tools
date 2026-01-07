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

    ''' <summary>
    ''' Steps through assemblies and initiates the replacement of components
    ''' (ComponentReplace) handles the actual replacement
    ''' </summary>
    ''' <param name="asyOcc"></param>
    ''' <param name="skelId"></param>
    Sub ReplaceOccurences(Optional ByRef asyOcc As ComponentOccurrence = Nothing)

        If asyOcc Is Nothing Then
            'this is the root assembly
            'create name value map of options for opening the root assembly
            Dim nameValueMap As Inventor.NameValueMap = _invApp.TransientObjects.CreateNameValueMap
            nameValueMap.Add("SkipAllUnresolvedFiles", True)

            ' we need to open the new assembly
            Dim newAsmDoc As Inventor.AssemblyDocument = _invApp.Documents.OpenWithOptions(nFullFileName, nameValueMap, True)

            'assign the assembly occurrence to the root occurrence
            asyOcc = newAsmDoc.ComponentDefinition.Occurrences
        End If

        Dim curAsyOccs As ComponentOccurrences = asyOcc.Definition.Occurrences

        'replace the parts in sub-assemblies
        If prtList.Count > 0 Then
            For Each part As InvtPartObj In prtList
                'skip content center parts and parts that are not enabled for copy
                If part.SubType IsNot "Content Center Part" And part.CopyEnabled = True Then
                    'This replaces all occurrences so no need to replace duplicates separately                    
                    Dim curOcc As ComponentOccurrence
                    Try
                        curOcc = curAsyOccs.ItemByName(part.OriginalComponentOccurrence.Name)
                    Catch ex As Exception
                        _form.Log("Could Not find occurrence by name: " & part.OriginalComponentOccurrence.Name & "; trying by document full name.")
                        curOcc = FindOccurrenceByDocumentFullName(curAsyOccs, part.OriginalFullFileName)
                    End Try

                    ComponentReplace(curOcc, part.NewFullFileName)

                    'we need to update the part number in the iProperties of components that have a new component name
                    If part.OriginalName IsNot part.NewName Then
                        Dim replacedPartDoc As PartDocument = _invApp.Documents.ItemByName(part.NewFullFileName)
                        replacedPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = part.NewName
                        curOcc.Name = part.NewName
                    End If
                End If
            Next
        End If

        'replace sub-assemblies of sub-assemblies
        If subAsyList.Count > 0 Then
            For Each subAsy As AssemblyCopyObject In subAsyList
                'only replace sub-assemblies that are enabled for copy
                If subAsy.CopyEnabled = True Then
                    'get the occurence of the current subAsy by searching for it by name using the original occurence name
                    Dim curOcc As ComponentOccurrence
                    Try
                        curOcc = curAsyOccs.ItemByName(subAsy.OriginalComponentOccurrence.Name)
                    Catch ex As Exception
                        _form.Log("Could not find occurrence by name: " & subAsy.OriginalComponentOccurrence.Name & "; trying by document full name.")
                        curOcc = FindOccurrenceByDocumentFullName(curAsyOccs, subAsy.OriginalFullFileName)
                    End Try

                    ComponentReplace(curOcc, subAsy.NewFullFileName)

                    'we need to update the part number in the iProperties of components that have a new component name
                    If subAsy.OriginalName IsNot subAsy.NewName Then
                        Dim replacedAsyDoc As AssemblyDocument = _invApp.Documents.ItemByName(subAsy.NewFullFileName)
                        replacedAsyDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = subAsy.NewName
                        curOcc.Name = subAsy.NewName
                    End If

                    If subAsy.SubType = "Frame" Then
                        subAsy.ReplaceFrame(curOcc)
                    Else
                        subAsy.ReplaceOccurences(curOcc)
                    End If
                End If
            Next
        End If

        'save the document once all replacements are done
        _invApp.ActiveDocument.Save2()

    End Sub

    ''' <summary>
    ''' Performs the actual replacing of a subassembly in an assembly
    ''' </summary>
    ''' <param name="origOcc"></param>
    ''' <param name="newFileName"></param>
    Private Sub ComponentReplace_OLD(ByVal origOcc As ComponentOccurrence, ByVal newFileName As String)

        Debug.WriteLine("Replacing " & origOcc.Name & " with: " & newFileName)
        _form.LB_CopyComplete.Text = "Replacing: " & origOcc.Name & " with " & newFileName

        If String.IsNullOrWhiteSpace(newFileName) Then
            Debug.WriteLine("Replacement filename empty; skipping.")
            Return
        End If

        If Not System.IO.File.Exists(newFileName) Then
            Debug.WriteLine("Replacement file does not exist: " & newFileName)
            _form.Log("Replacement file missing: " & newFileName)
            Return
        End If

        ' Try initial Replace without opening the document
        Try
            origOcc.Replace(newFileName, True)
            Return
        Catch ex As Exception
            _form.Log("Initial Replace failed: " & ex.Message)
            Debug.WriteLine("Initial Replace failed: " & ex.Message)
        End Try

        ' If initial Replace failed, attempt to open replacement document (if not already open) then retry
        Try
            If Not IsDocumentOpenByFullName(newFileName) Then
                Try
                    _invApp.Documents.Open(newFileName)
                    _form.Log("Opened replacement document: " & newFileName)
                    Debug.WriteLine("Opened replacement document: " & newFileName)
                    System.Threading.Thread.Sleep(200)
                Catch exOpen As Exception
                    _form.Log("Failed to open replacement document: " & exOpen.Message)
                    Debug.WriteLine("Failed to open replacement document: " & exOpen.Message)
                End Try
            Else
                _form.Log("Replacement document already open: " & newFileName)
                Debug.WriteLine("Replacement document already open: " & newFileName)
            End If
        Catch ex As Exception
            _form.Log("Error checking/opening replacement document: " & ex.Message)
            Debug.WriteLine("Error checking/opening replacement document: " & ex.Message)
        End Try

        ' Retry Replace with limited attempts and delay
        Dim attempts As Integer = 0
        Dim maxAttempts As Integer = 3
        Dim replaced As Boolean = False

        While attempts < maxAttempts AndAlso Not replaced
            Try
                origOcc.Replace(newFileName, True)
                replaced = True
            Catch ex As Exception
                attempts += 1
                Debug.WriteLine("Replace retry " & attempts.ToString() & " failed for " & newFileName & ": " & ex.Message)
                If attempts < maxAttempts Then
                    System.Threading.Thread.Sleep(500)
                End If
            End Try
        End While

        If Not replaced Then
            Debug.WriteLine("All Replace attempts failed for: " & newFileName)
            _form.Log("Replace failed for: " & newFileName)
        End If


    End Sub


    ' Find first occurrence in a collection by matching the component document file name
    Private Function FindOccurrenceByDocumentFullName(occurrences As ComponentOccurrences, targetFullName As String) As ComponentOccurrence
        Debug.WriteLine("Looking for: " & targetFullName)
        For Each occ As ComponentOccurrence In occurrences
            Debug.WriteLine(vbTab & occ.Definition.Document.FullFileName)
            Try
                If String.Compare(occ.Definition.Document.FullFileName, targetFullName, StringComparison.OrdinalIgnoreCase) = 0 Then
                    Return occ
                End If
            Catch ex As Exception
                Debug.WriteLine("Match Not Found..")
                ' ignore inaccessible occ or continue
            End Try
        Next
        Return Nothing
    End Function

    Sub AssignNodeTags()
        For Each part As InvtPartObj In prtList
            'parts in the root assembly
            Dim partNode As System.Windows.Forms.TreeNode = part.NewTreeNode
            'in the root assembly the component occurrence is the only thing you need for highlighting
            Dim occList As New List(Of Inventor.ComponentOccurrence)
            Dim occNames As New List(Of String) From {
                part.OriginalComponentOccurrence.Name
            }
            'Debug.WriteLine("Adding original occurrence name to search list: " & part.OriginalComponentOccurence.Name)
            If part.DuplicateOccurrences.Count > 0 Then
                For Each dupOcc As Inventor.ComponentOccurrence In part.DuplicateOccurrences
                    occNames.Add(dupOcc.Name)
                    'Debug.WriteLine("Adding duplicate occurrence name to search list: " & dupOcc.Name)
                Next
            End If

            For Each occName As String In occNames
                For Each occ As Inventor.ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
                    If occ.Name = occName Then
                        occList.Add(occ)
                        'Debug.WriteLine("Found matching occurrence proxy: " & occ.Name)
                        Exit For
                    End If
                Next
            Next

            partNode.Tag = occList
        Next

        For Each subAsy As AssemblyCopyObject In subAsyList
            'sub-assemblies in the root assembly
            Dim subAsmNode As System.Windows.Forms.TreeNode = subAsy.NewTreeNode
            'assemblies in the root assembly need a list of occurrences for highlighting
            Dim occList As New List(Of Inventor.ComponentOccurrence)
            Dim occNames As New List(Of String) From {
                subAsy.OriginalComponentOccurrence.Name
            }
            'Debug.WriteLine("Adding original occurrence name to search list: " & subAsy.OriginalComponentOccurrence.Name)
            If subAsy.DuplicateOccurrences.Count > 0 Then
                For Each dupOcc As Inventor.ComponentOccurrence In subAsy.DuplicateOccurrences
                    occNames.Add(dupOcc.Name)
                    'Debug.WriteLine("Adding duplicate occurrence name to search list: " & dupOcc.Name)
                Next
            End If

            For Each occName As String In occNames
                For Each occ As Inventor.ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
                    If occ.Name = occName Then
                        occList.Add(occ)
                        'Debug.WriteLine("Found matching occurrence proxy: " & occ.Name)
                        'now process the components in the subassembly
                        SubAssemblyNodeTagSetup(occ.SubOccurrences, subAsy)
                        Exit For
                    End If
                Next
            Next
            subAsmNode.Tag = occList
        Next
    End Sub

    Private Sub SubAssemblyNodeTagSetup(ByRef occurrences As Inventor.ComponentOccurrences, ByVal subAsy As AssemblyCopyObject)

        For Each part As InvtPartObj In subAsy.prtList
            'setup part name list
            Dim occNames As New List(Of String) From {
                part.OriginalComponentOccurrence.Name
            }
            If part.DuplicateOccurrences.Count > 0 Then
                For Each dupOcc As Inventor.ComponentOccurrence In part.DuplicateOccurrences
                    occNames.Add(dupOcc.Name)
                Next
            End If

            Dim index As Integer = 1
            Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
            For Each occName As String In occNames
                While index <= occurrences.Count
                    Dim occ As Inventor.ComponentOccurrenceProxy = occurrences.Item(index)
                    If occ.Name = occName Then
                        'Debug.WriteLine("Found matching occurrence proxy: " & occ.Name)
                        occProxyList.Add(occ)
                        Exit While
                    End If
                    index += 1
                End While
            Next
            part.NewTreeNode.Tag = occProxyList
        Next

        For Each asy As AssemblyCopyObject In subAsy.subAsyList
            Dim occNames As New List(Of String) From {
                asy.OriginalComponentOccurrence.Name
            }
            If asy.DuplicateOccurrences.Count > 0 Then
                For Each dupOcc As Inventor.ComponentOccurrence In asy.DuplicateOccurrences
                    occNames.Add(dupOcc.Name)
                Next
            End If

            Dim index As Integer = 1
            Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
            For Each occName As String In occNames
                While index <= occurrences.Count
                    Dim occ As Inventor.ComponentOccurrenceProxy = occurrences.Item(index)
                    If occ.Name = occName Then
                        Debug.WriteLine("Found matching occurrence proxy: " & occ.Name)
                        occProxyList.Add(occ)
                        'Process components in the sub-assembly
                        'This will work for duplicate components because the name "component:1" will be the same for all duplicates
                        SubAssemblyNodeTagSetup(occ.SubOccurrences, asy)
                        Exit While
                    End If
                    index += 1
                End While
            Next

            asy.NewTreeNode.Tag = occProxyList

        Next


    End Sub




End Module

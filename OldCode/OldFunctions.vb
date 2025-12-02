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
End Module

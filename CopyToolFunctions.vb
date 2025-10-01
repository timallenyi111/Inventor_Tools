Imports System.IO
Module CopyToolFunctions

    Sub SetupFilePaths(ByRef invtAsmObj As InvtAssemblyObj, ByRef newRootDirectory As String)
        invtAsmObj.NewFilePath = newRootDirectory
        invtAsmObj.NewName = invtAsmObj.NewTreeNode.Text
        DB("New Assembly: " & invtAsmObj.NewName)
        DB("Old" & invtAsmObj.OriginalFullFileName)
        DB("New" & invtAsmObj.NewFullFileName)
        DB(" ")
        For Each comp As InvtComponentObj In invtAsmObj.AssemblyComponents
            If comp.Type = "Part" AndAlso comp.PartObject IsNot Nothing Then
                Dim curPrtObj As InvtPartObj = comp.PartObject
                curPrtObj.NewName = comp.PartObject.NewTreeNode.Text
                curPrtObj.NewFilePath = newRootDirectory
            ElseIf comp.Type = "Assembly" AndAlso comp.AssemblyObject IsNot Nothing Then
                Dim curAsmObj As InvtAssemblyObj = comp.AssemblyObject
                SetupFilePaths(curAsmObj, invtAsmObj.NewFilePath)
            End If
        Next
    End Sub

    Sub CopyAssemblyFile(ByRef invtAsmObj As InvtAssemblyObj)

        If Directory.Exists(invtAsmObj.NewFilePath) = False Then
            Directory.CreateDirectory(invtAsmObj.NewFilePath)
        End If
        File.Copy(invtAsmObj.OriginalFullFileName, invtAsmObj.NewFullFileName, False)
        'For now all new components will have the same name as the original
        For Each comp As InvtComponentObj In invtAsmObj.AssemblyComponents
            If comp.Type = "Part" AndAlso comp.PartObject IsNot Nothing Then
                Dim curPrtObj As InvtPartObj = comp.PartObject

                If Directory.Exists(curPrtObj.NewFilePath) = False Then
                    Directory.CreateDirectory(curPrtObj.NewFilePath)
                End If
                File.Copy(curPrtObj.OriginalFullFileName, curPrtObj.NewFullFileName, False)
            ElseIf comp.Type = "Assembly" AndAlso comp.AssemblyObject IsNot Nothing Then
                Dim curAsmObj As InvtAssemblyObj = comp.AssemblyObject
                ' curAsmObj.NewName = curAsmObj.OriginalName
                ' curAsmObj.NewFilePath = rootDirectory
                CopyAssemblyFile(curAsmObj)
            End If
        Next

    End Sub

    Sub ReplaceAssemblyComponents(ByRef invtAsmObj As InvtAssemblyObj, ByRef AsmDoc As Inventor.AssemblyDocument, Optional ByRef mainAssembly As Boolean = False)
        Dim compOccs As Inventor.ComponentOccurrences = AsmDoc.ComponentDefinition.Occurrences

        If mainAssembly Then
            ' Nothing for now because the main assembly can't replace itself
            DB("Replacing components in :")
            DB(AsmDoc.FullFileName.ToString)
            DB(" ")
        Else
            'DB("Replacing Assembly: ")
            'DB(invtAsmObj.OriginalFullFileName)
            'DB("with")
            'DB(invtAsmObj.NewFullFileName)
            'DB(" ")

        End If

        Dim replacedOccList As New List(Of String)
        For Each occ As Inventor.ComponentOccurrence In compOccs
            If occ.Definition.Type = Inventor.ObjectTypeEnum.kAssemblyComponentDefinitionObject Then
                Dim oSubAsmDoc As Inventor.AssemblyDocument = occ.Definition.Document
                If replacedOccList.Contains(oSubAsmDoc.FullFileName) Then
                    ' occurence already replaced move on
                Else

                    DB("Replacing Assembly: ")
                    DB(oSubAsmDoc.FullFileName.ToString)

                    Dim subAsmObj As InvtAssemblyObj = invtAsmObj.GetComponentByOriginalFileName(oSubAsmDoc.FullFileName).AssemblyObject

                    DB("with")
                    DB(subAsmObj.NewFullFileName)
                    DB(" ")

                    occ.Replace(subAsmObj.NewFullFileName, True)
                    Dim nSubAsmDoc As Inventor.AssemblyDocument = occ.Definition.Document
                    ReplaceAssemblyComponents(subAsmObj, nSubAsmDoc)
                    replacedOccList.Add(subAsmObj.NewFullFileName)
                End If

            ElseIf occ.Definition.Type = Inventor.ObjectTypeEnum.kPartComponentDefinitionObject Then
                Dim prtDoc As Inventor.PartDocument = occ.Definition.Document
                If replacedOccList.Contains(prtDoc.FullFileName) Then
                    ' occurence already replaced move on
                Else
                    DB("Replacing Part: ")
                    DB(prtDoc.FullFileName.ToString)
                    DB("In Assembly:")
                    DB(AsmDoc.FullFileName.ToString)

                    Dim prtObj As InvtPartObj = invtAsmObj.GetComponentByOriginalFileName(prtDoc.FullFileName).PartObject

                    DB("with")
                    DB(prtObj.NewFullFileName)
                    DB(" ")

                    occ.Replace(prtObj.NewFullFileName, True)
                    replacedOccList.Add(prtDoc.FullFileName.ToString)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Writes text in a textbox and scrolls to the end
    ''' </summary>
    ''' <param name="textBox"></param>
    ''' <param name="msg"></param>
    Sub LongTextboxWrite(ByRef textBox As System.Windows.Forms.TextBox, ByRef msg As String)
        textBox.Text = msg
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub

    Sub MoveCaret(ByRef textBox As System.Windows.Forms.TextBox)
        textBox.SelectionStart = 0
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
        textBox.SelectionStart = textBox.Text.Length
        textBox.SelectionLength = 0
        textBox.ScrollToCaret()
    End Sub


End Module

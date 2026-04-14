Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Window
Imports Inventor

Module CopyAndReplaceFunctions
    '*COPY/REPLACE FILE STRATEGY
    'update new properties
    'copy all parts
    'copy all assemblies that don't have a frame
    'open frameObj assemblies 
    'change the frameObj skeleton id and save COPY as the new file name
    'change the frameObj skeleton id back, save, and close the original frameObj assembly
    'open the new root assembly
    'replace all assemblies
    'replace all parts

    Private _form As AssemblyCopyToolForm
    Private _app As Inventor.Application

    ''' <summary>
    ''' Coordinates the process of copying all of the files and then replacing the occurrences in the new assembly
    ''' </summary>
    ''' <param name="rootAssemblyObject"></param>
    Public Sub CopyAndReplace(ByRef rootAssemblyObject As InvtAssembly, ByRef form As AssemblyCopyToolForm, app As Inventor.Application)

        _form = form
        _app = app
        UpdateComponentProperties(rootAssemblyObject)

        LogAssembly(rootAssemblyObject, isRoot:=True, startingMessage:="*****UPDATED PROPERTIES*****")

        'used if the user cancels the copy process because they didn't want to overwrite existing files
        Dim processCancled As Boolean = False

        'copy the root assembly document
        _app.ActiveDocument.SaveAs(rootAssemblyObject.NewFullFileName, True)

        'close the original document to prevent errors during frame replacement
        _app.ActiveDocument.Save2()
        _app.ActiveDocument.Close(True)

        _form.Log("****STARTING TO COPY FILES****", numLinesBefore:=3)
        'open the new root assembly
        Dim newRootAsmDoc As AssemblyDocument = OpenAssemblyDocument(rootAssemblyObject.NewFullFileName, True)

        'copy all of the necessary files
        processCancled = CreateNewFiles(rootAssemblyObject)
        _form.Log("*****FILE COPY COMPLETE*****", numLinesAfter:=3)

        _form.Log("*****STARTING TO REPLACE FILES*****", numLinesBefore:=3)
        'replace the component occurrences in the assembly
        ReplaceComponents(rootAssemblyObject, newRootAsmDoc.ComponentDefinition.Occurrences)
        _form.Log("*****REPLACE COMPLETE*****", numLinesBefore:=3)

        newRootAsmDoc.Save()

        MessageBox.Show("Copy Complete")

    End Sub

    ''' <summary>
    ''' 'updates the component properties of the Invt Components that have been changed by the user
    ''' since the initial setup during load.
    ''' (HANDLES ASSEMBLIES, FRAMES, AND PARTS)
    ''' </summary>
    Private Sub UpdateComponentProperties(ByRef invtObject As Object)
        Dim tNode As TreeNode = invtObject.TreeNode
        'tree node forecolor gets changed to red if the user selects the "do not copy" option
        If tNode.ForeColor = System.Drawing.Color.Red Then
            invtObject.CopyEnabled = False
        End If

        'when the user changes the assembly name by changing the treenode text
        invtObject.NewName = tNode.Text

        If TypeOf (invtObject) Is InvtPart Then
            'do nothing because parts don't have subcomponents
        Else
            'update sub-components
            For Each part As InvtPart In invtObject.PartList
                UpdateComponentProperties(part)
            Next
            For Each asm As InvtAssembly In invtObject.AssemblyList
                UpdateComponentProperties(asm)
            Next
            For Each frm As InvtFrame In invtObject.FrameList
                UpdateComponentProperties(frm)
            Next
        End If

    End Sub

    ''' <summary>
    ''' Copies all of the files that are supposed to be copied to their respective directory. If the file alread exist
    ''' a message box will be shown to the user to be able to replace the existing files or cancel the operation.
    ''' </summary>
    ''' <param name="parentAssembly"></param>
    ''' <param name="isRoot"></param>
    ''' <returns>True if the user canceled the operation to avoid overwriting existing files</returns>
    Private Function CreateNewFiles(ByRef parentAssembly As InvtAssembly) As Boolean
        'used to end the copy process if the user decides to no overwrite files
        Dim endProcess As Boolean = False

        If endProcess Then
            Return endProcess
        End If

        For Each part As InvtPart In parentAssembly.PartList
            If part.CopyEnabled Then
                If part.IsContentCenter Then
                    'we don't copy content center parts
                    _form.Log(part.OriginalName & " skipped because it's a content center part", numLinesBefore:=1)
                Else
                    _form.Log("Copying: " & part.OriginalFullFileName, numLinesBefore:=1)
                    _form.Log("To: " & part.NewFullFileName)
                    endProcess = CopyFile(part.OriginalFullFileName, part.NewFullFileName)
                End If
                If endProcess Then
                    Return endProcess
                End If
            Else
                _form.Log(part.OriginalName & " skipped because copy is disabled*", numLinesBefore:=1)
            End If

        Next

        For Each subAsm As InvtAssembly In parentAssembly.AssemblyList
            If subAsm.CopyEnabled Then
                If subAsm.ContainsFrame Then
                    CopyFrameParent(subAsm)
                Else
                    _form.Log("Copying: " & subAsm.OriginalFullFileName, numLinesBefore:=1)
                    _form.Log("To: " & subAsm.NewFullFileName)
                    'make a copy of this sub assembly
                    endProcess = CopyFile(subAsm.OriginalFullFileName, subAsm.NewFullFileName)
                    'send the sub assembly to have all of it's components copied
                    endProcess = CreateNewFiles(subAsm)
                    If endProcess Then
                        Return endProcess
                    End If
                End If
            Else
                _form.Log(subAsm.OriginalName & " skipped because copy is disabled*", numLinesBefore:=1)
            End If
        Next

        For Each subFrm As InvtFrame In parentAssembly.FrameList
            If subFrm.CopyEnabled Then
                If subFrm.ContainsFrame Then
                    CopyFrameParent(subFrm.CoreAssemblyObject)
                Else
                    _form.Log("Copying: " & subFrm.OriginalFullFileName, numLinesBefore:=1)
                    _form.Log("To: " & subFrm.NewFullFileName)
                    'save a copy of the sub frame
                    endProcess = CopyFile(subFrm.OriginalFullFileName, subFrm.NewFullFileName)
                    'send the sub frame to have all of it's components copied
                    endProcess = CreateNewFiles(subFrm.CoreAssemblyObject)
                    If endProcess Then
                        Return endProcess
                    End If
                End If
            Else
                _form.Log(subFrm.OriginalName & " skipped because copy is disabled*", numLinesBefore:=1)
            End If
        Next

        Return endProcess
    End Function

    ''' <summary>
    ''' Copies the input files from the original file name to the new file name.
    ''' </summary>
    ''' <param name="oFile"></param>
    ''' <param name="nFile"></param>
    ''' <returns>True if the user has canceled the operation to avoid overwriting files</returns>
    Private Function CopyFile(oFile As String, nFile As String) As Boolean
        'will be set to true if the user decides to cancel the copy operation
        Dim endProcess As Boolean = False

        _form.Log("copying file: " & oFile)
        Dim nFilePath As String = nFile.Substring(0, nFile.LastIndexOf("\"))
        _form.Log("New File Path: " & nFilePath, numTabs:=1)

        'check if the new directory exists and if it doesn't create one.
        If Directory.Exists(nFilePath) = False Then
            Directory.CreateDirectory(nFilePath)
            _form.Log("New Directory Created")
        End If

        'check if new file already exists, if so log it
        If System.IO.File.Exists(nFile) Then
            '_form.Log("!!!!!!! FILE SKIPPED BECAUSE IT ALREADY EXISTS !!!!!!!")
            Dim result As DialogResult
            result = MessageBox.Show(nFile & " Already Exists", "Yes = Continue and overwrite all existing files \n No = Cancel the copy operation",
                                     MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                System.IO.File.Copy(oFile, nFile, True)
            Else
                endProcess = True
            End If
        Else
            System.IO.File.Copy(oFile, nFile, False)
            _form.Log("COPY SUCCESSFUL", numTabs:=1, numLinesAfter:=1)
            _form.LB_CopyComplete.Text = "Saving File: " & nFile
        End If

        Return endProcess
    End Function

    Private Sub CopyFrameParent(ByVal parentAsmObj As InvtAssembly)
        'send the frame parent to have all of its components saved.
        CreateNewFiles(parentAsmObj)

        'used to close the file that is opened if this frame isn't in the root assembly
        Dim openedAsmFlag As Boolean = False
        Dim oParentAsmDoc As Inventor.AssemblyDocument = Nothing

        If _app.ActiveDocument.FullFileName = parentAsmObj.OriginalFullFileName Then
            'the frame is in the root directory so we don't need to open anything.
            oParentAsmDoc = _app.ActiveDocument
        Else
            _form.Log("Opening: " & parentAsmObj.OriginalFullFileName & "for frame copying", numLinesBefore:=1)
            oParentAsmDoc = OpenAssemblyDocument(parentAsmObj.OriginalFullFileName, False)

            'change the opened assembly flag to true so we can close it when we are finished.
            openedAsmFlag = True
        End If

        For Each frameObj As InvtFrame In parentAsmObj.FrameList

            Dim frameOcc As ComponentOccurrence = oParentAsmDoc.ComponentDefinition.Occurrences.Item(frameObj.OccurrenceIndex)
            'replace the frame skeleton id attribute
            For Each attSet As Inventor.AttributeSet In frameOcc.Definition.AttributeSets
                For Each atri As Inventor.Attribute In attSet
                    If atri.Name = "Frame.Skeletons" Then
                        atri.Value = frameObj.NewSkeletonAttributeValue
                        _form.Log("Skeleton Attribute Replaced")
                    End If
                Next
            Next

            Dim skelOcc As ComponentOccurrence = oParentAsmDoc.ComponentDefinition.Occurrences.Item(frameObj.OriginalSkeletonPart.OccurrenceIndex)

            'replace the skeleton occurence skeleton id
            For Each attSet As AttributeSet In skelOcc.AttributeSets
                For Each att As Attribute In attSet
                    ' replace the old skeleton id with the new
                    If att.Name = "ID" Then
                        _form.Log("Replacing Skeleton Occurrence ID", numLinesBefore:=1)
                        _form.Log("ID: " & frameObj.NewSkeletonID)
                        att.Value = frameObj.NewSkeletonID
                    End If
                Next
            Next

            'replace the skeleton occurrence
            _form.Log("Replacing Skeleton Occurrence:", numLinesBefore:=1)
            _form.Log("Original: " & frameObj.OriginalSkeletonPart.OriginalFullFileName)
            _form.Log("New: " & frameObj.OriginalSkeletonPart.NewFullFileName)
            skelOcc.Replace(frameObj.OriginalSkeletonPart.NewFullFileName, True)

            'save a copy of the original parent assembly document as the new document.
            _form.Log("Saving new frame parent assembly", numLinesBefore:=1)
            _form.Log(parentAsmObj.NewFullFileName)

            oParentAsmDoc.SaveAs(parentAsmObj.NewFullFileName, True)

            'change the skeleton id attribute back to the original
            For Each attSet As Inventor.AttributeSet In frameOcc.Definition.AttributeSets
                For Each atri As Inventor.Attribute In attSet
                    If atri.Name = "Frame.Skeletons" Then
                        atri.Value = frameObj.OriginalSkeletonAttributeValue
                        Debug.WriteLine("Skeleton Attribute Replaced")
                    End If
                Next
            Next

            'chnage the skeleton occurrence back
            _form.Log("Replacing Skeleton Occurrence Back to Original:", numLinesBefore:=1)
            _form.Log("Original: " & frameObj.OriginalSkeletonPart.OriginalFullFileName)
            _form.Log("New: " & frameObj.OriginalSkeletonPart.NewFullFileName)
            skelOcc.Replace(frameObj.OriginalSkeletonPart.OriginalFullFileName, True)

            'change the skeleton occurrence id back to the original
            For Each attSet As AttributeSet In skelOcc.AttributeSets
                For Each att As Attribute In attSet
                    ' replace the old skeleton id with the new
                    If att.Name = "ID" Then
                        att.Value = frameObj.OriginalSkeletonID
                        _form.Log("Replacing Skeleton Occurrence ID with Original", numLinesBefore:=1)
                        _form.Log("ID: " & frameObj.OriginalSkeletonID)
                    End If
                Next
            Next

            'save the original document after changing everything back to normal
            _form.Log("Saving Frame Parent Assembly Document", numLinesBefore:=1)
            oParentAsmDoc.Save2()
        Next

        If openedAsmFlag Then
            oParentAsmDoc.Save2()
            _form.Log("Closing Original Frame Parent Assembly Document:", numLinesBefore:=1)
            _form.Log(oParentAsmDoc.FullFileName)
            oParentAsmDoc.Close(True)
        End If

    End Sub

    Private Sub ReplaceComponents(ByRef parentAsmObj As InvtAssembly, ByRef parentCompOccs As Inventor.ComponentOccurrences, Optional ByRef isRoot As Boolean = False)

        Dim openedAsmFlag As Boolean = False

        If parentAsmObj.ContainsFrame Then
            _form.Log("Opening: " & parentAsmObj.NewFullFileName, numLinesBefore:=1)
            _form.Log("Because it contains a frame")
            'we need to open this file to replace the frame
            Dim nParentAsmDoc As AssemblyDocument = OpenAssemblyDocument(parentAsmObj.NewFullFileName, False)
            parentCompOccs = nParentAsmDoc.ComponentDefinition.Occurrences
            openedAsmFlag = True
        End If

        'replace the subassemblies first
        For Each subAsm As InvtAssembly In parentAsmObj.AssemblyList
            If subAsm.CopyEnabled Then
                _form.Log("Replacing: " & subAsm.OriginalName, numLinesBefore:=1)
                _form.Log("With: " & subAsm.NewName)

                'get the component occurrence that for the sub assembly
                Dim compOcc As Inventor.ComponentOccurrence = parentCompOccs.Item(subAsm.OccurrenceIndex)
                compOcc.Replace(subAsm.NewFullFileName, True)

                'now replace the components in the sub-assembly
                ReplaceComponents(subAsm, compOcc.Definition.Occurrences)
            End If
        Next

        For Each subFrm As InvtFrame In parentAsmObj.FrameList
            If subFrm.CopyEnabled Then
                ReplaceFrame(parentAsmObj, subFrm)
            End If
        Next

        For Each subPart As InvtPart In parentAsmObj.PartList
            If subPart.CopyEnabled Then
                If subPart.IsContentCenter Then
                    'we don't replace content center parts
                    _form.Log("Skipping " & subPart.OriginalName & "because it is a content center part", numLinesBefore:=1)
                Else
                    _form.Log("Replacing: " & subPart.OriginalName, numLinesBefore:=1)
                    _form.Log("With: " & subPart.NewName)
                    Dim compOcc As Inventor.ComponentOccurrence = parentCompOccs.Item(subPart.OccurrenceIndex)
                    compOcc.Replace(subPart.NewFullFileName, True)
                End If
            End If
        Next

        'close the parent assembly if it isn't the root
        If openedAsmFlag Then
            _form.Log("Closing " & _app.ActiveDocument.FullFileName, numLinesBefore:=1)

            _app.ActiveDocument.Save2()
            _app.ActiveDocument.Close(True)
        End If
    End Sub

    Private Sub ReplaceFrame(ByRef parentAsmObj As InvtAssembly, ByRef frameObj As InvtFrame)
        'we have to open the frame assembly
        _form.Log("Opening Frame Assembly: " & frameObj.NewFullFileName, numLinesBefore:=1)
        Dim frameDoc As AssemblyDocument = OpenAssemblyDocument(frameObj.NewFullFileName, False)

        'now send the frame document to have all of its components replaced
        _form.Log("Replacing Components in Frame Assembly")
        ReplaceComponents(frameObj.CoreAssemblyObject, frameDoc.ComponentDefinition.Occurrences)

        _form.Log("Closing Frame Assembly...")
        'save and close the frame document
        frameDoc.Save2()
        frameDoc.Close(True)

        _form.Log("Setting the active doc as the frame parent document assembly", numLinesBefore:=1)
        _form.Log("Active Doc = " & _app.ActiveDocument.FullFileName)
        'the parent document should be the active document after closing the frame because the replace components sub opens assemblies with frames
        Dim parentDoc As AssemblyDocument = _app.ActiveDocument

        _form.Log("replacing the frame occurrence in the frame parent assembly document")
        'now replace the frame occurrence
        Dim frameOcc As ComponentOccurrence = parentDoc.ComponentDefinition.Occurrences.Item(frameObj.OccurrenceIndex)
        frameOcc.Replace(frameObj.NewFullFileName, True)
    End Sub

    Private Function OpenAssemblyDocument(ByRef fileName As String, ByRef visibility As Boolean) As Inventor.AssemblyDocument
        Dim nameValueMap As Inventor.NameValueMap = _app.TransientObjects.CreateNameValueMap
        nameValueMap.Add("SkipAllUnresolvedFiles", True)
        Dim newDoc As Inventor.AssemblyDocument = _app.Documents.OpenWithOptions(fileName, nameValueMap, True)
        Return newDoc
    End Function

End Module

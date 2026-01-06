Imports Inventor
Imports System.Windows.Forms
Imports System.IO




''' <summary>
''' This stores all of the information necessary to make a copy of an assembly
''' </summary>
Friend Class AssemblyCopyObject

    Private ReadOnly _form As AssemblyCopyToolForm
    Private ReadOnly _invApp As Inventor.Application
    Private _contentCenterPath As String
    Private prtList As List(Of InvtPartObj)
    Private subAsyList As List(Of AssemblyCopyObject)
    Private oAsyName As String
    Private nAsyName As String
    Private oFullFileName As String
    Private nFullFileName As String
    Private nRootDirectory As String
    Private oAsmDoc As AssemblyDocument
    Private nTreeNode As TreeNode
    Private oCompOcc As ComponentOccurrence
    Private _subType As String
    Private hltSet As HighlightSet
    Private ReadOnly duplicateOccurrenceList As List(Of ComponentOccurrence)
    Private _copyEnabled As Boolean = True
    Private _occurrenceIndex As Integer

    Public Sub New(form As AssemblyCopyToolForm, invApp As Inventor.Application)
        _form = form
        _invApp = invApp
        _contentCenterPath = _invApp.DesignProjectManager.ActiveDesignProject.ContentCenterPath

        prtList = New List(Of InvtPartObj)
        subAsyList = New List(Of AssemblyCopyObject)
        duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
    End Sub

#Region "setup functions"
    Sub InitialSetup(Optional asyOcc As ComponentOccurrence = Nothing, Optional rootDirectory As String = Nothing,
                     Optional nParentTreeNode As TreeNode = Nothing)

        'Check if the current assembly is the root assembly
        If asyOcc Is Nothing Then
            ' this is the root assembly         
            SetOriginalProperties(_invApp.ActiveDocument)
            ' define the root directory for the entire assembly
            Dim actProj As Inventor.DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
            Dim projectDir As String = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\") + 1)
            nRootDirectory = projectDir & _form.TB_Prefix.Text & oAsyName & _form.TB_Suffix.Text & "\"

            SetNewProperties()

            'create a highlight set to be used for highlighting model components selected in the treeview
            hltSet = _invApp.ActiveDocument.CreateHighlightSet()
        Else
            ' this is a subassembly
            SetOriginalProperties(asyOcc.Definition.Document, asyOcc)
            nRootDirectory = rootDirectory
            SetNewProperties(nParentTreeNode) ' sub assemblys don't automatically get the pre/suffix
            'hltSet = oAsmDoc.CreateHighlightSet()
        End If

        _form.Log("", numLines:=1)
        _form.Log("***** " & oAsyName & " component setup *****")

        Dim curOccIndex As Integer = 1

        'Perform the initial setup for all components in the assembly
        For Each curOcc As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences

            If curOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                'ReadOccurrenceDefinitionAttributes(curOcc)
                If CheckForDuplicateDocument(curOcc) = False Then
                    ' perform part setup
                    Dim curPartObject As New InvtPartObj
                    curPartObject.InitialSetup(curOcc, nRootDirectory, nTreeNode, _contentCenterPath)
                    curPartObject.OccurrenceIndex = curOccIndex
                    prtList.Add(curPartObject)
                Else
                    _form.Log(curOcc._DisplayName & " was a duplicate")
                End If

            ElseIf curOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                If CheckForDuplicateDocument(curOcc) = False Then
                    ' perform sub assembly setup

                    Dim curAsmObject As New AssemblyCopyObject(_form, _invApp)
                    If CheckIfOccurenceIsFrame(curOcc) Then
                        ' this is a frame assembly so we need to set the frame directory
                        Dim frameRootDirectory As String = nRootDirectory + nTreeNode.Text + "\Frame\"
                        curAsmObject.InitialSetup(curOcc, frameRootDirectory, nTreeNode)
                        curAsmObject.SubType = "Frame"
                        curAsmObject.OccurrenceIndex = curOccIndex
                    ElseIf CheckIfOccurrenceIsBoltedConnection(curOcc) Then
                        Dim boltedConnectionDirectory As String = nRootDirectory + nTreeNode.Text + "\Design Accelerator\"
                        curAsmObject.InitialSetup(curOcc, boltedConnectionDirectory, nTreeNode)
                        curAsmObject.SubType = "Bolted Connection"
                        curAsmObject.OccurrenceIndex = curOccIndex
                    Else
                        curAsmObject.InitialSetup(curOcc, nRootDirectory, nTreeNode)
                        curAsmObject.OccurrenceIndex = curOccIndex
                    End If
                    'Debug.WriteLine(curAsmObject.NewFullFileName)
                    subAsyList.Add(curAsmObject)
                Else
                    _form.Log(curOcc._DisplayName & " was a duplicate")
                End If
            Else
                'unknown component type
                _form.Log("Unknown component type found: " & curOcc._DisplayName)
            End If
            curOccIndex += 1
        Next

        'these node tags are used for highlighting the components in the assembly treeview
        AssignNodeTags()
    End Sub

    ''' <summary>
    ''' Sets up all the initial parameters for the original assembly file
    ''' </summary>
    ''' <param name="AsyOcc"></param>
    ''' <param name="ParentAssembly"></param>
    Sub SetOriginalProperties(ByRef AsyDoc As AssemblyDocument,
                              Optional ByRef AsyOcc As ComponentOccurrence = Nothing)

        oCompOcc = AsyOcc
        oAsmDoc = AsyDoc
        oFullFileName = oAsmDoc.FullFileName
        oAsyName = GetAssemblyName(oFullFileName)

        'if there is no parent occurence then it is the main assembly and so the first tree node has to be created
        If AsyOcc Is Nothing Then
            SubType = "Root"
        Else
            If CheckIfOccurenceIsFrame(oCompOcc) Then
                _subType = "Frame"
            End If
        End If
    End Sub

    Sub SetNewProperties(Optional ByRef nParentNode As TreeNode = Nothing)
        If nParentNode Is Nothing Then
            ' this is the root assembly

            nAsyName = _form.TB_Prefix.Text & oAsyName & _form.TB_Suffix.Text
            nFullFileName = nRootDirectory & nAsyName & ".iam"
            nTreeNode = New TreeNode(nAsyName)
        Else
            ' this is a subassembly
            If _subType = "Frame" Then
                nAsyName = GenerateNewFrameName(oAsyName)
                nFullFileName = nRootDirectory & nAsyName & ".iam"
                nTreeNode = nParentNode.Nodes.Add(nAsyName)
            Else
                nAsyName = oAsyName
                nFullFileName = nRootDirectory & nAsyName & ".iam"
                nTreeNode = nParentNode.Nodes.Add(nAsyName)
            End If

        End If
    End Sub

    Private Function GenerateNewFrameName(ByRef frameName As String) As String
        Dim nFrameName As String = "Frame "
        Dim rnd As New Random
        Dim count = 0
        While count < 13
            nFrameName = nFrameName + rnd.Next(0, 9).ToString
            count += 1
        End While
        Return nFrameName
    End Function

    ''' <summary>
    ''' returns an assembly name based on the original file name without the .iam
    ''' </summary>
    ''' <returns></returns>
    Private Function GetAssemblyName(ByRef fullFileName As String) As String
        Dim asyName As String = fullFileName.Substring(fullFileName.LastIndexOf("\") + 1)
        'now remove the .iam    
        asyName = asyName.Substring(0, asyName.Length - 4)
        Return asyName
    End Function

    'Private Function SetDefaultRootDirectory()
    '    Dim rootDirectory As String = GetProjectDirectory(_invApp) & _form.TB_Prefix.Text &
    '        oAsyName & _form.TB_Suffix.Text & "\"

    '    Return rootDirectory
    'End Function

    'Function GetProjectDirectory(_invApp As Inventor.Application) As String
    '    Dim actProj As Inventor.DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
    '    Dim projectDir As String = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\") + 1)
    '    Return projectDir
    'End Function

    ''' <summary>
    ''' Checks for an assembly or part that has the same original full file name in this AssemblyCopyObject Only
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <returns></returns>
    Function CheckForDuplicateDocument(ByRef occ As Inventor.ComponentOccurrence) As Boolean
        Dim doc As Document = occ.Definition.Document
        Dim isDuplicate As Boolean = False
        If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            ' this is a part so check the parts list
            For Each part As InvtPartObj In prtList
                If part.OriginalFullFileName = doc.FullFileName Then
                    isDuplicate = True
                    part.AddDuplicateOccurrence(occ)
                    'Debug.WriteLine("Found Duplicate Part: " & part.OriginalName)
                    Exit For
                End If
            Next
        ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            ' this is an assembly so check the sub assembly list
            For Each asy As AssemblyCopyObject In subAsyList
                If asy.OriginalFullFileName = doc.FullFileName Then
                    isDuplicate = True
                    asy.AddDuplicateOccurrence(occ)
                    Exit For
                End If
            Next
        End If

        Return isDuplicate
    End Function

    Sub AddDuplicateOccurrence(ByRef dupOcc As Inventor.ComponentOccurrence)
        duplicateOccurrenceList.Add(dupOcc)
    End Sub

    Function CheckIfOccurenceIsFrame(ByRef compOcc As ComponentOccurrence) As Boolean
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

    Function CheckIfOccurrenceIsBoltedConnection(ByRef compOcc As ComponentOccurrence) As Boolean
        Dim isBoltedConnection As Boolean = False
        If compOcc.AttributeSets.Count > 0 Then
            For Each attSet As AttributeSet In compOcc.AttributeSets
                If attSet.Name = "FDesign" Then
                    For Each atri As Inventor.Attribute In attSet
                        If VarType(atri.Value) = vbString Then
                            Dim atriValue As String = atri.Value
                            If atriValue.IndexOf("CABoltCon") >= 0 Then
                                isBoltedConnection = True
                                Debug.WriteLine(compOcc.Name & " is a bolted connection")
                            End If
                        End If
                    Next
                End If
            Next
        End If

        Return isBoltedConnection
    End Function
    Sub GenerateSetupLog(Optional ByRef isRoot As Boolean = True)
        If isRoot Then
            _form.Log("", numLines:=4)
            _form.Log("*****ROOT SETUP SUMMARY*****", numLines:=1)
            _form.Log("Root Assembly: " & oAsyName)
            _form.Log("Original File Name: " & oFullFileName, numTabs:=1)
            _form.Log("Defualt New Name: " & nAsyName, numTabs:=1)
            _form.Log("Default New File Name: " & nFullFileName, numTabs:=1, numLines:=1)
        Else
            _form.Log("", numLines:=2)
            _form.Log("*****SUB-ASSEMBLY SETUP*****")
            _form.Log("Sub Assembly: " & oAsyName)
            _form.Log("Original File Name: " & oFullFileName, numTabs:=1)
            _form.Log("Defualt New Name: " & nAsyName, numTabs:=1)
            _form.Log("Default New File Name: " & nFullFileName, numTabs:=1, numLines:=1)
        End If

        If prtList.Count > 0 Then
            _form.Log("***** PARTS LIST ******")
            _form.Log("_______________________", numLines:=1)
            For Each part As InvtPartObj In prtList
                _form.Log(part.OriginalName & ": part added")
                _form.Log("original file name: " & part.OriginalFullFileName, numTabs:=1)
                _form.Log("new file name: " & part.NewFullFileName, numTabs:=1, numLines:=1)
            Next
            _form.Log("*****EOL*****")
        End If

        If subAsyList.Count > 0 Then
            _form.Log("", numLines:=1)
            _form.Log("***** Sub-Assembly LIST ******")
            _form.Log("_______________________", numLines:=1)
            For Each subAsy As AssemblyCopyObject In subAsyList
                _form.Log(subAsy.OriginalName & ": assembly added")
                If subAsy.SubType = "Frame" Then
                    _form.Log("**FRAME**")
                End If
                _form.Log("original file name: " & subAsy.OriginalFullFileName, numTabs:=1)
                _form.Log("new file name: " & subAsy.NewFullFileName, numTabs:=1, numLines:=1)
            Next
            _form.Log("*****EOL*****")
            For Each subAsy As AssemblyCopyObject In subAsyList
                subAsy.GenerateSetupLog(False)
            Next
        End If


    End Sub

#End Region

#Region "Misc. Functions"
    Sub NameChange()
        ' remove the last \
        nRootDirectory = nRootDirectory.Substring(0, nRootDirectory.LastIndexOf("\"))
        ' remove the old sub directory but leave the \
        nRootDirectory = nRootDirectory.Substring(0, nRootDirectory.LastIndexOf("\") + 1)
        ' rename the sub directory and add \
        nRootDirectory = nRootDirectory & nAsyName & "\"

        ' rename the treeview node
        nTreeNode.Text = nAsyName
    End Sub

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

#End Region

#Region "File Copy Functions"

    ''' <summary>
    ''' 'update the "new properties" based on changes to the form since load
    ''' </summary>
    Sub UpdateNewProperties(Optional ByVal nRootDirectory = Nothing)
        'assembly already has its node from the original setup so we can just reference that for updates
        'the first run through is the root assembly so we need to add the prefix and suffix

        'update the root directory based on changes since load
        If nRootDirectory Is Nothing Then
            'this is the the root directory and the sub assemblies don't have access to the form
            nRootDirectory = _form.TB_newDir.Text
        End If

        If NewTreeNode.ForeColor = System.Drawing.Color.Red Then
            CopyEnabled = False
        Else
            nAsyName = NewTreeNode.Text
            nFullFileName = nRootDirectory & nAsyName & ".iam"

            For Each part As InvtPartObj In prtList
                part.UpdateNewProperties(nRootDirectory)
            Next

            For Each subAsy As AssemblyCopyObject In subAsyList
                If subAsy.SubType Is "Frame" Then
                    'frame assemblies have a different root directory
                    Dim frameRootDirectory As String = nRootDirectory + nAsyName + "\Frame\"
                    subAsy.UpdateNewProperties(frameRootDirectory)
                ElseIf subAsy.SubType Is "Bolted Connection" Then
                    Dim boltedConnectionDirectory As String = nRootDirectory + nAsyName + "\Design Accelerator\"
                    subAsy.UpdateNewProperties(boltedConnectionDirectory)
                Else
                    subAsy.UpdateNewProperties(nRootDirectory)
                End If
            Next
        End If

    End Sub


    Sub CreateNewFiles(Optional dryrun As Boolean = False)
        'copy the root assembly
        If dryrun Then
            CopyFile_DRYRUN(oFullFileName, nFullFileName)
        Else
            If CopyEnabled Then
                CopyFile(oFullFileName, nFullFileName)
            Else
                _form.Log("Skipping Assembly: " & oAsyName & " because copy enabled is false")
            End If
        End If

        For Each part As InvtPartObj In prtList
            If dryrun Then
                CopyFile_DRYRUN(part.OriginalFullFileName, part.NewFullFileName)
            Else
                If part.SubType = "Content Center Part" Or part.CopyEnabled = False Then
                    'we don't want to copy content center parts
                    If CopyEnabled = False Then
                        _form.Log("Skipping Part: " & part.OriginalName & " because copy enabled is false")
                    Else
                        _form.Log("Skipping Content Center Part: " & part.OriginalName)
                    End If
                Else
                    CopyFile(part.OriginalFullFileName, part.NewFullFileName)
                End If
            End If
        Next

        For Each subAsy As AssemblyCopyObject In subAsyList
            subAsy.CreateNewFiles(dryrun)
        Next

    End Sub

    Private Sub CopyFile(oFile As String, nFile As String)
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
            _form.Log("!!!!!!! FILE SKIPPED BECAUSE IT ALREADY EXISTS !!!!!!!")
        Else
            System.IO.File.Copy(oFile, nFile, False)
            _form.Log("COPY SUCCESSFUL", numTabs:=1, numLines:=1)
            _form.LB_CopyComplete.Text = "Saving File: " & nFile
        End If
    End Sub

    Sub ReplaceOccurrencesByIndex(Optional ByRef asyOcc As ComponentOccurrence = Nothing)

        Dim curAsyOccs As ComponentOccurrences
        'handle root assembly
        If asyOcc Is Nothing Then
            'this is the root assembly
            'create name value map of options for opening the root assembly
            Dim nameValueMap As Inventor.NameValueMap = _invApp.TransientObjects.CreateNameValueMap
            nameValueMap.Add("SkipAllUnresolvedFiles", True)

            ' we need to open the new assembly
            Dim newAsmDoc As Inventor.AssemblyDocument = _invApp.Documents.OpenWithOptions(nFullFileName, nameValueMap, True)

            'assign the assembly occurrence to the root occurrence
            curAsyOccs = newAsmDoc.ComponentDefinition.Occurrences
        Else
            curAsyOccs = asyOcc.Definition.Occurrences
        End If

        If subAsyList.Count > 0 Then
            For Each subAsy As AssemblyCopyObject In subAsyList

                If subAsy.CopyEnabled Then

                    'If subAsy.SubType = "Frame" Then
                    '    _form.Log("Opening Frame Parent Document: " & subAsy.NewFullFileName)
                    '    _invApp.Documents.Open(nFullFileName)
                    '    _form.Log("Active Document: " & _invApp.ActiveDocument.FullFileName)
                    'End If

                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(curAsyOccs, subAsy)

                    ComponentReplace(curOcc, subAsy)

                    'update part number if the name has changed
                    If subAsy.OriginalName IsNot subAsy.NewName Then
                        UpdatePartNumber(curOcc, subAsy, _invApp)
                    End If

                    If subAsy.SubType = "Frame" Then
                        subAsy.ReplaceFrame(curOcc)
                        '_invApp.ActiveDocument.Save2()
                        '_invApp.ActiveDocument.Close(True)
                    Else
                        subAsy.ReplaceOccurrencesByIndex(curOcc)
                    End If
                Else
                    _form.Log("Not replacing sub-assembly " & subAsy.OriginalName & " because copy enabled is false")
                End If
            Next
        End If

        If prtList.Count > 0 Then
            For Each part As InvtPartObj In prtList
                'skip content center parts and parts that are not enabled for copy
                If part.SubType = "Content Center Part" Then
                    _form.Log("Skipping Content Center Part: " & part.OriginalName, numLines:=1)
                ElseIf part.CopyEnabled = False Then
                    _form.Log("Skipping Part: " & part.OriginalName & " because copy enabled is false", numLines:=1)
                Else
                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(curAsyOccs, part)
                    ComponentReplace(curOcc, part)

                    'update part number if the name has changed
                    If part.OriginalName IsNot part.NewName Then
                        UpdatePartNumber(curOcc, part, _invApp)
                    End If
                End If
            Next
        End If


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

    Private Sub ComponentReplace(ByVal curOcc As ComponentOccurrence, ByVal component As Object)
        'only used in the case of a replacement failure
        Dim newFileName As String

        If TypeOf component Is InvtPartObj Then
            Dim part As InvtPartObj = CType(component, InvtPartObj)
            Try
                'replace all instances of this occurrence
                curOcc.Replace(part.NewFullFileName, True)
                _form.Log("Replaced " & part.OriginalName & " with:" & part.NewFullFileName, numLines:=1)
                Return
            Catch pRepEx As Exception
                _form.Log("******ERROR REPLACING PART******")
                _form.Log("Replacing " & part.OriginalName & " with:")
                _form.Log(part.NewFullFileName)
                _form.Log(pRepEx.Message, numLines:=1)
                newFileName = part.NewFullFileName
            End Try

        ElseIf TypeOf component Is AssemblyCopyObject Then
            Dim subAsy As AssemblyCopyObject = CType(component, AssemblyCopyObject)
            Try
                'replace all instances of this occurrence
                curOcc.Replace(subAsy.NewFullFileName, True)
                _form.Log("Replaced " & subAsy.OriginalName & " with:" & subAsy.NewFullFileName, numLines:=1)
                Return
            Catch asyRepEx As Exception
                _form.Log("******ERROR REPLACING SUB-ASSEMBLY******")
                _form.Log(subAsy.OriginalName & " with:")
                _form.Log(subAsy.NewFullFileName)
                _form.Log(asyRepEx.Message, numLines:=1)
                newFileName = subAsy.NewFullFileName
            End Try
        Else
            _form.Log("******ERROR REPLACING COMPONENT******")
            _form.Log("ComponentReplace called with invalid component type.")
            _form.Log(component.GetType.ToString(), numLines:=1)
            Return
        End If


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

    End Sub

    ''' <summary>
    ''' Returns the occurrence based on the occurrence index stored in the part or sub-assembly object
    ''' </summary>
    ''' <param name="occurrences"></param>
    ''' <param name="part"></param>
    ''' <param name="subAsy"></param>
    ''' <returns></returns>
    Private Function GetOccurrenceByIndex(occurrences As ComponentOccurrences, component As Object) As ComponentOccurrence
        Dim compOcc As ComponentOccurrence = Nothing

        If TypeOf component Is InvtPartObj Then
            Dim part As InvtPartObj = CType(component, InvtPartObj)
            Try
                compOcc = occurrences.Item(part.OccurrenceIndex)
            Catch partEx As Exception
                _form.Log("******ERROR*****")
                _form.Log("Error retrieving occurrence by index for part: " & part.OriginalName & "; " & partEx.Message, numLines:=1)
            End Try
        ElseIf TypeOf component Is AssemblyCopyObject Then
            Dim subAsy As AssemblyCopyObject = CType(component, AssemblyCopyObject)
            Try
                compOcc = occurrences.Item(subAsy.OccurrenceIndex)
            Catch asyEx As Exception
                _form.Log("******ERROR*****")
                _form.Log("Error retrieving occurrence by index For Sub-assembly: " & subAsy.OriginalName & "; " & asyEx.Message, numLines:=1)
            End Try
        Else
            _form.Log("******ERROR*****")
            _form.Log("GetOccurrenceByIndex called with invalid component type.")
            _form.Log(component.GetType.ToString(), numLines:=1)
        End If

        Return compOcc
    End Function

    Private Function IsDocumentOpenByFullName(ByVal fullName As String) As Boolean
        Try
            For Each doc As Inventor.Document In _invApp.Documents
                If String.Compare(doc.FullFileName, fullName, StringComparison.OrdinalIgnoreCase) = 0 Then
                    Return True
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("Error enumerating Inventor.Documents: " & ex.Message)
        End Try
        Return False
    End Function

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

#End Region

#Region "Frame Copy Functions"

    ''' <summary>
    ''' Replaces the frame assembly component along with all of its occurences
    ''' Changes the frame assembly id and frame skeleton component id to a new id
    ''' </summary>
    ''' <param name="frmOcc"></param>
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

        Dim curAsyOccs As ComponentOccurrences = frmOcc.Definition.Occurrences
        'replace the parts in frame assembly
        If prtList.Count > 0 Then
            For Each part As InvtPartObj In prtList
                'skip content center parts and parts that are not enabled for copy
                If part.SubType = "Content Center Part" Then
                    _form.Log("Skipping Content Center Part: " & part.OriginalName, numLines:=1)
                ElseIf part.CopyEnabled = False Then
                    _form.Log("Skipping Part: " & part.OriginalName & " because copy enabled is false", numLines:=1)
                Else
                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(curAsyOccs, part)
                    ComponentReplace(curOcc, part)
                    If part.OriginalName IsNot part.NewName Then
                        UpdatePartNumber(curOcc, part, _invApp)
                    End If
                End If
            Next
        End If

        'replace frame sub-assemblies
        If subAsyList.Count > 0 Then
            For Each subAsy As AssemblyCopyObject In subAsyList

                If subAsy.CopyEnabled Then
                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(curAsyOccs, subAsy)
                    ComponentReplace(curOcc, subAsy)
                    If subAsy.OriginalName IsNot subAsy.NewName Then
                        UpdatePartNumber(curOcc, subAsy, _invApp)
                    End If

                    If subAsy.SubType = "Frame" Then
                        subAsy.ReplaceFrame(curOcc)
                    Else
                        subAsy.ReplaceOccurrencesByIndex(curOcc)
                    End If
                Else
                    _form.Log("Not replacing sub-assembly " & subAsy.OriginalName & " because copy enabled is false")
                End If
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

    Function GetSkeletonOcc(ByVal frmOccs As ComponentOccurrences) As ComponentOccurrence
        Dim skeletonOcc As ComponentOccurrence = Nothing
        For Each occ As ComponentOccurrence In frmOccs
            For Each attSet In occ.AttributeSets
                For Each ati As Attribute In attSet
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


    ''' <summary>
    ''' Get the integer for the start of the skeleton id in the frame assembly attribute value
    ''' </summary>
    ''' <param name="atri"></param>
    ''' <returns></returns>
    Private Function GetSkelIdStartInt(ByVal atri As String) As Integer
        Dim skelIDStart = InStr(atri, "SkeletonID")
        Dim skelId As String = atri.Substring(skelIDStart)
        skelIDStart = skelIDStart + InStr(skelId, """")
        Return skelIDStart
    End Function


    ''' <summary>
    ''' Gets the integer for the end of the skeleton id in the frame assembly attribute value
    ''' </summary>
    ''' <param name="atri"></param>
    ''' <param name="skelIdStart"></param>
    ''' <returns></returns>
    Private Function GetSkelIdEndInt(ByVal atri As String, ByVal skelIdStart As Integer) As Integer
        Dim skelId As String = atri.Substring(skelIdStart)
        Dim skelIdEnd As Integer = skelIdStart + InStr(skelId, """") - 1
        Return skelIdEnd
    End Function


    ''' <summary>
    ''' Replaces everything after the final "-" in the original skeleton id with random integers
    ''' </summary>
    ''' <param name="oSkelId"></param>
    ''' <returns></returns>
    Private Function GenerateNewSkelId(ByVal oSkelId As String) As String
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



#End Region

#Region "Dry Run Functions"

    ''' <summary>
    ''' creates a log file of the copy but doesn't execute any file operations
    ''' </summary>
    ''' <param name="oFile"></param>
    ''' <param name="nFile"></param>
    Private Sub CopyFile_DRYRUN(oFile As String, nFile As String)
        _form.Log("copying file: " & oFile)
        Dim nFilePath As String = nFile.Substring(0, nFile.LastIndexOf("\"))
        _form.Log("New File Path: " & nFilePath, numTabs:=1)
        'check if the new directory exists and if it doesn't create one.
        If Directory.Exists(nFilePath) = False Then
            'Directory.CreateDirectory(nFilePath)
            _form.Log("New Directory Created")
        End If

        'check if new file already exists, if so tell them about it
        If System.IO.File.Exists(nFile) Then
            _form.Log("!!!!!!! FILE SKIPPED BECAUSE IT ALREADY EXISTS !!!!!!!")
        Else
            'System.IO.File.Copy(oFile, nFile, False)
            _form.Log("COPY SUCCESSFUL", numTabs:=1, numLines:=1)
        End If
    End Sub

#End Region

#Region "Properties"
    ReadOnly Property OriginalName As String
        Get
            Return oAsyName
        End Get
    End Property

    ReadOnly Property OriginalFullFileName As String
        Get
            Return oFullFileName
        End Get
    End Property

    ReadOnly Property OriginalAsmDocument As AssemblyDocument
        Get
            Return oAsmDoc
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurrence As ComponentOccurrence
        Get
            Return oCompOcc
        End Get
    End Property

    Property NewName As String
        Get
            Return nAsyName
        End Get
        Set(value As String)
            nAsyName = value
            NameChange()
        End Set
    End Property
    Property NewTreeNode As TreeNode
        Get
            Return nTreeNode
        End Get
        Set(value As TreeNode)
            nTreeNode = value
        End Set
    End Property

    Property NewFullFileName As String
        Get
            Return nFullFileName
        End Get
        Set(value As String)
            nFullFileName = value
        End Set
    End Property
    ReadOnly Property NewRootDirectory As String
        Get
            Return nRootDirectory
        End Get
    End Property

    Property SubType As String
        Get
            Return _subType
        End Get
        Set(value As String)
            _subType = value
        End Set
    End Property

    ReadOnly Property PartList As List(Of InvtPartObj)
        Get
            Return prtList
        End Get
    End Property

    ReadOnly Property DuplicateOccurrences As List(Of Inventor.ComponentOccurrence)
        Get
            Return duplicateOccurrenceList
        End Get
    End Property

    ReadOnly Property HighlightSet As HighlightSet
        Get
            Return hltSet
        End Get
    End Property

    Property CopyEnabled As Boolean
        Get
            Return _copyEnabled
        End Get
        Set(value As Boolean)
            _copyEnabled = value
        End Set
    End Property

    Property OccurrenceIndex As Integer
        Get
            Return _occurrenceIndex
        End Get
        Set(value As Integer)
            _occurrenceIndex = value
        End Set
    End Property

#End Region

End Class

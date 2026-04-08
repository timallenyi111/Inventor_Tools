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
    Private ReadOnly _duplicateOccIndexList As List(Of Integer)
    Private _copyEnabled As Boolean = True
    Private _occurrenceIndex As Integer
    Private _containsFrame As Boolean = False

    Public Sub New(form As AssemblyCopyToolForm, invApp As Inventor.Application)
        _form = form
        _invApp = invApp
        _contentCenterPath = _invApp.DesignProjectManager.ActiveDesignProject.ContentCenterPath

        prtList = New List(Of InvtPartObj)
        subAsyList = New List(Of AssemblyCopyObject)
        duplicateOccurrenceList = New List(Of Inventor.ComponentOccurrence)
        _duplicateOccIndexList = New List(Of Integer)
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
                If CheckForDuplicateDocument(curOcc, curOccIndex) = False Then
                    ' perform part setup
                    Dim curPartObject As New InvtPartObj
                    curPartObject.InitialSetup(curOcc, nRootDirectory, nTreeNode, _contentCenterPath)
                    curPartObject.OccurrenceIndex = curOccIndex
                    prtList.Add(curPartObject)
                Else
                    _form.Log(curOcc._DisplayName & " was a duplicate")
                End If

            ElseIf curOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                If CheckForDuplicateDocument(curOcc, curOccIndex) = False Then
                    ' perform sub assembly setup

                    Dim curAsmObject As New AssemblyCopyObject(_form, _invApp)
                    If CheckIfOccurenceIsFrame(curOcc) Then
                        'set the contains frame flag to true so that when replacing the components we will open this assembly
                        _containsFrame = True
                        ' this is a frame assembly so we need to set the frame directory
                        Dim frameRootDirectory As String = nRootDirectory + nTreeNode.Text + "\Frame\"
                        curAsmObject.InitialSetup(curOcc, frameRootDirectory, nTreeNode)
                        curAsmObject.SubType = "Frame"
                        curAsmObject.OccurrenceIndex = curOccIndex
                        _form.Log(curOcc._DisplayName & " was identified as a frame assembly")
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
        'AssignNodeTags()
        AssignNodeTagsByIndex()
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

    ''' <summary>
    ''' Checks for an assembly or part that has the same original full file name in this AssemblyCopyObject Only
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <returns></returns>
    Function CheckForDuplicateDocument(ByRef occ As Inventor.ComponentOccurrence, occIndex As Integer) As Boolean
        Dim doc As Document
        Dim isDuplicate As Boolean = False
        Try
            doc = occ.Definition.Document
            If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                ' this is a part so check the parts list
                For Each part As InvtPartObj In prtList
                    If part.OriginalFullFileName = doc.FullFileName Then
                        isDuplicate = True
                        part.AddDuplicateOccurrence(occ, occIndex)
                        'Debug.WriteLine("Found Duplicate Part: " & part.OriginalName)
                        Exit For
                    End If
                Next
            ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ' this is an assembly so check the sub assembly list
                For Each asy As AssemblyCopyObject In subAsyList
                    If asy.OriginalFullFileName = doc.FullFileName Then
                        isDuplicate = True
                        asy.AddDuplicateOccurrence(occ, occIndex)
                        Exit For
                    End If
                Next
            End If
        Catch
            Console.WriteLine("Not able to perform duplicate check for", occ.Name)
        End Try

        Return isDuplicate
    End Function

    Sub AddDuplicateOccurrence(ByRef dupOcc As Inventor.ComponentOccurrence, occIndex As Integer)
        duplicateOccurrenceList.Add(dupOcc)
        _duplicateOccIndexList.Add(occIndex)
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

    Sub AssignNodeTagsByIndex(Optional occurrences As Inventor.ComponentOccurrences = Nothing)
        'handle root assembly
        If occurrences Is Nothing Then
            'this is the root assembly
            occurrences = oAsmDoc.ComponentDefinition.Occurrences

            For Each part As InvtPartObj In prtList
                'parts in the root assembly
                Dim partNode As System.Windows.Forms.TreeNode = part.NewTreeNode
                'in the root assembly the component occurrence is the only thing you need for highlighting
                'Dim occ As Inventor.ComponentOccurrence = oAsmDoc.ComponentDefinition.Occurrences.Item(part.OccurrenceIndex)
                Dim occList As New List(Of Inventor.ComponentOccurrence)
                occList.Add(part.OriginalComponentOccurrence)
                For Each dupOcc As Inventor.ComponentOccurrence In part.DuplicateOccurrences
                    occList.Add(dupOcc)
                Next

                partNode.Tag = occList
            Next

            For Each subAsy As AssemblyCopyObject In subAsyList
                Dim asyNode As System.Windows.Forms.TreeNode = subAsy.NewTreeNode
                Dim occList As New List(Of Inventor.ComponentOccurrence)
                'Dim occProxy As Inventor.ComponentOccurrenceProxy = occurrences.Item(subAsy.OccurrenceIndex)

                occList.Add(subAsy.OriginalComponentOccurrence)
                For Each dupOcc As Inventor.ComponentOccurrence In subAsy.DuplicateOccurrences
                    occList.Add(dupOcc)
                Next

                asyNode.Tag = occList

                'process components in the sub-assembly
                subAsy.AssignNodeTagsByIndex(subAsy.OriginalComponentOccurrence.SubOccurrences)
            Next

        Else
            If prtList.Count > 0 Then
                For Each part As InvtPartObj In prtList
                    'parts in the root assembly
                    Dim partNode As System.Windows.Forms.TreeNode = part.NewTreeNode
                    'in the root assembly the component occurrence is the only thing you need for highlighting
                    'Dim occ As Inventor.ComponentOccurrence = oAsmDoc.ComponentDefinition.Occurrences.Item(part.OccurrenceIndex)
                    Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
                    Dim occProxy As Inventor.ComponentOccurrenceProxy = occurrences.Item(part.OccurrenceIndex)
                    occProxyList.Add(occProxy)
                    For Each occIndex As Integer In part.DuplicateOccurrenceIndexList
                        Dim dupOccProxy As Inventor.ComponentOccurrenceProxy = occurrences.Item(occIndex)
                        occProxyList.Add(dupOccProxy)
                    Next

                    partNode.Tag = occProxyList
                Next
            End If

            If subAsyList.Count > 0 Then
                For Each subAsy As AssemblyCopyObject In subAsyList
                    Dim asyNode As System.Windows.Forms.TreeNode = subAsy.NewTreeNode
                    Dim occProxyList As New List(Of Inventor.ComponentOccurrenceProxy)
                    Dim occProxy As Inventor.ComponentOccurrenceProxy = occurrences.Item(subAsy.OccurrenceIndex)
                    occProxyList.Add(occProxy)
                    For Each occIndex As Integer In subAsy.DuplicateOccurrenceIndexList
                        Dim dupOccProxy As Inventor.ComponentOccurrenceProxy = occurrences.Item(occIndex)
                        occProxyList.Add(dupOccProxy)
                    Next

                    asyNode.Tag = occProxyList

                    'process components in the sub-assembly
                    subAsy.AssignNodeTagsByIndex(occProxy.SubOccurrences)
                Next
            End If
        End If



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
                If ContainsFrame Then
                    _form.Log("This assembly contains a frame sub-assembly. When copying this assembly, the frame sub-assembly will be copied and replaced first before copying and replacing the rest of the components in this assembly because of the potential complexity of the frame assembly and the fact that it is often generated from iLogic with unique file naming schemes for the components within it. After the frame sub-assembly has been copied and replaced, the rest of the components in this assembly will be copied and replaced as normal.", numLines:=2)
                Else
                    CopyFile(oFullFileName, nFullFileName)
                End If
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
            If subAsy.SubType = "Frame" Then
                '_form.Log("Frame Assembly: " & subAsy.OriginalName)
                '_form.Log("New Frame Name: " & subAsy.NewName, numTabs:=1)
                '_form.Log("Frame assemblies will be copied and replaced but not processed for part/sub-assembly copying or replacement within them because of the potential complexity of the frame assemblies and the fact that they are often generated from iLogic with unique file naming schemes for the components within them.")
                CopyFrameFile(subAsy)
                subAsy.CreateNewFiles(dryrun)
            Else
                subAsy.CreateNewFiles(dryrun)
            End If

        Next
    End Sub

    Private Sub CopyFrameFile(frame As AssemblyCopyObject)
        'used to make sure we close the opened sub assembly before moving on.
        Dim openedNewAssembly As Boolean = False
        Dim nameValueMap As Inventor.NameValueMap = _invApp.TransientObjects.CreateNameValueMap
        nameValueMap.Add("SkipAllUnresolvedFiles", True)
        Dim oParentAsmDoc As AssemblyDocument = Nothing
        If _invApp.ActiveDocument.FullFileName = Me.OriginalFullFileName Then
            'the frame is in the root assembly so we don't need to open anything.
        Else
            ' we need to open the new assembly
            ' at this point the frame assembly is still a sub assembly of the parent
            oParentAsmDoc = _invApp.Documents.OpenWithOptions(oFullFileName, nameValueMap, True)

            openedNewAssembly = True
        End If

        Dim frmOcc As ComponentOccurrence = GetOccurrenceByIndex(OriginalAsmDocument.ComponentDefinition.Occurrences, frame)

        'the original frame attribute value (we will replace this once the copy has been completed)
        Dim oAtriVal As String = Nothing

        Dim nskelID As String = Nothing
        Dim oskelID As String = Nothing
        'try changing the skeleton ID and path ID before changing
        'replace the skelton id in the frame assembly attributes
        'change the frame attributes in the original file
        For Each attSet As AttributeSet In frmOcc.Definition.AttributeSets
            For Each atri As Attribute In attSet
                If atri.Name = "Frame.Skeletons" Then
                    oAtriVal = atri.Value
                    Dim skelIdStart As Integer = GetSkelIdStartInt(oAtriVal)
                    Dim skelIdEnd As Integer = GetSkelIdEndInt(oAtriVal, skelIdStart)
                    Dim skelPathStart As Integer = GetPathIdStartInt(oAtriVal)
                    Dim skelPathEnd As Integer = GetPathIdEndInt(oAtriVal, skelPathStart)

                    oskelID = oAtriVal.Substring(skelIdStart, skelIdEnd - skelIdStart)
                    Dim oSkelPath As String = oAtriVal.Substring(skelPathStart, skelPathEnd - skelPathStart)

                    nskelID = GenerateNewID(oSkelId)
                    Dim nskelPath As String = GenerateNewID(oSkelPath)


                    'replace the skeleton ID
                    Dim nAtriVal As String = oAtriVal.Substring(0, skelIdStart) & nskelID &
                        oAtriVal.Substring(skelIdEnd)

                    'nAtriVal = nAtriVal.Substring(0, skelPathStart) & nskelPath &
                    'nAtriVal.Substring(skelPathEnd)

                    atri.Value = nAtriVal
                End If
            Next
        Next

        Dim skelOcc As ComponentOccurrence = GetSkeletonOcc(frmOcc.Definition.Occurrences)

        For Each attSet As AttributeSet In skelOcc.AttributeSets
            For Each att As Attribute In attSet
                ' replace the old skeleton id with the new
                If att.Name = "ID" Then
                    att.Value = nskelID
                End If
            Next
        Next

        'save the new file
        _invApp.ActiveDocument.SaveAs(nFullFileName, True)

        'close the saved copy
        _invApp.ActiveDocument.Close()

        'reopen the original file        
        oParentAsmDoc = _invApp.Documents.OpenWithOptions(oFullFileName, nameValueMap, True)

        'reset the frame occurence to the occurrence in the original assembly
        frmOcc = GetOccurrenceByIndex(OriginalAsmDocument.ComponentDefinition.Occurrences, frame)


        'change the skeleton ID and path back to the original
        For Each attSet As AttributeSet In frmOcc.Definition.AttributeSets
            For Each atri As Attribute In attSet
                If atri.Name = "Frame.Skeletons" Then
                    atri.Value = oAtriVal
                End If
            Next
        Next


        skelOcc = GetSkeletonOcc(frmOcc.Definition.Occurrences)
        'set the original skeleton occurence id back to the original skeleton id
        For Each attSet As AttributeSet In skelOcc.AttributeSets
            For Each att As Attribute In attSet
                ' replace the old skeleton id with the new
                If att.Name = "ID" Then
                    att.Value = oskelID
                End If
            Next
        Next

        'resave the original assembly with the original skeleton ID and path
        _invApp.ActiveDocument.Save2()

        If openedNewAssembly Then
            'close the new assembly 
            _invApp.ActiveDocument.Close(True)
        End If

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

                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(curAsyOccs, subAsy)

                    If subAsy.SubType = "Frame" Then
                        'If nTreeNode.Parent IsNot Nothing Then
                        '    'this is not the root directory so we need to open this document
                        '    'we need to open the parent document of the frame assembly occurrence in order to replace the frame assembly occurrence
                        '    Dim nameValueMap As Inventor.NameValueMap = _invApp.TransientObjects.CreateNameValueMap
                        '    nameValueMap.Add("SkipAllUnresolvedFiles", True)

                        '    ' we need to open the new assembly
                        '    ' at this point the frame assembly is still a sub assembly of the parent
                        '    Dim parentAsmDoc As AssemblyDocument = _invApp.Documents.OpenWithOptions(nFullFileName, nameValueMap, True)

                        '    'we have to recapture the component occurrence because we opened a new file
                        '    Dim currentDoc As AssemblyDocument = _invApp.ActiveDocument
                        '    curOcc = GetOccurrenceByIndex(currentDoc.ComponentDefinition.Occurrences, subAsy)
                        '    Debug.WriteLine("Replacing frame assembly occurrence: " & curOcc.Name)
                        '    ComponentReplace(curOcc, subAsy)
                        '    'subAsy.ReplaceFrame(curOcc, subAsy)
                        '    subAsy.ReplaceOccurrencesByIndex(curOcc)

                        'Else
                        '    ComponentReplace(curOcc, subAsy)
                        '    'subAsy.ReplaceFrame(curOcc, subAsy)
                        '    subAsy.ReplaceOccurrencesByIndex(curOcc)
                        'End If

                        'ReplaceFrame(curOcc, subAsy)
                        'ComponentReplace(curOcc, subAsy)

                        'we need to open the frame assembly and replace all of the components prior to replacing it in the assembly
                        'if we call "replaceocccurencesByIndex" with no occurrence it will treat the frame like a root assembly and open the frame file
                        subAsy.ReplaceOccurrencesByIndex()

                        'the frame document should be the active document coming back from component replacing
                        Dim frameDoc As AssemblyDocument = _invApp.ActiveDocument
                        'save and close the document
                        frameDoc.Save2()
                        frameDoc.Close()

                        'now replace the frame assembly in the parent assembly with the new file that we just saved
                        ComponentReplace(curOcc, subAsy)


                        'If nTreeNode.Parent IsNot Nothing Then
                        '    'after replacing the frame assembly occurrence we can close the parent assembly document because the frame assembly is now in place and we don't need to access the parent assembly anymore
                        '    _invApp.ActiveDocument.Save2()
                        '    _invApp.ActiveDocument.Close()
                        'End If

                    Else
                        ComponentReplace(curOcc, subAsy)
                        'replace all of the components within the sub assembly
                        subAsy.ReplaceOccurrencesByIndex(curOcc)
                    End If

                    'ComponentReplace(curOcc, subAsy)

                    'update part number if the name has changed
                    If subAsy.OriginalName IsNot subAsy.NewName Then
                        UpdatePartNumber(curOcc, subAsy, _invApp)
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
                System.Threading.Thread.Sleep(1000)
                _form.Log("Trying again after a 1 second delay...")
                Try
                    curOcc.Replace(part.NewFullFileName, True)
                    _form.Log("SUCCESSFULLY REPLACED ON SECOND ATTEMPT")
                    _form.Log("Replaced " & part.OriginalName & " with:" & part.NewFullFileName, numLines:=1)
                Catch ex As Exception
                    _form.Log("Second attempt failed.")
                End Try

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
                _form.Log(curOcc.Name & " was not replaced successfully on the first attempt.")
                _form.Log("******ERROR REPLACING SUB-ASSEMBLY******")
                _form.Log(subAsy.OriginalName & " with:")
                _form.Log(subAsy.NewFullFileName)
                _form.Log(asyRepEx.Message, numLines:=1)
                newFileName = subAsy.NewFullFileName
                System.Threading.Thread.Sleep(1000)
                _form.Log("Trying again after a 1 second delay...")
                Try
                    curOcc.Replace(subAsy.NewFullFileName, True)
                    _form.Log("SUCCESSFULLY REPLACED ON SECOND ATTEMPT")
                    _form.Log("Replaced " & subAsy.OriginalName & " with:" & subAsy.NewFullFileName, numLines:=1)
                Catch ex As Exception
                    _form.Log("Second attempt failed.")
                End Try
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

#End Region

#Region "Frame Copy Functions"

    ''' <summary>
    ''' Replaces the frame assembly component along with all of its occurences
    ''' Changes the frame assembly id and frame skeleton component id to a new id
    ''' </summary>
    ''' <param name="frmOcc"></param>
    ''' <param name="frmAssemblyObject"></param>
    Private Sub ReplaceFrame(ByRef frmOcc As ComponentOccurrence, ByRef frmAssemblyObject As AssemblyCopyObject)
        Debug.WriteLine("Replacing Frame Assembly: " & frmOcc.Name)

        'replace the old skeleton id with a new one
        Dim nSkelId As String = Nothing

        'replace the skelton id in the frame assembly attributes
        'we are now doing this in the copy phase
        For Each attSet As AttributeSet In frmOcc.Definition.AttributeSets
            For Each atri As Attribute In attSet
                If atri.Name = "Frame.Skeletons" Then
                    Dim oAtriVal As String = atri.Value
                    Dim skelIdStart As Integer = GetSkelIdStartInt(oAtriVal)
                    Dim skelIdEnd As Integer = GetSkelIdEndInt(oAtriVal, skelIdStart)

                    Dim oSkelId As String = oAtriVal.Substring(skelIdStart, skelIdEnd - skelIdStart)

                    nSkelId = GenerateNewID(oSkelId)

                    Dim nAtriVal As String = oAtriVal.Substring(0, skelIdStart) & nSkelId &
                        oAtriVal.Substring(skelIdEnd)

                    atri.Value = nAtriVal
                End If
            Next
        Next

        Dim frameAsyOccs As ComponentOccurrences = frmOcc.Definition.Occurrences
        'replace the parts in frame assembly
        If prtList.Count > 0 Then
            For Each part As InvtPartObj In prtList
                'skip content center parts and parts that are not enabled for copy
                If part.SubType = "Content Center Part" Then
                    _form.Log("Skipping Content Center Part: " & part.OriginalName, numLines:=1)
                ElseIf part.CopyEnabled = False Then
                    _form.Log("Skipping Part: " & part.OriginalName & " because copy enabled is false", numLines:=1)
                Else
                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(frameAsyOccs, part)
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
                    Dim curOcc As ComponentOccurrence = GetOccurrenceByIndex(frameAsyOccs, subAsy)

                    If subAsy.OriginalName IsNot subAsy.NewName Then
                        UpdatePartNumber(curOcc, subAsy, _invApp)
                    End If

                    If subAsy.SubType = "Frame" Then
                        'subAsy.ReplaceFrame(curOcc, subAsy)
                        Debug.WriteLine("Replacing Frame Sub-Assembly: " & subAsy.OriginalName)
                        ReplaceFrame(curOcc, subAsy)
                    Else
                        ComponentReplace(curOcc, subAsy)
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


    Private Sub ReplaceFrame2(ByRef frmOcc As ComponentOccurrence, ByRef frmAssemblyObject As AssemblyCopyObject)
        Debug.WriteLine("Replacing Frame Assembly: " & frmOcc.Name)

        'open the frame assembly and replace all of the components within it before trying to place the frame in the new assembly.
        Dim nameValueMap As Inventor.NameValueMap = _invApp.TransientObjects.CreateNameValueMap
        nameValueMap.Add("SkipAllUnresolvedFiles", True)

        'we need to open the new assembly
        'at this point the frame assembly Is still a sub assembly of the parent
        Dim frameAsmDoc As AssemblyDocument = _invApp.Documents.OpenWithOptions(frmAssemblyObject.NewFullFileName, nameValueMap, True)

        Dim frameAsyOccs As ComponentOccurrences = frameAsmDoc.ComponentDefinition.Occurrences

        frmAssemblyObject.ReplaceOccurrencesByIndex()

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

    Private Function GetPathIdStartInt(ByVal atri As String) As Integer
        Dim pathIDStart = InStr(atri, "PathID")
        Dim pathId As String = atri.Substring(pathIDStart)
        pathIDStart = pathIDStart + InStr(pathId, """")
        Return pathIDStart
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

    Private Function GetPathIdEndInt(ByVal atri As String, ByVal pathIdStart As Integer) As Integer
        Dim pathId As String = atri.Substring(pathIdStart)
        Dim pathIdEnd As Integer = pathIdStart + InStr(pathId, """") - 1
        Return pathIdEnd
    End Function


    ''' <summary>
    ''' Replaces everything after the final "-" in the original skeleton id with random integers
    ''' </summary>
    ''' <param name="oSkelId"></param>
    ''' <returns></returns>
    Private Function GenerateNewID(ByVal oSkelId As String) As String
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

    ReadOnly Property DuplicateOccurrenceIndexList As List(Of Integer)
        Get
            Return _duplicateOccIndexList
        End Get
    End Property

    ReadOnly Property ContainsFrame As Boolean
        Get
            Return _containsFrame
        End Get
    End Property

#End Region

End Class

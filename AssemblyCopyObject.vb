Imports Inventor
Imports System.Windows.Forms
Imports System.IO

' part list
' frame list
' assembly list
' treenode

''' <summary>
''' This stores all of the information necessary to make a copy of an assembly
''' </summary>
Friend Class AssemblyCopyObject

    Private ReadOnly _form As AssemblyCopyToolForm
    Private ReadOnly _invApp As Inventor.Application
    Dim partList As List(Of InvtPartObj)
    Dim subAsyList As List(Of AssemblyCopyObject)
    Dim oAsyName As String
    Dim nAsyName As String
    Dim oFullFileName As String
    Dim nFullFileName As String
    Dim nRootDirectory As String
    Dim oAsmDoc As AssemblyDocument
    Dim oTreeNode As TreeNode
    Dim nTreeNode As TreeNode
    Dim oCompOcc As ComponentOccurrence
    Dim _subType As String
    Dim projectDirectory As String

    Public Sub New(form As AssemblyCopyToolForm, invApp As Inventor.Application)
        _form = form
        _invApp = invApp

        partList = New List(Of InvtPartObj)
        subAsyList = New List(Of AssemblyCopyObject)
    End Sub

#Region "setup functions"
    Sub InitialSetup(Optional asyOcc As ComponentOccurrence = Nothing, Optional rootDirectory As String = Nothing,
                     Optional oParentTreeNode As TreeNode = Nothing, Optional nParentTreeNode As TreeNode = Nothing)

        ' this is the root assembly
        If asyOcc Is Nothing Then
            ' this is the root assembly         
            SetOriginalProperties(_invApp.ActiveDocument)
            ' define the root directory for the entire assembly
            nRootDirectory = SetDefaultRootDirectory()
            SetNewProperties()
        Else
            ' this is a subassembly
            SetOriginalProperties(asyOcc.Definition.Document, asyOcc, oParentTreeNode)
            nRootDirectory = rootDirectory
            SetNewProperties(nParentTreeNode) ' sub assemblys don't automatically get the pre/suffix
        End If

        _form.Log("", numLines:=1)
        _form.Log("***** " & oAsyName & " component setup *****")
        For Each curOcc As ComponentOccurrence In oAsmDoc.ComponentDefinition.Occurrences
            If curOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                If CheckForDuplicateDocument(curOcc.Definition.Document) = False Then
                    ' perform part setup
                    Dim curPartObject As New InvtPartObj
                    curPartObject.InitialSetup(curOcc, nRootDirectory, oTreeNode, nTreeNode)
                    partList.Add(curPartObject)
                Else
                    _form.Log(curOcc._DisplayName & " was a duplicate")
                End If

            ElseIf curOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                If CheckForDuplicateDocument(curOcc.Definition.Document) = False Then
                    ' perform sub assembly setup

                    Dim curAsmObject As New AssemblyCopyObject(_form, _invApp)
                    curAsmObject.InitialSetup(curOcc, nRootDirectory, oTreeNode, nTreeNode)
                    subAsyList.Add(curAsmObject)
                Else
                    _form.Log(curOcc._DisplayName & " was a duplicate")
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Sets up all the initial parameters for the original assembly file
    ''' </summary>
    ''' <param name="AsyOcc"></param>
    ''' <param name="ParentAssembly"></param>
    Sub SetOriginalProperties(ByRef AsyDoc As AssemblyDocument,
                              Optional ByRef AsyOcc As ComponentOccurrence = Nothing,
                              Optional ByRef oParentNode As TreeNode = Nothing)

        oCompOcc = AsyOcc
        oAsmDoc = AsyDoc
        oFullFileName = oAsmDoc.FullFileName
        oAsyName = GetAssemblyName(oFullFileName)

        'if there is no parent occurence then it is the main assembly and so the first tree node has to be created
        If AsyOcc Is Nothing Then
            oTreeNode = New TreeNode(oAsyName)
            subType = "Root"
        Else
            If CheckIfOccurenceIsFrame(oCompOcc) Then
                _subType = "Frame"
            End If
            oTreeNode = oParentNode.Nodes.Add(oAsyName)
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
            nAsyName = oAsyName
            nFullFileName = nRootDirectory & nAsyName & ".iam"
            nTreeNode = nParentNode.Nodes.Add(nAsyName)
        End If
    End Sub

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

    Private Function SetDefaultRootDirectory()
        Dim rootDirectory As String = GetProjectDirectory(_invApp) & _form.TB_Prefix.Text &
            oAsyName & _form.TB_Suffix.Text & "\"

        Return rootDirectory
    End Function

    Function GetProjectDirectory(_invApp As Inventor.Application) As String
        Dim actProj As Inventor.DesignProject = _invApp.DesignProjectManager.ActiveDesignProject
        Dim projectDir As String = actProj.FullFileName.Substring(0, actProj.FullFileName.LastIndexOf("\") + 1)
        Return projectDir
    End Function

    ''' <summary>
    ''' Checks for an assembly or part that has the same original full file name in this AssemblyCopyObject Only
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <returns></returns>
    Function CheckForDuplicateDocument(ByRef doc As Inventor.Document) As Boolean
        Dim isDuplicate As Boolean = False
        If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            ' this is a part so check the parts list
            For Each part As InvtPartObj In partList
                If part.OriginalFullFileName = doc.FullFileName Then
                    isDuplicate = True
                    Exit For
                End If
            Next
        ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            ' this is an assembly so check the sub assembly list
            For Each asy As AssemblyCopyObject In subAsyList
                If asy.OriginalFullFileName = doc.FullFileName Then
                    isDuplicate = True
                    Exit For
                End If
            Next
        End If

        Return isDuplicate
    End Function

    Function CheckIfOccurenceIsFrame(ByRef compOcc As ComponentOccurrence) As Boolean
        Dim isFrame As Boolean = False
        For Each attSet As AttributeSet In compOcc.AttributeSets
            For Each atri As Inventor.Attribute In attSet
                If atri.Value = "MasterFrameOcc" Then
                    isFrame = True
                End If
            Next
        Next
        Return isFrame
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

        If partList.Count > 0 Then
            _form.Log("***** PARTS LIST ******")
            _form.Log("_______________________", numLines:=1)
            For Each part As InvtPartObj In partList
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

#Region "Update Functions"
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

#End Region

#Region "File Copy Functions"
    Sub CreateNewFiles(Optional dryrun As Boolean = False)

        'copy the root assembly
        If dryrun Then
            CopyFile_DRYRUN(oFullFileName, nFullFileName)
        Else
            CopyFile(oFullFileName, nFullFileName)
        End If

        For Each part As InvtPartObj In partList
            If dryrun Then
                CopyFile_DRYRUN(part.OriginalFullFileName, part.NewFullFileName)
            Else
                CopyFile(part.OriginalFullFileName, part.NewFullFileName)
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

        'check if new file already exists, if so tell them about it
        If System.IO.File.Exists(nFile) Then
            _form.Log("!!!!!!! FILE SKIPPED BECAUSE IT ALREADY EXISTS !!!!!!!")
        Else
            System.IO.File.Copy(oFile, nFile, False)
            _form.Log("COPY SUCCESSFUL", numTabs:=1, numLines:=1)
        End If
    End Sub

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

    ''' <summary>
    ''' Replaces the components in assemblys and sub-assemblies
    ''' </summary>
    ''' <param name="asyOcc"></param>
    ''' <param name="skelId"></param>
    Sub ReplaceOccurences(Optional ByRef asyOcc As ComponentOccurrence = Nothing,
                          Optional ByVal skelId As String = Nothing)
        If asyOcc Is Nothing Then
            ' this is the root assembly
            ' we need to open the new assembly
            Dim newAsmDoc As Inventor.AssemblyDocument = _invApp.Documents.Open(nFullFileName)
            Dim newAsmOccs As Inventor.ComponentOccurrences = newAsmDoc.ComponentDefinition.Occurrences

            'replace the parts in the root assembly
            If partList.Count > 0 Then
                For Each part As InvtPartObj In partList
                    newAsmOccs.ItemByName(part.OriginalComponentOccurence.Name).Replace(part.NewFullFileName, True)
                Next
            End If

            If subAsyList.Count > 0 Then
                For Each subAsy As AssemblyCopyObject In subAsyList
                    Dim curOcc As ComponentOccurrence = newAsmOccs.ItemByName(subAsy.OriginalComponentOccurence.Name)
                    curOcc.Replace(subAsy.NewFullFileName, True)
                    If subAsy.SubType = "Frame" Then
                        subAsy.ReplaceFrame(curOcc)
                    Else
                        subAsy.ReplaceOccurences(curOcc)
                    End If

                Next
            End If

        Else
            Dim subAsyOccs As ComponentOccurrences = asyOcc.Definition.Occurrences
            'replace the parts in the root assembly
            If partList.Count > 0 Then
                For Each part As InvtPartObj In partList
                    subAsyOccs.ItemByName(part.OriginalComponentOccurence.Name).Replace(part.NewFullFileName, True)
                Next
            End If

            If subAsyList.Count > 0 Then
                For Each subAsy As AssemblyCopyObject In subAsyList
                    'get the occurence of the current subAsy by searching for it by name using the original occurence name
                    Dim curOcc As ComponentOccurrence = subAsyOccs.ItemByName(subAsy.OriginalComponentOccurence.Name)

                    'recall this sub by getting the occurence of the component to be replaced by 
                    curOcc.Replace(subAsy.NewFullFileName, True)
                    If subAsy.SubType = "Frame" Then
                        subAsy.ReplaceFrame(curOcc)
                    Else
                        subAsy.ReplaceOccurences(curOcc)
                    End If

                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' Replaces the frame assembly component along with all of its occurences
    ''' Changes the frame assembly id and frame skeleton component id to a new id
    ''' </summary>
    ''' <param name="frmOcc"></param>
    Private Sub ReplaceFrame(ByRef frmOcc As ComponentOccurrence)
        'replace the old skeleton id with a new one
        Dim nSkelId As String = Nothing
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
        'replace the parts in the root assembly
        If partList.Count > 0 Then
            For Each part As InvtPartObj In partList
                subAsyOccs.ItemByName(part.OriginalComponentOccurence.Name).Replace(part.NewFullFileName, True)
            Next
        End If

        If subAsyList.Count > 0 Then
            For Each subAsy As AssemblyCopyObject In subAsyList
                'get the occurence of the current subAsy by searching for it by name using the original occurence name
                Dim curOcc As ComponentOccurrence = subAsyOccs.ItemByName(subAsy.OriginalComponentOccurence.Name)

                'recall this sub by getting the occurence of the component to be replaced by 
                curOcc.Replace(subAsy.NewFullFileName, True)
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

    ReadOnly Property OriginalPartDocument As PartDocument
        Get
            Return oAsmDoc
        End Get
    End Property

    ReadOnly Property OriginalTreeNode As TreeNode
        Get
            Return oTreeNode
        End Get
    End Property

    ReadOnly Property OriginalComponentOccurence As ComponentOccurrence
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

#End Region

End Class

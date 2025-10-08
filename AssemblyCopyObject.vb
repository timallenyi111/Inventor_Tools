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
    Dim subType As String
    Dim projectDirectory As String

    Public Sub New(form As AssemblyCopyToolForm, invApp As Inventor.Application)
        _form = form
        _invApp = invApp

        partList = New List(Of InvtPartObj)
        subAsyList = New List(Of AssemblyCopyObject)
    End Sub
    Sub InitialSetup(Optional asyOcc As ComponentOccurrence = Nothing, Optional rootDirectory As String = Nothing,
                     Optional oParentTreeNode As TreeNode = Nothing, Optional nParentTreeNode As TreeNode = Nothing)

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
                _form.Log("original file name: " & subAsy.OriginalFullFileName, numTabs:=1)
                _form.Log("new file name: " & subAsy.NewFullFileName, numTabs:=1, numLines:=1)
            Next
            _form.Log("*****EOL*****")
            For Each subAsy As AssemblyCopyObject In subAsyList
                subAsy.GenerateSetupLog(False)
            Next
        End If


    End Sub

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
    ''' <summary>
    ''' Sets up all the initial parameters for the original assembly file
    ''' </summary>
    ''' <param name="AsyOcc"></param>
    ''' <param name="ParentAssembly"></param>

#End Region

End Class

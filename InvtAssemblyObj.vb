Public Class InvtAssemblyObj
    Dim compDefs As List(Of Inventor.ComponentDefinition)
    Dim asmDef As Inventor.AssemblyComponentDefinition
    Dim asmCompList As List(Of InvtComponentObj)
    Dim oAsmDoc As Inventor.AssemblyDocument
    Dim nAsmDoc As Inventor.AssemblyDocument
    Dim assmName As String
    Dim oFileName As String
    Dim oFilePath As String
    Dim nName As String
    Dim nFileName As String
    Dim nFilePath As String
    Dim nFullFileName As String
    Dim tNode As New TreeNode

    Public Sub New()
        compDefs = New List(Of Inventor.ComponentDefinition)()
        asmCompList = New List(Of InvtComponentObj)()
    End Sub

    Property OriginalName As String
        Get
            Return assmName
        End Get
        Set(value As String)
            assmName = value
        End Set
    End Property

    Property OriginalFullFileName As String
        Get
            Return oFileName
        End Get
        Set(value As String)
            oFileName = value
        End Set
    End Property

    Property OriginalFilePath As String
        Get
            Return oFilePath
        End Get
        Set(value As String)
            oFilePath = value
        End Set
    End Property

    Property AssemblyComponents As List(Of InvtComponentObj)
        Get
            Return asmCompList
        End Get
        Set(value As List(Of InvtComponentObj))
            asmCompList = value
        End Set
    End Property

    ''' <summary>
    ''' The Inventor Assembly Document object.
    ''' </summary>
    ''' <returns></returns>
    Property OriginalAssemblyDocument As Inventor.AssemblyDocument
        Get
            Return oAsmDoc
        End Get
        Set(value As Inventor.AssemblyDocument)
            oAsmDoc = value
        End Set
    End Property

    Property NewAssemblyDocument As Inventor.AssemblyDocument
        Get
            Return nAsmDoc
        End Get
        Set(value As Inventor.AssemblyDocument)
            nAsmDoc = value
        End Set
    End Property

    ''' <summary>
    ''' The name of the assembly without the file extension.
    ''' </summary>
    ''' <returns></returns> 
    Property NewName As String
        Get
            Return nName
        End Get
        Set(value As String)
            nName = value
            NewFileName = nName & ".iam"
        End Set
    End Property

    Property NewFileName As String
        Get
            Return nFileName
        End Get
        Set(value As String)
            nFileName = value
        End Set
    End Property

    Property NewFilePath As String
        Get
            Return nFilePath
        End Get
        Set(value As String)
            nFilePath = value
        End Set
    End Property

    Property NewTreeNode As TreeNode
        Get
            Return tNode
        End Get
        Set(value As TreeNode)
            tNode = value
        End Set
    End Property

    ReadOnly Property NewFullFileName As String
        Get
            Return nFilePath & nFileName
        End Get
    End Property

    ''' <summary>
    ''' Returns the InvtComponentObj based on the file name.
    ''' If not found, returns Nothing.
    ''' </summary>
    ''' <param name="fName"></param>
    ''' <returns></returns>
    Function GetComponentByOriginalFileName(fName As String) As InvtComponentObj
        Dim compObj As InvtComponentObj = Nothing
        For Each comp As InvtComponentObj In asmCompList
            If comp.OriginalFileName = fName Then
                compObj = comp
                Exit For
            End If
        Next
        Return compObj
    End Function

    ''' <summary>
    ''' Check if the component is already part of the assembly object list.
    ''' addQty is optional and defaults to 0.
    ''' If it exists and addQty is greater than 0, the quantity of the component will be increased by addQty.
    ''' </summary>
    ''' <param name="fName"></param>
    ''' <param name="addQty"></param>
    ''' <returns></returns>
    Function CheckIfComponentExists(fName As String, Optional addQty As Integer = 0) As Boolean
        Dim exists As Boolean = False
        For Each comp As InvtComponentObj In asmCompList
            If comp.OriginalFileName = fName Then
                exists = True
                If addQty > 0 Then
                    comp.Quantity += addQty
                End If
                Exit For
            End If
        Next
        Return exists
    End Function

    ''' <summary>
    ''' Increases the quantity of the component by qtyToAdd.
    ''' </summary>
    ''' <param name="fName"></param>
    ''' <param name="qtyToAdd"></param>
    Sub AddComponentQty(fName As String, qtyToAdd As Integer)
        For Each comp As InvtComponentObj In asmCompList
            If comp.OriginalFileName = fName Then
                comp.Quantity += qtyToAdd
                Exit For
            End If
        Next
    End Sub

End Class

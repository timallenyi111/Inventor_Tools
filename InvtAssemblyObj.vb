Public Class InvtAssemblyObj
    Dim compDefs As List(Of Inventor.ComponentDefinition)
    Dim asmDef As Inventor.AssemblyComponentDefinition
    Dim asmCompList As List(Of InvtComponentObj)
    Dim asmDoc As Inventor.AssemblyDocument
    Dim assmName As String
    Dim oFileName As String
    Dim oFilePath As String
    Dim nName As String
    Dim nFileName As String
    Dim nFilePath As String
    Dim nFullFileName As String

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

    Property OriginalFileName As String
        Get
            Return oFileName
        End Get
        Set(value As String)
            oFileName = value
        End Set
    End Property

    Property ComponentDefinitions As List(Of Inventor.ComponentDefinition)
        Get
            Return compDefs
        End Get
        Set(value As List(Of Inventor.ComponentDefinition))
            compDefs = value
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

    Property AssemblyComponentDefinition As Inventor.AssemblyComponentDefinition
        Get
            Return asmDef
        End Get
        Set(value As Inventor.AssemblyComponentDefinition)
            asmDef = value
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

    Property AssemblyDocument As Inventor.AssemblyDocument
        Get
            Return asmDoc
        End Get
        Set(value As Inventor.AssemblyDocument)
            asmDoc = value
        End Set
    End Property

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

    ReadOnly Property NewFullFileName As String
        Get
            Return nFilePath & nFileName
        End Get
        'Set(value As String)
        '    nFullFileName = value
        'End Set
    End Property

    ''' <summary>
    ''' Returns the InvtComponentObj based on the file name.
    ''' If not found, returns Nothing.
    ''' </summary>
    ''' <param name="fName"></param>
    ''' <returns></returns>
    Function GetComponentByFileName(fName As String) As InvtComponentObj
        Dim compObj As InvtComponentObj = Nothing
        For Each comp As InvtComponentObj In asmCompList
            If comp.FileName = fName Then
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
            If comp.FileName = fName Then
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
            If comp.FileName = fName Then
                comp.Quantity += qtyToAdd
                Exit For
            End If
        Next
    End Sub

End Class

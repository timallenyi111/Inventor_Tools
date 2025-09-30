Public Class InvtAssemblyObj
    Dim compDefs As List(Of Inventor.ComponentDefinition)
    Dim asmDef As Inventor.AssemblyComponentDefinition
    Dim asmObjList As List(Of InvtAssemblyObj)
    Dim asmCompList As List(Of InvtComponentObj)
    Dim compNames As List(Of String)
    Dim asmNames As List(Of String)
    Dim prtNames As List(Of String)
    Dim asmDoc As Inventor.AssemblyDocument
    Dim assmName As String
    Dim assmFileName As String
    Dim assmFilePath As String

    Public Sub New()
        compDefs = New List(Of Inventor.ComponentDefinition)()
        asmObjList = New List(Of InvtAssemblyObj)()
        compNames = New List(Of String)()
        asmNames = New List(Of String)()
        prtNames = New List(Of String)()
        asmCompList = New List(Of InvtComponentObj)()
    End Sub

    Property Name As String
        Get
            Return assmName
        End Get
        Set(value As String)
            assmName = value
        End Set
    End Property

    Property FileName As String
        Get
            Return assmFileName
        End Get
        Set(value As String)
            assmFileName = value
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

    Property FilePath As String
        Get
            Return assmFilePath
        End Get
        Set(value As String)
            assmFilePath = value
        End Set
    End Property

    Property ComponentNames As List(Of String)
        Get
            Return compNames
        End Get
        Set(value As List(Of String))
            compNames = value
        End Set
    End Property

    Property SubAssemblyObjectList As List(Of InvtAssemblyObj)
        Get
            Return asmObjList
        End Get
        Set(value As List(Of InvtAssemblyObj))
            asmObjList = value
        End Set
    End Property

    Property AssemblyNames As List(Of String)
        Get
            Return asmNames
        End Get
        Set(value As List(Of String))
            asmNames = value
        End Set
    End Property

    Property PartNames As List(Of String)
        Get
            Return prtNames
        End Get
        Set(value As List(Of String))
            prtNames = value
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

    Sub AddComponentQty(fName As String, qtyToAdd As Integer)
        For Each comp As InvtComponentObj In asmCompList
            If comp.FileName = fName Then
                comp.Quantity += qtyToAdd
                Exit For
            End If
        Next
    End Sub

End Class

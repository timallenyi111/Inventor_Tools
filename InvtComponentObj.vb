Public Class InvtComponentObj
    Dim compName As String
    Dim compType As String
    Dim pObj As New InvtPartObj
    Dim aObj As New InvtAssemblyObj
    Dim qty As Integer
    Dim fName As String

    Public Sub New()
        Name = ""
        Type = ""
        qty = 1
    End Sub
    Property Name As String
        Get
            Return compName
        End Get
        Set(value As String)
            compName = value
        End Set
    End Property

    Property Type As String
        Get
            Return compType
        End Get
        Set(value As String)
            compType = value
        End Set
    End Property

    Property PartObject As InvtPartObj
        Get
            Return pObj
        End Get
        Set(value As InvtPartObj)
            pObj = value
        End Set
    End Property

    Property AssemblyObject As InvtAssemblyObj
        Get
            Return aObj
        End Get
        Set(value As InvtAssemblyObj)
            aObj = value
        End Set
    End Property

    Property Quantity As Integer
        Get
            Return qty
        End Get
        Set(value As Integer)
            qty = value
        End Set
    End Property

    ''' <summary>
    ''' The complete file name of the component
    ''' Example: **path**\***assembly Name ***.iam
    ''' </summary>
    ''' <returns></returns>
    Property FileName As String
        Get
            Return fName
        End Get
        Set(value As String)
            fName = value
        End Set
    End Property

    Sub AddQuantity(qtyToAdd As Integer)
        qty += qtyToAdd
    End Sub

End Class

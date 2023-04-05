<DebuggerStepThrough()>
Public Class ReturnStatus
#Region "variables"
    Private _BlnReturn As Boolean = False
    Private _StrReturn As String = String.Empty
    Private _ObjReturn As Object = Nothing
#End Region
#Region "Property"
    Public ReadOnly Property BlnReturn As Boolean
        Get
            Return _BlnReturn
        End Get
    End Property
    Public ReadOnly Property StrReturn As String
        Get
            Return _StrReturn
        End Get
    End Property
    Public ReadOnly Property ObjeReturn As Object
        Get
            Return _ObjReturn
        End Get
    End Property
#End Region
    Sub New()

    End Sub
    Sub New(_bln As Boolean)
        _BlnReturn = _bln
    End Sub
    Sub New(_bln As Boolean, _str As String)
        _BlnReturn = _bln : _StrReturn = _str
    End Sub
    Sub New(_bln As Boolean, _obj As Object)
        _BlnReturn = _bln : _ObjReturn = _obj
    End Sub
End Class

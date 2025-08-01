Public Class DocumentosTrans
    Private _Code As String
    Public Property Code() As String
        Get
            Return _Code
        End Get
        Set(ByVal value As String)
            _Code = value
        End Set
    End Property

    Private _DocEntry As Integer
    Public Property DocEntry() As Integer
        Get
            Return _DocEntry
        End Get
        Set(ByVal value As Integer)
            _DocEntry = value
        End Set
    End Property

    Private _ObjectType As Integer
    Public Property ObjectType() As Integer
        Get
            Return _ObjectType
        End Get
        Set(ByVal value As Integer)
            _ObjectType = value
        End Set
    End Property

    Private _DocSubType As String
    Public Property DocSubType() As String
        Get
            Return _DocSubType
        End Get
        Set(ByVal value As String)
            _DocSubType = value
        End Set
    End Property

    Private _SRI_Code As String
    Public Property SRI_Code() As String
        Get
            Return _SRI_Code
        End Get
        Set(ByVal value As String)
            _SRI_Code = value
        End Set
    End Property

    Private _Procesado As Integer
    Public Property Procesado() As Integer
        Get
            Return _Procesado
        End Get
        Set(ByVal value As Integer)
            _Procesado = value
        End Set
    End Property

    Private _Oberva As String
    Public Property Oberva() As String
        Get
            Return _Oberva
        End Get
        Set(ByVal value As String)
            _Oberva = value
        End Set
    End Property

    Sub New(Code As String, DocEntry As Integer, ObjectType As Integer, DocSubType As String, SRICode As String, Procesado As Integer, Observacion As String)
        _Code = Code
        _DocEntry = DocEntry
        _ObjectType = ObjectType
        _SRI_Code = SRI_Code
        _Procesado = Procesado
        _Oberva = Observacion
    End Sub

    Sub New()

    End Sub
End Class

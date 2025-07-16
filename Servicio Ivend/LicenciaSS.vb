Public Class LicenciaSS
    Private _NombreBaseSAP As String
    Public Property NombreBaseSAP() As String
        Get
            Return _NombreBaseSAP
        End Get
        Set(ByVal value As String)
            _NombreBaseSAP = value
        End Set
    End Property
    Private _Estado As Boolean
    Public Property Estado() As Boolean
        Get
            Return _Estado
        End Get
        Set(ByVal value As Boolean)
            _Estado = value
        End Set
    End Property
    Private _Opcion As String
    Public Property Opcion() As String
        Get
            Return _Opcion
        End Get
        Set(ByVal value As String)
            _Opcion = value
        End Set
    End Property
    Private _validoHasta As Integer
    Public Property validoHasta() As Integer
        Get
            Return _validoHasta
        End Get
        Set(ByVal value As Integer)
            _validoHasta = value
        End Set
    End Property

    Sub New(NombreBaseSAP As String, Opcion As String, ValidoHasta As Integer)
        _NombreBaseSAP = NombreBaseSAP
        _Opcion = Opcion
        _validoHasta = ValidoHasta
    End Sub
    Sub New()

    End Sub
End Class

Public Class Configuracion

    Private _Modulo As String
    Public Property Modulo() As String
        Get
            Return _Modulo
        End Get
        Set(ByVal value As String)
            _Modulo = value
        End Set
    End Property
    Private _Tipo As String
    Public Property Tipo() As String
        Get
            Return _Tipo
        End Get
        Set(ByVal value As String)
            _Tipo = value
        End Set
    End Property
    Private _SubTipo As String
    Public Property SubTipo() As String
        Get
            Return _SubTipo
        End Get
        Set(ByVal value As String)
            _SubTipo = value
        End Set
    End Property
    Private _Detalle As List(Of ConfiguracionDetalle)
    Public Property Detalle() As List(Of ConfiguracionDetalle)
        Get
            Return _Detalle
        End Get
        Set(ByVal value As List(Of ConfiguracionDetalle))
            _Detalle = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(ByVal Modulo As String, ByVal Tipo As String, ByVal SubTipo As String, ByVal Detalle As List(Of ConfiguracionDetalle))
        _Modulo = Modulo
        _Tipo = Tipo
        _SubTipo = SubTipo
        _Detalle = Detalle
    End Sub
End Class

Public Class ConfiguracionDetalle

    Private _Nombre As String
    Public Property Nombre() As String
        Get
            Return _Nombre
        End Get
        Set(ByVal value As String)
            _Nombre = value
        End Set
    End Property

    Private _Valor As String
    Public Property Valor() As String
        Get
            Return _Valor
        End Get
        Set(ByVal value As String)
            _Valor = value
        End Set
    End Property

    Sub New()

    End Sub
    Sub New(ByVal Nombre As String, ByVal Valor As String)
        _Nombre = Nombre
        _Valor = Valor
    End Sub
End Class

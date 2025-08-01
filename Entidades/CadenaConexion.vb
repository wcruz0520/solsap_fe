Public Class CadenaConexion

    Private _Api As String
    Public Property API() As String
        Get
            Return _Api
        End Get
        Set(ByVal value As String)
            _Api = value
        End Set
    End Property


    Private _Servidor As String
    Public Property Servidor() As String
        Get
            Return _Servidor
        End Get
        Set(ByVal value As String)
            _Servidor = value
        End Set
    End Property

    Private _BaseDeDatos As String
    Public Property BaseDeDatos() As String
        Get
            Return _BaseDeDatos
        End Get
        Set(ByVal value As String)
            _BaseDeDatos = value
        End Set
    End Property

    Private _Usuario As String
    Public Property Usuario() As String
        Get
            Return _Usuario
        End Get
        Set(ByVal value As String)
            _Usuario = value
        End Set
    End Property

    Private _Contraseña As String
    Public Property Contraseña() As String
        Get
            Return _Contraseña
        End Get
        Set(ByVal value As String)
            _Contraseña = value
        End Set
    End Property

    Private _SeguridadIntegrada As Boolean
    Public Property SeguridadIntegrada() As Boolean
        Get
            Return _SeguridadIntegrada
        End Get
        Set(ByVal value As Boolean)
            _SeguridadIntegrada = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(Api As String, Servidor As String, BaseDeDatos As String, Usuario As String, Contraseña As String, SeguridadIntegrada As Boolean)
        _Api = Api
        _Servidor = Servidor
        _BaseDeDatos = BaseDeDatos
        _Usuario = Usuario
        _Contraseña = Contraseña
        _SeguridadIntegrada = SeguridadIntegrada
    End Sub

End Class

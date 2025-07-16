Public Class DocumentoTipo


    Private _Sigla As String
    Public Property Sigla() As String
        Get
            Return _Sigla
        End Get
        Set(ByVal value As String)
            _Sigla = value
        End Set
    End Property

    Private _Descripcion As String
    Public Property Descripcion() As String
        Get
            Return _Descripcion
        End Get
        Set(ByVal value As String)
            _Descripcion = value
        End Set
    End Property

    Private _ObjectType As String
    Public Property ObjectType() As String
        Get
            Return _ObjectType
        End Get
        Set(ByVal value As String)
            _ObjectType = value
        End Set
    End Property

    Private _EsAdministrador As String
    Public Property EsAdministrador() As String
        Get
            Return _EsAdministrador
        End Get
        Set(ByVal value As String)
            _EsAdministrador = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(ByVal Sigla As String, Descripcion As String, ObjectType As String, EsAdministrador As String)
        _Sigla = Sigla
        _Descripcion = Descripcion
        _ObjectType = ObjectType
        _EsAdministrador = EsAdministrador
    End Sub


End Class

Public Class Usuario

    Private _UserCode As String
    Public Property UserCode() As String
        Get
            Return _UserCode
        End Get
        Set(ByVal value As String)
            _UserCode = value
        End Set
    End Property

    Private _UserName As String
    Public Property UserName() As String
        Get
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property

    Private _dpto As String
    Public Property Dpto() As String
        Get
            Return _dpto
        End Get
        Set(ByVal value As String)
            _dpto = value
        End Set
    End Property

    Private _DescripcionDpto As String
    Public Property DescripcionDpto() As String
        Get
            Return _DescripcionDpto
        End Get
        Set(ByVal value As String)
            _DescripcionDpto = value
        End Set
    End Property

    Sub New(usercode As String, username As String, dpto As String, descripciondpto As String)
        _UserCode = usercode
        _UserName = username
        _dpto = dpto
        _DescripcionDpto = descripciondpto
    End Sub

End Class

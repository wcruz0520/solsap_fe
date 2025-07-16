Public Class clsError

    Private _ErrorEstado As Boolean
    Public Property ErrorEstado() As Boolean
        Get
            Return _ErrorEstado
        End Get
        Set(ByVal value As Boolean)
            _ErrorEstado = value
        End Set
    End Property

    Private _ErrorDescripcion As String
    Public Property ErrorDescripcion() As String
        Get
            Return _ErrorDescripcion
        End Get
        Set(ByVal value As String)
            _ErrorDescripcion = value
        End Set
    End Property


    Sub New(ErrorEstado As Boolean, ErrorDescripcion As String)
        _ErrorEstado = ErrorEstado
        _ErrorDescripcion = ErrorDescripcion
    End Sub

    Sub New()

    End Sub


End Class

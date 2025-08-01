Public Class FacturaVenta

    Private _DocEntry As Integer
    Public Property DocEntry() As Integer
        Get
            Return _DocEntry
        End Get
        Set(ByVal value As Integer)
            _DocEntry = value
        End Set
    End Property

    Private _ValorARetener As Double
    Public Property ValorARetener() As Double
        Get
            Return _ValorARetener
        End Get
        Set(ByVal value As Double)
            _ValorARetener = value
        End Set
    End Property

    Private _Cuota As Integer
    Public Property Cuota() As Integer
        Get
            Return _Cuota
        End Get
        Set(ByVal value As Integer)
            _Cuota = value
        End Set
    End Property

    Private _Name As String
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property
    Sub New()

    End Sub

    Sub New(DocEnytry As Integer, ValorARetener As Double, Cuota As Integer, Name As String)
        _DocEntry = DocEntry
        _ValorARetener = ValorARetener
        _Cuota = Cuota
        _Name = Name
    End Sub


End Class

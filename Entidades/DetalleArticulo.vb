Public Class DetalleArticulo
    Private _claveAcceso As String
    Public Property ClaveAcceso() As String
        Get
            Return _claveAcceso
        End Get
        Set(ByVal value As String)
            _claveAcceso = value
        End Set
    End Property

    Private _codigoPrincipal As String
    Public Property CodigoPrincipal() As String
        Get
            Return _codigoPrincipal
        End Get
        Set(ByVal value As String)
            _codigoPrincipal = value
        End Set
    End Property

    Private _CodigoAuxiliar As String
    Public Property CodigoAuxiliar() As String
        Get
            Return _CodigoAuxiliar
        End Get
        Set(ByVal value As String)
            _CodigoAuxiliar = value
        End Set
    End Property

    Private _descripcion As String
    Public Property Descripcion() As String
        Get
            Return _descripcion
        End Get
        Set(ByVal value As String)
            _descripcion = value
        End Set
    End Property

    Private _Cantidad As Decimal
    Public Property Cantidad() As Decimal
        Get
            Return _Cantidad
        End Get
        Set(ByVal value As Decimal)
            _Cantidad = value
        End Set
    End Property

    Private _PrecioUnitario As Decimal
    Public Property PrecioUnitario() As Decimal
        Get
            Return _PrecioUnitario
        End Get
        Set(ByVal value As Decimal)
            _PrecioUnitario = value
        End Set
    End Property

    Private _Descuento As Decimal
    Public Property Descuento() As Decimal
        Get
            Return _Descuento
        End Get
        Set(ByVal value As Decimal)
            _Descuento = value
        End Set
    End Property

    Private _PrecioTotalSinImpuesto As Decimal
    Public Property PrecioTotalSinImpuesto() As Decimal
        Get
            Return _PrecioTotalSinImpuesto
        End Get
        Set(ByVal value As Decimal)
            _PrecioTotalSinImpuesto = value
        End Set
    End Property


    Sub New(ClaveAcceso As String, codigoPrincipal As String, CodigoAuxiliar As String, descripcion As String, _
            Cantidad As String, precioUnitario As Decimal, Descuento As Decimal, PrecioTotalSinImpuesto As String)
        _claveAcceso = ClaveAcceso
        _codigoPrincipal = codigoPrincipal
        _CodigoAuxiliar = CodigoAuxiliar
        _descripcion = descripcion
        _Cantidad = Cantidad
        _Descuento = Descuento
        _PrecioUnitario = precioUnitario
        _PrecioTotalSinImpuesto = PrecioTotalSinImpuesto

    End Sub

    Sub New()

    End Sub

End Class

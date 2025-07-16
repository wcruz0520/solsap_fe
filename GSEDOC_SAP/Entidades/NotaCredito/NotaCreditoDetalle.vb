Public Class NotaCreditoDetalle

    Public Property _codigoInterno As String
    Public Property _codigoAdicional As String
    Public Property _descripcion As String
    Public Property _cantidad As Decimal
    Public Property _precioUnitario As Decimal
    Public Property _descuento As Decimal
    Public Property _precioTotalSinImpuesto As Decimal
    Public Property _impuestos As List(Of NotaCreditoDetalleImpuesto)

End Class

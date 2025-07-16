Public Class NotaCreditoDetalle

    Public Property _codigoInterno As String
    Public Property _codigoAdicional As String
    Public Property _descripcion As String
    Public Property _cantidad As String
    Public Property _precioUnitario As String
    Public Property _descuento As String
    Public Property _precioTotalSinImpuesto As String
    Public Property _impuestos As List(Of NotaCreditoDetalleImpuesto)

End Class

Public Class FacturaDetalle

    Public Property _codigoPrincipal As String
    Public Property _codigoAuxiliar As String
    Public Property _descripcion As String
    Public Property _cantidad As Decimal
    Public Property _precioUnitario As Decimal
    Public Property _descuento As Decimal
    Public Property _precioTotalSinImpuesto As Decimal

    Public Property _impuestos As List(Of FacturaDetalleImpuesto)


End Class

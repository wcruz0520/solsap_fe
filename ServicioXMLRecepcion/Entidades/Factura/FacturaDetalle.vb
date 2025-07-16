Public Class FacturaDetalle

    Public Property _codigoPrincipal As String
    Public Property _codigoAuxiliar As String
    Public Property _descripcion As String
    Public Property _cantidad As String
    Public Property _precioUnitario As String
    Public Property _descuento As String
    Public Property _precioTotalSinImpuesto As String

    Public Property _impuestos As List(Of FacturaDetalleImpuesto)


End Class

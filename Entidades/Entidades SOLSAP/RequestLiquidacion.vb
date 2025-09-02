Public Class RequestLiquidacion
    Public Property infoTributaria As infoTributariaLQ
    Public Property infoLiquidacionCompra As infoLiquidacionCompraLQ
    Public Property detalles As List(Of detalleLQ)
    Public Property reembolsos As List(Of reembolsoLQ)
    Public Property infoAdicional As List(Of campoAdicionalLQ)
    Public Property campoAdicional1 As String
    Public Property campoAdicional2 As String
End Class

Public Class infoTributariaLQ
    Public Property ambiente As String
    Public Property tipoEmision As String
    Public Property claveAcceso As String
    Public Property razonSocial As String
    Public Property nombreComercial As String
    Public Property ruc As String
    Public Property codDoc As String
    Public Property estab As String
    Public Property ptoEmi As String
    Public Property secuencial As String
    Public Property dirMatriz As String
    Public Property diaEmission As String
    Public Property mesEmission As String
    Public Property anioEmission As String
    'Public Property campoAdicional1 As String
    'Public Property campoAdicional2 As String
End Class

Public Class infoLiquidacionCompraLQ
    Public Property fechaEmision As String
    Public Property dirEstablecimiento As String
    Public Property contribuyenteEspecial As String
    Public Property obligadoContabilidad As String
    Public Property tipoIdentificacionProveedor As String
    Public Property razonSocialProveedor As String
    Public Property identificacionProveedor As String
    Public Property direccionProveedor As String
    Public Property totalSinImpuestos As String
    Public Property totalDescuento As String
    Public Property codDocReembolso As String
    Public Property totalComprobantesReembolso As String
    Public Property totalBaseImponibleReembolso As String
    Public Property totalImpuestoReembolso As String
    Public Property totalConImpuestos As List(Of totalConImpuestoLQ)
    Public Property importeTotal As String
    Public Property moneda As String
    Public Property pagos As List(Of pagoLQ)
End Class

Public Class totalConImpuestoLQ
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
    Public Property tarifa As String
    Public Property descuentoAdicional As String
End Class

Public Class pagoLQ
    Public Property formaPago As String
    Public Property total As String
    Public Property plazo As String
    Public Property unidadTiempo As String
End Class

Public Class detalleLQ
    Public Property codigoPrincipal As String
    Public Property codigoAuxiliar As String
    Public Property descripcion As String
    Public Property cantidad As Decimal
    Public Property precioUnitario As String
    Public Property descuento As String
    Public Property precioTotalSinImpuesto As String
    Public Property detallesAdicionales As List(Of detallesAdicionalLQ)
    Public Property impuestos As List(Of impuestoLQ)
    Public Property unidadMedida As String
End Class

Public Class detallesAdicionalLQ
    Public Property nombre As String
    Public Property valor As String
End Class

Public Class impuestoLQ
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
    Public Property tarifa As String
End Class

Public Class reembolsoLQ
    Public Property tipoIdentificacionProveedorReembolso As String
    Public Property identificacionProveedorReembolso As String
    Public Property codPaisPagoProveedorReembolso As String
    Public Property tipoProveedorReembolso As String
    Public Property codDocReembolso As String
    Public Property estabDocReembolso As String
    Public Property ptoEmiDocReembolso As String
    Public Property secuencialDocReembolso As String
    Public Property fechaEmisionDocReembolso As String
    Public Property numeroautorizacionDocReemb As String
    Public Property detalleImpuestos As List(Of detalleImpuestoReembolsoLQ)
End Class

Public Class detalleImpuestoReembolsoLQ
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property tarifa As String
    Public Property baseImponibleReembolso As String
    Public Property impuestoReembolso As String
End Class

Public Class campoAdicionalLQ
    Public Property nombre As String
    Public Property valor As String
End Class
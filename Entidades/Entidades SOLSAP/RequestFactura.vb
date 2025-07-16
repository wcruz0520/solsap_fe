Public Class RequestFactura
    Public Property infoTributaria As infoTributariaFE
    Public Property infoFactura As infoFacturaFE
    Public Property detalles As List(Of detalleFE)
    Public Property infoAdicional As List(Of infoAdicionalFE)
End Class

Public Class infoTributariaFE
    Public Property ambiente As String
    Public Property claveAcceso As String
    Public Property razonSocial As String
    Public Property nombreComercial As String
    Public Property ruc As String
    Public Property tipoEmision As String
    Public Property codDoc As String
    Public Property estab As String
    Public Property ptoEmi As String
    Public Property secuencial As String
    Public Property dirMatriz As String
    Public Property diaEmission As String
    Public Property mesEmission As String
    Public Property anioEmission As String
End Class

Public Class infoFacturaFE
    Public Property fechaEmision As String
    Public Property dirEstablecimiento As String
    Public Property contribuyenteEspecial As String
    Public Property obligadoContabilidad As String
    Public Property comercioExterior As String
    Public Property IncoTermFactura As String
    Public Property lugarIncoTerm As String
    Public Property paisOrigen As String
    Public Property puertoEmbarque As String
    Public Property paisDestino As String
    Public Property paisAdquisicion As String
    Public Property tipoIdentificacionComprador As String
    Public Property guiaRemision As String
    Public Property razonSocialComprador As String
    Public Property identificacionComprador As String
    Public Property direccionComprador As String
    Public Property totalSinImpuestos As String
    Public Property incoTermTotalSinImpuestos As String
    Public Property totalDescuento As String
    Public Property totalConImpuestos As List(Of totalConImpuestosFE)
    Public Property propina As String
    Public Property fleteInternacional As String
    Public Property seguroInternacional As String
    Public Property gastosAduaneros As String
    Public Property gastosTransporteOtros As String
    Public Property importeTotal As String
    Public Property moneda As String
    Public Property pagos As List(Of pagosFE)
    Public Property valorRetIva As String
    Public Property valorRetRenta As String
End Class

Public Class totalConImpuestosFE
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
End Class

Public Class pagosFE
    Public Property formaPago As String
    Public Property total As String
    Public Property plazo As String
    Public Property unidadTiempo As String
End Class

Public Class detalleFE
    Public Property codigoPrincipal As String
    Public Property codigoAuxiliar As String
    Public Property descripcion As String
    Public Property cantidad As Integer
    Public Property precioUnitario As String
    Public Property descuento As String
    Public Property precioTotalSinImpuesto As String
    Public Property detallesAdicionales As List(Of detallesAdicionalesFE)
    Public Property impuestos As List(Of impuestosFE)
End Class

Public Class detallesAdicionalesFE
    Public Property nombre As String
    Public Property valor As String
End Class

Public Class impuestosFE
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
    Public Property tarifa As String
End Class

Public Class infoAdicionalFE
    Public Property nombre As String
    Public Property valor As String
End Class

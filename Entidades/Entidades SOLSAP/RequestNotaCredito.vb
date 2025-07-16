Public Class RequestNotaCredito
    Public Property infoTributaria As infoTributariaNCE
    Public Property infoNotaCredito As infoNotaCreditoNCE
    Public Property detalles As List(Of detalleNCE)
    Public Property infoAdicional As List(Of infoAdicionalNCE)
End Class

Public Class infoTributariaNCE
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

Public Class infoNotaCreditoNCE
    Public Property fechaEmision As String
    Public Property dirEstablecimiento As String
    Public Property tipoIdentificacionComprador As String
    Public Property razonSocialComprador As String
    Public Property identificacionComprador As String
    Public Property contribuyenteEspecial As String
    Public Property obligadoContabilidad As String
    Public Property rise As String
    Public Property codDocModificado As String
    Public Property numDocModificado As String
    Public Property fechaEmisionDocSustento As String
    Public Property totalSinImpuestos As String
    Public Property valorModificacion As String
    Public Property moneda As String
    Public Property totalConImpuestos As List(Of totalConImpuestosNCE)
    Public Property motivo As String
End Class

Public Class totalConImpuestosNCE
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
End Class

Public Class detalleNCE
    Public Property codigoPrincipal As String
    Public Property codigoAuxiliar As String
    Public Property descripcion As String
    Public Property cantidad As Integer
    Public Property precioUnitario As String
    Public Property descuento As String
    Public Property precioTotalSinImpuesto As String
    Public Property detallesAdicionales As List(Of detallesAdicionalesNCE)
    Public Property impuestos As List(Of impuestosNCE)
End Class

Public Class detallesAdicionalesNCE
    Public Property nombre As String
    Public Property valor As String
End Class

Public Class impuestosNCE
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
    Public Property tarifa As String
End Class

Public Class infoAdicionalNCE
    Public Property nombre As String
    Public Property valor As String
End Class
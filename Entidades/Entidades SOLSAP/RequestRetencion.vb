Public Class RequestRetencion
    Public Property infoTributaria As infoTributariaRET
    Public Property infoCompRetencion As infoNotaCreditoRET
    Public Property impuestos As List(Of impuestosRET)
    Public Property infoAdicional As List(Of infoAdicionalRET)
End Class

Public Class infoTributariaRET
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

Public Class infoNotaCreditoRET
    Public Property fechaEmision As String
    Public Property dirEstablecimiento As String
    Public Property contribuyenteEspecial As String
    Public Property obligadoContabilidad As String
    Public Property tipoIdentificacionSujetoRetenido As String
    Public Property razonSocialSujetoRetenido As String
    Public Property identificacionSujetoRetenido As String
    Public Property periodoFiscal As String
End Class

Public Class impuestosRET
    Public Property codigo As String
    Public Property codigoRetencion As String
    Public Property baseImponible As String
    Public Property porcentajeRetener As String
    Public Property valorRetenido As String
    Public Property codDocSustento As String
    Public Property numDocSustento As String
    Public Property fechaEmisionDocSustento As String
End Class

Public Class infoAdicionalRET
    Public Property nombre As String
    Public Property valor As String
End Class
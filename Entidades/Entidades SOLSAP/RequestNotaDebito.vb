Public Class RequestNotaDebito
    Public Property infoTributaria As infoTributariaND
    Public Property infoNotaDebito As infoNotaDebitoND
    Public Property motivos As List(Of motivoND)
    Public Property infoAdicional As List(Of infoAdicionalND)
End Class

Public Class infoTributariaND
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
    Public Property campoAdicional1 As String
    Public Property campoAdicional2 As String
End Class

Public Class infoNotaDebitoND
    Public Property fechaEmision As String
    Public Property dirEstablecimiento As String
    Public Property tipoIdentificacionComprador As String
    Public Property razonSocialComprador As String
    Public Property identificacionComprador As String
    Public Property contribuyenteEspecial As String
    Public Property obligadoContabilidad As String
    Public Property codDocModificado As String
    Public Property numDocModificado As String
    Public Property fechaEmisionDocSustento As String
    Public Property totalSinImpuestos As String
    Public Property impuestos As List(Of impuestoND)
    Public Property valorTotal As String
    Public Property pagos As List(Of pagoND)
End Class

Public Class impuestoND
    Public Property codigo As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property valor As String
    Public Property tarifa As String
End Class

Public Class pagoND
    Public Property formaPago As String
    Public Property total As String
    Public Property plazo As String
    Public Property unidadTiempo As String
End Class

Public Class motivoND
    Public Property razon As String
    Public Property valor As String
End Class

Public Class infoAdicionalND
    Public Property nombre As String
    Public Property valor As String
End Class
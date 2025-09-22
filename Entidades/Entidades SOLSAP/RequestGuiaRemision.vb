' Clase principal del request
Public Class RequestGuiaRemision
    Public Property infoTributaria As infoTributariaGR
    Public Property infoGuiaRemision As infoGuiaRemisionGR
    Public Property destinatarios As List(Of destinatarioGR)
    Public Property infoAdicional As List(Of infoAdicionalGR)
    Public Property campoAdicional1 As String
    Public Property campoAdicional2 As String
End Class

' Info Tributaria
Public Class infoTributariaGR
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
End Class

' Info Guía Remisión
Public Class infoGuiaRemisionGR
    Public Property dirEstablecimiento As String
    Public Property dirPartida As String
    Public Property razonSocialTransportista As String
    Public Property tipoIdentificacionTransportista As String
    Public Property rucTransportista As String
    Public Property rise As String
    Public Property obligadoContabilidad As String
    Public Property contribuyenteEspecial As String
    Public Property fechaIniTransporte As String
    Public Property fechaFinTransporte As String
    Public Property placa As String
End Class

' Destinatario
Public Class destinatarioGR
    Public Property identificacionDestinatario As String
    Public Property razonSocialDestinatario As String
    Public Property dirDestinatario As String
    Public Property motivoTraslado As String
    Public Property codEstabDestino As String
    Public Property codDocSustento As String
    Public Property numDocSustento As String
    Public Property numAutDocSustento As String
    Public Property fechaEmisionDocSustento As String
    Public Property docAduaneroUnico As String
    Public Property ruta As String
    Public Property detalles As List(Of detalleGR)
End Class

' Detalles dentro de cada destinatario
Public Class detalleGR
    Public Property codigoInterno As String
    Public Property codigoAdicional As String
    Public Property descripcion As String
    Public Property cantidad As Integer
    Public Property detallesAdicionales As List(Of detalleAdicionalGR)
End Class

' Detalles Adicionales de un item
Public Class detalleAdicionalGR
    Public Property nombre As String
    Public Property valor As String
End Class

' Info Adicional (al nivel de la guía)
Public Class infoAdicionalGR
    Public Property nombre As String
    Public Property valor As String
End Class


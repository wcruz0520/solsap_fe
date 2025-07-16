Public Class SalidaCabecera
    Public Property DocEntrySalida As String
    Public Property Fecha As Date

    Public Property EmpresaOrigen As String
    Public Property AlmacenOrigen As String
    Public Property EmpresaDestino As String
    Public Property AlmacenDestino As String
    Public Property comentario As String
    Public Property Detalles As List(Of SalidaDetalle)

End Class

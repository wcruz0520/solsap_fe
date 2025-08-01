Public Class NotaCreditoCabecera

    Public Property _DocEntry As Integer
    Public Property _NumeroAutorizacion As String
    Public Property _FechaAutorizacion As String
    Public Property _RazonSocial As String
    Public Property _ruc As String
    Public Property _claveAcceso As String
    Public Property _estab As String
    Public Property _ptoEmi As String
    Public Property _secuencial As String
    Public Property _fechaEmision As String
    Public Property _dirEstablecimiento As String
    Public Property _contribuyenteEspecial As String
    Public Property _razonSocialComprador As String
    Public Property _identificacionComprador As String
    Public Property _CodDocMod As String

    Public Property _numDocModificado As String
    Public Property _fechaEmisionDocSustento As String
    Public Property _valorModificacion As Decimal
    Public Property _direccionComprador As String
    Public Property _totalSinImpuestos As Decimal
    Public Property _totalDescuento As Decimal
    Public Property _motivo As String

    Public Property _importeTotal As Decimal
    Public Property _impuestos As List(Of NotaCreditoCabeceraImpuesto)

End Class

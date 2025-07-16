Public Class FacturaCabecera

    Public Property _DocEntry As Integer
    Public Property _NumeroAutorizacion As String
    Public Property _FechaAutorizacion As String
    Public Property _RazonSocial As String
    Public Property _ruc As String
    Public Property _claveAcceso As String
    Public Property _estab As String
    Public Property _ptoEmi As String
    Public Property _secuencial As String
    Public Property _fechaEmision As Date
    Public Property _dirEstablecimiento As String
    Public Property _contribuyenteEspecial As String
    Public Property _razonSocialComprador As String
    Public Property _identificacionComprador As String
    Public Property _direccionComprador As String
    Public Property _totalSinImpuestos As Decimal
    Public Property _totalDescuento As Decimal
    Public Property _importeTotal As Decimal
    Public Property _formaPago As String
    Public Property _totalFormaPago As Decimal
    Public Property _plazo As System.Nullable(Of Integer)
    Public Property _unidadTiempo As String
    Public Property _impuestos As List(Of FacturaCabeceraImpuestos)


End Class

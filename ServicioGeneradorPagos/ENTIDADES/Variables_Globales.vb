Public Class Variables_Globales

    Public Shared WS_Recepcion As String
    Public Shared WS_RecepcionCambiarEstado As String
    Public Shared WS_RecepcionClave As String
    Public Shared WS_RecepcionCargaEstados As String

    Public Shared SALIDA_POR_PROXY
    Public Shared Proxy_puerto
    Public Shared Proxy_IP
    Public Shared Proxy_Usuario
    Public Shared Proxy_Clave

    Public Shared Nombre_Proveedor

    Public Shared CampoNumRetencion
    Public Shared FechaEmisionRetencion
    Public Shared FechaEmisionRetencionP

    Public Shared _ValidarFechasCTK
    Public Shared _vgFechaEmisionRetencion
    Public Shared _vgFechaEmisionRetencionP

    Public Shared Dias

    Public Shared CantDiasLab
    Public Shared CantUltmsDias

    Public Shared ContSaldoPendMenor
    Public Shared CuentaSaldoFavor

    Public Shared IdSeriePR

    'Public Shared FechaEmisionRetencion

    Structure PROVEEDOR_DE_SAPBO
        Const EXXIS = "EXXIS"
        Const ONESOLUTIONS = "ONESOLUTIONS"
        Const SYPSOFT = "SYPSOFT"
        Const HEINSOHN = "HEINSOHN"
        Const TOPMANAGE = "TOPMANAGE"
        Const SOLSAP = "SOLSAP"
    End Structure

End Class

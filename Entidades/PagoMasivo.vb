Public Class PagoMasivo
    Public Property DocEntryFP As Integer 'U_DocEntry
    Public Property CodPro As String 'U_CodProv
    Public Property NomPro As String 'U_Prov
    Public Property Vencimiento As String 'U_Venc
    Public Property FechaVencimiento As String 'U_FecVec
    Public Property Monto As String 'U_Mon
    Public Property Saldo As String 'U_Sal
    Public Property Pagar As String 'U_Pag
    Public Property Cuota As String 'U_Cuo
    Public Property Sucursal As String 'U_Suc
    Public Property Proyecto As String 'U_Proy
    Public Property ObjType As String 'U_ObjType
    Public Property Cuenta As String
    Public Property DocEntry As Integer
    Public Property NumCheque As String
    Public Property MedioPago As String
    Public Property Banco As String
    Public Property Procesada As String
    Public Property LineId As String

    Public Property Consolidado As Boolean

End Class

Public Class SS_PM_AUT
    Public Property Usuario As String
    Public Property Nivel As String
End Class
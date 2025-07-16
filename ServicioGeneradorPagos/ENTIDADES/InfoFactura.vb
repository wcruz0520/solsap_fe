Public Class InfoFactura

    Public Property DocEntry As String

    Public Property NumFactura As String

    Public Property NumRet As String
    Public Property Subtotal As Decimal
    Public Property Iva As Decimal

    Public Property RetRenta As Decimal

    Public Property RetIva As Decimal


    Public Sub New(docEntry As String, numfactura As String, docnumRet As String, subtotal As Decimal, iva As Decimal, retrenta As Decimal, retiva As Decimal)
        Me.DocEntry = docEntry
        Me.NumFactura = numfactura
        Me.NumRet = docnumRet
        Me.Subtotal = subtotal
        Me.Iva = iva
        Me.RetRenta = retrenta
        Me.RetIva = retiva
    End Sub

End Class

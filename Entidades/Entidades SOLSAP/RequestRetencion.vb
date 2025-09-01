Public Class RequestRetencion
    Public Property infoTributaria As InfoTributariaRET
    Public Property infoCompRetencion As InfoCompRetencionRET
    Public Property docsSustento As List(Of DocSustentoRET)
    Public Property infoAdicional As List(Of InfoAdicionalRET)
End Class

Public Class InfoTributariaRET
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

Public Class InfoCompRetencionRET
    Public Property fechaEmision As String
    Public Property dirEstablecimiento As String
    Public Property contribuyenteEspecial As String
    Public Property obligadoContabilidad As String
    Public Property tipoIdentificacionSujetoRetenido As String
    Public Property tipoSujetoRetenido As String
    Public Property parteRel As String
    Public Property razonSocialSujetoRetenido As String
    Public Property identificacionSujetoRetenido As String
    Public Property periodoFiscal As String
End Class

Public Class DocSustentoRET
    Public Property codSustento As String
    Public Property codDocSustento As String
    Public Property numDocSustento As String
    Public Property factura_relacionada As String
    Public Property fechaEmisionDocSustento As String
    Public Property fechaRegistroContable As String
    Public Property numAutDocSustento As String
    Public Property pagoLocExt As String
    Public Property tipoRegi As String
    Public Property paisEfecPago As String
    Public Property aplicConvDobTrib As String
    Public Property pagExtSujRetNorLeg As String
    Public Property pagRegFis As String
    Public Property totalComprobantesReembolso As String
    Public Property totalBaseImponibleReembolso As String
    Public Property totalSinImpuestos As String
    Public Property importeTotal As String
    Public Property impuestosDocSustento As List(Of ImpuestoDocSustentoRET)
    Public Property retenciones As List(Of RetencionRET)
    Public Property pagos As List(Of PagoRET)
End Class

Public Class ImpuestoDocSustentoRET
    Public Property codImpuestoDocSustento As String
    Public Property codigoPorcentaje As String
    Public Property baseImponible As String
    Public Property tarifa As String
    Public Property valorImpuesto As String
End Class

Public Class RetencionRET
    Public Property codigo As String
    Public Property codigoRetencion As String
    Public Property baseImponible As String
    Public Property porcentajeRetener As String
    Public Property valorRetenido As String
    Public Property dividendos As DividendosRET
End Class

Public Class DividendosRET
    Public Property fechaPagoDiv As String
    Public Property imRentaSoc As String
    Public Property ejerFisUtDiv As String
End Class

Public Class PagoRET
    Public Property formaPago As String
    Public Property total As String
End Class

Public Class InfoAdicionalRET
    Public Property nombre As String
    Public Property valor As String
End Class

' NOTA: El código generado puede requerir, como mínimo, .NET Framework 4.5 o .NET Core/Standard 2.0.
'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False)>
Partial Public Class comprobanteRetencion

    Private infoTributariaField As comprobanteRetencionInfoTributaria

    Private infoCompRetencionField As comprobanteRetencionInfoCompRetencion

    Private impuestosField() As comprobanteRetencionImpuesto

    Private docsSustentoField As comprobanteRetencionDocsSustento

    Private infoAdicionalField() As comprobanteRetencionCampoAdicional

    Private idField As String

    Private versionField As String

    '''<remarks/>
    Public Property infoTributaria() As comprobanteRetencionInfoTributaria
        Get
            Return Me.infoTributariaField
        End Get
        Set
            Me.infoTributariaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property infoCompRetencion() As comprobanteRetencionInfoCompRetencion
        Get
            Return Me.infoCompRetencionField
        End Get
        Set
            Me.infoCompRetencionField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("impuesto", IsNullable:=False)>
    Public Property impuestos() As comprobanteRetencionImpuesto()
        Get
            Return Me.impuestosField
        End Get
        Set
            Me.impuestosField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property docsSustento() As comprobanteRetencionDocsSustento
        Get
            Return Me.docsSustentoField
        End Get
        Set
            Me.docsSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("campoAdicional", IsNullable:=False)>
    Public Property infoAdicional() As comprobanteRetencionCampoAdicional()
        Get
            Return Me.infoAdicionalField
        End Get
        Set
            Me.infoAdicionalField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute()>
    Public Property id() As String
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute()>
    Public Property version() As String
        Get
            Return Me.versionField
        End Get
        Set
            Me.versionField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionInfoTributaria

    Private ambienteField As String

    Private tipoEmisionField As String

    Private razonSocialField As String

    Private nombreComercialField As String

    Private rucField As String

    Private claveAccesoField As String

    Private codDocField As String

    Private estabField As String

    Private ptoEmiField As String

    Private secuencialField As String

    Private dirMatrizField As String

    Private agenteRetencionField As String

    Private contribuyenteRimpeField As String

    '''<remarks/>
    Public Property ambiente() As String
        Get
            Return Me.ambienteField
        End Get
        Set
            Me.ambienteField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoEmision() As String
        Get
            Return Me.tipoEmisionField
        End Get
        Set
            Me.tipoEmisionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property razonSocial() As String
        Get
            Return Me.razonSocialField
        End Get
        Set
            Me.razonSocialField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property nombreComercial() As String
        Get
            Return Me.nombreComercialField
        End Get
        Set
            Me.nombreComercialField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property ruc() As String
        Get
            Return Me.rucField
        End Get
        Set
            Me.rucField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property claveAcceso() As String
        Get
            Return Me.claveAccesoField
        End Get
        Set
            Me.claveAccesoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDoc() As String
        Get
            Return Me.codDocField
        End Get
        Set
            Me.codDocField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property estab() As String
        Get
            Return Me.estabField
        End Get
        Set
            Me.estabField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property ptoEmi() As String
        Get
            Return Me.ptoEmiField
        End Get
        Set
            Me.ptoEmiField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property secuencial() As String
        Get
            Return Me.secuencialField
        End Get
        Set
            Me.secuencialField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property dirMatriz() As String
        Get
            Return Me.dirMatrizField
        End Get
        Set
            Me.dirMatrizField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property agenteRetencion() As String
        Get
            Return Me.agenteRetencionField
        End Get
        Set
            Me.agenteRetencionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property contribuyenteRimpe() As String
        Get
            Return Me.contribuyenteRimpeField
        End Get
        Set
            Me.contribuyenteRimpeField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionInfoCompRetencion

    Private fechaEmisionField As String

    Private dirEstablecimientoField As String

    Private contribuyenteEspecialField As String

    Private obligadoContabilidadField As String

    Private tipoIdentificacionSujetoRetenidoField As String

    Private razonSocialSujetoRetenidoField As String

    Private identificacionSujetoRetenidoField As String

    Private periodoFiscalField As String

    Private tipoSujetoRetenidoField As String

    Private parteRelField As String


    '''<remarks/>
    Public Property fechaEmision() As String
        Get
            Return Me.fechaEmisionField
        End Get
        Set
            Me.fechaEmisionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property dirEstablecimiento() As String
        Get
            Return Me.dirEstablecimientoField
        End Get
        Set
            Me.dirEstablecimientoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property contribuyenteEspecial() As String
        Get
            Return Me.contribuyenteEspecialField
        End Get
        Set
            Me.contribuyenteEspecialField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property obligadoContabilidad() As String
        Get
            Return Me.obligadoContabilidadField
        End Get
        Set
            Me.obligadoContabilidadField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoIdentificacionSujetoRetenido() As String
        Get
            Return Me.tipoIdentificacionSujetoRetenidoField
        End Get
        Set
            Me.tipoIdentificacionSujetoRetenidoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property razonSocialSujetoRetenido() As String
        Get
            Return Me.razonSocialSujetoRetenidoField
        End Get
        Set
            Me.razonSocialSujetoRetenidoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property identificacionSujetoRetenido() As String
        Get
            Return Me.identificacionSujetoRetenidoField
        End Get
        Set
            Me.identificacionSujetoRetenidoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property periodoFiscal() As String
        Get
            Return Me.periodoFiscalField
        End Get
        Set
            Me.periodoFiscalField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoSujetoRetenido() As String
        Get
            Return Me.tipoSujetoRetenidoField
        End Get
        Set
            Me.tipoSujetoRetenidoField = Value
        End Set
    End Property


    '''<remarks/>
    Public Property parteRel() As String
        Get
            Return Me.parteRelField
        End Get
        Set
            Me.parteRelField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionImpuesto

    Private codigoField As Integer

    Private codigoRetencionField As String

    Private baseImponibleField As Decimal

    Private porcentajeRetenerField As Decimal

    Private valorRetenidoField As Decimal

    Private codDocSustentoField As String

    Private numDocSustentoField As String

    Private fechaEmisionDocSustentoField As String

    '''<remarks/>
    Public Property codigo() As Integer
        Get
            Return Me.codigoField
        End Get
        Set
            Me.codigoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codigoRetencion() As String
        Get
            Return Me.codigoRetencionField
        End Get
        Set
            Me.codigoRetencionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property baseImponible() As Decimal
        Get
            Return Me.baseImponibleField
        End Get
        Set
            Me.baseImponibleField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property porcentajeRetener() As Decimal
        Get
            Return Me.porcentajeRetenerField
        End Get
        Set
            Me.porcentajeRetenerField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property valorRetenido() As Decimal
        Get
            Return Me.valorRetenidoField
        End Get
        Set
            Me.valorRetenidoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDocSustento() As String
        Get
            Return Me.codDocSustentoField
        End Get
        Set
            Me.codDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property numDocSustento() As String
        Get
            Return Me.numDocSustentoField
        End Get
        Set
            Me.numDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaEmisionDocSustento() As String
        Get
            Return Me.fechaEmisionDocSustentoField
        End Get
        Set
            Me.fechaEmisionDocSustentoField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustento

    Private docSustentoField As comprobanteRetencionDocsSustentoDocSustento

    '''<remarks/>
    Public Property docSustento() As comprobanteRetencionDocsSustentoDocSustento
        Get
            Return Me.docSustentoField
        End Get
        Set
            Me.docSustentoField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustento

    Private codSustentoField As String

    Private codDocSustentoField As String

    Private numDocSustentoField As String

    Private fechaEmisionDocSustentoField As String

    Private fechaRegistroContableField As String

    Private numAutDocSustentoField As String

    Private pagoLocExtField As String

    Private tipoRegiField As String

    Private paisEfecPagoField As String

    Private aplicConvDobTribField As String

    Private pagExtSujRetNorLegField As String

    Private pagoRegFisField As String

    Private totalSinImpuestosField As Decimal

    Private importeTotalField As Decimal

    Private totalComprobantesReembolsoField As Decimal?

    Private totalBaseImponibleReembolsoField As Decimal?

    Private totalImpuestoReembolsoField As Decimal?

    Private impuestosDocSustentoField() As comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento

    Private retencionesField() As comprobanteRetencionDocsSustentoDocSustentoRetencion

    Private reembolsosField() As comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalle

    Private pagosField As comprobanteRetencionDocsSustentoDocSustentoPagos

    '''<remarks/>
    Public Property codSustento() As String
        Get
            Return Me.codSustentoField
        End Get
        Set
            Me.codSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDocSustento() As String
        Get
            Return Me.codDocSustentoField
        End Get
        Set
            Me.codDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property numDocSustento() As String
        Get
            Return Me.numDocSustentoField
        End Get
        Set
            Me.numDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaEmisionDocSustento() As String
        Get
            Return Me.fechaEmisionDocSustentoField
        End Get
        Set
            Me.fechaEmisionDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaRegistroContable() As String
        Get
            Return Me.fechaRegistroContableField
        End Get
        Set
            Me.fechaRegistroContableField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property numAutDocSustento() As String
        Get
            Return Me.numAutDocSustentoField
        End Get
        Set
            Me.numAutDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property pagoLocExt() As String
        Get
            Return Me.pagoLocExtField
        End Get
        Set
            Me.pagoLocExtField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoRegi() As String
        Get
            Return Me.tipoRegiField
        End Get
        Set
            Me.tipoRegiField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property paisEfecPago() As String
        Get
            Return Me.paisEfecPagoField
        End Get
        Set
            Me.paisEfecPagoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property aplicConvDobTrib() As String
        Get
            Return Me.aplicConvDobTribField
        End Get
        Set
            Me.aplicConvDobTribField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property pagExtSujRetNorLeg() As String
        Get
            Return Me.pagExtSujRetNorLegField
        End Get
        Set
            Me.pagExtSujRetNorLegField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property pagoRegFis() As String
        Get
            Return Me.pagoRegFisField
        End Get
        Set
            Me.pagoRegFisField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property totalSinImpuestos() As Decimal
        Get
            Return Me.totalSinImpuestosField
        End Get
        Set
            Me.totalSinImpuestosField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property importeTotal() As Decimal
        Get
            Return Me.importeTotalField
        End Get
        Set
            Me.importeTotalField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property totalComprobantesReembolso() As Decimal?
        Get
            Return Me.totalComprobantesReembolsoField
        End Get
        Set
            Me.totalComprobantesReembolsoField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property totalComprobantesReembolsoSpecified() As Boolean
        Get
            Return Me.totalComprobantesReembolso IsNot Nothing
        End Get
    End Property

    '''<remarks/>
    Public Property totalBaseImponibleReembolso() As Decimal?
        Get
            Return Me.totalBaseImponibleReembolsoField
        End Get
        Set
            Me.totalBaseImponibleReembolsoField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property totalBaseImponibleReembolsoSpecified() As Boolean
        Get
            Return Me.totalBaseImponibleReembolso IsNot Nothing
        End Get
    End Property


    '''<remarks/>
    Public Property totalImpuestoReembolso() As Decimal?
        Get
            Return Me.totalImpuestoReembolsoField
        End Get
        Set
            Me.totalImpuestoReembolsoField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property totalImpuestoReembolsoSpecified() As Boolean
        Get
            Return Me.totalImpuestoReembolso IsNot Nothing
        End Get
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("impuestoDocSustento", IsNullable:=False)>
    Public Property impuestosDocSustento() As comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento()
        Get
            Return Me.impuestosDocSustentoField
        End Get
        Set
            Me.impuestosDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("retencion", IsNullable:=False)>
    Public Property retenciones() As comprobanteRetencionDocsSustentoDocSustentoRetencion()
        Get
            Return Me.retencionesField
        End Get
        Set
            Me.retencionesField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("reembolsoDetalle", IsNullable:=False)>
    Public Property reembolsos() As comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalle()
        Get
            Return Me.reembolsosField
        End Get
        Set
            Me.reembolsosField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property pagos() As comprobanteRetencionDocsSustentoDocSustentoPagos
        Get
            Return Me.pagosField
        End Get
        Set
            Me.pagosField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento

    Private codImpuestoDocSustentoField As Integer

    Private codigoPorcentajeField As String

    Private baseImponibleField As Decimal

    Private tarifaField As Decimal

    Private valorImpuestoField As Decimal

    '''<remarks/>
    Public Property codImpuestoDocSustento() As Integer
        Get
            Return Me.codImpuestoDocSustentoField
        End Get
        Set
            Me.codImpuestoDocSustentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codigoPorcentaje() As String
        Get
            Return Me.codigoPorcentajeField
        End Get
        Set
            Me.codigoPorcentajeField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property baseImponible() As Decimal
        Get
            Return Me.baseImponibleField
        End Get
        Set
            Me.baseImponibleField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tarifa() As Decimal
        Get
            Return Me.tarifaField
        End Get
        Set
            Me.tarifaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property valorImpuesto() As Decimal
        Get
            Return Me.valorImpuestoField
        End Get
        Set
            Me.valorImpuestoField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustentoRetencion

    Private codigoField As Integer

    Private codigoRetencionField As String

    Private baseImponibleField As Decimal

    Private porcentajeRetenerField As Decimal

    Private valorRetenidoField As Decimal

    '''<remarks/>
    Public Property codigo() As Integer
        Get
            Return Me.codigoField
        End Get
        Set
            Me.codigoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codigoRetencion() As String
        Get
            Return Me.codigoRetencionField
        End Get
        Set
            Me.codigoRetencionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property baseImponible() As Decimal
        Get
            Return Me.baseImponibleField
        End Get
        Set
            Me.baseImponibleField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property porcentajeRetener() As Decimal
        Get
            Return Me.porcentajeRetenerField
        End Get
        Set
            Me.porcentajeRetenerField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property valorRetenido() As Decimal
        Get
            Return Me.valorRetenidoField
        End Get
        Set
            Me.valorRetenidoField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalle

    Private tipoIdentificacionProveedorReembolsoField As Byte

    Private identificacionProveedorReembolsoField As ULong

    Private codPaisPagoProveedorReembolsoField As UShort

    Private tipoProveedorReembolsoField As Byte

    Private codDocReembolsoField As Byte

    Private estabDocReembolsoField As Byte

    Private ptoEmiDocReembolsoField As Byte

    Private secuencialDocReembolsoField As UShort

    Private fechaEmisionDocReembolsoField As String

    Private numeroAutorizacionDocReembField As UInteger

    Private detalleImpuestosField() As comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto

    '''<remarks/>
    Public Property tipoIdentificacionProveedorReembolso() As Byte
        Get
            Return Me.tipoIdentificacionProveedorReembolsoField
        End Get
        Set
            Me.tipoIdentificacionProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property identificacionProveedorReembolso() As ULong
        Get
            Return Me.identificacionProveedorReembolsoField
        End Get
        Set
            Me.identificacionProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codPaisPagoProveedorReembolso() As UShort
        Get
            Return Me.codPaisPagoProveedorReembolsoField
        End Get
        Set
            Me.codPaisPagoProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoProveedorReembolso() As Byte
        Get
            Return Me.tipoProveedorReembolsoField
        End Get
        Set
            Me.tipoProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDocReembolso() As Byte
        Get
            Return Me.codDocReembolsoField
        End Get
        Set
            Me.codDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property estabDocReembolso() As Byte
        Get
            Return Me.estabDocReembolsoField
        End Get
        Set
            Me.estabDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property ptoEmiDocReembolso() As Byte
        Get
            Return Me.ptoEmiDocReembolsoField
        End Get
        Set
            Me.ptoEmiDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property secuencialDocReembolso() As UShort
        Get
            Return Me.secuencialDocReembolsoField
        End Get
        Set
            Me.secuencialDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaEmisionDocReembolso() As String
        Get
            Return Me.fechaEmisionDocReembolsoField
        End Get
        Set
            Me.fechaEmisionDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property numeroAutorizacionDocReemb() As UInteger
        Get
            Return Me.numeroAutorizacionDocReembField
        End Get
        Set
            Me.numeroAutorizacionDocReembField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("detalleImpuesto", IsNullable:=False)>
    Public Property detalleImpuestos() As comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto()
        Get
            Return Me.detalleImpuestosField
        End Get
        Set
            Me.detalleImpuestosField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto

    Private codigoField As Byte

    Private codigoPorcentajeField As Byte

    Private tarifaField As Byte

    Private baseImponibleReembolsoField As Decimal

    Private impuestoReembolsoField As Decimal

    '''<remarks/>
    Public Property codigo() As Byte
        Get
            Return Me.codigoField
        End Get
        Set
            Me.codigoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codigoPorcentaje() As Byte
        Get
            Return Me.codigoPorcentajeField
        End Get
        Set
            Me.codigoPorcentajeField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tarifa() As Byte
        Get
            Return Me.tarifaField
        End Get
        Set
            Me.tarifaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property baseImponibleReembolso() As Decimal
        Get
            Return Me.baseImponibleReembolsoField
        End Get
        Set
            Me.baseImponibleReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property impuestoReembolso() As Decimal
        Get
            Return Me.impuestoReembolsoField
        End Get
        Set
            Me.impuestoReembolsoField = Value
        End Set
    End Property
End Class


'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustentoPagos

    Private pagoField As comprobanteRetencionDocsSustentoDocSustentoPagosPago

    '''<remarks/>
    Public Property pago() As comprobanteRetencionDocsSustentoDocSustentoPagosPago
        Get
            Return Me.pagoField
        End Get
        Set
            Me.pagoField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionDocsSustentoDocSustentoPagosPago

    Private formaPagoField As String

    Private totalField As Decimal

    '''<remarks/>
    Public Property formaPago() As String
        Get
            Return Me.formaPagoField
        End Get
        Set
            Me.formaPagoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property total() As Decimal
        Get
            Return Me.totalField
        End Get
        Set
            Me.totalField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class comprobanteRetencionCampoAdicional

    Private nombreField As String

    Private valueField As String

    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute()>
    Public Property nombre() As String
        Get
            Return Me.nombreField
        End Get
        Set
            Me.nombreField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlTextAttribute()>
    Public Property Value() As String
        Get
            Return Me.valueField
        End Get
        Set
            Me.valueField = Value
        End Set
    End Property
End Class


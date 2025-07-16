
' NOTA: El código generado puede requerir, como mínimo, .NET Framework 4.5 o .NET Core/Standard 2.0.
'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False)>
Partial Public Class factura


    Private infoTributariaField As facturaInfoTributaria

    Private infoFacturaField As facturaInfoFactura

    Private detallesField() As facturaDetalle

    Private reembolsosField() As reembolsoDetalle

    Private infoAdicionalField() As facturaCampoAdicional

    Private idField As String

    Private versionField As String

    '''<remarks/>
    Public Property infoTributaria() As facturaInfoTributaria
        Get
            Return Me.infoTributariaField
        End Get
        Set
            Me.infoTributariaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property infoFactura() As facturaInfoFactura
        Get
            Return Me.infoFacturaField
        End Get
        Set
            Me.infoFacturaField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("detalle", IsNullable:=False)>
    Public Property detalles() As facturaDetalle()
        Get
            Return Me.detallesField
        End Get
        Set
            Me.detallesField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("reembolsoDetalle", IsNullable:=False)>
    Public Property reembolsos() As reembolsoDetalle()
        Get
            Return Me.reembolsosField
        End Get
        Set
            Me.reembolsosField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("campoAdicional", IsNullable:=False)>
    Public Property infoAdicional() As facturaCampoAdicional()
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
Partial Public Class facturaInfoTributaria

    Private ambienteField As String

    Private tipoEmisionField As String

    Private razonSocialField As String

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
    Public Property ruc() As String
        Get
            Return Me.rucField
        End Get
        Set
            Me.rucField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(DataType:="integer")>
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

    Public Property agenteRetencion() As String
        Get
            Return Me.agenteRetencionField
        End Get
        Set
            Me.agenteRetencionField = Value
        End Set
    End Property

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
Partial Public Class facturaInfoFactura

    Private fechaEmisionField As String

    Private dirEstablecimientoField As String

    Private contribuyenteEspecialField As String

    Private obligadoContabilidadField As String

    Private comercioExteriorField As String

    Private incoTermFacturaField As String

    Private lugarIncoTermField As String

    Private paisOrigenField As String

    Private puertoEmbarqueField As String

    Private puertoDestinoField As String

    Private paisDestinoField As String

    Private paisAdquisicionField As String

    Private tipoIdentificacionCompradorField As String

    Private guiaRemisionField As String

    Private razonSocialCompradorField As String

    Private identificacionCompradorField As String

    Private direccionCompradorField As String

    Private incoTermTotalSinImpuestosField As String

    Private totalSinImpuestosField As Decimal

    Private totalDescuentoField As Decimal

    Private codDocReembolsoField As String

    Private totalComprobantesReembolsoField As Decimal?

    Private totalBaseImponibleReembolsoField As Decimal?

    Private totalImpuestoReembolsoField As Decimal?

    Private totalConImpuestosField() As facturaInfoFacturaTotalImpuesto

    Private propinaField As Decimal

    Private fleteInternacionalField As Decimal?

    Private seguroInternacionalField As Decimal?

    Private GastosAduanerosField As Decimal?

    Private GastosTransporteOtrosField As Decimal?

    Private importeTotalField As Decimal

    Private monedaField As String

    Private pagosField As facturaInfoFacturaPagos

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
    Public Property tipoIdentificacionComprador() As String
        Get
            Return Me.tipoIdentificacionCompradorField
        End Get
        Set
            Me.tipoIdentificacionCompradorField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property guiaRemision() As String
        Get
            Return Me.guiaRemisionField
        End Get
        Set
            Me.guiaRemisionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property razonSocialComprador() As String
        Get
            Return Me.razonSocialCompradorField
        End Get
        Set
            Me.razonSocialCompradorField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property identificacionComprador() As String
        Get
            Return Me.identificacionCompradorField
        End Get
        Set
            Me.identificacionCompradorField = Value
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
    Public Property totalDescuento() As Decimal
        Get
            Return Me.totalDescuentoField
        End Get
        Set
            Me.totalDescuentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDocReembolso() As String
        Get
            Return Me.codDocReembolsoField
        End Get
        Set
            Me.codDocReembolsoField = Value
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
    <System.Xml.Serialization.XmlArrayItemAttribute("totalImpuesto", IsNullable:=False)>
    Public Property totalConImpuestos() As facturaInfoFacturaTotalImpuesto()
        Get
            Return Me.totalConImpuestosField
        End Get
        Set
            Me.totalConImpuestosField = Value
        End Set
    End Property


    Public Property comercioExterior() As String
        Get
            Return Me.comercioExteriorField
        End Get
        Set
            Me.comercioExteriorField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property incoTermFactura() As String
        Get
            Return Me.incoTermFacturaField
        End Get
        Set
            Me.incoTermFacturaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property lugarIncoTerm() As String
        Get
            Return Me.lugarIncoTermField
        End Get
        Set
            Me.lugarIncoTermField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property paisOrigen() As String
        Get
            Return Me.paisOrigenField
        End Get
        Set
            Me.paisOrigenField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property puertoEmbarque() As String
        Get
            Return Me.puertoEmbarqueField
        End Get
        Set
            Me.puertoEmbarqueField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property puertoDestino() As String
        Get
            Return Me.puertoDestinoField
        End Get
        Set
            Me.puertoDestinoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property paisDestino() As String
        Get
            Return Me.paisDestinoField
        End Get
        Set
            Me.paisDestinoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property paisAdquisicion() As String
        Get
            Return Me.paisAdquisicionField
        End Get
        Set
            Me.paisAdquisicionField = Value
        End Set
    End Property
    Public Property direccionComprador() As String
        Get
            Return Me.direccionCompradorField
        End Get
        Set
            Me.direccionCompradorField = Value
        End Set
    End Property


    Public Property incoTermTotalSinImpuestos() As String
        Get
            Return Me.incoTermTotalSinImpuestosField
        End Get
        Set
            Me.incoTermTotalSinImpuestosField = Value
        End Set
    End Property


    Public Property fleteInternacional() As Decimal?
        Get
            Return Me.fleteInternacionalField
        End Get
        Set
            Me.fleteInternacionalField = Value
        End Set
    End Property

    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property fleteInternacionalSpecified() As Boolean
        Get
            Return Me.fleteInternacional IsNot Nothing
        End Get
    End Property

    '''<remarks/>

    Public Property seguroInternacional() As Decimal?
        Get
            Return Me.seguroInternacionalField
        End Get
        Set
            Me.seguroInternacionalField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property seguroInternacionalSpecified() As Boolean
        Get
            Return Me.seguroInternacional IsNot Nothing
        End Get
    End Property



    Public Property GastosAduaneros() As Decimal?
        Get
            Return Me.GastosAduanerosField
        End Get
        Set
            Me.GastosAduanerosField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property GastosAduanerosSpecified() As Boolean
        Get
            Return Me.GastosAduaneros IsNot Nothing
        End Get
    End Property
    '''<remarks/>

    Public Property GastosTransporteOtros() As Decimal?
        Get
            Return Me.GastosTransporteOtrosField
        End Get
        Set
            Me.GastosTransporteOtrosField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property GastosTransporteOtrosSpecified() As Boolean
        Get
            Return Me.GastosTransporteOtros IsNot Nothing
        End Get
    End Property

    '''<remarks/>
    Public Property propina() As Decimal
        Get
            Return Me.propinaField
        End Get
        Set
            Me.propinaField = Value
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
    Public Property moneda() As String
        Get
            Return Me.monedaField
        End Get
        Set
            Me.monedaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property pagos() As facturaInfoFacturaPagos
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
Partial Public Class facturaInfoFacturaTotalImpuesto

    Private codigoField As Byte

    Private codigoPorcentajeField As String

    Private descuentoAdicionalField As Decimal?

    Private baseImponibleField As Decimal

    Private tarifaField As Decimal

    Private valorField As Decimal

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
    Public Property codigoPorcentaje() As String
        Get
            Return Me.codigoPorcentajeField
        End Get
        Set
            Me.codigoPorcentajeField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property descuentoAdicional() As Decimal?
        Get
            Return Me.descuentoAdicionalField
        End Get
        Set
            Me.descuentoAdicionalField = Value
        End Set
    End Property
    <System.Xml.Serialization.XmlIgnore>
    Public ReadOnly Property descuentoAdicionalSpecified() As Boolean
        Get
            Return Me.descuentoAdicional IsNot Nothing
        End Get
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
    Public Property valor() As Decimal
        Get
            Return Me.valorField
        End Get
        Set
            Me.valorField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class facturaInfoFacturaPagos

    Private pagoField As facturaInfoFacturaPagosPago

    '''<remarks/>
    Public Property pago() As facturaInfoFacturaPagosPago
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
Partial Public Class facturaInfoFacturaPagosPago

    Private formaPagoField As String

    Private totalField As Decimal

    Private plazoField As Integer

    Private unidadTiempoField As String

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

    '''<remarks/>
    Public Property plazo() As Integer
        Get
            Return Me.plazoField
        End Get
        Set
            Me.plazoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property unidadTiempo() As String
        Get
            Return Me.unidadTiempoField
        End Get
        Set
            Me.unidadTiempoField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class facturaDetalle

    Private codigoPrincipalField As String

    Private codigoAuxiliarField As String

    Private descripcionField As String

    Private cantidadField As Decimal

    Private precioUnitarioField As Decimal

    Private descuentoField As Decimal

    Private precioTotalSinImpuestoField As Decimal

    Private detallesAdicionalesField() As facturadetAdicional

    Private impuestosField() As facturaDetalleImpuesto

    '''<remarks/>
    Public Property codigoPrincipal() As String
        Get
            Return Me.codigoPrincipalField
        End Get
        Set
            Me.codigoPrincipalField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codigoAuxiliar() As String
        Get
            Return Me.codigoAuxiliarField
        End Get
        Set
            Me.codigoAuxiliarField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property descripcion() As String
        Get
            Return Me.descripcionField
        End Get
        Set
            Me.descripcionField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property cantidad() As Decimal
        Get
            Return Me.cantidadField
        End Get
        Set
            Me.cantidadField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property precioUnitario() As Decimal
        Get
            Return Me.precioUnitarioField
        End Get
        Set
            Me.precioUnitarioField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property descuento() As Decimal
        Get
            Return Me.descuentoField
        End Get
        Set
            Me.descuentoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property precioTotalSinImpuesto() As Decimal
        Get
            Return Me.precioTotalSinImpuestoField
        End Get
        Set
            Me.precioTotalSinImpuestoField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("detAdicional", IsNullable:=False)>
    Public Property detallesAdicionales As facturadetAdicional()
        Get
            Return Me.detallesAdicionalesField
        End Get
        Set
            Me.detallesAdicionalesField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("impuesto", IsNullable:=False)>
    Public Property impuestos() As facturaDetalleImpuesto()
        Get
            Return Me.impuestosField
        End Get
        Set
            Me.impuestosField = Value
        End Set
    End Property
End Class

'''<remarks/>
'<System.SerializableAttribute(),
'    System.ComponentModel.DesignerCategoryAttribute("code"),
'    System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
'Partial Public Class facturaDetalleDetallesAdicionales

'    Private detAdicionalField As facturadetAdicional()

'    '''<remarks/>
'    Public Property detAdicional As facturaDetalleDetallesAdicionalesDetAdicional()
'        Get
'            Return Me.detAdicionalField
'        End Get
'        Set
'            Me.detAdicionalField = Value
'        End Set
'    End Property
'End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class facturadetAdicional

    Private nombreField As String

    Private valorField As String

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
    <System.Xml.Serialization.XmlAttributeAttribute()>
    Public Property valor() As String
        Get
            Return Me.valorField
        End Get
        Set
            Me.valorField = Value
        End Set
    End Property
End Class

''''<remarks/>
'<System.SerializableAttribute(),
' System.ComponentModel.DesignerCategoryAttribute("code"),
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
'Partial Public Class facturaDetalleImpuestos

'    Private impuestoField As facturaDetalleImpuestosImpuesto()

'    '''<remarks/>
'    Public Property impuesto As facturaDetalleImpuestosImpuesto()
'        Get
'            Return Me.impuestoField
'        End Get
'        Set
'            Me.impuestoField = Value
'        End Set
'    End Property
'End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class facturaDetalleImpuesto

    Private codigoField As Integer

    Private codigoPorcentajeField As String

    Private tarifaField As Decimal

    Private baseImponibleField As Decimal

    Private valorField As Decimal

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
    Public Property codigoPorcentaje() As String
        Get
            Return Me.codigoPorcentajeField
        End Get
        Set
            Me.codigoPorcentajeField = Value
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
    Public Property baseImponible() As Decimal
        Get
            Return Me.baseImponibleField
        End Get
        Set
            Me.baseImponibleField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property valor() As Decimal
        Get
            Return Me.valorField
        End Get
        Set
            Me.valorField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class reembolsoDetalle

    Private tipoIdentificacionProveedorReembolsoField As String

    Private identificacionProveedorReembolsoField As String

    Private codPaisPagoProveedorReembolsoField As String

    Private tipoProveedorReembolsoField As String

    Private codDocReembolsoField As String

    Private estabDocReembolsoField As String

    Private ptoEmiDocReembolsoField As String

    Private secuencialDocReembolsoField As String

    Private fechaEmisionDocReembolsoField As Date

    Private numeroautorizacionDocReembField As String

    Private detalleImpuestosField() As facturaDetalleReembolsoImpuesto

    '''<remarks/>
    Public Property tipoIdentificacionProveedorReembolso() As String
        Get
            Return Me.tipoIdentificacionProveedorReembolsoField
        End Get
        Set
            Me.tipoIdentificacionProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property identificacionProveedorReembolso() As String
        Get
            Return Me.identificacionProveedorReembolsoField
        End Get
        Set
            Me.identificacionProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codPaisPagoProveedorReembolso() As String
        Get
            Return Me.codPaisPagoProveedorReembolsoField
        End Get
        Set
            Me.codPaisPagoProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoProveedorReembolso() As String
        Get
            Return Me.tipoProveedorReembolsoField
        End Get
        Set
            Me.tipoProveedorReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDocReembolso() As String
        Get
            Return Me.codDocReembolsoField
        End Get
        Set
            Me.codDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property estabDocReembolso() As String
        Get
            Return Me.estabDocReembolsoField
        End Get
        Set
            Me.estabDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property ptoEmiDocReembolso() As String
        Get
            Return Me.ptoEmiDocReembolsoField
        End Get
        Set
            Me.ptoEmiDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property secuencialDocReembolso() As String
        Get
            Return Me.secuencialDocReembolsoField
        End Get
        Set
            Me.secuencialDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaEmisionDocReembolso() As Date
        Get
            Return Me.fechaEmisionDocReembolsoField
        End Get
        Set
            Me.fechaEmisionDocReembolsoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property numeroautorizacionDocReemb() As String
        Get
            Return Me.numeroautorizacionDocReembField
        End Get
        Set
            Me.numeroautorizacionDocReembField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("detalleImpuesto", IsNullable:=False)>
    Public Property detalleImpuestos() As facturaDetalleReembolsoImpuesto()
        Get
            Return Me.detalleImpuestosField
        End Get
        Set
            Me.detalleImpuestosField = Value
        End Set
    End Property
End Class


<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class facturaDetalleReembolsoImpuesto

    Private codigoField As Integer

    Private codigoPorcentajeField As Integer

    Private tarifaField As Decimal

    Private baseImponibleReembolsoField As Decimal

    Private impuestoReembolsoField As Decimal

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
    Public Property codigoPorcentaje() As Integer
        Get
            Return Me.codigoPorcentajeField
        End Get
        Set
            Me.codigoPorcentajeField = Value
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
Partial Public Class facturaCampoAdicional

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



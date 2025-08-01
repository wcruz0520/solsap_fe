
' NOTA: El código generado puede requerir, como mínimo, .NET Framework 4.5 o .NET Core/Standard 2.0.
'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False)>
Partial Public Class notaCredito


    Private infoTributariaField As notaCreditoInfoTributaria

    Private infoNotaCreditoField As notaCreditoInfonotaCredito

    Private detallesField() As notaCreditoDetalle

    Private infoAdicionalField() As notaCreditoCampoAdicional

    Private idField As String

    Private versionField As String

    '''<remarks/>
    Public Property infoTributaria() As notaCreditoInfoTributaria
        Get
            Return Me.infoTributariaField
        End Get
        Set
            Me.infoTributariaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property infoNotaCredito() As notaCreditoInfonotaCredito
        Get
            Return Me.infoNotaCreditoField
        End Get
        Set
            Me.infoNotaCreditoField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("detalle", IsNullable:=False)>
    Public Property detalles() As notaCreditoDetalle()
        Get
            Return Me.detallesField
        End Get
        Set
            Me.detallesField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("campoAdicional", IsNullable:=False)>
    Public Property infoAdicional() As notaCreditoCampoAdicional()
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
Partial Public Class notaCreditoInfoTributaria

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
Partial Public Class notaCreditoInfonotaCredito

    Private fechaEmisionField As String

    Private dirEstablecimientoField As String

    Private tipoIdentificacionCompradorField As String

    Private razonSocialCompradorField As String

    Private identificacionCompradorField As String

    Private contribuyenteEspecialField As String

    Private obligadoContabilidadField As String

    Private riseField As String

    Private codDocModificadoField As String

    Private numDocModificadoField As String

    Private fechaEmisionDocSustentoField As String

    Private totalSinImpuestosField As Decimal

    Private valorModificacionField As Decimal

    Private monedaField As String

    Private totalConImpuestosField() As notaCreditoInfonotaCreditoTotalImpuesto

    Private motivoField As String

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
    Public Property tipoIdentificacionComprador() As String
        Get
            Return Me.tipoIdentificacionCompradorField
        End Get
        Set
            Me.tipoIdentificacionCompradorField = Value
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
    Public Property rise() As String
        Get
            Return Me.riseField
        End Get
        Set
            Me.riseField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codDocModificado() As String
        Get
            Return Me.codDocModificadoField
        End Get
        Set
            Me.codDocModificadoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property numDocModificado() As String
        Get
            Return Me.numDocModificadoField
        End Get
        Set
            Me.numDocModificadoField = Value
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
    Public Property totalSinImpuestos() As Decimal
        Get
            Return Me.totalSinImpuestosField
        End Get
        Set
            Me.totalSinImpuestosField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property valorModificacion() As Decimal
        Get
            Return Me.valorModificacionField
        End Get
        Set
            Me.valorModificacionField = Value
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
    <System.Xml.Serialization.XmlArrayItemAttribute("totalImpuesto", IsNullable:=False)>
    Public Property totalConImpuestos() As notaCreditoInfonotaCreditoTotalImpuesto()
        Get
            Return Me.totalConImpuestosField
        End Get
        Set
            Me.totalConImpuestosField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property motivo() As String
        Get
            Return Me.motivoField
        End Get
        Set
            Me.motivoField = Value
        End Set
    End Property


End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class notaCreditoInfonotaCreditoTotalImpuesto

    Private codigoField As Byte

    Private codigoPorcentajeField As String

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
Partial Public Class notaCreditoDetalle

    Private codigoInternoField As String

    Private codigoAdicionalField As String

    Private descripcionField As String

    Private cantidadField As Decimal

    Private precioUnitarioField As Decimal

    Private descuentoField As Decimal

    Private precioTotalSinImpuestoField As Decimal

    Private detallesAdicionalesField() As notaCreditodetAdicional

    Private impuestosField() As notaCreditoDetalleImpuesto

    '''<remarks/>
    Public Property codigoInterno() As String
        Get
            Return Me.codigoInternoField
        End Get
        Set
            Me.codigoInternoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codigoAdicional() As String
        Get
            Return Me.codigoAdicionalField
        End Get
        Set
            Me.codigoAdicionalField = Value
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
    Public Property detallesAdicionales As notaCreditodetAdicional()
        Get
            Return Me.detallesAdicionalesField
        End Get
        Set
            Me.detallesAdicionalesField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("impuesto", IsNullable:=False)>
    Public Property impuestos() As notaCreditoDetalleImpuesto()
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
'Partial Public Class notaCreditoDetalleDetallesAdicionales

'    Private detAdicionalField As notaCreditodetAdicional()

'    '''<remarks/>
'    Public Property detAdicional As notaCreditoDetalleDetallesAdicionalesDetAdicional()
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
Partial Public Class notaCreditodetAdicional

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
'Partial Public Class notaCreditoDetalleImpuestos

'    Private impuestoField As notaCreditoDetalleImpuestosImpuesto()

'    '''<remarks/>
'    Public Property impuesto As notaCreditoDetalleImpuestosImpuesto()
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
Partial Public Class notaCreditoDetalleImpuesto

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
Partial Public Class notaCreditoCampoAdicional

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



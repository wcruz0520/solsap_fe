
' NOTA: El código generado puede requerir, como mínimo, .NET Framework 4.5 o .NET Core/Standard 2.0.
'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False)>
Partial Public Class notaDebito


    Private infoTributariaField As notaDebitoInfoTributaria

    Private infoNotaDebitoField As notaDebitoInfonotaDebito

    Private motivosField() As notaDebitomotivo

    Private infoAdicionalField() As notaDebitoCampoAdicional

    Private idField As String

    Private versionField As String

    '''<remarks/>
    Public Property infoTributaria() As notaDebitoInfoTributaria
        Get
            Return Me.infoTributariaField
        End Get
        Set
            Me.infoTributariaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property infoNotaDebito() As notaDebitoInfonotaDebito
        Get
            Return Me.infoNotaDebitoField
        End Get
        Set
            Me.infoNotaDebitoField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("motivo", IsNullable:=False)>
    Public Property motivos() As notaDebitomotivo()
        Get
            Return Me.motivosField
        End Get
        Set
            Me.motivosField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("campoAdicional", IsNullable:=False)>
    Public Property infoAdicional() As notaDebitoCampoAdicional()
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
Partial Public Class notaDebitoInfoTributaria

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
Partial Public Class notaDebitoInfonotaDebito

    Private fechaEmisionField As String

    Private dirEstablecimientoField As String

    Private tipoIdentificacionCompradorField As String

    Private razonSocialCompradorField As String

    Private identificacionCompradorField As String

    Private contribuyenteEspecialField As String

    Private obligadoContabilidadField As String

    Private direccionCompradorField As String

    Private codDocModificadoField As String

    Private numDocModificadoField As String

    Private fechaEmisionDocSustentoField As String

    Private totalSinImpuestosField As Decimal

    Private impuestosField() As notaDebitoInfonotaDebitoTotalImpuesto

    Private valorTotalField As Decimal

    Private pagosField As notaDebitoInfonotaDebitoPagos

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
    Public Property direccionComprador() As String
        Get
            Return Me.direccionCompradorField
        End Get
        Set
            Me.direccionCompradorField = Value
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
    <System.Xml.Serialization.XmlArrayItemAttribute("impuesto", IsNullable:=False)>
    Public Property impuestos() As notaDebitoInfonotaDebitoTotalImpuesto()
        Get
            Return Me.impuestosField
        End Get
        Set
            Me.impuestosField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property valorTotal() As Decimal
        Get
            Return Me.valorTotalField
        End Get
        Set
            Me.valorTotalField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property pagos() As notaDebitoInfonotaDebitoPagos
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
Partial Public Class notaDebitoInfonotaDebitoTotalImpuesto

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
Partial Public Class notaDebitoInfonotaDebitoPagos

    Private pagoField As notaDebitoInfonotaDebitoPagosPago

    '''<remarks/>
    Public Property pago() As notaDebitoInfonotaDebitoPagosPago
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
Partial Public Class notaDebitoInfonotaDebitoPagosPago

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
Partial Public Class notaDebitomotivo

    Private razonField As String
    Private valorField As Decimal


    '''<remarks/>
    Public Property razon() As String
        Get
            Return Me.razonField
        End Get
        Set
            Me.razonField = Value
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
Partial Public Class notaDebitoCampoAdicional

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




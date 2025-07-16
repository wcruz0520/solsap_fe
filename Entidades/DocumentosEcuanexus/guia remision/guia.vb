'Public Class guia

'End Class


' NOTA: El código generado puede requerir, como mínimo, .NET Framework 4.5 o .NET Core/Standard 2.0.
'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False)>
Partial Public Class guiaRemision

    Private infoTributariaField As guiaRemisionInfoTributaria

    Private infoGuiaRemisionField As guiaRemisionInfoGuiaRemision

    Private destinatariosField() As guiaRemisionDestinatario

    Private infoAdicionalField() As guiaRemisionCampoAdicional

    Private idField As String

    Private versionField As String

    '''<remarks/>
    Public Property infoTributaria() As guiaRemisionInfoTributaria
        Get
            Return Me.infoTributariaField
        End Get
        Set
            Me.infoTributariaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property infoGuiaRemision() As guiaRemisionInfoGuiaRemision
        Get
            Return Me.infoGuiaRemisionField
        End Get
        Set
            Me.infoGuiaRemisionField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("destinatario", IsNullable:=False)>
    Public Property destinatarios() As guiaRemisionDestinatario()
        Get
            Return Me.destinatariosField
        End Get
        Set
            Me.destinatariosField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlArrayItemAttribute("campoAdicional", IsNullable:=False)>
    Public Property infoAdicional() As guiaRemisionCampoAdicional()
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
Partial Public Class guiaRemisionInfoTributaria

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

    '''<remarks/>
    Public Property agenteRetencion() As Byte
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
Partial Public Class guiaRemisionInfoGuiaRemision

    Private dirEstablecimientoField As String

    Private dirPartidaField As String

    Private razonSocialTransportistaField As String

    Private tipoIdentificacionTransportistaField As String

    Private rucTransportistaField As String

    Private obligadoContabilidadField As String

    Private contribuyenteEspecialField As String

    Private fechaIniTransporteField As String

    Private fechaFinTransporteField As String

    Private placaField As String

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
    Public Property dirPartida() As String
        Get
            Return Me.dirPartidaField
        End Get
        Set
            Me.dirPartidaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property razonSocialTransportista() As String
        Get
            Return Me.razonSocialTransportistaField
        End Get
        Set
            Me.razonSocialTransportistaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property tipoIdentificacionTransportista() As String
        Get
            Return Me.tipoIdentificacionTransportistaField
        End Get
        Set
            Me.tipoIdentificacionTransportistaField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property rucTransportista() As String
        Get
            Return Me.rucTransportistaField
        End Get
        Set
            Me.rucTransportistaField = Value
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
    Public Property contribuyenteEspecial() As String
        Get
            Return Me.contribuyenteEspecialField
        End Get
        Set
            Me.contribuyenteEspecialField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaIniTransporte() As String
        Get
            Return Me.fechaIniTransporteField
        End Get
        Set
            Me.fechaIniTransporteField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property fechaFinTransporte() As String
        Get
            Return Me.fechaFinTransporteField
        End Get
        Set
            Me.fechaFinTransporteField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property placa() As String
        Get
            Return Me.placaField
        End Get
        Set
            Me.placaField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class guiaRemisionDestinatario

    Private identificacionDestinatarioField As String

    Private razonSocialDestinatarioField As String

    Private dirDestinatarioField As String

    Private motivoTrasladoField As String

    Private codEstabDestinoField As String

    Private rutaField As String

    Private codDocSustentoField As String

    Private numDocSustentoField As String

    Private numAutDocSustentoField As String

    Private fechaEmisionDocSustentoField As String

    Private detallesField() As guiaRemisionDestinatarioDetalle

    '''<remarks/>
    Public Property identificacionDestinatario() As String
        Get
            Return Me.identificacionDestinatarioField
        End Get
        Set
            Me.identificacionDestinatarioField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property razonSocialDestinatario() As String
        Get
            Return Me.razonSocialDestinatarioField
        End Get
        Set
            Me.razonSocialDestinatarioField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property dirDestinatario() As String
        Get
            Return Me.dirDestinatarioField
        End Get
        Set
            Me.dirDestinatarioField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property motivoTraslado() As String
        Get
            Return Me.motivoTrasladoField
        End Get
        Set
            Me.motivoTrasladoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property codEstabDestino() As String
        Get
            Return Me.codEstabDestinoField
        End Get
        Set
            Me.codEstabDestinoField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property ruta() As String
        Get
            Return Me.rutaField
        End Get
        Set
            Me.rutaField = Value
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
    Public Property numAutDocSustento() As String
        Get
            Return Me.numAutDocSustentoField
        End Get
        Set
            Me.numAutDocSustentoField = Value
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
    <System.Xml.Serialization.XmlArrayItemAttribute("detalle", IsNullable:=False)>
    Public Property detalles() As guiaRemisionDestinatarioDetalle()
        Get
            Return Me.detallesField
        End Get
        Set
            Me.detallesField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class guiaRemisionDestinatarioDetalle

    Private codigoInternoField As String

    Private codigoAdicionalField As String

    Private descripcionField As String

    Private cantidadField As Decimal

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
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class guiaRemisionCampoAdicional

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



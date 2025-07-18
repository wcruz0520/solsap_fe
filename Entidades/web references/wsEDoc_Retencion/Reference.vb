﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Este código fue generado por una herramienta.
'     Versión de runtime:4.0.30319.42000
'
'     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
'     se vuelve a generar el código.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'Microsoft.VSDesigner generó automáticamente este código fuente, versión=4.0.30319.42000.
'
Namespace wsEDoc_Retencion
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="BasicHttpBinding_IWSEDOCNUBE_RETENCIONES", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class WSEDOCNUBE_RETENCIONES
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private EnviarRetencionSRIOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = Global.Entidades.My.MySettings.Default.Entidades_wsEDoc_Retencion_WSEDOC_RETENCIONES
            If (Me.IsLocalFileSystemWebService(Me.Url) = true) Then
                Me.UseDefaultCredentials = true
                Me.useDefaultCredentialsSetExplicitly = false
            Else
                Me.useDefaultCredentialsSetExplicitly = true
            End If
        End Sub
        
        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = true)  _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = false))  _
                            AndAlso (Me.IsLocalFileSystemWebService(value) = false)) Then
                    MyBase.UseDefaultCredentials = false
                End If
                MyBase.Url = value
            End Set
        End Property
        
        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set
                MyBase.UseDefaultCredentials = value
                Me.useDefaultCredentialsSetExplicitly = true
            End Set
        End Property
        
        '''<remarks/>
        Public Event EnviarRetencionSRICompleted As EnviarRetencionSRICompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/IWSEDOCNUBE_RETENCIONES/EnviarRetencionSRI", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function EnviarRetencionSRI(ByVal Credencial As String, ByVal Entorno As String, ByVal Retencion As ENTRetencion, ByRef mensaje As String) As RespuestaEDOC
            Dim results() As Object = Me.Invoke("EnviarRetencionSRI", New Object() {Credencial, Entorno, Retencion, mensaje})
            mensaje = CType(results(1),String)
            Return CType(results(0),RespuestaEDOC)
        End Function
        
        '''<remarks/>
        Public Overloads Sub EnviarRetencionSRIAsync(ByVal Credencial As String, ByVal Entorno As String, ByVal Retencion As ENTRetencion, ByVal mensaje As String)
            Me.EnviarRetencionSRIAsync(Credencial, Entorno, Retencion, mensaje, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub EnviarRetencionSRIAsync(ByVal Credencial As String, ByVal Entorno As String, ByVal Retencion As ENTRetencion, ByVal mensaje As String, ByVal userState As Object)
            If (Me.EnviarRetencionSRIOperationCompleted Is Nothing) Then
                Me.EnviarRetencionSRIOperationCompleted = AddressOf Me.OnEnviarRetencionSRIOperationCompleted
            End If
            Me.InvokeAsync("EnviarRetencionSRI", New Object() {Credencial, Entorno, Retencion, mensaje}, Me.EnviarRetencionSRIOperationCompleted, userState)
        End Sub
        
        Private Sub OnEnviarRetencionSRIOperationCompleted(ByVal arg As Object)
            If (Not (Me.EnviarRetencionSRICompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent EnviarRetencionSRICompleted(Me, New EnviarRetencionSRICompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
        
        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing)  _
                        OrElse (url Is String.Empty)) Then
                Return false
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024)  _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return true
            End If
            Return false
        End Function
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ENTRetencion
        
        Private campo1Field As String
        
        Private campo2Field As String
        
        Private campo3Field As String
        
        Private campo4Field As String
        
        Private campo5Field As String
        
        Private campo6Field As String
        
        Private campo7Field As String
        
        Private campo8Field As String
        
        Private campo9Field As String
        
        Private campo10Field As String
        
        Private baseImponibleField As Decimal
        
        Private codigoTransaccionERPField As String
        
        Private usuarioTransaccionERPField As String
        
        Private idRetencionField As Long
        
        Private autorizacionSRIField As String
        
        Private fechaAutorizacionField As System.Nullable(Of Date)
        
        Private ambienteField As Integer
        
        Private tipoEmisionField As Integer
        
        Private razonSocialField As String
        
        Private nombreComercialField As String
        
        Private rucField As String
        
        Private claveAccesoField As String
        
        Private codigoDocumentoField As String
        
        Private puntoEmisionField As String
        
        Private establecimientoField As String
        
        Private secuencialField As String
        
        Private direccionMatrizField As String
        
        Private fechaEmisionField As Date
        
        Private direccionEstablecimientoField As String
        
        Private contribuyenteEspecialField As String
        
        Private obligadoContabilidadField As String
        
        Private tipoIdentificacionSujetoRetenidoField As String
        
        Private identificacionSujetoRetenidoField As String
        
        Private razonSocialSujetoRetenidoField As String
        
        Private periodoFiscalField As String
        
        Private totalRetencionField As Decimal
        
        Private estadoField As Integer
        
        Private secuencialERPField As String
        
        Private emailResponsableField As String
        
        Private eNTDatoAdicionalRetencionField() As ENTDatoAdicionalRetencion
        
        Private eNTDetalleRetencionField() As ENTDetalleRetencion
        
        Private eNTDatosOpcionalesField As ENTDatosOpcionales
        
        Private datosFacturadorManualField As DatosFacturadorManual
        
        '''<remarks/>
        Public Property Campo1() As String
            Get
                Return Me.campo1Field
            End Get
            Set
                Me.campo1Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo2() As String
            Get
                Return Me.campo2Field
            End Get
            Set
                Me.campo2Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo3() As String
            Get
                Return Me.campo3Field
            End Get
            Set
                Me.campo3Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo4() As String
            Get
                Return Me.campo4Field
            End Get
            Set
                Me.campo4Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo5() As String
            Get
                Return Me.campo5Field
            End Get
            Set
                Me.campo5Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo6() As String
            Get
                Return Me.campo6Field
            End Get
            Set
                Me.campo6Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo7() As String
            Get
                Return Me.campo7Field
            End Get
            Set
                Me.campo7Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo8() As String
            Get
                Return Me.campo8Field
            End Get
            Set
                Me.campo8Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo9() As String
            Get
                Return Me.campo9Field
            End Get
            Set
                Me.campo9Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Campo10() As String
            Get
                Return Me.campo10Field
            End Get
            Set
                Me.campo10Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property BaseImponible() As Decimal
            Get
                Return Me.baseImponibleField
            End Get
            Set
                Me.baseImponibleField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property CodigoTransaccionERP() As String
            Get
                Return Me.codigoTransaccionERPField
            End Get
            Set
                Me.codigoTransaccionERPField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property UsuarioTransaccionERP() As String
            Get
                Return Me.usuarioTransaccionERPField
            End Get
            Set
                Me.usuarioTransaccionERPField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property IdRetencion() As Long
            Get
                Return Me.idRetencionField
            End Get
            Set
                Me.idRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property AutorizacionSRI() As String
            Get
                Return Me.autorizacionSRIField
            End Get
            Set
                Me.autorizacionSRIField = value
            End Set
        End Property
        
        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute(IsNullable:=true)>  _
        Public Property FechaAutorizacion() As System.Nullable(Of Date)
            Get
                Return Me.fechaAutorizacionField
            End Get
            Set
                Me.fechaAutorizacionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Ambiente() As Integer
            Get
                Return Me.ambienteField
            End Get
            Set
                Me.ambienteField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property TipoEmision() As Integer
            Get
                Return Me.tipoEmisionField
            End Get
            Set
                Me.tipoEmisionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property RazonSocial() As String
            Get
                Return Me.razonSocialField
            End Get
            Set
                Me.razonSocialField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property NombreComercial() As String
            Get
                Return Me.nombreComercialField
            End Get
            Set
                Me.nombreComercialField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Ruc() As String
            Get
                Return Me.rucField
            End Get
            Set
                Me.rucField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ClaveAcceso() As String
            Get
                Return Me.claveAccesoField
            End Get
            Set
                Me.claveAccesoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property CodigoDocumento() As String
            Get
                Return Me.codigoDocumentoField
            End Get
            Set
                Me.codigoDocumentoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property PuntoEmision() As String
            Get
                Return Me.puntoEmisionField
            End Get
            Set
                Me.puntoEmisionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Establecimiento() As String
            Get
                Return Me.establecimientoField
            End Get
            Set
                Me.establecimientoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Secuencial() As String
            Get
                Return Me.secuencialField
            End Get
            Set
                Me.secuencialField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property DireccionMatriz() As String
            Get
                Return Me.direccionMatrizField
            End Get
            Set
                Me.direccionMatrizField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property FechaEmision() As Date
            Get
                Return Me.fechaEmisionField
            End Get
            Set
                Me.fechaEmisionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property DireccionEstablecimiento() As String
            Get
                Return Me.direccionEstablecimientoField
            End Get
            Set
                Me.direccionEstablecimientoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ContribuyenteEspecial() As String
            Get
                Return Me.contribuyenteEspecialField
            End Get
            Set
                Me.contribuyenteEspecialField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ObligadoContabilidad() As String
            Get
                Return Me.obligadoContabilidadField
            End Get
            Set
                Me.obligadoContabilidadField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property TipoIdentificacionSujetoRetenido() As String
            Get
                Return Me.tipoIdentificacionSujetoRetenidoField
            End Get
            Set
                Me.tipoIdentificacionSujetoRetenidoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property IdentificacionSujetoRetenido() As String
            Get
                Return Me.identificacionSujetoRetenidoField
            End Get
            Set
                Me.identificacionSujetoRetenidoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property RazonSocialSujetoRetenido() As String
            Get
                Return Me.razonSocialSujetoRetenidoField
            End Get
            Set
                Me.razonSocialSujetoRetenidoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property PeriodoFiscal() As String
            Get
                Return Me.periodoFiscalField
            End Get
            Set
                Me.periodoFiscalField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property TotalRetencion() As Decimal
            Get
                Return Me.totalRetencionField
            End Get
            Set
                Me.totalRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Estado() As Integer
            Get
                Return Me.estadoField
            End Get
            Set
                Me.estadoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property SecuencialERP() As String
            Get
                Return Me.secuencialERPField
            End Get
            Set
                Me.secuencialERPField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property EmailResponsable() As String
            Get
                Return Me.emailResponsableField
            End Get
            Set
                Me.emailResponsableField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ENTDatoAdicionalRetencion() As ENTDatoAdicionalRetencion()
            Get
                Return Me.eNTDatoAdicionalRetencionField
            End Get
            Set
                Me.eNTDatoAdicionalRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ENTDetalleRetencion() As ENTDetalleRetencion()
            Get
                Return Me.eNTDetalleRetencionField
            End Get
            Set
                Me.eNTDetalleRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ENTDatosOpcionales() As ENTDatosOpcionales
            Get
                Return Me.eNTDatosOpcionalesField
            End Get
            Set
                Me.eNTDatosOpcionalesField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property DatosFacturadorManual() As DatosFacturadorManual
            Get
                Return Me.datosFacturadorManualField
            End Get
            Set
                Me.datosFacturadorManualField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ENTDatoAdicionalRetencion
        
        Private idDatoAdicionalRetencionField As Long
        
        Private nombreField As String
        
        Private descripcionField As String
        
        Private idRetencionField As Long
        
        '''<remarks/>
        Public Property IdDatoAdicionalRetencion() As Long
            Get
                Return Me.idDatoAdicionalRetencionField
            End Get
            Set
                Me.idDatoAdicionalRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Nombre() As String
            Get
                Return Me.nombreField
            End Get
            Set
                Me.nombreField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Descripcion() As String
            Get
                Return Me.descripcionField
            End Get
            Set
                Me.descripcionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property IdRetencion() As Long
            Get
                Return Me.idRetencionField
            End Get
            Set
                Me.idRetencionField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class EComprobante
        
        Private claveAccesoField As String
        
        Private mensajesField() As EMensaje
        
        '''<remarks/>
        Public Property claveAcceso() As String
            Get
                Return Me.claveAccesoField
            End Get
            Set
                Me.claveAccesoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property mensajes() As EMensaje()
            Get
                Return Me.mensajesField
            End Get
            Set
                Me.mensajesField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class EMensaje
        
        Private identificadorField As String
        
        Private mensaje1Field As String
        
        Private informacionAdicionalField As String
        
        Private tipoField As String
        
        '''<remarks/>
        Public Property identificador() As String
            Get
                Return Me.identificadorField
            End Get
            Set
                Me.identificadorField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property mensaje1() As String
            Get
                Return Me.mensaje1Field
            End Get
            Set
                Me.mensaje1Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property informacionAdicional() As String
            Get
                Return Me.informacionAdicionalField
            End Get
            Set
                Me.informacionAdicionalField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property tipo() As String
            Get
                Return Me.tipoField
            End Get
            Set
                Me.tipoField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class EAutorizacion
        
        Private estadoField As String
        
        Private numeroAutorizacionField As String
        
        Private claveAccesoField As String
        
        Private fechaAutorizacionField As Date
        
        Private fechaAutorizacionFieldSpecified As Boolean
        
        Private fechaAutorizacionSpecified1Field As Boolean
        
        Private ambienteField As String
        
        Private comprobanteField As String
        
        Private mensajesField() As EMensaje
        
        '''<remarks/>
        Public Property estado() As String
            Get
                Return Me.estadoField
            End Get
            Set
                Me.estadoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property numeroAutorizacion() As String
            Get
                Return Me.numeroAutorizacionField
            End Get
            Set
                Me.numeroAutorizacionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ClaveAcceso() As String
            Get
                Return Me.claveAccesoField
            End Get
            Set
                Me.claveAccesoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property fechaAutorizacion() As Date
            Get
                Return Me.fechaAutorizacionField
            End Get
            Set
                Me.fechaAutorizacionField = value
            End Set
        End Property
        
        '''<remarks/>
        <System.Xml.Serialization.XmlIgnoreAttribute()>  _
        Public Property fechaAutorizacionSpecified() As Boolean
            Get
                Return Me.fechaAutorizacionFieldSpecified
            End Get
            Set
                Me.fechaAutorizacionFieldSpecified = value
            End Set
        End Property
        
        '''<remarks/>
        <System.Xml.Serialization.XmlElementAttribute("fechaAutorizacionSpecified")>  _
        Public Property fechaAutorizacionSpecified1() As Boolean
            Get
                Return Me.fechaAutorizacionSpecified1Field
            End Get
            Set
                Me.fechaAutorizacionSpecified1Field = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ambiente() As String
            Get
                Return Me.ambienteField
            End Get
            Set
                Me.ambienteField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property comprobante() As String
            Get
                Return Me.comprobanteField
            End Get
            Set
                Me.comprobanteField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property mensajes() As EMensaje()
            Get
                Return Me.mensajesField
            End Get
            Set
                Me.mensajesField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class RespuestaEDOC
        
        Private claveAccesoField As String
        
        Private numeroComprobantesField As String
        
        Private estadoField As String
        
        Private autorizacionesField() As EAutorizacion
        
        Private comprobantesField() As EComprobante
        
        '''<remarks/>
        Public Property ClaveAcceso() As String
            Get
                Return Me.claveAccesoField
            End Get
            Set
                Me.claveAccesoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property NumeroComprobantes() As String
            Get
                Return Me.numeroComprobantesField
            End Get
            Set
                Me.numeroComprobantesField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Estado() As String
            Get
                Return Me.estadoField
            End Get
            Set
                Me.estadoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property autorizaciones() As EAutorizacion()
            Get
                Return Me.autorizacionesField
            End Get
            Set
                Me.autorizacionesField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Comprobantes() As EComprobante()
            Get
                Return Me.comprobantesField
            End Get
            Set
                Me.comprobantesField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class DatosFacturadorManual
        
        Private facturadorManualField As Boolean
        
        Private usaDirectorioField As Integer
        
        '''<remarks/>
        Public Property FacturadorManual() As Boolean
            Get
                Return Me.facturadorManualField
            End Get
            Set
                Me.facturadorManualField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property UsaDirectorio() As Integer
            Get
                Return Me.usaDirectorioField
            End Get
            Set
                Me.usaDirectorioField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ENTDatosOpcionales
        
        Private mailResponsableField As String
        
        Private usuarioCreadorField As String
        
        Private directorioDocumentoField As String
        
        Private nombreDocumentoField As String
        
        '''<remarks/>
        Public Property MailResponsable() As String
            Get
                Return Me.mailResponsableField
            End Get
            Set
                Me.mailResponsableField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property UsuarioCreador() As String
            Get
                Return Me.usuarioCreadorField
            End Get
            Set
                Me.usuarioCreadorField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property DirectorioDocumento() As String
            Get
                Return Me.directorioDocumentoField
            End Get
            Set
                Me.directorioDocumentoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property NombreDocumento() As String
            Get
                Return Me.nombreDocumentoField
            End Get
            Set
                Me.nombreDocumentoField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ENTDetalleRetencion
        
        Private idDetalleRetencionField As Long
        
        Private codigoField As Integer
        
        Private codigoRetencionField As String
        
        Private baseImponibleField As Decimal
        
        Private porcentajeRetenerField As Decimal
        
        Private valorRetenidoField As Decimal
        
        Private codDocRetenerField As String
        
        Private numDocRetenerField As String
        
        Private fechaEmisionDocRetenerField As Date
        
        Private idRetencionField As Long
        
        '''<remarks/>
        Public Property IdDetalleRetencion() As Long
            Get
                Return Me.idDetalleRetencionField
            End Get
            Set
                Me.idDetalleRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Codigo() As Integer
            Get
                Return Me.codigoField
            End Get
            Set
                Me.codigoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property CodigoRetencion() As String
            Get
                Return Me.codigoRetencionField
            End Get
            Set
                Me.codigoRetencionField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property BaseImponible() As Decimal
            Get
                Return Me.baseImponibleField
            End Get
            Set
                Me.baseImponibleField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property PorcentajeRetener() As Decimal
            Get
                Return Me.porcentajeRetenerField
            End Get
            Set
                Me.porcentajeRetenerField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property ValorRetenido() As Decimal
            Get
                Return Me.valorRetenidoField
            End Get
            Set
                Me.valorRetenidoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property CodDocRetener() As String
            Get
                Return Me.codDocRetenerField
            End Get
            Set
                Me.codDocRetenerField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property NumDocRetener() As String
            Get
                Return Me.numDocRetenerField
            End Get
            Set
                Me.numDocRetenerField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property FechaEmisionDocRetener() As Date
            Get
                Return Me.fechaEmisionDocRetenerField
            End Get
            Set
                Me.fechaEmisionDocRetenerField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property IdRetencion() As Long
            Get
                Return Me.idRetencionField
            End Get
            Set
                Me.idRetencionField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub EnviarRetencionSRICompletedEventHandler(ByVal sender As Object, ByVal e As EnviarRetencionSRICompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class EnviarRetencionSRICompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As RespuestaEDOC
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),RespuestaEDOC)
            End Get
        End Property
        
        '''<remarks/>
        Public ReadOnly Property mensaje() As String
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(1),String)
            End Get
        End Property
    End Class
End Namespace

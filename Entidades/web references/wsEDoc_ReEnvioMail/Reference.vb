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
Namespace wsEDoc_ReEnvioMail
    
    'CODEGEN: No se controló el elemento de extensión WSDL opcional 'PolicyReference' del espacio de nombres 'http://schemas.xmlsoap.org/ws/2004/09/policy'.
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="BasicHttpBinding_IWSEDOC_ENVIARMAIL", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class WSEDOC_ENVIARMAIL
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private EnviarCorreoDocumentoEmitidoOperationCompleted As System.Threading.SendOrPostCallback
        
        Private ReenvioMailEnLineaOperationCompleted As System.Threading.SendOrPostCallback
        
        Private ResendArchiveOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = Global.Entidades.My.MySettings.Default.Entidades_wsEDoc_ReEnvioMail_WSEDOC_ENVIARMAIL
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
        Public Event EnviarCorreoDocumentoEmitidoCompleted As EnviarCorreoDocumentoEmitidoCompletedEventHandler
        
        '''<remarks/>
        Public Event ReenvioMailEnLineaCompleted As ReenvioMailEnLineaCompletedEventHandler
        
        '''<remarks/>
        Public Event ResendArchiveCompleted As ResendArchiveCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/IWSEDOC_ENVIARMAIL/EnviarCorreoDocumentoEmitido", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Sub EnviarCorreoDocumentoEmitido(<System.Xml.Serialization.XmlElementAttribute(IsNullable:=true)> ByVal claveacceso As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=true)> ByVal correocliente As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=true)> ByRef mensaje As String, ByRef EnviarCorreoDocumentoEmitidoResult As Boolean, <System.Xml.Serialization.XmlIgnoreAttribute()> ByRef EnviarCorreoDocumentoEmitidoResultSpecified As Boolean)
            Dim results() As Object = Me.Invoke("EnviarCorreoDocumentoEmitido", New Object() {claveacceso, correocliente, mensaje})
            mensaje = CType(results(0),String)
            EnviarCorreoDocumentoEmitidoResult = CType(results(1),Boolean)
            EnviarCorreoDocumentoEmitidoResultSpecified = CType(results(2),Boolean)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub EnviarCorreoDocumentoEmitidoAsync(ByVal claveacceso As String, ByVal correocliente As String, ByVal mensaje As String)
            Me.EnviarCorreoDocumentoEmitidoAsync(claveacceso, correocliente, mensaje, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub EnviarCorreoDocumentoEmitidoAsync(ByVal claveacceso As String, ByVal correocliente As String, ByVal mensaje As String, ByVal userState As Object)
            If (Me.EnviarCorreoDocumentoEmitidoOperationCompleted Is Nothing) Then
                Me.EnviarCorreoDocumentoEmitidoOperationCompleted = AddressOf Me.OnEnviarCorreoDocumentoEmitidoOperationCompleted
            End If
            Me.InvokeAsync("EnviarCorreoDocumentoEmitido", New Object() {claveacceso, correocliente, mensaje}, Me.EnviarCorreoDocumentoEmitidoOperationCompleted, userState)
        End Sub
        
        Private Sub OnEnviarCorreoDocumentoEmitidoOperationCompleted(ByVal arg As Object)
            If (Not (Me.EnviarCorreoDocumentoEmitidoCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent EnviarCorreoDocumentoEmitidoCompleted(Me, New EnviarCorreoDocumentoEmitidoCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/IWSEDOC_ENVIARMAIL/ReenvioMailEnLinea", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function ReenvioMailEnLinea(ByVal MailEnLinea As ClsMailEnLinea, ByRef mensaje As String) As Boolean
            Dim results() As Object = Me.Invoke("ReenvioMailEnLinea", New Object() {MailEnLinea, mensaje})
            mensaje = CType(results(1),String)
            Return CType(results(0),Boolean)
        End Function
        
        '''<remarks/>
        Public Overloads Sub ReenvioMailEnLineaAsync(ByVal MailEnLinea As ClsMailEnLinea, ByVal mensaje As String)
            Me.ReenvioMailEnLineaAsync(MailEnLinea, mensaje, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub ReenvioMailEnLineaAsync(ByVal MailEnLinea As ClsMailEnLinea, ByVal mensaje As String, ByVal userState As Object)
            If (Me.ReenvioMailEnLineaOperationCompleted Is Nothing) Then
                Me.ReenvioMailEnLineaOperationCompleted = AddressOf Me.OnReenvioMailEnLineaOperationCompleted
            End If
            Me.InvokeAsync("ReenvioMailEnLinea", New Object() {MailEnLinea, mensaje}, Me.ReenvioMailEnLineaOperationCompleted, userState)
        End Sub
        
        Private Sub OnReenvioMailEnLineaOperationCompleted(ByVal arg As Object)
            If (Not (Me.ReenvioMailEnLineaCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent ReenvioMailEnLineaCompleted(Me, New ReenvioMailEnLineaCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/IWSEDOC_ENVIARMAIL/ResendArchive", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function ResendArchive(ByVal clave As String, ByVal info As ClsResendArchive, ByRef mensaje As String) As ResponseResendArchive
            Dim results() As Object = Me.Invoke("ResendArchive", New Object() {clave, info, mensaje})
            mensaje = CType(results(1),String)
            Return CType(results(0),ResponseResendArchive)
        End Function
        
        '''<remarks/>
        Public Overloads Sub ResendArchiveAsync(ByVal clave As String, ByVal info As ClsResendArchive, ByVal mensaje As String)
            Me.ResendArchiveAsync(clave, info, mensaje, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub ResendArchiveAsync(ByVal clave As String, ByVal info As ClsResendArchive, ByVal mensaje As String, ByVal userState As Object)
            If (Me.ResendArchiveOperationCompleted Is Nothing) Then
                Me.ResendArchiveOperationCompleted = AddressOf Me.OnResendArchiveOperationCompleted
            End If
            Me.InvokeAsync("ResendArchive", New Object() {clave, info, mensaje}, Me.ResendArchiveOperationCompleted, userState)
        End Sub
        
        Private Sub OnResendArchiveOperationCompleted(ByVal arg As Object)
            If (Not (Me.ResendArchiveCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent ResendArchiveCompleted(Me, New ResendArchiveCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
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
    Partial Public Class ClsMailEnLinea
        
        Private credencialField As String
        
        Private tipoField As TipoMail
        
        Private companiaField As String
        
        Private nicknameField As String
        
        Private claveAccesoField As String
        
        Private destinatarioField As String
        
        '''<remarks/>
        Public Property Credencial() As String
            Get
                Return Me.credencialField
            End Get
            Set
                Me.credencialField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Tipo() As TipoMail
            Get
                Return Me.tipoField
            End Get
            Set
                Me.tipoField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Compania() As String
            Get
                Return Me.companiaField
            End Get
            Set
                Me.companiaField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Nickname() As String
            Get
                Return Me.nicknameField
            End Get
            Set
                Me.nicknameField = value
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
        Public Property Destinatario() As String
            Get
                Return Me.destinatarioField
            End Get
            Set
                Me.destinatarioField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Public Enum TipoMail
        
        '''<remarks/>
        Cliente_RecuperarClave
        
        '''<remarks/>
        Emision_DocumentoEmitido
        
        '''<remarks/>
        Cliente_Bienvenida
        
        '''<remarks/>
        Emision_MailResponsable
        
        '''<remarks/>
        Recepcion_AutoResponse
        
        '''<remarks/>
        Recepcion_Asignacion
        
        '''<remarks/>
        Recepcion_Validacion
        
        '''<remarks/>
        Recepcion_ExpiroDia
        
        '''<remarks/>
        Recepcion_MailResponsable
        
        '''<remarks/>
        Recepcion_MailResponsableError
        
        '''<remarks/>
        Recepcion_OtrosAdjunto
        
        '''<remarks/>
        Recepcion_SinAdjunto
        
        '''<remarks/>
        Compania_RecuperarClave
        
        '''<remarks/>
        Compania_Bienvenida
        
        '''<remarks/>
        Emision_DocumentoProcesamiento5
        
        '''<remarks/>
        Emision_DocumentoProcesamiento7
        
        '''<remarks/>
        Emision_NotificarCorreoErroneo
        
        '''<remarks/>
        Compania_AlertaCertificado
        
        '''<remarks/>
        Recepcion_Notifica_Comercial
    End Enum
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ResponseResendArchive
        
        Private statusSendEmailField As Boolean
        
        Private statusSendAddressPrintField As Boolean
        
        Private messageSendMailField As String
        
        Private messageSendAddressPrintField As String
        
        '''<remarks/>
        Public Property StatusSendEmail() As Boolean
            Get
                Return Me.statusSendEmailField
            End Get
            Set
                Me.statusSendEmailField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property StatusSendAddressPrint() As Boolean
            Get
                Return Me.statusSendAddressPrintField
            End Get
            Set
                Me.statusSendAddressPrintField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property MessageSendMail() As String
            Get
                Return Me.messageSendMailField
            End Get
            Set
                Me.messageSendMailField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property MessageSendAddressPrint() As String
            Get
                Return Me.messageSendAddressPrintField
            End Get
            Set
                Me.messageSendAddressPrintField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.8.9037.0"),  _
     System.SerializableAttribute(),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ClsResendArchive
        
        Private sendEmailField As Boolean
        
        Private sendAddressPrintField As Boolean
        
        Private emailField As String
        
        Private addressPrintField As String
        
        Private billNumberField As String
        
        Private dateOfIssueField As Date
        
        '''<remarks/>
        Public Property SendEmail() As Boolean
            Get
                Return Me.sendEmailField
            End Get
            Set
                Me.sendEmailField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property SendAddressPrint() As Boolean
            Get
                Return Me.sendAddressPrintField
            End Get
            Set
                Me.sendAddressPrintField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property Email() As String
            Get
                Return Me.emailField
            End Get
            Set
                Me.emailField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property AddressPrint() As String
            Get
                Return Me.addressPrintField
            End Get
            Set
                Me.addressPrintField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property BillNumber() As String
            Get
                Return Me.billNumberField
            End Get
            Set
                Me.billNumberField = value
            End Set
        End Property
        
        '''<remarks/>
        Public Property DateOfIssue() As Date
            Get
                Return Me.dateOfIssueField
            End Get
            Set
                Me.dateOfIssueField = value
            End Set
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub EnviarCorreoDocumentoEmitidoCompletedEventHandler(ByVal sender As Object, ByVal e As EnviarCorreoDocumentoEmitidoCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class EnviarCorreoDocumentoEmitidoCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property mensaje() As String
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),String)
            End Get
        End Property
        
        '''<remarks/>
        Public ReadOnly Property EnviarCorreoDocumentoEmitidoResult() As Boolean
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(1),Boolean)
            End Get
        End Property
        
        '''<remarks/>
        Public ReadOnly Property EnviarCorreoDocumentoEmitidoResultSpecified() As Boolean
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(2),Boolean)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub ReenvioMailEnLineaCompletedEventHandler(ByVal sender As Object, ByVal e As ReenvioMailEnLineaCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class ReenvioMailEnLineaCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As Boolean
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),Boolean)
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
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0")>  _
    Public Delegate Sub ResendArchiveCompletedEventHandler(ByVal sender As Object, ByVal e As ResendArchiveCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.8.9037.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class ResendArchiveCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As ResponseResendArchive
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),ResponseResendArchive)
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

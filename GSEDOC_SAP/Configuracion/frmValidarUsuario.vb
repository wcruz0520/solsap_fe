Imports System.Net

'HTTPS
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Public Class frmValidarUsuario

    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioValidarUsuario()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmValidarUsuario") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmValidarUsuario.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmValidarUsuario").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmValidarUsuario")
            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent
        Try
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                  And pVal.FormTypeEx = "frmValidarUsuario" Then
                If pVal.BeforeAction = False And pVal.ItemUID = "btnIngre" Then

                    oForm = rsboApp.Forms.Item("frmValidarUsuario")

                    '----- Consumo WS licencias
                    Dim wsSSLIC As New Entidades.wsSS_LICENCIA_SAP.Licencia

                    ' MANEJO PROXY
                    Dim SALIDA_POR_PROXY As String = ""
                    SALIDA_POR_PROXY = ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "PROXY", "PROXY")
                    Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
                    Dim Proxy_puerto As String = ""
                    Dim Proxy_IP As String = ""
                    Dim Proxy_Usuario As String = ""
                    Dim Proxy_Clave As String = ""

                    If SALIDA_POR_PROXY = "Y" Then
                        Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                        Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                        Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                        Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")
                        Dim proxyobject As System.Net.WebProxy = Nothing
                        Dim cred As System.Net.NetworkCredential = Nothing

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        wsSSLIC.Proxy = proxyobject
                        wsSSLIC.Credentials = cred
                    End If
                    ' END  MANEJO PROXY
                    Dim TimeOutEmision As String = Functions.VariablesGlobales._gTimeOut_Emision

                    If TimeOutEmision = "" Then
                        wsSSLIC.Timeout = 3600
                    Else
                        wsSSLIC.Timeout = Integer.Parse(TimeOutEmision)
                    End If

                    '------ Clave
                    Dim txtClave As SAPbouiCOM.EditText
                    txtClave = oForm.Items.Item("txtClave").Specific

                    '----- Usuario
                    Dim txtUser As SAPbouiCOM.EditText
                    txtUser = oForm.Items.Item("txtUser").Specific

                    '----- Motivo Ingreso
                    Dim txtMotivo As SAPbouiCOM.EditText
                    txtMotivo = oForm.Items.Item("txtMotivo").Specific

                    If txtMotivo.Value = "" Then
                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - Por favor ingresar un Motivo..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        Exit Sub
                    End If

                    '---------- llamado to Home
                    '---------- Identificacion

                    'Creo obj Credencial
                    Dim objCRED As New Entidades.wsSS_LICENCIA_SAP.Credencial

                    With objCRED
                        .Usuario = txtUser.Value
                        .Clave = txtClave.Value
                        .IPLocal = rCompany.Server
                        .MotivoIngreso = "Pantalla Administracion Addon: " + NombreAddon + " - Motivo: " + txtMotivo.Value
                        .NombreBD = rCompany.CompanyDB
                        .IPPublica = rCompany.LicenseServer
                        .NombreAddon = NombreAddon
                        .NombreCliente = rCompany.CompanyName
                        .Pais = CodigoPais
                        .VersionAddon = rEstructura.VersionAddon

                    End With
                    ' Consumo el metodo de autenticacion
                    Dim respWS As Boolean = False
                    Dim msgg As String = ""

                    Dim URLWS As String = ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "WsLicencia")
                    If Not String.IsNullOrEmpty(URLWS) Then
                        wsSSLIC.Url = URLWS
                        Utilitario.Util_Log.Escribir_Log("URLWS: " + URLWS.ToString(), "frmValidarUsuario")
                        'rSboApp.StatusBar.SetText(wsSSLIC.Url, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

                    Try
                        If Functions.VariablesGlobales._vgHttps = "Y" Then
                            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)

                        End If

                        respWS = wsSSLIC.ValidarClave(objCRED, msgg)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("ex ValidarClave: " + ex.Message.ToString(), "frmValidarUsuario")
                        msgg = "Error WS Autenticacion ! verificar la Conexion"
                    End Try

                    'If txtClave.Value.ToString().Equals("S0ls@p2o1R*") Then
                    If respWS Then
                        ofrmConfMenu.CargaFormularioMenuDeConfiguraciones()
                        oForm.Close()
                    Else
                        If msgg = "" Then
                            rsboApp.StatusBar.SetText(NombreAddon + " - Clave Incorrecta!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            rsboApp.StatusBar.SetText(NombreAddon + " - " + msgg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmValidarUsuario")
            System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try

    End Sub


    Private Function RecorreFormulario(ByVal oApp As SAPbouiCOM.Application, ByVal Formulario As String) As Boolean
        Try
            For Each oForm In oApp.Forms
                Select Case oForm.UniqueID
                    Case Formulario
                        oForm.Visible = True
                        oForm.Select()
                        Return True
                End Select
            Next

            For Each oForm In oApp.Forms
                If oForm.UniqueID = Formulario Then
                    oForm.Visible = True
                    oForm.Select()
                    ' oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Shared Function customCertValidation(ByVal sender As Object, _
                                                 ByVal cert As X509Certificate, _
                                                 ByVal chain As X509Chain, _
                                                 ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function
End Class

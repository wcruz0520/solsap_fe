Imports System.Net

Public Class frmConfClaveLE

    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioValidarClave()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConfClaveLE") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmConfClaveLE.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConfClaveLE").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmConfClaveLE")
            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                  And pVal.FormTypeEx = "frmConfClaveLE" Then
                If pVal.BeforeAction = False And pVal.ItemUID = "btnIngre" Then

                    oForm = rsboApp.Forms.Item("frmConfClaveLE")

                    '----- Consumo WS licencias
                    Dim wsSSLIC As New Entidades.wsSS_LICENCIA_SAP.Licencia

                    ' MANEJO PROXY
                    Dim SALIDA_POR_PROXY As String = ""
                    SALIDA_POR_PROXY = ofrmParametrosAddonLE.ConsultaParametro(Functions.VariablesGlobales._gNombreAddOn, "PARAMETROS", "PROXY", "PROXY")
                    Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
                    Dim Proxy_puerto As String = ""
                    Dim Proxy_IP As String = ""
                    Dim Proxy_Usuario As String = ""
                    Dim Proxy_Clave As String = ""

                    If SALIDA_POR_PROXY = "Y" Then
                        Proxy_puerto = ofrmParametrosAddonLE.ConsultaParametro(Functions.VariablesGlobales._gNombreAddOn, "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ofrmParametrosAddonLE.ConsultaParametro(Functions.VariablesGlobales._gNombreAddOn, "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ofrmParametrosAddonLE.ConsultaParametro(Functions.VariablesGlobales._gNombreAddOn, "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ofrmParametrosAddonLE.ConsultaParametro(Functions.VariablesGlobales._gNombreAddOn, "PARAMETROS", "PROXY", "PROXY_CLAVE")

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

                    Dim URLWS As String = ofrmParametrosAddonLE.ConsultaParametro(Functions.VariablesGlobales._gNombreAddOn, "PARAMETROS", "CONFIGURACION", "WS_Licencia")
                    If Not String.IsNullOrEmpty(URLWS) Then
                        wsSSLIC.Url = URLWS
                        Utilitario.Util_Log.Escribir_Log("URLWS: " + URLWS.ToString(), "frmConfClave")
                        'rSboApp.StatusBar.SetText(wsSSLIC.Url, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

                    Try
                        If Not Functions.VariablesGlobales._gNO_ConsumirMetodoHTTPS = "Y" Then
                            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                        End If

                        respWS = wsSSLIC.ValidarClave(objCRED, msgg)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("ex ValidarClave: " + ex.Message.ToString(), "frmConfClave")
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
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmConfClave")
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

End Class

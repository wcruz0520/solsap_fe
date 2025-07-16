Option Strict Off
Option Explicit On
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports System.Xml

'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
'
Module SubMain

    Public Const NombreAddon As String = "SAED"
    Public Const NombreAddonLOC As String = "eLocalizacion"
    Public Const sKey As String = "S01s7p1" ' CLAVE DE ENCRIPTACION LICENCIA S01s7p1
    Public Const CodigoPais As String = "EC"
    'Public DirDelLog As String = IO.Path.GetDirectoryName(Application.ExecutablePath) + "LOG" ' RUTA DEL LOG
    Public Nombre_Proveedor_SAP_BO As String = ""

    Private Const conectarse As Boolean = True
    Public oFuncionesB1 As Functions.FuncionesB1
    Public oFuncionesAddon As Functions.FuncionesAddon
    Public oManejoDocumentos As Negocio.ManejoDeDocumentos

    Public oManejoDocumentosEcua As Negocio.ManejoDeDocumentosEcua

    'JP 17/06/2025
    Public oManejoDocumentosSolsap As Negocio.ManejoDeDocumentoSolsap

    'EMISION
    Public ofrmDocumentosEnviados As frmDocumentosEnviados
    Public ofrmLogEmision As frmLogEmision
    Public ofrmImpresionPorBloque As frmImpresionPorBloque

    'RECEPCION
    Public ofrmDocumentosRecibidos As frmDocumentosRecibidos
    Public ofrmMapeo As frmMapeo
    Public ofrmDocumento As frmDocumento
    Public ofrmDocumentoNC As frmDocumentoNC
    Public ofrmDocumentoRE As frmDocumentoRE
    Public ofrmConsultaOrdenes As frmConsultaOrdenes
    Public ofrmDocumentosIntegrados As frmDocumentosIntegrados
    Public ofrmParametrosRecepcion As frmParametrosRecepcion
    Public ofrmSubirArchivo As frmSubirArchivo
    Public ofrmProcesoLote As frmProcesoLote
    Public ofrmProcesoLote2 As frmProcesoLote2

    Public ofrmAcercaDe As frmAcercaDe

    ' OPCIONES DE CONFIGURACION / PARAMETRIZACIONES
    Public ofrmConfClave As frmConfClave
    Public ofrmConfMenu As frmConfMenu
    Public ofrmParametrosAddon As frmParametrosAddon
    Public ofrmProxy As frmProxy
    'add Artur
    Public ofrmConsultasDB As frmConsultasDB
    Public ofrmConsultasDB_RE As frmConsultasDB_RE

    Public oLicencia As New Licencia
    Public ofrmValidarUsuario As frmValidarUsuario
    Public ofrmProcesoLoteManamer As frmProcesoLoteManamer

    Public ofrmDocumentosRecibidosXML As frmDocumentosRecibidosXML
    Public ofrmDocumentoXML As frmDocumentoXML
    Public ofrmDocumentoNCXML As frmDocumentoNCXML
    Public ofrmDocumentoREXML As frmDocumentoREXML
    Public ofrmProcesoLoteXML As frmProcesoLoteXML

    Public ofrmProcesoLoteC As frmProcesoLoteC
    Public OfrmListaAEnviar As frmListaAEnviar


#Region "Variables de Addon"
    Public oFiltros As SAPbouiCOM.EventFilters
    Public oFiltro As SAPbouiCOM.EventFilter

    Public rSboApp As SAPbouiCOM.Application
    Public rSboGui As SAPbouiCOM.SboGuiApi
    Public rCompany As SAPbobsCOM.Company
    Public roFuncionesB1Srv As SAPbobsCOM.CompanyService
    ' Public oSapBobs As SAPbobsCOM.SBObob
    'Public oUserBusFor As SAPbobsCOM.FormattedSearches
    Public rEventoEmision As EventosEmision
    Public rEventoRecepion As EventosRecepcion
    Public rEvento As Eventos
    Public rEstructura As Estructura
    Public pMan As SAPbouiCOM.Application

    Dim strSQL As String
    Dim ret As Integer

#End Region

#Region "Variables de Localizacion"

    'Tomado de Localizacion

    ' Public rEstructura As EstructuraL_EC
    '  Public rEvento As EventosLE
    ' Private Const conectarse As Boolean = True

    'Dim strSQL As String
    'Dim ret As Integer

    'Public oFuncionesB1 As Funciones_SAP.FuncionesB1
    'Public oFuncionesAddon As Functions.FuncionesAddon

    Public ofrmAcercaDeLE As frmAcercaDeLE
    Public ofrmConfClaveLE As frmConfClaveLE
    Public ofrmConfMenuLE As frmConfMenuLE
    Public ofrmParametrosAddonLE As frmParametrosAddonLE
    Public ofrmClaveLE As frmClave
    Public ofrmGeneradorATS As frmGeneradorATS
    Public ofrmGenerarRPT As frmGenerarRPT

    'add artur 08092022
    Public ofrmConsultasDbLE As frmConsultasDbLE

    'add artur 09092022 carga de RPTS

    Public ofrmCargaRPT As frmCargaRPT

    'add artur 12092022
    Public ofrmAnexoCompras As frmAnexoCompras

    Public ofrmAnexoVentas As frmAnexoVentas

    'add artur 13092022

    Public ofrmSRI As frmSRI
    Public ofrmSRIConsulta As frmSRIConsulta
    Public ofrmDinardap As frmDinardap
    '--------------FILTROS---------------------------
    'Public oFiltros As SAPbouiCOM.EventFilters
    'Public oFiltro As SAPbouiCOM.EventFilter

    ' MANEJO DE CHEQUES PROTESTOS RO
    Public ofrmChequeP As frmChequeP
    Public ofrmChequePD As frmChequePD

    'MANEJOR TRANSFERENCIAS ENTRE COMPAÑIAS
    Public ofrmTransEntreCompanias As frmTransEntreCompanias
    Public ofrmConsultaDetalleTrans As frmConsultaDetalleTrans
    Public ofrmConsultaSalidaEntrada As frmConsultaSalidaEntrada
    Public ofrmConsultaBodega As frmConsultaBodega


    'Agregados 14052024
    'ReyArturo

    Public ofrmGuiasRemision As frmGuiasRemision
    Public ofrmPagosMasivos As frmPagosMasivos
    Public ofrmPagosAprobacion As frmPagosAprobacion
    Public ofrmImprimir As frmImprimir
    'add 02072024
    Public ofrmCashManagement As frmCashManagemet
    Public ofrmMapeoCuentas As frmMapeoCuentasCM

    'add 10092024
    Public ofrmServiciosBasicos As frmServiciosBasicos

#End Region

    Public Sub main()
        Try

            Dim strTest(4) As String, sCookie As String
            Dim strConnString As String
            Dim textoMensajeSinLicencia As String = ""

            strConnString = vbNullString
            strTest = System.Environment.GetCommandLineArgs()

            ' Validaciones de seguridad del AddOn
            If strTest.Length > 0 Then
                If strTest.Length > 1 Then
                    If strTest(0).LastIndexOf("\") > 0 Then
                        strConnString = strTest(1)
                    Else
                        strConnString = strTest(0)
                    End If
                Else
                    If strTest(0).LastIndexOf("\") = -1 Then
                        strConnString = strTest(0)
                    Else
                        System.Windows.Forms.MessageBox.Show("El Add-on se debe ejecutar desde SAP Business One. (" & NombreAddon & "-Err1)")
                        End
                    End If
                End If
            Else
                System.Windows.Forms.MessageBox.Show("El Add-on se debe ejecutar desde SAP Business One. (" & NombreAddon & "-Err2)")
                End
            End If

            ' Conexión
            If strConnString.Length > 0 Then
                Try
                    rSboGui = New SAPbouiCOM.SboGuiApi
                    rCompany = New SAPbobsCOM.Company()
                    ' Conexión con el UI
                    rSboGui.Connect(strConnString)
                    rSboApp = rSboGui.GetApplication()

                    ' Conexión con el DI
                    If conectarse Then
                        rCompany = rSboApp.Company.GetDICompany()
                    Else
                        sCookie = rCompany.GetContextCookie
                        ret = rCompany.SetSboLoginContext(rSboApp.Company.GetConnectionContext(sCookie))
                        If ret = 0 Then
                            ret = rCompany.Connect()
                            If ret <> 0 Then
                                rCompany.GetLastError(ret, strSQL)
                                rSboApp.StatusBar.SetText("Error al Conectar el Add-On " & NombreAddon & ": " & strSQL, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Utilitario.Util_Log.Escribir_Log("Error al Conectar el Add-On " & NombreAddon & ": " & strSQL, "SubMain")
                                End
                            End If
                        Else
                            rSboApp.StatusBar.SetText("No se ha Conectado con AddOn " & NombreAddon & ": Error " & ret, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log("No se ha Conectado con AddOn " & NombreAddon & ": Error " & ret, "SubMain")
                            End
                        End If

                    End If
                    ' **** INICIALIZACIÓN DE LAS CLASES ****

                    'ESTRUCTURA DE DATOS / BUSQ. FORMATEADAS
                    rEstructura = New Estructura
                    'pEst = Est.rSboApp

                    'EVENTOS GENERALES - ACERCA DE
                    rEvento = New Eventos

                    ' FORMULARIOS DE CONFIGURACION / PARAMETRIZACION
                    ofrmAcercaDe = New frmAcercaDe(rCompany, rSboApp)
                    ofrmConfClave = New frmConfClave(rCompany, rSboApp)
                    ofrmConfMenu = New frmConfMenu(rCompany, rSboApp)
                    ofrmParametrosAddon = New frmParametrosAddon(rCompany, rSboApp)
                    ofrmProxy = New frmProxy(rCompany, rSboApp)
                    ofrmConsultasDB = New frmConsultasDB(rCompany, rSboApp)
                    ofrmConsultasDB_RE = New frmConsultasDB_RE(rCompany, rSboApp)

                    Functions.VariablesGlobales._vgNombreAddOn = NombreAddon
                    Functions.VariablesGlobales._vgVersionAddOn = rEstructura.VersionAddon

                    Dim RsFe As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        'CONFIGURACION
                        Try
                            Dim QueryDias As String = ""
                            QueryDias = "Select Top 1 ""Code"" from ""@SS_DIAS_PARAM"" where ""Name"" ='" + rCompany.UserName + "'"
                            RsFe.DoQuery(QueryDias)
                            While RsFe.EoF = False
                                Functions.VariablesGlobales._UsuarioParamDias = RsFe.Fields.Item("Code").Value.ToString

                                RsFe.MoveNext()
                            End While
                        Catch ex As Exception
                            Functions.VariablesGlobales._UsuarioParamDias = ""
                        End Try

                        Dim Query As String = ""

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = '" & Functions.VariablesGlobales._vgNombreAddOn & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'CONFIGURACION'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = '" & Functions.VariablesGlobales._vgNombreAddOn & "' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'CONFIGURACION'"
                        End If



                        RsFe.DoQuery(Query)

                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString

                                Case "GuardarLog"
                                    Functions.VariablesGlobales._vgGuardarLog = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "foliacionLQ"
                                    Functions.VariablesGlobales._vgFolioLQUDF = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ImpresionBloque"
                                    Functions.VariablesGlobales._vgImpBlo = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PreliminaresLote"
                                    Functions.VariablesGlobales._vgPreLot = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "VisualizaPDF_Bytes"
                                    Functions.VariablesGlobales._VisualizaPDFByte = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FechaSalidaEnVivo"
                                    Functions.VariablesGlobales._vgFechaSalidaEnVivo = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ProcesoLoteManamer"
                                    Functions.VariablesGlobales._vgProcesoLoteManamer = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "NoEnviarRT"
                                    Functions.VariablesGlobales._vgNoEnviarRT = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "BloquearReenviarSRI"
                                    Functions.VariablesGlobales._vgBloquearReenviarSRI = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ServerNode"
                                    Functions.VariablesGlobales._vgServerNode = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "BD_User"
                                    Functions.VariablesGlobales._vgUserBD = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "BD_Pass"
                                    Functions.VariablesGlobales._vgPassBD = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "MostrarLogoSolSap"
                                    Functions.VariablesGlobales._vgMostrarLogo = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "QUERY_CORREO"
                                    Functions.VariablesGlobales._vgQueryCorreo = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_RecepcionConsulta"
                                    Functions.VariablesGlobales._WS_Recepcion = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_RecepcionEstado"
                                    Functions.VariablesGlobales._WS_RecepcionCambiarEstado = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RecepcionClave"
                                    Functions.VariablesGlobales._WS_RecepcionClave = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Estados_docs"
                                    Functions.VariablesGlobales._WS_RecepcionCargaEstados = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_RecepcionConsultaArchivo"
                                    Functions.VariablesGlobales._WS_RecepcionConsultaArchivo = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Registros_por_paginas"
                                    Functions.VariablesGlobales._nRegistros = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Adicional_FC"
                                    Functions.VariablesGlobales._Adicional_FC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Adicional_NC"
                                    Functions.VariablesGlobales._Adicional_NC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Adicional_RET"
                                    Functions.VariablesGlobales._Adicional_RET = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Nombre_CA"
                                    Functions.VariablesGlobales._Nombre_CA = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RecepcionLite"
                                    Functions.VariablesGlobales._RecepcionLite = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Ruta_Compartida"
                                    Functions.VariablesGlobales._Ruta_Compartida = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PagoRecibido_Seidor_exxis"
                                    Functions.VariablesGlobales._PagoRecibido_Seidor_exxis = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "AsignarNumeroDocEnNumAtCard"
                                    Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "CrearFacturaDeResarvaProveedores"
                                    Functions.VariablesGlobales._CrearFCdeReservaProveedores = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "NombreCampoNumRet"
                                    Functions.VariablesGlobales._CampoNumRetencion = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "EmisionTipo"
                                    Functions.VariablesGlobales._TipoEmision = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "EmisionClave"
                                    Functions.VariablesGlobales._wsClaveEmision = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "TipoWebServices"
                                    Functions.VariablesGlobales._TipoWS = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionFC"
                                    Functions.VariablesGlobales._wsEmisionFactura = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionGuia"
                                    Functions.VariablesGlobales._wsEmisionGuiaRemision = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_LiquidacionCompra"
                                    Functions.VariablesGlobales._wsEmisionLiquidacionCompra = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionNC"
                                    Functions.VariablesGlobales._wsEmisionNotaCredito = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionND"
                                    Functions.VariablesGlobales._wsEmisionNotaDebito = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionRetencion"
                                    Functions.VariablesGlobales._wsEmisionRetencion = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionConsulta"
                                    Functions.VariablesGlobales._wsConsultaEmision = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WS_EmisionReenvioMail"
                                    Functions.VariablesGlobales._wsReenvioMail = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FoliacionPOSTN"
                                    Functions.VariablesGlobales._FoliacionPostin = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "NumeroDocEnNumAtCardCancelar"
                                    Functions.VariablesGlobales._AnuDocNumAtCard = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "EnviaDocumentosEnBackGround"
                                    Functions.VariablesGlobales._EnviarBackGroung = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "XMLRecepcionHesion"
                                    Functions.VariablesGlobales._XMLRecepcionHeison = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RutaFC"
                                    Functions.VariablesGlobales._RutaFC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RutaNC"
                                    Functions.VariablesGlobales._RutaNC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RutaRT"
                                    Functions.VariablesGlobales._RutaRT = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RutaProFC"
                                    Functions.VariablesGlobales._RutaProFC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RutaProNC"
                                    Functions.VariablesGlobales._RutaProNC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RutaProRT"
                                    Functions.VariablesGlobales._RutaProRT = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PreliminarLotesXML"
                                    Functions.VariablesGlobales._PreliminarLoteXML = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ImprimirDocAut"
                                    Functions.VariablesGlobales._ImpDocAut = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "IntegracionEcuanexus"
                                    Functions.VariablesGlobales._IntegracionEcuanexus = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WsEmisionEcu"
                                    Functions.VariablesGlobales._WsEmisionEcua = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WsConsultaEcu"
                                    Functions.VariablesGlobales._WsEmisionConsultaEcua = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "TokenEcu"
                                    Functions.VariablesGlobales._Token = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "NombreWsEcu"
                                    Functions.VariablesGlobales._NombreWsEcua = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "QrySecuencial"
                                    Functions.VariablesGlobales._ConsultaFolioSS = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "AsignarFolioReenvio"
                                    Functions.VariablesGlobales._AsignarFolioalReenviarSolsap = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "SINCRO_RET"
                                    Functions.VariablesGlobales._SINCRO_RT = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "SINCRO_LQE"
                                    Functions.VariablesGlobales._SINCRO_LQE = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "SINCRO_DOC"
                                    Functions.VariablesGlobales._SINCRO_DOC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ReenviarDoc"
                                    Functions.VariablesGlobales._ReenviarDocsPantala = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ReenviarListaDocEnv"
                                    Functions.VariablesGlobales._ReenviarListaDocEnv = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ImpresionDobleCara"
                                    Functions.VariablesGlobales._ImpresionDobleCara = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ProveedorSAP"
                                    Functions.VariablesGlobales._ProveedorSAP = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "SalidaporHttps"
                                    Functions.VariablesGlobales._vgHttps = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "WsLicencia"
                                    Functions.VariablesGlobales._WsLicencia = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "RucCompañia"
                                    Functions.VariablesGlobales._RucCompañia = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "TipoWsLicencia"
                                    Functions.VariablesGlobales._TipoWsLicencia = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ValidarCamposNulos"
                                    Functions.VariablesGlobales._ValidarCamposNulos = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'add ArturDev 05042024
                                Case "ActivarLocalizacionEC"
                                    Functions.VariablesGlobales._ActivarLocalizacionEC = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "RutaIntegracionXML"
                                    Functions.VariablesGlobales._RutaIntegracionXML = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "TablasNativasReplace"
                                    Functions.VariablesGlobales._TablasNativasReplace = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'ADD 09072024
                                Case "ActivarCMFML"
                                    Functions.VariablesGlobales._ActivarCMFML = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "QueryGRUdo"
                                    Functions.VariablesGlobales._SINCRO_GRUDO = RsFe.Fields.Item("U_Valor").Value.ToString

                                     'ADD 03/09/2024 
                                Case "CantabilizarPRProcLot"
                                    Functions.VariablesGlobales._ContabilizarPRPL = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'ADD DM 17/12/2024 
                                Case "NomCampoPedInfoAdicional"
                                    Functions.VariablesGlobales._NombreCampoPedidoInfoAd = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'ADD DM 20/02/2025
                                Case "MostrarFechaAutorizacion"
                                    Functions.VariablesGlobales._MostrarFechaAutorizacion = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "SeriesFEUDF"
                                    Functions.VariablesGlobales._vgSerieUDF = RsFe.Fields.Item("U_Valor").Value.ToString

                                    '16/06/2025
                                Case "ActivaIntSS"
                                    Functions.VariablesGlobales._ActApiSS = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "APIAUTSS"
                                    Functions.VariablesGlobales._ApiAutSS = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "APIFACTSS"
                                    Functions.VariablesGlobales._ApiFactEmiSS = RsFe.Fields.Item("U_Valor").Value.ToString
                            End Select

                            RsFe.MoveNext()
                        End While

                        'Se Recupera La Seccion de Querys
                        'Add Artur

                        Query = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = '" & NombreAddon & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'BD'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = '" & NombreAddon & "' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'BD'"
                        End If
                        RsFe.DoQuery(Query)
                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString
                                'Querys Encryptados
                                Case "Query_FacturaSeccion01"
                                    Functions.VariablesGlobales._Query_FacturaSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_FacturaSeccion02"
                                    Functions.VariablesGlobales._Query_FacturaSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_FacturaAnticipoSeccion01"
                                    Functions.VariablesGlobales._Query_FacturaAnticipoSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_FacturaAnticipoSeccion02"
                                    Functions.VariablesGlobales._Query_FacturaAnticipoSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_NotaCreditoSeccion01"
                                    Functions.VariablesGlobales._Query_NotaCreditoSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_NotaCreditoSeccion02"
                                    Functions.VariablesGlobales._Query_NotaCreditoSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_NotaDebitoSeccion01"
                                    Functions.VariablesGlobales._Query_NotaDebitoSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString
                                    'Ruta integracion XML

                                Case "Query_NotaDebitoSeccion02"
                                    Functions.VariablesGlobales._Query_NotaDebitoSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString


                                Case "Query_CompleExportacion"
                                    Functions.VariablesGlobales._Query_CompleExportacion = RsFe.Fields.Item("U_Valor").Value.ToString
                                    'Ruta integracion XML

                                Case "Query_CompleReembolso"
                                    Functions.VariablesGlobales._Query_CompleReembolso = RsFe.Fields.Item("U_Valor").Value.ToString


                                Case "Query_GuiaRemisionSeccion01"
                                    Functions.VariablesGlobales._Query_GuiaRemisionSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString
                                    'Ruta integracion XML

                                Case "Query_GuiaRemisionSeccion02"
                                    Functions.VariablesGlobales._Query_GuiaRemisionSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_RetencionSeccion01"
                                    Functions.VariablesGlobales._Query_RetencionSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString
                                    'Ruta integracion XML

                                Case "Query_RetencionSeccion02"
                                    Functions.VariablesGlobales._Query_RetencionSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_LiquidacionSeccion01"
                                    Functions.VariablesGlobales._Query_LiquidacionSeccion01 = RsFe.Fields.Item("U_Valor").Value.ToString


                                Case "Query_LiquidacionSeccion02"
                                    Functions.VariablesGlobales._Query_LiquidacionSeccion02 = RsFe.Fields.Item("U_Valor").Value.ToString


                                Case "Query_DocumentosEnviados"
                                    Functions.VariablesGlobales._Query_DocumentosEnviados = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'Query Recepcion

                                Case "Query_ReDocumentosMarcados"
                                    Functions.VariablesGlobales._Query_ReDocumentosMarcados = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "Query_ReDocumentosIntegrados"
                                    Functions.VariablesGlobales._Query_ReDocumentosIntegrados = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'Guias desatendidas

                                Case "Query_GuiasDesatendidas01"
                                    Functions.VariablesGlobales._Query_GuiasDesatendidas01 = RsFe.Fields.Item("U_Valor").Value.ToString


                                Case "Query_GuiasDesatendidas02"
                                    Functions.VariablesGlobales._Query_GuiasDesatendidas02 = RsFe.Fields.Item("U_Valor").Value.ToString




                            End Select
                            RsFe.MoveNext()
                        End While



                        '-- fin --- querys-------------------


                        'PROXY
                        Query = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = '" & Functions.VariablesGlobales._vgNombreAddOn & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'PROXY'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = '" & Functions.VariablesGlobales._vgNombreAddOn & "' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'PROXY'"
                        End If
                        RsFe.DoQuery(Query)
                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString
                                Case "PROXY_IP"
                                    Functions.VariablesGlobales._vgProxy_IP = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PROXY_USER"
                                    Functions.VariablesGlobales._vgProxy_Usuario = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PROXY_CLAVE"
                                    Functions.VariablesGlobales._vgProxy_Clave = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PROXY_PUERTO"
                                    Functions.VariablesGlobales._vgProxy_puerto = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PROXY"
                                    Functions.VariablesGlobales._SALIDA_POR_PROXY = RsFe.Fields.Item("U_Valor").Value.ToString

                            End Select
                            RsFe.MoveNext()
                        End While

                        'RECEPCION FACTURA
                        Query = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'FC'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'FC'"
                        End If
                        RsFe.DoQuery(Query)
                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString
                                Case "FechaEmisionFactura"
                                    Functions.VariablesGlobales._vgFechaEmisionFactura = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FechaEmisionFacturaP"
                                    Functions.VariablesGlobales._vgFechaEmisionFacturaP = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "CreaPedido"
                                    Functions.VariablesGlobales._CreaPedido = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "PermiteDescuadre"
                                    Functions.VariablesGlobales._PermiteDescuadre = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Prefijo"
                                    Functions.VariablesGlobales._Prefijo_FC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Cuenta"
                                    Functions.VariablesGlobales._Cuenta_FC = RsFe.Fields.Item("U_Valor").Value.ToString
                                        ' Case "Cuenta"
                                       ' Functions.VariablesGlobales._Cuenta_RE = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "MarcarDocFC"
                                    Functions.VariablesGlobales._MarcarContabiliadosManualFC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FormaPagoCompras"
                                    Functions.VariablesGlobales._FormaPagoCompras = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "ValorFormaPagoCompras"
                                    Functions.VariablesGlobales._ValorFormaPagoCompras = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "FechaAutEnFechaContabFC"
                                    Functions.VariablesGlobales._FechaAutEnFechaContabFC = RsFe.Fields.Item("U_Valor").Value.ToString

                            End Select
                            RsFe.MoveNext()
                        End While

                        'RECEPCION NOTA CREDITO
                        Query = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'NC'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'NC'"
                        End If
                        RsFe.DoQuery(Query)
                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString
                                Case "FechaEmisionNotaCredito"
                                    Functions.VariablesGlobales._vgFechaEmisionNotaCredito = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FechaEmisionNotaCreditoP"
                                    Functions.VariablesGlobales._vgFechaEmisionNotaCreditoP = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Prefijo"
                                    Functions.VariablesGlobales._Prefijo_NC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "Cuenta"
                                    Functions.VariablesGlobales._Cuenta_NC = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "MarcarDocNC"
                                    Functions.VariablesGlobales._MarcarContabiliadosManualNC = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "FechaAutEnFechaContabNC"
                                    Functions.VariablesGlobales._FechaAutEnFechaContabNC = RsFe.Fields.Item("U_Valor").Value.ToString

                            End Select
                            RsFe.MoveNext()
                        End While

                        'RECEPCION RETENCION
                        Query = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'RE'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'RE'"
                        End If
                        RsFe.DoQuery(Query)
                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString
                                Case "FechaEmisionRetencion"
                                    Functions.VariablesGlobales._vgFechaEmisionRetencion = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FechaEmisionRetencionP"
                                    Functions.VariablesGlobales._vgFechaEmisionRetencionP = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "BFR"
                                    Functions.VariablesGlobales._BFR = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "CodigoRetencion"
                                    Functions.VariablesGlobales._CodigoRetencion = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "CodigoRetencionR"
                                    Functions.VariablesGlobales._CodigoRetencionR = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "MarcarDocRT"
                                    Functions.VariablesGlobales._MarcarContabiliadosManualRT = RsFe.Fields.Item("U_Valor").Value.ToString
                                    'Case "ValidarFechasCTK"
                                    '    Functions.VariablesGlobales._ValidarFechasCTK = RsFe.Fields.Item("U_Valor").Value.ToString
                                    'Case "DiasValidarProcesoLote"
                                    '    Functions.VariablesGlobales._DiasValidarProcesoLote = RsFe.Fields.Item("U_Valor").Value.ToString
                                Case "FechaAutEnFechaContabRT"
                                    Functions.VariablesGlobales._FechaAutEnFechaContabRT = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "ComentarioPago"
                                    Functions.VariablesGlobales._ComentarioPago = RsFe.Fields.Item("U_Valor").Value.ToString

                                Case "FechaFinMesAnterior"
                                    Functions.VariablesGlobales.FechaFinMesAnterior = RsFe.Fields.Item("U_Valor").Value.ToString

                            End Select
                            RsFe.MoveNext()
                        End While


                        'PROCESO LOTE
                        Query = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                            Query += "FROM ""@GS_CONFD"" A INNER JOIN "
                            Query += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                            Query += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                            Query += " AND B.""U_Subtipo"" = 'PL'"
                        Else
                            Query = "SELECT A.U_Nombre,A.U_Valor "
                            Query += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                            Query += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                            Query += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                            Query += " AND  B.U_Subtipo = 'PL'"
                        End If
                        RsFe.DoQuery(Query)
                        While RsFe.EoF = False
                            Select Case RsFe.Fields.Item("U_Nombre").Value.ToString

                                Case "ValidarFechasCTK"
                                    Functions.VariablesGlobales._ValidarFechasCTK = IIf(String.IsNullOrEmpty(RsFe.Fields.Item("U_Valor").Value.ToString), "N", RsFe.Fields.Item("U_Valor").Value.ToString)

                                Case "DiasValidarProcesoLote"
                                    Functions.VariablesGlobales._DiasValidarProcesoLote = IIf(String.IsNullOrEmpty(RsFe.Fields.Item("U_Valor").Value.ToString), "0", RsFe.Fields.Item("U_Valor").Value.ToString)

                                Case "3Fechas"
                                    Functions.VariablesGlobales._PL3FECHAS = IIf(String.IsNullOrEmpty(RsFe.Fields.Item("U_Valor").Value.ToString), "N", RsFe.Fields.Item("U_Valor").Value.ToString)

                                Case "FechaContabilizacion"
                                    Functions.VariablesGlobales._PL1FECHAS = IIf(String.IsNullOrEmpty(RsFe.Fields.Item("U_Valor").Value.ToString), "N", RsFe.Fields.Item("U_Valor").Value.ToString)

                                Case "FechaFinMesAnteriorPL"
                                    Functions.VariablesGlobales.FechaFinMesAnteriorPL = RsFe.Fields.Item("U_Valor").Value.ToString

                            End Select
                            RsFe.MoveNext()
                        End While



                        '-------------------------Localizacion-----

                        If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then



                            Query = ""

                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                                Query += "FROM ""@SS_CONFD"" A INNER JOIN "
                                Query += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                                Query += " WHERE  B.""U_Modulo"" = '" & NombreAddonLOC & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                                Query += " AND B.""U_Subtipo"" = 'BD'"
                            Else
                                Query = "SELECT A.U_Nombre,A.U_Valor "
                                Query += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                                Query += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                                Query += " WHERE  B.U_Modulo = '" & NombreAddonLOC & "' AND  B.U_Tipo = 'PARAMETROS' "
                                Query += " AND  B.U_Subtipo = 'BD'"
                            End If


                            RsFe.DoQuery(Query)

                            While RsFe.EoF = False

                                Select Case RsFe.Fields.Item("U_Nombre").Value.ToString

                                    Case "ComprasQRY"
                                        Functions.VariablesGlobales._SS_ComprasQRY = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "VentasQRY"
                                        Functions.VariablesGlobales._SS_VentasQRY = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "CalculoFolioQRY"
                                        Functions.VariablesGlobales._SS_CalculoFolioQRY = RsFe.Fields.Item("U_Valor").Value.ToString


                                End Select

                                RsFe.MoveNext()
                            End While

                            'Se cargan los Bytes de los RPT
                            Query = ""

                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                                Query += "FROM ""@SS_CONFD"" A INNER JOIN "
                                Query += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                                Query += " WHERE  B.""U_Modulo"" = '" & NombreAddonLOC & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                                Query += " AND B.""U_Subtipo"" = 'NORPT'"
                            Else
                                Query = "SELECT A.U_Nombre,A.U_Valor "
                                Query += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                                Query += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                                Query += " WHERE  B.U_Modulo = '" & NombreAddonLOC & "' AND  B.U_Tipo = 'PARAMETROS' "
                                Query += " AND  B.U_Subtipo = 'NORPT'"
                            End If


                            RsFe.DoQuery(Query)

                            While RsFe.EoF = False

                                Select Case RsFe.Fields.Item("U_Nombre").Value.ToString

                                    Case "RPTFacturasCompras"
                                        Functions.VariablesGlobales._SS_RPTFacturasCompras = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "RPTFormularioSRI103"
                                        Functions.VariablesGlobales._SS_RPTFormularioSRI103 = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "RPTListadoVentas"
                                        Functions.VariablesGlobales._SS_RPTListadoVentas = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "RPTRetencionClientes"
                                        Functions.VariablesGlobales._SS_RPTRetencionClientes = RsFe.Fields.Item("U_Valor").Value.ToString

                                End Select

                                RsFe.MoveNext()
                            End While


                            'Se cargaran los ajustes Globales

                            Query = ""

                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                Query = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                                Query += "FROM ""@SS_CONFD"" A INNER JOIN "
                                Query += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                                Query += " WHERE  B.""U_Modulo"" = '" & NombreAddonLOC & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                                Query += " AND B.""U_Subtipo"" = 'CONFIGURACION'"
                            Else
                                Query = "SELECT A.U_Nombre,A.U_Valor "
                                Query += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                                Query += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                                Query += " WHERE  B.U_Modulo = '" & NombreAddonLOC & "' AND  B.U_Tipo = 'PARAMETROS' "
                                Query += " AND  B.U_Subtipo = 'CONFIGURACION'"
                            End If


                            RsFe.DoQuery(Query)

                            While RsFe.EoF = False

                                Select Case RsFe.Fields.Item("U_Nombre").Value.ToString

                                    Case "IpServer"
                                        Functions.VariablesGlobales._ipServer = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "UserDB"
                                        Functions.VariablesGlobales._gUsuarioDB = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "PassDB"
                                        Functions.VariablesGlobales._gPasswordDB = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "RutaRPT"
                                        Functions.VariablesGlobales._RutaRPT = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "AbrirRPTCargadosDB"
                                        Functions.VariablesGlobales._SS_AbrirRPTCargadosDB = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "Dinardap"
                                        Functions.VariablesGlobales._SS_DINARDAP = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "GenerarFolio"
                                        Functions.VariablesGlobales._SS_GenerarFolio = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "ValidarSocioNegociosUDF"
                                        Functions.VariablesGlobales._SS_ValidarSocioNegociosUDF = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "ValidarDocumentosUDF"
                                        Functions.VariablesGlobales._SS_ValidarDocumentosUDF = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "ValidarDocumentosUDF"
                                        Functions.VariablesGlobales._SS_ValidarDocumentosUDF = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "IDSerie"
                                        Functions.VariablesGlobales._SS_IDSerie = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "ImpuestoMontoCheque"
                                        Functions.VariablesGlobales._SS_ImpuestoMontoCheque = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "NombreImpuestoMontoCheque"
                                        Functions.VariablesGlobales._SS_NombreImpuestoMontoCheque = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "impuestoMontoProtesto"
                                        Functions.VariablesGlobales._SS_impuestoMontoProtesto = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "NombreImpuestoMontoProtesto"
                                        Functions.VariablesGlobales._SS_NombreImpuestoMontoProtesto = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "ActivarOpcionesParaManejoChequesProtestos"
                                        Functions.VariablesGlobales._SS_ActivarOpcionesParaManejoChequesProtestos = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "ActivarOpcionesParaTrasnferStockEntreBDs"
                                        Functions.VariablesGlobales._SS_ActivarOpcionesParaTrasnferStockEntreBDs = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "CuentaTransferencia"
                                        Functions.VariablesGlobales._SS_CuentaTransferencia = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "UserBaseDatos"
                                        Functions.VariablesGlobales._SS_UserDB = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "PassBaseDatos"
                                        Functions.VariablesGlobales._SS_PassDB = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "UserSAP"
                                        Functions.VariablesGlobales._SS_UserSAP = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "PassSAP"
                                        Functions.VariablesGlobales._SS_PassSAP = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "RutaRporteND"
                                        Functions.VariablesGlobales._SS_RutaReporteND = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "SetearCamposUsuario"
                                        Functions.VariablesGlobales._SS_SetearCamposUsuario = RsFe.Fields.Item("U_Valor").Value.ToString

                                        'Arturo add 14072023
                                    Case "MostrarBotonImpresion"
                                        Functions.VariablesGlobales._SS_MostrarBotonImpresion = RsFe.Fields.Item("U_Valor").Value.ToString


                                    Case "SConexionHana"
                                        Functions.VariablesGlobales._SS_SConexionHana = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "DriverHana"
                                        Functions.VariablesGlobales._SS_DriverHana = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "ParametroRPT"
                                        Functions.VariablesGlobales._SS_ParametroRPT = RsFe.Fields.Item("U_Valor").Value.ToString

                                    'items de los tabs paramtros
                                    Case "PosicionItemTabX"
                                        Functions.VariablesGlobales._PosicionItemTabX = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "PosicionItemTabY"
                                        Functions.VariablesGlobales._PosicionItemTabY = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "GuiasDesatendidas"
                                        Functions.VariablesGlobales._GuiasDesatendidas = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "QryGuiasDesSerie"
                                        Functions.VariablesGlobales._QueryGuiasDesatendidasSeries = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "QryGuiasDesNumDoc"
                                        Functions.VariablesGlobales._QueryGuiasDesatendidasProximoDocNum = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "PagosMasivos"
                                        Functions.VariablesGlobales._PagosMasivos = RsFe.Fields.Item("U_Valor").Value.ToString

                                        'Add JP 23/08/2024
                                    Case "RutaArchivoRPTPM"
                                        Functions.VariablesGlobales._RutaArchivoRPTPM = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "CadenaConexionRPTPM"
                                        Functions.VariablesGlobales._CadenaConexionRPTPM = RsFe.Fields.Item("U_Valor").Value.ToString

                                        'ADD JP 25/10/2024
                                    Case "CuentaTransitoriaPM"
                                        Functions.VariablesGlobales._CuentaTransitoriaPM = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "ManejoCuenta"
                                        Functions.VariablesGlobales._ManejoCuenta = RsFe.Fields.Item("U_Valor").Value.ToString
                                    Case "RutaArchivoTxt"
                                        Functions.VariablesGlobales._RutaArchivoTxt = RsFe.Fields.Item("U_Valor").Value.ToString

                                        'Add JP 23/08/2024
                                    Case "SBMenuPadreRptPreInf"
                                        Functions.VariablesGlobales._SBMenuPadreRptPreInf = RsFe.Fields.Item("U_Valor").Value.ToString

                                          'Add JP 23/08/2024
                                    'Case "RutaArchivoCEPM"   'Se utilizara general para rpt pm
                                    '    Functions.VariablesGlobales._RutaArchivoCEPM = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "RutaArchivoCHQPM"
                                        Functions.VariablesGlobales._RutaArchivoCHQPM = RsFe.Fields.Item("U_Valor").Value.ToString

                                        'Add DM 6/12/2024
                                    Case "NombreColumnasAnexos"
                                        Functions.VariablesGlobales._NombreColumnbasAnexo = RsFe.Fields.Item("U_Valor").Value.ToString


                                    Case "RutaRepCM"
                                        Functions.VariablesGlobales._RutaReposCM = RsFe.Fields.Item("U_Valor").Value.ToString

                                    Case "ActivarServiciosBasicos"
                                        Functions.VariablesGlobales._ActServiciosBasicos = RsFe.Fields.Item("U_Valor").Value.ToString

                                End Select

                                RsFe.MoveNext()
                            End While


                            'agregado el 14072023 para temas de  impresion
                            If Functions.VariablesGlobales._SS_MostrarBotonImpresion = "Y" Then


                                RsFe.DoQuery("Select * from ""@SS_IMPRESIONES"" ")

                                While RsFe.EoF = False

                                    Functions.VariablesGlobales._SS_ConfImpresoras.Add(New Functions.DatosImpresora With {.IdReporte = RsFe.Fields.Item("U_SS_IdReporte").Value.ToString(),
                                                                                                                                    .ImpresoraCompartida = RsFe.Fields.Item("U_SS_IMPCOM").Value.ToString(),
                                                                                                                                    .ImpresoraIndividual = RsFe.Fields.Item("U_SS_IMPIND").Value.ToString(),
                                                                                                                                    .TipoDocumento = RsFe.Fields.Item("U_SS_TipDoc").Value.ToString(),
                                                                                                                                    .Usuario = RsFe.Fields.Item("U_SS_Usuario").Value.ToString()})
                                    RsFe.MoveNext()

                                End While
                            End If

                        End If


                        '----------------------- fin parametros Localizacion




                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error al Obtener Configuracion Inicial del Addon " & ex.Message, "SubMain")

                    Finally

                        oFuncionesB1.Release(RsFe)
                    End Try
                    'RECUPERO EL NOMBRE DEL PROVEEDOR DE SAP BO PARA 

                    rSboApp.StatusBar.SetText(NombreAddon + " Obteniendo Nombre del Proveedor de SAP BO, desde la tabla de configuración... ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'Nombre_Proveedor_SAP_BO = ofrmParametrosAddon.ConsultaParametro(NombreAddon, "PARAMETROS", "CONFIGURACION", "ProveedorSAP")
                    Nombre_Proveedor_SAP_BO = Functions.VariablesGlobales._ProveedorSAP
                    If Nombre_Proveedor_SAP_BO = "" Then
                        rSboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización de Proveedor SAP BO", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "SubMain")

                    ' Seteo los filtros
                    SetFiltros()

                    ' SE AGREGO ESTA REFERENCIA PARA USAR EL METODO GUARDAR LOG, LO USA EMISION Y RECEPCION
                    oFuncionesAddon = New Functions.FuncionesAddon(rCompany, rSboApp, True, True, NombreAddon)
                    'VALIDA LICENCIA MEDIANTE ARCHIVO TXT
                    ' TIENE LICENCIA
                    rSboApp.StatusBar.SetText(NombreAddon + " Validando Licencia... ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)



                    Dim wsSSLIC As New Entidades.wsSS_LICENCIA_SAP.Licencia


                    Dim Proxy_puerto As String = ""
                    Dim Proxy_IP As String = ""
                    Dim Proxy_Usuario As String = ""
                    Dim Proxy_Clave As String = ""
                    If Functions.VariablesGlobales._SALIDA_POR_PROXY = "Y" Then

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
                    '****************************
                    ActivarTLS()

                    Dim TimeOutEmision As String = Functions.VariablesGlobales._gTimeOut_Emision
                    If TimeOutEmision = "" Then
                        wsSSLIC.Timeout = 30000 ' 30 segundos
                    Else
                        wsSSLIC.Timeout = Integer.Parse(TimeOutEmision)
                    End If

                    Dim msgg As String = ""
                    Dim oLicencia As Licencia = Nothing
                    Dim RucCliente As String = ""
                    Dim listaParametros() As Entidades.wsSS_LICENCIA_SAP.ClsConfigValores = Nothing
                    Try
                        Dim param_ambiente As Integer = 0
                        Dim param_inhouse As Boolean
                        Dim respLIC As Entidades.wsSS_LICENCIA_SAP.CLsRespLic = Nothing

                        Dim URLWS As String = Functions.VariablesGlobales._WsLicencia  'ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "WsLicencia")
                        If Not URLWS = "" Then
                            wsSSLIC.Url = URLWS
                        End If

                        RucCliente = Functions.VariablesGlobales._RucCompañia ' ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "RucCompañia")
                        'If (ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "TipoWebServices")) <> "LOCAL" Then
                        If Functions.VariablesGlobales._TipoWS <> "LOCAL" Then
                            param_inhouse = False
                        Else
                            param_inhouse = True
                        End If

                        If String.IsNullOrEmpty(URLWS) Then
                            rSboApp.StatusBar.SetText("No se encontro parametrización de WS Control, verificar por favor!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If

                        If String.IsNullOrEmpty(RucCliente) Then
                            RucCliente = "000000000"
                        End If

                        'If Not String.IsNullOrEmpty(ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "TipoWsLicencia")) Then

                        If Not String.IsNullOrEmpty(Functions.VariablesGlobales._TipoWsLicencia) Then
                            'param_ambiente = IIf(ofrmParametrosAddon.ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "TipoWsLicencia") = "PRUEBAS", 1, 2)
                            param_ambiente = IIf(Functions.VariablesGlobales._TipoWsLicencia = "PRUEBAS", 1, 2)
                            Functions.VariablesGlobales._vgAmbiente = param_ambiente ' DUDA
                        End If

                        'If Not String.IsNullOrEmpty(ofrmParametrosAddonNot.ConsultaParametro(Funciones_SAP.VariablesGlobales._gNombreAddOn, "PARAMETROS", "CONFIGURACION", "Param_Inhouse")) Then
                        '    param_inhouse = IIf(ofrmParametrosAddonNot.ConsultaParametro(Funciones_SAP.VariablesGlobales._gNombreAddOn, "PARAMETROS", "CONFIGURACION", "Param_Inhouse") = 1, True, False)
                        'End If

                        Dim SApdatos As New Entidades.wsSS_LICENCIA_SAP.DatosSap
                        With SApdatos
                            .RucEmpresa = RucCliente
                            .DireccionIPSERVER = rCompany.Server
                            .NombreDB = rCompany.CompanyDB
                            .NombreProducto = NombreAddon
                            .VersionProducto = rEstructura.VersionAddon
                            .Ambiente = param_ambiente
                            .Inhouse = param_inhouse
                            'Funciones_SAP.VariablesGlobales._gIsInHouse = param_inhouse
                        End With

                        'If Functions.VariablesGlobales._vgHttps = "Y" Then
                        '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                        'End If

                        SetProtocolosdeSeguridad()

                        'Descomentar antes de instalar en cliente
                        respLIC = wsSSLIC.ValidarLicencia(SApdatos, msgg)

                        If Not IsNothing(respLIC) Then
                            oLicencia = New Licencia
                            oLicencia.Opcion = respLIC.TipoLic

                            Functions.VariablesGlobales._vgTipoLicenciaAddOn = respLIC.TipoLic

                            oLicencia.NombreBaseSAP = rCompany.CompanyDB

                            oLicencia.Estado = CBool(respLIC.Estado)
                            Functions.VariablesGlobales._vgTieneLicenciaActivaAddOn = CBool(respLIC.Estado)

                            oLicencia.validoHasta = 1000
                            Functions.VariablesGlobales._vgCorreoResponsable = respLIC.MailResponsable
                            Functions.VariablesGlobales._vgVersionSRI = respLIC.VersionTributaria
                            listaParametros = respLIC.ListaUrlWS

                            Functions.VariablesGlobales._vgVersionDisponibleAddOn = respLIC.VersionProducto
                            Functions.VariablesGlobales._vgReleaseNoteAddOn = respLIC.ReleaseNoteProducto

                        End If
                    Catch ex As Exception
                        msgg = "GS Error de autenticacion con el WS " & ex.Message
                        oLicencia = Nothing
                    End Try


                    'esto es para probar offline Quitar Luego

                    'oLicencia = New Licencia

                    'oLicencia.Opcion = "full"
                    'oLicencia.Estado = True
                    'oLicencia.NombreBaseSAP = rCompany.CompanyDB
                    'Functions.VariablesGlobales._vgTieneLicenciaActivaAddOn = True

                    'fin offline--------------------------------


                    If oLicencia Is Nothing Then
                        Functions.VariablesGlobales._vgCorreoResponsable = ""
                        Functions.VariablesGlobales._vgTipoLicenciaAddOn = ""
                        Functions.VariablesGlobales._vgTieneLicenciaActivaAddOn = False

                        ''Menu_SinLicencia()
                        ''Menu_ParaPropuestaComercial()
                        'ofrmValidarUsuario = New frmValidarUsuario(rCompany, rSboApp)
                        'Functions.VariablesGlobales._vgTipoLicenciaAddOn = ""
                        rSboApp.StatusBar.SetText(NombreAddon + " - No se inició el AddOn " + NombreAddon + ", Se esta presentando un Inconveniente, contactese con un asesor de SolSap S.A. ! - " + msgg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Menu_SinLicencia()

                    Else
                        Functions.VariablesGlobales._gTimeOut_Emision = "30000"

                        rSboApp.StatusBar.SetText(NombreAddon + " Inicializando Variables del AddOn..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If oLicencia.Estado = False Then
                            Menu_SinLicencia()
                            'Menu_ParaPropuestaComercial()
                            textoMensajeSinLicencia = NombreAddon + " Su licencia se encuentra Inactiva , contactese con un asesor de SolSap360 S.A. !"
                        ElseIf Not oLicencia.NombreBaseSAP = rCompany.CompanyDB Then
                            textoMensajeSinLicencia = "La licencia para el addon " + NombreAddon + " no corresponde a la Sociedad que ha iniciado sesión!, Contactese con un asesor de SolSap S.A. !"
                            Menu_SinLicencia()

                        ElseIf oLicencia.Opcion.ToLower = "emision" Then
                            rSboApp.StatusBar.SetText(NombreAddon + " Licencia valida para el manejo de Emisión de Documentos ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Menu_LicenciaEMISION()
                            rEventoEmision = New EventosEmision()
                            ofrmLogEmision = New frmLogEmision(rCompany, rSboApp)
                            If Functions.VariablesGlobales._vgImpBlo = "Y" Then
                                ofrmImpresionPorBloque = New frmImpresionPorBloque(rCompany, rSboApp)
                            End If
                            OfrmListaAEnviar = New frmListaAEnviar(rCompany, rSboApp)
                            ofrmValidarUsuario = New frmValidarUsuario(rCompany, rSboApp)


                            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                                InstanciarObjetosLocalizacionEC()

                            End If

                            'ADD 09072024
                            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then IntanciarObjetosCM()

                        ElseIf oLicencia.Opcion.ToLower = "recepcion" Then
                            rSboApp.StatusBar.SetText(NombreAddon + " Licencia valida para el manejo de Recepción de Documentos ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Menu_LicenciaRECEPCION()
                            rEventoRecepion = New EventosRecepcion()

                            ofrmMapeo = New frmMapeo(rCompany, rSboApp)

                            ofrmConsultaOrdenes = New frmConsultaOrdenes(rCompany, rSboApp)
                            ofrmDocumentosIntegrados = New frmDocumentosIntegrados(rCompany, rSboApp)
                            ofrmParametrosRecepcion = New frmParametrosRecepcion(rCompany, rSboApp)

                            ofrmSubirArchivo = New frmSubirArchivo(rCompany, rSboApp)


                            If Functions.VariablesGlobales._vgPreLot = "Y" Then
                                ofrmProcesoLote = New frmProcesoLote(rCompany, rSboApp)
                                'ofrmProcesoLote2 = New frmProcesoLote2(rCompany, rSboApp)
                            End If
                            If Functions.VariablesGlobales._vgProcesoLoteManamer = "Y" Then
                                ofrmProcesoLoteManamer = New frmProcesoLoteManamer(rCompany, rSboApp)
                            End If

                            ofrmValidarUsuario = New frmValidarUsuario(rCompany, rSboApp)
                            OfrmListaAEnviar = New frmListaAEnviar(rCompany, rSboApp)
                            ofrmProcesoLoteC = New frmProcesoLoteC(rCompany, rSboApp)

                            If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                ofrmDocumentosRecibidosXML = New frmDocumentosRecibidosXML(rCompany, rSboApp)
                                ofrmDocumentoXML = New frmDocumentoXML(rCompany, rSboApp)
                                ofrmDocumentoNCXML = New frmDocumentoNCXML(rCompany, rSboApp)
                                ofrmDocumentoREXML = New frmDocumentoREXML(rCompany, rSboApp)

                                If Functions.VariablesGlobales._PreliminarLoteXML = "Y" Then
                                    ofrmProcesoLoteXML = New frmProcesoLoteXML(rCompany, rSboApp)
                                End If

                            Else
                                ofrmDocumentosRecibidos = New frmDocumentosRecibidos(rCompany, rSboApp)
                                ofrmDocumento = New frmDocumento(rCompany, rSboApp)
                                ofrmDocumentoNC = New frmDocumentoNC(rCompany, rSboApp)
                                ofrmDocumentoRE = New frmDocumentoRE(rCompany, rSboApp)
                            End If

                            'Si esta activado el param se inicia
                            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                                InstanciarObjetosLocalizacionEC()

                            End If

                            'ADD 09072024
                            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then IntanciarObjetosCM()

                        ElseIf oLicencia.Opcion.ToLower = "full" Then
                            rSboApp.StatusBar.SetText(NombreAddon + " Licencia FULL valida para el manejo de Emisión de Documentos y Recepción de Documentos ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                            Menu_LicenciaFULL()
                            rEventoEmision = New EventosEmision()
                            ofrmLogEmision = New frmLogEmision(rCompany, rSboApp)

                            rEventoRecepion = New EventosRecepcion()

                            ofrmMapeo = New frmMapeo(rCompany, rSboApp)

                            ofrmConsultaOrdenes = New frmConsultaOrdenes(rCompany, rSboApp)
                            ofrmDocumentosIntegrados = New frmDocumentosIntegrados(rCompany, rSboApp)
                            ofrmParametrosRecepcion = New frmParametrosRecepcion(rCompany, rSboApp)

                            ofrmSubirArchivo = New frmSubirArchivo(rCompany, rSboApp)

                            If Functions.VariablesGlobales._vgPreLot = "Y" Then
                                ofrmProcesoLote = New frmProcesoLote(rCompany, rSboApp)
                                'ofrmProcesoLote2 = New frmProcesoLote2(rCompany, rSboApp)
                            End If
                            If Functions.VariablesGlobales._vgProcesoLoteManamer = "Y" Then
                                ofrmProcesoLoteManamer = New frmProcesoLoteManamer(rCompany, rSboApp)
                            End If

                            ofrmValidarUsuario = New frmValidarUsuario(rCompany, rSboApp)
                            ofrmProcesoLoteC = New frmProcesoLoteC(rCompany, rSboApp)
                            OfrmListaAEnviar = New frmListaAEnviar(rCompany, rSboApp)
                            If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                ofrmDocumentosRecibidosXML = New frmDocumentosRecibidosXML(rCompany, rSboApp)
                                ofrmDocumentoXML = New frmDocumentoXML(rCompany, rSboApp)
                                ofrmDocumentoNCXML = New frmDocumentoNCXML(rCompany, rSboApp)
                                ofrmDocumentoREXML = New frmDocumentoREXML(rCompany, rSboApp)
                                ofrmProcesoLoteXML = New frmProcesoLoteXML(rCompany, rSboApp)
                            Else
                                ofrmDocumentosRecibidos = New frmDocumentosRecibidos(rCompany, rSboApp)
                                ofrmDocumento = New frmDocumento(rCompany, rSboApp)
                                ofrmDocumentoNC = New frmDocumentoNC(rCompany, rSboApp)
                                ofrmDocumentoRE = New frmDocumentoRE(rCompany, rSboApp)

                            End If

                            'If Functions.VariablesGlobales._vgImpBlo = "Y" Then
                            '    ofrmImpresionPorBloque = New frmImpresionPorBloque(rCompany, rSboApp)
                            'End If


                            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                                InstanciarObjetosLocalizacionEC()

                            End If

                            'ADD 09072024
                            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then IntanciarObjetosCM()

                            rEstructura.LicenciaAddon = "FULL"
                            rSboApp.StatusBar.SetText("Se Conecto el Add-On " & NombreAddon, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        End If
                    End If

                    ''FIN LICENCIA TXT
                    If Not textoMensajeSinLicencia = "" Then
                        rSboApp.MessageBox(textoMensajeSinLicencia)
                    End If

                    ''CargarMenuDesdeXML()

                    ''rSboApp.StatusBar.SetText("Validando Conexion ADO ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    ''oManejoDocumentos.ValidarConexionADO()

                    ''rSboApp.StatusBar.SetText("Se Conecto el Add-On " & NombreAddon, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Catch ex As Exception
                    System.Windows.Forms.MessageBox.Show("No se ha Conectado con AddOn " & NombreAddon & ": " & ex.Message)

                    End
                End Try
            End If

            System.Windows.Forms.Application.Run()

            Exit Sub

        Catch exMain As Exception
            System.Windows.Forms.MessageBox.Show("Error iniciando el Add-On " & NombreAddon & ": " & exMain.Message)
        End Try
    End Sub

    Private Sub InstanciarObjetosLocalizacionEC()

        ofrmClaveLE = New frmClave(rCompany, rSboApp)
        ofrmAcercaDeLE = New frmAcercaDeLE(rCompany, rSboApp)
        ofrmConfClaveLE = New frmConfClaveLE(rCompany, rSboApp)
        ofrmParametrosAddonLE = New frmParametrosAddonLE(rCompany, rSboApp)
        ofrmConfMenuLE = New frmConfMenuLE(rCompany, rSboApp)

        '--------------------------


        ofrmGeneradorATS = New frmGeneradorATS(rCompany, rSboApp)
        ofrmGenerarRPT = New frmGenerarRPT(rCompany, rSboApp)

        ofrmAnexoVentas = New frmAnexoVentas(rCompany, rSboApp)

        ofrmAnexoCompras = New frmAnexoCompras(rCompany, rSboApp)

        ofrmDinardap = New frmDinardap(rCompany, rSboApp)

        ofrmSRI = New frmSRI(rCompany, rSboApp)

        ofrmSRIConsulta = New frmSRIConsulta(rCompany, rSboApp)

        ofrmChequeP = New frmChequeP(rCompany, rSboApp)
        ofrmChequePD = New frmChequePD(rCompany, rSboApp)

        ofrmTransEntreCompanias = New frmTransEntreCompanias(rCompany, rSboApp)
        ofrmConsultaDetalleTrans = New frmConsultaDetalleTrans(rCompany, rSboApp)
        ofrmConsultaSalidaEntrada = New frmConsultaSalidaEntrada(rCompany, rSboApp)
        ofrmConsultaBodega = New frmConsultaBodega(rCompany, rSboApp)

        ofrmConsultasDbLE = New frmConsultasDbLE(rCompany, rSboApp)

        ofrmCargaRPT = New frmCargaRPT(rCompany, rSboApp)

        'add 14052024
        'guias y pagos masivos

        ofrmGuiasRemision = New frmGuiasRemision(rCompany, rSboApp)
        ofrmPagosMasivos = New frmPagosMasivos(rCompany, rSboApp)
        ofrmImprimir = New frmImprimir(rCompany, rSboApp)
        ofrmPagosAprobacion = New frmPagosAprobacion(rCompany, rSboApp)

        'add 02072024
        ofrmCashManagement = New frmCashManagemet(rCompany, rSboApp)
        ofrmMapeoCuentas = New frmMapeoCuentasCM(rCompany, rSboApp)

        'add 10092024 Servicios basicos
        ofrmServiciosBasicos = New frmServiciosBasicos(rCompany, rSboApp)

    End Sub

    Private Sub IntanciarObjetosCM()
        ofrmCashManagement = New frmCashManagemet(rCompany, rSboApp)
        ofrmMapeoCuentas = New frmMapeoCuentasCM(rCompany, rSboApp)
    End Sub

    'Private Sub SetFiltros_Localizacion(ByRef oFiltros As SAPbouiCOM.EventFilters)

    '    'oFiltros = New SAPbouiCOM.EventFilters()
    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
    '    oFiltro.AddEx("133") 'FACTURA DE CLIENTES
    '    oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
    '    oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
    '    oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
    '    oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
    '    oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
    '    oFiltro.AddEx("140") 'GUIA DE REMISION
    '    oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
    '    oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
    '    oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
    '    oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
    '    oFiltro.AddEx("181") 'NOTA DE CREDITO PROVEEDOR/RETENCION - PARA RECEPCION
    '    oFiltro.AddEx("170") 'PAGO RECIBIDO CLIENTE/RETENCION - PARA RECEPCION

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
    '    oFiltro.AddEx("133") 'FACTURA DE CLIENTES
    '    oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
    '    oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
    '    oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
    '    oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
    '    oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
    '    oFiltro.AddEx("140") 'GUIA DE REMISION
    '    oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
    '    oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
    '    oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
    '    oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
    '    oFiltro.AddEx("181") 'NOTA DE CREDITO PROVEEDOR/RETENCION - PARA RECEPCION

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
    '    oFiltro.AddEx("133") 'FACTURA DE CLIENTES
    '    oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
    '    oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
    '    oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
    '    oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
    '    oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
    '    oFiltro.AddEx("140") 'GUIA DE REMISION
    '    oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
    '    oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
    '    oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("frmAcercaDeLE")
    '    oFiltro.AddEx("frmAnexoCompras")
    '    oFiltro.AddEx("frmAnexoVentas")
    '    oFiltro.AddEx("frmCargaRPT")
    '    oFiltro.AddEx("frmChequeP")
    '    oFiltro.AddEx("frmChequePD")
    '    oFiltro.AddEx("frmClave")
    '    oFiltro.AddEx("frmConfClaveLE")
    '    oFiltro.AddEx("frmConfMenuLE")
    '    oFiltro.AddEx("frmDinardap")
    '    oFiltro.AddEx("frmGeneradorATS")
    '    oFiltro.AddEx("frmGenerarRPT")
    '    oFiltro.AddEx("frmParametrosAddonLE")
    '    oFiltro.AddEx("frmSRI")
    '    oFiltro.AddEx("frmSRIConsulta")
    '    oFiltro.AddEx("frmTransEntreCompanias")

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
    '    oFiltro.AddEx("133") 'FACTURA DE CLIENTES
    '    oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
    '    oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
    '    oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
    '    oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
    '    oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
    '    oFiltro.AddEx("140") 'GUIA DE REMISION
    '    oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
    '    oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
    '    oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
    '    oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
    '    oFiltro.AddEx("181") 'NOTA DE CREDITO PROVEEDOR/RETENCION - PARA RECEPCION

    '    oFiltro.AddEx("frmSRI") 'formulario del SRI

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
    '    oFiltro.AddEx("133") 'FACTURA DE CLIENTES
    '    oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
    '    oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
    '    oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
    '    oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
    '    oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
    '    oFiltro.AddEx("140") 'GUIA DE REMISION
    '    oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
    '    oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
    '    oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("-141") 'FACTURA DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
    '    oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
    '    oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
    '    oFiltro.AddEx("-60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
    '    oFiltro.AddEx("frmTransEntreCompanias")

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
    '    oFiltro.AddEx("frmChequeP")

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
    '    oFiltro.AddEx("frmAcercaDeLE")
    '    oFiltro.AddEx("frmAnexoCompras")
    '    oFiltro.AddEx("frmAnexoVentas")
    '    oFiltro.AddEx("frmCargaRPT")
    '    oFiltro.AddEx("frmChequeP")
    '    oFiltro.AddEx("frmChequePD")
    '    oFiltro.AddEx("frmClave")
    '    oFiltro.AddEx("frmConfClaveLE")
    '    oFiltro.AddEx("frmConfMenuLE")
    '    oFiltro.AddEx("frmDinardap")
    '    oFiltro.AddEx("frmGeneradorATS")
    '    oFiltro.AddEx("frmGenerarRPT")
    '    oFiltro.AddEx("frmParametrosAddonLE")
    '    oFiltro.AddEx("frmSRI")
    '    oFiltro.AddEx("frmSRIConsulta")
    '    oFiltro.AddEx("frmTransEntreCompanias")

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
    '    oFiltro.AddEx("frmChequePD")
    '    oFiltro.AddEx("frmTransEntreCompanias")

    '    oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
    '    oFiltro.AddEx("frmTransEntreCompanias")
    'End Sub

    Private Sub SetFiltros()

        Try
            ' Creo un nuevo objecto EventFilters
            oFiltros = New SAPbouiCOM.EventFilters()

            ' Agrego el tipo de evento al contenedor
            ' Este metodo retorna un objecto EventFilter
            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            'oFiltro.AddEx("65307") 'FACTURA DE exportacion
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            oFiltro.AddEx("181") 'NOTA DE CREDITO PROVEEDOR/RETENCION - PARA RECEPCION
            oFiltro.AddEx("170") 'PAGO RECIBIDO CLIENTE/RETENCION - PARA RECEPCION


            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                oFiltro.AddEx("frmGuiasRemision")
                'oFiltro.AddEx("frmPagosMasivos")

            End If

            '60092
            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            oFiltro.AddEx("170") 'PAGO RECIBIDO CLIENTE/RETENCION - PARA RECEPCION -  para controlar el evento cancelar
            'se agregan por nueva funcionalidad
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            'oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            '-------------------------------------
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                oFiltro.AddEx("UDO_FT_TM_RETV")
            End If
            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                ' oFiltro.AddEx("frmPagosMasivos")
            End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            ' oFiltro.AddEx("65303") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            oFiltro.AddEx("181") 'NOTA DE CREDITO PROVEEDOR/RETENCION - PARA RECEPCION
            'oFiltro.AddEx("170") 'PAGO RECIBIDO CLIENTE/RETENCION - PARA RECEPCION
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
               Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
               Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                oFiltro.AddEx("146") 'PAGO RECIBIDO CLIENTE/RETENCION RCT3 - PARA RECEPCION

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then

                oFiltro.AddEx("170") 'PAGO RECIBIDO CLIENTE/ORCT - PARA RECEPCION

            End If
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                oFiltro.AddEx("UDO_FT_TM_RETV")
            End If
            oFiltro.AddEx("85")
            oFiltro.AddEx("720")


            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                oFiltro.AddEx("frmGuiasRemision")
                'oFiltro.AddEx("frmPagosMasivos")

            End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            oFiltro.AddEx("65307") 'NOTA DE DEBITO DE CLIENTES
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("frmDocumento")
            oFiltro.AddEx("frmDocumentoNC")
            oFiltro.AddEx("frmDocumentoRE")
            oFiltro.AddEx("frmDocumentosEnviados")
            oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            oFiltro.AddEx("frmConfClave") ' FORMULARIO DE VALIDACION CLAVE PARA INGRESAR AL MODULO DE CONFIGURACION
            oFiltro.AddEx("frmParametrosAddon")
            oFiltro.AddEx("frmProxy")
            oFiltro.AddEx("frmDocumentoXML")
            oFiltro.AddEx("frmDocumentoNCXML")
            oFiltro.AddEx("frmDocumentoREXML")

            'proceso en lote

            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLote2")
            oFiltro.AddEx("frmProcesoLoteManamer")

            oFiltro.AddEx("frmImpresionPorBloque")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")
            oFiltro.AddEx("frmListaAEnviar")
            'impresion por bloque dibeal

            'validar usuario
            oFiltro.AddEx("frmValidarUsuario")

            oFiltro.AddEx("181") 'NOTA DE CREDITO PROVEEDOR/RETENCION - PARA RECEPCION
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
               Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
               Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
               Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                oFiltro.AddEx("146") 'PAGO RECIBIDO CLIENTE/RETENCION RCT3 - PARA RECEPCION

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then

                oFiltro.AddEx("170") 'PAGO RECIBIDO CLIENTE/ORCT - PARA RECEPCION

            End If
            oFiltro.AddEx("frmSubirArchivo")
            oFiltro.AddEx("85")
            oFiltro.AddEx("720")
            'Artur 10042024
            oFiltro.AddEx("frmConsultasDB")
            oFiltro.AddEx("frmConsultasDB_RE")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                oFiltro.AddEx("frmAcercaDeLE")
                oFiltro.AddEx("frmAnexoCompras")
                oFiltro.AddEx("frmAnexoVentas")
                oFiltro.AddEx("frmCargaRPT")
                oFiltro.AddEx("frmChequeP")
                oFiltro.AddEx("frmChequePD")
                oFiltro.AddEx("frmClave")
                oFiltro.AddEx("frmConfClaveLE")
                oFiltro.AddEx("frmConfMenuLE")
                oFiltro.AddEx("frmDinardap")
                oFiltro.AddEx("frmGeneradorATS")
                oFiltro.AddEx("frmGenerarRPT")
                oFiltro.AddEx("frmParametrosAddonLE")
                oFiltro.AddEx("frmSRI")
                oFiltro.AddEx("frmSRIConsulta")
                oFiltro.AddEx("frmTransEntreCompanias")
                'frmConsultasDbLE
                oFiltro.AddEx("frmConsultasDbLE")

                oFiltro.AddEx("frmGuiasRemision")

                oFiltro.AddEx("frmPagosMasivos")
                oFiltro.AddEx("frmPagosAprobacion")
                oFiltro.AddEx("frmMapeoCuentasCM")
                oFiltro.AddEx("frmCashManagement")
                oFiltro.AddEx("frmServiciosBasicos")
                oFiltro.AddEx("385")
            End If

            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then
                'CashManagement 02072024
                oFiltro.AddEx("frmCashManagement")
                oFiltro.AddEx("frmMapeoCuentasCM")
            End If


            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmDocumento")
            oFiltro.AddEx("frmMapeo")
            oFiltro.AddEx("frmConsultaOrdenes")
            oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            'oFiltro.AddEx("frmSubirArchivo")

            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLote2")
            oFiltro.AddEx("frmProcesoLoteManamer")

            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmDocumentoXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")

            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                oFiltro.AddEx("UDO_FT_TM_RETV")
            End If

            'oFiltro.AddEx("frmImpresionPorBloque")
            oFiltro.AddEx("85")

            oFiltro.AddEx("frmSRI") 'formulario del SRI

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmGuiasRemision")
                oFiltro.AddEx("frmServiciosBasicos")
                oFiltro.AddEx("385")
            End If


            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("-141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            oFiltro.AddEx("-60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmDocumentosIntegrados")

            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLote2")
            oFiltro.AddEx("frmProcesoLoteManamer")

            oFiltro.AddEx("frmParametrosAddon")

            oFiltro.AddEx("frmImpresionPorBloque")

            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmDocumentosEnviados")
            oFiltro.AddEx("frmProcesoLoteC")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                oFiltro.AddEx("frmTransEntreCompanias")
                oFiltro.AddEx("frmGuiasRemision")

                'oFiltro.AddEx("frmPagosMasivos")
                'oFiltro.AddEx("frmCashManagement")
                'oFiltro.AddEx("frmMapeoCuentasCM")
            End If


            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmDocumentosIntegrados")
            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmProcesoLoteC")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                oFiltro.AddEx("frmChequeP")

                oFiltro.AddEx("frmSRI")
                oFiltro.AddEx("frmPagosAprobacion")
                'oFiltro.AddEx("frmPagosMasivos")
            End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmDocumento")
            oFiltro.AddEx("frmMapeo")
            oFiltro.AddEx("frmConsultaOrdenes")
            oFiltro.AddEx("frmDocumentosIntegrados")
            oFiltro.AddEx("frmParametrosRecepcion")
            oFiltro.AddEx("frmAcercaDe")
            oFiltro.AddEx("frmConfMenu")
            oFiltro.AddEx("frmDocumentoNC")
            oFiltro.AddEx("frmDocumentoRE")
            'oFiltro.AddEx("frmImpresionPorBloque")

            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLote2")
            oFiltro.AddEx("frmProcesoLoteManamer")
            oFiltro.AddEx("169") ' MAIN MENU
            oFiltro.AddEx("frmValidarUsuario")

            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmDocumentoXML")
            oFiltro.AddEx("frmDocumentoNCXML")
            oFiltro.AddEx("frmDocumentoREXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                '----localizacion
                oFiltro.AddEx("frmAcercaDeLE")
                oFiltro.AddEx("frmAnexoCompras")
                oFiltro.AddEx("frmAnexoVentas")
                oFiltro.AddEx("frmCargaRPT")
                oFiltro.AddEx("frmChequeP")
                oFiltro.AddEx("frmChequePD")
                oFiltro.AddEx("frmClave")
                oFiltro.AddEx("frmConfClaveLE")
                oFiltro.AddEx("frmConfMenuLE")
                oFiltro.AddEx("frmDinardap")
                oFiltro.AddEx("frmGeneradorATS")
                oFiltro.AddEx("frmGenerarRPT")
                oFiltro.AddEx("frmParametrosAddonLE")
                oFiltro.AddEx("frmSRI")
                oFiltro.AddEx("frmSRIConsulta")
                oFiltro.AddEx("frmTransEntreCompanias")
                'guias de remision
                oFiltro.AddEx("frmGuiasRemision")

                'pagos masivos add 30052024
                oFiltro.AddEx("frmPagosMasivos")
                oFiltro.AddEx("frmPagosAprobacion")
                oFiltro.AddEx("frmImprimir")
                '----fin
                'CashManagement 02072024
                oFiltro.AddEx("frmCashManagement")
                oFiltro.AddEx("frmMapeoCuentasCM")

                'ADD 10/09/2024
                oFiltro.AddEx("frmServiciosBasicos")
            End If

            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then
                'CashManagement 02072024
                oFiltro.AddEx("frmCashManagement")
                oFiltro.AddEx("frmMapeoCuentasCM")
            End If


            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")
            oFiltro.AddEx("frmPagosMasivos")

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmDocumentosIntegrados")
            oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            oFiltro.AddEx("65307") 'FACTURA DE EXPORTACION
            oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            oFiltro.AddEx("140") 'GUIA DE REMISION
            oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION

            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLoteManamer")
            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("85")
            oFiltro.AddEx("720")
            oFiltro.AddEx("frmProcesoLoteC")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmTransEntreCompanias")
                oFiltro.AddEx("frmSRI")
                oFiltro.AddEx("frmSRIConsulta")
                oFiltro.AddEx("frmGuiasRemision")

                oFiltro.AddEx("frmPagosMasivos")
            End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmMapeo")
            oFiltro.AddEx("frmDocumentosIntegrados")
            oFiltro.AddEx("frmParametrosRecepcion")
            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLoteManamer")
            oFiltro.AddEx("frmDocumentosEnviados")
            oFiltro.AddEx("frmImpresionPorBloque")
            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmChequePD")
                oFiltro.AddEx("frmTransEntreCompanias")
                oFiltro.AddEx("frmGuiasRemision")
                oFiltro.AddEx("frmPagosMasivos")

                oFiltro.AddEx("frmAnexoVentas")
                oFiltro.AddEx("frmAnexoCompras")

                oFiltro.AddEx("133")
                oFiltro.AddEx("141")
            End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
            oFiltro.AddEx("frmDocumentosRecibidos")
            oFiltro.AddEx("frmDocumentosEnviados")
            oFiltro.AddEx("frmDocumentosIntegrados")

            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLoteManamer")
            oFiltro.AddEx("frmImpresionPorBloque")
            oFiltro.AddEx("frmDocumentosRecibidosXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")
            oFiltro.AddEx("frmListaAEnviar")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmAnexoCompras")
                oFiltro.AddEx("frmAnexoVentas")
                oFiltro.AddEx("frmPagosMasivos")
            End If

            'et_FORM_RESIZE
            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
            oFiltro.AddEx("frmDocumento")
            oFiltro.AddEx("frmDocumentoNC")
            oFiltro.AddEx("frmDocumentoRE")
            oFiltro.AddEx("frmProcesoLote")
            oFiltro.AddEx("frmProcesoLoteManamer")
            oFiltro.AddEx("frmDocumentoXML")
            oFiltro.AddEx("frmDocumentoNCXML")
            oFiltro.AddEx("frmDocumentoREXML")
            oFiltro.AddEx("frmProcesoLoteXML")
            oFiltro.AddEx("frmProcesoLoteC")

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmPagosMasivos")
                oFiltro.AddEx("frmServiciosBasicos")
            End If


            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            oFiltro.AddEx("frmGuiasRemision")
            'If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
            '    oFiltro.AddEx("frmPagosMasivos")
            'End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmServiciosBasicos")
            End If

            oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)
            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                oFiltro.AddEx("frmPagosMasivos")
                oFiltro.AddEx("frmServiciosBasicos")
            End If

            'FoSboForm.EnableMenu('1281', false); //'find record
            'FoSboForm.EnableMenu('1282', false); //'add new record
            'FoSboForm.EnableMenu('1283',false); //Delete record
            'FoSboForm.EnableMenu('1288', false); //'next record
            'FoSboForm.EnableMenu('1289', false); //'previous record
            'FoSboForm.EnableMenu('1290', false); //'first record
            'FoSboForm.EnableMenu('1291', false); //'last record

            ' Seteo la aplicacion con los filtros
            'oFiltro = oFiltros.Add(SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            'oFiltro.AddEx("133") 'FACTURA DE CLIENTES
            'oFiltro.AddEx("60090") 'FACTURA DE DEUDOR + PAGO
            'oFiltro.AddEx("60091") 'FACTURA DE RESERVA CLIENTES
            'oFiltro.AddEx("65300") 'FACTURA DE ANTICIPO DE CLIENTES
            'oFiltro.AddEx("179") 'NOTA DE CREDITO DE CLIENTES
            'oFiltro.AddEx("65303") 'NOTA DE DEBITO DE CLIENTES
            'oFiltro.AddEx("140") 'GUIA DE REMISION
            'oFiltro.AddEx("940") 'GUIA DE REMISION - TRANSFERENCIAS
            'oFiltro.AddEx("1250000940") 'GUIA DE REMISION -  SOLICITUD TRANSLADO
            'oFiltro.AddEx("141") 'FACTURA DE PROVEEDOR/RETENCION
            'oFiltro.AddEx("65306") 'NOTA DE DEBITO DE PROVEEDOR/RETENCION
            'oFiltro.AddEx("frmDocumentosRecibidos")
            'oFiltro.AddEx("frmDocumento")
            'oFiltro.AddEx("frmMapeo")
            'oFiltro.AddEx("frmConsultaOrdenes")
            'oFiltro.AddEx("65301") 'FACTURA DE ANTICIPO PROVEEDOR/RETENCION
            'oFiltro.AddEx("60092") 'FACTURA DE RESERVA PROVEEDOR/RETENCION

            '-----------------filtros Localizacion
            '  SetFiltros_Localizacion(oFiltros)
            '--------------------fin Filtros LOC
            'oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
            'oFiltro.AddEx("frmPagosMasivos")

            rSboApp.SetFilter(oFiltros)

        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al setear los filtros: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    Private Sub Menu_LicenciaFULL()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipal"
            oCreationPackage.String = "SAED"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logo11.png"
            oCreationPackage.Position = 15
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipal")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "GS10"
            oCreationPackage.String = "Emisión de Documentos Electrónicos"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "GS20"
            oCreationPackage.String = "Recepción de Documentos Electrónicos"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'ADD 09072024
            'If Functions.VariablesGlobales._ActivarCMFML = "Y" Then Menu_CM(oCreationPackage, oMenus, oMenuItem)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS30"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("GS10")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS11"
            oCreationPackage.String = "Documentos Enviados"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'impresion por bloque dibeal
            If Functions.VariablesGlobales._vgImpBlo = "Y" Then
                oMenuItem = rSboApp.Menus.Item("GS10")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GS12"
                oCreationPackage.String = "Generar PDF por Bloque"
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If

            oMenuItem = rSboApp.Menus.Item("GS20")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS21"
            oCreationPackage.String = "Documentos Recibidos"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS22"
            oCreationPackage.String = "Documentos Integrados"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'MENU PARA PROCESO EN LOTE
            If Functions.VariablesGlobales._vgPreLot = "Y" Or Functions.VariablesGlobales._PreliminarLoteXML = "Y" Then
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GS24"
                oCreationPackage.String = "Proceso de Documentos en Lote"
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If

            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS23"
            oCreationPackage.String = "Configuración de Parametros"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try




            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                Menu_Localizacion(oCreationPackage, oMenus, oMenuItem)

            End If



        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu Licencia FULL" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub Menu_SinLicencia()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipal"
            oCreationPackage.String = "SAED"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logo11.png"
            oCreationPackage.Position = 15
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


            oMenuItem = rSboApp.Menus.Item("mnPrincipal")
            oMenus = oMenuItem.SubMenus
         
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS30"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu SIN Licencia" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub Menu_LicenciaEMISION()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipal"
            oCreationPackage.String = "SAED"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logo11.png"
            oCreationPackage.Position = 15
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipal")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "GS10"
            oCreationPackage.String = "Emisión de Documentos Electrónicos"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try
           
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS30"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("GS10")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS11"
            oCreationPackage.String = "Documentos Enviados"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'impresion por bloque dibeal
            If Functions.VariablesGlobales._vgImpBlo = "Y" Then
                oMenuItem = rSboApp.Menus.Item("GS10")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GS12"
                oCreationPackage.String = "Generar PDF por Bloque"
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If


            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                Menu_Localizacion(oCreationPackage, oMenus, oMenuItem)

            End If

            'ADD 09072024
            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then Menu_CM(oCreationPackage, oMenus, oMenuItem)

        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu Licencia EMISION" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub Menu_LicenciaRECEPCION()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipal"
            oCreationPackage.String = "SAED"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logo11.png"
            oCreationPackage.Position = 15
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipal")
            oMenus = oMenuItem.SubMenus
           
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "GS20"
            oCreationPackage.String = "Recepción de Documentos Electrónicos"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


            oMenuItem = rSboApp.Menus.Item("GS20")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS21"
            oCreationPackage.String = "Documentos Recibidos"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS22"
            oCreationPackage.String = "Documentos Integrados"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS23"
            oCreationPackage.String = "Configuración de Parametros"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try



            'MENU PARA PROCESO EN LOTE
            If Functions.VariablesGlobales._vgPreLot = "Y" Or Functions.VariablesGlobales._PreliminarLoteXML = "Y" Then
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GS24"
                oCreationPackage.String = "Proceso de Documentos en Lote"
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If

            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS23"
            oCreationPackage.String = "Configuración de Parametros"
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'oMenus = oMenuItem.SubMenus
            'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            'oCreationPackage.UniqueID = "GS24"
            'oCreationPackage.String = "Cargar de Documentos - TXT"
            'Try
            '    oMenus.AddEx(oCreationPackage)
            'Catch ex As Exception
            'End Try

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                Menu_Localizacion(oCreationPackage, oMenus, oMenuItem)

            End If

            'ADD 09072024
            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then Menu_CM(oCreationPackage, oMenus, oMenuItem)

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS30"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu Licencia RECEPCION" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub


    Private Sub Menu_Localizacion(oCreationPackage As SAPbouiCOM.MenuCreationParams, oMenus As SAPbouiCOM.Menus, oMenuItem As SAPbouiCOM.MenuItem)

        Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipal")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipalLoc"
            oCreationPackage.String = "Localización Ecuador"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GSLoc5"
            oCreationPackage.String = "Parametrización SRI"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "GSLoc10"
            oCreationPackage.String = "Generador de ATS"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            If Functions.VariablesGlobales._SS_AbrirRPTCargadosDB = "Y" Then

                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc20"
                oCreationPackage.String = "Informes Legales"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

            Else

                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "GSLoc20_1"
                oCreationPackage.String = "Informes Legales"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try


                oMenuItem = rSboApp.Menus.Item("GSLoc20_1")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc20_2"
                oCreationPackage.String = "Informe facturas de compras"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try


                oMenuItem = rSboApp.Menus.Item("GSLoc20_1")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc20_3"
                oCreationPackage.String = "Informe para Formulario SRI 103"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try


                oMenuItem = rSboApp.Menus.Item("GSLoc20_1")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc20_4"
                oCreationPackage.String = "Informe Listado de Ventas"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try


                oMenuItem = rSboApp.Menus.Item("GSLoc20_1")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc20_5"
                oCreationPackage.String = "Informe Retenciones de clientes"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc20_1")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc20_6"
                oCreationPackage.String = "Informe 104"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                If Functions.VariablesGlobales._SS_AbrirRPTCargadosDB = "Y" Then
                    oMenuItem = rSboApp.Menus.Item("GSLoc20_1")
                    oMenus = oMenuItem.SubMenus
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "GSLoc20_7"
                    oCreationPackage.String = "Informe DINARDAP"
                    oCreationPackage.Enabled = True
                    Try
                        oMenus.AddEx(oCreationPackage)
                    Catch ex As Exception
                    End Try
                End If


            End If


            oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GSLoc30"
            oCreationPackage.String = "Anexos de Compras"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GSLoc40"
            oCreationPackage.String = "Anexos de Ventas"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


            oMenuItem = rSboApp.Menus.Item("GSLoc10")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GSLoc10_1"
            oCreationPackage.String = "Generar XML"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            If Functions.VariablesGlobales._SS_ActivarOpcionesParaManejoChequesProtestos = "Y" Then

                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "GSLoc50"
                oCreationPackage.String = "Cheques Protesto"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc50")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc50_1"
                oCreationPackage.String = "Generar Protesto"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

            End If

            If Functions.VariablesGlobales._SS_ActivarOpcionesParaTrasnferStockEntreBDs = "Y" Then

                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "GSLoc60"
                oCreationPackage.String = "Transferencia entre Compañias"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc60")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc60_1"
                oCreationPackage.String = "Generar Transferencia"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc60")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc60_2"
                oCreationPackage.String = "Transferencias Generadas"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try


            End If


            'Nuevas Funcionalidades
            If Functions.VariablesGlobales._PagosMasivos = "Y" Then


                'Pagos Masivos
                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "GSLoc70"
                oCreationPackage.String = "Procesamiento de Pagos"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc70")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc70_1"
                oCreationPackage.String = "Pagos Masivos"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc70")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc70_2"
                oCreationPackage.String = "Aprobacion de Pagos Masivos"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If

            If Functions.VariablesGlobales._ActivarCMFML = "Y" Then


                'CashManagement
                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "GSLoc90"
                oCreationPackage.String = "Cash Management"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc90")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc90_1"
                oCreationPackage.String = "Configuraciones CM"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc90")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc90_2"
                oCreationPackage.String = "Generar archivo"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If


            'Guias Remision Desatendidas
            If Functions.VariablesGlobales._GuiasDesatendidas = "Y" Then


                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
                oCreationPackage.UniqueID = "GSLoc80"
                oCreationPackage.String = "Guias de Remision"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

                oMenuItem = rSboApp.Menus.Item("GSLoc80")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc80_1"
                oCreationPackage.String = "Guia Desatendida"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try

            End If
            '-----------------------------

            'ADD 10/09/2024 JP
            If Functions.VariablesGlobales._ActServiciosBasicos = "Y" Then

                oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
                oMenus = oMenuItem.SubMenus
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "GSLoc_2"
                oCreationPackage.String = "Servicios Básicos"
                oCreationPackage.Enabled = True
                Try
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                End Try
            End If

            oMenuItem = rSboApp.Menus.Item("mnPrincipalLoc")
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GSLoc_1"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu Licencia FULL" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub Menu_CM(oCreationPackage As SAPbouiCOM.MenuCreationParams, oMenus As SAPbouiCOM.Menus, oMenuItem As SAPbouiCOM.MenuItem)

        'CashManagement
        oMenuItem = rSboApp.Menus.Item("mnPrincipal")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
        oCreationPackage.UniqueID = "GSLoc90"
        oCreationPackage.String = "Cash Management"
        oCreationPackage.Enabled = True
        Try
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception
        End Try

        oMenuItem = rSboApp.Menus.Item("GSLoc90")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "GSLoc90_1"
        oCreationPackage.String = "Configuraciones CM"
        oCreationPackage.Enabled = True
        Try
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception
        End Try

        oMenuItem = rSboApp.Menus.Item("GSLoc90")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "GSLoc90_2"
        oCreationPackage.String = "Generar archivo"
        oCreationPackage.Enabled = True
        Try
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Menu_ParaPropuestaComercial()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipal"
            oCreationPackage.String = "Asistente de Pagos"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logo11.png"
            oCreationPackage.Position = 15
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


            oMenuItem = rSboApp.Menus.Item("mnPrincipal")
            oMenus = oMenuItem.SubMenus

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS28"
            oCreationPackage.String = "Pago con Cheque"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try
            'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            'oCreationPackage.UniqueID = "GS29"
            'oCreationPackage.String = "Configuración Parametros"
            'oCreationPackage.Enabled = True
            'Try
            '    oMenus.AddEx(oCreationPackage)
            'Catch ex As Exception
            'End Try
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS30"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu SIN Licencia" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    'Private Sub CargarMenuDesdeXML()

    '    rEvento.rSboApp.Forms.GetFormByTypeAndCount(169, 1).Freeze(True)

    '    Dim sPath As String

    '    Try
    '        sPath = Application.StartupPath & "\"
    '        Dim xmlD As System.Xml.XmlDocument
    '        xmlD = New System.Xml.XmlDocument
    '        xmlD.Load(sPath & "Menu.xml")

    '        Dim oNodes As System.Xml.XmlNodeList
    '        Dim oNodeItem As System.Xml.XmlNode
    '        oNodes = xmlD.GetElementsByTagName("Menu")
    '        For Each oNodeItem In oNodes
    '            If oNodeItem.Attributes.GetNamedItem("Image").Value <> "" Then
    '                oNodeItem.Attributes.GetNamedItem("Image").Value = sPath & oNodeItem.Attributes.GetNamedItem("Image").Value
    '            End If
    '        Next
    '        rEvento.rSboApp.LoadBatchActions(xmlD.InnerXml)
    '    Catch ex As System.IO.FileNotFoundException
    '        rEvento.rSboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    Catch ex As Exception
    '        rEvento.rSboApp.MessageBox("Mensaje: " & ex.Message & " Pila: " & ex.StackTrace)
    '    End Try

    '    rEvento.rSboApp.Forms.GetFormByTypeAndCount(169, 1).Freeze(False)
    '    rEvento.rSboApp.Forms.GetFormByTypeAndCount(169, 1).Update()

    'End Sub
    Public Sub ActivarTLS()
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or 768 Or 3072
    End Sub

    Public Sub SetProtocolosdeSeguridad()



        'PARA TLS 1.2
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999



        'PARA HTTPS



        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)



    End Sub
    Function customCertValidation(ByVal sender As Object, _
                                     ByVal cert As X509Certificate, _
                                     ByVal chain As X509Chain, _
                                     ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function

End Module
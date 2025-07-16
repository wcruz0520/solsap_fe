Public Class VariablesGlobales


    'fin localizacion

    Public Shared _vgNombreAddOn As String
    Public Shared _vgVersionAddOn As String

    Public Shared _vgGuardarLog As String
    Public Shared _vgFolioLQUDF As String
    Public Shared _vgImpBlo As String
    Public Shared _vgPreLot As String
    Public Shared _vgSerieUDF As String
    Public Shared _vgFechaSalidaEnVivo As String

    Public Shared _vgProxy_puerto As String
    Public Shared _vgProxy_IP As String
    Public Shared _vgProxy_Usuario As String
    Public Shared _vgProxy_Clave As String

    Public Shared _gTimeOut_Emision As String
    Public Shared _vgHttps As String
    Public Shared _vgAmbiente As Integer = 0

    Public Shared _vgTipoLicenciaAddOn As String = ""

    Public Shared _vgTieneLicenciaActivaAddOn As Boolean = False
    Public Shared _vgCorreoResponsable As String
    Public Shared _vgVersionSRI As String
    Public Shared _vgVersionDisponibleAddOn As String
    Public Shared _vgReleaseNoteAddOn As String

    Public Shared _vgProcesoLoteManamer As String

    Public Shared _vgNoEnviarRT As String
    Public Shared _vgBloquearReenviarSRI As String


    Public Shared _vgServerNode As String
    Public Shared _vgUserBD As String
    Public Shared _vgPassBD As String

    Public Shared _vgMostrarLogo As String

    Public Shared _vgPruebaSinWSLIC As String

    Public Shared _vgQueryCorreo As String

    Public Shared _vgFechaEmisionFactura As String
    Public Shared _vgFechaEmisionNotaCredito As String
    Public Shared _vgFechaEmisionRetencion As String

    Public Shared _vgFechaEmisionFacturaP As String
    Public Shared _vgFechaEmisionNotaCreditoP As String
    Public Shared _vgFechaEmisionRetencionP As String

    Public Shared _SALIDA_POR_PROXY As String
    Public Shared _VisualizaPDFByte As String
    ' VARIABLES GLOBALES RECEPCION
    Public Shared _WS_Recepcion As String
    Public Shared _WS_RecepcionCambiarEstado As String
    Public Shared _WS_RecepcionClave As String
    Public Shared _WS_RecepcionCargaEstados As String
    Public Shared _WS_RecepcionConsultaArchivo As String
    Public Shared _nRegistros As String
    Public Shared _Adicional_FC As String
    Public Shared _Adicional_NC As String
    Public Shared _Adicional_RET As String
    Public Shared _MostrarFechaAutorizacion As String
    Public Shared _Nombre_CA As String ' NOMBRE CAMPO ADICIONAL
    Public Shared _CreaPedido As String
    Public Shared _PermiteDescuadre As String
    Public Shared _RecepcionLite As String
    Public Shared _Ruta_Compartida As String
    Public Shared _Prefijo_FC As String
    Public Shared _Cuenta_FC As String
    Public Shared _Prefijo_NC As String
    Public Shared _Cuenta_NC As String

    Public Shared _BFR As String ' RETENCION - RELACIONAR FACTURAS
    Public Shared _PagoRecibido_Seidor_exxis As String
    Public Shared _CodigoRetencion As String
    Public Shared _CodigoRetencionR As String
    Public Shared _Cuenta_RE As String
    Public Shared _CampoNumRetencion As String
    Public Shared _FormaPagoCompras As String
    Public Shared _ValorFormaPagoCompras As String

    Public Shared _AsignarNumeroDocEnNumAtCard As String
    Public Shared _CrearFCdeReservaProveedores As String

    Public Shared _wsEmisionFactura As String
    Public Shared _wsEmisionNotaCredito As String
    Public Shared _wsEmisionGuiaRemision As String
    Public Shared _wsEmisionNotaDebito As String
    Public Shared _wsEmisionRetencion As String
    Public Shared _wsEmisionLiquidacionCompra As String
    Public Shared _wsConsultaEmision As String
    Public Shared _wsReenvioMail As String
    Public Shared _wsClaveEmision As String
    Public Shared _TipoEmision As String
    Public Shared _TipoWS As String
    Public Shared _FoliacionPostin As String
    Public Shared _EnviarBackGroung As String

    Public Shared _MarcarContabiliadosManualFC As String
    Public Shared _MarcarContabiliadosManualNC As String
    Public Shared _MarcarContabiliadosManualRT As String

    Public Shared _AnuDocNumAtCard As String

    Public Shared _XMLRecepcionHeison As String
    Public Shared _RutaFC As String
    Public Shared _RutaNC As String
    Public Shared _RutaRT As String
    Public Shared _RutaProFC As String
    Public Shared _RutaProNC As String
    Public Shared _RutaProRT As String
    Public Shared _PreliminarLoteXML As String

    Public Shared _ImpDocAut As String

    Public Shared _IntegracionEcuanexus As String
    Public Shared _WsEmisionEcua As String
    Public Shared _WsEmisionConsultaEcua As String
    Public Shared _Token As String
    Public Shared _NombreWsEcua As String
    Public Shared _ConsultaFolioSS As String

    Public Shared _AsignarFolioalReenviarSolsap As String

    Public Shared _SINCRO_RT As String
    Public Shared _SINCRO_LQE As String
    Public Shared _SINCRO_DOC As String

    Public Shared _ReenviarDocsPantala As String
    Public Shared _ReenviarListaDocEnv As String

    Public Shared _FacturaGuiaRemision As String
    Public Shared _SalidaMercanciasGuiaRemision As String

    Public Shared _ValidarFechasCTK As String

    Public Shared _ValidarCamposNulos As String

    Public Shared _ImpresionDobleCara As String

    Public Shared _SS_Impresoras As New List(Of DatosImpresora)

    Public Shared _SS_FacturaExportacion As String

    Public Shared _ProveedorSAP As String

    Public Shared _WsLicencia As String

    Public Shared _RucCompañia As String

    Public Shared _TipoWsLicencia As String

    Public Shared _DiasValidarProcesoLote As String

    Public Shared _PL3FECHAS As String

    Public Shared _PL1FECHAS As String

    'variables de Localizacion
    Public Shared _gNombreAddOn As String
    Public Shared _gVersionAddOn As String

    Public Shared _CabeceraQRY As String

    Public Shared _EmiReceptorQRY As String

    Public Shared _ImpuestosQRY As String

    Public Shared _DesCargosQRY As String

    Public Shared _InformacionFiscalQRY As String

    Public Shared _ReteQRY As String

    Public Shared _AnticipoQRY As String

    Public Shared _ItotalesQRY As String

    Public Shared _FpagoQRY As String

    Public Shared _DetalleQry As String

    Public Shared _AdicionalFCQry As String

    Public Shared _AdicionalNCQry As String

    Public Shared _AdicionalNDQry As String

    Public Shared _DocsEnviadosQRY As String

    Public Shared _DocsIntegradosQRY As String

    Public Shared _EntregaQRY As String
    'VARIABLES GOBALES PARA LAS URLS DE LOS WS

    Public Shared _gEmisionTipo As String
    Public Shared _gEmisionClave As String
    Public Shared _gWS_RecepcionClave As String

    Public Shared _gGuardaLogEmision As String

    Public Shared _gWS_EmisionFC As String = ""
    Public Shared _gWS_EmisionND As String = ""
    Public Shared _gWS_EmisionNC As String = ""
    Public Shared _gWS_EmisionConsulta As String = ""
    Public Shared _gWS_EmisionConsultaFiles As String = ""
    Public Shared _gWS_ReenvioMail As String = ""
    Public Shared _gWS_Utilidades As String = ""

    'RECEPCION
    Public Shared _gWS_RecepcionConsulta As String = ""
    Public Shared _gWS_RecepcionEstado As String = ""
    Public Shared _gWS_RecepcionMR As String = ""

    'OTRAS AJUSTES GLOBALES
    ' Public Shared _gTimeOut_Emision As String = ""
    'Public Shared _gIsInHouse As String = ""
    Public Shared _gNO_ConsumirMetodoHTTPS As String = ""
    Public Shared _gRUCEmisor As String = ""
    Public Shared _gTipoSocioNegocio As String = ""
    Public Shared _gGeneracion_Cufe_QR As String = ""
    Public Shared _gRutaRPT As String = ""
    Public Shared _gUsuarioDB As String = ""
    Public Shared _gPasswordDB As String = ""
    Public Shared _gAdjuntos As String = ""
    '- Necesario para Generacion de Cufe
    Public Shared _gAmbiente As Integer = 0
    Public Shared _gPinSoftware As String = ""

    'LICENCIA
    Public Shared _gTipoLicenciaAddOn As String = ""
    Public Shared _gTieneLicenciaActivaAddOn As Boolean = False
    Public Shared _CorreoResponsable As String = ""
    Public Shared _gVersiondelMinisterioDeHacienda As String = ""

    Public Shared _gVersionDisponibleAddOn As String = ""
    Public Shared _gReleaseNoteAddOn As String = ""

    'VALIDACIONES
    Public Shared _gValidacionNit As String = ""
    Public Shared _gValidacionObligacionFiscales As String = ""
    Public Shared _gVTipoFactura As String = ""
    Public Shared _gVTipoOperacionDoc As String = ""
    Public Shared _gVTipoNotaCredito As String = ""
    Public Shared _VTipoNotaDebito As String = ""
    Public Shared _gVtipoDescuento As String = ""
    Public Shared _gVMediodePago As String = ""
    Public Shared _gVInfoReferencia As String = ""
    Public Shared _gVImpuestosMapeados As String = ""
    Public Shared _gVCamposExportacion As String = ""
    Public Shared _gVmailReceptor As String = ""
    Public Shared _gNombreCampoTipoIdentificacion As String = ""
    Public Shared _gNombreCampoDireccion As String = ""
    Public Shared _gNombreCampoMunicipio As String = ""

    Public Shared _TipoLicenciaAddonLE As String = ""

    Public Shared _RutaRPT As String = ""
    Public Shared _ipServer As String = ""


    'Aqui se guardara los datos de los querys de los anexos de compra y venta

    Public Shared _SS_ComprasQRY As String = ""
    Public Shared _SS_VentasQRY As String = ""
    Public Shared _SS_CalculoFolioQRY As String = ""


    ' Se mantentendran en memoria los Archivos de los RPTS , Se realizaran pruebas
    ' de factibilidad 

    Public Shared _SS_RPTFacturasCompras As String = ""
    Public Shared _SS_RPTFormularioSRI103 As String = ""
    Public Shared _SS_RPTListadoVentas As String = ""
    Public Shared _SS_RPTRetencionClientes As String = ""

    'add 13092022
    Public Shared _SS_AbrirRPTCargadosDB As String = ""

    'add 07102022
    Public Shared _SS_GenerarFolio As String = ""

    'add 14102022
    Public Shared _SS_ValidarSocioNegociosUDF As String = ""
    Public Shared _SS_ValidarDocumentosUDF As String = ""

    'Variables para el manejo de Cheques Protestados
    Public Shared _SS_ActivarOpcionesParaManejoChequesProtestos As String = ""

    Public Shared _SS_IDSerie As String = ""

    Public Shared _SS_ImpuestoMontoCheque As String = ""
    Public Shared _SS_NombreImpuestoMontoCheque As String = ""
    Public Shared _SS_impuestoMontoProtesto As String = ""
    Public Shared _SS_NombreImpuestoMontoProtesto As String = ""
    Public Shared _SS_RutaReporteND As String = ""

    'Variables para el manejo de transferencias de sctock entre almacenes
    Public Shared _SS_ActivarOpcionesParaTrasnferStockEntreBDs As String = ""

    Public Shared _SS_DINARDAP As String = ""

    Public Shared _SS_CuentaTransferencia As String = ""

    Public Shared _SS_UserDB As String = ""
    Public Shared _SS_PassDB As String = ""
    Public Shared _SS_UserSAP As String = ""
    Public Shared _SS_PassSAP As String = ""

    Public Shared _SS_IdSalidaMercancias As String = ""
    Public Shared _SS_IdEntradaMercancias As String = ""

    Public Shared _SS_SetearCamposUsuario As String = ""

    'add Arturito 14072023
    Public Shared _SS_MostrarBotonImpresion As String = ""

    Public Shared _SS_ConfImpresoras As New List(Of DatosImpresora)

    Public Shared _SS_SConexionHana As String = ""

    Public Shared _SS_DriverHana As String = ""

    Public Shared _SS_ParametroRPT As String = ""

    Public Shared _SS_CamposFV As New List(Of NombreCampos)

    'ArturoDEv 05042024 
    'se agrega variable para el parametro de la localizacion

    Public Shared _ActivarLocalizacionEC As String = ""

    'add 09072024 activar cash management fuera de Menu de loc
    Public Shared _ActivarCMFML As String = ""

    'add 03/09/2024 Parametro para contabilizar pago recibido desde proceso lote
    Public Shared _ContabilizarPRPL As String = ""

    'ruta del los XML

    Public Shared _RutaIntegracionXML As String = ""

    'Replace datatables
    Public Shared _TablasNativasReplace As String = ""

    'Artur Querys encryptados

    Public Shared _Query_FacturaSeccion01 As String = ""
    Public Shared _Query_FacturaSeccion02 As String = ""
    Public Shared _Query_FacturaAnticipoSeccion01 As String = ""
    Public Shared _Query_FacturaAnticipoSeccion02 As String = ""
    Public Shared _Query_NotaCreditoSeccion01 As String = ""
    Public Shared _Query_NotaCreditoSeccion02 As String = ""
    Public Shared _Query_NotaDebitoSeccion01 As String = ""
    Public Shared _Query_NotaDebitoSeccion02 As String = ""

    Public Shared _Query_CompleExportacion As String = ""
    Public Shared _Query_CompleReembolso As String = ""
    Public Shared _Query_GuiaRemisionSeccion01 As String = ""
    Public Shared _Query_GuiaRemisionSeccion02 As String = ""
    Public Shared _Query_RetencionSeccion01 As String = ""
    Public Shared _Query_RetencionSeccion02 As String = ""
    Public Shared _Query_LiquidacionSeccion01 As String = ""
    Public Shared _Query_LiquidacionSeccion02 As String = ""

    Public Shared _Query_DocumentosEnviados As String = ""

    'Guias desatendidas

    Public Shared _Query_GuiasDesatendidas01 As String = ""
    Public Shared _Query_GuiasDesatendidas02 As String = ""


    'Querys Recepcion

    Public Shared _Query_ReDocumentosMarcados As String = ""
    Public Shared _Query_ReDocumentosIntegrados As String = ""

    'add 052024
    'KingArtur
    Public Shared _PosicionItemTabX As String = ""
    Public Shared _PosicionItemTabY As String = ""

    'query guia udo
    Public Shared _SINCRO_GRUDO As String = ""

    Public Shared _GuiasDesatendidas As String = ""
    Public Shared _QueryGuiasDesatendidasSeries As String = ""
    Public Shared _QueryGuiasDesatendidasProximoDocNum As String = ""
    Public Shared _PagosMasivos As String = ""
    Public Shared _ManejoCuenta As String = ""
    Public Shared _RutaArchivoTxt As String = ""

    Public Shared _UsuarioParamDias As String

    'Add JP 23/08/2024
    Public Shared _RutaArchivoRPTPM As String = ""
    Public Shared _RutaArchivoCHQPM As String = ""
    Public Shared _CadenaConexionRPTPM As String = ""
    Public Shared _CuentaTransitoriaPM As String = ""

    'Add JP 13/11/2024
    Public Shared _SBMenuPadreRptPreInf As String = ""
    Public Shared _RutaArchivoCEPM As String = ""

    'Add DM 6/12/2024
    Public Shared _NombreColumnbasAnexo As String = ""

    'Add DM 17/12/2024
    Public Shared _NombreCampoPedidoInfoAd As String = ""

    'Add DM 14/01/2025
    Public Shared _FechaAutEnFechaContabFC As String = ""
    Public Shared _FechaAutEnFechaContabNC As String = ""
    Public Shared _FechaAutEnFechaContabRT As String = ""

    'Add JP 15/01/2025
    Public Shared _RutaReposCM As String = ""

    'Add DM 16/01/2025
    Public Shared _ComentarioPago As String = ""

    Public Shared _ActServiciosBasicos As String = ""

    Public Shared FechaFinMesAnterior As String = ""
    Public Shared FechaFinMesAnteriorPL As String = ""

    '16/06/2025 JP
    Public Shared _ActApiSS As String = ""
    Public Shared _ApiAutSS As String = ""
    Public Shared _ApiFactEmiSS As String = ""

End Class

'Public Class DatosImpresora

'    Public Property Usuario As String

'    Public Property Impresora As String

'End Class

Public Class DatosImpresora

    Public Property TipoDocumento As String
    Public Property IdReporte As String

    Public Property ImpresoraCompartida As String

    Public Property Usuario As String

    Public Property Impresora As String

    Public Property ImpresoraIndividual As String

End Class

Public Class NombreCampos

    Public Property Descripcion As String
    Public Property Nombre As String

End Class
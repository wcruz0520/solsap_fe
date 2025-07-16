Imports System.Timers
Imports System.Text
Imports System.IO
Imports System.Configuration
Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Net
Imports System.Threading
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Web.Compilation
Imports System.Xml.Serialization

Public Class Servico_Ivend_Citikold


    Private ReadOnly miobj As New Object
    Private ReadOnly lockgetreader As New Object
    Private ConsultaDocumentos As String = "GS_SAP_FE_ObtenerDocumentosPendientes_v2"
    Private oTimer As System.Timers.Timer = Nothing
    Private oTimerSincro As System.Timers.Timer = Nothing
    Private blnInicio As Boolean = False
    Private blnInicio_sincro As Boolean = False

    Public oCompany As SAPbobsCOM.Company

    Dim observacion As String, ProveedorSapBo As String
    Dim RangodeRegistrosEmision As Integer, RangodeRegistrosSincronizacion As Integer

    Private sServerType As Integer = 0

    Dim APIKEY_SG As String = "", CORREOASUNTO_SG As String = "", CORREOFROM_SG As String = "", CORREOTO_SG As String = "", URL_SG As String = "", MAXDOC_ALERTA As Integer = 0

    Private Const NombreServicio As String = "SAED_Procesador_EC"
    Private Const VersionServicio As String = "1.5"
    Private Const CodigoPais As String = "EC"

    Dim estadoLicencias As Boolean = False
    Dim versionTributaria As String = ""

    Dim RucCliente As String = ""



    Dim listaParametros() As Entidades.wsSS_LICENCIA_SAP.ClsConfigValores = Nothing



    'Sub New()
    '    'descomentar el sub new para pruebas y limpiar y generar nuevamente para que no de error
    '    ' Llamada necesaria para el diseñador.
    '    InitializeComponent()
    '    servicio_Core()



    '    ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    'End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)

        servicio_Core()

    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        oTimer.Stop()
    End Sub

    Protected Overrides Sub OnPause()
        oTimer.Stop()
    End Sub
    Protected Overrides Sub OnContinue()
        oTimer.Start()
    End Sub

    Private Sub servicio_Core()

        GuardaLog("Servicio Iniciado..Procediendo a Conectarse")
        GuardaLog("El motor de la Base de Datos Detectado es de tipo  : " + ConfigurationManager.AppSettings("DevServerType"))
        sServerType = Convert.ToInt32(ConfigurationManager.AppSettings("DevServerType"))
        ProveedorSapBo = ConfigurationManager.AppSettings("Tipo_Pch_Sap").ToString()
        RangodeRegistrosEmision = Convert.ToInt32(ConfigurationManager.AppSettings("RangodeRegistrosEmision"))
        RangodeRegistrosSincronizacion = Convert.ToInt32(ConfigurationManager.AppSettings("RangodeRegistrosSincronizacion"))

        Try

            URL_SG = ConfigurationManager.AppSettings("URL_SG").ToString()

        Catch ex As Exception
            URL_SG = ""
        End Try

        Try

            APIKEY_SG = ConfigurationManager.AppSettings("APIKEY_SG").ToString()

        Catch ex As Exception
            APIKEY_SG = ""
        End Try

        Try

            CORREOASUNTO_SG = ConfigurationManager.AppSettings("CORREOASUNTO_SG").ToString()

        Catch ex As Exception
            CORREOASUNTO_SG = ""
        End Try



        Try

            CORREOFROM_SG = ConfigurationManager.AppSettings("CORREOFROM_SG").ToString()

        Catch ex As Exception
            CORREOFROM_SG = ""
        End Try


        Try

            CORREOTO_SG = ConfigurationManager.AppSettings("CORREOTO_SG").ToString()

        Catch ex As Exception
            CORREOTO_SG = ""
        End Try

        Try

            MAXDOC_ALERTA = Convert.ToInt32(ConfigurationManager.AppSettings("MAXDOC_ALERTA").ToString())

        Catch ex As Exception
            MAXDOC_ALERTA = ""
        End Try

        If conectSAP() Then


            If (oCompany.Connected) Then

                GuardaLog(String.Format("Conexion SAP exitosa, CompanyName={0} ,DataBase={1}  ", oCompany.CompanyName, oCompany.CompanyDB))

            End If

            If CheckLIC() And estadoLicencias Then

                Try
                    GETPArametros_INIT()
                    GuardaLog("Parametros Obtenidos")

                    'procesar_docEnvios()

                    Dim timer As Int64 = ConfigurationManager.AppSettings("timer") * 1000
                    oTimer = New System.Timers.Timer(timer)
                    oTimerSincro = New System.Timers.Timer(timer)
                    GuardaLog("Temporizador Para Sincronizacion y Envio de Facturas Seteados a: " + timer.ToString())
                    oTimer.AutoReset = True
                    oTimer.Enabled = False
                    oTimerSincro.AutoReset = True
                    oTimerSincro.Enabled = False

                    AddHandler oTimer.Elapsed, AddressOf oTimer_Elapsed

                    AddHandler oTimerSincro.Elapsed, AddressOf oTimerSincro_Elapsed

                    oTimer.Start()
                    oTimerSincro.Start()

                    GuardaLog("Temporizadores  Iniciados...")


                Catch ex As ConfigurationErrorsException
                    GuardaLog("Error al Inicializar Los Temporizadores: " + ex.Message.ToString())
                End Try

            End If




        Else

            GuardaLog("Error en la Conexion SAP , Validar el Archivo de Configuracion e Iniciar nuevamente el Servicio ")
        End If


    End Sub


    Private Sub GETPArametros_INIT()

        Functions.VariablesGlobales._SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
        Functions.VariablesGlobales._IntegracionEcuanexus = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "IntegracionEcuanexus")
        Functions.VariablesGlobales._WsEmisionConsultaEcua = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WsConsultaEcu")
        Functions.VariablesGlobales._vgGuardarLog = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog")
        Functions.VariablesGlobales._Token = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TokenEcu")
        Functions.VariablesGlobales._ValidarCamposNulos = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "ValidarCamposNulos")

        Functions.VariablesGlobales._Query_CompleExportacion = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_CompleExportacion")

        Functions.VariablesGlobales._Query_CompleReembolso = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_CompleReembolso")


        Functions.VariablesGlobales._Query_FacturaSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_FacturaSeccion01")
        Functions.VariablesGlobales._Query_FacturaSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_FacturaSeccion02")

        Functions.VariablesGlobales._Query_FacturaAnticipoSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_FacturaAnticipoSeccion01")
        Functions.VariablesGlobales._Query_FacturaAnticipoSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_FacturaAnticipoSeccion02")

        Functions.VariablesGlobales._Query_NotaCreditoSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_NotaCreditoSeccion01")
        Functions.VariablesGlobales._Query_NotaCreditoSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_NotaCreditoSeccion02")

        Functions.VariablesGlobales._Query_NotaDebitoSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_NotaDebitoSeccion01")
        Functions.VariablesGlobales._Query_NotaDebitoSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_NotaDebitoSeccion02")

        Functions.VariablesGlobales._Query_GuiaRemisionSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_GuiaRemisionSeccion01")
        Functions.VariablesGlobales._Query_GuiaRemisionSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_GuiaRemisionSeccion02")

        Functions.VariablesGlobales._Query_RetencionSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_RetencionSeccion01")
        Functions.VariablesGlobales._Query_RetencionSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_RetencionSeccion02")


        Functions.VariablesGlobales._Query_LiquidacionSeccion01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_LiquidacionSeccion01")
        Functions.VariablesGlobales._Query_LiquidacionSeccion02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "Query_LiquidacionSeccion02")

        Functions.VariablesGlobales._Query_GuiasDesatendidas01 = ConsultaParametro("SAED", "PARAMETROS", "BD", "ValidarCamposNulos")
        Functions.VariablesGlobales._Query_GuiasDesatendidas02 = ConsultaParametro("SAED", "PARAMETROS", "BD", "ValidarCamposNulos")


    End Sub


    'Private Sub oTimer_Elapsed()

    '    Try
    '        If blnInicio = False Then
    '            blnInicio = True
    '            GuardaLog("Ingresando al Hilo de Procesamiento de Documentos!!")
    '            Dim hilo As New Thread(AddressOf procesoUnHilo)
    '            hilo.Start()
    '            hilo.Join()
    '            GuardaLog("Retomando Control Hilo del Temporizador!!")
    '            blnInicio = False
    '        End If


    '    Catch ex As Exception
    '        GuardaLog("Error al tratar de Invocar la funcion de procesar_docEnvios en el Ciclo oTimer_Elapsed " + ex.Message.ToString())
    '    End Try

    'End Sub
    Private Sub oTimer_Elapsed()

        Try
            If blnInicio = False Then
                blnInicio = True
                GuardaLog("Ingresando al Hilo de Procesamiento de Documentos!!")
                'Dim hilo As New Thread(AddressOf procesoUnHilo)
                'hilo.Start()
                'hilo.Join()
                procesar_docEnvios()
                GuardaLog("Retomando Control Hilo del Temporizador!!")
                blnInicio = False
            End If


        Catch ex As Exception
            GuardaLog("Error al tratar de Invocar la funcion de procesar_docEnvios en el Ciclo oTimer_Elapsed " + ex.Message.ToString())
        End Try

    End Sub
    Private Sub oTimerSincro_Elapsed(sender As Object, e As ElapsedEventArgs)

        Try
            If blnInicio_sincro = False Then
                blnInicio_sincro = True

                procesar_docSincro()

                blnInicio_sincro = False
            End If


        Catch ex As Exception

            GuardaLog("Error al tratar de Invocar la funcion de procesar_docSincro en el Ciclo oTimer_Elapsed " & ex.Message.ToString())

        End Try

    End Sub

    Public Sub procesoUnHilo()


        Try
            ClearMemory()

            Dim dataTable As DataTable = Nothing
            Dim dataTable2 As DataTable = Nothing

            GuardaLog("Consultando Documentos para Emision..")

            Try

                dataTable = GetDataTable(ConsultaDocumentos, "ENVIO")

            Catch ex As Exception

                dataTable = Nothing

                GuardaLog("Exepcion en consultar funcion GetDataTable EMISION ex:" + ex.Message)

            End Try

            GuardaLog("Consultando Documentos para Sincronizacion..")

            Try

                dataTable2 = GetDataTable(ConsultaDocumentos, "SINCRO")

            Catch ex As Exception

                dataTable2 = Nothing

                GuardaLog("Exepcion en consultar funcion GetDataTable SINCRONIZACION ex:" + ex.Message)

            End Try

            If (dataTable Is Nothing) Then

                GuardaLog("tabla de Documentos Para Emitir es nothing")

            Else

                GuardaLog("Documentos a procesar EMISION  ==> " + dataTable.Rows.Count.ToString)
                If (dataTable.Rows.Count > 0) Then

                    ProcesoEnvioAlerta(dataTable.Rows.Count)

                    Dim list = dataTable.Select().ToList()

                    Dim num = 0
                    num = IIf(RangodeRegistrosEmision <= list.Count, RangodeRegistrosEmision, list.Count)
                    trabajo_hilo_emision(list.GetRange(0, num))
                End If

            End If

            If (dataTable2 Is Nothing) Then

                GuardaLog("tabla de Documentos Para Sincronizar es nothing")

            Else

                GuardaLog("Documentos a SINCRONIZAR  ==> " + dataTable2.Rows.Count.ToString)
                If (dataTable2.Rows.Count > 0) Then

                    Dim list2 = dataTable2.Select().ToList()
                    Dim num2 = 0
                    num2 = IIf(RangodeRegistrosSincronizacion <= list2.Count, RangodeRegistrosSincronizacion, list2.Count)
                    trabajo_hilo_sincro(list2.GetRange(0, num2))
                End If

            End If

            GuardaLog("Proceso Hilos Completado Correctamente!!")

        Catch ex As Exception
            GuardaLog("Error en Funcion procesoOptimizado, EX  ==> " + ex.Message)
        End Try

    End Sub

    Private Sub procesar_docSincro()
        Try

            ClearMemory()

            Dim dtDoc As DataTable = Nothing
            Try
                dtDoc = GetDataTable(ConsultaDocumentos, "SINCRO")
            Catch ex As Exception
                dtDoc = Nothing
                GuardaLog("Exepcion en consultar funcion GetDataTable ex:" + ex.Message)
            End Try


            If dtDoc Is Nothing Then
                GuardaLog("tabla de consultaDocuemtnos es nothing")
                Exit Sub
            End If

            GuardaLog("Documentos a procesar SINCRONIZACION  ==> " + dtDoc.Rows.Count.ToString)

            'Si hay registros procesamos 
            If dtDoc.Rows.Count > 0 Then
                If Not oCompany.Connected Then

                    Dim ErrCode = oCompany.Connect()

                    If ErrCode <> 0 Then

                        GuardaLog("Error al conectarse a SAP ,funcion :procesar_docSincro : " + oCompany.GetLastErrorDescription)

                    End If

                End If

                If oCompany.Connected Then


                    Dim rows_sincronizacion As List(Of DataRow) = dtDoc.Select().ToList

                    GuardaLog(String.Format("Filas para Sincronizacion {0}", rows_sincronizacion.Count.ToString))

                    Dim hilos_sincronizacion As Integer = IIf(ConfigurationManager.AppSettings("TotalHilosSincro") = "", 1, CInt(ConfigurationManager.AppSettings("TotalHilosSincro")))

                    'TRABAJO DE HILOS

                    Try


                        GuardaLog(String.Format("lista hilos {0}", hilos_sincronizacion.ToString))
                        Dim listaHilosSincronizacion As New List(Of Thread)


                        If rows_sincronizacion.Count > 0 Then

                            For he = 1 To hilos_sincronizacion

                                listaHilosSincronizacion.Add(New Thread(AddressOf trabajo_hilo_sincro) With {.Name = "hilo_sincronizacion" & CStr(he)})

                            Next

                        End If
                        GuardaLog(String.Format("1 lista hilos {0}", hilos_sincronizacion.ToString))

                        '----------------------------------------------------
                        Dim promhilo_sincro = Decimal.Truncate(rows_sincronizacion.Count / hilos_sincronizacion)
                        Dim promhilo_sincro_R = CInt(rows_sincronizacion.Count Mod hilos_sincronizacion)
                        For Each he As Thread In listaHilosSincronizacion

                            If promhilo_sincro = 0 Then

                                he.Start(rows_sincronizacion)

                                Exit For
                            End If

                            he.Start(rows_sincronizacion.GetRange(0, promhilo_sincro + promhilo_sincro_R))
                            rows_sincronizacion.RemoveRange(0, promhilo_sincro + promhilo_sincro_R)

                            promhilo_sincro_R = 0

                        Next



                        'espero a que los hilos finalicen
                        For Each he As Thread In listaHilosSincronizacion
                            he.Join()

                        Next



                        listaHilosSincronizacion.Clear()

                        rows_sincronizacion.Clear()

                        listaHilosSincronizacion = Nothing

                        rows_sincronizacion = Nothing

                    Catch ex As Exception

                        GuardaLog("Error al crear Hilos para Sincronizar ==> " + ex.Message.ToString)

                    End Try



                End If
            End If



        Catch ex As Exception
            GuardaLog("Error general en la funcion de procesar_docSincro " & ex.Message.ToString())
        End Try
    End Sub

    Private Declare Auto Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal procHandle As IntPtr, ByVal min As Int32, ByVal max As Int32) As Boolean
    'Funcion de liberacion de memoria
    Shared Function ClearMemory(Optional ByRef msgg = "") As Boolean
        Try
            If Process.GetCurrentProcess().WorkingSet64 / 1048576 > 150 Then
                GC.Collect()
                Try
                    GC.WaitForPendingFinalizers()
                    Try
                        Dim Mem As Process
                        Mem = Process.GetCurrentProcess()
                        SetProcessWorkingSetSize(Mem.Handle, -1, -1)
                    Catch ex As Exception
                        msgg = ex.ToString
                    End Try
                Catch ex As Exception
                    msgg = ex.ToString
                End Try
            End If
        Catch ex As Exception
        End Try
        Return True
    End Function

    Private Sub procesar_docEnvios()
        Try
            ClearMemory()

            Dim dtDoc As DataTable = Nothing

            Try
                dtDoc = GetDataTable(ConsultaDocumentos, "ENVIO")
            Catch ex As Exception
                dtDoc = Nothing
                GuardaLog("Exepcion en consultar funcion GetDataTable ex:" + ex.Message)
            End Try



            If dtDoc Is Nothing Then
                GuardaLog("tabla de consultaDocuemtnos es nothing")
                Exit Sub
            End If

            GuardaLog("Documentos a procesar EMISION  ==> " + dtDoc.Rows.Count.ToString)
            'Si hay registros procesamos 
            If dtDoc.Rows.Count > 0 Then

                'envio de mail de alerta

                ProcesoEnvioAlerta(dtDoc.Rows.Count)

                'fin alerta

                If Not oCompany.Connected Then

                    Dim ErrCode = oCompany.Connect()

                    If ErrCode <> 0 Then

                        GuardaLog("Error al conectarse a SAP ,funcion :procesar_docEnvios : " + oCompany.GetLastErrorDescription)

                    End If

                End If

                If oCompany.Connected Then


                    Dim rows_emision As List(Of DataRow) = dtDoc.Select().ToList

                    GuardaLog(String.Format("Filas para Emision {0}", rows_emision.Count.ToString))

                    Dim hilos_emision As Integer = IIf(ConfigurationManager.AppSettings("TotalHilosEmision") = "", 1, CInt(ConfigurationManager.AppSettings("TotalHilosEmision")))

                    'TRABAJO DE HILOS

                    Try



                        Dim listaHilosEmision As New List(Of Thread)


                        If rows_emision.Count > 0 Then

                            For he = 1 To hilos_emision

                                listaHilosEmision.Add(New Thread(AddressOf trabajo_hilo_emision) With {.Name = "hilo_emision" & CStr(he)})

                            Next

                        End If


                        '----------------------------------------------------
                        Dim promhilo_emision = Decimal.Truncate(rows_emision.Count / hilos_emision)
                        Dim promhilo_emision_R = CInt(rows_emision.Count Mod hilos_emision)
                        For Each he As Thread In listaHilosEmision

                            If promhilo_emision = 0 Then

                                he.Start(rows_emision)

                                Exit For
                            End If

                            he.Start(rows_emision.GetRange(0, promhilo_emision + promhilo_emision_R))
                            rows_emision.RemoveRange(0, promhilo_emision + promhilo_emision_R)

                            promhilo_emision_R = 0

                        Next



                        'espero a que los hilos finalicen
                        For Each he As Thread In listaHilosEmision
                            he.Join()

                        Next



                        listaHilosEmision.Clear()

                        rows_emision.Clear()

                        listaHilosEmision = Nothing

                        rows_emision = Nothing

                    Catch ex As Exception

                        GuardaLog("Error al crear Hilos ==> " + ex.Message.ToString)

                    End Try



                End If
            End If



        Catch ex As Exception
            GuardaLog("Error general en la funcion de procesar_docEnvios " & ex.Message.ToString())
        End Try
    End Sub

    Private Sub trabajo_hilo_emision(objdr As List(Of DataRow))

        Try
            Dim oNegocioEcua As Negocio.ManejoDeDocumentosEcua = Nothing
            Dim oNegocio As Negocio.ManejoDeDocumentos = Nothing
            Dim nombrehilo As String = ""
            nombrehilo = Thread.CurrentThread.Name
            Dim tipodocumento As String = ""

            GuardaLog(String.Format("Ejecutando funcion trabajo_hilo_emision,{0} documentos a Proce x hilo {1} ", objdr.Count.ToString, nombrehilo))

            Try
                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then

                    oNegocioEcua = New Negocio.ManejoDeDocumentosEcua(oCompany, Nothing, "S", ConfigurationManager.AppSettings("Tipo_Pch_Sap"))


                Else

                    oNegocio = New Negocio.ManejoDeDocumentos(oCompany, Nothing, "S", ConfigurationManager.AppSettings("Tipo_Pch_Sap"))


                End If

            Catch ex As Exception
                GuardaLog("Error al instanciar la clase ManejoDeDocumentos: Error" + ex.Message.ToString)
                Exit Sub
            End Try
            '-------------------------


            For Each row As DataRow In objdr


                Dim oDocumentoTraza As New DocumentosTrans()
                oDocumentoTraza.Code = row("Code")
                oDocumentoTraza.DocEntry = row("U_DocEntry")
                oDocumentoTraza.DocSubType = row("U_DocSubType")
                oDocumentoTraza.SRI_Code = CStr(row("U_SRI_Code"))
                oDocumentoTraza.Oberva = ""

                Try
                    If oDocumentoTraza.SRI_Code.ToString().Equals("01") Then ' FACTURA

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "FCE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "FCE")
                        End If


                        tipodocumento = "Factura"
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("05") Then ' NOTA DE DEBITO

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "NDE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "NDE")
                        End If


                        tipodocumento = "NOTA DEBITO"
                        ' GUIA DE REMISION - ENTREGA

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("04") Then ' NOTA DE CREDITO

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "NCE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "NCE")
                        End If


                        tipodocumento = "NOTA DE CREDITO"
                        ' GUIA DE REMISION - ENTREGA
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("15") Then

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "GRE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "GRE")
                        End If


                        tipodocumento = "GUIA DE REMISION - ENTREGA"
                        ' GUIA DE REMISION- TRANSFERENCIA
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("67") Then

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "TRE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "TRE")
                        End If

                        tipodocumento = "GUIA DE REMISION- TRANSFERENCIA"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("1250000001") Then

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "TLE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "TLE")
                        End If


                        tipodocumento = "GUIA DE SOLICITUD DE TRASLADO"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("07") Then

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "REE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "REE")
                        End If


                        tipodocumento = "RETENCION"
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("03") Or oDocumentoTraza.DocSubType.ToString().Equals("41") Then

                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                            oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "LQE")
                        Else
                            oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "LQE")
                        End If


                        tipodocumento = "LIQUIDACION"

                    End If

                    GuardaLog(oDocumentoTraza.Oberva + "- Tipo de Documento : " + tipodocumento + ": Escrito por hilo :" + CStr(nombrehilo) + " Docentry " + oDocumentoTraza.DocEntry.ToString)
                    ' ActualizaTabla(oDocumentoTraza)


                Catch ex As Exception

                    GuardaLog("trabajo_hilo_emision excepcion,ProcesaEnvioDocumento - " & ex.Message.ToString)

                End Try

                oDocumentoTraza = Nothing


            Next

            oNegocio = Nothing
            oNegocioEcua = Nothing

        Catch ex As Exception
            GuardaLog("trabajo_hilo_emision excepcion,ERROR GENERAL FUNCION - " + ex.Message.ToString())
        End Try

        GC.Collect()


    End Sub


    Private Sub trabajo_hilo_sincro(objdr As List(Of DataRow))

        Try
            Dim oNegocioEcua As Negocio.ManejoDeDocumentosEcua = Nothing
            Dim oNegocio As Negocio.ManejoDeDocumentos = Nothing
            Dim nombrehilo As String = "", prefijoDLL As String = "", postSincro As String = "", modoSincro As String = ""
            nombrehilo = Thread.CurrentThread.Name
            'bandera para saber si se usara el metodo de envio por el WS de sincronizacion o emision
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PosSincro")) Then
                postSincro = ConfigurationManager.AppSettings("PosSincro")
            Else
                postSincro = "1"
            End If

            GuardaLog(String.Format("Ejecutando funcion trabajo_hilo_sincro,{0} documentos a Proce x hilo {1} ", objdr.Count.ToString, nombrehilo))

            Try
                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then

                    oNegocioEcua = New Negocio.ManejoDeDocumentosEcua(oCompany, Nothing, "S", ConfigurationManager.AppSettings("Tipo_Pch_Sap"))



                Else

                    oNegocio = New Negocio.ManejoDeDocumentos(oCompany, Nothing, "S", ConfigurationManager.AppSettings("Tipo_Pch_Sap"))


                End If



            Catch ex As Exception
                GuardaLog("Error al instanciar la clase ManejoDeDocumentos: Error" + ex.Message.ToString)
                Exit Sub
            End Try
            '-------------------------

            Dim tipodocumento As String = ""

            For Each row As DataRow In objdr

                Dim oDocumentoTraza As New DocumentosTrans()

                oDocumentoTraza.Code = row("Code")
                oDocumentoTraza.DocEntry = row("U_DocEntry")
                oDocumentoTraza.DocSubType = row("U_DocSubType")
                oDocumentoTraza.SRI_Code = CStr(row("U_SRI_Code"))
                oDocumentoTraza.Oberva = ""
                prefijoDLL = ""

                Try
                    If oDocumentoTraza.SRI_Code.ToString().Equals("01") Then ' FACTURA

                        ' oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "FCE")

                        prefijoDLL = "FCE"
                        tipodocumento = "Factura"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("05") Then ' NOTA DE Debito

                        ' oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "NDE")

                        prefijoDLL = "NDE"
                        tipodocumento = "NOTA DEBITO"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("04") Then ' NOTA DE CREDITO

                        ' oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "NCE")

                        prefijoDLL = "NCE"
                        tipodocumento = "NOTA DE CREDITO"
                        ' GUIA DE REMISION - ENTREGA
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("15") Then

                        ' oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "GRE")
                        prefijoDLL = "GRE"
                        tipodocumento = "GUIA DE REMISION - ENTREGA"
                        ' GUIA DE REMISION- TRANSFERENCIA
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("67") Then

                        '   oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "TRE")
                        prefijoDLL = "TRE"
                        tipodocumento = "GUIA DE REMISION- TRANSFERENCIA"
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("1250000001") Then

                        ' oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "TLE")
                        prefijoDLL = "TLE"
                        tipodocumento = "GUIA DE SOLICITUD DE TRASLADO"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("07") Then

                        '  oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "REE")
                        prefijoDLL = "REE"
                        tipodocumento = "RETENCION"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("03") Or oDocumentoTraza.DocSubType.ToString().Equals("41") Then

                        ' oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(row("U_DocEntry"), "TLE")
                        prefijoDLL = "LQE"
                        tipodocumento = "LIQUIDACION"

                    End If



                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                        oDocumentoTraza.Oberva = oNegocioEcua.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, prefijoDLL, True)
                        modoSincro = "Webservices Sincro"
                    Else
                        'If postSincro = "1" Then
                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, prefijoDLL, True)
                        modoSincro = "Webservices Sincro"
                        'Else
                        'oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, prefijoDLL)
                        'modoSincro = "ReenvioDoc"
                        'End If
                    End If

                    GuardaLog(oDocumentoTraza.Oberva + "- Tipo de Documento : " + tipodocumento + " - DocEntry: " + oDocumentoTraza.DocEntry.ToString + " - Modo sincro: " + modoSincro)

                Catch ex As Exception

                    GuardaLog("trabajo_hilo_sincro excepcion,ProcesaEnvioDocumento - " + ex.Message.ToString())

                End Try

                oDocumentoTraza = Nothing

            Next

            oNegocio = Nothing
            oNegocioEcua = Nothing
        Catch ex As Exception
            GuardaLog("trabajo_hilo_sincro excepcion,ERROR GENERAL FUNCION - " + ex.Message.ToString())
        End Try



        GC.Collect()


    End Sub


    Private Sub GuardaLog(texto As String)

        If ConfigurationManager.AppSettings("GuardaLog").ToString().Equals("1") Then

            Dim nombreHilo = IIf(Thread.CurrentThread.Name Is Nothing, "SN", Thread.CurrentThread.Name)

            Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "ProcesadorFE_Log_" & Date.Now.ToString("dd_MM_yyyy") & " - " & nombreHilo.ToString & ".txt"
            'Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "ProcesadorFE_Log_" & Date.Now.ToString("dd_MM_yyyy") & ".txt"
            Dim sTexto As New StringBuilder

            sTexto.AppendLine("FECHA: " & Now)
            sTexto.AppendLine("----------------------------------------------------------")
            sTexto.AppendLine(texto.ToString())

            'SyncLock (miobj)

            Try
                Dim oTextWriter As TextWriter = New StreamWriter(sRuta, True)
                oTextWriter.WriteLine(sTexto.ToString)
                oTextWriter.Flush()
                oTextWriter.Close()
                oTextWriter = Nothing


            Catch ex As Exception
                ' EventLog.WriteEntry("MyWindowsService", "Error: " & ex.Message.ToString)
            Finally

            End Try

            'End SyncLock

        End If

    End Sub


#Region "Conexion SAP"



    Private Function conectSAP() As Boolean
        Try
            Dim ErrCode As Long
            Dim ErrMsg As String = ""
            GuardaLog(String.Format("Conexion SAP"))
            oCompany = New SAPbobsCOM.Company()

            'oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.DbServerType = ConfigurationManager.AppSettings("DevServerType")
            oCompany.UseTrusted = ConfigurationManager.AppSettings("UseTrusted")
            oCompany.CompanyDB = ConfigurationManager.AppSettings("DevDatabase")
            oCompany.UserName = ConfigurationManager.AppSettings("DevSBOUser")
            oCompany.Password = ConfigurationManager.AppSettings("DevSBOPassword")

            Try
                If CInt(ConfigurationManager.AppSettings("SAP_VERSION")) < 10 Then
                    oCompany.Server = ConfigurationManager.AppSettings("DevServer")
                    oCompany.LicenseServer = ConfigurationManager.AppSettings("LicenseServer")
                    oCompany.DbUserName = ConfigurationManager.AppSettings("DevDBUser")
                    oCompany.DbPassword = ConfigurationManager.AppSettings("DevDBPassword")
                Else
                    oCompany.Server = ConfigurationManager.AppSettings("DevServer")
                    'oCompany.SLDServer = ConfigurationManager.AppSettings("LicenseServer")
                End If
            Catch ex As Exception
                GuardaLog(ex.Message)
            End Try

            GuardaLog("DevServerType " + ConfigurationManager.AppSettings("DevServerType").ToString())
            GuardaLog("UseTrusted " + ConfigurationManager.AppSettings("UseTrusted").ToString())
            GuardaLog("DevDatabase " + ConfigurationManager.AppSettings("DevDatabase").ToString())
            GuardaLog("DevSBOUser " + ConfigurationManager.AppSettings("DevSBOUser").ToString())
            GuardaLog("DevSBOPassword " + ConfigurationManager.AppSettings("DevSBOPassword").ToString())
            GuardaLog("DevServer " + ConfigurationManager.AppSettings("DevServer").ToString())
            GuardaLog("LicenseServer " + ConfigurationManager.AppSettings("LicenseServer").ToString())

            If oCompany.Connected Then
                GuardaLog("Company en estado Conectado a SAP BO " + oCompany.CompanyDB.ToString())
                Return True
            End If



            ErrCode = oCompany.Connect()

            If ErrCode <> 0 Then
                oCompany.GetLastError(ErrCode, ErrMsg)
                GuardaLog("Error al conectarse a SAP ,funcion : oCompany.Connect :" + ErrCode.ToString() + " - " + ErrMsg.ToString)
                Return False
            Else
                ' GuardaLog("Conectado a SAP BO" + oCompany.CompanyDB.ToString())
                Return True
            End If

        Catch ex As Exception
            GuardaLog("Error al conectarse a SAP , funcion: conectSAP , EX :" + ex.Message.ToString())
            Return False
        End Try

    End Function
#End Region


#Region "Funciones de conexión ADO"


    Private Function conexionADOHANA() As Odbc.OdbcConnection

        Try
            Dim ConexionHana As String = String.Empty
            If (IntPtr.Size = 8) Then
                ConexionHana = ConexionHana & "Driver={HDBODBC};"
            Else
                ConexionHana = ConexionHana & "Driver={HDBODBC32};"
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DevServerNode")) Then
                ConexionHana = ConexionHana & "ServerNode=" & ConfigurationManager.AppSettings("DevServerNode") & ";"
                ConexionHana = ConexionHana & "UID=" & ConfigurationManager.AppSettings("DevDBUser") & ";"
                ConexionHana = ConexionHana & "PWD=" & ConfigurationManager.AppSettings("DevDBPassword") & ";"
                ConexionHana = ConexionHana & "CS=" & ConfigurationManager.AppSettings("DevDatabase")
            Else
                ConexionHana = ConexionHana & "ServerNode=" & ConfigurationManager.AppSettings("DevServer") & ";"
                ConexionHana = ConexionHana & "UID=" & ConfigurationManager.AppSettings("DevDBUser") & ";"
                ConexionHana = ConexionHana & "PWD=" & ConfigurationManager.AppSettings("DevDBPassword") & ";"
                ConexionHana = ConexionHana & "CS=" & ConfigurationManager.AppSettings("DevDatabase")
            End If

            'pswBD_HANA
            ' GuardaLog("Cadena de CONEXION_HANA : " + ConexionHana)
            Return New Odbc.OdbcConnection(ConexionHana)
        Catch ex As Exception
            GuardaLog("Error Al Instanciar la conexion HANA , funcion conexionADOHANA : " + ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Private Function conexionADOSQL() As SqlConnection
        Try
            Return New SqlConnection("Data Source = " & ConfigurationManager.AppSettings("DevServer").ToString() & "; Initial Catalog = " & ConfigurationManager.AppSettings("DevDatabase").ToString() & "; User Id=" & ConfigurationManager.AppSettings("DevDBUser") & ";Password=" & ConfigurationManager.AppSettings("DevDBPassword"))
        Catch ex As Exception
            GuardaLog("Error Al Instanciar la conexion SQL , funcion conexionADOHANA : " + ex.Message.ToString)
            Return Nothing
        End Try

    End Function

#End Region

#Region "Obtener Registros"

    Public Function GetDataTable(ByVal prmSQL As String, ByVal parametro As String) As DataTable

        Dim dtt As New DataTable
        Try
            If sServerType = 9 Then
                ' SI ES HANA

                Try

                    Using db As Odbc.OdbcConnection = conexionADOHANA()

                        db.Open()
                        prmSQL = String.Format("CALL {0}.{1}('{2}') ", ConfigurationManager.AppSettings("DevDatabase").ToString(), prmSQL, parametro)

                        GuardaLog("Consulta a la BD: " + prmSQL)

                        Using DapTable As New Odbc.OdbcDataAdapter(prmSQL, db)

                            DapTable.Fill(dtt)

                        End Using

                        db.Close()
                    End Using

                Catch ex As Exception
                    GuardaLog("Error al recuperar DataDable Session HANA, FUNCION GetDataTable, ex " & ex.Message)

                End Try


            Else

                Try
                    Using db As SqlConnection = conexionADOSQL()

                        db.Open()
                        prmSQL = String.Format("EXEC {0} '{1}'", prmSQL, parametro)
                        GuardaLog("Consulta a la BD: " + prmSQL)

                        Using DapTable As New SqlClient.SqlDataAdapter(prmSQL, db)

                            DapTable.Fill(dtt)

                        End Using

                        db.Close()

                    End Using

                Catch ex As Exception
                    GuardaLog("Error al recuperar DataDable Session SQL, FUNCION GetDataTable, ex " & ex.Message)
                End Try

            End If

        Catch ex As Exception
            GuardaLog("Error en la funcion principal GetDataTable, excepcion  " + ex.Message)

        End Try

        Return dtt

    End Function


#End Region

    Public Function ConsultaParametro(ByVal Modulo As String, ByVal Tipo As String, ByVal Subtipo As String, ByVal Nombre As String) As String
        Try
            Dim valor As String = ""
            Dim sQueryPrefijo As String = ""
            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryPrefijo = "SELECT A.""U_Valor"" "
                sQueryPrefijo += "FROM ""@GS_CONFD"" A INNER JOIN "
                sQueryPrefijo += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                sQueryPrefijo += " WHERE  B.""U_Modulo"" = '" + Modulo + "' AND B.""U_Tipo"" = '" + Tipo + "' "
                sQueryPrefijo += " AND B.""U_Subtipo"" = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.""U_Nombre"" = '" + Nombre + "'"
            Else
                sQueryPrefijo = "SELECT A.U_Valor "
                sQueryPrefijo += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                sQueryPrefijo += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                sQueryPrefijo += " WHERE B.U_Modulo = '" + Modulo + "' AND  B.U_Tipo = '" + Tipo + "' "
                sQueryPrefijo += " AND B.U_Subtipo = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.U_Nombre = '" + Nombre + "'"
            End If

            valor = getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#Region "Funciones Nuevas"


    'Funcion Comprobar Datos Mail

    Private Sub SeteaCamposMail()

        Try
            URL_SG = ConfigurationManager.AppSettings("URL_SG").ToString
        Catch ex As Exception
            URL_SG = ""
        End Try

        Try
            APIKEY_SG = ConfigurationManager.AppSettings("APIKEY_SG").ToString
        Catch ex As Exception
            APIKEY_SG = ""
        End Try

        Try
            CORREOASUNTO_SG = ConfigurationManager.AppSettings("CORREOASUNTO_SG").ToString
        Catch ex As Exception
            CORREOASUNTO_SG = ""
        End Try

        Try
            CORREOFROM_SG = ConfigurationManager.AppSettings("CORREOFROM_SG").ToString
        Catch ex As Exception
            CORREOFROM_SG = ""
        End Try

        Try
            CORREOTO_SG = ConfigurationManager.AppSettings("CORREOTO_SG").ToString
        Catch ex As Exception
            CORREOTO_SG = ""
        End Try

        Try
            MAXDOC_ALERTA = CInt(ConfigurationManager.AppSettings("MAXDOC_ALERTA").ToString)
        Catch ex As Exception
            MAXDOC_ALERTA = 0
        End Try

    End Sub


    'PROCESO ENVIO ALERTA

    Private Sub ProcesoEnvioAlerta(ByVal numRegistros As Integer)
        Try

            If URL_SG <> "" And APIKEY_SG <> "" And CORREOASUNTO_SG <> "" And CORREOFROM_SG <> "" And CORREOTO_SG <> "" _
                And numRegistros >= MAXDOC_ALERTA And MAXDOC_ALERTA > 0 Then

                Dim x As New SendGridMailInfo
                Dim key As String = APIKEY_SG.Trim, m As String = "", urlsg As String = URL_SG.Trim
                Dim listadestinatarios() As String = CORREOTO_SG.Split(";")
                x.Asunto = CORREOASUNTO_SG.Trim
                x.CuerpoMensaje = String.Format("Se Ha Detectado un Incremento en los Documentos a Ser Procesados : Numero de Documentos por Procesar {0}, La alarma se dispara al llegar a los {1} documentos", numRegistros.ToString, MAXDOC_ALERTA.ToString)
                x.CorreoDesde = New KeyValuePair(Of String, String)("Alarma Solsap360", CORREOFROM_SG.ToString.Trim)
                x.ListaDestinatarios = New List(Of String)

                For Each d In listadestinatarios

                    x.ListaDestinatarios.Add(d)

                Next

                If EnviarEmailSendGrid(urlsg, key, x, m) Then

                    GuardaLog("Mail de Alerta Enviado Correctamente!! Se disparo al llegar a  " & numRegistros.ToString & " documentos no Procesados")
                Else
                    GuardaLog("Error al tratar de Enviar Mail de Alerta , EX : " & m)
                End If


            End If

        Catch ex As Exception

            GuardaLog("Se genero una Excepcion en la funcion ProcesoEnvioAlerta , EX : " & ex.Message)

        End Try



    End Sub

    'FUNCION SENDGRID
    Private Function EnviarEmailSendGrid(ByVal URLWS As String, ByVal APIKEY As String, ByVal datosMail As SendGridMailInfo, Optional ByRef men As String = "") As Boolean

        Try
            Dim request = CType(WebRequest.Create(URLWS), HttpWebRequest)
            Dim postData As String = "", _personalizations As String = "", _from As String = "", _subject As String = "", _content As String = ""
            Dim destinatarios As New List(Of String)
            For Each destina In datosMail.ListaDestinatarios

                destinatarios.Add("{" & String.Format("""email"": ""{0}""", destina) & "}")

            Next

            _personalizations = """personalizations"": [{""to"": [" & String.Join(",", destinatarios.ToArray) & "]}]"

            _from = """from"": {""email"": """ & datosMail.CorreoDesde.Value & """,""name"": """ & datosMail.CorreoDesde.Key & """}"

            _subject = """subject"": """ & datosMail.Asunto & """"

            _content = """content"": [{""type"": ""text/plain"", ""value"": """ & datosMail.CuerpoMensaje.Replace("""", "").Replace("'", "") & """}]"

            postData = "{" & _personalizations & "," & _from & "," & _subject & "," & _content & "}"


            Dim data = Encoding.UTF8.GetBytes(postData)
            request.Headers.Add("authorization", "Bearer " & APIKEY.ToString().Trim())
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = data.Length
            Using stream = request.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using
            Dim response = CType(request.GetResponse(), HttpWebResponse)

            If response.StatusCode = HttpStatusCode.Accepted Then
                Dim menserver As String = ""
                menserver = New StreamReader(response.GetResponseStream()).ReadToEnd()
                If Not String.IsNullOrWhiteSpace(menserver) Then
                    men = menserver
                End If

                Return True

            End If


        Catch ex As Exception

            men = ex.Message

        End Try


        Return False

    End Function



#End Region


    Public Function getRSvalue(ByVal query As String, ByVal columnaRet As String, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRSvalue Catch:" + ex.Message().ToString() + "-QUERY: " + query, "FuncionesB1")
        End Try
        Return ret
    End Function


    Public Function getRecordSet(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
        End Try
        Return fRS
    End Function


    Public Function nzString(ByVal unString As String, Optional ByVal formatoSQL As Boolean = False, Optional ByVal valorSiNulo As String = "") As String
        Try
            If Not IsDBNull(unString) Then
                If formatoSQL Then
                    unString = unString.Replace("'", "' + CHAR(39) + '")
                End If
                valorSiNulo = unString
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("nzString Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
        Return valorSiNulo
    End Function

    Public Sub Release(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Release Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
    End Sub

    Private Function CheckLIC() As Boolean

        ' TIENE LICENCIA
        GuardaLog(String.Format("Validando la Licencia para Producto : {0} Version {1}", NombreServicio, VersionServicio))

        ''MANEJO LICENCIA WEB
        Dim wsSSLIC As New Entidades.wsSS_LICENCIA_SAP.Licencia

        Dim msgg As String = ""
        Dim oLicencia As LicenciaSS = Nothing
        Dim RucCliente As String = ""

        Try
            Dim param_ambiente As Integer = 0
            Dim param_inhouse As Boolean = False
            Dim respLIC As Entidades.wsSS_LICENCIA_SAP.CLsRespLic = Nothing

            Dim URLWS As String = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WsLicencia")
            wsSSLIC.Url = URLWS

            RucCliente = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RucCompañia")

            If String.IsNullOrEmpty(URLWS) Then
                GuardaLog("No se encontro parametrización de WS Control, verificar por favor!")
            End If

            If String.IsNullOrEmpty(RucCliente) Then
                RucCliente = "000000000"
            End If

            If Not String.IsNullOrEmpty(ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWsLicencia")) Then
                param_ambiente = IIf(ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWsLicencia") = "PRUEBAS", 1, 2)

            End If

            If (ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")) <> "LOCAL" Then
                param_inhouse = False
            Else
                param_inhouse = True
            End If

            Dim SApdatos As New Entidades.wsSS_LICENCIA_SAP.DatosSap
            With SApdatos
                .RucEmpresa = RucCliente
                .DireccionIPSERVER = oCompany.Server
                .NombreDB = oCompany.CompanyDB
                .NombreProducto = NombreServicio
                .VersionProducto = VersionServicio
                .Ambiente = param_ambiente
                .Inhouse = param_inhouse

            End With

            'If Not ConsultaParametro("eDoc", "PARAMETROS", "CONFIGURACION", "NO_ConsumirMetodoHTTPS") = "Y" Then
            '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
            'End If
            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Licencia" + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de Busqueda", "ServicioIvend")

                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsSS_LICENCIA_SAP.DatosSap))
                    Dim writer As TextWriter = New StreamWriter(sRuta)
                    x.Serialize(writer, SApdatos)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda" + sRuta, "ServicioIvend")
                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ServicioIvend")
            End Try

            SetProtocolosdeSeguridad()

            respLIC = wsSSLIC.ValidarLicencia(SApdatos, msgg)
            If Not IsNothing(respLIC) Then
                oLicencia = New LicenciaSS
                oLicencia.Opcion = respLIC.TipoLic
                oLicencia.NombreBaseSAP = oCompany.CompanyDB
                oLicencia.Estado = CBool(respLIC.Estado)
                oLicencia.validoHasta = 1000
                'oLicencia.VersionTributaria = respLIC.VersionTributaria
                listaParametros = respLIC.ListaUrlWS
            End If

            If oLicencia Is Nothing Then

                estadoLicencias = False
                versionTributaria = ""
                GuardaLog("Se Encontro un inconveniente al Consultar la Licencia, mensaje Referencia: " & msgg)
            Else

                estadoLicencias = oLicencia.Estado
                GuardaLog(String.Format("Respuesta desde el WS={0} ,Estado de Licencia={1},Tipo de Licencia={2}", msgg, oLicencia.Estado.ToString, oLicencia.Opcion.ToString))
            End If


            Return True
        Catch ex As Exception
            GuardaLog("GS Error de autenticacion con el WS " & ex.Message)
            oLicencia = Nothing
            Return False
        End Try
        ''END LICENCIA WEB
        Return False

    End Function

    Public Sub SetProtocolosdeSeguridad()



        'PARA TLS 1.2
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999



        'PARA HTTPS



        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)



    End Sub

    Function customCertValidation(ByVal sender As Object,
                                   ByVal cert As X509Certificate,
                                   ByVal chain As X509Chain,
                                   ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function
    'para las pruebas descomentar y colocar el proyecto servicio ivend como inicio
    'Public Sub New()

    '    ' Llamada necesaria para el diseñador.
    '    InitializeComponent()
    '    servicio_Core()

    '    ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    'End Sub
End Class

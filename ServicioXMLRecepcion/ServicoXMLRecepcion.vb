Imports System.Timers
Imports System.Text
Imports System.IO
Imports System.Configuration
Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Net
Imports System.Threading
Imports System.Xml
Imports System.Globalization
Imports System.Xml.Serialization


'Imports GS.EDOC.RAD.DAL
'Imports GS.EDOC.RAD.ENTIDADES
'Imports GS.EDOC.RAD.RPTVIEWER
'Imports GS.EDOC.RAD.BLL
'Imports GS.EDOC.ENTIDADES


Public Class ServicoXMLRecepcion


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

    Private Const NombreServicio As String = "SAED_XMLRecepcion"
    Private Const VersionServicio As String = "1.0"
    Private Const CodigoPais As String = "EC"

    Dim estadoLicencias As Boolean = False
    Dim versionTributaria As String = ""

    Dim RucCliente As String = ""

    Dim VarMonedaDolares As String = "###,###,###,###;##0.00"

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

        GuardaLog("Servicio XML Recepcion Iniciado..Procediendo a Conectarse")
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



            Try
                GETPArametros_INIT()
                Dim timer As Int64 = ConfigurationManager.AppSettings("timer") * 1000
                oTimer = New System.Timers.Timer(timer)
                'oTimerSincro = New System.Timers.Timer(timer)
                GuardaLog("Temporizador para lectura de XML iniciado: " + timer.ToString())
                oTimer.AutoReset = True
                oTimer.Enabled = False
                ' oTimerSincro.AutoReset = True
                ' oTimerSincro.Enabled = False

                AddHandler oTimer.Elapsed, AddressOf oTimer_Elapsed

                ' AddHandler oTimerSincro.Elapsed, AddressOf oTimerSincro_Elapsed

                oTimer.Start()
                'oTimerSincro.Start()

                GuardaLog("Temporizadores  Iniciados...")

                'descomentar para las pruebas
                'procesoUnHilo()
            Catch ex As ConfigurationErrorsException
                GuardaLog("Error al Inicializar Los Temporizadores: " + ex.Message.ToString())
            End Try


        Else

            GuardaLog("Error en la Conexion SAP , Validar el Archivo de Configuracion e Iniciar nuevamente el Servicio ")
        End If


    End Sub


    Private Sub GETPArametros_INIT()

        Functions.VariablesGlobales._SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")

    End Sub


    Private Sub oTimer_Elapsed()

        Try
            If blnInicio = False Then
                blnInicio = True
                GuardaLog("Ingresando al Hilo de Procesamiento de XML!!")
                Dim hilo As New Thread(AddressOf procesoUnHilo)
                hilo.Start()
                hilo.Join()
                GuardaLog("Retomando Control Hilo del Temporizador!!")
                blnInicio = False
            End If


        Catch ex As Exception
            GuardaLog("Error al tratar de Invocar la funcion de procesar_docEnvios en el Ciclo oTimer_Elapsed " + ex.Message.ToString())
        End Try

    End Sub


    Public Sub procesoUnHilo()


        Try
            ClearMemory()

            'Dim rutacompartida As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\FC"
            GuardaLog("Proceso Iniciado")

            HiloXML()

            GuardaLog("Proceso Terminado")

        Catch ex As Exception
            GuardaLog("Error en Funcion procesoOptimizado, EX  ==> " + ex.Message)
        End Try

    End Sub


    Private Declare Auto Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal procHandle As IntPtr, ByVal min As Int32, ByVal max As Int32) As Boolean
    'Funcion de liberacion de memoria
    Shared Function ClearMemory() As Boolean
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
                        'Control de errores
                    End Try
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
        End Try
        Return True
    End Function



    Private Sub trabajo_hilo_emision(objdr As List(Of DataRow))

        Try

            Dim oNegocio As Negocio.ManejoDeDocumentos
            Dim nombrehilo As String = ""
            nombrehilo = Thread.CurrentThread.Name
            Dim tipodocumento As String = ""

            GuardaLog(String.Format("Ejecutando funcion trabajo_hilo_emision,{0} documentos a Proce x hilo {1} ", objdr.Count.ToString, nombrehilo))

            Try

                oNegocio = New Negocio.ManejoDeDocumentos(oCompany, Nothing, "S", ConfigurationManager.AppSettings("Tipo_Pch_Sap"))
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


                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "FCE")

                        tipodocumento = "Factura"
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("05") Then ' NOTA DE DEBITO

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "NDE")

                        tipodocumento = "NOTA DEBITO"
                        ' GUIA DE REMISION - ENTREGA

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("04") Then ' NOTA DE CREDITO

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "NCE")

                        tipodocumento = "NOTA DE CREDITO"
                        ' GUIA DE REMISION - ENTREGA
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("15") Then

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "GRE")

                        tipodocumento = "GUIA DE REMISION - ENTREGA"
                        ' GUIA DE REMISION- TRANSFERENCIA
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("67") Then

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "TRE")

                        tipodocumento = "GUIA DE REMISION- TRANSFERENCIA"
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("06") And oDocumentoTraza.DocSubType.ToString().Equals("1250000001") Then

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "TLE")

                        tipodocumento = "GUIA DE SOLICITUD DE TRASLADO"

                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("07") Then

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "REE")

                        tipodocumento = "RETENCION"
                    ElseIf oDocumentoTraza.SRI_Code.ToString().Equals("03") Or oDocumentoTraza.DocSubType.ToString().Equals("41") Then

                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, "LQE")

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

        Catch ex As Exception
            GuardaLog("trabajo_hilo_emision excepcion,ERROR GENERAL FUNCION - " + ex.Message.ToString())
        End Try

        GC.Collect()


    End Sub


    Private Sub trabajo_hilo_sincro(objdr As List(Of DataRow))

        Try
            Dim oNegocio As Negocio.ManejoDeDocumentos = Nothing

            Dim nombrehilo As String = "", prefijoDLL As String = "", postSincro As String = "", modoSincro As String = ""
            nombrehilo = Thread.CurrentThread.Name
            'bandera para saber si se usara el metodo de envio por el WS de sincronizacion o emision
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PosSincro")) Then
                postSincro = ConfigurationManager.AppSettings("PosSincro")
            End If

            GuardaLog(String.Format("Ejecutando funcion trabajo_hilo_sincro,{0} documentos a Proce x hilo {1} ", objdr.Count.ToString, nombrehilo))

            Try

                oNegocio = New Negocio.ManejoDeDocumentos(oCompany, Nothing, "S", ConfigurationManager.AppSettings("Tipo_Pch_Sap"))

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

                    If postSincro = "1" Then
                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, prefijoDLL, True)
                        modoSincro = "Webservices Sincro"
                    Else
                        oDocumentoTraza.Oberva = oNegocio.ProcesaEnvioDocumento(oDocumentoTraza.DocEntry, prefijoDLL)
                        modoSincro = "ReenvioDoc"
                    End If

                    GuardaLog(oDocumentoTraza.Oberva + "- Tipo de Documento : " + tipodocumento + " - DocEntry: " + oDocumentoTraza.DocEntry.ToString + " - Modo sincro: " + modoSincro)

                Catch ex As Exception

                    GuardaLog("trabajo_hilo_sincro excepcion,ProcesaEnvioDocumento - " + ex.Message.ToString())

                End Try

                oDocumentoTraza = Nothing

            Next

            oNegocio = Nothing
        Catch ex As Exception
            GuardaLog("trabajo_hilo_sincro excepcion,ERROR GENERAL FUNCION - " + ex.Message.ToString())
        End Try



        GC.Collect()


    End Sub


    Public Sub GuardaLog(texto As String)

        If ConfigurationManager.AppSettings("GuardaLog").ToString().Equals("1") Then



            Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "ServicioRecepcionXML " & Date.Now.ToString("dd_MM_yyyy") & ".txt"

            Dim sTexto As New StringBuilder

            sTexto.AppendLine("FECHA: " & Now)
            sTexto.AppendLine("----------------------------------------------------------")
            sTexto.AppendLine(texto.ToString())

            SyncLock (miobj)

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

            End SyncLock

        End If

    End Sub

    Public Sub HiloXML()
        Dim nombreArchivo As String = ""

        Try

            'PROCESO FACTURA
            Dim rutafc = ConfigurationManager.AppSettings("RutaFC").ToString()
            Dim rutaProfc = ConfigurationManager.AppSettings("RutaProFC").ToString()
            Dim sArchivos As String() = Directory.GetFiles(rutafc)

            GuardaLog("XML Factura a procesar: " + sArchivos.Count.ToString)

            If sArchivos.Count > 0 Then

                For Each rutafc In sArchivos
                    nombreArchivo = Path.GetFileName(rutafc)
                    Dim extension = Path.GetExtension(rutafc)
                    If extension = ".xml" Then
                        Dim factura = LeerXMLFactura2(rutafc)


                        Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                        Dim sRuta As String = sRutaCarpeta & factura.FacturaCabecera._claveAcceso.ToString() + ".xml"
                        If System.IO.Directory.Exists(sRutaCarpeta) Then
                            Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")

                            Dim x As XmlSerializer = New XmlSerializer(GetType(Factura))
                            Dim writer As TextWriter = New StreamWriter(sRuta)
                            x.Serialize(writer, factura)
                            writer.Close()
                            Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")

                        End If

                        If insertarEntidadFC(factura, nombreArchivo) Then
                            MoverXML(rutafc, Path.GetFileName(rutafc), rutaProfc, ".xml")
                        End If

                    ElseIf extension = ".pdf" Then
                        MoverXML(rutafc, Path.GetFileName(rutafc), rutaProfc, ".pdf")
                    End If

                Next
                GuardaLog("Proceso Facturas termino con exito")
            End If

            'PROCESO NOTA CREDITO
            Dim rutanc = ConfigurationManager.AppSettings("RutaNC").ToString()
            Dim rutaPronc = ConfigurationManager.AppSettings("RutaProNC").ToString()
            Dim sArchivosNC As String() = Directory.GetFiles(rutanc)
            GuardaLog("XML Nota credito a procesar: " + sArchivosNC.Count.ToString)

            If sArchivosNC.Count > 0 Then



                For Each rutanc In sArchivosNC
                    nombreArchivo = Path.GetFileName(rutanc)

                    Dim extension = Path.GetExtension(rutanc)

                    If extension = ".xml" Then
                        Dim ncredito = LeerXMLNotaCredito2(rutanc)
                        If insertarEntidadNC(ncredito, Path.GetFileName(rutanc)) Then
                            MoverXML(rutanc, Path.GetFileName(rutanc), rutaPronc, ".xml")
                        End If

                    ElseIf extension = ".pdf" Then
                        MoverXML(rutanc, Path.GetFileName(rutanc), rutaPronc, ".pdf")

                    End If


                Next
                GuardaLog("Proceso notas credito termino con exito")
            End If


            'PROCESO RETENCION
            'Dim _ruta = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\RT\0304202207091630242500120051000000009711234567817_validarestructura.xml"
            Dim rutart = ConfigurationManager.AppSettings("RutaRT").ToString()
            Dim rutaPrort = ConfigurationManager.AppSettings("RutaProRT").ToString()
            'Dim mensaje As String = ""
            'Dim rete As DatosLeerXml = BLLLeerXml.LeerDatosRetencionXml(_ruta, Nothing, mensaje, Nothing, Nothing)

            Dim sArchivosRT As String() = Directory.GetFiles(rutart)
            GuardaLog("XML Retencion a procesar: " + sArchivosRT.Count.ToString)
            If sArchivosRT.Count > 0 Then

                For Each rutart In sArchivosRT
                    nombreArchivo = Path.GetFileName(rutart)

                    Dim extension = Path.GetExtension(rutart)
                    If extension = ".xml" Then
                        Dim retencion = LeerXMLRetencion2(rutart)
                        If insertarEntidadRT(retencion, Path.GetFileName(rutart)) Then
                            MoverXML(rutart, Path.GetFileName(rutart), rutaPrort, ".xml")
                        End If

                    ElseIf extension = ".pdf" Then
                        MoverXML(rutart, Path.GetFileName(rutart), rutaPrort, ".pdf")

                    End If


                Next
                GuardaLog("Proceso retencion termino con exito")
            End If

        Catch ex As Exception
            GuardaLog("Error al procesar XML: " + ex.Message.ToString + " con nombre: " + nombreArchivo.ToString)
        End Try


    End Sub
    Public Function LeerXMLFactura(ruta As String) As Factura

        GuardaLog("Consultando XML Facturas a procesar..")
        Try

            'Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\RFS.xml"

            Dim mensaje As String = ""

            Dim ms As New MemoryStream


            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList

            m_xmld = New XmlDocument()

            m_xmld.Load(ruta)

            '*******************Gs***************+++
            'Dim TipoDoc As Integer = BLLLeerXml.retornarTipoDocumento(RUTA, Nothing, mensaje)

            'Dim xml = BLLLeerXml.LeerDatosFacturaXml(RUTA, Nothing, mensaje, Nothing, Nothing)
            '******************FIN GS***************

            Dim Factura As New Factura
            Factura.FacturaCabecera = New FacturaCabecera

            Factura.FacturaCabecera._impuestos = New List(Of FacturaCabeceraImpuestos)

            Factura.facturaDetalle = New List(Of FacturaDetalle)




            Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")
            For Each facAut As XmlNode In nodoAut.ChildNodes
                Select Case facAut.Name
                    Case "numeroAutorizacion"
                        Factura.FacturaCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                    Case "fechaAutorizacion"
                        Factura.FacturaCabecera._FechaAutorizacion = CDate(facAut.InnerText)
                End Select
            Next

            Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

            Dim comprobante = nodo.InnerText

            Dim nodofactura As New XmlDocument()
            nodofactura.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

            Dim razonSocial As String = ""
            Dim ruc As String = ""
            Dim estab As String = ""
            Dim ptoEmi As String = ""
            Dim secuencial As String = ""


            For Each fac As XmlNode In nodofactura.ChildNodes
                Dim ad = fac.Name
                Select Case ad
                    Case "factura"
                        For Each nodoInfTri As XmlNode In fac.ChildNodes
                            Dim p = nodoInfTri.Name
                            Select Case p
                                Case "infoTributaria"
                                    For Each n As XmlNode In nodoInfTri.ChildNodes
                                        Select Case n.Name
                                            Case "razonSocial"
                                                razonSocial = n.InnerText.ToString
                                                Factura.FacturaCabecera._RazonSocial = razonSocial
                                            Case "ruc"
                                                ruc = n.InnerText.ToString
                                                Factura.FacturaCabecera._ruc = ruc
                                            Case "estab"
                                                estab = n.InnerText.ToString
                                                Factura.FacturaCabecera._estab = estab
                                            Case "ptoEmi"
                                                ptoEmi = n.InnerText.ToString
                                                Factura.FacturaCabecera._ptoEmi = ptoEmi
                                            Case "secuencial"
                                                secuencial = n.InnerText.ToString
                                                Factura.FacturaCabecera._secuencial = secuencial
                                            Case "claveAcceso"
                                                Factura.FacturaCabecera._claveAcceso = n.InnerText.ToString
                                        End Select

                                    Next
                                Case "infoFactura"
                                    For Each infoFac As XmlNode In nodoInfTri.ChildNodes
                                        Select Case infoFac.Name
                                            Case "fechaEmision"
                                                Dim fechaemision = infoFac.InnerText
                                                Factura.FacturaCabecera._fechaEmision = fechaemision
                                            Case "contribuyenteEspecial"
                                                Dim contribuyenteEspecial = infoFac.InnerText
                                                Factura.FacturaCabecera._contribuyenteEspecial = contribuyenteEspecial
                                            Case "dirEstablecimiento"
                                                Dim dirEstablecimiento = infoFac.InnerText
                                                Factura.FacturaCabecera._dirEstablecimiento = dirEstablecimiento
                                            Case "razonSocialComprador"
                                                Dim razonSocialComprador = infoFac.InnerText
                                                Factura.FacturaCabecera._razonSocialComprador = razonSocialComprador
                                            Case "identificacionComprador"
                                                Dim identificacionComprador = infoFac.InnerText
                                                Factura.FacturaCabecera._identificacionComprador = identificacionComprador
                                            Case "direccionComprador"
                                                Dim direccionComprador = infoFac.InnerText
                                                Factura.FacturaCabecera._direccionComprador = direccionComprador
                                            Case "totalSinImpuestos"
                                                Dim totalSinImpuestos = infoFac.InnerText
                                                Factura.FacturaCabecera._totalSinImpuestos = CDec(totalSinImpuestos)
                                            Case "totalDescuento"
                                                Dim totalDescuento = infoFac.InnerText
                                                Factura.FacturaCabecera._totalDescuento = CDec(totalDescuento)

                                            Case "totalConImpuestos"
                                                For Each totalConImpuestos As XmlNode In infoFac.ChildNodes
                                                    Select Case totalConImpuestos.Name
                                                        Case "totalImpuesto"
                                                            Dim facCabImp As New FacturaCabeceraImpuestos

                                                            For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                                Select Case totalImpuesto.Name
                                                                    Case "codigo"
                                                                        Dim codigo = totalImpuesto.InnerText
                                                                        facCabImp._codigo = CInt(codigo)
                                                                    Case "codigoPorcentaje"
                                                                        Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                        facCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                    Case "baseImponible"
                                                                        Dim baseImponible = totalImpuesto.InnerText
                                                                        facCabImp._baseImponible = CDec(baseImponible)
                                                                    Case "tarifa"
                                                                        Dim tarifa = totalImpuesto.InnerText
                                                                        facCabImp._tarifa = CDec(tarifa)
                                                                    Case "valor"
                                                                        Dim valor = totalImpuesto.InnerText
                                                                        facCabImp._valor = CDec(valor)
                                                                End Select

                                                            Next
                                                            Factura.FacturaCabecera._impuestos.Add(facCabImp)
                                                    End Select
                                                Next

                                            Case "importeTotal"
                                                Dim importeTotal = infoFac.InnerText
                                                Factura.FacturaCabecera._importeTotal = CDec(importeTotal)
                                            Case "moneda"
                                                Dim moneda = infoFac.InnerText

                                            Case "pagos"
                                                For Each pagos As XmlNode In infoFac.ChildNodes
                                                    Select Case pagos.Name
                                                        Case "pago"
                                                            For Each pago As XmlNode In pagos.ChildNodes
                                                                Select Case pago.Name
                                                                    Case "formaPago"
                                                                        Dim formaPago = pago.InnerText
                                                                        Factura.FacturaCabecera._formaPago = CInt(formaPago)
                                                                    Case "total"
                                                                        Dim total = pago.InnerText
                                                                        Factura.FacturaCabecera._totalFormaPago = CDec(total)
                                                                    Case "plazo"
                                                                        Dim plazo = pago.InnerText
                                                                        Factura.FacturaCabecera._plazo = CInt(plazo)
                                                                    Case "unidadTiempo"
                                                                        Dim unidadTiempo = pago.InnerText
                                                                        Factura.FacturaCabecera._unidadTiempo = unidadTiempo
                                                                End Select
                                                            Next
                                                    End Select
                                                Next
                                        End Select
                                    Next
                                Case "detalles"

                                    For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                        Select Case detalles.Name
                                            Case "detalle"
                                                Dim FacDet As New FacturaDetalle
                                                Dim FacDetImp As New FacturaDetalleImpuesto
                                                FacDet._impuestos = New List(Of FacturaDetalleImpuesto)
                                                For Each detalle As XmlNode In detalles.ChildNodes
                                                    Select Case detalle.Name
                                                        Case "codigoPrincipal"
                                                            Dim codigoPrincipal = detalle.InnerText
                                                            FacDet._codigoPrincipal = codigoPrincipal
                                                        Case "codigoAuxiliar"
                                                            Dim codigoAuxiliar = detalle.InnerText
                                                            FacDet._codigoAuxiliar = codigoAuxiliar
                                                        Case "descripcion"
                                                            Dim descripcion = detalle.InnerText
                                                            FacDet._descripcion = descripcion
                                                        Case "cantidad"
                                                            Dim cantidad = detalle.InnerText
                                                            FacDet._cantidad = CDec(cantidad)
                                                        Case "precioUnitario"
                                                            Dim precioUnitario = detalle.InnerText
                                                            FacDet._precioUnitario = CDec(precioUnitario)
                                                        Case "descuento"
                                                            Dim descuento = detalle.InnerText
                                                            FacDet._descuento = CDec(descuento)
                                                        Case "precioTotalSinImpuesto"
                                                            Dim precioTotalSinImpuesto = detalle.InnerText
                                                            FacDet._precioTotalSinImpuesto = CDec(precioTotalSinImpuesto)
                                                        Case "impuestos"
                                                            For Each impuestos As XmlNode In detalle.ChildNodes
                                                                Select Case impuestos.Name
                                                                    Case "impuesto"
                                                                        For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                            Select Case impuesto.Name
                                                                                Case "codigo"
                                                                                    Dim codigo = impuesto.InnerText
                                                                                    FacDetImp._codigo = CInt(codigo)
                                                                                Case "codigoPorcentaje"
                                                                                    Dim codigoPorcentaje = impuesto.InnerText
                                                                                    FacDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                Case "tarifa"
                                                                                    Dim tarifa = impuesto.InnerText
                                                                                    FacDetImp._tarifa = CDec(tarifa)
                                                                                Case "baseImponible"
                                                                                    Dim baseImponible = impuesto.InnerText
                                                                                    FacDetImp._baseImponible = CDec(baseImponible)
                                                                                Case "valor"
                                                                                    Dim valor = impuesto.InnerText
                                                                                    FacDetImp._valor = CDec(valor)
                                                                            End Select
                                                                        Next
                                                                        FacDet._impuestos.Add(FacDetImp)
                                                                End Select
                                                            Next
                                                    End Select
                                                Next
                                                Factura.facturaDetalle.Add(FacDet)
                                        End Select
                                    Next

                            End Select
                        Next
                End Select
            Next


            GuardaLog("Proceso Hilos Completado Correctamente!!")
            Return Factura
        Catch ex As Exception
            GuardaLog("Error al leer XML, EX  ==> " + ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function LeerXMLFactura2(ruta As String) As Factura

        GuardaLog("Consultando XML Facturas a procesar..")
        Try
            'Dim razonSocial As String = ""
            'Dim ruc As String = ""
            'Dim estab As String = ""
            'Dim ptoEmi As String = ""
            'Dim secuencial As String = ""

            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList

            Dim Factura As New Factura
            Factura.FacturaCabecera = New FacturaCabecera

            Factura.FacturaCabecera._impuestos = New List(Of FacturaCabeceraImpuestos)

            Factura.facturaDetalle = New List(Of FacturaDetalle)

            m_xmld = New XmlDocument()

            m_xmld.Load(ruta)

            Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")
            If (Not (nodoAut) Is Nothing) Then

                For Each facAut As XmlNode In nodoAut.ChildNodes
                    Select Case facAut.Name
                        Case "numeroAutorizacion"
                            Factura.FacturaCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                        Case "fechaAutorizacion"
                            Factura.FacturaCabecera._FechaAutorizacion = CDate(facAut.InnerText)
                    End Select
                Next

                Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

                Dim comprobante = nodo.InnerText

                Dim nodofactura As New XmlDocument()
                nodofactura.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable


                For Each fac As XmlNode In nodofactura.ChildNodes
                    Dim ad = fac.Name
                    Select Case ad
                        Case "factura"
                            For Each nodoInfTri As XmlNode In fac.ChildNodes
                                Dim p = nodoInfTri.Name
                                Select Case p
                                    Case "infoTributaria"
                                        For Each n As XmlNode In nodoInfTri.ChildNodes
                                            Select Case n.Name
                                                Case "razonSocial"
                                                    Dim razonSocial = n.InnerText.ToString
                                                    Factura.FacturaCabecera._RazonSocial = razonSocial
                                                Case "ruc"
                                                    Dim ruc = n.InnerText.ToString
                                                    Factura.FacturaCabecera._ruc = ruc
                                                Case "estab"
                                                    Dim estab = n.InnerText.ToString
                                                    Factura.FacturaCabecera._estab = estab
                                                Case "ptoEmi"
                                                    Dim ptoEmi = n.InnerText.ToString
                                                    Factura.FacturaCabecera._ptoEmi = ptoEmi
                                                Case "secuencial"
                                                    Dim secuencial = n.InnerText.ToString
                                                    Factura.FacturaCabecera._secuencial = secuencial
                                                Case "claveAcceso"
                                                    Factura.FacturaCabecera._claveAcceso = n.InnerText.ToString
                                            End Select

                                        Next
                                    Case "infoFactura"
                                        For Each infoFac As XmlNode In nodoInfTri.ChildNodes
                                            Select Case infoFac.Name
                                                Case "fechaEmision"
                                                    Dim fechaemision = infoFac.InnerText
                                                    Factura.FacturaCabecera._fechaEmision = fechaemision
                                                Case "contribuyenteEspecial"
                                                    Dim contribuyenteEspecial = infoFac.InnerText
                                                    Factura.FacturaCabecera._contribuyenteEspecial = contribuyenteEspecial
                                                Case "dirEstablecimiento"
                                                    Dim dirEstablecimiento = infoFac.InnerText
                                                    Factura.FacturaCabecera._dirEstablecimiento = dirEstablecimiento
                                                Case "razonSocialComprador"
                                                    Dim razonSocialComprador = infoFac.InnerText
                                                    Factura.FacturaCabecera._razonSocialComprador = razonSocialComprador
                                                Case "identificacionComprador"
                                                    Dim identificacionComprador = infoFac.InnerText
                                                    Factura.FacturaCabecera._identificacionComprador = identificacionComprador
                                                Case "direccionComprador"
                                                    Dim direccionComprador = infoFac.InnerText
                                                    Factura.FacturaCabecera._direccionComprador = direccionComprador
                                                Case "totalSinImpuestos"
                                                    Dim totalSinImpuestos = infoFac.InnerText
                                                    Factura.FacturaCabecera._totalSinImpuestos = totalSinImpuestos
                                                Case "totalDescuento"
                                                    Dim totalDescuento = infoFac.InnerText
                                                    Factura.FacturaCabecera._totalDescuento = totalDescuento 'CDec(totalDescuento)

                                                Case "totalConImpuestos"
                                                    For Each totalConImpuestos As XmlNode In infoFac.ChildNodes
                                                        Select Case totalConImpuestos.Name
                                                            Case "totalImpuesto"
                                                                Dim facCabImp As New FacturaCabeceraImpuestos

                                                                For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                                    Select Case totalImpuesto.Name
                                                                        Case "codigo"
                                                                            Dim codigo = totalImpuesto.InnerText
                                                                            facCabImp._codigo = CInt(codigo)
                                                                        Case "codigoPorcentaje"
                                                                            Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                            facCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                        Case "baseImponible"
                                                                            Dim baseImponible = totalImpuesto.InnerText
                                                                            facCabImp._baseImponible = baseImponible 'CDec(baseImponible)
                                                                        Case "tarifa"
                                                                            Dim tarifa = totalImpuesto.InnerText
                                                                            facCabImp._tarifa = tarifa 'CDec(tarifa)
                                                                        Case "valor"
                                                                            Dim valor = totalImpuesto.InnerText
                                                                            facCabImp._valor = valor 'CDec(valor)
                                                                    End Select

                                                                Next
                                                                Factura.FacturaCabecera._impuestos.Add(facCabImp)
                                                        End Select
                                                    Next

                                                Case "importeTotal"
                                                    Dim importeTotal = infoFac.InnerText
                                                    Factura.FacturaCabecera._importeTotal = importeTotal
                                                Case "moneda"
                                                    Dim moneda = infoFac.InnerText

                                                Case "pagos"
                                                    For Each pagos As XmlNode In infoFac.ChildNodes
                                                        Select Case pagos.Name
                                                            Case "pago"
                                                                For Each pago As XmlNode In pagos.ChildNodes
                                                                    Select Case pago.Name
                                                                        Case "formaPago"
                                                                            Dim formaPago = pago.InnerText
                                                                            Factura.FacturaCabecera._formaPago = formaPago
                                                                        Case "total"
                                                                            Dim total = pago.InnerText
                                                                            Factura.FacturaCabecera._totalFormaPago = total
                                                                        Case "plazo"
                                                                            Dim plazo = pago.InnerText
                                                                            Factura.FacturaCabecera._plazo = CInt(plazo)
                                                                        Case "unidadTiempo"
                                                                            Dim unidadTiempo = pago.InnerText
                                                                            Factura.FacturaCabecera._unidadTiempo = unidadTiempo
                                                                    End Select
                                                                Next
                                                        End Select
                                                    Next
                                            End Select
                                        Next
                                    Case "detalles"

                                        For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                            Select Case detalles.Name
                                                Case "detalle"
                                                    Dim FacDet As New FacturaDetalle
                                                    Dim FacDetImp As New FacturaDetalleImpuesto
                                                    FacDet._impuestos = New List(Of FacturaDetalleImpuesto)
                                                    For Each detalle As XmlNode In detalles.ChildNodes
                                                        Select Case detalle.Name
                                                            Case "codigoPrincipal"
                                                                Dim codigoPrincipal = detalle.InnerText
                                                                FacDet._codigoPrincipal = codigoPrincipal
                                                            Case "codigoAuxiliar"
                                                                Dim codigoAuxiliar = detalle.InnerText
                                                                FacDet._codigoAuxiliar = codigoAuxiliar
                                                            Case "descripcion"
                                                                Dim descripcion = detalle.InnerText
                                                                FacDet._descripcion = descripcion
                                                            Case "cantidad"
                                                                Dim cantidad = detalle.InnerText
                                                                FacDet._cantidad = cantidad 'String.Format(CultureInfo.InvariantCulture, "{0:N0}", cantidad) 'CDec(cantidad)
                                                            Case "precioUnitario"
                                                                Dim precioUnitario = detalle.InnerText
                                                                FacDet._precioUnitario = precioUnitario 'CDec(precioUnitario)
                                                            Case "descuento"
                                                                Dim descuento = detalle.InnerText
                                                                FacDet._descuento = descuento 'CDec(descuento)
                                                            Case "precioTotalSinImpuesto"
                                                                Dim precioTotalSinImpuesto = detalle.InnerText
                                                                FacDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                            Case "impuestos"
                                                                For Each impuestos As XmlNode In detalle.ChildNodes
                                                                    Select Case impuestos.Name
                                                                        Case "impuesto"
                                                                            For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                                Select Case impuesto.Name
                                                                                    Case "codigo"
                                                                                        Dim codigo = impuesto.InnerText
                                                                                        FacDetImp._codigo = CInt(codigo)
                                                                                    Case "codigoPorcentaje"
                                                                                        Dim codigoPorcentaje = impuesto.InnerText
                                                                                        FacDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                    Case "tarifa"
                                                                                        Dim tarifa = impuesto.InnerText
                                                                                        FacDetImp._tarifa = tarifa 'CDec(tarifa)
                                                                                    Case "baseImponible"
                                                                                        Dim baseImponible = impuesto.InnerText
                                                                                        FacDetImp._baseImponible = baseImponible 'CDec(baseImponible)
                                                                                    Case "valor"
                                                                                        Dim valor = impuesto.InnerText
                                                                                        FacDetImp._valor = valor 'CDec(valor)
                                                                                End Select
                                                                            Next
                                                                            FacDet._impuestos.Add(FacDetImp)
                                                                    End Select
                                                                Next
                                                        End Select
                                                    Next
                                                    Factura.facturaDetalle.Add(FacDet)
                                            End Select
                                        Next

                                End Select
                            Next
                    End Select
                Next

            Else

                Dim nodoFac = m_xmld.SelectSingleNode("factura")

                For Each fac As XmlNode In nodoFac.ChildNodes
                    Dim p = fac.Name
                    Select Case p
                        Case "infoTributaria"
                            For Each n As XmlNode In fac.ChildNodes
                                Select Case n.Name
                                    Case "razonSocial"
                                        Dim razonSocial = n.InnerText.ToString
                                        Factura.FacturaCabecera._RazonSocial = razonSocial
                                    Case "ruc"
                                        Dim ruc = n.InnerText.ToString
                                        Factura.FacturaCabecera._ruc = ruc
                                    Case "estab"
                                        Dim estab = n.InnerText.ToString
                                        Factura.FacturaCabecera._estab = estab
                                    Case "ptoEmi"
                                        Dim ptoEmi = n.InnerText.ToString
                                        Factura.FacturaCabecera._ptoEmi = ptoEmi
                                    Case "secuencial"
                                        Dim secuencial = n.InnerText.ToString
                                        Factura.FacturaCabecera._secuencial = secuencial
                                    Case "claveAcceso"
                                        Factura.FacturaCabecera._claveAcceso = n.InnerText.ToString
                                End Select

                            Next
                        Case "infoFactura"
                            For Each infoFac As XmlNode In fac.ChildNodes
                                Select Case infoFac.Name
                                    Case "fechaEmision"
                                        Dim fechaemision = infoFac.InnerText
                                        Factura.FacturaCabecera._fechaEmision = fechaemision
                                    Case "contribuyenteEspecial"
                                        Dim contribuyenteEspecial = infoFac.InnerText
                                        Factura.FacturaCabecera._contribuyenteEspecial = contribuyenteEspecial
                                    Case "dirEstablecimiento"
                                        Dim dirEstablecimiento = infoFac.InnerText
                                        Factura.FacturaCabecera._dirEstablecimiento = dirEstablecimiento
                                    Case "razonSocialComprador"
                                        Dim razonSocialComprador = infoFac.InnerText
                                        Factura.FacturaCabecera._razonSocialComprador = razonSocialComprador
                                    Case "identificacionComprador"
                                        Dim identificacionComprador = infoFac.InnerText
                                        Factura.FacturaCabecera._identificacionComprador = identificacionComprador
                                    Case "direccionComprador"
                                        Dim direccionComprador = infoFac.InnerText
                                        Factura.FacturaCabecera._direccionComprador = direccionComprador
                                    Case "totalSinImpuestos"
                                        Dim totalSinImpuestos = infoFac.InnerText
                                        Factura.FacturaCabecera._totalSinImpuestos = totalSinImpuestos
                                    Case "totalDescuento"
                                        Dim totalDescuento = infoFac.InnerText
                                        Factura.FacturaCabecera._totalDescuento = totalDescuento

                                    Case "totalConImpuestos"
                                        For Each totalConImpuestos As XmlNode In infoFac.ChildNodes
                                            Select Case totalConImpuestos.Name
                                                Case "totalImpuesto"
                                                    Dim facCabImp As New FacturaCabeceraImpuestos

                                                    For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                        Select Case totalImpuesto.Name
                                                            Case "codigo"
                                                                Dim codigo = totalImpuesto.InnerText
                                                                facCabImp._codigo = CInt(codigo)
                                                            Case "codigoPorcentaje"
                                                                Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                facCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                            Case "baseImponible"
                                                                Dim baseImponible = totalImpuesto.InnerText
                                                                facCabImp._baseImponible = baseImponible
                                                            Case "tarifa"
                                                                Dim tarifa = totalImpuesto.InnerText
                                                                facCabImp._tarifa = tarifa
                                                            Case "valor"
                                                                Dim valor = totalImpuesto.InnerText
                                                                facCabImp._valor = valor
                                                        End Select

                                                    Next
                                                    Factura.FacturaCabecera._impuestos.Add(facCabImp)
                                            End Select
                                        Next

                                    Case "importeTotal"
                                        Dim importeTotal = infoFac.InnerText
                                        Factura.FacturaCabecera._importeTotal = importeTotal
                                    Case "moneda"
                                        Dim moneda = infoFac.InnerText

                                    Case "pagos"
                                        For Each pagos As XmlNode In infoFac.ChildNodes
                                            Select Case pagos.Name
                                                Case "pago"
                                                    For Each pago As XmlNode In pagos.ChildNodes
                                                        Select Case pago.Name
                                                            Case "formaPago"
                                                                Dim formaPago = pago.InnerText
                                                                Factura.FacturaCabecera._formaPago = formaPago
                                                            Case "total"
                                                                Dim total = pago.InnerText
                                                                Factura.FacturaCabecera._totalFormaPago = total
                                                            Case "plazo"
                                                                Dim plazo = pago.InnerText
                                                                Factura.FacturaCabecera._plazo = CInt(plazo)
                                                            Case "unidadTiempo"
                                                                Dim unidadTiempo = pago.InnerText
                                                                Factura.FacturaCabecera._unidadTiempo = unidadTiempo
                                                        End Select
                                                    Next
                                            End Select
                                        Next
                                End Select
                            Next
                        Case "detalles"

                            For Each detalles As XmlNode In fac.ChildNodes
                                Select Case detalles.Name
                                    Case "detalle"
                                        Dim FacDet As New FacturaDetalle
                                        Dim FacDetImp As New FacturaDetalleImpuesto
                                        FacDet._impuestos = New List(Of FacturaDetalleImpuesto)
                                        For Each detalle As XmlNode In detalles.ChildNodes
                                            Select Case detalle.Name
                                                Case "codigoPrincipal"
                                                    Dim codigoPrincipal = detalle.InnerText
                                                    FacDet._codigoPrincipal = codigoPrincipal
                                                Case "codigoAuxiliar"
                                                    Dim codigoAuxiliar = detalle.InnerText
                                                    FacDet._codigoAuxiliar = codigoAuxiliar
                                                Case "descripcion"
                                                    Dim descripcion = detalle.InnerText
                                                    FacDet._descripcion = descripcion
                                                Case "cantidad"
                                                    Dim cantidad = detalle.InnerText
                                                    FacDet._cantidad = cantidad
                                                Case "precioUnitario"
                                                    Dim precioUnitario = detalle.InnerText
                                                    FacDet._precioUnitario = precioUnitario
                                                Case "descuento"
                                                    Dim descuento = detalle.InnerText
                                                    FacDet._descuento = descuento
                                                Case "precioTotalSinImpuesto"
                                                    Dim precioTotalSinImpuesto = detalle.InnerText
                                                    FacDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                Case "impuestos"
                                                    For Each impuestos As XmlNode In detalle.ChildNodes
                                                        Select Case impuestos.Name
                                                            Case "impuesto"
                                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                    Select Case impuesto.Name
                                                                        Case "codigo"
                                                                            Dim codigo = impuesto.InnerText
                                                                            FacDetImp._codigo = CInt(codigo)
                                                                        Case "codigoPorcentaje"
                                                                            Dim codigoPorcentaje = impuesto.InnerText
                                                                            FacDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                        Case "tarifa"
                                                                            Dim tarifa = impuesto.InnerText
                                                                            FacDetImp._tarifa = tarifa
                                                                        Case "baseImponible"
                                                                            Dim baseImponible = impuesto.InnerText
                                                                            FacDetImp._baseImponible = baseImponible
                                                                        Case "valor"
                                                                            Dim valor = impuesto.InnerText
                                                                            FacDetImp._valor = valor
                                                                    End Select
                                                                Next
                                                                FacDet._impuestos.Add(FacDetImp)
                                                        End Select
                                                    Next
                                            End Select
                                        Next
                                        Factura.facturaDetalle.Add(FacDet)
                                End Select
                            Next

                    End Select
                Next

            End If


            GuardaLog("FC clave de acceso: " + Factura.FacturaCabecera._claveAcceso + " leido Correctamente!!" + " - nombre del archivo: " + Path.GetFileName(ruta))

            Return Factura
        Catch ex As Exception
            GuardaLog("Error al leer XML, EX  ==> " + ex.Message + " con nombre: " + Path.GetFileName(ruta))
            Return Nothing
        End Try

    End Function

    Public Function LeerXMLNotaCredito(ruta As String) As NotaCredito

        GuardaLog("Consultando XML Notas de Credito a procesar..")

        Try

            Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\ncprueba.xml"

            Dim mensaje As String = ""
            Dim m_xmld As XmlDocument
            m_xmld = New XmlDocument()

            m_xmld.Load(ruta)


            Dim Nc As New NotaCredito
            Nc.NotaCreditoCabecera = New NotaCreditoCabecera

            Nc.NotaCreditoCabecera._impuestos = New List(Of NotaCreditoCabeceraImpuesto)

            Nc.NotaCreditoDetalle = New List(Of NotaCreditoDetalle)




            Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")
            For Each NcAut As XmlNode In nodoAut.ChildNodes
                Select Case NcAut.Name
                    Case "numeroAutorizacion"
                        Nc.NotaCreditoCabecera._NumeroAutorizacion = NcAut.InnerText.ToString
                    Case "fechaAutorizacion"
                        Nc.NotaCreditoCabecera._FechaAutorizacion = CDate(NcAut.InnerText)
                End Select
            Next

            Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

            Dim comprobante = nodo.InnerText

            Dim nodoNc As New XmlDocument()
            nodoNc.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

            Dim razonSocial As String = ""
            Dim ruc As String = ""
            Dim estab As String = ""
            Dim ptoEmi As String = ""
            Dim secuencial As String = ""


            For Each notac As XmlNode In nodoNc.ChildNodes
                Dim ad = notac.Name
                Select Case ad
                    Case "notaCredito"
                        For Each nodoInfTri As XmlNode In notac.ChildNodes
                            Dim p = nodoInfTri.Name
                            Select Case p
                                Case "infoTributaria"
                                    For Each n As XmlNode In nodoInfTri.ChildNodes
                                        Select Case n.Name
                                            Case "razonSocial"
                                                razonSocial = n.InnerText.ToString
                                                Nc.NotaCreditoCabecera._RazonSocial = razonSocial
                                            Case "ruc"
                                                ruc = n.InnerText.ToString
                                                Nc.NotaCreditoCabecera._ruc = ruc
                                            Case "estab"
                                                estab = n.InnerText.ToString
                                                Nc.NotaCreditoCabecera._estab = estab
                                            Case "ptoEmi"
                                                ptoEmi = n.InnerText.ToString
                                                Nc.NotaCreditoCabecera._ptoEmi = ptoEmi
                                            Case "secuencial"
                                                secuencial = n.InnerText.ToString
                                                Nc.NotaCreditoCabecera._secuencial = secuencial
                                            Case "claveAcceso"
                                                Dim claveAcceso = n.InnerText.ToString
                                                Nc.NotaCreditoCabecera._claveAcceso = claveAcceso
                                        End Select

                                    Next
                                Case "infoNotaCredito"
                                    For Each infoNc As XmlNode In nodoInfTri.ChildNodes
                                        Select Case infoNc.Name
                                            Case "fechaEmision"
                                                Dim fechaemision = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._fechaEmision = fechaemision
                                            Case "dirEstablecimiento"
                                                Dim dirEstablecimiento = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._dirEstablecimiento = dirEstablecimiento
                                            Case "razonSocialComprador"
                                                Dim razonSocialComprador = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._razonSocialComprador = razonSocialComprador
                                            Case "identificacionComprador"
                                                Dim identificacionComprador = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._identificacionComprador = identificacionComprador
                                            Case "direccionComprador"
                                                Dim direccionComprador = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._direccionComprador = direccionComprador
                                            Case "codDocModificado"
                                                Dim codDocModificado = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._CodDocMod = codDocModificado
                                            Case "numDocModificado"
                                                Dim numDocModificado = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._numDocModificado = numDocModificado
                                            Case "fechaEmisionDocSustento"
                                                Dim fechaEmisionDocSustento = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._fechaEmisionDocSustento = CDate(fechaEmisionDocSustento)
                                            Case "totalSinImpuestos"
                                                Dim totalSinImpuestos = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._totalSinImpuestos = CDec(totalSinImpuestos)
                                            Case "motivo"
                                                Dim motivo = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._motivo = motivo
                                            Case "valorModificacion"
                                                Dim valorModificacion = infoNc.InnerText
                                                Nc.NotaCreditoCabecera._valorModificacion = CDec(valorModificacion)
                                            Case "totalConImpuestos"
                                                For Each totalConImpuestos As XmlNode In infoNc.ChildNodes
                                                    Select Case totalConImpuestos.Name
                                                        Case "totalImpuesto"
                                                            Dim NcCabImp As New NotaCreditoCabeceraImpuesto
                                                            For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                                Select Case totalImpuesto.Name
                                                                    Case "codigo"
                                                                        Dim codigo = totalImpuesto.InnerText
                                                                        NcCabImp._codigo = CInt(codigo)
                                                                    Case "codigoPorcentaje"
                                                                        Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                        NcCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                    Case "baseImponible"
                                                                        Dim baseImponible = totalImpuesto.InnerText
                                                                        NcCabImp._baseImponible = CDec(baseImponible)
                                                                    Case "tarifa"
                                                                        Dim tarifa = totalImpuesto.InnerText
                                                                        NcCabImp._tarifa = CDec(tarifa)
                                                                    Case "valor"
                                                                        Dim valor = totalImpuesto.InnerText
                                                                        NcCabImp._valor = CDec(valor)
                                                                End Select

                                                            Next
                                                            Nc.NotaCreditoCabecera._impuestos.Add(NcCabImp)
                                                    End Select
                                                Next


                                        End Select
                                    Next
                                Case "detalles"

                                    For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                        Select Case detalles.Name
                                            Case "detalle"
                                                Dim NcDet As New NotaCreditoDetalle
                                                Dim NcDetImp As New NotaCreditoDetalleImpuesto
                                                NcDet._impuestos = New List(Of NotaCreditoDetalleImpuesto)
                                                For Each detalle As XmlNode In detalles.ChildNodes
                                                    Select Case detalle.Name
                                                        Case "codigoInterno"
                                                            Dim codigoPrincipal = detalle.InnerText
                                                            NcDet._codigoInterno = codigoPrincipal
                                                        Case "codigoAdicional"
                                                            Dim codigoAuxiliar = detalle.InnerText
                                                            NcDet._codigoAdicional = codigoAuxiliar
                                                        Case "descripcion"
                                                            Dim descripcion = detalle.InnerText
                                                            NcDet._descripcion = descripcion
                                                        Case "cantidad"
                                                            Dim cantidad = detalle.InnerText
                                                            NcDet._cantidad = CDec(cantidad)
                                                        Case "precioUnitario"
                                                            Dim precioUnitario = detalle.InnerText
                                                            NcDet._precioUnitario = CDec(precioUnitario)
                                                        Case "descuento"
                                                            Dim descuento = detalle.InnerText
                                                            NcDet._descuento = CDec(descuento)
                                                        Case "precioTotalSinImpuesto"
                                                            Dim precioTotalSinImpuesto = detalle.InnerText
                                                            NcDet._precioTotalSinImpuesto = CDec(precioTotalSinImpuesto)
                                                        Case "impuestos"
                                                            For Each impuestos As XmlNode In detalle.ChildNodes
                                                                Select Case impuestos.Name
                                                                    Case "impuesto"
                                                                        For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                            Select Case impuesto.Name
                                                                                Case "codigo"
                                                                                    Dim codigo = impuesto.InnerText
                                                                                    NcDetImp._codigo = CInt(codigo)
                                                                                Case "codigoPorcentaje"
                                                                                    Dim codigoPorcentaje = impuesto.InnerText
                                                                                    NcDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                Case "tarifa"
                                                                                    Dim tarifa = impuesto.InnerText
                                                                                    NcDetImp._tarifa = CDec(tarifa)
                                                                                Case "baseImponible"
                                                                                    Dim baseImponible = impuesto.InnerText
                                                                                    NcDetImp._baseImponible = CDec(baseImponible)
                                                                                Case "valor"
                                                                                    Dim valor = impuesto.InnerText
                                                                                    NcDetImp._valor = CDec(valor)
                                                                            End Select
                                                                        Next
                                                                        NcDet._impuestos.Add(NcDetImp)
                                                                End Select
                                                            Next
                                                    End Select
                                                Next
                                                Nc.NotaCreditoDetalle.Add(NcDet)
                                        End Select
                                    Next

                            End Select
                        Next
                End Select
            Next


            GuardaLog("XML NC Leido correctamente clave: " + Nc.NotaCreditoCabecera._claveAcceso + " con nombre: " + Path.GetFileName(ruta))
            Return Nc
        Catch ex As Exception
            GuardaLog("Error al leer XML, EX  ==> " + ex.Message + " con nombre: " + Path.GetFileName(ruta))
            Return Nothing
        End Try

    End Function

    Public Function LeerXMLNotaCredito2(ruta As String) As NotaCredito

        GuardaLog("Consultando XML Notas de Credito a procesar..")

        Try

            Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\ncprueba.xml"

            Dim mensaje As String = ""
            Dim m_xmld As XmlDocument
            m_xmld = New XmlDocument()

            m_xmld.Load(ruta)


            Dim Nc As New NotaCredito
            Nc.NotaCreditoCabecera = New NotaCreditoCabecera

            Nc.NotaCreditoCabecera._impuestos = New List(Of NotaCreditoCabeceraImpuesto)

            Nc.NotaCreditoDetalle = New List(Of NotaCreditoDetalle)




            Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")

            If (Not (nodoAut) Is Nothing) Then

                For Each NcAut As XmlNode In nodoAut.ChildNodes
                    Select Case NcAut.Name
                        Case "numeroAutorizacion"
                            Nc.NotaCreditoCabecera._NumeroAutorizacion = NcAut.InnerText.ToString
                        Case "fechaAutorizacion"
                            Nc.NotaCreditoCabecera._FechaAutorizacion = CDate(NcAut.InnerText)
                    End Select
                Next

                Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

                Dim comprobante = nodo.InnerText

                Dim nodoNc As New XmlDocument()
                nodoNc.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

                Dim razonSocial As String = ""
                Dim ruc As String = ""
                Dim estab As String = ""
                Dim ptoEmi As String = ""
                Dim secuencial As String = ""


                For Each notac As XmlNode In nodoNc.ChildNodes
                    Dim ad = notac.Name
                    Select Case ad
                        Case "notaCredito"
                            For Each nodoInfTri As XmlNode In notac.ChildNodes
                                Dim p = nodoInfTri.Name
                                Select Case p
                                    Case "infoTributaria"
                                        For Each n As XmlNode In nodoInfTri.ChildNodes
                                            Select Case n.Name
                                                Case "razonSocial"
                                                    razonSocial = n.InnerText.ToString
                                                    Nc.NotaCreditoCabecera._RazonSocial = razonSocial
                                                Case "ruc"
                                                    ruc = n.InnerText.ToString
                                                    Nc.NotaCreditoCabecera._ruc = ruc
                                                Case "estab"
                                                    estab = n.InnerText.ToString
                                                    Nc.NotaCreditoCabecera._estab = estab
                                                Case "ptoEmi"
                                                    ptoEmi = n.InnerText.ToString
                                                    Nc.NotaCreditoCabecera._ptoEmi = ptoEmi
                                                Case "secuencial"
                                                    secuencial = n.InnerText.ToString
                                                    Nc.NotaCreditoCabecera._secuencial = secuencial
                                                Case "claveAcceso"
                                                    Dim claveAcceso = n.InnerText.ToString
                                                    Nc.NotaCreditoCabecera._claveAcceso = claveAcceso
                                            End Select

                                        Next
                                    Case "infoNotaCredito"
                                        For Each infoNc As XmlNode In nodoInfTri.ChildNodes
                                            Select Case infoNc.Name
                                                Case "fechaEmision"
                                                    Dim fechaemision = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._fechaEmision = fechaemision
                                                Case "dirEstablecimiento"
                                                    Dim dirEstablecimiento = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._dirEstablecimiento = dirEstablecimiento
                                                Case "razonSocialComprador"
                                                    Dim razonSocialComprador = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._razonSocialComprador = razonSocialComprador
                                                Case "identificacionComprador"
                                                    Dim identificacionComprador = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._identificacionComprador = identificacionComprador
                                                Case "direccionComprador"
                                                    Dim direccionComprador = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._direccionComprador = direccionComprador
                                                Case "codDocModificado"
                                                    Dim codDocModificado = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._CodDocMod = codDocModificado
                                                Case "numDocModificado"
                                                    Dim numDocModificado = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._numDocModificado = numDocModificado
                                                Case "fechaEmisionDocSustento"
                                                    Dim fechaEmisionDocSustento = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._fechaEmisionDocSustento = CDate(fechaEmisionDocSustento)
                                                Case "totalSinImpuestos"
                                                    Dim totalSinImpuestos = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._totalSinImpuestos = totalSinImpuestos
                                                Case "motivo"
                                                    Dim motivo = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._motivo = motivo
                                                Case "valorModificacion"
                                                    Dim valorModificacion = infoNc.InnerText
                                                    Nc.NotaCreditoCabecera._valorModificacion = valorModificacion
                                                Case "totalConImpuestos"
                                                    For Each totalConImpuestos As XmlNode In infoNc.ChildNodes
                                                        Select Case totalConImpuestos.Name
                                                            Case "totalImpuesto"
                                                                Dim NcCabImp As New NotaCreditoCabeceraImpuesto
                                                                For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                                    Select Case totalImpuesto.Name
                                                                        Case "codigo"
                                                                            Dim codigo = totalImpuesto.InnerText
                                                                            NcCabImp._codigo = CInt(codigo)
                                                                        Case "codigoPorcentaje"
                                                                            Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                            NcCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                        Case "baseImponible"
                                                                            Dim baseImponible = totalImpuesto.InnerText
                                                                            NcCabImp._baseImponible = baseImponible
                                                                        Case "tarifa"
                                                                            Dim tarifa = totalImpuesto.InnerText
                                                                            NcCabImp._tarifa = tarifa
                                                                        Case "valor"
                                                                            Dim valor = totalImpuesto.InnerText
                                                                            NcCabImp._valor = valor
                                                                    End Select

                                                                Next
                                                                Nc.NotaCreditoCabecera._impuestos.Add(NcCabImp)
                                                        End Select
                                                    Next


                                            End Select
                                        Next
                                    Case "detalles"

                                        For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                            Select Case detalles.Name
                                                Case "detalle"
                                                    Dim NcDet As New NotaCreditoDetalle
                                                    Dim NcDetImp As New NotaCreditoDetalleImpuesto
                                                    NcDet._impuestos = New List(Of NotaCreditoDetalleImpuesto)
                                                    For Each detalle As XmlNode In detalles.ChildNodes
                                                        Select Case detalle.Name
                                                            Case "codigoInterno"
                                                                Dim codigoPrincipal = detalle.InnerText
                                                                NcDet._codigoInterno = codigoPrincipal
                                                            Case "codigoAdicional"
                                                                Dim codigoAuxiliar = detalle.InnerText
                                                                NcDet._codigoAdicional = codigoAuxiliar
                                                            Case "descripcion"
                                                                Dim descripcion = detalle.InnerText
                                                                NcDet._descripcion = descripcion
                                                            Case "cantidad"
                                                                Dim cantidad = detalle.InnerText
                                                                NcDet._cantidad = cantidad
                                                            Case "precioUnitario"
                                                                Dim precioUnitario = detalle.InnerText
                                                                NcDet._precioUnitario = precioUnitario
                                                            Case "descuento"
                                                                Dim descuento = detalle.InnerText
                                                                NcDet._descuento = descuento
                                                            Case "precioTotalSinImpuesto"
                                                                Dim precioTotalSinImpuesto = detalle.InnerText
                                                                NcDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                            Case "impuestos"
                                                                For Each impuestos As XmlNode In detalle.ChildNodes
                                                                    Select Case impuestos.Name
                                                                        Case "impuesto"
                                                                            For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                                Select Case impuesto.Name
                                                                                    Case "codigo"
                                                                                        Dim codigo = impuesto.InnerText
                                                                                        NcDetImp._codigo = CInt(codigo)
                                                                                    Case "codigoPorcentaje"
                                                                                        Dim codigoPorcentaje = impuesto.InnerText
                                                                                        NcDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                    Case "tarifa"
                                                                                        Dim tarifa = impuesto.InnerText
                                                                                        NcDetImp._tarifa = tarifa
                                                                                    Case "baseImponible"
                                                                                        Dim baseImponible = impuesto.InnerText
                                                                                        NcDetImp._baseImponible = baseImponible
                                                                                    Case "valor"
                                                                                        Dim valor = impuesto.InnerText
                                                                                        NcDetImp._valor = valor
                                                                                End Select
                                                                            Next
                                                                            NcDet._impuestos.Add(NcDetImp)
                                                                    End Select
                                                                Next
                                                        End Select
                                                    Next
                                                    Nc.NotaCreditoDetalle.Add(NcDet)
                                            End Select
                                        Next

                                End Select
                            Next
                    End Select
                Next

            Else


                Dim notaCredito = m_xmld.SelectSingleNode("notaCredito")

                For Each nodoInfTri As XmlNode In notaCredito.ChildNodes
                    Dim p = nodoInfTri.Name
                    Select Case p
                        Case "infoTributaria"
                            For Each n As XmlNode In nodoInfTri.ChildNodes
                                Select Case n.Name
                                    Case "razonSocial"
                                        Dim razonSocial = n.InnerText.ToString
                                        Nc.NotaCreditoCabecera._RazonSocial = razonSocial
                                    Case "ruc"
                                        Dim ruc = n.InnerText.ToString
                                        Nc.NotaCreditoCabecera._ruc = ruc
                                    Case "estab"
                                        Dim estab = n.InnerText.ToString
                                        Nc.NotaCreditoCabecera._estab = estab
                                    Case "ptoEmi"
                                        Dim ptoEmi = n.InnerText.ToString
                                        Nc.NotaCreditoCabecera._ptoEmi = ptoEmi
                                    Case "secuencial"
                                        Dim secuencial = n.InnerText.ToString
                                        Nc.NotaCreditoCabecera._secuencial = secuencial
                                    Case "claveAcceso"
                                        Dim claveAcceso = n.InnerText.ToString
                                        Nc.NotaCreditoCabecera._claveAcceso = claveAcceso
                                End Select

                            Next
                        Case "infoNotaCredito"
                            For Each infoNc As XmlNode In nodoInfTri.ChildNodes
                                Select Case infoNc.Name
                                    Case "fechaEmision"
                                        Dim fechaemision = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._fechaEmision = fechaemision
                                    Case "dirEstablecimiento"
                                        Dim dirEstablecimiento = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._dirEstablecimiento = dirEstablecimiento
                                    Case "razonSocialComprador"
                                        Dim razonSocialComprador = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._razonSocialComprador = razonSocialComprador
                                    Case "identificacionComprador"
                                        Dim identificacionComprador = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._identificacionComprador = identificacionComprador
                                    Case "direccionComprador"
                                        Dim direccionComprador = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._direccionComprador = direccionComprador
                                    Case "codDocModificado"
                                        Dim codDocModificado = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._CodDocMod = codDocModificado
                                    Case "numDocModificado"
                                        Dim numDocModificado = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._numDocModificado = numDocModificado
                                    Case "fechaEmisionDocSustento"
                                        Dim fechaEmisionDocSustento = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._fechaEmisionDocSustento = CDate(fechaEmisionDocSustento)
                                    Case "totalSinImpuestos"
                                        Dim totalSinImpuestos = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._totalSinImpuestos = totalSinImpuestos
                                    Case "motivo"
                                        Dim motivo = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._motivo = motivo
                                    Case "valorModificacion"
                                        Dim valorModificacion = infoNc.InnerText
                                        Nc.NotaCreditoCabecera._valorModificacion = valorModificacion
                                    Case "totalConImpuestos"
                                        For Each totalConImpuestos As XmlNode In infoNc.ChildNodes
                                            Select Case totalConImpuestos.Name
                                                Case "totalImpuesto"
                                                    Dim NcCabImp As New NotaCreditoCabeceraImpuesto
                                                    For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                        Select Case totalImpuesto.Name
                                                            Case "codigo"
                                                                Dim codigo = totalImpuesto.InnerText
                                                                NcCabImp._codigo = CInt(codigo)
                                                            Case "codigoPorcentaje"
                                                                Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                NcCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                            Case "baseImponible"
                                                                Dim baseImponible = totalImpuesto.InnerText
                                                                NcCabImp._baseImponible = baseImponible
                                                            Case "tarifa"
                                                                Dim tarifa = totalImpuesto.InnerText
                                                                NcCabImp._tarifa = tarifa
                                                            Case "valor"
                                                                Dim valor = totalImpuesto.InnerText
                                                                NcCabImp._valor = valor
                                                        End Select

                                                    Next
                                                    Nc.NotaCreditoCabecera._impuestos.Add(NcCabImp)
                                            End Select
                                        Next


                                End Select
                            Next
                        Case "detalles"

                            For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                Select Case detalles.Name
                                    Case "detalle"
                                        Dim NcDet As New NotaCreditoDetalle
                                        Dim NcDetImp As New NotaCreditoDetalleImpuesto
                                        NcDet._impuestos = New List(Of NotaCreditoDetalleImpuesto)
                                        For Each detalle As XmlNode In detalles.ChildNodes
                                            Select Case detalle.Name
                                                Case "codigoInterno"
                                                    Dim codigoPrincipal = detalle.InnerText
                                                    NcDet._codigoInterno = codigoPrincipal
                                                Case "codigoAdicional"
                                                    Dim codigoAuxiliar = detalle.InnerText
                                                    NcDet._codigoAdicional = codigoAuxiliar
                                                Case "descripcion"
                                                    Dim descripcion = detalle.InnerText
                                                    NcDet._descripcion = descripcion
                                                Case "cantidad"
                                                    Dim cantidad = detalle.InnerText
                                                    NcDet._cantidad = cantidad
                                                Case "precioUnitario"
                                                    Dim precioUnitario = detalle.InnerText
                                                    NcDet._precioUnitario = precioUnitario
                                                Case "descuento"
                                                    Dim descuento = detalle.InnerText
                                                    NcDet._descuento = descuento
                                                Case "precioTotalSinImpuesto"
                                                    Dim precioTotalSinImpuesto = detalle.InnerText
                                                    NcDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                Case "impuestos"
                                                    For Each impuestos As XmlNode In detalle.ChildNodes
                                                        Select Case impuestos.Name
                                                            Case "impuesto"
                                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                    Select Case impuesto.Name
                                                                        Case "codigo"
                                                                            Dim codigo = impuesto.InnerText
                                                                            NcDetImp._codigo = CInt(codigo)
                                                                        Case "codigoPorcentaje"
                                                                            Dim codigoPorcentaje = impuesto.InnerText
                                                                            NcDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                        Case "tarifa"
                                                                            Dim tarifa = impuesto.InnerText
                                                                            NcDetImp._tarifa = tarifa
                                                                        Case "baseImponible"
                                                                            Dim baseImponible = impuesto.InnerText
                                                                            NcDetImp._baseImponible = baseImponible
                                                                        Case "valor"
                                                                            Dim valor = impuesto.InnerText
                                                                            NcDetImp._valor = valor
                                                                    End Select
                                                                Next
                                                                NcDet._impuestos.Add(NcDetImp)
                                                        End Select
                                                    Next
                                            End Select
                                        Next
                                        Nc.NotaCreditoDetalle.Add(NcDet)
                                End Select
                            Next

                    End Select
                Next

            End If
            GuardaLog("NC clave de acceso: " + Nc.NotaCreditoCabecera._claveAcceso + " leido Correctamente!!" + " nombre del archivo: " + Path.GetFileName(ruta))
            Return Nc
        Catch ex As Exception
            GuardaLog("Error al leer XML NC  ==> " + ex.Message + " con nombre: " + Path.GetFileName(ruta))
            Return Nothing
        End Try

    End Function

    Public Function LeerXMLRetencion(ruta As String) As Retencion

        Try

            GuardaLog("Consultando XML a procesar..")

            Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\ret.xml"

            Dim mensaje As String = ""

            Dim ms As New MemoryStream


            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList

            m_xmld = New XmlDocument()

            ' m_xmld.Load(_RUTA
            m_xmld.Load(ruta)

            Dim Retencion As New Retencion
            Retencion.RetCabecera = New RetCabecera
            Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

            Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")
            For Each facAut As XmlNode In nodoAut.ChildNodes
                Select Case facAut.Name
                    Case "numeroAutorizacion"
                        Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                    Case "fechaAutorizacion"
                        Retencion.RetCabecera._FechaAutorizacion = CDate(facAut.InnerText)
                End Select
            Next

            Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

            Dim comprobante = nodo.InnerText

            Dim nodoRet As New XmlDocument()
            nodoRet.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

            Dim razonSocial As String = ""
            Dim ruc As String = ""
            Dim estab As String = ""
            Dim ptoEmi As String = ""
            Dim secuencial As String = ""


            For Each Ret As XmlNode In nodoRet.ChildNodes
                Select Case Ret.Name
                    Case "comprobanteRetencion"
                        For Each nodoInfTri As XmlNode In Ret.ChildNodes
                            Select Case nodoInfTri.Name
                                Case "infoTributaria"
                                    For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                        Select Case infoTri.Name
                                            Case "razonSocial"
                                                razonSocial = infoTri.InnerText.ToString
                                                Retencion.RetCabecera._RazonSocial = razonSocial
                                            Case "ruc"
                                                ruc = infoTri.InnerText.ToString
                                                Retencion.RetCabecera._ruc = ruc
                                            Case "estab"
                                                estab = infoTri.InnerText.ToString
                                                Retencion.RetCabecera._estab = estab
                                            Case "ptoEmi"
                                                ptoEmi = infoTri.InnerText.ToString
                                                Retencion.RetCabecera._ptoEmi = ptoEmi
                                            Case "secuencial"
                                                secuencial = infoTri.InnerText.ToString
                                                Retencion.RetCabecera._secuencial = secuencial
                                            Case "claveAcceso"
                                                Dim claveAcceso = infoTri.InnerText.ToString
                                                Retencion.RetCabecera._claveAcceso = claveAcceso
                                        End Select
                                    Next
                                Case "infoCompRetencion"
                                    For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                        Select Case infoRt.Name
                                            Case "fechaEmision"
                                                Retencion.RetCabecera._fechaEmision = CDate(infoRt.InnerText)
                                            Case "dirEstablecimiento"
                                                Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                            Case "razonSocialSujetoRetenido"
                                                Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                            Case "identificacionSujetoRetenido"
                                                Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                            Case "periodoFiscal"
                                                Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                            Case "impuestos"
                                        End Select
                                    Next
                                Case "impuestos"
                                    Dim RTDetImp As New RetDetalleImpuestos
                                    For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                        Select Case impuestos.Name
                                            Case "impuesto"
                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                    Select Case impuesto.Name
                                                        Case "codigo"
                                                            RTDetImp._codigo = CInt(impuesto.InnerText)
                                                        Case "codigoRetencion"
                                                            RTDetImp._codigoRetencion = impuesto.InnerText
                                                        Case "baseImponible"
                                                            RTDetImp._baseImponible = CDec(impuesto.InnerText)
                                                        Case "porcentajeRetener"
                                                            RTDetImp._porcentajeRetener = CDec(impuesto.InnerText)
                                                        Case "valorRetenido"
                                                            RTDetImp._valorRetenido = CDec(impuesto.InnerText)
                                                        Case "codDocSustento"
                                                            RTDetImp._codDocSustento = impuesto.InnerText
                                                        Case "numDocSustento"
                                                            RTDetImp._numDocSustento = impuesto.InnerText
                                                        Case "fechaEmisionDocSustento"
                                                            RTDetImp._fechaEmisionDocSustento = CDate(impuesto.InnerText)
                                                    End Select
                                                Next
                                                Retencion.RetDetalleImp.Add(RTDetImp)
                                        End Select
                                    Next

                            End Select
                        Next

                End Select
            Next
            GuardaLog("Proceso Hilos Completado Correctamente!!")
            Return Retencion
        Catch ex As Exception
            GuardaLog("Error al leer XML: " + ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Public Function LeerXMLRetencion2(ruta As String) As Retencion

        Try

            GuardaLog("Consultando XML Retencion a procesar..")

            'Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\ret.xml"

            Dim mensaje As String = ""

            'Dim ms As New MemoryStream


            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList

            m_xmld = New XmlDocument()

            m_xmld.Load(ruta)




            Dim nodoAutSchema = m_xmld.SelectSingleNode("autorizacion")


            Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")

            Dim _nodoAut = m_xmld.SelectSingleNode("respuestaComprobante/autorizaciones/autorizacion")

            If (Not (nodoAutSchema) Is Nothing) Then

                Dim Retencion As New Retencion
                Retencion.RetCabecera = New RetCabecera
                Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                For Each facAut As XmlNode In nodoAut.ChildNodes
                    Select Case facAut.Name
                        Case "numeroAutorizacion"
                            Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                        Case "fechaAutorizacion"
                            Retencion.RetCabecera._FechaAutorizacion = facAut.InnerText
                        Case "comprobante"
                            For Each comp As XmlNode In facAut.ChildNodes
                                Select Case comp.Name
                                    Case "comprobanteRetencion"
                                        For Each compRet As XmlNode In comp.ChildNodes
                                            Select Case compRet.Name
                                                Case "infoTributaria"
                                                    For Each infoTri As XmlNode In compRet.ChildNodes
                                                        Select Case infoTri.Name
                                                            Case "razonSocial"
                                                                Dim razonSocial = infoTri.InnerText.ToString
                                                                Retencion.RetCabecera._RazonSocial = razonSocial
                                                            Case "ruc"
                                                                Dim ruc = infoTri.InnerText.ToString
                                                                Retencion.RetCabecera._ruc = ruc
                                                            Case "estab"
                                                                Dim estab = infoTri.InnerText.ToString
                                                                Retencion.RetCabecera._estab = estab
                                                            Case "ptoEmi"
                                                                Dim ptoEmi = infoTri.InnerText.ToString
                                                                Retencion.RetCabecera._ptoEmi = ptoEmi
                                                            Case "secuencial"
                                                                Dim secuencial = infoTri.InnerText.ToString
                                                                Retencion.RetCabecera._secuencial = secuencial
                                                            Case "claveAcceso"
                                                                Dim claveAcceso = infoTri.InnerText.ToString
                                                                Retencion.RetCabecera._claveAcceso = claveAcceso
                                                        End Select
                                                    Next

                                                Case "infoCompRetencion"
                                                    For Each infoRt As XmlNode In compRet.ChildNodes
                                                        Select Case infoRt.Name
                                                            Case "fechaEmision"
                                                                Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                            Case "dirEstablecimiento"
                                                                Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                            Case "razonSocialSujetoRetenido"
                                                                Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                            Case "identificacionSujetoRetenido"
                                                                Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                            Case "periodoFiscal"
                                                                Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                            Case "impuestos"
                                                        End Select
                                                    Next

                                                Case "impuestos"

                                                    For Each impuestos As XmlNode In compRet.ChildNodes
                                                        Select Case impuestos.Name
                                                            Case "impuesto"
                                                                Dim RTDetImp As New RetDetalleImpuestos
                                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                    Select Case impuesto.Name
                                                                        Case "codigo"
                                                                            RTDetImp._codigo = CInt(impuesto.InnerText)
                                                                        Case "codigoRetencion"
                                                                            RTDetImp._codigoRetencion = impuesto.InnerText
                                                                        Case "baseImponible"
                                                                            RTDetImp._baseImponible = impuesto.InnerText
                                                                        Case "porcentajeRetener"
                                                                            RTDetImp._porcentajeRetener = impuesto.InnerText
                                                                        Case "valorRetenido"
                                                                            RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                        Case "codDocSustento"
                                                                            RTDetImp._codDocSustento = impuesto.InnerText
                                                                        Case "numDocSustento"
                                                                            RTDetImp._numDocSustento = impuesto.InnerText
                                                                        Case "fechaEmisionDocSustento"
                                                                            RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                                    End Select
                                                                Next
                                                                Retencion.RetDetalleImp.Add(RTDetImp)
                                                        End Select
                                                    Next

                                            End Select


                                        Next
                                End Select
                            Next

                    End Select
                Next


                If Not IsNothing(Retencion.RetCabecera._RazonSocial) Then
                    Return Retencion
                End If

            End If

            If (Not (_nodoAut) Is Nothing) Then

                Dim Retencion As New Retencion
                Retencion.RetCabecera = New RetCabecera
                Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                For Each facAut As XmlNode In _nodoAut.ChildNodes
                    Select Case facAut.Name
                        Case "numeroAutorizacion"
                            Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                        Case "fechaAutorizacion"
                            Retencion.RetCabecera._FechaAutorizacion = facAut.InnerText.ToString
                    End Select
                Next

                Dim nodo = m_xmld.SelectSingleNode("respuestaComprobante/autorizaciones/autorizacion/comprobante")

                Dim comprobante = nodo.InnerText

                Dim nodoRet As New XmlDocument()
                nodoRet.LoadXml(comprobante)


                For Each comp As XmlNode In nodoRet.ChildNodes
                    Select Case comp.Name
                        Case "comprobanteRetencion"
                            For Each info As XmlNode In comp.ChildNodes
                                Select Case info.Name
                                    Case "infoTributaria"
                                        For Each infoTri As XmlNode In info.ChildNodes
                                            Select Case infoTri.Name
                                                Case "razonSocial"
                                                    Dim razonSocial = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._RazonSocial = razonSocial
                                                Case "ruc"
                                                    Dim ruc = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._ruc = ruc
                                                Case "estab"
                                                    Dim estab = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._estab = estab
                                                Case "ptoEmi"
                                                    Dim ptoEmi = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._ptoEmi = ptoEmi
                                                Case "secuencial"
                                                    Dim secuencial = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._secuencial = secuencial
                                                Case "claveAcceso"
                                                    Dim claveAcceso = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._claveAcceso = claveAcceso
                                            End Select
                                        Next

                                    Case "infoCompRetencion"
                                        For Each infoRt As XmlNode In info.ChildNodes
                                            Select Case infoRt.Name
                                                Case "fechaEmision"
                                                    Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                Case "dirEstablecimiento"
                                                    Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                Case "razonSocialSujetoRetenido"
                                                    Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                Case "identificacionSujetoRetenido"
                                                    Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                Case "periodoFiscal"
                                                    Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                Case "impuestos"
                                            End Select
                                        Next

                                    Case "impuestos"

                                        For Each impuestos As XmlNode In info.ChildNodes
                                            Select Case impuestos.Name
                                                Case "impuesto"
                                                    Dim RTDetImp As New RetDetalleImpuestos
                                                    For Each impuesto As XmlNode In impuestos.ChildNodes
                                                        Select Case impuesto.Name
                                                            Case "codigo"
                                                                RTDetImp._codigo = CInt(impuesto.InnerText)
                                                            Case "codigoRetencion"
                                                                RTDetImp._codigoRetencion = impuesto.InnerText
                                                            Case "baseImponible"
                                                                RTDetImp._baseImponible = impuesto.InnerText
                                                            Case "porcentajeRetener"
                                                                RTDetImp._porcentajeRetener = impuesto.InnerText
                                                            Case "valorRetenido"
                                                                RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                            Case "codDocSustento"
                                                                RTDetImp._codDocSustento = impuesto.InnerText
                                                            Case "numDocSustento"
                                                                RTDetImp._numDocSustento = impuesto.InnerText
                                                            Case "fechaEmisionDocSustento"
                                                                RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                        End Select
                                                    Next
                                                    Retencion.RetDetalleImp.Add(RTDetImp)
                                            End Select
                                        Next

                                End Select
                            Next
                    End Select
                Next

                GuardaLog("Retencion clave de acceso: " + Retencion.RetCabecera._claveAcceso + " leido Correctamente!!")
                Return Retencion

            End If


            If (Not (nodoAut) Is Nothing) Then

                Dim Retencion As New Retencion
                Retencion.RetCabecera = New RetCabecera
                Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                For Each facAut As XmlNode In nodoAut.ChildNodes
                    Select Case facAut.Name
                        Case "numeroAutorizacion"
                            Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                        Case "fechaAutorizacion"
                            Retencion.RetCabecera._FechaAutorizacion = facAut.InnerText
                    End Select
                Next

                Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

                Dim comprobante = nodo.InnerText

                Dim nodoRet As New XmlDocument()
                nodoRet.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

                Dim razonSocial As String = ""
                Dim ruc As String = ""
                Dim estab As String = ""
                Dim ptoEmi As String = ""
                Dim secuencial As String = ""


                For Each Ret As XmlNode In nodoRet.ChildNodes
                    Select Case Ret.Name
                        Case "comprobanteRetencion"
                            For Each nodoInfTri As XmlNode In Ret.ChildNodes
                                Select Case nodoInfTri.Name
                                    Case "infoTributaria"
                                        For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                            Select Case infoTri.Name
                                                Case "razonSocial"
                                                    razonSocial = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._RazonSocial = razonSocial
                                                Case "ruc"
                                                    ruc = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._ruc = ruc
                                                Case "estab"
                                                    estab = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._estab = estab
                                                Case "ptoEmi"
                                                    ptoEmi = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._ptoEmi = ptoEmi
                                                Case "secuencial"
                                                    secuencial = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._secuencial = secuencial
                                                Case "claveAcceso"
                                                    Dim claveAcceso = infoTri.InnerText.ToString
                                                    Retencion.RetCabecera._claveAcceso = claveAcceso
                                            End Select
                                        Next
                                    Case "infoCompRetencion"
                                        For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                            Select Case infoRt.Name
                                                Case "fechaEmision"
                                                    Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                Case "dirEstablecimiento"
                                                    Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                Case "razonSocialSujetoRetenido"
                                                    Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                Case "identificacionSujetoRetenido"
                                                    Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                Case "periodoFiscal"
                                                    Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                Case "impuestos"
                                            End Select
                                        Next

                                    Case "docsSustento"
                                        Dim codDocSustento As String = ""
                                        Dim _numDocSustento As String = ""
                                        Dim fechaEmisionDocSustento As String = ""
                                        For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                            Select Case impuestos.Name
                                                Case "docSustento"
                                                    For Each docSustento As XmlNode In impuestos.ChildNodes
                                                        Select Case docSustento.Name

                                                            Case "codDocSustento"
                                                                codDocSustento = docSustento.InnerText
                                                            Case "numDocSustento"
                                                                _numDocSustento = docSustento.InnerText
                                                            Case "fechaEmisionDocSustento"
                                                                fechaEmisionDocSustento = docSustento.InnerText

                                                            Case "retenciones"
                                                                For Each retenciones As XmlNode In docSustento.ChildNodes
                                                                    Select Case retenciones.Name
                                                                        Case "retencion"
                                                                            Dim RTDetImp As New RetDetalleImpuestos
                                                                            For Each _retencion As XmlNode In retenciones.ChildNodes
                                                                                Select Case _retencion.Name
                                                                                    Case "codigo"
                                                                                        RTDetImp._codigo = CInt(_retencion.InnerText)
                                                                                    Case "codigoRetencion"
                                                                                        RTDetImp._codigoRetencion = _retencion.InnerText
                                                                                    Case "baseImponible"
                                                                                        RTDetImp._baseImponible = _retencion.InnerText
                                                                                    Case "porcentajeRetener"
                                                                                        RTDetImp._porcentajeRetener = _retencion.InnerText
                                                                                    Case "valorRetenido"
                                                                                        RTDetImp._valorRetenido = validaSiEmpiezaPunto(_retencion.InnerText)
                                                                                        'Case "codDocSustento"
                                                                                        '    RTDetImp._codDocSustento = _retencion.InnerText
                                                                                        'Case "numDocSustento"
                                                                                        '    RTDetImp._numDocSustento = _retencion.InnerText
                                                                                        'Case "fechaEmisionDocSustento"
                                                                                        '    RTDetImp._fechaEmisionDocSustento = _retencion.InnerText.ToString
                                                                                        RTDetImp._codDocSustento = codDocSustento
                                                                                        RTDetImp._numDocSustento = _numDocSustento
                                                                                        RTDetImp._fechaEmisionDocSustento = fechaEmisionDocSustento
                                                                                End Select
                                                                            Next
                                                                            Retencion.RetDetalleImp.Add(RTDetImp)
                                                                    End Select

                                                                Next


                                                        End Select
                                                    Next

                                            End Select
                                        Next

                                    Case "impuestos"

                                        For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                            Select Case impuestos.Name
                                                Case "impuesto"
                                                    Dim RTDetImp As New RetDetalleImpuestos
                                                    For Each impuesto As XmlNode In impuestos.ChildNodes
                                                        Select Case impuesto.Name
                                                            Case "codigo"
                                                                RTDetImp._codigo = CInt(impuesto.InnerText)
                                                            Case "codigoRetencion"
                                                                RTDetImp._codigoRetencion = impuesto.InnerText
                                                            Case "baseImponible"
                                                                RTDetImp._baseImponible = impuesto.InnerText
                                                            Case "porcentajeRetener"
                                                                RTDetImp._porcentajeRetener = impuesto.InnerText
                                                            Case "valorRetenido"
                                                                RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                            Case "codDocSustento"
                                                                RTDetImp._codDocSustento = impuesto.InnerText
                                                            Case "numDocSustento"
                                                                RTDetImp._numDocSustento = impuesto.InnerText
                                                            Case "fechaEmisionDocSustento"
                                                                RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                        End Select
                                                    Next
                                                    Retencion.RetDetalleImp.Add(RTDetImp)
                                            End Select
                                        Next

                                End Select
                            Next

                    End Select
                Next

                GuardaLog("Retencion clave de acceso: " + Retencion.RetCabecera._claveAcceso + " leido Correctamente!!")
                Return Retencion

            Else

                Dim rt = m_xmld.SelectSingleNode("comprobanteRetencion")
                If (Not (nodoAut) Is Nothing) Then

                    Dim Retencion As New Retencion
                    Retencion.RetCabecera = New RetCabecera
                    Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                    For Each nodoInfTri As XmlNode In rt.ChildNodes
                        Select Case nodoInfTri.Name
                            Case "infoTributaria"
                                For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                    Select Case infoTri.Name
                                        Case "razonSocial"
                                            Dim razonSocial = infoTri.InnerText.ToString
                                            Retencion.RetCabecera._RazonSocial = razonSocial
                                        Case "ruc"
                                            Dim ruc = infoTri.InnerText.ToString
                                            Retencion.RetCabecera._ruc = ruc
                                        Case "estab"
                                            Dim estab = infoTri.InnerText.ToString
                                            Retencion.RetCabecera._estab = estab
                                        Case "ptoEmi"
                                            Dim ptoEmi = infoTri.InnerText.ToString
                                            Retencion.RetCabecera._ptoEmi = ptoEmi
                                        Case "secuencial"
                                            Dim secuencial = infoTri.InnerText.ToString
                                            Retencion.RetCabecera._secuencial = secuencial
                                        Case "claveAcceso"
                                            Dim claveAcceso = infoTri.InnerText.ToString
                                            Retencion.RetCabecera._claveAcceso = claveAcceso
                                    End Select
                                Next
                            Case "infoCompRetencion"
                                For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                    Select Case infoRt.Name
                                        Case "fechaEmision"
                                            Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                        Case "dirEstablecimiento"
                                            Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                        Case "razonSocialSujetoRetenido"
                                            Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                        Case "identificacionSujetoRetenido"
                                            Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                        Case "periodoFiscal"
                                            Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                        Case "impuestos"
                                    End Select
                                Next
                            Case "impuestos"

                                For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                    Select Case impuestos.Name
                                        Case "impuesto"
                                            Dim RTDetImp As New RetDetalleImpuestos
                                            For Each impuesto As XmlNode In impuestos.ChildNodes
                                                Select Case impuesto.Name
                                                    Case "codigo"
                                                        RTDetImp._codigo = CInt(impuesto.InnerText)
                                                    Case "codigoRetencion"
                                                        RTDetImp._codigoRetencion = impuesto.InnerText
                                                    Case "baseImponible"
                                                        RTDetImp._baseImponible = impuesto.InnerText
                                                    Case "porcentajeRetener"
                                                        RTDetImp._porcentajeRetener = impuesto.InnerText
                                                    Case "valorRetenido"
                                                        RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                    Case "codDocSustento"
                                                        RTDetImp._codDocSustento = impuesto.InnerText
                                                    Case "numDocSustento"
                                                        RTDetImp._numDocSustento = impuesto.InnerText
                                                    Case "fechaEmisionDocSustento"
                                                        RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                End Select
                                            Next
                                            Retencion.RetDetalleImp.Add(RTDetImp)
                                    End Select
                                Next

                        End Select
                    Next
                    GuardaLog("Retencion clave de acceso: " + Retencion.RetCabecera._claveAcceso + " leido Correctamente!!")
                    Return Retencion
                Else

                    Dim Retencion As New Retencion
                    Retencion.RetCabecera = New RetCabecera
                    Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                    Dim lector = New XmlDocument()
                    lector.Load(ruta)

                    Dim cvev As String = lector.GetNamespaceOfPrefix("http://ec.gob.sri.ws.autorizacion")
                    Dim xmlns As New Xml.XmlNamespaceManager(lector.NameTable)
                    xmlns.AddNamespace("q1", "http://ec.gob.sri.ws.autorizacion")
                    Dim xnodo As Xml.XmlNode
                    xnodo = lector.SelectSingleNode("/q1:respuestaComprobante/autorizaciones/autorizacion", xmlns)

                    If (Not (xnodo) Is Nothing) Then

                        For Each Aut As XmlNode In xnodo.ChildNodes
                            Select Case Aut.Name
                                Case "numeroAutorizacion"
                                    Retencion.RetCabecera._NumeroAutorizacion = Aut.InnerText

                                Case "fechaAutorizacion"
                                    Retencion.RetCabecera._FechaAutorizacion = Aut.InnerText

                            End Select
                        Next

                        Dim nodo = lector.SelectSingleNode("/q1:respuestaComprobante/autorizaciones/autorizacion/comprobante", xmlns)

                        Dim comprobante = nodo.InnerText

                        Dim nodoRet As New XmlDocument()
                        nodoRet.LoadXml(comprobante)

                        For Each Ret As XmlNode In nodoRet.ChildNodes
                            Select Case Ret.Name
                                Case "comprobanteRetencion"
                                    For Each nodoInfTri As XmlNode In Ret.ChildNodes
                                        Select Case nodoInfTri.Name
                                            Case "infoTributaria"
                                                For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                                    Select Case infoTri.Name
                                                        Case "razonSocial"
                                                            Dim razonSocial = infoTri.InnerText.ToString
                                                            Retencion.RetCabecera._RazonSocial = razonSocial
                                                        Case "ruc"
                                                            Dim ruc = infoTri.InnerText.ToString
                                                            Retencion.RetCabecera._ruc = ruc
                                                        Case "estab"
                                                            Dim estab = infoTri.InnerText.ToString
                                                            Retencion.RetCabecera._estab = estab
                                                        Case "ptoEmi"
                                                            Dim ptoEmi = infoTri.InnerText.ToString
                                                            Retencion.RetCabecera._ptoEmi = ptoEmi
                                                        Case "secuencial"
                                                            Dim secuencial = infoTri.InnerText.ToString
                                                            Retencion.RetCabecera._secuencial = secuencial
                                                        Case "claveAcceso"
                                                            Dim claveAcceso = infoTri.InnerText.ToString
                                                            Retencion.RetCabecera._claveAcceso = claveAcceso
                                                    End Select
                                                Next

                                            Case "infoCompRetencion"
                                                For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                                    Select Case infoRt.Name
                                                        Case "fechaEmision"
                                                            Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                        Case "dirEstablecimiento"
                                                            Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                        Case "razonSocialSujetoRetenido"
                                                            Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                        Case "identificacionSujetoRetenido"
                                                            Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                        Case "periodoFiscal"
                                                            Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                        Case "impuestos"
                                                    End Select
                                                Next

                                            Case "impuestos"

                                                For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                                    Select Case impuestos.Name
                                                        Case "impuesto"
                                                            Dim RTDetImp As New RetDetalleImpuestos

                                                            For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                Select Case impuesto.Name
                                                                    Case "codigo"
                                                                        RTDetImp._codigo = CInt(impuesto.InnerText)
                                                                    Case "codigoRetencion"
                                                                        RTDetImp._codigoRetencion = impuesto.InnerText
                                                                    Case "baseImponible"
                                                                        RTDetImp._baseImponible = impuesto.InnerText
                                                                    Case "porcentajeRetener"
                                                                        RTDetImp._porcentajeRetener = impuesto.InnerText
                                                                    Case "valorRetenido"
                                                                        RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                    Case "codDocSustento"
                                                                        RTDetImp._codDocSustento = impuesto.InnerText
                                                                    Case "numDocSustento"
                                                                        RTDetImp._numDocSustento = impuesto.InnerText
                                                                    Case "fechaEmisionDocSustento"
                                                                        RTDetImp._fechaEmisionDocSustento = impuesto.InnerText
                                                                End Select
                                                            Next
                                                            Retencion.RetDetalleImp.Add(RTDetImp)
                                                    End Select
                                                Next
                                        End Select
                                    Next
                            End Select
                        Next


                    End If
                    GuardaLog("Retencion clave de acceso: " + Retencion.RetCabecera._claveAcceso + " leido Correctamente" + " - nombre del archivo: " + Path.GetFileName(ruta))
                    Return Retencion

                End If




            End If


        Catch ex As Exception
            GuardaLog("Error al leer XML retencion: " + ex.Message.ToString + " con nombre: " + Path.GetFileName(ruta))
            Return Nothing
        End Try

    End Function

    Public Sub MoverXML(ByVal FullName As String, ByVal Name As String, ByVal Directorio As String, ByVal Extension As String)
        Try
            'Name = Name.ToLower
            Dim temp = Name
            Dim flag = True
            Dim idarchivo As Integer = 0
            While flag = True
                If System.IO.File.Exists(Directorio & temp) Then
                    idarchivo += 1
                    temp = Name.Replace(Extension, "") & idarchivo & Extension
                Else
                    System.IO.File.Move(FullName, Directorio & temp)
                    flag = False
                End If
            End While
            GuardaLog("Ruta destino: " & Directorio & temp)
            If System.IO.File.Exists(Directorio & temp) Then
                GuardaLog("Archivo movido exitosamente")
            Else
                GuardaLog("No se pudo mover el archivo")
            End If
        Catch ex As Exception
            GuardaLog("MoverArchivo: Fallido -> " & ex.Message)
        End Try
    End Sub

    Public Sub MoverXMLFC(ByVal FullName As String, ByVal Name As String, ByVal Directorio As String, ByVal Extension As String)
        Try
            'Name = Name.ToLower
            Dim temp = Name
            Dim flag = True
            Dim idarchivo As Integer = 0
            While flag = True
                If System.IO.File.Exists(Directorio & temp) Then
                    idarchivo += 1
                    temp = Name.Replace(Extension, "") & idarchivo & Extension
                Else
                    System.IO.File.Move(FullName, Directorio & temp)
                    flag = False
                End If
            End While
            GuardaLog("Ruta destino: " & Directorio & temp)
            If System.IO.File.Exists(Directorio & temp) Then
                GuardaLog("Archivo movido exitosamente")
            Else
                GuardaLog("No se pudo mover el archivo")
            End If
        Catch ex As Exception
            GuardaLog("MoverArchivo: Fallido -> " & ex.Message)
        End Try
    End Sub

    Public Function insertarEntidadFC(ByVal oFactura As Factura, ByVal nombreAr As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Dim DocEntryUdoFC As String = getRSvalue("Select ""DocEntry"" from ""@GS_FC"" where ""U_ClaAcc""= '" + oFactura.FacturaCabecera._claveAcceso.ToString + "'", "DocEntry", "")
        If DocEntryUdoFC = "0" Or DocEntryUdoFC = "" Then
            Try
                'GuardaLog("ImporteTotal: " + Convert.ToDouble(oFactura.FacturaCabecera._importeTotal).ToString)
                'GuardaLog("ImporteTotal: " + formatDecimal(Convert.ToDouble(oFactura.FacturaCabecera._importeTotal).ToString()).ToString())
                'GuardaLog("ImporteTotal: " + CDec(oFactura.FacturaCabecera._importeTotal).ToString)
                'GuardaLog("ImporteTotal: " + oFactura.FacturaCabecera._importeTotal.ToString)
                'GuardaLog("ImporteTotal: " + Convert.ToDecimal(oFactura.FacturaCabecera._importeTotal).ToString("###0.00"))

                oCompanyService = oCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_FC")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                If String.IsNullOrEmpty(oFactura.FacturaCabecera._NumeroAutorizacion) Then
                    oGeneralData.SetProperty("U_NumAut", oFactura.FacturaCabecera._claveAcceso.ToString)
                Else
                    oGeneralData.SetProperty("U_NumAut", oFactura.FacturaCabecera._NumeroAutorizacion.ToString)
                End If

                If String.IsNullOrEmpty(oFactura.FacturaCabecera._FechaAutorizacion) Then
                    oGeneralData.SetProperty("U_FechaAut", oFactura.FacturaCabecera._fechaEmision.ToString)
                Else
                    oGeneralData.SetProperty("U_FechaAut", oFactura.FacturaCabecera._FechaAutorizacion.ToString)
                End If

                oGeneralData.SetProperty("U_Ruc", oFactura.FacturaCabecera._ruc)
                oGeneralData.SetProperty("U_RazSoc", oFactura.FacturaCabecera._RazonSocial)
                oGeneralData.SetProperty("U_ClaAcc", oFactura.FacturaCabecera._claveAcceso)
                oGeneralData.SetProperty("U_Est", oFactura.FacturaCabecera._estab)
                oGeneralData.SetProperty("U_PuntoEmi", oFactura.FacturaCabecera._ptoEmi)
                oGeneralData.SetProperty("U_Sec", oFactura.FacturaCabecera._secuencial)
                oGeneralData.SetProperty("U_FecEmi", oFactura.FacturaCabecera._fechaEmision.ToString)
                oGeneralData.SetProperty("U_DirEst", IIf(IsNothing(oFactura.FacturaCabecera._dirEstablecimiento), "", oFactura.FacturaCabecera._dirEstablecimiento))
                oGeneralData.SetProperty("U_ConEsp", IIf(IsNothing(oFactura.FacturaCabecera._contribuyenteEspecial), "", oFactura.FacturaCabecera._contribuyenteEspecial))
                oGeneralData.SetProperty("U_RazSocComp", oFactura.FacturaCabecera._razonSocialComprador)
                oGeneralData.SetProperty("U_IdenComp", oFactura.FacturaCabecera._identificacionComprador)
                oGeneralData.SetProperty("U_DirComp", IIf(IsNothing(oFactura.FacturaCabecera._direccionComprador), "", oFactura.FacturaCabecera._direccionComprador))

                oGeneralData.SetProperty("U_TotSinImp", Convert.ToDouble(formatDecimal(oFactura.FacturaCabecera._totalSinImpuestos)))
                oGeneralData.SetProperty("U_TotDesc", Convert.ToDouble(formatDecimal(oFactura.FacturaCabecera._totalDescuento)))
                oGeneralData.SetProperty("U_ImpTotal", Convert.ToDouble(formatDecimal(oFactura.FacturaCabecera._importeTotal)))

                If Not String.IsNullOrEmpty(oFactura.FacturaCabecera._formaPago) Then

                    oGeneralData.SetProperty("U_FormaPago", oFactura.FacturaCabecera._formaPago.ToString)
                    oGeneralData.SetProperty("U_TotalPago", Convert.ToDouble(formatDecimal(oFactura.FacturaCabecera._totalFormaPago)))
                    oGeneralData.SetProperty("U_PlazoPago", IIf(IsNothing(oFactura.FacturaCabecera._plazo.ToString), "", oFactura.FacturaCabecera._plazo.ToString))

                    If Not String.IsNullOrEmpty(oFactura.FacturaCabecera._unidadTiempo) Then
                        oGeneralData.SetProperty("U_UniTiempo", IIf(IsNothing(oFactura.FacturaCabecera._unidadTiempo.ToString), "", oFactura.FacturaCabecera._unidadTiempo.ToString))
                    End If
                End If

                oGeneralData.SetProperty("U_Estado", "Importado")
                'oGeneralData.SetProperty("U_UniTiempo", oFactura.FacturaCabecera._unidadTiempo)

                For Each CabImp In oFactura.FacturaCabecera._impuestos
                    If CabImp._codigo = 2 Then
                        If CabImp._codigoPorcentaje = 0 Then

                            oGeneralData.SetProperty("U_Cod0", CabImp._codigo.ToString)
                            oGeneralData.SetProperty("U_CodPorc0", CabImp._codigoPorcentaje.ToString)
                            oGeneralData.SetProperty("U_BaseImp0", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralData.SetProperty("U_Tarifa0", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralData.SetProperty("U_Valor0", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 8 Then

                            oGeneralData.SetProperty("U_Cod8", CabImp._codigo.ToString)
                            oGeneralData.SetProperty("U_CodPorc8", CabImp._codigoPorcentaje.ToString)
                            oGeneralData.SetProperty("U_BaseImp8", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralData.SetProperty("U_Tarifa8", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralData.SetProperty("U_Valor8", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 2 Then

                            oGeneralData.SetProperty("U_Cod12", CabImp._codigo.ToString)
                            oGeneralData.SetProperty("U_CodPorc12", CabImp._codigoPorcentaje.ToString)
                            oGeneralData.SetProperty("U_BaseImp12", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralData.SetProperty("U_Tarifa12", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralData.SetProperty("U_Valor12", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 6 Then

                            oGeneralData.SetProperty("U_CodNoi", CabImp._codigo.ToString)
                            oGeneralData.SetProperty("U_CodPorcNoi", CabImp._codigoPorcentaje.ToString)
                            oGeneralData.SetProperty("U_BaseImpNoi", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralData.SetProperty("U_TarifaNoi", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralData.SetProperty("U_ValorNoi", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 7 Then

                            oGeneralData.SetProperty("U_CodExe", CabImp._codigo.ToString)
                            oGeneralData.SetProperty("U_CodPorcExe", CabImp._codigoPorcentaje.ToString)
                            oGeneralData.SetProperty("U_BaseImpExe", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralData.SetProperty("U_TarifaExe", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralData.SetProperty("U_ValorExe", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                    ElseIf CabImp._codigo = 3 Then

                        oGeneralData.SetProperty("U_CodIce", CabImp._codigo.ToString)
                        oGeneralData.SetProperty("U_CodPorcIce", CabImp._codigoPorcentaje.ToString)
                        oGeneralData.SetProperty("U_BaseImpIce", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                        oGeneralData.SetProperty("U_TarifaIce", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                        oGeneralData.SetProperty("U_ValorIce", Convert.ToDouble(formatDecimal(CabImp._valor)))

                    End If
                Next

                oChildren = oGeneralData.Child("GS_FCDET")

                For Each detalleFC In oFactura.facturaDetalle

                    oChild = oChildren.Add
                    oChild.SetProperty("U_CodPrin", IIf(IsNothing(detalleFC._codigoPrincipal), Left(detalleFC._descripcion, 50), detalleFC._codigoPrincipal))
                    oChild.SetProperty("U_CodAuxi", IIf(IsNothing(detalleFC._codigoAuxiliar), "", detalleFC._codigoAuxiliar))
                    oChild.SetProperty("U_Descripc", Left(detalleFC._descripcion, 254))
                    oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(detalleFC._cantidad)))
                    oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(detalleFC._precioUnitario)))
                    oChild.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(detalleFC._descuento)))
                    oChild.SetProperty("U_TotSinImp", Convert.ToDouble(formatDecimal(detalleFC._precioTotalSinImpuesto)))

                    For Each detalleFCImp In detalleFC._impuestos

                        If detalleFCImp._codigo = 2 Then

                            oChild.SetProperty("U_Cod", detalleFCImp._codigo.ToString)
                            oChild.SetProperty("U_CodPorc", detalleFCImp._codigoPorcentaje.ToString)
                            oChild.SetProperty("U_BaseImp", Convert.ToDouble(formatDecimal(detalleFCImp._baseImponible)))
                            oChild.SetProperty("U_Tarifa", Convert.ToDouble(formatDecimal(detalleFCImp._tarifa)))
                            oChild.SetProperty("U_Valor", Convert.ToDouble(formatDecimal(detalleFCImp._valor)))

                        ElseIf detalleFCImp._codigo = 3 Then

                            oChild.SetProperty("U_CodIce", detalleFCImp._codigo.ToString)
                            oChild.SetProperty("U_CodPorcIce", detalleFCImp._codigoPorcentaje.ToString)
                            oChild.SetProperty("U_BaseImpIce", Convert.ToDouble(formatDecimal(detalleFCImp._baseImponible)))
                            oChild.SetProperty("U_TarifaIce", Convert.ToDouble(formatDecimal(detalleFCImp._tarifa)))
                            oChild.SetProperty("U_ValorIce", Convert.ToDouble(formatDecimal(detalleFCImp._valor)))

                        End If

                    Next

                Next

                oGeneralParams = oGeneralService.Add(oGeneralData)
                Dim DocEntryUdo = oGeneralParams.GetProperty("DocEntry")
                GuardaLog("Registro creado con exito. docentry: " + DocEntryUdo.ToString + " - nombre archivo " + nombreAr)
                Return True
            Catch ex As Exception

                GuardaLog("Error al insertar Factura Recibida con clave: " + oFactura.FacturaCabecera._claveAcceso + " - " + ex.Message.ToString() + " - nombre archivo " + nombreAr)
                Return False

            End Try

        Else
            GuardaLog("Ya se encuentra importada la factura con clave: " + oFactura.FacturaCabecera._claveAcceso.ToString + " - Proveedor " + oFactura.FacturaCabecera._RazonSocial + " - nombre archivo " + nombreAr)
            Return False
        End If



    End Function

    Public Function insertarEntidadNC(ByVal oNcredito As NotaCredito, ByVal nombreAr As String) As Boolean

        Dim oGeneralServiceNC As SAPbobsCOM.GeneralService
        Dim oGeneralDataNC As SAPbobsCOM.GeneralData
        Dim oChildNC As SAPbobsCOM.GeneralData
        Dim oChildrenNC As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParamsNC As SAPbobsCOM.GeneralDataParams
        Dim oCompanyServiceNC As SAPbobsCOM.CompanyService

        Dim _DocEntryUdoNC As String = getRSvalue("Select ""DocEntry"" from ""@GS_NC"" where ""U_ClaAcc""= '" + oNcredito.NotaCreditoCabecera._claveAcceso.ToString + "'", "DocEntry", "")

        If _DocEntryUdoNC = "0" Or _DocEntryUdoNC = "" Then

            Try
                oCompanyServiceNC = oCompany.GetCompanyService
                oGeneralServiceNC = oCompanyServiceNC.GetGeneralService("GS_NC")
                oGeneralDataNC = oGeneralServiceNC.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)


                If String.IsNullOrEmpty(oNcredito.NotaCreditoCabecera._NumeroAutorizacion) Then
                    oGeneralDataNC.SetProperty("U_NumAut", oNcredito.NotaCreditoCabecera._claveAcceso.ToString)
                Else
                    oGeneralDataNC.SetProperty("U_NumAut", oNcredito.NotaCreditoCabecera._NumeroAutorizacion.ToString)
                End If

                If String.IsNullOrEmpty(oNcredito.NotaCreditoCabecera._FechaAutorizacion) Then
                    oGeneralDataNC.SetProperty("U_FechaAut", oNcredito.NotaCreditoCabecera._fechaEmision.ToString)
                Else
                    oGeneralDataNC.SetProperty("U_FechaAut", oNcredito.NotaCreditoCabecera._FechaAutorizacion.ToString)
                End If

                oGeneralDataNC.SetProperty("U_Ruc", oNcredito.NotaCreditoCabecera._ruc)
                oGeneralDataNC.SetProperty("U_RazSoc", oNcredito.NotaCreditoCabecera._RazonSocial)
                oGeneralDataNC.SetProperty("U_ClaAcc", oNcredito.NotaCreditoCabecera._claveAcceso)
                oGeneralDataNC.SetProperty("U_Est", oNcredito.NotaCreditoCabecera._estab)
                oGeneralDataNC.SetProperty("U_PuntoEmi", oNcredito.NotaCreditoCabecera._ptoEmi)
                oGeneralDataNC.SetProperty("U_Sec", oNcredito.NotaCreditoCabecera._secuencial)
                oGeneralDataNC.SetProperty("U_FecEmi", oNcredito.NotaCreditoCabecera._fechaEmision.ToString)
                oGeneralDataNC.SetProperty("U_DirEst", IIf(IsNothing(oNcredito.NotaCreditoCabecera._dirEstablecimiento), "", oNcredito.NotaCreditoCabecera._dirEstablecimiento))
                oGeneralDataNC.SetProperty("U_ConEsp", IIf(IsNothing(oNcredito.NotaCreditoCabecera._contribuyenteEspecial), "", oNcredito.NotaCreditoCabecera._contribuyenteEspecial))
                oGeneralDataNC.SetProperty("U_RazSocComp", oNcredito.NotaCreditoCabecera._razonSocialComprador)
                oGeneralDataNC.SetProperty("U_IdenComp", oNcredito.NotaCreditoCabecera._identificacionComprador)
                oGeneralDataNC.SetProperty("U_DirComp", IIf(IsNothing(oNcredito.NotaCreditoCabecera._direccionComprador), "", oNcredito.NotaCreditoCabecera._direccionComprador))

                oGeneralDataNC.SetProperty("U_TotSinImp", Convert.ToDouble(formatDecimal(oNcredito.NotaCreditoCabecera._totalSinImpuestos)))
                oGeneralDataNC.SetProperty("U_TotDesc", Convert.ToDouble(formatDecimal(oNcredito.NotaCreditoCabecera._totalDescuento)))
                oGeneralDataNC.SetProperty("U_ImpTotal", Convert.ToDouble(formatDecimal(oNcredito.NotaCreditoCabecera._importeTotal)))

                oGeneralDataNC.SetProperty("U_CodDocMod", oNcredito.NotaCreditoCabecera._CodDocMod)
                oGeneralDataNC.SetProperty("U_NumDocMod", oNcredito.NotaCreditoCabecera._numDocModificado)
                oGeneralDataNC.SetProperty("U_FecDocMod", oNcredito.NotaCreditoCabecera._fechaEmisionDocSustento.ToString)
                'oGeneralData.SetProperty("U_FormaPago", oNcredito.NotaCreditoCabecera._formaPago.ToString)
                'oGeneralData.SetProperty("U_TotalPago", Convert.ToDouble(oNcredito.NotaCreditoCabecera._totalFormaPago))
                'oGeneralData.SetProperty("U_PlazoPago", oFactura.FacturaCabecera._plazo.ToString)
                'oGeneralData.SetProperty("U_UniTiempo", oFactura.FacturaCabecera._unidadTiempo.ToString)
                oGeneralDataNC.SetProperty("U_ValorMod", Convert.ToDouble(formatDecimal(oNcredito.NotaCreditoCabecera._valorModificacion)))
                oGeneralDataNC.SetProperty("U_Estado", "Importado")
                'oGeneralData.SetProperty("U_UniTiempo", oFactura.FacturaCabecera._unidadTiempo)

                For Each CabImp In oNcredito.NotaCreditoCabecera._impuestos
                    If CabImp._codigo = 2 Then
                        If CabImp._codigoPorcentaje = 0 Then

                            oGeneralDataNC.SetProperty("U_Cod0", CabImp._codigo.ToString)
                            oGeneralDataNC.SetProperty("U_CodPorc0", CabImp._codigoPorcentaje.ToString)
                            oGeneralDataNC.SetProperty("U_BaseImp0", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralDataNC.SetProperty("U_Tarifa0", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralDataNC.SetProperty("U_Valor0", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 8 Then

                            oGeneralDataNC.SetProperty("U_Cod8", CabImp._codigo.ToString)
                            oGeneralDataNC.SetProperty("U_CodPorc8", CabImp._codigoPorcentaje.ToString)
                            oGeneralDataNC.SetProperty("U_BaseImp8", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralDataNC.SetProperty("U_Tarifa8", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralDataNC.SetProperty("U_Valor8", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 2 Then

                            oGeneralDataNC.SetProperty("U_Cod12", CabImp._codigo.ToString)
                            oGeneralDataNC.SetProperty("U_CodPorc12", CabImp._codigoPorcentaje.ToString)
                            oGeneralDataNC.SetProperty("U_BaseImp12", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralDataNC.SetProperty("U_Tarifa12", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralDataNC.SetProperty("U_Valor12", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 6 Then

                            oGeneralDataNC.SetProperty("U_CodNoi", CabImp._codigo.ToString)
                            oGeneralDataNC.SetProperty("U_CodPorcNoi", CabImp._codigoPorcentaje.ToString)
                            oGeneralDataNC.SetProperty("U_BaseImpNoi", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralDataNC.SetProperty("U_TarifaNoi", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralDataNC.SetProperty("U_ValorNoi", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                        If CabImp._codigoPorcentaje = 7 Then

                            oGeneralDataNC.SetProperty("U_CodExe", CabImp._codigo.ToString)
                            oGeneralDataNC.SetProperty("U_CodPorcExe", CabImp._codigoPorcentaje.ToString)
                            oGeneralDataNC.SetProperty("U_BaseImpExe", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                            oGeneralDataNC.SetProperty("U_TarifaExe", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                            oGeneralDataNC.SetProperty("U_ValorExe", Convert.ToDouble(formatDecimal(CabImp._valor)))

                        End If

                    ElseIf CabImp._codigo = 3 Then

                        oGeneralDataNC.SetProperty("U_CodIce", CabImp._codigo.ToString)
                        oGeneralDataNC.SetProperty("U_CodPorcIce", CabImp._codigoPorcentaje.ToString)
                        oGeneralDataNC.SetProperty("U_BaseImpIce", Convert.ToDouble(formatDecimal(CabImp._baseImponible)))
                        oGeneralDataNC.SetProperty("U_TarifaIce", Convert.ToDouble(formatDecimal(CabImp._tarifa)))
                        oGeneralDataNC.SetProperty("U_ValorIce", Convert.ToDouble(formatDecimal(CabImp._valor)))

                    End If
                Next

                oChildrenNC = oGeneralDataNC.Child("GS_NCDET")

                For Each detalleNC In oNcredito.NotaCreditoDetalle

                    oChildNC = oChildrenNC.Add
                    oChildNC.SetProperty("U_CodPrin", IIf(IsNothing(detalleNC._codigoInterno), Left(detalleNC._descripcion, 50), detalleNC._codigoInterno))
                    oChildNC.SetProperty("U_CodAuxi", IIf(IsNothing(detalleNC._codigoAdicional), "", detalleNC._codigoAdicional))
                    oChildNC.SetProperty("U_Descripc", detalleNC._descripcion)
                    oChildNC.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(detalleNC._cantidad)))
                    oChildNC.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(detalleNC._precioUnitario)))
                    oChildNC.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(detalleNC._descuento)))
                    oChildNC.SetProperty("U_TotSinImp", Convert.ToDouble(formatDecimal(detalleNC._precioTotalSinImpuesto)))

                    For Each detalleNCImp In detalleNC._impuestos

                        If detalleNCImp._codigo = 2 Then

                            oChildNC.SetProperty("U_Cod", detalleNCImp._codigo.ToString)
                            oChildNC.SetProperty("U_CodPorc", detalleNCImp._codigoPorcentaje.ToString)
                            oChildNC.SetProperty("U_BaseImp", Convert.ToDouble(formatDecimal(detalleNCImp._baseImponible)))
                            oChildNC.SetProperty("U_Tarifa", Convert.ToDouble(formatDecimal(detalleNCImp._tarifa)))
                            oChildNC.SetProperty("U_Valor", Convert.ToDouble(formatDecimal(detalleNCImp._valor)))

                        ElseIf detalleNCImp._codigo = 3 Then

                            oChildNC.SetProperty("U_CodIce", detalleNCImp._codigo.ToString)
                            oChildNC.SetProperty("U_CodPorcIce", detalleNCImp._codigoPorcentaje.ToString)
                            oChildNC.SetProperty("U_BaseImpIce", Convert.ToDouble(formatDecimal(detalleNCImp._baseImponible)))
                            oChildNC.SetProperty("U_TarifaIce", Convert.ToDouble(formatDecimal(detalleNCImp._tarifa)))
                            oChildNC.SetProperty("U_ValorIce", Convert.ToDouble(formatDecimal(detalleNCImp._valor)))

                        End If

                    Next

                Next

                oGeneralParamsNC = oGeneralServiceNC.Add(oGeneralDataNC)
                Dim DocEntryUdoNC = oGeneralParamsNC.GetProperty("DocEntry")
                GuardaLog("Registro NC creado con exito. docentry: " + DocEntryUdoNC.ToString + " - nombre archivo " + nombreAr)
                Return True
            Catch ex As Exception

                GuardaLog("Error al insertar nota credito Recibida con clave: " + oNcredito.NotaCreditoCabecera._claveAcceso + " - " + ex.Message.ToString() + " - nombre archivo " + nombreAr)
                Return False

            End Try

        Else
            GuardaLog("Ya se encuentra importada la nota de credito con clave: " + oNcredito.NotaCreditoCabecera._claveAcceso.ToString + " - Proveedor " + oNcredito.NotaCreditoCabecera._RazonSocial + " - nombre archivo " + nombreAr)
            Return False
        End If


    End Function

    Public Function insertarEntidadRT(ByVal oRT As Retencion, ByVal nombreAr As String) As Boolean

        Dim oGeneralServiceRT As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Dim _DocEntryUdoRT As String = getRSvalue("Select ""DocEntry"" from ""@GS_RT"" where ""U_ClaAcc""= '" + oRT.RetCabecera._claveAcceso.ToString + "'", "DocEntry", "")

        If _DocEntryUdoRT = "0" Or _DocEntryUdoRT = "" Then
            Try
                oCompanyService = oCompany.GetCompanyService
                oGeneralServiceRT = oCompanyService.GetGeneralService("GS_RT")
                oGeneralData = oGeneralServiceRT.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                'oGeneralData.SetProperty("U_NumAut", IIf(IsNothing(oRT.RetCabecera._NumeroAutorizacion), oRT.RetCabecera._claveAcceso, oRT.RetCabecera._NumeroAutorizacion))
                'oGeneralData.SetProperty("U_FechaAut", IIf(IsNothing(oRT.RetCabecera._FechaAutorizacion.ToString), "", oRT.RetCabecera._FechaAutorizacion.ToString))

                If String.IsNullOrEmpty(oRT.RetCabecera._NumeroAutorizacion) Then
                    oGeneralData.SetProperty("U_NumAut", oRT.RetCabecera._claveAcceso.ToString)
                Else
                    oGeneralData.SetProperty("U_NumAut", oRT.RetCabecera._NumeroAutorizacion.ToString)
                End If

                If String.IsNullOrEmpty(oRT.RetCabecera._FechaAutorizacion) Then
                    oGeneralData.SetProperty("U_FechaAut", oRT.RetCabecera._fechaEmision.ToString)
                Else
                    oGeneralData.SetProperty("U_FechaAut", oRT.RetCabecera._FechaAutorizacion.ToString)
                End If

                oGeneralData.SetProperty("U_Ruc", oRT.RetCabecera._ruc)
                oGeneralData.SetProperty("U_RazSoc", Left(oRT.RetCabecera._RazonSocial, 99))
                oGeneralData.SetProperty("U_ClaAcc", oRT.RetCabecera._claveAcceso)
                oGeneralData.SetProperty("U_Est", oRT.RetCabecera._estab)
                oGeneralData.SetProperty("U_PuntoEmi", oRT.RetCabecera._ptoEmi)
                oGeneralData.SetProperty("U_Sec", oRT.RetCabecera._secuencial)
                oGeneralData.SetProperty("U_FecEmi", oRT.RetCabecera._fechaEmision.ToString)
                oGeneralData.SetProperty("U_DirEst", IIf(IsNothing(oRT.RetCabecera._dirEstablecimiento), "", oRT.RetCabecera._dirEstablecimiento))
                oGeneralData.SetProperty("U_ConEsp", IIf(IsNothing(oRT.RetCabecera._contribuyenteEspecial), "", oRT.RetCabecera._contribuyenteEspecial))
                oGeneralData.SetProperty("U_RazSocRet", oRT.RetCabecera._razonSocialSujetoRetenido)
                oGeneralData.SetProperty("U_IdenRet", oRT.RetCabecera._identificacionSujetoRetenido)
                oGeneralData.SetProperty("U_Periodo", oRT.RetCabecera._periodoFiscal)

                oGeneralData.SetProperty("U_Estado", "Importado")
                'oGeneralData.SetProperty("U_UniTiempo", oFactura.FacturaCabecera._unidadTiempo)


                oChildren = oGeneralData.Child("GS_RTDET")

                For Each detalleRT In oRT.RetDetalleImp

                    oChild = oChildren.Add
                    oChild.SetProperty("U_Codigo", detalleRT._codigo.ToString)
                    oChild.SetProperty("U_CodRet", detalleRT._codigoRetencion.ToString)
                    oChild.SetProperty("U_BaseImp", Convert.ToDouble(formatDecimal(detalleRT._baseImponible)))
                    oChild.SetProperty("U_PorcRet", detalleRT._porcentajeRetener.ToString)
                    oChild.SetProperty("U_ValorRet", Convert.ToDouble(formatDecimal(detalleRT._valorRetenido)))
                    oChild.SetProperty("U_CodDocSus", detalleRT._codDocSustento.ToString)
                    If String.IsNullOrEmpty(detalleRT._numDocSustento) Then
                        oChild.SetProperty("U_NumDocSus", "")
                    Else
                        oChild.SetProperty("U_NumDocSus", detalleRT._numDocSustento.ToString)
                    End If
                    If String.IsNullOrEmpty(detalleRT._fechaEmisionDocSustento) Then
                        oChild.SetProperty("U_FemiDocSus", "")
                    Else
                        oChild.SetProperty("U_FemiDocSus", detalleRT._fechaEmisionDocSustento.ToString)
                    End If


                Next

                oGeneralParams = oGeneralServiceRT.Add(oGeneralData)
                Dim DocEntryUdo = oGeneralParams.GetProperty("DocEntry")
                GuardaLog("Registro RT creado con exito. docentry: " + DocEntryUdo.ToString + " - nombre archivo " + nombreAr)
                Return True
            Catch ex As Exception

                GuardaLog("Error al insertar retencion Recibida con clave: " + oRT.RetCabecera._claveAcceso + " - " + ex.Message.ToString() + " - nombre archivo " + nombreAr)
                Return False

            End Try
        Else
            GuardaLog("Ya se encuentra importada la Retencion con clave: " + oRT.RetCabecera._claveAcceso.ToString + " - Cliente " + oRT.RetCabecera._RazonSocial + " - nombre archivo " + nombreAr)
        End If


    End Function


#Region "Conexion SAP"



    Private Function conectSAP() As Boolean
        Try
            Dim ErrCode As Long
            Dim ErrMsg As String = ""
            GuardaLog(String.Format("Conexion SAP"))
            oCompany = New SAPbobsCOM.Company()
            oCompany.Server = ConfigurationManager.AppSettings("DevServer")
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.DbServerType = ConfigurationManager.AppSettings("DevServerType")
            oCompany.LicenseServer = ConfigurationManager.AppSettings("LicenseServer")

            'oCompany.UseTrusted = False
            oCompany.UseTrusted = ConfigurationManager.AppSettings("UseTrusted")
            oCompany.DbUserName = ConfigurationManager.AppSettings("DevDBUser")
            oCompany.DbPassword = ConfigurationManager.AppSettings("DevDBPassword")
            oCompany.CompanyDB = ConfigurationManager.AppSettings("DevDatabase")
            oCompany.UserName = ConfigurationManager.AppSettings("DevSBOUser")
            oCompany.Password = ConfigurationManager.AppSettings("DevSBOPassword")

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


    'para las pruebas descomentar y colocar el proyecto servicio ivend como inicio
    'Public Sub New()

    '    ' Llamada necesaria para el diseñador.
    '    InitializeComponent()
    '    servicio_Core()

    '    ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    'End Sub


    Public Shared Function formatDecimal(ByVal numero As String) As Decimal

        Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
        Dim result As Double = 0
        Try
            If numero = "" Then
                numero = "0"
            End If
            If numero IsNot Nothing Then
                If Not numero.Contains(",") Then
                    result = Double.Parse(numero, CultureInfo.InvariantCulture)
                Else
                    result = Convert.ToDouble(numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
                    'result = Double.Parse((numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString())), CultureInfo.InvariantCulture)
                End If
            End If
        Catch e As Exception
            Try
                'result = Convert.ToDouble(numero)
                result = Double.Parse(numero, CultureInfo.InvariantCulture)
            Catch
                Try
                    'result = Convert.ToDouble(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."))
                    result = Double.Parse(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."), CultureInfo.InvariantCulture)
                Catch
                    Throw New Exception("Wrong string-to-double format")
                End Try
            End Try
        End Try
        Return result

    End Function

    Public Function validaSiEmpiezaPunto(ByVal valor As String) As String
        Dim _valor As String = valor

        Try
            If valor.StartsWith(".") Then
                _valor = "0" & valor

            End If
        Catch ex As Exception

            _valor = Nothing
        End Try

        Return _valor
    End Function

End Class

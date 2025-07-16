Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Principal
Imports System.Text
Imports System.Threading
Imports System.Timers
Imports System.Xml.Serialization
Imports SAPbobsCOM

Public Class SSGENERADORDEPAGOS_Atcotrans

    Public oCompanyGP As SAPbobsCOM.Company
    'Public oCompanyBase2 As SAPbobsCOM.Company
    Private Const NombreServicioGP As String = "SS_GENERADOR_PAGOS"
    Private Const VersionServicioGP As String = "1.0"
    Private Const CodigoPaisGP As String = "EC"
    Dim estadoLicenciasGP As Boolean = False
    Dim versionTributariaGP As String = ""
    Private oTimerGP As System.Timers.Timer = Nothing
    'Private oTimerSincroItem As System.Timers.Timer = Nothing
    Private blnInicio_GP As Boolean = False
    'Private blnInicio_sincroItem As Boolean = False
    'Private ConsultaSN As String = "SS_CONSULTA_SOCIONEGOCIO"

    Private CONEXION_SQL As SqlConnection
    Private CONEXION_HANA As Odbc.OdbcConnection

    Dim listaParametros() As WS_LICENCIA.ClsConfigValores = Nothing

    Dim _WS_RecepcionGP As String = ""
    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential
    Private listaRetRecibidas As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
    Public oFuncionesB1GP As Functions.FuncionesB1
    Private oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
    Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
    Dim oFacturaVenta As Entidades.FacturaVenta

    Dim ExisteFacRel As Boolean = False
    Dim sCardCode As String = ""
    Dim _sRUC As String = ""
    Dim NumDocPago As String = ""
    Dim customCulture As CultureInfo
    Dim Fechalaboral5toDia As Date?
    Dim _Fechalaboral5toDia As Date?
    Dim ValidoDiasLab As Boolean
    Dim SaldoPendiente As Decimal = 0

    Private serviceThread As Thread


    'Public oEntidades As Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA


    'Sub New()
    '    'descomentar el sub New para pruebas y limpiar y generar nuevamente para que no de error
    '    'Llamada necesaria para el diseñador.
    '    InitializeComponent()
    '    servicio_Core()



    '    'Agregue cualquier inicialización después de la llamada a InitializeComponent().

    'End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        oCompanyGP = Nothing
        'servicio_Core()
        serviceThread = New Thread(AddressOf servicio_Core)
        serviceThread.Start()

    End Sub

    Protected Overrides Sub OnStop()
        Try

            serviceThread.Abort()
            GuardaLog("Hilo serviceThread finalizado.")

            oTimerGP.Dispose()
            GuardaLog("Timer oTimerGP liberado.")
            oTimerGP.Stop()
            GuardaLog("Timer oTimerGP detenido.")



        Catch ex As Exception
            GuardaLog("Error en OnStop " + ex.Message.ToString)
        Finally
            serviceThread.Abort()
            oTimerGP.Stop()
        End Try


    End Sub

    Protected Overrides Sub OnPause()
        oTimerGP.Stop()
    End Sub

    Protected Overrides Sub OnContinue()
        oTimerGP.Start()
    End Sub

    Private Sub servicio_Core()

        Try

            GuardaLog("Servicio Generador Pagos Iniciado, procediendo a conectarse..!")
            GuardaLog("El motor de la Base de Datos Detectado es de tipo  : " + ConfigurationManager.AppSettings("DevServerType"))
            GuardaLog("ruta contabilizados:" + ConfigurationManager.AppSettings("RutaLogProcesados").ToString())
            GuardaLog("ruta no contabilizados:" + ConfigurationManager.AppSettings("RutaLogNOProcesados").ToString())

            If conectSAP() Then

                If ValidarLicencia() And estadoLicenciasGP Then
                    ObtenerVariables()

                    Try
                        Dim windowsIdentity As WindowsIdentity = WindowsIdentity.GetCurrent()
                        Dim userNameSer As String = windowsIdentity.Name

                        GuardaLog("Usuario utilizado por el servicio: " + userNameSer.ToString)

                    Catch ex As Exception
                        GuardaLog("Error al obtener nombre de usuario de windows: " + ex.Message.ToString())
                    End Try

                    Try

                        customCulture = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
                        customCulture.NumberFormat.NumberDecimalSeparator = "."
                        customCulture.NumberFormat.NumberGroupSeparator = ","
                        GuardaLog("Cultura configurada: Separador decimal= " & customCulture.NumberFormat.NumberDecimalSeparator & " Separador de Miles= " & customCulture.NumberFormat.NumberGroupSeparator)


                    Catch ex As Exception
                        GuardaLog("Error al obtener setear separadores de decimles y miles: " + ex.Message.ToString())
                    End Try

                    If (oCompanyGP.Connected) Then
                        GuardaLog(String.Format("Conexion SAP exitosa, CompanyName={0} ,DataBase={1}  ", oCompanyGP.CompanyName, oCompanyGP.CompanyDB))
                    End If

                    'GETPArametros_INIT()
                    Dim timer As Int64 = ConfigurationManager.AppSettings("timer") * 1000
                    'oTimer = New System.Timers.Timer(timer)
                    oTimerGP = New System.Timers.Timer(timer)
                    GuardaLog("Temporizador para contabilizacion de retenciones recibidas: " + oTimerGP.ToString())
                    oTimerGP.AutoReset = True
                    oTimerGP.Enabled = False
                    AddHandler oTimerGP.Elapsed, AddressOf oTimerGP_Elapsed
                    oTimerGP.Start()

                    'Proceso_Generar_Pagos()
                    GuardaLog("Temporizadores  Iniciados...")
                End If


            End If

        Catch ex As Exception
            'EventLog.WriteEntry("Error en servicio_Core: " & ex.Message, EventLogEntryType.Error)
            GuardaLog("Error en servicio_Core: " & ex.Message.ToString)
            Me.Stop()
        End Try


    End Sub

    Public Sub GuardaLog(texto As String)

        If ConfigurationManager.AppSettings("GuardaLog").ToString().Equals("1") Then

            Dim nombreHilo = IIf(Thread.CurrentThread.Name Is Nothing, "SN", Thread.CurrentThread.Name)

            Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "Generador Pagos " & Date.Now.ToString("dd_MM_yyyy") & " - " & nombreHilo.ToString & ".txt"
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

    Public Sub GuardaLogNoContabilizado(texto As String)

        'If ConfigurationManager.AppSettings("GuardaLog").ToString().Equals("1") Then

        Dim nombreHilo = IIf(Thread.CurrentThread.Name Is Nothing, "SN", Thread.CurrentThread.Name)
        Dim _RUTA = ConfigurationManager.AppSettings("RutaLogNOProcesados").ToString()
        'Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "Generador Pagos " & Date.Now.ToString("dd_MM_yyyy") & " - " & nombreHilo.ToString & ".txt"
        Dim sRuta As String = _RUTA & "Generador Pagos No Contabilizados " & Date.Now.ToString("dd_MM_yyyy") & " - " & nombreHilo.ToString & ".txt"
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
            GuardaLog("Error al guaardr dato no contabilizados:" + ex.Message.ToString)
        Finally

        End Try

        'End SyncLock

        'End If

    End Sub

    Public Sub GuardaLogContabilizado(texto As String)

        'If ConfigurationManager.AppSettings("GuardaLog").ToString().Equals("1") Then

        Dim nombreHilo = IIf(Thread.CurrentThread.Name Is Nothing, "SN", Thread.CurrentThread.Name)
        Dim _RUTA = ConfigurationManager.AppSettings("RutaLogProcesados").ToString()
        'Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "Generador Pagos " & Date.Now.ToString("dd_MM_yyyy") & " - " & nombreHilo.ToString & ".txt"
        Dim sRuta As String = _RUTA & "Generador Pagos Contabilizados " & Date.Now.ToString("dd_MM_yyyy") & " - " & nombreHilo.ToString & ".txt"
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

        ' End If

    End Sub

    Private Function conectSAP() As Boolean
        Try
            Dim ErrCode As Long
            Dim ErrMsg As String = ""
            GuardaLog(String.Format("Conexion SAP"))
            Try
                If IsNothing(oCompanyGP) Then
                    oCompanyGP = New SAPbobsCOM.Company()
                    GuardaLog("Instancia oCompanyGP creada por primera vez")
                End If
            Catch ex As Exception
                GuardaLog("Error al querer instanciar el objetooCompanyGP " + ex.Message.ToString())
            End Try



            If Not IsNothing(oCompanyGP) Then
                If oCompanyGP.Connected Then
                    GuardaLog("Company en estado Conectado a SAP BO " + oCompanyGP.CompanyDB.ToString())
                    Return True
                End If
            End If


            oCompanyGP.DbServerType = ConfigurationManager.AppSettings("DevServerType")
            oCompanyGP.UseTrusted = ConfigurationManager.AppSettings("UseTrusted")
            oCompanyGP.CompanyDB = ConfigurationManager.AppSettings("DevDatabase")
            oCompanyGP.UserName = ConfigurationManager.AppSettings("DevSBOUser")
            oCompanyGP.Password = ConfigurationManager.AppSettings("DevSBOPassword")

            Try
                If CInt(ConfigurationManager.AppSettings("SAP_VERSION")) < 10 Then
                    oCompanyGP.Server = ConfigurationManager.AppSettings("DevServer")
                    oCompanyGP.LicenseServer = ConfigurationManager.AppSettings("LicenseServer")
                    oCompanyGP.DbUserName = ConfigurationManager.AppSettings("DevDBUser")
                    oCompanyGP.DbPassword = ConfigurationManager.AppSettings("DevDBPassword")
                Else
                    oCompanyGP.Server = ConfigurationManager.AppSettings("DevServer")
                    'oCompanyGP.SLDServer = ConfigurationManager.AppSettings("LicenseServer")
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

            ErrCode = oCompanyGP.Connect()

            If ErrCode <> 0 Then
                oCompanyGP.GetLastError(ErrCode, ErrMsg)
                GuardaLog("Error al conectarse a SAP ,funcion : oCompanyGP.Connect :" + ErrCode.ToString() + " - " + ErrMsg.ToString)
                Return False
            Else
                ' GuardaLog("Conectado a SAP BO" + oCompanyGP.CompanyDB.ToString())
                Return True
            End If

        Catch ex As Exception
            GuardaLog("Error al conectarse a SAP , funcion: conectSAP , EX :" + ex.Message.ToString())
            Return False
        End Try

    End Function

    Private Sub oTimerGP_Elapsed(sender As Object, e As ElapsedEventArgs)

        Try
            If blnInicio_GP = False Then
                blnInicio_GP = True

                Proceso_Generar_Pagos()

                blnInicio_GP = False
            End If


        Catch ex As Exception

            GuardaLog("Error al tratar de Invocar la funcion de procesar_docSincro en el Ciclo oTimer_Elapsed " & ex.Message.ToString())

        End Try

    End Sub

    Private Function ValidarLicencia() As Boolean

        ' TIENE LICENCIA
        GuardaLog(String.Format("Validando la Licencia para Producto : {0} Version {1}", NombreServicioGP, VersionServicioGP))

        ''MANEJO LICENCIA WEB
        Dim wsSSLIC As New WS_LICENCIA.Licencia

        Dim msgg As String = ""
        Dim oLicencia As Licencia = Nothing
        Dim RucCliente As String = ""

        Try
            Dim param_ambiente As Integer = 0
            Dim param_inhouse As Boolean = False
            Dim respLIC As WS_LICENCIA.CLsRespLic = Nothing

            Dim URLWS As String = ConfigurationManager.AppSettings("WsLicencia")
            wsSSLIC.Url = URLWS

            RucCliente = ConfigurationManager.AppSettings("RucCompañia")

            If String.IsNullOrEmpty(URLWS) Then
                GuardaLog("No se encontro parametrización de WS Control, verificar por favor!")
            End If

            If String.IsNullOrEmpty(RucCliente) Then
                RucCliente = "000000000"
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TipoWsLicencia")) Then
                param_ambiente = IIf(ConfigurationManager.AppSettings("TipoWsLicencia") = "PRUEBAS", 1, 2)

            End If

            param_inhouse = False

            Dim SApdatos As New WS_LICENCIA.DatosSap
            With SApdatos
                .RucEmpresa = RucCliente
                .DireccionIPSERVER = oCompanyGP.Server
                .NombreDB = oCompanyGP.CompanyDB
                .NombreProducto = NombreServicioGP
                .VersionProducto = VersionServicioGP
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
                    GuardaLog("Serializando, Parametros de Busqueda")

                    Dim x As XmlSerializer = New XmlSerializer(GetType(WS_LICENCIA.DatosSap))
                    Dim writer As TextWriter = New StreamWriter(sRuta)
                    x.Serialize(writer, SApdatos)
                    writer.Close()
                    GuardaLog("Serializado, Parametros de Busqueda" + sRuta)
                End If

            Catch ex As Exception
                GuardaLog("Serializado. Error: " + ex.Message.ToString())
            End Try

            SetProtocolosdeSeguridad()

            respLIC = wsSSLIC.ValidarLicencia(SApdatos, msgg)
            If Not IsNothing(respLIC) Then
                oLicencia = New Licencia
                oLicencia.Opcion = respLIC.TipoLic
                oLicencia.NombreBaseSAP = oCompanyGP.CompanyDB
                oLicencia.Estado = CBool(respLIC.Estado)
                oLicencia.validoHasta = 1000
                'oLicencia.VersionTributaria = respLIC.VersionTributaria
                listaParametros = respLIC.ListaUrlWS
            End If

            If oLicencia Is Nothing Then

                estadoLicenciasGP = False
                versionTributariaGP = ""
                GuardaLog("Se Encontro un inconveniente al Consultar la Licencia, mensaje Referencia: " & msgg)
            Else

                estadoLicenciasGP = oLicencia.Estado
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

    Private Sub Proceso_Generar_Pagos()
        Try

            ClearMemory()

            If ConsultarRetencionesRecibidas() Then

                If Not oCompanyGP.Connected Then

                    Dim ErrCode = oCompanyGP.Connect()

                    If ErrCode <> 0 Then

                        GuardaLog("Error al conectarse a SAP ,funcion  : " + oCompanyGP.GetLastErrorDescription)

                    End If

                End If

                If oCompanyGP.Connected Then


                    Dim hilos_GP As Integer = IIf(ConfigurationManager.AppSettings("TotalHilosGP") = "", 1, CInt(ConfigurationManager.AppSettings("TotalHilosGP")))

                    Try
                        GuardaLog("Iniciando proceso para calculo de documentos por hilo, total documentos " + listaRetRecibidas.Count.ToString)
                        GuardaLog(String.Format("Hilos Configurados {0}", hilos_GP.ToString))
                        Dim listaHilosGP As New List(Of Thread)
                        'Dim oListaFacturaVenta As New List(Of Entidades.FacturaVenta)

                        If listaRetRecibidas.Count > 0 Then

                            For he = 1 To hilos_GP

                                'oListaFacturaVenta.Add(New Entidades.FacturaVenta With {.Name = "FacturasHilo" & CStr(he)})
                                listaHilosGP.Add(New Thread(AddressOf trabajo_hilo_GP) With {.Name = "hilo_GP" & CStr(he)})


                            Next

                        End If
                        'GuardaLog(String.Format("1 lista hilos {0}", hilos_sincronizacionSN.ToString))
                        Dim sum = 0
                        '----------------------------------------------------
                        Dim promhilo_GP = Decimal.Truncate(listaRetRecibidas.Count / hilos_GP)
                        Dim promhilo_GP_R = CInt(listaRetRecibidas.Count Mod hilos_GP)
                        For Each he As Thread In listaHilosGP

                            If promhilo_GP = 0 Then

                                he.Start(listaRetRecibidas)

                                Exit For
                            End If

                            GuardaLog("RangoHilo " & he.Name.ToString & "de " + sum.ToString + " a " + (sum + (promhilo_GP + promhilo_GP_R)).ToString)
                            sum = sum + promhilo_GP + promhilo_GP_R
                            he.Start(listaRetRecibidas.GetRange(0, promhilo_GP + promhilo_GP_R))
                            listaRetRecibidas.RemoveRange(0, promhilo_GP + promhilo_GP_R)

                            promhilo_GP_R = 0

                        Next

                        'espero a que los hilos finalicen
                        For Each he As Thread In listaHilosGP
                            he.Join()

                        Next

                        listaHilosGP.Clear()

                        listaHilosGP = Nothing

                    Catch ex As Exception
                        GuardaLog("Error al crear Hilos para generar pagos ==> " + ex.Message.ToString)
                    End Try

                End If

            End If

        Catch ex As Exception
            GuardaLog("Error general en la funcion de procesar_docSincro " & ex.Message.ToString())
        End Try
    End Sub

    Private Declare Auto Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal procHandle As IntPtr, ByVal min As Int32, ByVal max As Int32) As Boolean
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

    Private Sub ObtenerVariables()

        Variables_Globales.WS_Recepcion = ConfigurationManager.AppSettings("WsRecepcion")
        Variables_Globales.WS_RecepcionCambiarEstado = ConfigurationManager.AppSettings("WsCambiarEstado")
        Variables_Globales.WS_RecepcionCargaEstados = ConfigurationManager.AppSettings("Estados")
        Variables_Globales.WS_RecepcionClave = ConfigurationManager.AppSettings("ClaveRecepcion")

        Variables_Globales.SALIDA_POR_PROXY = ConfigurationManager.AppSettings("SalidaProxy")
        Variables_Globales.Proxy_puerto = ConfigurationManager.AppSettings("ProxyPuerto")
        Variables_Globales.Proxy_IP = ConfigurationManager.AppSettings("ProxyIp")
        Variables_Globales.Proxy_Usuario = ConfigurationManager.AppSettings("ProxyUsuario")
        Variables_Globales.Proxy_Clave = ConfigurationManager.AppSettings("ProxyClave")

        Variables_Globales.CampoNumRetencion = ConfigurationManager.AppSettings("CampoNumRetencion")
        Variables_Globales.FechaEmisionRetencion = ConfigurationManager.AppSettings("FechaEmisionRetencion")
        Variables_Globales.FechaEmisionRetencionP = ConfigurationManager.AppSettings("FechaEmisionRetencionP")

        Variables_Globales.Nombre_Proveedor = ConfigurationManager.AppSettings("NombreProveedor")



        Variables_Globales._ValidarFechasCTK = ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "ValidarFechasCTK")
        Variables_Globales._vgFechaEmisionRetencion = ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "FechaEmisionRetencion")
        Variables_Globales._vgFechaEmisionRetencionP = ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "FechaEmisionRetencionP")

        Variables_Globales.CantUltmsDias = IIf(ConfigurationManager.AppSettings("CantUltmsDia") = "", "3", ConfigurationManager.AppSettings("CantUltmsDia"))
        Variables_Globales.CantDiasLab = IIf(ConfigurationManager.AppSettings("CantDiasLab") = "", "5", ConfigurationManager.AppSettings("CantDiasLab"))

        Variables_Globales.ContSaldoPendMenor = IIf(ConfigurationManager.AppSettings("ContabilizaSaldoPendienteMenor") = "", "NO", ConfigurationManager.AppSettings("ContabilizaSaldoPendienteMenor"))
        Variables_Globales.CuentaSaldoFavor = ConfigurationManager.AppSettings("CuentaSaldoFavor")
        Variables_Globales.IdSeriePR = ConfigurationManager.AppSettings("IdSeriePR")

    End Sub

    Private Function ConsultarRetencionesRecibidas() As Boolean

        GuardaLog("Consuta de documentos recibidos")

        If Variables_Globales.WS_Recepcion = "" Then
            GuardaLog("No existe parametrización del Web Services de Recepcion, Revisar archivo de configuracion")
            Return False
        End If

        'Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA

        WS.Url = Variables_Globales.WS_Recepcion

        'MANEJO DE PROXY
        Dim SALIDA_POR_PROXY As String = ""
        SALIDA_POR_PROXY = Variables_Globales.SALIDA_POR_PROXY
        GuardaLog("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString)
        Dim Proxy_puerto As String = ""
        Dim Proxy_IP As String = ""
        Dim Proxy_Usuario As String = ""
        Dim Proxy_Clave As String = ""
        If SALIDA_POR_PROXY = "Y" Then

            Proxy_puerto = Variables_Globales.Proxy_puerto
            Proxy_IP = Variables_Globales.Proxy_IP
            Proxy_Usuario = Variables_Globales.Proxy_Usuario
            Proxy_Clave = Variables_Globales.Proxy_Clave

            GuardaLog("Proxy_puerto : " + Proxy_puerto.ToString)
            GuardaLog("Proxy_IP : " + Proxy_IP.ToString)
            GuardaLog("Proxy_Usuario : " + Proxy_Usuario.ToString)
            GuardaLog("Proxy_Clave : " + Proxy_Clave.ToString)

            If Not Proxy_puerto = "" Then
                proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
            Else
                proxyobject = New System.Net.WebProxy(Proxy_IP)
            End If
            cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

            proxyobject.Credentials = cred

            WS.Proxy = proxyobject
            WS.Credentials = cred

        End If
        ' END MANEJO DE PROXY

        Dim oFiltrosRecepcionECRT As New Entidades.wsEDoc_ConsultaRecepcion.ClsBusqueda
        oFiltrosRecepcionECRT.CiaTipoAlojamientoKey = Variables_Globales.WS_RecepcionClave
        oFiltrosRecepcionECRT.RucProveedor = Nothing
        oFiltrosRecepcionECRT.Estado = IIf(Variables_Globales.WS_RecepcionCargaEstados = "", "1", Variables_Globales.WS_RecepcionCargaEstados)
        oFiltrosRecepcionECRT.FechaEmisionDesde = Nothing
        oFiltrosRecepcionECRT.FechaEmisionHasta = Nothing
        'oFiltrosRecepcionECRT.NumDocumento = "001-064-000557451"

        Dim z As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
        SetProtocolosdeSeguridad()
        Dim mensaje = ""

        z = WS.ConsultarRetencion_CabeceraBuscar(oFiltrosRecepcionECRT, mensaje).ToList

        If z.Count > 0 Then
            GuardaLog("Retenciones Recibidas : " + z.Count.ToString)

            listaRetRecibidas = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))
            Return True

        Else
            GuardaLog("Nº Total Documentos Recibidos :" + z.Count.ToString())
            Return False
        End If

    End Function

    Private Sub trabajo_hilo_GP(LisRet As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))

        Try
            ''For Each oFactura As Factura In listaFCs.GetRange(NumIni, RegistrosXPaginas)
            Dim Factura As String = ""
            Dim NumCuota As Integer = 0
            Dim DocEntryRERecibida_UDO As String = "0"
            Dim sDocEntryPreliminar As String = "0"
            Dim mensajeError As String
            Dim nombrehilo As String = ""
            Dim bandera As Boolean = False
            Dim BaseImponible
            Dim banderaBases As Boolean

            Dim BaseRenta As Decimal
            Dim BaseIVA As Decimal

            WS.Url = Variables_Globales.WS_Recepcion
            nombrehilo = Thread.CurrentThread.Name


            Thread.CurrentThread.CurrentCulture = customCulture
            Thread.CurrentThread.CurrentUICulture = customCulture

            GuardaLog(String.Format("Ejecutando funcion trabajo_hilo_GP,{0} documentos a Proce x hilo {1} ", LisRet.Count.ToString, nombrehilo))

            Dim oListaFacturaVenta As New List(Of Entidades.FacturaVenta)


            For Each oRet In LisRet

                oListaFacturaVenta.Clear()
                Dim QueryExisteCliente As String = ""
                If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                    Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                    Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE _
                    Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        QueryExisteCliente = "SELECT ""CardCode"" FROM " + oCompanyGP.CompanyDB + ".""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + oRet.Ruc + "'"
                    Else
                        QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + oRet.Ruc + "'"
                    End If
                ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                        QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""U_DOCUMENTO"" = '" + oRet.Ruc + "'"
                    Else

                        QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND U_DOCUMENTO = '" + oRet.Ruc + "'"
                    End If
                End If
                BaseRenta = 0
                BaseIVA = 0
                Dim numdocretenr = ""
                Dim FechaEmisionRetencion = oRet.FechaEmision
                banderaBases = False
                bandera = False
                GuardaLog("Query info cardcode: " + QueryExisteCliente.ToString)
                sCardCode = getRSvalue(QueryExisteCliente, "CardCode", "")
                _sRUC = oRet.Ruc

                oRetencion = Nothing
                Dim mensaje = ""
                Dim numfac = ""

                If sCardCode <> "" Then
                    GuardaLog("Codigo SN encontrado :" + sCardCode.ToString())
                    SetProtocolosdeSeguridad()
                    oRetencion = WS.ConsultarRetencion_Detalle(Variables_Globales.WS_RecepcionClave, oRet.IdRetencion, mensaje)
                    numfac = oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString

                    If ValidarSiExiste(oRetencion, sCardCode, numfac) Then

                        'validar si la fecha de emision de la factura esta dentro de los ultimos dias del mes anterior
                        'If Calculo5toDiaLaborable(oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener) And ValidoDiasLab = False Then
                        ValidoDiasLab = False
                        Dim distinctDetalleRet = oRetencion.ENTDetalleRetencion.Select(Function(r) New With {Key r.NumDocRetener, Key r.FechaEmisionDocRetener}).Distinct().ToList()

                        'For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRetencion.ENTDetalleRetencion
                        For Each DetalleRet In distinctDetalleRet
                            Fechalaboral5toDia = Calculo5toDiaLaborable(DetalleRet.FechaEmisionDocRetener)

                            If Not IsNothing(Fechalaboral5toDia) Then
                                Dim fecha5dia As Date = Fechalaboral5toDia
                                Dim fechaActual = CInt(Date.Now.ToString("yyyyMMdd"))
                                Dim _fecha5dia = CInt(fecha5dia.ToString("yyyyMMdd"))
                                If fechaActual <= _fecha5dia Then
                                    ValidoDiasLab = True
                                    GuardaLog("Si esta dentro del rango, fecha calculada: " + CDate(Fechalaboral5toDia).ToString("yyyy-MM-dd"))
                                Else
                                    ValidoDiasLab = False
                                    GuardaLogNoContabilizado("Razon Social: " + oRet.RazonSocial.ToString & vbCrLf &
                                             "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                                             "Numero de factura: " + DetalleRet.NumDocRetener.ToString().Substring(0, 3) + "-" + DetalleRet.NumDocRetener.ToString().Substring(3, 3) + "-" + DetalleRet.NumDocRetener.Substring(6, 9) & vbCrLf &
                                             "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                                             "Motivo: La fecha calculada: " + CDate(Fechalaboral5toDia).ToString("yyyy-MM-dd") + " no esta dentro de los " + Variables_Globales.CantDiasLab + " dias habiles configurados")
                                    Exit For
                                End If
                            End If

                        Next

                        If ValidoDiasLab = False And Not IsNothing(Fechalaboral5toDia) Then
                            Continue For
                        End If

                        'validacion de que si ya se encuentra contabilizados.
                        If ValidarRetContabilizadaCancelada(oRetencion, sCardCode, numfac) = False Then


                            If ValidarRetContabilizadas(oRetencion, sCardCode, numfac) = False Then

                                mensajeError = ""

                                'validacion de dias permitidos para contabilizar ret del mes anterior
                                'por mes
                                Variables_Globales.Dias = ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "DiasValidarProcesoLote")
                                GuardaLog("Dias Parametrizados: " + Variables_Globales.Dias.ToString)

                                If CDate(oRet.FechaAutorizacion).Month <= Date.Now.Month And oRet.FechaEmision.Month <= Date.Now.Month And oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener.Month < Date.Now.Month And ValidoDiasLab = False Then
                                    If Date.Now.Day <= Variables_Globales.Dias Then
                                        bandera = True
                                    Else
                                        bandera = False
                                    End If
                                    'cuando es enero y recibe doc de diciembre
                                ElseIf CDate(oRet.FechaAutorizacion).Year > Date.Now.Year Then
                                    If CDate(oRet.FechaAutorizacion).Month >= Date.Now.Month And oRet.FechaEmision.Month >= Date.Now.Month And oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener.Month > Date.Now.Month And ValidoDiasLab = False Then
                                        If Date.Now.Day <= Variables_Globales.Dias Then
                                            bandera = True
                                        Else
                                            bandera = False
                                        End If
                                    Else
                                        bandera = True
                                    End If
                                Else
                                    bandera = True
                                End If

                                If bandera = True Then

                                    'VALIDA LAS BSES PARA RENTA E IVA, SALDO PENDIENTE Y SI EXISTE LA FACTURA
                                    If ValidarBasesRetRec(oRetencion, sCardCode, mensajeError, oListaFacturaVenta) = True Then

                                        Dim IdRet = oRetencion.IdRetencion
                                        GuardaLog("IdRet recibida: " + oRetencion.IdRetencion.ToString)
                                        mensajeError = ""

                                        If Guarda_DocumentoRecibido_RE_Udo(DocEntryRERecibida_UDO, oRetencion, sCardCode) Then
                                            If CrearPagoRecibido(DocEntryRERecibida_UDO, sDocEntryPreliminar, oRetencion, oListaFacturaVenta) Then
                                                ' LE CAMBIA EL ESTADO A LA FACTURA UDO A DOCFINAL
                                                ActualizadoEstado_DocumentoRecibido_RE(DocEntryRERecibida_UDO, "docFinal", 1)
                                                ' ACTUALIZA EL CAMPO SINCRO A 1, ESTE CAMPO IDENTIFICA QUE YA ESTA SINCRONIZADA EN SAP

                                                ' MARCA EL DOCUMENTO COMO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                                MarcarVisto(Integer.Parse(IdRet), 2, mensaje, DocEntryRERecibida_UDO)
                                                GuardaLog("Clave Retencion: " + oRet.ClaveAcceso + " id Retencion: " + Integer.Parse(IdRet).ToString + " Proceso terminado Exitosamente!")
                                                GuardaLogContabilizado("Razon Social:" + oRet.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                                                "Numero de factura: " + numfac.ToString().Substring(0, 3) + "-" + numfac.ToString().Substring(3, 3) + "-" + numfac.Substring(6, 9) & vbCrLf &
                                                "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                                                "Numero de Paogo Recibido: " + NumDocPago.ToString & vbCrLf &
                                                "Motivo: Contabilizado correctamente ")
                                                NumDocPago = ""
                                                oListaFacturaVenta.Clear()
                                                GuardaLog("Lista Limpia..!")
                                                oFacturaVenta = Nothing
                                                Continue For
                                            Else
                                                Continue For
                                            End If
                                        Else
                                            Continue For
                                        End If
                                    Else
                                        oListaFacturaVenta.Clear()
                                        GuardaLog("Lista Limpia..!")
                                        oFacturaVenta = Nothing
                                        Continue For
                                    End If
                                Else
                                    GuardaLogNoContabilizado("Razon Social: " + oRet.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                                                             "Numero de factura: " + numfac.ToString().Substring(0, 3) + "-" + numfac.ToString().Substring(3, 3) + "-" + numfac.Substring(6, 9) & vbCrLf &
                                                             "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                                                             "Motivo: La fecha de autorizacion de la retenccion: " + CDate(oRet.FechaAutorizacion).ToString("yyyy-MM-dd") + " no se encuentra dentro del rango de dias parametrizados(" + Variables_Globales.Dias.ToString + ")")

                                    Continue For
                                End If
                            Else

                                'GuardaLogNoContabilizado("Razon Social:" + oRet.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                                '     "Numero de factura: " + numfac.ToString().Substring(0, 3) + "-" + numfac.ToString().Substring(3, 3) + "-" + numfac.Substring(6, 9) & vbCrLf &
                                '     "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                                '     "Motivo: Ya existe un pago recibido por tarjetaa de credito apuntando a la factura.") 'se cambia dentro de la funcion ya ue existen corporacion el reosado donde la retencion tiene varias facturas

                                Continue For

                            End If

                        Else

                            'GuardaLogNoContabilizado("Razon Social:" + oRet.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                            '      "Numero de factura: " + numfac.ToString().Substring(0, 3) + "-" + numfac.ToString().Substring(3, 3) + "-" + numfac.Substring(6, 9) & vbCrLf &
                            '      "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                            '      "Motivo: La retencion ya se registró y se encuentra cancelada en SAP") 'se cambia dentro de la funcion ya ue existen corporacion el reosado donde la retencion tiene varias facturas

                            Continue For

                        End If

                    Else

                        'GuardaLogNoContabilizado("Razon Social:" + oRet.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                        '      "Numero de factura: " + numfac.ToString().Substring(0, 3) + "-" + numfac.ToString().Substring(3, 3) + "-" + numfac.Substring(6, 9) & vbCrLf &
                        '      "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                        '      "Motivo: No se encuentro la factura con numeracion " + numfac.ToString + " para el socio de negocio") 'se cambia dentro de la funcion ya ue existen corporacion el reosado donde la retencion tiene varias facturas
                        Continue For

                    End If

                Else

                    GuardaLogNoContabilizado("Razon Social:" + oRet.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + oRet.Establecimiento + "-" + oRet.PuntoEmision + "-" + oRet.Secuencial & vbCrLf &
                              "Clave de acceso: " + oRet.ClaveAcceso & vbCrLf &
                              "Motivo: Ruc " + oRet.Ruc.ToString + " no encontrado ")
                    Continue For
                End If


            Next


        Catch ex As Exception
            GuardaLog("Error en try general funcion trabajo_hilo_GP :" + ex.Message().ToString())
        End Try


    End Sub
    Public Function ValidarRetContabilizadas(_oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef CardCode As String, ByRef NumDocRetener As String) As Boolean

        Try
            Dim query = ""

            Dim SecRet = _oRetencion.Secuencial.ToString
            Dim SerieRet = _oRetencion.Establecimiento.ToString + _oRetencion.PuntoEmision.ToString
            Dim total = _oRetencion.TotalRetencion

            Dim NumDocRetenerUnicos As List(Of String) = _oRetencion.ENTDetalleRetencion.Select(Function(x) x.NumDocRetener).Distinct().ToList()

            'For Each odetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oRetencion.ENTDetalleRetencion
            For Each NumDocSutento In NumDocRetenerUnicos

                If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        query = " Select COUNT(*) as   ""Contador"" "
                        query += " From ORCT T0 inner Join RCT3 T1"
                        query += " On T0.""DocEntry""=T1.""DocNum"" WHERE"
                        query += " T1.""VoucherNum""='" + NumDocSutento.ToString + "'"
                        query += " And T0.""ObjType"" ='24'"
                        query += " And T0.""CardCode""='" + CardCode + "'"
                        'query += " And t0.""DocTotal""=" + total.ToString("###0.00") 'SE COMENTA DEBIDO A QUE PUEDE EMITIRSE OTRA RETENCION A LA MISMA FACTURA Y ASI MISMO PUEDE SER A LA MISMA PERO CON OTRO VALOR
                        'query += " And right('000000000'||T1.""U_CXS_NUM_RETE"",9)='" + SecRet + "'"
                        'query += " And T1.""U_CXS_SER_PTO_RET""='" + SerieRet + "'"
                        query += " And T0.""Canceled""='N'"


                    Else

                        query = " Select COUNT(*) as Contador"
                        query += " From ORCT T0 inner Join RCT3 T1"
                        query += " On T0.""DocEntry""=T1.""DocNum"" WHERE"
                        query += " T1.""VoucherNum""='" + NumDocSutento.ToString + "'"
                        query += " And T0.""ObjType"" ='24'"
                        query += " And T0.""CardCode""='" + CardCode + "'"
                        'query += " And t0.""DocTotal""=" + total.ToString("###0.00")
                        'query += " And right('000000000'+T1.""U_CXS_NUM_RETE"",9)='" + SecRet + "'"
                        'query += " And T1.""U_CXS_SER_PTO_RET""='" + SerieRet + "'"
                        query += " And T0.""Canceled""='N'"

                    End If

                    GuardaLog("Query para validar si la factura ya tiene retencion: " + query + "")

                    Dim resultado = CInt(getRSvalue(query, "Contador", "0"))


                    If resultado > 0 Then
                        GuardaLog("Resultado de la consulta:" + resultado.ToString)
                        GuardaLogNoContabilizado("Razon Social:" + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                     "Numero de factura: " + NumDocSutento.ToString().Substring(0, 3) + "-" + NumDocSutento.ToString().Substring(3, 3) + "-" + NumDocSutento.Substring(6, 9) & vbCrLf &
                                     "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                     "Motivo: Ya existe un pago recibido por tarjetaa de credito apuntando a la factura.")
                        Return True
                    Else
                        GuardaLog("Resultado de la consulta: " + resultado.ToString)
                        Continue For
                    End If


                End If
            Next

            Return False

        Catch ex As Exception
            GuardaLog("Error en funcion ValidarRetContabilizadas: " + ex.Message.ToString)
            Return True
        End Try


    End Function


    Public Function ValidarFechaEmiFac(_oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef CardCode As String, ByRef NumDocRetener As String) As Boolean

        Try
            Dim query = ""
            'Dim NumDocRetener = ""

            'For Each odetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oRetencion.ENTDetalleRetencion
            '    NumDocRetener = odetalle.NumDocRetener.ToString
            '    Exit For
            'Next

            Dim SecRet = _oRetencion.Secuencial.ToString
            Dim SerieRet = _oRetencion.Establecimiento.ToString + _oRetencion.PuntoEmision.ToString
            Dim total = _oRetencion.TotalRetencion

            If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    query = " Select COUNT(*) as   ""Contador"" "
                    query += " From ORCT T0 inner Join RCT3 T1"
                    query += " On T0.""DocEntry""=T1.""DocNum"" WHERE"
                    query += " T1.""VoucherNum""='" + NumDocRetener + "'"
                    query += " And T0.""ObjType"" ='24'"
                    query += " And T0.""CardCode""='" + CardCode + "'"
                    'query += " And t0.""DocTotal""=" + total.ToString("###0.00") 'SE COMENTA DEBIDO A QUE PUEDE EMITIRSE OTRA RETENCION A LA MISMA FACTURA Y ASI MISMO PUEDE SER A LA MISMA PERO CON OTRO VALOR
                    'query += " And right('000000000'||T1.""U_CXS_NUM_RETE"",9)='" + SecRet + "'"
                    'query += " And T1.""U_CXS_SER_PTO_RET""='" + SerieRet + "'"
                    query += " And T0.""Canceled""='N'"


                Else

                    query = " Select COUNT(*) as Contador"
                    query += " From ORCT T0 inner Join RCT3 T1"
                    query += " On T0.""DocEntry""=T1.""DocNum"" WHERE"
                    query += " T1.""VoucherNum""='" + NumDocRetener + "'"
                    query += " And T0.""ObjType"" ='24'"
                    query += " And T0.""CardCode""='" + CardCode + "'"
                    'query += " And t0.""DocTotal""=" + total.ToString("###0.00")
                    'query += " And right('000000000'+T1.""U_CXS_NUM_RETE"",9)='" + SecRet + "'"
                    'query += " And T1.""U_CXS_SER_PTO_RET""='" + SerieRet + "'"
                    query += " And T0.""Canceled""='N'"

                End If

                GuardaLog("Query para validar si la factura ya tiene retencion: " + query + "")

                Dim resultado = CInt(getRSvalue(query, "Contador", "0"))


                If resultado > 0 Then
                    GuardaLog("Resultado de la consulta:" + resultado.ToString)
                    Return True
                Else
                    GuardaLog("Resultado de la consulta: " + resultado.ToString)

                End If

            End If

            Return False

        Catch ex As Exception
            GuardaLog("Error en funcion ValidarRetContabilizadas: " + ex.Message.ToString)
            Return True
        End Try


    End Function


    Public Function ValidarSiExiste(_oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef CardCode As String, ByRef NumDocRetener As String) As Boolean

        Try
            Dim queryExiste = ""

            Dim NumDocRetenerUnicos As List(Of String) = _oRetencion.ENTDetalleRetencion.Select(Function(x) x.NumDocRetener).Distinct().ToList()

            'For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oRetencion.ENTDetalleRetencion.Distinct()
            For Each NumDocSutento In NumDocRetenerUnicos

                If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                        queryExiste = " SELECT count(""DocEntry"") as ""ContadorF"" "
                        queryExiste += " FROM " + oCompanyGP.CompanyDB + ".""OINV"" WHERE ""CardCode"" = '" + CardCode + "'"
                        queryExiste += " AND ""U_SER_EST"" = '" & NumDocSutento.ToString().Substring(0, 3) & "'"
                        queryExiste += " AND ""U_SER_PE"" = '" & NumDocSutento.ToString().Substring(3, 3) & "'"
                        queryExiste += " AND ""FolioNum"" = " & Integer.Parse(NumDocSutento.Substring(6, 9))
                        queryExiste += " AND ""DocStatus""='O' AND ""CANCELED""='N' "


                    Else

                        queryExiste = " SELECT count(DocEntry) as ""ContadorF"" "
                        queryExiste += " FROM ""OINV"" WHERE ""CardCode"" = '" + CardCode + "'"
                        queryExiste += " AND U_SER_EST = '" & NumDocSutento.ToString().Substring(0, 3) & "'"
                        queryExiste += " AND U_SER_PE = '" & NumDocSutento.ToString().Substring(3, 3) & "'"
                        queryExiste += " AND FolioNum = " & Integer.Parse(NumDocSutento.Substring(6, 9))
                        queryExiste += " AND DocStatus='O' AND CANCELED='N' "

                    End If

                    GuardaLog("Query para validar si existe la factura: " + queryExiste.ToString)
                    'Factura = getRSvalue(sQueryFactura, "DocEntry", "")
                    Dim resultado = CInt(getRSvalue(queryExiste, "ContadorF", "0"))


                    If resultado = 0 Then
                        GuardaLog("Resultado de la consulta existe factura: " + resultado.ToString)
                        GuardaLogNoContabilizado("Razon Social:" + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                              "Numero de factura: " + NumDocSutento.ToString().Substring(0, 3) + "-" + NumDocSutento.ToString().Substring(3, 3) + "-" + NumDocSutento.Substring(6, 9) & vbCrLf &
                              "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                              "Motivo: No se encuentro la factura con numeracion " + NumDocSutento.ToString + " para el socio de negocio")
                        Return False
                    Else
                        GuardaLog("Resultado de la consulta existe factura: " + resultado.ToString)
                        Continue For
                    End If

                End If


            Next

            Return True
            'si es mayor a  es porue si existe la factura

        Catch ex As Exception
            GuardaLog("Error en funcion ValidarSiExiste: " + ex.Message.ToString)
            Return False
        End Try


    End Function


    Public Function ValidarRetContabilizadaCancelada(_oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef CardCode As String, ByRef NumDocRetener As String) As Boolean

        Try
            Dim query = ""
            'Dim NumDocRetener = ""
            Dim SecRet = _oRetencion.Secuencial.ToString
            Dim SerieRet = _oRetencion.Establecimiento.ToString + _oRetencion.PuntoEmision.ToString
            Dim total = _oRetencion.TotalRetencion

            Dim NumDocRetenerUnicos As List(Of String) = _oRetencion.ENTDetalleRetencion.Select(Function(x) x.NumDocRetener).Distinct().ToList()

            'For Each odetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oRetencion.ENTDetalleRetencion
            For Each NumDocSustento In NumDocRetenerUnicos

                If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        query = " Select COUNT(*) as   ""ContadorC"" "
                        query += " From ORCT T0 inner Join RCT3 T1"
                        query += " On T0.""DocEntry""=T1.""DocNum"" WHERE"
                        query += " T1.""VoucherNum""='" + NumDocSustento.ToString + "'"
                        query += " And T0.""ObjType"" ='24'"
                        query += " And T0.""CardCode""='" + CardCode + "'"
                        query += " And t0.""DocTotal""=" + total.ToString("###0.00")
                        query += " And right('000000000'||T1.""U_CXS_NUM_RETE"",9)='" + SecRet + "'"
                        query += " And T1.""U_CXS_SER_PTO_RET""='" + SerieRet + "'"
                        query += " And T0.""Canceled""='Y'"


                    Else

                        query = " Select COUNT(*) as ContadorC"
                        query += " From ORCT T0 inner Join RCT3 T1"
                        query += " On T0.""DocEntry""=T1.""DocNum"" WHERE"
                        query += " T1.""VoucherNum""='" + NumDocSustento.ToString + "'"
                        query += " And T0.""ObjType"" ='24'"
                        query += " And T0.""CardCode""='" + CardCode + "'"
                        query += " And t0.""DocTotal""=" + total.ToString("###0.00")
                        query += " And right('000000000'+T1.""U_CXS_NUM_RETE"",9)='" + SecRet + "'"
                        query += " And T1.""U_CXS_SER_PTO_RET""='" + SerieRet + "'"
                        query += " And T0.""Canceled""='Y'"

                    End If


                End If

                GuardaLog("Query para validar si ya existe la retencion Cancelada: " + query + "")

                Dim resultado = CInt(getRSvalue(query, "ContadorC", "0"))


                If resultado = 0 Then
                    GuardaLog("Resultado de la consulta pago cancelado:" + resultado.ToString)
                    Continue For
                Else
                    GuardaLog("Resultado de la consulta pago cancelado: " + resultado.ToString)
                    GuardaLogNoContabilizado("Razon Social:" + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                  "Numero de factura: " + NumDocSustento.ToString().Substring(0, 3) + "-" + NumDocSustento.ToString().Substring(3, 3) + "-" + NumDocSustento.Substring(6, 9) & vbCrLf &
                                  "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                  "Motivo: La retencion ya se registró y se encuentra cancelada en SAP")
                    Return True ' es decir ue si se enconetro una factura cancelada
                End If

            Next


            Return False 'es decir ue no se encontro la retencion cancelada




        Catch ex As Exception
            GuardaLog("Error en funcion ValidarRetContabilizadas: " + ex.Message.ToString)
            Return True
        End Try


    End Function

    Public Function ValidarBasesRetRec(_oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef CardCode As String, ByRef mensajeError As String, ByRef oListaFacturaVenta As List(Of Entidades.FacturaVenta)) As Boolean

        Try
            Dim numeroRet As String = _oRetencion.Establecimiento + _oRetencion.PuntoEmision + _oRetencion.Secuencial
            'Dim query = ""
            Dim sQueryFactura = ""
            Dim NumDocRetener As String = ""
            Dim SecRet = _oRetencion.Secuencial.ToString
            Dim SerieRet = _oRetencion.Establecimiento.ToString + _oRetencion.PuntoEmision.ToString
            Dim total = _oRetencion.TotalRetencion
            Dim baseRetRec As Decimal = 0
            Dim IvaRetRec As Decimal = 0
            Dim baseFac As Decimal = 0
            Dim IvaFac As Decimal = 0
            Dim Factura As String = ""
            Dim listaInfoFacturas As New List(Of InfoFactura)()
            Dim conta As Integer = 0
            Dim NumCuota As Integer = 0

            For Each odetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oRetencion.ENTDetalleRetencion
                NumDocRetener = odetalle.NumDocRetener.ToString
                If Not IsNothing(NumDocRetener) Then
                    If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"", "
                            sQueryFactura += " (select SUM(""LineTotal"") from INV1 where ""DocEntry""=""OINV"".""DocEntry"")-""OINV"".""DiscSum"" as ""SubTotal"", ""VatSum"", "
                            sQueryFactura += " ifnull(""DocTotal"",0)-ifnull(""PaidToDate"",0) as ""SaldoPendiente"" "
                            sQueryFactura += " FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND ""U_SER_EST"" = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND ""U_SER_PE"" = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "

                            'sQueryFacturaAnt = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""ODPI"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ""ODPI"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            'sQueryFacturaAnt += " AND ""U_SER_EST"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            'sQueryFacturaAnt += " AND ""U_SER_PE"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            'sQueryFacturaAnt += " AND ""FolioNum"" = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            'sQueryFacturaAnt += " AND ""DocStatus""='O' AND ""CANCELED""='N' "

                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(InstlmntID) from INV6 where DocEntry=OINV.DocEntry and Status='O') as ""Cuota"" , "
                            sQueryFactura += " (select SUM(""LineTotal"") from INV1 where ""DocEntry""=""OINV"".""DocEntry"" )-""OINV"".""DiscSum"" as ""SubTotal"", ""VatSum"", "
                            sQueryFactura += " isnull(""Doctotal"",0)-isnull(""PaidToDate"",0) as ""SaldoPendiente"" "
                            sQueryFactura += " FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND U_SER_EST = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND U_SER_PE = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND FolioNum = " & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "

                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM ""OINV"" A INNER JOIN "
                            sQueryFactura += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND B.""BeginStr"" = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.""EndStr"" = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND A.""FolioNum"" =" & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "


                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM OINV A INNER JOIN "
                            sQueryFactura += " NNM1 B ON A.Series = B.Series "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND B.BeginStr = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.EndStr = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND A.FolioNum =" & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "


                        End If
                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM ""OINV"" A "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND right(A.""NumAtCard"",17) = '" & NumDocRetener.ToString().Substring(0, 3) _
                                                    & "-" & NumDocRetener.ToString().Substring(3, 3) & "-" & NumDocRetener.ToString().Substring(6, 9) & "'"
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "



                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM OINV A  "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND right(A.NumAtCard,17) = '" & NumDocRetener.ToString().Substring(0, 3) _
                                                    & "-" & NumDocRetener.ToString().Substring(3, 3) & "-" & NumDocRetener.ToString().Substring(6, 9) & "'"
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "



                        End If
                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM ""OINV"" A INNER JOIN "
                            sQueryFactura += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND B.""BeginStr"" = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.""EndStr"" = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.""DocNum"",7),'0',' ')),' ','0') =" & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "


                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM OINV A INNER JOIN "
                            sQueryFactura += " NNM1 B ON A.Series = B.Series "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND B.BeginStr = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.EndStr = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.DocNum,7),'0',' ')),' ','0') =" & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "

                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND ""U_SS_Est"" = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND ""U_SS_Pemi"" = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(NumDocRetener.Substring(6, 9))
                            sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "

                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND U_SS_Est = '" & NumDocRetener.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND U_SS_Pemi = '" & NumDocRetener.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND FolioNum = " & Integer.Parse(NumDocRetener.ToString().Substring(6, 9))
                            sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "

                        End If
                    End If

                    GuardaLog("Query info factura: " + sQueryFactura.ToString)
                    Factura = getRSvalue(sQueryFactura, "DocEntry", "")
                    If Factura = "" Then
                        Factura = "0"
                    End If
                    NumCuota = CInt(getRSvalue(sQueryFactura, "Cuota", ""))
                    GuardaLog("Query Factura Relacionada cuota:" + NumCuota.ToString)

                    If Not Factura = "0" Then
                        oFacturaVenta = New Entidades.FacturaVenta
                        oFacturaVenta.DocEntry = Factura
                        oFacturaVenta.ValorARetener = Convert.ToDouble(odetalle.ValorRetenido)
                        oFacturaVenta.Cuota = NumCuota
                        Dim query As System.Collections.Generic.IEnumerable(Of Entidades.FacturaVenta)
                        query = oListaFacturaVenta.Where(Function(q As Entidades.FacturaVenta) q.DocEntry = Factura)
                        If query.Count() > 0 Then
                            query.Single().ValorARetener += Convert.ToDouble(odetalle.ValorRetenido)
                        Else
                            oListaFacturaVenta.Add(oFacturaVenta)
                        End If
                        ExisteFacRel = True
                        'oFacturaVenta = Nothing
                        'END GUARDO LA INFO DE LA FACTURA Y EL VALOR A RETENER, PARA AL CREAR EL PAGO DESCONTARLE EL VALOR DE LA RETENCIÓN A CADA FACTURA
                    Else

                        GuardaLog("No existe factura: " + NumDocRetener + " query:" + sQueryFactura.ToString())
                        oFacturaVenta = Nothing
                        GuardaLogNoContabilizado("Razon Social: " + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                          "Numero de factura: " + NumDocRetener.ToString().Substring(0, 3) + "-" + NumDocRetener.ToString().Substring(3, 3) + "-" + NumDocRetener.Substring(6, 9) & vbCrLf &
                                          "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                          "Motivo: No se encuentra la factura con numeracion " + NumDocRetener)
                        'uardaLog("No se encontro la factura: " + oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString + " de la retencion: " + oRetencion.Establecimiento + oRetencion.PuntoEmision + oRetencion.Secuencial + " Clave de Acceso: " + oRetencion.ClaveAcceso)
                        Return False

                    End If

                    'Dim _DocEntryFac As System.Collections.Generic.IEnumerable(Of InfoFactura)
                    '_DocEntryFac = listaInfoFacturas.Where(Function(q As InfoFactura) q.DocEntry = Factura)

                    Dim VAL = listaInfoFacturas.Any(Function(f) f.DocEntry = Factura)

                    If VAL = False Then
                        baseFac = CDec(getRSvalue(sQueryFactura, "SubTotal", "0"))
                        IvaFac = CDec(getRSvalue(sQueryFactura, "VatSum", "0"))
                        SaldoPendiente = CDec(getRSvalue(sQueryFactura, "SaldoPendiente", "0"))
                    Else
                        baseFac = 0
                        IvaFac = 0
                    End If

                    If odetalle.Codigo = 1 Then 'RENTA

                        Dim Infofactura As New InfoFactura(Factura, NumDocRetener, numeroRet, baseFac, IvaFac, odetalle.BaseImponible, 0)
                        listaInfoFacturas.Add(Infofactura)

                    ElseIf odetalle.Codigo = 2 Then 'IVA
                        Dim Infofactura As New InfoFactura(Factura, NumDocRetener, numeroRet, baseFac, IvaFac, 0, odetalle.BaseImponible)
                        listaInfoFacturas.Add(Infofactura)

                    ElseIf odetalle.Codigo = 6 Then 'ISD
                        'banderaBases = True
                        mensajeError = "Retencion ISD"
                        Return False

                    End If
                    conta += 1


                    If SaldoPendiente < _oRetencion.TotalRetencion And Variables_Globales.ContSaldoPendMenor = "NO" Then
                        'mensajeError = "El saldo pendiente " + SaldoPendiente.ToString("###0.00") + " de la factura " + NumDocRetener.ToString() + " es menor al total de la retencion recibida " + _oRetencion.TotalRetencion.ToString("###0.00")
                        GuardaLogNoContabilizado("Razon Social: " + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                          "Numero de factura: " + NumDocRetener.ToString().Substring(0, 3) + "-" + NumDocRetener.ToString().Substring(3, 3) + "-" + NumDocRetener.Substring(6, 9) & vbCrLf &
                                          "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                          "Motivo: El saldo pendiente " + SaldoPendiente.ToString("###0.00") + " de la factura " + NumDocRetener.ToString() + " es menor al total de la retencion recibida " + _oRetencion.TotalRetencion.ToString("###0.00"))
                        Return False
                    End If

                Else
                    GuardaLogNoContabilizado("Razon Social: " + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                          "Numero de factura: " + NumDocRetener.ToString().Substring(0, 3) + "-" + NumDocRetener.ToString().Substring(3, 3) + "-" + NumDocRetener.Substring(6, 9) & vbCrLf &
                                          "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                          "Motivo: No contiene un numero de documento sustento")
                    Return False
                End If

            Next

            Dim facturasAgrupadas = listaInfoFacturas.
            GroupBy(Function(f) New With {Key f.DocEntry, Key f.NumFactura, Key f.NumRet}).
            Select(Function(g) New InfoFactura(
                g.Key.DocEntry,
                g.Key.NumFactura,
                g.Key.NumRet,
                g.Sum(Function(f) f.Subtotal),
                g.Sum(Function(f) f.Iva),
                g.Sum(Function(f) f.RetRenta),
                g.Sum(Function(f) f.RetIva)))

            For Each _InfoFacturas In facturasAgrupadas

                Dim sumaSubtotalFac As Decimal = _InfoFacturas.Subtotal
                Dim sumaIvaFac As Decimal = _InfoFacturas.Iva
                Dim RentaRec As Decimal = _InfoFacturas.RetRenta
                Dim RentaIvaRec As Decimal = _InfoFacturas.RetIva

                If RentaRec > 0 Then
                    If sumaSubtotalFac <> RentaRec Then
                        mensajeError = "La base recibida(Renta) " + RentaRec.ToString("###0.00") + " es diferente del subtotal de la factura " + sumaSubtotalFac.ToString("###0.00")
                        GuardaLogNoContabilizado("Razon Social: " + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                              "Numero de factura: " + NumDocRetener.ToString().Substring(0, 3) + "-" + NumDocRetener.ToString().Substring(3, 3) + "-" + NumDocRetener.Substring(6, 9) & vbCrLf &
                                              "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                              "Motivo:La base recibida(Renta) " + RentaRec.ToString("###0.00") + " es diferente del subtotal de la factura " + sumaSubtotalFac.ToString("###0.00"))
                        Return False
                    End If
                End If

                If RentaIvaRec > 0 Then
                    If sumaIvaFac <> RentaIvaRec Then
                        mensajeError = "La base recibida(IVA) " + RentaIvaRec.ToString("###0.00") + " es diferente del impuesto de la factura " + sumaIvaFac.ToString("###0.00")
                        GuardaLogNoContabilizado("Razon Social: " + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                              "Numero de factura: " + NumDocRetener.ToString().Substring(0, 3) + "-" + NumDocRetener.ToString().Substring(3, 3) + "-" + NumDocRetener.Substring(6, 9) & vbCrLf &
                                              "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                              "Motivo: La base recibida(IVA) " + RentaIvaRec.ToString("###0.00") + " es diferente del impuesto de la factura " + sumaIvaFac.ToString("###0.00"))
                        Return False
                    End If
                End If
            Next
            listaInfoFacturas.Clear()
            Return True

        Catch ex As Exception
            GuardaLog("Error en funcion ValidarBasesRetRec: " + ex.Message.ToString)
            Return False
        End Try


    End Function


    Public Function Guarda_DocumentoRecibido_RE_Udo(ByRef DocEntryFacturaRecibida_UDO As String, ByRef _oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef CardCode As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim _ClaveAcceso As String = ""
        Dim _numDocRet As String = ""
        Try


            Dim numDoc As String = _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial
            _ClaveAcceso = _oRetencion.ClaveAcceso

            GuardaLog("Grabando Retencion con clave: " + _oRetencion.ClaveAcceso + " - numero de doc: " + numDoc + " - Cod Cliente: " + CardCode)

            oCompanyService = oCompanyGP.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralData.SetProperty("Code", conta)
            oGeneralData.SetProperty("U_RUC", _oRetencion.Ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(_oRetencion.RazonSocial.ToString(), 99))
            oGeneralData.SetProperty("U_CardCode", CardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", oForm.Items.Item("lbMapp").Specific.Value.ToString())
            oGeneralData.SetProperty("U_ClaAcc", _oRetencion.ClaveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", _oRetencion.AutorizacionSRI.ToString())
            oGeneralData.SetProperty("U_FecAut", _oRetencion.FechaAutorizacion.ToString())
            oGeneralData.SetProperty("U_NumDoc", numDoc.ToString())
            oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString())
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(formatDecimal(_oRetencion.TotalRetencion.ToString())))
            oGeneralData.SetProperty("U_IdGS", _oRetencion.IdRetencion.ToString())
            oGeneralData.SetProperty("U_Sincro", 0)
            oGeneralData.SetProperty("U_Estado", "docPreliminar")


            oChildren = oGeneralData.Child("GS0_RER")
            For Each detalleRet As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oRetencion.ENTDetalleRetencion

                Dim ejeFiscal As String = _oRetencion.PeriodoFiscal

                oChild = oChildren.Add
                oChild.SetProperty("U_CodRet", "Factura" + " - " + detalleRet.CodigoRetencion.ToString())
                If IsNothing(detalleRet.NumDocRetener) Then
                    _numDocRet = "0"
                Else
                    _numDocRet = detalleRet.NumDocRetener.ToString
                End If
                oChild.SetProperty("U_NumDocRe", _numDocRet)
                oChild.SetProperty("U_Fecha", detalleRet.FechaEmisionDocRetener.ToString())
                oChild.SetProperty("U_pFiscal", ejeFiscal.ToString())
                oChild.SetProperty("U_Base", Convert.ToDouble(formatDecimal(detalleRet.BaseImponible.ToString())))
                If detalleRet.CodigoRetencion = "1" Then
                    oChild.SetProperty("U_Impuesto", "RENTA")
                Else
                    oChild.SetProperty("U_Impuesto", "IVA")
                End If
                oChild.SetProperty("U_Porcent", Convert.ToDouble(formatDecimal(detalleRet.PorcentajeRetener.ToString())))
                oChild.SetProperty("U_valorR", Convert.ToDouble(formatDecimal(detalleRet.ValorRetenido.ToString())))

            Next
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            GuardaLog("PRR " + _oRetencion.ClaveAcceso.ToString() + " Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString())
            'rsboApp.StatusBar.SetText(NombreAddon + " - Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Catch ex As Exception
            GuardaLog("Error al Guarda_DocumentoRecibido_RE_Udo " + _ClaveAcceso + " Ocurrior un error al crear registro de Pago Recibido(Retencion) Recibida UDO: " + ex.Message.ToString())
            GuardaLogNoContabilizado("Razon Social: " + _oRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + _oRetencion.Establecimiento + "-" + _oRetencion.PuntoEmision + "-" + _oRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + _numDocRet.ToString().Substring(0, 3) + "-" + _numDocRet.ToString().Substring(3, 3) + "-" + _numDocRet.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + _oRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: Error al guardar informacion en udo: " + ex.Message.ToString)
            Return False
        End Try
    End Function

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

    Private Function CrearPagoRecibido(DocEntry_UDO As String, ByRef DocEntry_Preliminar As String, ByRef ObjetoRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef oListaFacturaVenta As List(Of Entidades.FacturaVenta)) As Boolean

        If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
            Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
            Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
            Or Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

            Return CrearPagoRecibido_E_O(DocEntry_Preliminar, DocEntry_UDO, ObjetoRetencion, oListaFacturaVenta)

        ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

            'Return CrearPagoRecibido_Exxis_seidor(DocEntry_Preliminar, DocEntry_UDO)
            'Return CrearPagoRecibido_S(sDocEntryPreliminar, DocEntryRERecibida_UDO)
        ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then

            'Return CrearPagoRecibido_TopManage(DocEntry_Preliminar, DocEntry_UDO)
        End If
        Return True ' eliminr linea
    End Function

    Private Function CrearPagoRecibido_E_O(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String, ByRef ObjetoRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef oListaFacturaVenta As List(Of Entidades.FacturaVenta)) As Boolean

        Try

            Dim sQueryCB As String = ""
            'sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardCode""= '" + oCardCode.ToString + "' " se comento ya que el codigo cambia para el caso de crear un pago tipo proveedor
            sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardType"" = 'C' and ""LicTradNum""= '" + ObjetoRetencion.Ruc.ToString + "' "
            Dim clienteBancario As String = getRSvalue(sQueryCB, "U_SSCLIENTEBANCO", "")
            GuardaLog("Validar Cliente Bancario - query " + sQueryCB + "Resultado :" + clienteBancario.ToString())
            If clienteBancario = "SI" Then
                Dim estado As Boolean
                estado = CrearPagoRecibido_E_OCB(sDocEntryPreliminar, DocEntryRERecibida_UDO, ObjetoRetencion)
                Return estado

                'FINCLIENFTEBANCARIO
            ElseIf clienteBancario = "PROVEEDOR" Then
                Dim estado As Boolean
                estado = CrearPagoRecibido_E_OProveedor(sDocEntryPreliminar, DocEntryRERecibida_UDO, ObjetoRetencion)
                Return estado
            Else
                Dim estado As Boolean
                GuardaLog("Pago recibido normal")
                estado = CrearPagoRecibido_E_ONormal(sDocEntryPreliminar, DocEntryRERecibida_UDO, ObjetoRetencion, oListaFacturaVenta)
                Return estado
            End If

        Catch ex As Exception

            GuardaLog("Exepcion en funcion CrearPagoRecibido_E_O " + ex.Message.ToString)

        End Try




    End Function

    Private Function CrearPagoRecibido_E_OCB(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String, ByRef ObjetoRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCBMP As Integer = 1
        Try

            'CLIENTE BANCARIO 

            GuardaLog("Creando Pago Recibido Tipo Cuenta(Retencion) ruc: " + ObjetoRetencion.Ruc + " con clave: " + ObjetoRetencion.ClaveAcceso)

            vPay = oCompanyGP.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rAccount

            vPay.DocCurrency = "USD"

            If Variables_Globales.FechaEmisionRetencion = "Y" Then
                vPay.DocDate = ObjetoRetencion.FechaEmision
                vPay.TaxDate = ObjetoRetencion.FechaEmision
                vPay.DueDate = ObjetoRetencion.FechaEmision
            ElseIf Variables_Globales.FechaEmisionRetencionP = "Y" Then
                vPay.DocDate = ObjetoRetencion.FechaEmision
                vPay.TaxDate = ObjetoRetencion.FechaEmision
            Else

                vPay.DocDate = Date.Now
                vPay.TaxDate = Date.Now
                vPay.DueDate = Date.Now
            End If

            vPay.DocRate = 0

            'AGREGAR DETALLE DEL PAGO
            ' OBTENCION CUENTA CONTABLE
            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" Where ""CardCode"" =  '" + sCardCode + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardCode =  '" + sCardCode + "'"
            End If
            FormatCodeProveedor = getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            Dim FormatCode As String = ""
            Dim sQueryAcctCode As String = ""
            If FormatCodeProveedor = "" Then
                FormatCode = Functions.VariablesGlobales._Cuenta_RE
            Else
                FormatCode = FormatCodeProveedor
            End If

            If FormatCode = "" Then
                GuardaLog("REE " + ObjetoRetencion.ClaveAcceso + " ERROR - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!")
                Return False
            End If
            If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode + "'"
            End If
            Dim Cuenta As String = getRSvalue(sQueryAcctCode, "AcctCode", "")
            ' END OBTENCION CUENTA CONTABLE
            Try
                vPay.AccountPayments.AccountCode = Cuenta
            Catch ex As Exception
                GuardaLog("Agregando cuenta" + ex.Message.ToString + Cuenta.ToString())
            End Try

            'vPay.AccountPayments.AccountName = NombreCuentaRetencionCB
            Try
                vPay.AccountPayments.SumPaid = ObjetoRetencion.TotalRetencion
            Catch ex As Exception
                GuardaLog("Agregando total retencion " + ex.Message.ToString + ObjetoRetencion.TotalRetencion.ToString())
            End Try

            Try
                vPay.AccountPayments.Add()
                vPay.AccountPayments.SetCurrentLine(1)
            Catch ex As Exception
                GuardaLog("error al agregar linea pago recibido " + ex.Message.ToString)
            End Try

            'END AGREGAR DETALLE DEL PAGO


            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In ObjetoRetencion.ENTDetalleRetencion

                If oDetalle.ValorRetenido > 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle.Codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
                            oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    ElseIf oDetalle.Codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    ElseIf oDetalle.Codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    GuardaLog("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString())

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    GuardaLog("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString())

                    Dim TypeCode As Integer = CrTypeCode
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        GuardaLog("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString())
                    End If

                    vPay.CreditCards.CreditAcct = CuentaRetencion
                    vPay.CreditCards.CreditCard = CodRetencion
                    vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    vPay.CreditCards.CreditCardNumber = ObjetoRetencion.Secuencial
                    vPay.CreditCards.CreditSum = oDetalle.ValorRetenido
                    vPay.CreditCards.FirstPaymentSum = ObjetoRetencion.TotalRetencion

                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        Try
                            vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion
                        Catch ex As Exception
                            GuardaLog("agregando fecha autroizacion : " + ObjetoRetencion.FechaAutorizacion.ToString())
                        End Try


                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If

                        Catch ex As Exception
                            GuardaLog("agregando fecha VoucherNum : " + oDetalle.NumDocRetener.ToString())
                        End Try
                        '
                        Try
                            If checkCampoBD("RCT3", "MONTO_BASE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                            End If
                        Catch ex As Exception
                            GuardaLog("agregando monto base : " + oDetalle.BaseImponible.ToString())
                        End Try
                        '
                        Try
                            If checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                            End If
                        Catch ex As Exception
                            GuardaLog("agregando monto base : " + oDetalle.BaseImponible.ToString())
                        End Try
                        '
                        Try
                            If checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = ObjetoRetencion.Secuencial
                            End If
                        Catch ex As Exception
                            GuardaLog("agregando numero de retencion : " + ObjetoRetencion.Secuencial.ToString())
                        End Try
                        '
                        Try
                            If checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = ObjetoRetencion.AutorizacionSRI
                            End If
                        Catch ex As Exception
                            GuardaLog("agregando numero de autorizacion de retencion : " + ObjetoRetencion.AutorizacionSRI.ToString())
                        End Try
                        '
                        Try
                            If checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If
                        Catch ex As Exception
                            GuardaLog("agregando est y punto de emision : " + ObjetoRetencion.Establecimiento.ToString() + ObjetoRetencion.PuntoEmision)
                        End Try
                        Try
                            If checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = sCardCode
                            End If
                        Catch ex As Exception
                            GuardaLog("agregando est y punto de emision : " + sCardCode.ToString)
                        End Try

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision + ObjetoRetencion.Secuencial
                        If checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = ObjetoRetencion.FechaAutorizacion
                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        Try
                            vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion
                        Catch ex As Exception
                            GuardaLog("agregando fecha autroizacion : " + ObjetoRetencion.FechaAutorizacion.ToString())
                        End Try


                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If

                        Catch ex As Exception
                            GuardaLog("agregando NumDocRetener : " + oDetalle.NumDocRetener.ToString())
                        End Try
                        If checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = ObjetoRetencion.Secuencial
                        End If
                        If checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                        End If
                        If checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = sCardCode
                        End If

                        If checkCampoBD("RCT3", "SS_NombreSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_NombreSN").Value = Left(ObjetoRetencion.RazonSocial.ToString, 100)
                        End If

                        If String.IsNullOrEmpty(Variables_Globales.CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = ObjetoRetencion.Secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Variables_Globales.CampoNumRetencion).Value = ObjetoRetencion.Secuencial
                        End If
                    End If


                    If checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    Try
                        If checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                        End If
                    Catch ex As Exception
                        GuardaLog("agregando docentry udo : " + DocEntryRERecibida_UDO.ToString)
                    End Try

                    Try
                        vPay.CreditCards.Add()
                        vPay.CreditCards.SetCurrentLine(secuencial)
                        secuencial += 1
                    Catch ex As Exception
                        GuardaLog("agregando lineas medio de pago : " + ex.Message.ToString)
                    End Try


                End If
            Next
            Try
                RetVal = vPay.Add()
            Catch ex As Exception
                GuardaLog("erro al agregar pago recibido tipo cuenta : " + ex.Message.ToString)
            End Try

            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode.ToString() + ObjetoRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        GuardaLog("Serializando...")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    GuardaLog("EEROR " + ex.Message)
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso)

                oCompanyGP.GetLastError(ErrCode, ErrMsg)

                GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString())

                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode.ToString() + ObjetoRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        GuardaLog("Serializando...")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    GuardaLog("EEROR serializar" + ex.Message)
                End Try
                oCompanyGP.GetNewObjectCode(sDocEntryPreliminar)
                GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString())
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar, ObjetoRetencion.ClaveAcceso)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso)
            GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + "Error:" + ex.Message.ToString())
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function

    Private Function CrearPagoRecibido_E_OProveedor(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String, ByRef ObjetoRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCB As Integer = 1
        Dim secuencialCBMP As Integer = 1
        'Dim fechaVencRtMP As Date
        Try

            'Dim vPay As SAPbobsCOM.Documents
            GuardaLog("Creando Pago Recibido Tipo Proveedor(Retencion) ruc: " + ObjetoRetencion.Ruc + " con clave: " + ObjetoRetencion.ClaveAcceso)

            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" WHERE ""CardType"" = 'C' and ""LicTradNum""= '" + _sRUC.ToString + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardType = 'C' and LicTradNum =  '" + _sRUC.ToString + "'"
            End If


            FormatCodeProveedor = getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            Dim CUENTA As String = getRSvalue("SELECT ""AcctCode"" FROM OACT WHERE  ""ActId""= '" + FormatCodeProveedor + "'", "AcctCode", "")

            vPay = oCompanyGP.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rSupplier
            vPay.CardCode = sCardCode
            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.DocCurrency = "USD"
            'vPay.DocDate = DateSerisal(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now
            'vPay.DocDate = _oDocumento.FechaAutorizacion
            'vPay.TaxDate = _oDocumento.FechaEmision
            If Variables_Globales.FechaEmisionRetencion = "Y" Then
                vPay.DocDate = ObjetoRetencion.FechaEmision
                vPay.DueDate = ObjetoRetencion.FechaEmision
                vPay.TaxDate = ObjetoRetencion.FechaEmision
            ElseIf Variables_Globales.FechaEmisionRetencionP = "Y" Then
                vPay.DocDate = ObjetoRetencion.FechaEmision
                vPay.TaxDate = ObjetoRetencion.FechaEmision
            Else
                vPay.DocDate = Date.Now
                vPay.TaxDate = ObjetoRetencion.FechaEmision
            End If


            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            'vPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.Reference1 = ""
            '' vPay.Series = 0
            'vPay.TaxDate = DateSerial(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now            
            ' vPay.TaxDate = Date.Now


            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In ObjetoRetencion.ENTDetalleRetencion

                If oDetalle.ValorRetenido > 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle.Codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
                            oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    ElseIf oDetalle.Codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    ElseIf oDetalle.Codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    GuardaLog("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString())

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Dim TypeCode As Integer = CrTypeCode
                    GuardaLog("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString())
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        GuardaLog("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString())
                    End If

                    vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        GuardaLog("CrTypeCode Asignado RENTA - QUERY: " + CrTypeCode.ToString())
                    End Try

                    vPay.CreditCards.CreditSum = oDetalle.ValorRetenido ' _oDocumento.TotalRetencion ' formatDecimal(_oDocumento.TotalRetencion.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    vPay.CreditCards.FirstPaymentSum = ObjetoRetencion.TotalRetencion
                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion
                        'fechaVencRtMP = _oDocumento.FechaAutorizacion
                        'vPay.CreditCards.CardValidUntil = fechaVencRtMP

                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
                                'Left(odt.GetValue(0, i).ToString(), 99))

                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Length)
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If

                        Catch ex As Exception
                        End Try


                        If checkCampoBD("RCT3", "MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = ObjetoRetencion.Secuencial
                        End If
                        If checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                        End If
                        If checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = sCardCode
                        End If
                        'If oFuncionesB1.checkCampoBD("RCT3", "REPL_NUM_RETE") Then
                        '    vPay.CreditCards.UserFields.Fields.Item("U_REPL_NUM_RETE").Value = _oDocumento.Secuencial
                        'End If
                        If String.IsNullOrEmpty(Variables_Globales.CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = ObjetoRetencion.Secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Variables_Globales.CampoNumRetencion).Value = ObjetoRetencion.Secuencial
                        End If


                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision + ObjetoRetencion.Secuencial
                        If checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = ObjetoRetencion.FechaAutorizacion
                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion
                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If
                        Catch ex As Exception
                        End Try
                        If checkCampoBD("ORCT", "HBT_TIP_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_TIP_RET").Value = "E"
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_SERIE_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_SERIE_RET").Value = ObjetoRetencion.Establecimiento.ToString + "-" + ObjetoRetencion.PuntoEmision.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_RET").Value = ObjetoRetencion.Secuencial.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_FEC_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_FEC_RET").Value = CDate(ObjetoRetencion.FechaAutorizacion)
                        End If
                        If checkCampoBD("ORCT", "HBT_CADRET") Then
                            vPay.UserFields.Fields.Item("U_HBT_CADRET").Value = CDate(ObjetoRetencion.FechaAutorizacion)
                        End If
                        If checkCampoBD("RCT3", "HBT_Depositado") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_HBT_Depositado").Value = "SI"
                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                        vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion

                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If

                        Catch ex As Exception
                        End Try


                        If checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = ObjetoRetencion.Secuencial
                        End If
                        If checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                        End If
                        If checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = sCardCode
                        End If
                        If checkCampoBD("RCT3", "SS_NombreSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_NombreSN").Value = Left(ObjetoRetencion.RazonSocial.ToString, 100)
                        End If
                        If String.IsNullOrEmpty(Variables_Globales.CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = ObjetoRetencion.Secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Variables_Globales.CampoNumRetencion).Value = ObjetoRetencion.Secuencial
                        End If

                    End If


                    If checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    If checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                    End If

                    vPay.CreditCards.Add()
                    vPay.CreditCards.SetCurrentLine(secuencial)
                    secuencial += 1


                End If


            Next
            'dibeal
            Try
                If checkCampoBD("ORCT", "DIB_TipoOperacion") Then
                    vPay.UserFields.Fields.Item("U_DIB_TipoOperacion").Value = "A-0015"
                End If
            Catch ex As Exception
                GuardaLog("DIB_TipoOperacion error: " + ex.Message.ToString)
            End Try


            vPay.ControlAccount = CUENTA

            RetVal = vPay.Add()

            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode + ObjetoRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        GuardaLog("Serializando...")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    GuardaLog("EEROR " + ex.Message)
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso)
                oCompanyGP.GetLastError(ErrCode, ErrMsg)
                GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Ocurrio Error al grabar Pago Recibido tipo proveedor (Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString())
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode.ToString() + ObjetoRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        GuardaLog("Serializando...")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    GuardaLog("EEROR " + ex.Message)
                End Try
                oCompanyGP.GetNewObjectCode(sDocEntryPreliminar)
                GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString())
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar, ObjetoRetencion.ClaveAcceso.ToString)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso.ToString)
            GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Error:" + ex.Message.ToString())
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function

    Private Function CrearPagoRecibido_E_ONormal(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String, ByRef ObjetoRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef oListaFacturaVenta As List(Of Entidades.FacturaVenta)) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim queryDocDate As String = ""
        Dim numFact As String = ""
        Dim vPay As SAPbobsCOM.Payments

        GuardaLog("Entrando a funcion CrearPagoRecibido_E_ONormal")

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCB As Integer = 1
        Dim secuencialCBMP As Integer = 1

        Dim diasValidar As Integer = ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "DiasValidarProcesoLote")
        Dim FecEmiFac As Date = ObjetoRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener
        Dim FecAutRet As Date = ObjetoRetencion.FechaAutorizacion
        Dim FecEmiRet As Date = ObjetoRetencion.FechaEmision
        Dim _numDocRetener As String = ObjetoRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(0, 3).ToString + "-" + ObjetoRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(3, 3).ToString + "-" + CLng(Right(ObjetoRetencion.ENTDetalleRetencion(0).NumDocRetener, 9)).ToString
        Dim est As String = ObjetoRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(0, 3).ToString
        Dim punEmi As String = ObjetoRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(3, 3).ToString
        Dim folio As String = CLng(Right(ObjetoRetencion.ENTDetalleRetencion(0).NumDocRetener, 9)).ToString

        GuardaLog("Entrando a validar query docdate")

        If Variables_Globales.PROVEEDOR_DE_SAPBO.EXXIS = Variables_Globales.Nombre_Proveedor Then
            If CInt(ConfigurationManager.AppSettings("DevServerType")) = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryDocDate = " SELECT ""DocDate"" from  OINV where ""DocStatus""='O' AND ""CANCELED""='N' AND ""CardCode""='" + sCardCode.ToString + "' and ""U_SER_EST""='" + est + "' and ""U_SER_PE""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryDocDate = " SELECT DocDate from oinv where ""DocStatus""='O' AND ""CANCELED""='N' AND CardCode='" + sCardCode + "' and (U_SER_EST +'-'+ U_SER_PE +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        ElseIf Variables_Globales.PROVEEDOR_DE_SAPBO.SOLSAP = Variables_Globales.Nombre_Proveedor Then
            If CInt(ConfigurationManager.AppSettings("DevServerType")) = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryDocDate = " SELECT ""DocDate"" from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SS_Est""='" + est + "' and ""U_SS_Pemi""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryDocDate = " SELECT DocDate from oinv where CardCode='" + sCardCode + "' and (U_SS_Est +'-'+ U_SS_Pemi +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        End If
        GuardaLog("Query fecha contabilizacion factura " + queryDocDate.ToString)
        Dim DocDate = CDate(getRSvalue(queryDocDate, "DocDate", ""))
        GuardaLog("DocDate " + DocDate.ToString)
        Dim SfechaConFac = CInt(DocDate.ToString("yyyyMMdd"))
        Dim UltDiaFecha As Date = DateSerial(Year(DocDate), Month(DocDate) + 1, 0)
        GuardaLog("UltDiaFecha " + UltDiaFecha.ToString)

        Try

            GuardaLog("Creando Pago Recibido Normaal(Retencion) ruc: " + ObjetoRetencion.Ruc + " con clave: " + ObjetoRetencion.ClaveAcceso)
            GuardaLog("Instanciando inicio de transaccion")

            If oCompanyGP.InTransaction Then
                GuardaLog("Transaccion en curso, se procede hacer rollback")
                oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)

            End If

            oCompanyGP.StartTransaction()

            vPay = oCompanyGP.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            vPay.CardCode = sCardCode
            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.DocCurrency = "USD"

            If Variables_Globales._ValidarFechasCTK = "Y" Then

                If ValidoDiasLab = True Then
                    'vPay.DocDate = ObjetoRetencion.FechaEmision
                    'vPay.DueDate = ObjetoRetencion.FechaEmision
                    vPay.DocDate = UltDiaFecha
                    vPay.DueDate = UltDiaFecha

                ElseIf FecAutRet.Month <= Date.Now.Month Then

                    If FecEmiRet.Month = Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month = Date.Now.Month Then
                        vPay.DocDate = Date.Now
                        vPay.DueDate = Date.Now
                        'fechaVencRtMP = Date.Now
                    ElseIf FecEmiRet.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                        vPay.DocDate = UltDiaFecha 'Date.Now
                        vPay.DueDate = UltDiaFecha
                        ' fechaVencRtMP = UltDiaFecha

                    ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Month < Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                        vPay.DocDate = UltDiaFecha
                        vPay.DueDate = UltDiaFecha
                        'fechaVencRtMP = UltDiaFecha
                    ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Month = Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                        vPay.DocDate = UltDiaFecha
                        vPay.DueDate = UltDiaFecha
                        'fechaVencRtMP = UltDiaFecha
                    ElseIf Date.Now.Day > diasValidar And FecEmiRet.Month = Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                        'vPay.DocDate = Date.Now
                        vPay.DocDate = UltDiaFecha
                        vPay.DueDate = UltDiaFecha
                        'fechaVencRtMP = Date.Now
                    ElseIf Date.Now.Day > diasValidar And FecEmiRet.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                        vPay.DocDate = Date.Now
                        vPay.DueDate = Date.Now
                        'fechaVencRtMP = Date.Now
                    ElseIf FecEmiRet.Year < Date.Now.Year Then
                        If Date.Now.Day <= diasValidar And FecEmiRet.Year < Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                            vPay.DocDate = UltDiaFecha
                            vPay.DueDate = UltDiaFecha
                            'fechaVencRtMP = UltDiaFecha
                        ElseIf FecEmiRet.Month < Date.Now.Year And FecAutRet.Year < Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                            vPay.DocDate = UltDiaFecha 'Date.Now
                            vPay.DueDate = UltDiaFecha
                            'fechaVencRtMP = UltDiaFecha
                        ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Year = Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                            vPay.DocDate = UltDiaFecha
                            vPay.DueDate = UltDiaFecha
                            'fechaVencRtMP = UltDiaFecha
                        Else
                            vPay.DocDate = Date.Now
                            vPay.DueDate = Date.Now
                            'fechaVencRtMP = Date.Now
                        End If
                    Else
                        vPay.DocDate = Date.Now
                        vPay.DueDate = Date.Now
                        'fechaVencRtMP = Date.Now
                    End If
                ElseIf FecEmiRet.Year < Date.Now.Year Then
                    If Date.Now.Day <= diasValidar And FecEmiRet.Year < Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                        vPay.DocDate = UltDiaFecha
                        vPay.DueDate = UltDiaFecha
                        'fechaVencRtMP = UltDiaFecha
                    ElseIf FecEmiRet.Year < Date.Now.Year And FecAutRet.Year < Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                        vPay.DocDate = UltDiaFecha 'Date.Now
                        vPay.DueDate = UltDiaFecha
                        ' fechaVencRtMP = UltDiaFecha
                    ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Year = Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                        vPay.DocDate = UltDiaFecha
                        vPay.DueDate = UltDiaFecha
                        'fechaVencRtMP = UltDiaFecha
                    ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Year < Date.Now.Year And FecAutRet.Year < Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                        vPay.DocDate = UltDiaFecha
                        vPay.DueDate = UltDiaFecha
                        'fechaVencRtMP = UltDiaFecha
                    Else
                        vPay.DocDate = Date.Now
                        vPay.DueDate = Date.Now
                    End If
                Else
                    vPay.DocDate = Date.Now
                    vPay.DueDate = Date.Now
                    'fechaVencRtMP = Date.Now
                End If

                'ElseIf Variables_Globales._vgFechaEmisionRetencion = "Y" Then
                '    vPay.DocDate = oRetencion.FechaEmision
                '    vPay.DueDate = oRetencion.FechaEmision
                '    vPay.TaxDate = oRetencion.FechaEmision
                'ElseIf Variables_Globales._vgFechaEmisionRetencionP = "Y" Then
                '    vPay.DocDate = oRetencion.FechaEmision
                '    vPay.TaxDate = oRetencion.FechaEmision

            Else
                vPay.DocDate = oRetencion.FechaEmision
                vPay.DueDate = oRetencion.FechaEmision
            End If

            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            vPay.Reference1 = ""
            vPay.Series = CInt(Variables_Globales.IdSeriePR)
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""
            Dim valorRetencion0 As Boolean = False
            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""
            Dim conRet0 As Integer = 0
            Dim conRetM0 As Integer = 0
            Dim secuencial As Integer = 1
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In ObjetoRetencion.ENTDetalleRetencion

                If ObjetoRetencion.TotalRetencion > 0 And oDetalle.ValorRetenido > 0 Then
                    numFact = oDetalle.NumDocRetener.ToString
                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    Dim qryDocDateFact = ""

                    If Variables_Globales.PROVEEDOR_DE_SAPBO.EXXIS = Variables_Globales.Nombre_Proveedor Then
                        If CInt(ConfigurationManager.AppSettings("DevServerType")) = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            qryDocDateFact = " SELECT ""DocDate"" from  OINV where ""DocStatus""='O' AND ""CANCELED""='N' AND ""CardCode""='" + sCardCode.ToString + "' and ""U_SER_EST""='" + numFact.ToString().Substring(0, 3) + "' and ""U_SER_PE""='" + numFact.ToString().Substring(3, 3) + "' and cast(""FolioNum"" as varchar)='" + CLng(Right(numFact, 9)).ToString + "'"
                        Else
                            qryDocDateFact = " SELECT DocDate from oinv where ""DocStatus""='O' AND ""CANCELED""='N' AND CardCode='" + sCardCode + "' and ""U_SER_EST""='" + numFact.ToString().Substring(0, 3) + "' and ""U_SER_PE""='" + numFact.ToString().Substring(3, 3) + "' and cast(""FolioNum"" as varchar)='" + CLng(Right(numFact, 9)).ToString + "'"
                        End If
                    ElseIf Variables_Globales.PROVEEDOR_DE_SAPBO.SOLSAP = Variables_Globales.Nombre_Proveedor Then
                        If CInt(ConfigurationManager.AppSettings("DevServerType")) = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            qryDocDateFact = " SELECT ""DocDate"" from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SS_Est""='" + numFact.ToString().Substring(0, 3) + "' and ""U_SS_Pemi""='" + numFact.ToString().Substring(3, 3) + "' and cast(""FolioNum"" as varchar)='" + CLng(Right(numFact, 9)).ToString + "'"
                        Else
                            qryDocDateFact = " SELECT DocDate from oinv where CardCode='" + sCardCode + "' and ""U_SS_Est""='" + numFact.ToString().Substring(0, 3) + "' and ""U_SS_Pemi""='" + numFact.ToString().Substring(3, 3) + "' and cast(""FolioNum"" as varchar)='" + CLng(Right(numFact, 9)).ToString + "'"
                        End If
                    End If

                    GuardaLog("Consultando Fecha secuencial " + oDetalle.NumDocRetener.ToString)
                    Dim FechaCont = CDate(getRSvalue(qryDocDateFact, "DocDate", ""))
                    GuardaLog("Fecha secuencial " + oDetalle.NumDocRetener.ToString + " consultada: " + DocDate.ToString)

                    Dim _SfechaConFac = CInt(FechaCont.ToString("yyyyMMdd"))
                    Dim _SfechaEmiDocSus = CInt(oDetalle.FechaEmisionDocRetener.ToString("yyyyMMdd"))

                    If _SfechaConFac <> _SfechaEmiDocSus Then

                        GuardaLog("La fecha de emision del documento sustento: " + oDetalle.FechaEmisionDocRetener.ToString("yyyy-MM-dd") + " es diferente a la fecha de contabilizacion de la factura: " + DocDate.ToString())
                        GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                "Motivo: La fecha de emision del documento sustento: " + oDetalle.FechaEmisionDocRetener.ToString("yyyy-MM-dd") + " es diferente a la fecha de contabilizacion de la factura: " + FechaCont.ToString("yyyy-MM-dd"))
                        oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)

                        Return False

                    End If

                    If oDetalle.Codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
                            GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
                            oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
                            'oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    ElseIf oDetalle.Codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
                            'oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    ElseIf oDetalle.Codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        GuardaLog("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString())
                        If CodRetencion = "" Then
                            GuardaLog("No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
                            'oFuncionesB1GP.Release(vPay)
                            Return False
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    GuardaLog("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString())

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Dim TypeCode As Integer = CrTypeCode
                    GuardaLog("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString())
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        GuardaLog("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString())
                    End If

                    vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        GuardaLog("CrTypeCode Asignado RENTA - QUERY: " + CrTypeCode.ToString())
                    End Try




                    vPay.CreditCards.CreditSum = oDetalle.ValorRetenido ' _oDocumento.TotalRetencion ' formatDecimal(_oDocumento.TotalRetencion.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    vPay.CreditCards.FirstPaymentSum = ObjetoRetencion.TotalRetencion
                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    If Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion

                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then

                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If

                        Catch ex As Exception
                        End Try


                        If checkCampoBD("RCT3", "MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = ObjetoRetencion.Secuencial
                        End If
                        If checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                        End If
                        If checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = sCardCode
                        End If
                        'If oFuncionesB1.checkCampoBD("RCT3", "REPL_NUM_RETE") Then
                        '    vPay.CreditCards.UserFields.Fields.Item("U_REPL_NUM_RETE").Value = _oDocumento.Secuencial
                        'End If
                        If String.IsNullOrEmpty(Variables_Globales.CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = ObjetoRetencion.Secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Variables_Globales.CampoNumRetencion).Value = ObjetoRetencion.Secuencial
                        End If


                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision + ObjetoRetencion.Secuencial
                        If checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = ObjetoRetencion.FechaAutorizacion
                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion
                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If
                        Catch ex As Exception
                        End Try
                        If checkCampoBD("ORCT", "HBT_TIP_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_TIP_RET").Value = "E"
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_SERIE_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_SERIE_RET").Value = ObjetoRetencion.Establecimiento.ToString + "-" + ObjetoRetencion.PuntoEmision.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_RET").Value = ObjetoRetencion.Secuencial.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = ObjetoRetencion.AutorizacionSRI.ToString
                        End If
                        If checkCampoBD("ORCT", "HBT_FEC_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_FEC_RET").Value = CDate(ObjetoRetencion.FechaAutorizacion)
                        End If
                        If checkCampoBD("ORCT", "HBT_CADRET") Then
                            vPay.UserFields.Fields.Item("U_HBT_CADRET").Value = CDate(ObjetoRetencion.FechaAutorizacion)
                        End If
                        If checkCampoBD("RCT3", "HBT_Depositado") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_HBT_Depositado").Value = "SI"
                        End If

                    ElseIf Variables_Globales.Nombre_Proveedor = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                        vPay.CreditCards.FirstPaymentDue = ObjetoRetencion.FechaAutorizacion

                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                            End If

                        Catch ex As Exception
                        End Try


                        'If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                        '    vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        'End If
                        If checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                        End If
                        If checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = ObjetoRetencion.Secuencial
                        End If
                        If checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = ObjetoRetencion.AutorizacionSRI
                        End If
                        If checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = ObjetoRetencion.Establecimiento + ObjetoRetencion.PuntoEmision
                        End If
                        If checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = sCardCode
                        End If

                        If checkCampoBD("RCT3", "SS_NombreSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_NombreSN").Value = Left(ObjetoRetencion.RazonSocial.ToString, 100)
                        End If

                        If String.IsNullOrEmpty(Variables_Globales.CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = ObjetoRetencion.Secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Variables_Globales.CampoNumRetencion).Value = ObjetoRetencion.Secuencial
                        End If

                    End If


                    If checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    If checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                    End If

                    vPay.CreditCards.Add()
                    vPay.CreditCards.SetCurrentLine(secuencial)
                    secuencial += 1
                    conRetM0 += 1
                    'valorRetencion0 = True
                Else
                    conRet0 += 1
                End If


            Next
            'dibeal
            If conRet0 > 0 And conRetM0 = 0 Then

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso)
                GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: existe un codigo de renta o iva 0%")
                oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
                Return False

            End If

            Try
                If checkCampoBD("ORCT", "DIB_TipoOperacion") Then
                    vPay.UserFields.Fields.Item("U_DIB_TipoOperacion").Value = "A-0015"
                End If
            Catch ex As Exception
                GuardaLog("DIB_TipoOperacion error: " + ex.Message.ToString)
            End Try

            Try
                If checkCampoBD("ORCT", "SS_SERPL") Then
                    vPay.UserFields.Fields.Item("U_SS_SERPL").Value = "SI"
                End If
                If checkCampoBD("ORCT", "SS_SERPLF") Then
                    vPay.UserFields.Fields.Item("U_SS_SERPLF").Value = Date.Now.ToString + " " + Thread.CurrentThread.Name.ToString
                End If
            Catch ex As Exception
                GuardaLog("DIB_TipoOperacion error: " + ex.Message.ToString)
            End Try

            ' FACTURAS
            If ExisteFacRel Then
                For Each o As Entidades.FacturaVenta In oListaFacturaVenta
                    GuardaLog("Datos Docentry: " & o.DocEntry.ToString & " valor a retener: " & o.ValorARetener.ToString)
                    vPay.Invoices.DocEntry = o.DocEntry
                    vPay.Invoices.SumApplied = o.ValorARetener
                    vPay.Invoices.InstallmentId = o.Cuota
                    vPay.Invoices.Add()
                Next
            End If

            If SaldoPendiente < ObjetoRetencion.TotalRetencion And Variables_Globales.ContSaldoPendMenor = "SI" Then

                Dim qryCuentaControl = " select ""AcctCode"" from OACT WHERE ""FormatCode""='" + Variables_Globales.CuentaSaldoFavor + "'"
                GuardaLog("qryCuentaControl: " & qryCuentaControl.ToString)
                Dim resultqry = getRSvalue(qryCuentaControl.ToString, "AcctCode", "")
                GuardaLog("resultado qryCuentaControl: " & resultqry.ToString)
                If resultqry <> "" Then
                    vPay.ControlAccount = resultqry
                    vPay.JournalRemarks = "SALDO A FAVOR #RET " + ObjetoRetencion.Secuencial.ToString() + " #FAC " + numFact.Substring(6, 9)
                    vPay.Remarks = "SALDO A FAVOR #RET " + ObjetoRetencion.Establecimiento.ToString() + "-" + ObjetoRetencion.PuntoEmision.ToString() + "-" + ObjetoRetencion.Secuencial.ToString() + " #FAC " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9)
                Else

                    GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: No se encontro codigo de la cuenta a favor " + Variables_Globales.CuentaSaldoFavor.ToString)
                    oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
                    Return False

                End If

            End If

            RetVal = vPay.Add()

            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode + ObjetoRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        GuardaLog("Serializando...")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    GuardaLog("EEROR " + ex.Message)
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso)
                oCompanyGP.GetLastError(ErrCode, ErrMsg)
                GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString())
                GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: Error al crear pago: " + ErrCode.ToString + " - " + ErrMsg.ToString())
                'oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode.ToString() + ObjetoRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        GuardaLog("Serializando...")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    GuardaLog("EEROR " + ex.Message)
                End Try
                oCompanyGP.GetNewObjectCode(sDocEntryPreliminar)
                Dim qryNumPago = "Select ""DocNum"" from ORCT where ""DocEntry""=" + sDocEntryPreliminar.ToString
                GuardaLog("Query numero de pago: " + qryNumPago.ToString)
                NumDocPago = getRSvalue(qryNumPago, "DocNum", "0")
                GuardaLog("PRR " + ObjetoRetencion.ClaveAcceso + " Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString())
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar, ObjetoRetencion.ClaveAcceso)
                If oCompanyGP.InTransaction Then
                    oCompanyGP.EndTransaction(BoWfTransOpt.wf_Commit)
                End If
                GuardaLog("Finalizo la transaccion de la retencion " + ObjetoRetencion.ClaveAcceso)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO, ObjetoRetencion.ClaveAcceso)
            GuardaLog("PRR" + ObjetoRetencion.ClaveAcceso + " Error: " + ex.Message.ToString())
            GuardaLogNoContabilizado("Razon Social: " + ObjetoRetencion.RazonSocial.ToString & vbCrLf & "Numero de Retencion: " + ObjetoRetencion.Establecimiento + "-" + ObjetoRetencion.PuntoEmision + "-" + ObjetoRetencion.Secuencial & vbCrLf &
                                    "Numero de factura: " + numFact.ToString().Substring(0, 3) + "-" + numFact.ToString().Substring(3, 3) + "-" + numFact.Substring(6, 9) & vbCrLf &
                                    "Clave de acceso: " + ObjetoRetencion.ClaveAcceso & vbCrLf &
                                    "Motivo: Error en funcun crear pago normal: " + ex.Message.ToString())
            If oCompanyGP.InTransaction Then
                oCompanyGP.EndTransaction(BoWfTransOpt.wf_RollBack)
            End If

            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function


    Public Sub Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String, ByRef ClaveAcceso As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            Dim query As String
            Dim CodeExist As String = "0"
            If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & oCompanyGP.CompanyDB & """.""@GS_RER"" Where ""DocEntry"" = '" + DocEntryRERecibida_UDO + "' "
            Else
                query = "Select DocEntry From [@GS_RER] Where DocEntry = '" + DocEntryRERecibida_UDO + "' "
            End If
            CodeExist = getRSvalue(query, "DocEntry")

            GuardaLog("PRR Eliminando Documento Recibido UDO Retención # " + DocEntryRERecibida_UDO.ToString())

            If CodeExist = "0" Or CodeExist = "" Then ' 
                GuardaLog("PRR Actualiza_DocumentoRecibido_RE " + ClaveAcceso + " No existen registros coincidentes, la consulta no trajo registro: " + query.ToString())
            Else
                oCompanyService = oCompanyGP.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_RER")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", DocEntryRERecibida_UDO)

                'oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                'oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

                oGeneralService.Delete(oGeneralParams)
                GuardaLog("PRR Eliminado" + DocEntryRERecibida_UDO.ToString())
            End If




            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            GuardaLog("PRR " + ClaveAcceso + " Error: Eliminando Documento Recibido UDO Retención..: " + ex.Message().ToString())
        End Try
    End Sub

    Public Sub Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String, DocEntryPreliminar As String, ByRef ClaveAcceso As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            GuardaLog("PRR Actualiza_DocumentoRecibido_RE" + ClaveAcceso + "Actualizando Numero de Documento Preliminar en Documento Recibido UDO")

            Dim query As String
            Dim CodeExist As String = "0"
            If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & oCompanyGP.CompanyDB & """.""@GS_RER"" Where ""DocEntry"" = '" + DocEntryRERecibida_UDO + "' "
            Else
                query = "Select DocEntry From [@GS_RER] Where DocEntry = '" + DocEntryRERecibida_UDO + "' "
            End If
            CodeExist = getRSvalue(query, "DocEntry")

            If CodeExist = "0" Or CodeExist = "" Then ' 
                GuardaLog("PRR Actualiza_DocumentoRecibido_RE " + ClaveAcceso + " No existen registros coincidentes, la consulta no trajo registro: " + query.ToString())
            Else
                oCompanyService = oCompanyGP.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_RER")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

                oGeneralService.Update(oGeneralData)

            End If


        Catch ex As Exception
            GuardaLog("PRR Actualiza_DocumentoRecibido_RE " + ClaveAcceso + " Error: Actualizando Numero de Documento Preliminar en Documento Recibido UDO: " + ex.Message().ToString())
        End Try
    End Sub

    Private Function ConvertToDouble(s As String) As Double
        Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
        Dim result As Double = 0
        Try
            If s IsNot Nothing Then
                If Not s.Contains(",") Then
                    result = Double.Parse(s, CultureInfo.InvariantCulture)
                Else
                    result = Convert.ToDouble(s.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
                End If
            End If
        Catch e As Exception
            Try
                result = Convert.ToDouble(s)
            Catch
                Try
                    result = Convert.ToDouble(s.Replace(",", ";").Replace(".", ",").Replace(";", "."))
                Catch
                    Throw New Exception("Wrong string-to-double format")
                End Try
            End Try
        End Try
        Return result
    End Function

    Public Function ConsultaParametro(ByVal Modulo As String, ByVal Tipo As String, ByVal Subtipo As String, ByVal Nombre As String) As String
        Try
            Dim valor As String = ""
            Dim sQueryPrefijo As String = ""
            If oCompanyGP.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
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

    Public Function getRSvalue(ByVal query As String, ByVal columnaRet As String, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
            GuardaLog("getRSvalue Catch:" + ex.Message().ToString() + "-QUERY: " + query)
        End Try
        Return ret
    End Function

    Public Function getRecordSet(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = oCompanyGP.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
            GuardaLog("nzString Catch:" + ex.Message().ToString())
        End Try
        Return valorSiNulo
    End Function

    Public Sub Release(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            GuardaLog("Release Catch:" + ex.Message().ToString())
        End Try

    End Sub

    Public Function checkCampoBD(ByVal Tabla As String, ByVal Campo As String) As Boolean
        Dim retorno As Boolean = False
        Dim strSQLBD As String = ""
        Dim oLocalBD As SAPbobsCOM.Recordset = Nothing

        Try
            oLocalBD = oCompanyGP.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oCompanyGP.DbServerType = "9" Then
                strSQLBD = " SELECT ""TableID""  FROM """ & oCompanyGP.CompanyDB & """.""CUFD""  WHERE ""TableID"" ='" & Tabla & "' AND ""AliasID"" = '" & Campo & "'"
            Else
                strSQLBD = "SELECT TableID  FROM CUFD  WHERE TableID ='" & Tabla & "' AND AliasID = '" & Campo & "'"
            End If
            'strSQLBD = "SELECT column_name "
            'strSQLBD &= "FROM [" & fCompany.CompanyDB & "].INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = '" & Campo & "' AND Table_Name ='" & Tabla & "'"
            oLocalBD.DoQuery(strSQLBD)
            If oLocalBD.EoF = False Then
                retorno = True
            End If
            Release(oLocalBD)
        Catch ex As Exception
            GuardaLog("FUN_CreaCampos_checkCampoBD_Catch, Query: " + strSQLBD)
            GuardaLog("FUN_CreaCampos_checkCampoBD_Catch, Nombre Tabla: " + Tabla + "- Nombre Campo: " + Campo + "-Error: " + ex.Message.ToString())
        End Try
        Return retorno
    End Function

    Public Function ActualizadoEstado_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Estado As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            GuardaLog("PRR " + sCardCode + " Actualizando el estado a : " + Estado.ToString() + " al codigo: " + DocEntryFacturaRecibida_UDO.ToString)
            oCompanyService = oCompanyGP.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", Estado)
            oGeneralData.SetProperty("U_FechaS", Integer.Parse(Date.Now.ToString("yyyyMMdd")))
            oGeneralData.SetProperty("U_Sincro", Sincronizado)

            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            GuardaLog("PRR " + DocEntryFacturaRecibida_UDO + " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString())
            Return False
        End Try
    End Function

    Public Function MarcarVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String) As Boolean
        Try
            '_WS_Recepcion = Variables_Globales.WS_Recepcion
            'If _WS_Recepcion = "" Then
            '    rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            'End If
            '_WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            '_WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
            WS.Url = Variables_Globales.WS_RecepcionCambiarEstado
            ' MANEJO PROXY
            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = IIf(Variables_Globales.SALIDA_POR_PROXY = Nothing, "N", Variables_Globales.SALIDA_POR_PROXY)
            GuardaLog("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString)
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            If SALIDA_POR_PROXY = "Y" Then
                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                GuardaLog("Proxy_puerto : " + Proxy_puerto.ToString)
                GuardaLog("Proxy_IP : " + Proxy_IP.ToString)
                GuardaLog("Proxy_Usuario : " + Proxy_Usuario.ToString)
                GuardaLog("Proxy_Clave : " + Proxy_Clave.ToString)

                If Not Proxy_puerto = "" Then
                    proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                Else
                    proxyobject = New System.Net.WebProxy(Proxy_IP)
                End If
                cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                proxyobject.Credentials = cred

                WS.Proxy = proxyobject
                WS.Credentials = cred
            End If
            ' END  MANEJO PROXY

            SetProtocolosdeSeguridad()
            If WS.MarcarVisto(Variables_Globales.WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            GuardaLog("Error MarcarVisto : " + ex.Message.ToString)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            GuardaLog("PRR" + DocEntryFacturaRecibida_UDO + " Actualizando a Sincronizado EDOC = " + Sincronizado.ToString())
            oCompanyService = oCompanyGP.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_SincroE", Sincronizado)

            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            GuardaLog("Error ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE : " + ex.Message.ToString)

            Return False
        End Try

    End Function

    Public Function Calculo5toDiaLaborable(ByRef FechaEmisionFac As Date) As Date?

        Dim FechaRetorno As Date?
        Try

            FechaRetorno = Nothing
            Dim today As Date = Date.Today
            ' Obtener el primer día del mes actual
            Dim firstDayOfCurrentMonth As Date = DateSerial(Year(today), Month(today), 1)
            GuardaLog("Dia Actual: " + firstDayOfCurrentMonth.ToString("yyyy-MM-dd"))
            ' Obtener el último día del mes anterior
            Dim lastDayOfPreviousMonth As Date = firstDayOfCurrentMonth.AddDays(-1)
            GuardaLog("Ultimo dia del mes anterior: " + lastDayOfPreviousMonth.ToString("yyyy-MM-dd"))
            ' Obtener el penúltimo día del mes anterior, la variable va a depender 
            Dim secondToLastDayOfPreviousMonth As Date = lastDayOfPreviousMonth.AddDays(-(CInt(Variables_Globales.CantUltmsDias) - 1))
            GuardaLog("Fecha ultimos dias del mes: " + secondToLastDayOfPreviousMonth.ToString("yyyy-MM-dd"))

            If FechaEmisionFac >= secondToLastDayOfPreviousMonth AndAlso FechaEmisionFac <= lastDayOfPreviousMonth Then
                ' Lista para almacenar los días laborales
                Dim businessDays As New List(Of Date)

                ' Contador de días laborales
                'Dim currentDay As Date = secondToLastDayOfPreviousMonth
                Dim currentDay As Date = FechaEmisionFac 'se cambia a la fecha de emision de la factura porque desde esta fecha debe de correr los dias hbailes

                ' Mientras no tengamos 5 días laborales
                While businessDays.Count < CInt(Variables_Globales.CantDiasLab)
                    ' Si es un día laboral (lunes a viernes)
                    If currentDay.DayOfWeek >= DayOfWeek.Monday AndAlso currentDay.DayOfWeek <= DayOfWeek.Friday Then
                        GuardaLog("Dia: " + currentDay.ToString("yyyy-MM-dd"))
                        businessDays.Add(currentDay)
                    End If
                    ' Avanzar al siguiente día
                    currentDay = currentDay.AddDays(1)
                End While

                'fecha quinto dia laborable
                FechaRetorno = businessDays(Variables_Globales.CantDiasLab - 1)
            End If

            Return FechaRetorno

        Catch ex As Exception
            GuardaLog("Error en funcion Calculo5toDiaLaborable: " + ex.Message.ToString)
            Return FechaRetorno
        End Try
    End Function
End Class

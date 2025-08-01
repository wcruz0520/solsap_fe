Imports System.Data.SqlClient
Imports Functions
Imports System.IO
Imports System.Text
Imports System.Xml.Serialization

'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Imports SAPbouiCOM
Imports SAPbobsCOM
Imports Spire.Pdf
Imports System.Drawing.Printing
Imports System.Globalization

Imports itextsharp.text.pdf
Imports itextsharp.text


Public Class ManejoDeDocumentosEcua

    Private rCompanyEcua As SAPbobsCOM.Company
    Private rsboAppEcua As SAPbouiCOM.Application
    Private _EstadoAutorizacionEcua As String = ""
    Private _ClaveAccesoEcua As String = ""
    Private _statusMsg As String = ""
    Private _ObservacionEcua As String = ""
    Private _NumAutorizacionEcua As String = ""
    Private _FechaAutorizacionEcua As String
    Private _EstadoSAPEcua As String = ""
    Private _ErrorEcua As String = ""
    Dim mensajeEcua As String = ""
    Dim oObjetoEcua As Object
    'Dim ObjetoRespuesta As Object = Nothing
    Dim oFuncionesAddonEcua As FuncionesAddon

    Dim oFuncionesB1Ecua As FuncionesB1

    Dim _tipoManejoEcua As String
    Dim _errorMensajeWSEnvíoEcua As String
    Public _Nombre_Proveedor_SAP_BOEcua As String = ""

    Dim _mensajeSRI As String

    Dim proxyobjectEcua As System.Net.WebProxy
    Dim credEcua As System.Net.NetworkCredential

    Dim oDocumentoEcua As SAPbobsCOM.Documents

    Dim _GuardarLogEcua As String = "N"

    Private _NumeroDeDocumentoSRIEcua As String = ""

    Dim mensajeDocAutEcua As String = ""

    Public CONEXION As Odbc.OdbcConnection

    Private CoreRest As CoreRest

    Private accessKey As String = ""
    Private authorizationDate As String = ""
    Private authorizationNumber As String = ""
    Private state As String = ""
    Private statusMsg As String = ""

    Private Base64 As String = ""
    Private NombreXML As String = ""
    Private _estado As String = ""

    Private ResulCode As String = ""
    Private ResultMensaje As String = ""

    Private xml As String = ""
    Private pdf As String = ""

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application, tipoManejo As String, ByVal ProveedorSAPBO As String)
        'Utilitario.Util_Log.Escribir_Log("SubNew Inicio", "ManejoDeDocumentos")
        rCompanyEcua = Company
        _tipoManejoEcua = tipoManejo
        _Nombre_Proveedor_SAP_BOEcua = ProveedorSAPBO
        If tipoManejo = "A" Then
            rsboAppEcua = sboApp
            oFuncionesAddonEcua = New Functions.FuncionesAddon(rCompanyEcua, rsboAppEcua, True, False)
            oFuncionesB1Ecua = New Functions.FuncionesB1(rCompanyEcua, rsboAppEcua, True, False)
        Else
            ' SI ES SERVICIO INSTANCIO ESTA CLASE, YA QUE NO USA LA UIAPI
            oFuncionesAddonEcua = New Functions.FuncionesAddon(rCompanyEcua, rsboAppEcua, True, False)
        End If

        CoreRest = New CoreRest()

        If tipoManejo = "A" Then
            CoreRest.WS_EnvioDocumento = Functions.VariablesGlobales._WsEmisionEcua
            CoreRest.WS_ConsultaDocumento = Functions.VariablesGlobales._WsEmisionConsultaEcua
        Else
            CoreRest.WS_EnvioDocumento = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WsEmisionEcu")
            CoreRest.WS_ConsultaDocumento = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WsConsultaEcu")
        End If




    End Sub

    Public Function ProcesaEnvioDocumento(DocEntry As Integer, TipoDocumento As String, Optional ByVal sincronizado As Boolean = False) As String

        Try
            Dim result As Boolean = False
            Dim objetoRespuestaEcu As Object = Nothing
            Dim _objetoRespuestaEcu As Entidades.RespuestaEnvio = Nothing
            Dim _RespuestaConsultaDocEcu As Entidades.ConsultaDocRespuesta = Nothing
            Dim TipoWS As String = "LOCAL"
            Dim BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo
            Dim sSQL As String = ""

            If _tipoManejoEcua = "S" Then
                TipoWS = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
            Else
                TipoWS = Functions.VariablesGlobales._TipoWS
            End If

            Utilitario.Util_Log.Escribir_Log("TIPO WEB SERVICES: " + TipoWS, "ManejoDeDocumentos")
            'Se escribe el log

            If sincronizado = True Then

                'Se valida el parametro sincronizado 
                'en caso de ser verdadero solo se llamara al metodo sincronizar
                'caso contrario se Enviara el documento
                'llamar metodo sincronizador
                _RespuestaConsultaDocEcu = SincronizarDocumento(TipoDocumento, DocEntry, TipoWS)
                'objetoRespuesta = EnviaDocumentoSRI(oObjeto, TipoDocumento, DocEntry, TipoWS)

                If Not _RespuestaConsultaDocEcu Is Nothing Then

                    accessKey = _RespuestaConsultaDocEcu.accessKey

                    authorizationDate = _RespuestaConsultaDocEcu.authorizationDate
                    authorizationNumber = _RespuestaConsultaDocEcu.authorizationNumber
                    state = _RespuestaConsultaDocEcu.state
                    statusMsg = _RespuestaConsultaDocEcu.statusMsg

                    ResulCode = _RespuestaConsultaDocEcu.result.code
                    ResultMensaje = _RespuestaConsultaDocEcu.result.message
                    xml = _RespuestaConsultaDocEcu.xml
                    pdf = _RespuestaConsultaDocEcu.pdf

                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "SS_SINCRO_Respuesta del SRI: " + _EstadoAutorizacionEcua.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                    'VUELVO A VACIAR TODAS LAS VARIABLES YA QUE SE ESTA ALMACENANDO LA AUTORIZACION DEL DOCUMENTO ANTERIOR 
                    _NumAutorizacionEcua = ""
                    _ClaveAccesoEcua = ""
                    _EstadoAutorizacionEcua = ""
                    _ObservacionEcua = ""
                    _FechaAutorizacionEcua = ""

                    If statusMsg = "SIN ENVIAR AL SRI" Then 'documento en procesamiento

                        _ClaveAccesoEcua = accessKey
                        _NumAutorizacionEcua = ""
                        _FechaAutorizacionEcua = ""
                        _ObservacionEcua = "Respuesta Ws- Codigo: " + ResulCode + " - Mensaje Codigo: " + ResultMensaje + " - Mensaje: " + IIf(statusMsg = Nothing, "", statusMsg)
                        _EstadoAutorizacionEcua = "1"

                    ElseIf ResulCode = "600" Then 'documento no encontrado

                        _ClaveAccesoEcua = ""
                        _NumAutorizacionEcua = ""
                        _FechaAutorizacionEcua = ""
                        _ObservacionEcua = "Respuesta Ws- Codigo: " + ResulCode + " - Mensaje Codigo: " + ResultMensaje + " - Mensaje: " + IIf(statusMsg = Nothing, "", statusMsg)
                        _EstadoAutorizacionEcua = "4"

                    ElseIf state = "10" Then 'documento con error 

                        _ClaveAccesoEcua = accessKey
                        _NumAutorizacionEcua = ""
                        _FechaAutorizacionEcua = ""
                        _ObservacionEcua = "Respuesta Ws- Codigo: " + ResulCode + " - Mensaje Codigo: " + ResultMensaje + " - Mensaje: " + IIf(statusMsg = Nothing, "", statusMsg)
                        _EstadoAutorizacionEcua = "6"

                    ElseIf ResulCode = "1000" And ResultMensaje = "OK" And state = "100" And statusMsg = "OK" Then 'documento autorizado

                        _ClaveAccesoEcua = accessKey
                        _NumAutorizacionEcua = authorizationNumber
                        _FechaAutorizacionEcua = authorizationDate
                        _ObservacionEcua = "Respuesta Ws- Codigo: " + ResulCode + " - Mensaje Codigo: " + ResultMensaje + " - Mensaje: " + IIf(statusMsg = Nothing, "", statusMsg)
                        _EstadoAutorizacionEcua = "2"

                    End If

                    If _ObservacionEcua <> "" Then
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#193;", "A")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#201;", "E")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#205;", "I")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#211;", "O")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#218;", "U")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#225;", "a")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#233;", "e")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#237;", "i")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#243;", "o")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#250;", "u")
                        _ObservacionEcua = _ObservacionEcua.Replace("&amp;#39;", "'")

                    End If

                    Dim mensajeError As String = ""
                    If ResulCode = "1000" And ResultMensaje = "OK" And state = "100" And statusMsg = "OK" Then
                        Try
                            _NumAutorizacionEcua = authorizationNumber
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("SS_SINCRO Numero Aut SRI..!!" + _NumAutorizacionEcua.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            End If

                            _ObservacionEcua = "SS_SINCRO: Documento AUTORIZADO AUTORIZACION # " & _NumAutorizacionEcua
                        Catch ex As Exception

                        End Try
                    Else
                        _NumAutorizacionEcua = "0000000000"
                        If _tipoManejoEcua = "A" Then
                            rsboAppEcua.SetStatusBarMessage("SS_SINCRO Numero Aut SRI..!!" + _NumAutorizacionEcua.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        End If

                        mensajeError = _ObservacionEcua.ToString
                    End If



                    If _tipoManejoEcua = "A" Then
                        rsboAppEcua.SetStatusBarMessage("SS_SINCRO_Grabando respuesta de SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "GS_SINCRO_Grabando Respuesta del SRI en Documento - " + TipoDocumento + " - " + DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                    End If

                    '_ObservacionEcua = String.Format("SINCRO Estado:{0} - # AUTORIZACION {1} - Mensaje - {2}", _EstadoAutorizacionEcua.ToString, _NumAutorizacionEcua.ToString, mensajeError)

                    ' Mando a Grabar a SAP
                    If TipoDocumento = "LQE" Then
                        result = GrabaDatosAutorizacion_LiquidacionCompra(DocEntry, TipoDocumento)

                    Else
                        result = GrabaDatosAutorizacion(DocEntry, TipoDocumento)
                    End If
                    If result Then
                        If _tipoManejoEcua = "A" Then
                            rsboAppEcua.SetStatusBarMessage("Proceso terminado con exito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        End If

                    Else
                        If _tipoManejoEcua = "A" Then
                            rsboAppEcua.SetStatusBarMessage("Ocurrio un Error al Guardar los datos de Autorización..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End If

                    End If

                    'AQUI FUNCION IMPRIMIR DOCUMENTO
                    If _tipoManejoEcua = "A" Then
                        If Functions.VariablesGlobales._ImpDocAut = "Y" Then
                            If ResulCode = "1000" And ResultMensaje = "OK" And state = "100" And statusMsg = "OK" Then
                                rsboAppEcua.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + " - Imprimiendo Documento por favor esperar... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                mensajeDocAutEcua = ""
                                If ImprmirDOcAut(pdf) Then 'validar impresion
                                    rsboAppEcua.SetStatusBarMessage("El documento se imprimió con éxito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                Else
                                    rsboAppEcua.SetStatusBarMessage("Error al imprimir PDF: " + mensajeDocAutEcua.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                            End If
                        End If
                    End If


                Else

                    'No se pudo sincronizar
                    _ObservacionEcua = "SS_SINCRO - ProcesaEnvioDocumento - ObjetoRespuesta vacio  " + DocEntry.ToString() + " - " + mensajeEcua.ToString()
                    _ErrorEcua = "SS_SINCRO-No se recibio respuesta inmediata del servicio : " + _errorMensajeWSEnvíoEcua.ToString()
                    If _tipoManejoEcua = "A" Then
                        rsboAppEcua.SetStatusBarMessage("SS_SINCRO-No se recibio respuesta, Presione nuevamente el boton de Consultar Autorizacion.", SAPbouiCOM.BoMessageTime.bmt_Short, True)

                        oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "GS_SINCRO-No se recibió respuesta de los Web Services", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                    End If

                    Try
                        If TipoDocumento = "LQE" Then
                            GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry, TipoDocumento, _ErrorEcua)
                        Else
                            GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _ErrorEcua)
                        End If
                    Catch ex As Exception
                    End Try

                End If


            Else

                _NumAutorizacionEcua = ""
                _ClaveAccesoEcua = ""
                _EstadoAutorizacionEcua = ""
                _ObservacionEcua = ""
                _FechaAutorizacionEcua = ""

                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Seteando informacion a enviar..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
                Dim doc As String = ""
                If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" Then
                    doc = "factura"
                    oObjetoEcua = ConsultarFactura_Ecuanexus(TipoDocumento, DocEntry, TipoWS)
                ElseIf TipoDocumento = "NDE" Then
                    doc = "nota de debito"
                    oObjetoEcua = ConsultarNotadeDebito_Ecuanexus(TipoDocumento, DocEntry, TipoWS)
                ElseIf TipoDocumento = "NCE" Then
                    doc = "nota de credito"
                    oObjetoEcua = ConsultarNotadeCredito_Ecuanexus(TipoDocumento, DocEntry, TipoWS)
                    'ElseIf TipoDocumento = "GRE" Or TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then 'AGREGADO TLE solicitud de traslado
                    '    oObjeto = ConsultarGuiaDeRemision(TipoDocumento, DocEntry, TipoWS)
                ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then
                    doc = "retención"
                    oObjetoEcua = ConsultarRetencion_Ecuanexus(TipoDocumento, DocEntry, TipoWS)
                ElseIf TipoDocumento = "LQE" Then
                    doc = "liquidación de compras"
                    oObjetoEcua = ConsultarLiquidacionCompra_Ecuanexus(TipoDocumento, DocEntry, TipoWS)
                ElseIf TipoDocumento = "GRE" Or TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then 'AGREGADO TLE solicitud de traslado
                    oObjetoEcua = ConsultarGuiaDeRemision_Ecuanexus(TipoDocumento, DocEntry, TipoWS)
                    'ElseIf TipoDocumento = "RDM" Then
                    '    oObjeto = ConsultarRetencionND(TipoDocumento, DocEntry, TipoWS)
                End If


                'si el objeto no esta vacio

                If Not oObjetoEcua Is Nothing Then

                    Try
                        If Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = "Y" Then
                            _NumeroDeDocumentoSRIEcua = ""
                            _NumeroDeDocumentoSRIEcua = oObjetoEcua.infoTributaria.estab + "-" + oObjetoEcua.infoTributaria.ptoEmi + "-" + oObjetoEcua.infoTributaria.secuencial
                            Utilitario.Util_Log.Escribir_Log("NumeroDeDocumentoSRI: " + _NumeroDeDocumentoSRIEcua.ToString(), "ManejoDeDocumentos")
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error al setear NumeroDeDocumentoSRI: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    Utilitario.Util_Log.Escribir_Log("Enviando documento a Ecuanexus, por favor espere..!!", "ManejoDeDocumentos")
                    If _tipoManejoEcua = "A" Then
                        rsboAppEcua.SetStatusBarMessage("Enviando documento a Ecuanexus, por favor espere..!!", SAPbouiCOM.BoMessageTime.bmt_Long, False)

                        oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Envíando Documento a Ecuanexus", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                    Dim respuesta_WS As String = ""

                    Try

                        Dim enviodoc As New Entidades.EnvioDocumento
                        enviodoc.File = Base64
                        enviodoc.FileName = NombreXML
                        enviodoc.Token = Functions.VariablesGlobales._Token

                        Try
                            Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                            Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_EMISION" + ".xml"
                            If System.IO.Directory.Exists(sRutaCarpeta) Then
                                Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de envio", "ManejoDeDocumentos")

                                Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.EnvioDocumento))
                                Dim writer As TextWriter = New StreamWriter(sRuta)
                                x.Serialize(writer, enviodoc)
                                writer.Close()
                                Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de envio" + sRuta, "frmDocumentosRecibidos")
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                        End Try
                        'Se consulta estado del documento
                        _objetoRespuestaEcu = CoreRest.EnvioDocumento(enviodoc, respuesta_WS)

                        Utilitario.Util_Log.Escribir_Log("ECUANEXUS", "ManejoDeDocumentos")

                    Catch ex As Exception
                        rsboAppEcua.SetStatusBarMessage("Error al consultar estado del documento:  " & ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Function
                    End Try

                    If Not _objetoRespuestaEcu Is Nothing Then

                        Utilitario.Util_Log.Escribir_Log("NO DEVOLVIO NOTHING", "ManejoDeDocumentos")

                        ResulCode = _objetoRespuestaEcu.result.code
                        ResultMensaje = _objetoRespuestaEcu.result.message

                        If ResulCode = "100" And ResultMensaje = "OK" Then
                            oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Respuesta del Ws Code: " + ResulCode + " - mensaje: " + ResultMensaje, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If

                        If _tipoManejoEcua = "A" Then
                            rsboAppEcua.SetStatusBarMessage("Recibiendo respuesta..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        End If

                        _ObservacionEcua = "Respuesta ws - Código: " + IIf(_objetoRespuestaEcu.result.code = Nothing, "", _objetoRespuestaEcu.result.code) + " - Mensaje: " + IIf(_objetoRespuestaEcu.result.message = Nothing, "", _objetoRespuestaEcu.result.message)

                        _EstadoAutorizacionEcua = "1"

                        If TipoDocumento = "LQE" Then
                            result = GrabaDatosAutorizacion_LiquidacionCompra(DocEntry, TipoDocumento)

                        Else
                            result = GrabaDatosAutorizacion(DocEntry, TipoDocumento)
                        End If

                        If result Then
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Proceso terminado con exito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            End If

                        Else
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Ocurrio un Error al Guardar los datos de Autorización..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If

                        End If


                        'If _tipoManejoEcua = "A" Then se comenta debido a que no se obtiene inmeditamente la respuesta
                        '    If Functions.VariablesGlobales._ImpDocAut = "Y" Then
                        '        If state = "100" And String.IsNullOrEmpty(statusMsg) Then
                        '            rsboAppEcua.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + " - Imprimiendo Documento por favor esperar... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        '            mensajeDocAutEcua = ""
                        '            If ImprmirDOcAut(_NumAutorizacionEcua) Then
                        '                rsboAppEcua.SetStatusBarMessage("El documento se imprimió con éxito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        '            Else
                        '                rsboAppEcua.SetStatusBarMessage("Error al imprimir PDF: " + mensajeDocAutEcua.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            End If
                        '        End If
                        '    End If
                        'End If
                    Else

                        ' Controlo Error si no se seteo la factura con los datos de base 
                        _ObservacionEcua = "Ocurrio un error al Consultar los datos de la " + doc + ": " & DocEntry.ToString() + " - " + respuesta_WS.ToString
                        _ErrorEcua = _ObservacionEcua
                        If _tipoManejoEcua = "A" Then
                            rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la " + doc + " en la Base, DocEntry:  " & DocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            rsboAppEcua.SetStatusBarMessage("Error:  " & respuesta_WS.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Ocurrio un error al consultar datos de la " + doc + " en la Base, DocEntry: " & DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                        End If

                        Try
                            If TipoDocumento = "LQE" Then
                                GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry, TipoDocumento, _ErrorEcua)
                            Else
                                GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _ErrorEcua)
                            End If
                        Catch ex As Exception
                        End Try

                    End If

                End If
            End If
            Return _ObservacionEcua

        Catch ex As Exception
            _ErrorEcua = ex.Message
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Error:  " & ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            Return _ErrorEcua + _errorMensajeWSEnvíoEcua
        End Try
    End Function

    Public Function SincronizarDocumento(ByVal tipoDocumento As String, DocEntry As String, ByVal TipoWS As String) As Object
        ' ws.Url = Url ' Seteo la URL en el servicio web
        ' Entorno 2- en Linea, 1- en Batch

        Dim ObjetoRespuesta As Object = Nothing
        Dim mensajeRespuesta As String = ""
        _errorMensajeWSEnvíoEcua = ""
        Dim url As String = ""

        Dim SALIDA_POR_PROXY As String = ""
        SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
        Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
        Dim Proxy_puerto As String = ""
        Dim Proxy_IP As String = ""
        Dim Proxy_Usuario As String = ""
        Dim Proxy_Clave As String = ""

        'Dim wsauto As New WSAutorizacionComp.AutorizacionComprobantesService
        'wsauto.Url = url
        'wsauto.Timeout = 10000

        Try


            Dim WS_EmisionConsul As String

            If _tipoManejoEcua = "S" Then
                WS_EmisionConsul = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WsConsultaEcu")
            Else
                WS_EmisionConsul = Functions.VariablesGlobales._WsEmisionConsultaEcua
            End If



            'OBTENER INFORMACION COMPANIA PARA SINCRONIZAR
            'RUC COMPANIA
            'TIPO DOC
            'NUMDOC xxx-xxx-xxxxxxxxx  en este formato
            'SECERP

            Dim Sincro_ruc As String = "", Sincro_Tipo_doc As String = "", Sincro_sec_ERP As String = "", Sincro_Num_Doc As String



            '--------------------------
            'numero de documento xxx-xxx-xxxxxxxxx  en este formato y ruc
            Dim info_company_numdoc() As String = Get_company_numdoc_by_proveedor(_Nombre_Proveedor_SAP_BOEcua, DocEntry, tipoDocumento)
            'RUC compania

            Sincro_ruc = info_company_numdoc(0) 'cero para ruc
            'NUM Doc

            Sincro_Num_Doc = info_company_numdoc(1) 'uno numero doc


            'tipo documento
            Sincro_Tipo_doc = ObtnerIdTipoDocumentoSRI(tipoDocumento)
            'secuencial
            Sincro_sec_ERP = DocEntry

            If Sincro_ruc = "" Or Sincro_Num_Doc = "" Then
                Return Nothing
            End If

            Utilitario.Util_Log.Escribir_Log("Sincro_ruc " + Sincro_ruc.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Sincro_sec_ERP " + Sincro_sec_ERP.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Sincro_Num_Doc " + Sincro_Num_Doc.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("WS_EmisionConsul " + WS_EmisionConsul.ToString, "ManejoDeDocumentos")

            If WS_EmisionConsul = "" Then
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Else
                    Utilitario.Util_Log.Escribir_Log("No existe informacion del Web Service, revisar Parametrización", "ManejoDeDocumentos")
                End If
                Return Nothing
            End If

            'Dim ws As Object
            Dim respuesta_WS As String = ""
            ObjetoRespuesta = New Entidades.ConsultaDocRespuesta

            Dim ConsultarEstadoDoc As New Entidades.ConsultaDocumento

            If _tipoManejoEcua = "S" Then
                ConsultarEstadoDoc.NombreWs = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "NombreWsEcu")
                ConsultarEstadoDoc.clave = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TokenEcu")
            Else
                ConsultarEstadoDoc.NombreWs = Functions.VariablesGlobales._NombreWsEcua
                ConsultarEstadoDoc.clave = Functions.VariablesGlobales._Token
            End If

            ConsultarEstadoDoc.ruc = Sincro_ruc
            ConsultarEstadoDoc.docType = Sincro_Tipo_doc
            ConsultarEstadoDoc.docNumber = Sincro_Num_Doc

            Utilitario.Util_Log.Escribir_Log("ConsultarEstadoDoc.NombreWs " + ConsultarEstadoDoc.NombreWs.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ConsultarEstadoDoc.clave " + ConsultarEstadoDoc.clave.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ConsultarEstadoDoc.ruc " + ConsultarEstadoDoc.ruc.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ConsultarEstadoDoc.docType " + ConsultarEstadoDoc.docType.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ConsultarEstadoDoc.docNumber " + ConsultarEstadoDoc.docNumber.ToString, "ManejoDeDocumentos")

            ObjetoRespuesta = CoreRest.ConsultaDocumento(ConsultarEstadoDoc, respuesta_WS)
            'ObjetoRespuesta = ws.ConsultarProcesoSincronizadorAX(Sincro_ruc, Sincro_Tipo_doc, Sincro_Num_Doc, Sincro_sec_ERP)

            If Not respuesta_WS = "" Then
                If _tipoManejoEcua = "A" Then

                    oFuncionesAddonEcua.GuardaLOG(tipoDocumento, DocEntry, respuesta_WS, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _errorMensajeWSEnvíoEcua = respuesta_WS
            End If

            If Not ObjetoRespuesta Is Nothing Then
                _NumeroDeDocumentoSRIEcua = Sincro_Num_Doc.ToString
                Return ObjetoRespuesta
            Else
                Utilitario.Util_Log.Escribir_Log("No se recibio respuesta", "ManejoDeDocumentos")
                Return Nothing
            End If

        Catch tx As TimeoutException
            'resp.Estado = "7 En Espera SRI"
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("(SS) WS : " + tx.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Else
                Utilitario.Util_Log.Escribir_Log("(SS) WS : " + tx.Message.ToString(), "ManejoDeDocumentos")
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(tipoDocumento, DocEntry, "WS : " + tx.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            Else
                Utilitario.Util_Log.Escribir_Log("(SS) WS : " + tx.Message.ToString(), "ManejoDeDocumentos")
            End If
            Return ObjetoRespuesta
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("(SS) WS : " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Else
                Utilitario.Util_Log.Escribir_Log("(SS) WS : " + ex.Message.ToString(), "ManejoDeDocumentos")
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(tipoDocumento, DocEntry, "WS : " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            Else
                Utilitario.Util_Log.Escribir_Log("(SS) WS : " + ex.Message.ToString(), "ManejoDeDocumentos")
            End If
            Return Nothing
        End Try

    End Function

    Private Function ObtnerTipoDocumentoEDOC(ByVal tipoDocumento As String) As String

        If tipoDocumento = "FCE" Or tipoDocumento = "FRE" Or tipoDocumento = "FAE" Then
            Return "01"
        ElseIf tipoDocumento = "NDE" Then
            Return "04"
        ElseIf tipoDocumento = "NCE" Then
            Return "03"
        ElseIf tipoDocumento = "TRE" Or tipoDocumento = "TLE" Or tipoDocumento = "GRE" Then
            Return "05"
        ElseIf tipoDocumento = "REE" Or tipoDocumento = "REA" Or tipoDocumento = "RER" Then
            Return "02"
        ElseIf tipoDocumento = "LQE" Then
            Return "06"
        End If

        Return ""

    End Function

    Public Function ObtnerIdTipoDocumentoSRI(ByVal tipoDocumento As String) As String

        If tipoDocumento = "FCE" Or tipoDocumento = "FRE" Or tipoDocumento = "FAE" Then
            Return "01"
        ElseIf tipoDocumento = "NDE" Then
            Return "05"
        ElseIf tipoDocumento = "NCE" Then
            Return "04"
        ElseIf tipoDocumento = "TRE" Or tipoDocumento = "TLE" Or tipoDocumento = "GRE" Then
            Return "06"
        ElseIf tipoDocumento = "REE" Or tipoDocumento = "REA" Or tipoDocumento = "RER" Then
            Return "07"
        ElseIf tipoDocumento = "LQE" Then
            Return "03"
        End If

        Return ""

    End Function


    Public Function Get_company_numdoc_by_proveedor(ByVal nombreProveedor As String, ByVal DocEnty As String, ByVal tipoDocumento As String) As String()

        Dim tabla_SAP As String = ""
        Dim ruc_numdoc() As String = {"", ""}

        If tipoDocumento = "FCE" Or tipoDocumento = "FRE" Or tipoDocumento = "FAE" Or tipoDocumento = "NDE" Then
            tabla_SAP = "OINV"
        ElseIf tipoDocumento = "NCE" Then
            tabla_SAP = "ORIN"
        ElseIf tipoDocumento = "TRE" Then
            tabla_SAP = "OWTR"
        ElseIf tipoDocumento = "GRE" Then
            tabla_SAP = "ODLN"
        ElseIf tipoDocumento = "TLE" Then
            tabla_SAP = "OWTQ"
        ElseIf tipoDocumento = "REE" Or tipoDocumento = "REA" Or tipoDocumento = "RER" Or tipoDocumento = "LQE" Then
            tabla_SAP = "OPCH"
        End If

        'obtener informacion de los textbox

        Dim querySincro As String = ""
        If _tipoManejoEcua = "A" Then

            If tabla_SAP = "OPCH" Then
                If tipoDocumento = "LQE" Then
                    querySincro = Functions.VariablesGlobales._SINCRO_LQE
                Else
                    querySincro = Functions.VariablesGlobales._SINCRO_RT
                End If
            Else

                querySincro = Functions.VariablesGlobales._SINCRO_DOC

            End If
        Else
            If tabla_SAP = "OPCH" Then
                If tipoDocumento = "LQE" Then
                    querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SINCRO_LQE")
                Else
                    querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SINCRO_RET")
                End If
            Else

                querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SINCRO_DOC")

            End If
        End If




        'hacer un replace de tabla actual por lo que esta en la plantilla

        querySincro = querySincro.Replace("TABLA", tabla_SAP)
        querySincro = querySincro.Replace("IDENTIFICADOR", DocEnty)


        Try

            'Realizo Consulta
            Dim dir_est As String = "", dir_pe As String = "", secuencial As String = "", ruc_compania As String = ""
            Dim numeroDOC As String = ""


            Dim r As SAPbobsCOM.Recordset = oFuncionesAddonEcua.getRecordSet(querySincro)

            If r.RecordCount > 0 Then

                dir_est = oFuncionesAddonEcua.nzString(r.Fields.Item("A").Value)
                dir_pe = oFuncionesAddonEcua.nzString(r.Fields.Item("B").Value)
                secuencial = oFuncionesAddonEcua.nzString(r.Fields.Item("C").Value)
                ruc_compania = oFuncionesAddonEcua.nzString(r.Fields.Item("R").Value)

                If Not secuencial.Length = 9 Then
                    secuencial = secuencial.PadLeft(9, "0")
                End If

                numeroDOC = dir_est & "-" & dir_pe & "-" & secuencial

                If numeroDOC.Length = 17 And String.IsNullOrEmpty(ruc_compania) = False Then

                    ruc_numdoc(0) = ruc_compania
                    ruc_numdoc(1) = numeroDOC

                    Return ruc_numdoc
                End If

            End If


        Catch ex As Exception

        End Try



        Return ruc_numdoc

    End Function

#Region "Consulta de Documentos"

    Public Function ConsultarGuiaDeRemision_Ecuanexus(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oGuiaRemision As New Entidades.guiaRemision
        Dim oDestinatario As Entidades.guiaRemisionDestinatario = Nothing
        Dim listaDetinatarios As New List(Of Entidades.guiaRemisionDestinatario)
        Dim listaDetalle As New List(Of Entidades.guiaRemisionDestinatarioDetalle)
        Dim listaDatosAdicional As New List(Of Entidades.guiaRemisionCampoAdicional)

        Try
            Dim SP As String = ""
            If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_ObtenerGuiaDeRemisionEntrega_Ecuanexus "
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_ObtenerGuiaDeRemisionTransferencia_Ecuanexus "
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If

            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONENTREGA_Ecuanexus "
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONTRANSFERENCIA_Ecuanexus "
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONSOLICITUDTRASLADO_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONENTREGA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONTRANSFERENCIA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONENTREGA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONTRANSFERENCIA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONENTREGA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONTRANSFERENCIA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If

            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_SS_OBTENERGUIAREMISIONENTREGA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_SS_OBTENERGUIAREMISIONTRANSFERENCIA_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If


            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            oGuiaRemision.id = "comprobante"
                            oGuiaRemision.version = "1.0.0"

                            Dim oGuiaRemisionInfoTributaria As New Entidades.guiaRemisionInfoTributaria

                            ' OFFLINE 14 NOVIEMBRE 2017
                            'If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                            'oGuiaRemision.ClaveAcceso = Nothing
                            'fechaemision As String = r("FechaEmision").ToString.Replace("/", "")
                            'Dim clave = GenerarClave(fechaemision, r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))
                            'oGuiaRemisionInfoTributaria.claveAcceso = IIf(clave = "", Nothing, clave)
                            'Else
                            'aRemisionInfoTributaria.claveAcceso = r("ClaveAcceso")
                            'Else
                            '    oGuiaRemision.ClaveAcceso = r("ClaveAcceso")
                            'End If
                            oGuiaRemisionInfoTributaria.claveAcceso = Nothing
                            oGuiaRemisionInfoTributaria.ambiente = r("Ambiente")
                            oGuiaRemisionInfoTributaria.tipoEmision = r("TipoEmision")
                            oGuiaRemisionInfoTributaria.razonSocial = r("RazonSocial")

                            If Not r("NombreComercial") = "" Then
                                oGuiaRemisionInfoTributaria.nombreComercial = r("NombreComercial")
                            End If
                            oGuiaRemisionInfoTributaria.ruc = r("Ruc")
                            oGuiaRemisionInfoTributaria.codDoc = r("CodigoDocumento")
                            oGuiaRemisionInfoTributaria.estab = r("Establecimiento")
                            oGuiaRemisionInfoTributaria.ptoEmi = r("PuntoEmision")
                            oGuiaRemisionInfoTributaria.secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemisionInfoTributaria.secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemisionInfoTributaria.secuencial = oGuiaRemisionInfoTributaria.secuencial.PadLeft(9, "0")
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oGuiaRemisionInfoTributaria.agenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oGuiaRemisionInfoTributaria.contribuyenteRimpe = r("ContribuyenteRimpe")
                            End If


                            oGuiaRemisionInfoTributaria.dirMatriz = r("DireccionMatriz")

                            oGuiaRemision.infoTributaria = oGuiaRemisionInfoTributaria


                            Dim oGuiaRemisionInfoGuiaRemision As New Entidades.guiaRemisionInfoGuiaRemision

                            oGuiaRemisionInfoGuiaRemision.dirEstablecimiento = r("DireccionEstablecimiento")
                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemisionInfoGuiaRemision.contribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemisionInfoGuiaRemision.contribuyenteEspecial = Nothing
                            End If



                            oGuiaRemisionInfoGuiaRemision.obligadoContabilidad = r("ObligadoContabilidad")

                            'oGuiaRemision.FechaEmision = r("FechaEmision")
                            oGuiaRemisionInfoGuiaRemision.dirPartida = r("DireccionPartida")
                            oGuiaRemisionInfoGuiaRemision.razonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemisionInfoGuiaRemision.tipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemisionInfoGuiaRemision.rucTransportista = r("RucTranportista")


                            oGuiaRemisionInfoGuiaRemision.fechaIniTransporte = r("FechaInicioTransporte")
                            oGuiaRemisionInfoGuiaRemision.fechaFinTransporte = r("FechaFinTransporte")
                            oGuiaRemisionInfoGuiaRemision.placa = r("Placa")


                            oGuiaRemision.infoGuiaRemision = oGuiaRemisionInfoGuiaRemision

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows

                            oDestinatario = New Entidades.guiaRemisionDestinatario

                            oDestinatario.identificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.razonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.dirDestinatario = r("DirDestinatario")

                            oDestinatario.motivoTraslado = r("MotivoTraslado")
                            oDestinatario.codEstabDestino = r("CodEstabDestino")

                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.ruta = r("Ruta")
                            End If

                            If Not r("CodDocSustento").ToString() = "" Then
                                oDestinatario.codDocSustento = r("CodDocSustento")
                            End If
                            If Not r("NumDocSustento").ToString() = "" Then
                                oDestinatario.numDocSustento = r("NumDocSustento")
                            End If
                            If Not r("NumAutDocSustento").ToString() = "" Then
                                oDestinatario.numAutDocSustento = r("NumAutDocSustento")
                            End If
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.fechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            listaDetinatarios.Add(oDestinatario)

                        Next

                        ' Dim oDestinatarios As New Entidades.gui

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As New Entidades.guiaRemisionDestinatarioDetalle
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.CodigoInterno = r("CodigoPrincipal")
                            itemDetalle.CodigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.Descripcion = r("Descripcion")
                            itemDetalle.Cantidad = r("Cantidad")

                            'Dim listaDetalleDatoAdicional As Object
                            'listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle)

                            ''Adicional 1
                            'If Not r("ConceptoAdicional1") = "0" Then
                            '    Dim itemDetalleDatoAdicional As Object
                            '    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                            '    itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                            '    itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                            '    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            'End If

                            'If Not r("ConceptoAdicional2") = "0" Then
                            '    Dim itemDetalleDatoAdicional As Object
                            '    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                            '    itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                            '    itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                            '    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            'End If

                            'If Not r("ConceptoAdicional3") = "0" Then
                            '    Dim itemDetalleDatoAdicional As Object
                            '    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                            '    itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                            '    itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                            '    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            'End If

                            'itemDetalle.ENTDatoAdicionalGuiaRemisionDetalle = listaDetalleDatoAdicional.ToArray

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalle)
                        Next
                        oDestinatario.detalles = listaDetalle.ToArray
                        ' La variable 'oDestinatario' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        oGuiaRemision.destinatarios = listaDetinatarios.ToArray()
                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicional As New Entidades.guiaRemisionCampoAdicional
                            itemDatoAdicional.nombre = r("Concepto")
                            itemDatoAdicional.Value = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicional)
                        Next
                        oGuiaRemision.infoAdicional = listaDatosAdicional.ToArray
                    End If
                Next
            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & "GR-" & oGuiaRemision.infoTributaria.estab.ToString() & oGuiaRemision.infoTributaria.ptoEmi.ToString() & oGuiaRemision.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando GR...", "ManejoDeDocumentos")

                    'Dim ms As New MemoryStream

                    Dim xmlns As New XmlSerializerNamespaces()
                    xmlns.Add("", "")
                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.guiaRemision))
                    Dim writer As TextWriter = New StreamWriter(sRuta)

                    x.Serialize(writer, oGuiaRemision, xmlns)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado GR..." + sRuta, "ManejoDeDocumentos")

                    'Dim XMLbyte As Byte() = ms.ToArray

                    'Dim base64String As String = ""
                    'base64String = Convert.ToBase64String(XMLbyte)
                End If

                Dim ms As New MemoryStream

                Dim _xmlns As New XmlSerializerNamespaces()
                _xmlns.Add("", "")
                Dim _x As XmlSerializer = New XmlSerializer(GetType(Entidades.guiaRemision))
                Dim _writer As TextWriter = New StreamWriter(ms)
                _x.Serialize(_writer, oGuiaRemision, _xmlns)
                _writer.Close()


                Dim XMLbyte As Byte() = ms.ToArray

                Dim base64String As String = ""
                base64String = Convert.ToBase64String(XMLbyte)

                Utilitario.Util_Log.Escribir_Log("Serializado GR base 64..." + base64String, "ManejoDeDocumentos")

                Utilitario.Util_Log.Escribir_Log("GR CONSULTADA", "ManejoDeDocumentos")

                Base64 = base64String
                NombreXML = "GR-" + oGuiaRemision.infoTributaria.estab + "-" + oGuiaRemision.infoTributaria.ptoEmi + "-" + oGuiaRemision.infoTributaria.secuencial + ".xml"



            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oGuiaRemision
        Catch x As ArgumentException
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Guia de Remisión en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "ArgumentException-Error al Consultar Guia de Remisión con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la oGuiaRemision en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoGR, DocEntry, "Error al Consultar Guia de Remisión con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try


    End Function

    Public Function ConsultarFactura_Ecuanexus(ByVal TipoFactura As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object
        Dim oFactura As Entidades.factura = Nothing
        Dim oFacturaConImpuesto As Entidades.facturaInfoFacturaTotalImpuesto
        Dim listaDetalle As List(Of Entidades.facturaDetalle)
        Dim listaDatosAdicional As List(Of Entidades.facturaCampoAdicional)
        Dim FormasdePago As List(Of Entidades.facturaInfoFacturaPagos)

        'If TipoWS = "NUBE_4_1" Then
        listaDetalle = New List(Of Entidades.facturaDetalle)
        listaDatosAdicional = New List(Of Entidades.facturaCampoAdicional)
        FormasdePago = New List(Of Entidades.facturaInfoFacturaPagos)
        'End If
        Dim ListareembolsoFc As New List(Of Entidades.reembolsoDetalle)
        Dim aplicadoDescuentoAdicional As Boolean = False

        Try
            Dim SP As String = ""

            If TipoFactura = "FAE" Then
                If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "SS_SAP_FE_ObtenerFacturadeVentaAnticipo "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "SS_SAP_FE_ONE_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "SS_SAP_FE_HEI_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "SS_SAP_FE_SYP_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "SS_SAP_FE_TM_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "SS_SAP_FE_SS_ObtenerFacturadeVentaAnticipo "
                End If

                If _tipoManejoEcua = "A" Then
                    If Functions.VariablesGlobales._vgGuardarLog = "Y" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoFactura, DocEntry, "Tipo de factura = " + TipoFactura.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        oFuncionesAddonEcua.GuardaLOG(TipoFactura, DocEntry, "Consultando Factura con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                    End If
                End If


            Else
                If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "SS_SAP_FE_ObtenerFacturadeVenta_Ecuanexus "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "SS_SAP_FE_ONE_OBTENERFACTURADEVENTA_Ecuanexus "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "SS_SAP_FE_HEI_OBTENERFACTURADEVENTA_Ecuanexus "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "SS_SAP_FE_SYP_OBTENERFACTURADEVENTA_Ecuanexus "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "SS_SAP_FE_TM_OBTENERFACTURADEVENTA_Ecuanexus "
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "SS_SAP_FE_SS_ObtenerFacturadeVenta_Ecuanexus "
                End If

                If _tipoManejoEcua = "A" Then
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

                    oFuncionesAddonEcua.GuardaLOG(TipoFactura, DocEntry, "Tipo de factura = " + TipoFactura.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    oFuncionesAddonEcua.GuardaLOG(TipoFactura, DocEntry, "Consultando Factura con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                    'End If
                End If


            End If
            Utilitario.Util_Log.Escribir_Log("SP: " + SP.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ANTES A CONSULTAR", "ManejoDeDocumentos")
            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("INGRESANDO A CONSULTAR", "ManejoDeDocumentos")
            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oFactura = New Entidades.factura
                Dim _tipoFactura As String = ""
                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try

                            For Each r As DataRow In ds.Tables(0).Rows

                                ' MANEJO DE FACTURAS DE EXPORTACION Y REEMBOLSO - 2018-02-18
                                ' Indica que tipo de factura es (0.- Normal, 1.- Exportadores, 2.- Reembolsos)
                                Try
                                    If r("TipoFactura").ToString() = "" Then
                                        _tipoFactura = 0
                                    Else
                                        _tipoFactura = r("TipoFactura")
                                    End If
                                    Utilitario.Util_Log.Escribir_Log(" (0.- Normal, 1.- Exportadores, 2.- Reembolsos)", "ManejoDeDocumentos")
                                    Utilitario.Util_Log.Escribir_Log("Tipo Factura : " & _tipoFactura.ToString(), "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                                    _tipoFactura = 0
                                End Try
                                'Utilitario.Util_Log.Escribir_Log("TipoFactura: " + oFactura.Tipo.ToString, "ManejoDeDocumentos")
                                ' OFFLINE 14 NOVIEMBRE 2017
                                'FAMC 18/02/2019
                                oFactura.id = "comprobante"
                                oFactura.version = "1.0.0"
                                Dim OfacturaInfoTributaria As New Entidades.facturaInfoTributaria

                                'Dim fecha As String = r("FechaEmision").ToString("ddMMyyyy")
                                'Dim idate As String = DateTime.Parse(clinvoicedate.SelectedDate.ToString).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                                'Dim dDate As DateTime = r("FechaEmision")
                                'Dim strDayFirst As String = Format(dDate, "dd/MM/yyyy")
                                'OfacturaInfoTributaria.claveAcceso = GenerarClave(Format(r("FechaEmision"), "dd/MM/yyyy").ToString.Replace("/", ""), r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))

                                'If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                OfacturaInfoTributaria.claveAcceso = Nothing
                                'Dim fechaemision As String = r("FechaEmision").ToString.Replace("/", "")
                                'Dim clave = GenerarClave(fechaemision, r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))
                                'OfacturaInfoTributaria.claveAcceso = IIf(clave = "", Nothing, clave)
                                'Else
                                'turaInfoTributaria.claveAcceso = r("ClaveAcceso")

                                'End If

                                OfacturaInfoTributaria.ambiente = r("Ambiente")
                                OfacturaInfoTributaria.tipoEmision = r("TipoEmision")
                                OfacturaInfoTributaria.razonSocial = r("RazonSocial")

                                If Not r("NombreComercial") = "" Then
                                    OfacturaInfoTributaria.razonSocial = r("NombreComercial")
                                End If

                                'Utilitario.Util_Log.Escribir_Log("NombreComercial: " + oFactura.infoFactura.razonSocialComprador.ToString, "ManejoDeDocumentos")
                                OfacturaInfoTributaria.ruc = r("RUC")
                                'oFactura.Ruc = "0992737964001"
                                OfacturaInfoTributaria.codDoc = r("CodigoDocumento")
                                OfacturaInfoTributaria.estab = r("Establecimiento")
                                OfacturaInfoTributaria.ptoEmi = r("PuntoEmision")
                                OfacturaInfoTributaria.secuencial = r("SecuencialDocumento")
                                If Not OfacturaInfoTributaria.ToString().Length.Equals("9") Then
                                    OfacturaInfoTributaria.secuencial = OfacturaInfoTributaria.secuencial.ToString().PadLeft(9, "0")
                                End If
                                Utilitario.Util_Log.Escribir_Log("oFactura.Secuencial : " & OfacturaInfoTributaria.ToString(), "ManejoDeDocumentos")
                                OfacturaInfoTributaria.dirMatriz = r("DireccionMatriz")

                                If Not r("AgenteRetencion") = "0" Then
                                    OfacturaInfoTributaria.agenteRetencion = r("AgenteRetencion")
                                End If

                                'If Not r("RegimenMicroempresas") = "0" Then
                                '    oFactura.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                                'End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    OfacturaInfoTributaria.contribuyenteRimpe = r("ContribuyenteRimpe")
                                End If


                                oFactura.infoTributaria = OfacturaInfoTributaria

                                Dim OfacturaInfoFactura As New Entidades.facturaInfoFactura

                                OfacturaInfoFactura.fechaEmision = r("FechaEmision")
                                OfacturaInfoFactura.dirEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    OfacturaInfoFactura.contribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    OfacturaInfoFactura.contribuyenteEspecial = Nothing
                                End If

                                OfacturaInfoFactura.obligadoContabilidad = r("ObligadoContabilidad")

                                If Not r("GuiaRemision") = "0" Then
                                    OfacturaInfoFactura.guiaRemision = r("GuiaRemision")
                                End If



                                OfacturaInfoFactura.tipoIdentificacionComprador = r("TipoIdentificadorComprador")

                                OfacturaInfoFactura.razonSocialComprador = r("RazonSocialComprador")
                                OfacturaInfoFactura.identificacionComprador = r("IdentificacionComprador")

                                If Not r("DirComprador") = "" Then
                                    OfacturaInfoFactura.direccionComprador = r("DirComprador")
                                End If


                                OfacturaInfoFactura.totalSinImpuestos = r("TotalSinImpuesto")
                                OfacturaInfoFactura.totalDescuento = r("TotalDescuento")

                                OfacturaInfoFactura.propina = r("Propina")
                                OfacturaInfoFactura.importeTotal = r("ImporteTotal")
                                OfacturaInfoFactura.moneda = r("Moneda")
                                If _tipoFactura = 1 Then

                                    rsboAppEcua.SetStatusBarMessage("Favor espere... Consultando Datos de Factura de Exportación, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                    OfacturaInfoFactura.comercioExterior = r("ComercioExterior")
                                    OfacturaInfoFactura.incoTermFactura = r("IncoTermFactura")
                                    OfacturaInfoFactura.lugarIncoTerm = r("LugarIncoTerm")
                                    OfacturaInfoFactura.paisOrigen = r("PaisOrigen")
                                    OfacturaInfoFactura.puertoEmbarque = r("PuertoEmbarque")
                                    OfacturaInfoFactura.paisDestino = r("PaisDestino")

                                    If r("PuertoDestino") <> "" Then
                                        OfacturaInfoFactura.puertoDestino = r("PuertoDestino")
                                    End If

                                    If r("PaisAdquisicion") <> "" Then
                                        OfacturaInfoFactura.paisAdquisicion = r("PaisAdquisicion")
                                    End If

                                    OfacturaInfoFactura.incoTermTotalSinImpuestos = r("IncoTermTotalSinImpuestos")

                                    If r("FleteInternacional") <> "0" Then
                                        OfacturaInfoFactura.fleteInternacional = r("FleteInternacional")
                                    End If
                                    If r("SeguroInternacional") <> "0" Then
                                        OfacturaInfoFactura.seguroInternacional = r("SeguroInternacional")
                                    End If
                                    If r("GastosAduaneros") <> "0" Then
                                        OfacturaInfoFactura.GastosAduaneros = r("GastosAduaneros")
                                    End If
                                    If r("GastosTransporteOtros") <> "0" Then
                                        OfacturaInfoFactura.GastosTransporteOtros = r("GastosTransporteOtros")
                                    End If


                                End If

                                If _tipoFactura = 2 Then

                                    rsboAppEcua.SetStatusBarMessage("Favor espere... Consultando Datos de Factura de Reembolso Cabecera, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                    If r("CodDocReemb") <> "" Then
                                        OfacturaInfoFactura.codDocReembolso = r("CodDocReemb")
                                    End If

                                    If r("TotalComprobantesReembolso") <> "0" Then
                                        OfacturaInfoFactura.totalComprobantesReembolso = r("TotalComprobantesReembolso")
                                    End If
                                    If r("TotalBaseImponibleReembolso") <> "0" Then
                                        OfacturaInfoFactura.totalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                                    End If
                                    If r("TotalImpuestoReembolso") <> "0" Then
                                        OfacturaInfoFactura.totalImpuestoReembolso = r("TotalImpuestoReembolso")
                                    End If

                                End If
                                'IMPUESTO FACTURA
                                'Impuestos totalizados en la factura.
                                oFactura.infoFactura = OfacturaInfoFactura
                                'Dim lstoFacturaConImpuesto As Entidades.facturaInfoFacturaTotalImpuesto
                                Dim lstoFacturaConImpuesto = New List(Of Entidades.facturaInfoFacturaTotalImpuesto)
                                If r("Base8") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("Codigo8")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentaje8")
                                    oFacturaConImpuesto.baseImponible = r("Base8")
                                    oFacturaConImpuesto.tarifa = r("TarifaIva8")
                                    oFacturaConImpuesto.valor = r("ValorIva8")

                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)

                                End If


                                If r("Base12") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("Codigo12")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentaje12")
                                    oFacturaConImpuesto.baseImponible = r("Base12")
                                    oFacturaConImpuesto.tarifa = r("Tarifa12")
                                    oFacturaConImpuesto.valor = r("ValorIva12")

                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)
                                End If

                                If r("Base13") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("Codigo13")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentaje13")
                                    oFacturaConImpuesto.baseImponible = r("Base13")
                                    oFacturaConImpuesto.tarifa = r("Tarifa13")
                                    oFacturaConImpuesto.valor = r("ValorIva13")
                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If


                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)
                                End If
                                If r("Base0") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("Codigo0")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentaje0")
                                    oFacturaConImpuesto.baseImponible = r("Base0")
                                    oFacturaConImpuesto.tarifa = r("Tarifa0")
                                    oFacturaConImpuesto.valor = r("ValorIva0")

                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If


                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)
                                End If

                                If r("BaseNoi") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("CodigoNoi")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    oFacturaConImpuesto.baseImponible = r("BaseNoi")
                                    oFacturaConImpuesto.tarifa = r("TarifaNoi")
                                    oFacturaConImpuesto.valor = r("ValorIvaNoi")
                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)
                                End If

                                If r("BaseExen") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("CodigoExen")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    oFacturaConImpuesto.baseImponible = r("BaseExen")
                                    oFacturaConImpuesto.tarifa = r("TarifaExen")
                                    oFacturaConImpuesto.valor = r("ValorIvaExen")
                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)
                                End If

                                If r("BaseIce") <> 0 Then

                                    oFacturaConImpuesto = New Entidades.facturaInfoFacturaTotalImpuesto

                                    oFacturaConImpuesto.codigo = r("CodigoIce")
                                    oFacturaConImpuesto.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    oFacturaConImpuesto.baseImponible = r("BaseIce")
                                    oFacturaConImpuesto.tarifa = r("TarifaIce")
                                    oFacturaConImpuesto.valor = r("ValorIvaIce")
                                    If r("DescuentoAdicional") <> "0" Then
                                        oFacturaConImpuesto.descuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstoFacturaConImpuesto.Add(oFacturaConImpuesto)
                                End If
                                oFactura.infoFactura.totalConImpuestos = lstoFacturaConImpuesto.ToArray
                            Next
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim oFacturaDetalle = New Entidades.facturaDetalle

                                oFacturaDetalle.codigoPrincipal = r("CodigoPrincipal")
                                oFacturaDetalle.codigoAuxiliar = r("CodigoAuxiliar")
                                oFacturaDetalle.descripcion = r("Descripcion")
                                oFacturaDetalle.cantidad = r("Cantidad")
                                oFacturaDetalle.precioUnitario = r("PrecioUnitario")
                                oFacturaDetalle.descuento = r("Descuento")
                                oFacturaDetalle.precioTotalSinImpuesto = r("PrecioTotalSinImpuesto")

                                If r("ConceptoAdicional1") = "0" And r("ConceptoAdicional2") = "0" And r("ConceptoAdicional3") = "0" Then
                                    oFacturaDetalle.detallesAdicionales = Nothing
                                Else
                                    Dim listaDetalleDatoAdicional = New List(Of Entidades.facturadetAdicional)
                                    'Adicional1
                                    If Not r("ConceptoAdicional1") = "0" Then
                                        Dim itemDetalleDatoAdicional As New Entidades.facturadetAdicional
                                        itemDetalleDatoAdicional.nombre = r("ConceptoAdicional1")
                                        itemDetalleDatoAdicional.valor = r("NombreAdicional1")
                                        listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)

                                    End If
                                    'oFacturaDetalle.detallesAdicionales. listaDetalleDatoAdici.ToArray

                                    'Adicional2
                                    If Not r("ConceptoAdicional2") = "0" Then
                                        Dim itemDetalleDatoAdicional As New Entidades.facturadetAdicional
                                        itemDetalleDatoAdicional.nombre = r("ConceptoAdicional2")
                                        itemDetalleDatoAdicional.valor = r("NombreAdicional2")
                                        listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)

                                    End If

                                    'Adicional3
                                    If Not r("ConceptoAdicional3") = "0" Then
                                        Dim itemDetalleDatoAdicional As New Entidades.facturadetAdicional
                                        itemDetalleDatoAdicional.nombre = r("ConceptoAdicional3")
                                        itemDetalleDatoAdicional.valor = r("NombreAdicional3")
                                        listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)

                                        'oFacturaDetalle.detallesAdicionales.detAdicional = itemDetalleDatoAdicional
                                    End If

                                    'listaDetalleDatoAdici.detAdicional = listaDetalleDatoAdicional.ToArray

                                    oFacturaDetalle.detallesAdicionales = listaDetalleDatoAdicional.ToArray
                                End If





                                'Dim listaImpuestosDetalles = New Entidades.facturaDetalleImpuestos
                                Dim listaImpuestosDetalles = New List(Of Entidades.facturaDetalleImpuesto)

                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%

                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA8" Then ' 12%
                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA" Then ' 12%
                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA13" Then ' 12%
                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                    Dim impdetalleIVA As New Entidades.facturaDetalleImpuesto
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestosDetalles.Add(impdetalleIVA)
                                End If



                                oFacturaDetalle.impuestos = listaImpuestosDetalles.ToArray

                                'agrego detalle a la lista
                                listaDetalle.Add(oFacturaDetalle)
                            Next
                            oFactura.detalles = listaDetalle.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalFac As New Entidades.facturaCampoAdicional

                                itemDatoAdicionalFac.nombre = r("Concepto")
                                itemDatoAdicionalFac.Value = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalFac)
                            Next
                            oFactura.infoAdicional = listaDatosAdicional.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Cabecera Campo Adicional: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Informacion Adicional: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 3 Then
                        Try
                            Dim pagos = New Entidades.facturaInfoFacturaPagos
                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim Pago As New Entidades.facturaInfoFacturaPagosPago

                                If r("FormaPago") <> "" Then


                                    Pago.formaPago = r("FormaPago")
                                    Pago.total = r("Total")
                                    If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                        Pago.plazo = Nothing
                                    Else
                                        Pago.plazo = r("Plazo")
                                    End If
                                    If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                        Pago.unidadTiempo = Nothing
                                    Else
                                        Pago.unidadTiempo = r("UnidadTiempo")
                                    End If

                                    pagos.pago = Pago
                                Else
                                    pagos.pago = Nothing
                                End If
                            Next

                            oFactura.infoFactura.pagos = pagos
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 4 Then
                        Try
                            If _tipoFactura = 2 Then

                                rsboAppEcua.SetStatusBarMessage("Favor espere... Consultando Datos de Factura de Reembolso detalle, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                For Each r As DataRow In ds.Tables(4).Rows

                                    Dim reembolsoFc As New Entidades.reembolsoDetalle

                                    If Not r("TipoIdentificacionProveedorReembolso") = "" Then
                                        reembolsoFc.tipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                                    End If
                                    If Not r("IdentificacionProveedorReembolso") = "" Then
                                        reembolsoFc.identificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                                    End If
                                    If Not r("CodPaisPagoProveedorReembolso") = "" Then
                                        reembolsoFc.codPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                                    End If
                                    If Not r("TipoProveedorReembolso") = "" Then
                                        reembolsoFc.tipoProveedorReembolso = r("TipoProveedorReembolso")
                                    End If
                                    If Not r("CodDocReembolso") = "" Then
                                        reembolsoFc.codDocReembolso = r("CodDocReembolso")
                                    End If
                                    If Not r("EstabDocReembolso") = "" Then
                                        reembolsoFc.estabDocReembolso = r("EstabDocReembolso")
                                    End If
                                    If Not r("PtoEmiDocReembolso") = "" Then
                                        reembolsoFc.ptoEmiDocReembolso = r("PtoEmiDocReembolso")
                                    End If
                                    If Not r("SecuencialDocReembolso") = "" Then
                                        reembolsoFc.secuencialDocReembolso = r("SecuencialDocReembolso")
                                    End If
                                    If Not r("FechaEmisionDocReembolso") = "" Then
                                        reembolsoFc.fechaEmisionDocReembolso = r("FechaEmisionDocReembolso")
                                    End If
                                    If Not r("NumeroAutorizacionDocReem") = "" Then
                                        reembolsoFc.numeroautorizacionDocReemb = r("NumeroAutorizacionDocReem")
                                    End If

                                    'Dim listaImpReembolsoLQ As New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto)
                                    Dim ListareembolsoFcImp As New List(Of Entidades.facturaDetalleReembolsoImpuesto)

                                    If r("Base8") <> 0 Then

                                        Dim reembolsoFcImp As New Entidades.facturaDetalleReembolsoImpuesto

                                        reembolsoFcImp.codigo = r("Codigo8")
                                        reembolsoFcImp.codigoPorcentaje = r("CodigoPorcentaje8")
                                        reembolsoFcImp.tarifa = r("Tarifa8")
                                        reembolsoFcImp.baseImponibleReembolso = r("Base8")
                                        reembolsoFcImp.impuestoReembolso = r("ValorIvaReem8")

                                        ListareembolsoFcImp.Add(reembolsoFcImp)
                                    End If

                                    If r("Base12") <> 0 Then
                                        Dim reembolsoFcImp As New Entidades.facturaDetalleReembolsoImpuesto
                                        reembolsoFcImp.codigo = r("Codigo12")
                                        reembolsoFcImp.codigoPorcentaje = r("CodigoPorcentaje12")
                                        reembolsoFcImp.tarifa = r("Tarifa12")
                                        reembolsoFcImp.baseImponibleReembolso = r("Base12")
                                        reembolsoFcImp.impuestoReembolso = r("ValorIvaReem12")

                                        ListareembolsoFcImp.Add(reembolsoFcImp)
                                    End If
                                    If r("Base13") <> 0 Then
                                        Dim reembolsoFcImp As New Entidades.facturaDetalleReembolsoImpuesto
                                        reembolsoFcImp.codigo = r("Codigo13")
                                        reembolsoFcImp.codigoPorcentaje = r("CodigoPorcentaje13")
                                        reembolsoFcImp.tarifa = r("Tarifa13")
                                        reembolsoFcImp.baseImponibleReembolso = r("Base13")
                                        reembolsoFcImp.impuestoReembolso = r("ValorIvaReem13")

                                        ListareembolsoFcImp.Add(reembolsoFcImp)
                                    End If
                                    If r("Base0") <> 0 Then
                                        Dim reembolsoFcImp As New Entidades.facturaDetalleReembolsoImpuesto
                                        reembolsoFcImp.codigo = r("Codigo0")
                                        reembolsoFcImp.codigoPorcentaje = r("CodigoPorcentaje0")
                                        reembolsoFcImp.tarifa = r("Tarifa0")
                                        reembolsoFcImp.baseImponibleReembolso = r("Base0")
                                        reembolsoFcImp.impuestoReembolso = r("ValorIvaReem0")

                                        ListareembolsoFcImp.Add(reembolsoFcImp)
                                    End If
                                    If r("BaseNoi") <> 0 Then
                                        Dim reembolsoFcImp As New Entidades.facturaDetalleReembolsoImpuesto
                                        reembolsoFcImp.codigo = r("CodigoNoi")
                                        reembolsoFcImp.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                        reembolsoFcImp.tarifa = r("TarifaNoi")
                                        reembolsoFcImp.baseImponibleReembolso = r("BaseNoi")
                                        reembolsoFcImp.impuestoReembolso = r("ValorIvaReemNoi")

                                        ListareembolsoFcImp.Add(reembolsoFcImp)
                                    End If
                                    If r("BaseExen") <> 0 Then
                                        Dim reembolsoFcImp As New Entidades.facturaDetalleReembolsoImpuesto
                                        reembolsoFcImp.codigo = r("CodigoExen")
                                        reembolsoFcImp.codigoPorcentaje = r("CodigoPorcentajeExen")
                                        reembolsoFcImp.tarifa = r("TarifaExen")
                                        reembolsoFcImp.baseImponibleReembolso = r("BaseExen")
                                        reembolsoFcImp.impuestoReembolso = r("ValorIvaReemExen")

                                        ListareembolsoFcImp.Add(reembolsoFcImp)
                                    End If
                                    reembolsoFc.detalleImpuestos = ListareembolsoFcImp.ToArray

                                    ListareembolsoFc.Add(reembolsoFc)

                                Next
                                oFactura.reembolsos = ListareembolsoFc.ToArray
                            Else
                                oFactura.reembolsos = Nothing
                            End If
                            'If ListareembolsoFc.Count > 0 Then
                            '    oFactura.reembolsos = ListareembolsoFc.ToArray
                            'Else
                            '    oFactura.reembolsos = Nothing
                            'End If

                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Reembolso: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Reembolso: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    End If

                Next

            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & "FC-" & oFactura.infoTributaria.estab.ToString() & oFactura.infoTributaria.ptoEmi.ToString() & oFactura.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")

                    'Dim ms As New MemoryStream

                    Dim xmlns As New XmlSerializerNamespaces()
                    xmlns.Add("", "")
                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.factura))
                    Dim writer As TextWriter = New StreamWriter(sRuta)

                    x.Serialize(writer, oFactura, xmlns)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")

                    'Dim XMLbyte As Byte() = ms.ToArray

                    'Dim base64String As String = ""
                    'base64String = Convert.ToBase64String(XMLbyte)
                End If

                Dim ms As New MemoryStream

                Dim _xmlns As New XmlSerializerNamespaces()
                _xmlns.Add("", "")
                Dim _x As XmlSerializer = New XmlSerializer(GetType(Entidades.factura))
                Dim _writer As TextWriter = New StreamWriter(ms)
                _x.Serialize(_writer, oFactura, _xmlns)
                _writer.Close()


                Dim XMLbyte As Byte() = ms.ToArray

                Dim base64String As String = ""
                base64String = Convert.ToBase64String(XMLbyte)

                Utilitario.Util_Log.Escribir_Log("Serializado base 64..." + base64String, "ManejoDeDocumentos")

                Utilitario.Util_Log.Escribir_Log("FACTURA CONSULTADA", "ManejoDeDocumentos")

                Base64 = base64String
                NombreXML = "FC-" + oFactura.infoTributaria.estab + "-" + oFactura.infoTributaria.ptoEmi + "-" + oFactura.infoTributaria.secuencial + ".xml"



            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try
            Return oFactura

        Catch x As ArgumentException
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoFactura, DocEntry, "ArgumentException-Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoFactura, DocEntry, "Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try
    End Function

    Public Function ConsultarNotadeCredito_Ecuanexus(ByVal TipoNC As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object


        Dim oNotaCredito As New Entidades.notaCredito
        Dim listaDetalle As New List(Of Entidades.notaCreditoDetalle)
        Dim listaDatosAdicional As New List(Of Entidades.notaCreditoCampoAdicional)

        Dim SP As String = ""

        Try
            If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "SS_SAP_FE_ObtenerNotaDeCredito_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "SS_SAP_FE_ONE_OBTENERNOTADECREDITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "SS_SAP_FE_HEI_OBTENERNOTADECREDITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "SS_SAP_FE_SYP_OBTENERNOTADECREDITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "SS_SAP_FE_TM_OBTENERNOTADECREDITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                SP = "SS_SAP_FE_SS_OBTENERNOTADECREDITO_Ecuanexus "
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoNC, DocEntry, "Consultando Nota de Crédito con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            Dim ds As DataSet = EjecutarSP(SP, DocEntry)
            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows
                            Try
                                oNotaCredito.id = "comprobante"
                                oNotaCredito.version = "1.0.0"

                                'INFORMACION TRIBUTARIA
                                Dim oNcInfoTributaria As New Entidades.notaCreditoInfoTributaria
                                ' OFFLINE 14 NOVIEMBRE 2017
                                'If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oNcInfoTributaria.claveAcceso = Nothing
                                'Dim clave = GenerarClave(r("FechaEmision").ToString.Replace("/", ""), r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))
                                'oNcInfoTributaria.claveAcceso = IIf(clave = "", Nothing, clave)
                                'Else
                                'oNcInfoTributaria.claveAcceso = r("ClaveAcceso")
                                'End If


                                oNcInfoTributaria.ambiente = r("Ambiente")
                                oNcInfoTributaria.tipoEmision = r("TipoEmision")

                                oNcInfoTributaria.razonSocial = r("RazonSocial")

                                If Not r("NombreComercial") = "" Then
                                    oNcInfoTributaria.nombreComercial = r("NombreComercial")
                                End If

                                oNcInfoTributaria.ruc = r("Ruc")
                                'oNotaCredito.Ruc = "0992737964001"

                                oNcInfoTributaria.codDoc = r("CodigoDocumento")
                                oNcInfoTributaria.estab = r("Establecimiento")
                                oNcInfoTributaria.ptoEmi = r("PuntoEmision")
                                oNcInfoTributaria.secuencial = r("SecuencialDocumento")
                                If Not oNcInfoTributaria.secuencial.ToString().Length.Equals("9") Then
                                    oNcInfoTributaria.secuencial = oNcInfoTributaria.secuencial.PadLeft(9, "0")
                                End If
                                oNcInfoTributaria.dirMatriz = r("DireccionMatriz")

                                If Not r("AgenteRetencion") = "0" Then
                                    oNcInfoTributaria.agenteRetencion = r("AgenteRetencion")
                                End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    oNcInfoTributaria.contribuyenteRimpe = r("ContribuyenteRimpe")
                                End If

                                oNotaCredito.infoTributaria = oNcInfoTributaria
                                'FIN INFORMACION TRIBUTARIA

                                'INFO NOTA CREDITO
                                Dim oNcInfoNc As New Entidades.notaCreditoInfonotaCredito

                                oNcInfoNc.fechaEmision = r("FechaEmision")
                                oNcInfoNc.dirEstablecimiento = r("DireccionEstablecimiento")

                                oNcInfoNc.tipoIdentificacionComprador = r("TipoIdentificadorComprador")

                                oNcInfoNc.razonSocialComprador = r("RazonSocialComprador")
                                oNcInfoNc.identificacionComprador = r("IdentificacionComprador")


                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oNcInfoNc.contribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oNcInfoNc.contribuyenteEspecial = Nothing
                                End If

                                oNcInfoNc.obligadoContabilidad = r("ObligadoContabilidad")

                                If Not r("Rise") = "" Then
                                    oNcInfoNc.rise = r("Rise")
                                End If

                                oNcInfoNc.codDocModificado = r("codDocModificado")
                                oNcInfoNc.numDocModificado = r("numDocModificado")
                                oNcInfoNc.fechaEmisionDocSustento = r("FechaEmisionDocModificado")



                                oNcInfoNc.totalSinImpuestos = r("TotalSinImpuesto")
                                oNcInfoNc.valorModificacion = r("ValorModificacion")
                                oNcInfoNc.motivo = r("Motivo")
                                oNcInfoNc.moneda = r("Moneda")

                                oNotaCredito.infoNotaCredito = oNcInfoNc
                                'FIN INFO NOTA CREDITO

                                Dim lstimpNc As New List(Of Entidades.notaCreditoInfonotaCreditoTotalImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("Codigo8")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impNcIVA.tarifa = r("Tarifa8")
                                    impNcIVA.baseImponible = r("Base8")
                                    impNcIVA.valor = r("ValorIva8")

                                    lstimpNc.Add(impNcIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("Codigo12")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impNcIVA.tarifa = r("Tarifa12")
                                    impNcIVA.baseImponible = r("Base12")
                                    impNcIVA.valor = r("ValorIva12")

                                    lstimpNc.Add(impNcIVA)
                                End If

                                If r("Base13") <> 0 Then

                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("Codigo13")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impNcIVA.tarifa = r("Tarifa13")
                                    impNcIVA.baseImponible = r("Base13")
                                    impNcIVA.valor = r("ValorIva13")

                                    lstimpNc.Add(impNcIVA)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("Codigo0")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impNcIVA.tarifa = r("Tarifa0")
                                    impNcIVA.baseImponible = r("Base0")
                                    impNcIVA.valor = r("ValorIva0")

                                    lstimpNc.Add(impNcIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("CodigoNoi")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impNcIVA.tarifa = r("TarifaNoi")
                                    impNcIVA.baseImponible = r("BaseNoi")
                                    impNcIVA.valor = r("ValorIvaNoi")

                                    lstimpNc.Add(impNcIVA)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("CodigoExen")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impNcIVA.tarifa = r("TarifaExen")
                                    impNcIVA.baseImponible = r("BaseExen")
                                    impNcIVA.valor = r("ValorIvaExen")

                                    lstimpNc.Add(impNcIVA)
                                End If

                                If r("BaseIce") <> 0 Then
                                    Dim impNcIVA As New Entidades.notaCreditoInfonotaCreditoTotalImpuesto

                                    impNcIVA.codigo = r("CodigoIce")
                                    impNcIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    impNcIVA.tarifa = r("TarifaIce")
                                    impNcIVA.baseImponible = r("BaseIce")
                                    impNcIVA.valor = r("ValorIvaIce")

                                    lstimpNc.Add(impNcIVA)
                                End If
                                oNotaCredito.infoNotaCredito.totalConImpuestos = lstimpNc.ToArray
                            Catch ex As Exception
                                If _tipoManejoEcua = "A" Then
                                    rsboAppEcua.SetStatusBarMessage("Cabecera nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                                _ErrorEcua = "Cabecera: " + ex.Message.ToString()
                                Utilitario.Util_Log.Escribir_Log("Cabcera nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                Return Nothing
                            End Try

                        Next
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows
                                Dim DetalleNC As New Entidades.notaCreditoDetalle
                                Try
                                    DetalleNC.codigoInterno = r("CodigoPrincipal")
                                    DetalleNC.codigoAdicional = r("CodigoAuxiliar")
                                    DetalleNC.descripcion = r("Descripcion")
                                    DetalleNC.cantidad = r("Cantidad")
                                    DetalleNC.precioUnitario = r("PrecioUnitario")
                                    DetalleNC.descuento = r("Descuento")
                                    DetalleNC.precioTotalSinImpuesto = r("PrecioTotalSinImpuesto")
                                Catch ex As Exception
                                    If _tipoManejoEcua = "A" Then
                                        rsboAppEcua.SetStatusBarMessage("DETALLE nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End If
                                    Utilitario.Util_Log.Escribir_Log("DETALLE nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                End Try

                                Try

                                    ''Datos adicionales de cada detalle del item                                                          

                                    If r("ConceptoAdicional1") = "0" And r("ConceptoAdicional2") = "0" And r("ConceptoAdicional3") = "0" Then
                                        DetalleNC.detallesAdicionales = Nothing
                                    Else
                                        Dim listaDetalleDatoAdicional As New List(Of Entidades.notaCreditodetAdicional)
                                        'Adicional1
                                        If Not r("ConceptoAdicional1") = "0" Then
                                            Dim itemDetalleDatoAdicional As New Entidades.notaCreditodetAdicional
                                            itemDetalleDatoAdicional.nombre = r("ConceptoAdicional1")
                                            itemDetalleDatoAdicional.valor = r("NombreAdicional1")
                                            listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                                        End If

                                        'Adicional2
                                        If Not r("ConceptoAdicional2") = "0" Then
                                            Dim itemDetalleDatoAdicional2 As New Entidades.notaCreditodetAdicional
                                            itemDetalleDatoAdicional2.nombre = r("ConceptoAdicional2")
                                            itemDetalleDatoAdicional2.valor = r("NombreAdicional2")
                                            listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional2)
                                        End If

                                        'Adicional3
                                        If Not r("ConceptoAdicional3") = "0" Then
                                            Dim itemDetalleDatoAdicional3 As New Entidades.notaCreditodetAdicional
                                            itemDetalleDatoAdicional3.nombre = r("ConceptoAdicional3")
                                            itemDetalleDatoAdicional3.valor = r("NombreAdicional3")
                                            listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional3)
                                        End If

                                        DetalleNC.detallesAdicionales = listaDetalleDatoAdicional.ToArray
                                    End If


                                Catch ex As Exception
                                    If _tipoManejoEcua = "A" Then
                                        rsboAppEcua.SetStatusBarMessage("ADICIONAL detalle nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End If
                                    Utilitario.Util_Log.Escribir_Log("ADICIONAL detalle nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                    _ErrorEcua = "ADICIONAL nota de credito error " + ex.Message.ToString()
                                    Return Nothing
                                End Try

                                Dim lstimpdetalle As New List(Of Entidades.notaCreditoDetalleImpuesto)

                                Try
                                    If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("Codigo")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                        impdetalleIVA.tarifa = r("Tarifa")
                                        impdetalleIVA.baseImponible = r("BaseImponible")
                                        impdetalleIVA.valor = r("TotalIva")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If

                                    If r("TaxCodeAp") = "IVA8" Then ' 0%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("Codigo")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                        impdetalleIVA.tarifa = r("Tarifa")
                                        impdetalleIVA.baseImponible = r("BaseImponible")
                                        impdetalleIVA.valor = r("TotalIva")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If

                                    If r("TaxCodeAp") = "IVA" Then ' 12%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("Codigo")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                        impdetalleIVA.tarifa = r("Tarifa")
                                        impdetalleIVA.baseImponible = r("BaseImponible")
                                        impdetalleIVA.valor = r("TotalIva")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If

                                    If r("TaxCodeAp") = "IVA13" Then ' 12%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("Codigo")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                        impdetalleIVA.tarifa = r("Tarifa")
                                        impdetalleIVA.baseImponible = r("BaseImponible")
                                        impdetalleIVA.valor = r("TotalIva")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If

                                    If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("Codigo")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                        impdetalleIVA.tarifa = r("Tarifa")
                                        impdetalleIVA.baseImponible = r("BaseImponible")
                                        impdetalleIVA.valor = r("TotalIva")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If

                                    If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("Codigo")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                        impdetalleIVA.tarifa = r("Tarifa")
                                        impdetalleIVA.baseImponible = r("BaseImponible")
                                        impdetalleIVA.valor = r("TotalIva")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If

                                    If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                        Dim impdetalleIVA As New Entidades.notaCreditoDetalleImpuesto

                                        impdetalleIVA.codigo = r("CodigoIce")
                                        impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                        impdetalleIVA.tarifa = r("TarifaIce")
                                        impdetalleIVA.baseImponible = r("BaseImponibleIce")
                                        impdetalleIVA.valor = r("TotalIvaIce")

                                        lstimpdetalle.Add(impdetalleIVA)
                                    End If
                                Catch ex As Exception
                                    If _tipoManejoEcua = "A" Then
                                        rsboAppEcua.SetStatusBarMessage("IMPUESTO nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End If
                                    Utilitario.Util_Log.Escribir_Log("IMPUESTO nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                    _ErrorEcua = "IMPUESTO nota de credito error : " + ex.Message.ToString()
                                    Return Nothing
                                End Try

                                DetalleNC.impuestos = lstimpdetalle.ToArray

                                'agrego detalle a la lista
                                listaDetalle.Add(DetalleNC)
                            Next
                            oNotaCredito.detalles = listaDetalle.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("detalle nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("detalle nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            _ErrorEcua = "detalle nota de credito error : " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalFac As New Entidades.notaCreditoCampoAdicional
                                itemDatoAdicionalFac.nombre = r("Concepto")
                                itemDatoAdicionalFac.Value = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalFac)
                            Next
                            oNotaCredito.infoAdicional = listaDatosAdicional.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("informacion adicional nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("informacion adicional nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            _ErrorEcua = "adicionales de nota de credito error : " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    End If
                Next

            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & "NC-" & oNotaCredito.infoTributaria.estab.ToString() & oNotaCredito.infoTributaria.ptoEmi.ToString() & oNotaCredito.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando NC...", "ManejoDeDocumentos")

                    'Dim ms As New MemoryStream

                    Dim xmlns As New XmlSerializerNamespaces()
                    xmlns.Add("", "")
                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.notaCredito))
                    Dim writer As TextWriter = New StreamWriter(sRuta)

                    x.Serialize(writer, oNotaCredito, xmlns)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado NC..." + sRuta, "ManejoDeDocumentos")

                    'Dim XMLbyte As Byte() = ms.ToArray

                    'Dim base64String As String = ""
                    'base64String = Convert.ToBase64String(XMLbyte)
                End If

                Dim ms As New MemoryStream

                Dim _xmlns As New XmlSerializerNamespaces()
                _xmlns.Add("", "")
                Dim _x As XmlSerializer = New XmlSerializer(GetType(Entidades.notaCredito))
                Dim _writer As TextWriter = New StreamWriter(ms)
                _x.Serialize(_writer, oNotaCredito, _xmlns)
                _writer.Close()


                Dim XMLbyte As Byte() = ms.ToArray

                Dim base64String As String = ""
                base64String = Convert.ToBase64String(XMLbyte)

                Utilitario.Util_Log.Escribir_Log("NC Serializado base 64..." + base64String, "ManejoDeDocumentos")

                Utilitario.Util_Log.Escribir_Log("NC CONSULTADA", "ManejoDeDocumentos")

                Base64 = base64String
                NombreXML = "NC-" + oNotaCredito.infoTributaria.estab + "-" + oNotaCredito.infoTributaria.ptoEmi + "-" + oNotaCredito.infoTributaria.secuencial + ".xml"



            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error NC: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oNotaCredito

        Catch x As ArgumentException
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Nota de Credito en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoNC, DocEntry, "ArgumentException-Error al Consultar Nota de Credito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            Return Nothing
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la oNotaCredito en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoNC, DocEntry, "Error al Consultar Nota de Credito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try
    End Function

    Public Function ConsultarNotadeDebito_Ecuanexus(ByVal TipoND As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oNotaDebito As New Entidades.notaDebito
        Dim listaMotivo As New List(Of Entidades.notaDebitomotivo)
        Dim listaDatosAdicional As New List(Of Entidades.notaDebitoCampoAdicional)
        Dim FormasdePago As New List(Of Entidades.notaDebitoInfonotaDebitoPagosPago)

        Try

            Dim SP As String = ""

            If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "SS_SAP_FE_ObtenerNotaDeDebito_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "SS_SAP_FE_ONE_OBTENERNOTADEDEBITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "SS_SAP_FE_HEI_OBTENERNOTADEDEBITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "SS_SAP_FE_SYP_OBTENERNOTADEDEBITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "SS_SAP_FE_TM_OBTENERNOTADEDEBITO_Ecuanexus "
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                SP = "SS_SAP_FE_SS_OBTENERNOTADEDEBITO_Ecuanexus "
            End If

            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoND, DocEntry, "Consultando Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If



            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then


                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            Try
                                'INFO TRIBUTARIA
                                oNotaDebito.id = "comprobante"
                                oNotaDebito.version = "1.0.0"

                                Dim oNdInfoTributaria As New Entidades.notaDebitoInfoTributaria

                                'If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oNdInfoTributaria.claveAcceso = Nothing
                                'Dim clave = GenerarClave(r("FechaEmision").ToString.Replace("/", ""), r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))
                                'oNdInfoTributaria.claveAcceso = IIf(clave = "", Nothing, clave)
                                'Else
                                'oNdInfoTributaria.claveAcceso = r("ClaveAcceso")
                                '

                                oNdInfoTributaria.ambiente = r("Ambiente")
                                oNdInfoTributaria.tipoEmision = r("TipoEmision")

                                oNdInfoTributaria.razonSocial = r("RazonSocial")

                                If Not r("NombreComercial") = "" Then
                                    oNdInfoTributaria.nombreComercial = r("NombreComercial")
                                End If

                                oNdInfoTributaria.ruc = r("Ruc")


                                oNdInfoTributaria.codDoc = r("CodigoDocumento")
                                oNdInfoTributaria.estab = r("Establecimiento")
                                oNdInfoTributaria.ptoEmi = r("PuntoEmision")
                                oNdInfoTributaria.secuencial = r("SecuencialDocumento")
                                If Not oNdInfoTributaria.secuencial.ToString().Length.Equals("9") Then
                                    oNdInfoTributaria.secuencial = oNdInfoTributaria.secuencial.PadLeft(9, "0")
                                End If
                                oNdInfoTributaria.dirMatriz = r("DireccionMatriz")

                                If Not r("AgenteRetencion") = "0" Then
                                    oNdInfoTributaria.agenteRetencion = r("AgenteRetencion")
                                End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    oNdInfoTributaria.contribuyenteRimpe = r("ContribuyenteRimpe")
                                End If

                                oNotaDebito.infoTributaria = oNdInfoTributaria

                                'FIN INFO TRIBUTARIA

                                'INFO NOTA DEBITO
                                Dim oNdInfoNotaDebito As New Entidades.notaDebitoInfonotaDebito

                                oNdInfoNotaDebito.fechaEmision = r("FechaEmision")
                                oNdInfoNotaDebito.dirEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oNdInfoNotaDebito.contribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oNdInfoNotaDebito.contribuyenteEspecial = Nothing
                                End If


                                oNdInfoNotaDebito.obligadoContabilidad = r("ObligadoContabilidad")

                                oNdInfoNotaDebito.codDocModificado = r("codDocModificado")
                                oNdInfoNotaDebito.numDocModificado = r("numDocModificado")
                                oNdInfoNotaDebito.fechaEmisionDocSustento = r("FechaEmisionDocModificado")

                                oNdInfoNotaDebito.tipoIdentificacionComprador = r("TipoIdentificadorComprador")

                                oNdInfoNotaDebito.razonSocialComprador = r("RazonSocialComprador")
                                oNdInfoNotaDebito.identificacionComprador = r("IdentificacionComprador")

                                oNdInfoNotaDebito.totalSinImpuestos = r("TotalSinImpuesto")

                                oNdInfoNotaDebito.valorTotal = r("ImporteTotal")

                                If Not r("DirComprador") = "" Then
                                    oNdInfoNotaDebito.direccionComprador = r("DirComprador")
                                End If

                                oNotaDebito.infoNotaDebito = oNdInfoNotaDebito

                                Dim lstimpNd As New List(Of Entidades.notaDebitoInfonotaDebitoTotalImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As New Entidades.notaDebitoInfonotaDebitoTotalImpuesto

                                    impfaIVA.codigo = r("Codigo8")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impfaIVA.tarifa = r("Tarifa8")
                                    impfaIVA.baseImponible = r("Base8")
                                    impfaIVA.valor = r("ValorIva8")

                                    lstimpNd.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As New Entidades.notaDebitoInfonotaDebitoTotalImpuesto
                                    'impfaIVA.Codigo = "2"
                                    'impfaIVA.CodigoPorcentaje = "2"
                                    'impfaIVA.Tarifa = "12"
                                    impfaIVA.codigo = r("Codigo12")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impfaIVA.tarifa = r("Tarifa12")
                                    impfaIVA.baseImponible = r("Base12")
                                    'impfaIVA.Valor = r("ImpuestoTotal")
                                    impfaIVA.valor = r("ValorIva12")
                                    lstimpNd.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As New Entidades.notaDebitoInfonotaDebitoTotalImpuesto
                                    'impfaIVA.Codigo = "2"
                                    'impfaIVA.CodigoPorcentaje = "3"
                                    'impfaIVA.Tarifa = "14"
                                    impfaIVA.codigo = r("Codigo13")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impfaIVA.tarifa = r("Tarifa13")
                                    impfaIVA.baseImponible = r("Base13")
                                    'impfaIVA.Valor = r("ImpuestoTotal")
                                    impfaIVA.valor = r("ValorIva13")
                                    lstimpNd.Add(impfaIVA)
                                End If


                                If r("Base0") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.notaDebitoInfonotaDebitoTotalImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "0"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.codigo = r("Codigo0")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impfaNOIVA.tarifa = r("Tarifa0")
                                    impfaNOIVA.baseImponible = r("Base0")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.valor = r("ValorIva0")
                                    lstimpNd.Add(impfaNOIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.notaDebitoInfonotaDebitoTotalImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "6"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.codigo = r("CodigoNoi")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impfaNOIVA.tarifa = r("TarifaNoi")
                                    impfaNOIVA.baseImponible = r("BaseNoi")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.valor = r("ValorIvaNoi")
                                    lstimpNd.Add(impfaNOIVA)

                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.notaDebitoInfonotaDebitoTotalImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "7"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.codigo = r("CodigoExen")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impfaNOIVA.tarifa = r("TarifaExen")
                                    impfaNOIVA.baseImponible = r("BaseExen")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.valor = r("ValorIvaExen")
                                    lstimpNd.Add(impfaNOIVA)

                                End If

                                oNotaDebito.infoNotaDebito.impuestos = lstimpNd.ToArray

                            Catch ex As Exception
                                If _tipoManejoEcua = "A" Then
                                    rsboAppEcua.SetStatusBarMessage("CabEcera nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                                'Utilitario.Util_Log.Escribir_Log("DETALLE nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                Utilitario.Util_Log.Escribir_Log("CabEcera nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                Return Nothing
                            End Try

                        Next
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleND As New Entidades.notaDebitomotivo
                                'itemDetalleND.ValorSpecified = True
                                itemDetalleND.razon = r("Descripcion")
                                itemDetalleND.valor = r("PrecioTotalSinImpuesto")

                                'agrego detalle a la lista
                                listaMotivo.Add(itemDetalleND)
                            Next
                            oNotaDebito.motivos = listaMotivo.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("detalle nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("detalle nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            Return Nothing
                        End Try


                    ElseIf i = 2 Then
                        Try
                            If ds.Tables(2).Rows.Count > 0 Then
                                For Each r As DataRow In ds.Tables(2).Rows
                                    Dim itemDatoAdicionalFac As New Entidades.notaDebitoCampoAdicional
                                    itemDatoAdicionalFac.nombre = r("Concepto")
                                    itemDatoAdicionalFac.Value = r("Descripcion")
                                    listaDatosAdicional.Add(itemDatoAdicionalFac)
                                Next
                                oNotaDebito.infoAdicional = listaDatosAdicional.ToArray
                            Else
                                oNotaDebito.infoAdicional = Nothing
                            End If

                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("informacion adicional nota de debito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("informacion adicional nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            Return Nothing
                        End Try


                    ElseIf i = 3 Then
                        Try
                            If ds.Tables(3).Rows.Count > 0 Then
                                Dim pagos = New Entidades.notaDebitoInfonotaDebitoPagos
                                For Each r As DataRow In ds.Tables(3).Rows
                                    Dim Pago As New Entidades.notaDebitoInfonotaDebitoPagosPago
                                    Pago.formaPago = r("FormaPago")
                                    Pago.total = r("Total")
                                    If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                        Pago.plazo = Nothing
                                    Else
                                        Pago.plazo = r("Plazo")
                                    End If
                                    If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                        Pago.unidadTiempo = Nothing
                                    Else
                                        Pago.unidadTiempo = r("UnidadTiempo")
                                    End If

                                    pagos.pago = Pago
                                Next
                                oNotaDebito.infoNotaDebito.pagos = pagos
                            End If

                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("FORMA PAGO nota de debito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("FORMA PAGO  nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            Return Nothing
                        End Try

                    End If
                Next

            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & "ND-" & oNotaDebito.infoTributaria.estab.ToString() & oNotaDebito.infoTributaria.ptoEmi.ToString() & oNotaDebito.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando ND...", "ManejoDeDocumentos")

                    'Dim ms As New MemoryStream

                    Dim xmlns As New XmlSerializerNamespaces()
                    xmlns.Add("", "")
                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.notaDebito))
                    Dim writer As TextWriter = New StreamWriter(sRuta)

                    x.Serialize(writer, oNotaDebito, xmlns)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado ND..." + sRuta, "ManejoDeDocumentos")

                    'Dim XMLbyte As Byte() = ms.ToArray

                    'Dim base64String As String = ""
                    'base64String = Convert.ToBase64String(XMLbyte)
                End If

                Dim ms As New MemoryStream

                Dim _xmlns As New XmlSerializerNamespaces()
                _xmlns.Add("", "")
                Dim _x As XmlSerializer = New XmlSerializer(GetType(Entidades.notaDebito))
                Dim _writer As TextWriter = New StreamWriter(ms)
                _x.Serialize(_writer, oNotaDebito, _xmlns)
                _writer.Close()


                Dim XMLbyte As Byte() = ms.ToArray

                Dim base64String As String = ""
                base64String = Convert.ToBase64String(XMLbyte)

                Utilitario.Util_Log.Escribir_Log("ND Serializado base 64..." + base64String, "ManejoDeDocumentos")

                Utilitario.Util_Log.Escribir_Log("ND CONSULTADA", "ManejoDeDocumentos")

                Base64 = base64String
                NombreXML = "ND-" + oNotaDebito.infoTributaria.estab + "-" + oNotaDebito.infoTributaria.ptoEmi + "-" + oNotaDebito.infoTributaria.secuencial + ".xml"



            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error NC: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oNotaDebito
        Catch x As ArgumentException
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Nota de Debito7 en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoND, DocEntry, "ArgumentException-Error al Consultar Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            Return Nothing
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la oNotaDebito en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoND, DocEntry, "Error al Consultar Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

    End Function

    Public Function ConsultarRetencion_Ecuanexus(ByVal TipoRE As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object


        'Dim lstimpfact As Object
        'Dim listaDatosAdicional As Object


        'Dim listaRetDocSus As New List(Of Entidades.comprobanteRetencionDocsSustentoDocSustento)
        'Dim listaRetPago As New List(Of Entidades.comprobanteRetencionDocsSustentoDocSustentoPagosPago)
        'lstimpfact = New List(Of Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion)
        Dim listaDatosAdicional As New List(Of Entidades.comprobanteRetencionCampoAdicional)
        'Dim listaRetDocSusRet As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoRetencion)

        Dim oRetencionDocSustento As New Entidades.comprobanteRetencionDocsSustentoDocSustento
        Dim ListoRetencionDocSustentoReem As New List(Of Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalle)

        Dim oRetencion As New Entidades.comprobanteRetencion

        Dim DocSSustento As New Entidades.comprobanteRetencionDocsSustento
        Dim DocSustento As New Entidades.comprobanteRetencionDocsSustentoDocSustento

        Dim ListRetencionDocSustentoImp As New List(Of Entidades.comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento)


        Dim tipoRetencion As Integer
        Try
            Dim SP As String = ""
            If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If TipoRE = "REA" Then
                    SP = "SS_SAP_FE_ObtenerRetencionAnticipo_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "SS_SAP_FE_ObtenerRetencion_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If TipoRE = "REA" Then
                    SP = "SS_SAP_FE_ONE_OBTENERRETENCIONANTICIPO_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "SS_SAP_FE_ONE_OBTENERRETENCION_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If TipoRE = "REA" Then
                    SP = "SS_SAP_FE_HEI_OBTENERRETENCIONANTICIPO_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "SS_SAP_FE_HEI_OBTENERRETENCION_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                If TipoRE = "REA" Then
                    SP = "SS_SAP_FE_SYP_OBTENERRETENCIONANTICIPO_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "SS_SAP_FE_SYP_OBTENERRETENCION_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If TipoRE = "REA" Then
                    SP = "SS_SAP_FE_TM_OBTENERRETENCIONANTICIPO_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "SS_SAP_FE_TM_OBTENERRETENCION_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If TipoRE = "REA" Then
                    SP = "SS_SAP_FE_SS_OBTENERRETENCIONANTICIPO_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "SS_SAP_FE_SS_OBTENERRETENCION_Ecuanexus"
                    If _tipoManejoEcua = "A" Then

                        oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then



                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            Try
                                If r("TipoRetencion").ToString() = "" Then
                                    tipoRetencion = 0
                                Else
                                    tipoRetencion = r("TipoRetencion")
                                End If
                                'Utilitario.Util_Log.Escribir_Log(" (0.- Normal, 1.- Exportadores, 2.- Reembolsos)", "ManejoDeDocumentos")
                                Utilitario.Util_Log.Escribir_Log("Tipo Factura : " & tipoRetencion.ToString(), "ManejoDeDocumentos")
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                                tipoRetencion = 0
                            End Try

                            oRetencion.id = "comprobante"
                            oRetencion.version = "1.0.0"

                            'INICIO INFO TRIBUTARIA
                            Dim oRtInfoTributaria As New Entidades.comprobanteRetencionInfoTributaria
                            ' OFFLINE 14 NOVIEMBRE 2017
                            'If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                            oRtInfoTributaria.claveAcceso = Nothing
                            'Dim clave = GenerarClave(r("FechaEmision").ToString.Replace("/", ""), r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))
                            'oRtInfoTributaria.claveAcceso = IIf(clave = "", Nothing, clave)
                            'Else
                            'nfoTributaria.claveAcceso = r("ClaveAcceso")
                            'End If

                            oRtInfoTributaria.ambiente = r("Ambiente")
                            oRtInfoTributaria.tipoEmision = r("TipoEmision")

                            oRtInfoTributaria.razonSocial = r("RazonSocial")
                            If Not r("NombreComercial") = "" Then
                                oRtInfoTributaria.nombreComercial = r("NombreComercial")
                            End If

                            oRtInfoTributaria.ruc = r("Ruc")

                            oRtInfoTributaria.codDoc = r("CodigoDocumento")
                            oRtInfoTributaria.estab = r("Establecimiento")
                            oRtInfoTributaria.ptoEmi = r("PuntoEmision")
                            oRtInfoTributaria.secuencial = r("SecuencialDocumento")

                            If Not oRtInfoTributaria.secuencial.ToString().Length.Equals("9") Then
                                oRtInfoTributaria.secuencial = oRtInfoTributaria.secuencial.ToString().PadLeft(9, "0")
                            End If
                            Utilitario.Util_Log.Escribir_Log("oRetencion.Secuencial : " & oRtInfoTributaria.secuencial.ToString(), "ManejoDeDocumentos")

                            oRtInfoTributaria.dirMatriz = r("DireccionMatriz")

                            If Not r("AgenteRetencion") = "0" Then
                                oRtInfoTributaria.agenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oRtInfoTributaria.contribuyenteRimpe = r("ContribuyenteRimpe")
                            End If

                            oRetencion.infoTributaria = oRtInfoTributaria

                            'FIN INFOR TRIBUTARIA

                            'INFO RETENCION
                            Dim oRtInfoComprobanteRetencion As New Entidades.comprobanteRetencionInfoCompRetencion

                            oRtInfoComprobanteRetencion.fechaEmision = r("FechaEmision")
                            oRtInfoComprobanteRetencion.dirEstablecimiento = r("DireccionEstablecimiento")

                            If Not r("ContribuyenteEspecial") = "0" Then
                                oRtInfoComprobanteRetencion.contribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oRtInfoComprobanteRetencion.contribuyenteEspecial = Nothing
                            End If



                            oRtInfoComprobanteRetencion.obligadoContabilidad = r("ObligadoContabilidad")

                            oRtInfoComprobanteRetencion.tipoIdentificacionSujetoRetenido = r("TipoIdentificacionSujetoRetenido")
                            oRtInfoComprobanteRetencion.razonSocialSujetoRetenido = r("RazonSocialSujetoRetenido")
                            oRtInfoComprobanteRetencion.identificacionSujetoRetenido = r("IdentificacionSujetoRetenido")
                            oRtInfoComprobanteRetencion.periodoFiscal = r("PeriodoFiscal")

                            If tipoRetencion = 1 Then
                                'oRetencion.Tipo = r("TipoRetencion")
                                If r("TipoSujetoRetenido") <> "" Then
                                    oRtInfoComprobanteRetencion.tipoSujetoRetenido = r("TipoSujetoRetenido")
                                End If

                                oRtInfoComprobanteRetencion.parteRel = r("ParteRel")
                            End If

                            oRetencion.infoCompRetencion = oRtInfoComprobanteRetencion
                        Next
                    ElseIf i = 1 Then
                        If ds.Tables(1).Rows.Count > 0 Then
                            If tipoRetencion = 0 Then

                                Dim ListoRetencionImp As New List(Of Entidades.comprobanteRetencionImpuesto)

                                For Each r As DataRow In ds.Tables(1).Rows

                                    Dim oRetencionImp As New Entidades.comprobanteRetencionImpuesto

                                    oRetencionImp.codigo = r("Codigo")
                                    oRetencionImp.codigoRetencion = r("CodigoRetencion")
                                    oRetencionImp.baseImponible = r("BaseImponible")
                                    oRetencionImp.porcentajeRetener = r("PorcentajeRetener")
                                    oRetencionImp.valorRetenido = r("ValorRetenido")
                                    oRetencionImp.codDocSustento = r("CodDocRetener")
                                    oRetencionImp.numDocSustento = r("NumDocRetener")
                                    oRetencionImp.fechaEmisionDocSustento = r("FechaEmisionDocRetener")

                                    ListoRetencionImp.Add(oRetencionImp)
                                Next
                                oRetencion.impuestos = ListoRetencionImp.ToArray

                            Else
                                oRetencion.impuestos = Nothing
                            End If

                            If tipoRetencion = 1 Then

                                DocSustento.codDocSustento = ds.Tables(1).Rows(0)("CodDocRetener")
                                DocSustento.numDocSustento = ds.Tables(1).Rows(0)("NumDocRetener")
                                DocSustento.fechaEmisionDocSustento = ds.Tables(1).Rows(0)("FechaEmisionDocRetener")
                                DocSustento.codSustento = ds.Tables(1).Rows(0)("CodSustento") ' r("CodSustento")
                                DocSustento.fechaRegistroContable = ds.Tables(1).Rows(0)("FechaRegistroContable")
                                DocSustento.numAutDocSustento = ds.Tables(1).Rows(0)("NumAutDocSustento")
                                DocSustento.pagoLocExt = ds.Tables(1).Rows(0)("PagoLocExt")
                                If DocSustento.pagoLocExt = "02" Then
                                    DocSustento.tipoRegi = ds.Tables(1).Rows(0)("TipoRegi")
                                    DocSustento.paisEfecPago = ds.Tables(1).Rows(0)("PaisEfecPago")
                                    DocSustento.aplicConvDobTrib = ds.Tables(1).Rows(0)("AplicConvDobTrib")
                                    If DocSustento.aplicConvDobTrib = "NO" Then
                                        DocSustento.PagExtSujRetNorLeg = ds.Tables(1).Rows(0)("PagExtSujRetNorLeg")
                                    End If
                                    DocSustento.pagoRegFis = ds.Tables(1).Rows(0)("PagoRegFis")
                                End If

                                If ds.Tables(1).Rows(0)("TotalComprobantesReembolso") <> 0 Then
                                    DocSustento.totalComprobantesReembolso = ds.Tables(1).Rows(0)("TotalComprobantesReembolso")
                                End If
                                If ds.Tables(1).Rows(0)("TotalBaseImponibleReembolso") <> 0 Then
                                    DocSustento.totalBaseImponibleReembolso = ds.Tables(1).Rows(0)("TotalBaseImponibleReembolso")
                                End If
                                If ds.Tables(1).Rows(0)("TotalImpuestoReembolso") <> 0 Then
                                    DocSustento.totalImpuestoReembolso = ds.Tables(1).Rows(0)("TotalImpuestoReembolso")
                                End If

                                DocSustento.totalSinImpuestos = ds.Tables(1).Rows(0)("TotalSinImpuestos")
                                DocSustento.importeTotal = ds.Tables(1).Rows(0)("ImporteTotal")


                                If ds.Tables(1).Rows(0)("Base8") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento
                                    impoRetencionDocSustentoImp.codImpuestoDocSustento = ds.Tables(1).Rows(0)("CodImpDocSus8")
                                    impoRetencionDocSustentoImp.codigoPorcentaje = ds.Tables(1).Rows(0)("CodPor8")
                                    impoRetencionDocSustentoImp.baseImponible = ds.Tables(1).Rows(0)("Base8")
                                    impoRetencionDocSustentoImp.tarifa = ds.Tables(1).Rows(0)("Tarifa8")
                                    impoRetencionDocSustentoImp.valorImpuesto = ds.Tables(1).Rows(0)("ValorImpuesto8")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If ds.Tables(1).Rows(0)("Base12") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento
                                    impoRetencionDocSustentoImp.codImpuestoDocSustento = ds.Tables(1).Rows(0)("CodImpDocSus12")
                                    impoRetencionDocSustentoImp.codigoPorcentaje = ds.Tables(1).Rows(0)("CodPor12")
                                    impoRetencionDocSustentoImp.baseImponible = ds.Tables(1).Rows(0)("Base12")
                                    impoRetencionDocSustentoImp.tarifa = ds.Tables(1).Rows(0)("Tarifa12")
                                    impoRetencionDocSustentoImp.valorImpuesto = ds.Tables(1).Rows(0)("ValorImpuesto12")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If ds.Tables(1).Rows(0)("Base0") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento
                                    impoRetencionDocSustentoImp.codImpuestoDocSustento = ds.Tables(1).Rows(0)("CodImpDocSus0")
                                    impoRetencionDocSustentoImp.codigoPorcentaje = ds.Tables(1).Rows(0)("CodPor0")
                                    impoRetencionDocSustentoImp.baseImponible = ds.Tables(1).Rows(0)("Base0")
                                    impoRetencionDocSustentoImp.tarifa = ds.Tables(1).Rows(0)("Tarifa0")
                                    impoRetencionDocSustentoImp.valorImpuesto = ds.Tables(1).Rows(0)("ValorImpuesto0")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If ds.Tables(1).Rows(0)("BaseNoi") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento
                                    impoRetencionDocSustentoImp.codImpuestoDocSustento = ds.Tables(1).Rows(0)("CodImpDocSusNoi")
                                    impoRetencionDocSustentoImp.codigoPorcentaje = ds.Tables(1).Rows(0)("CodPorNoi")
                                    impoRetencionDocSustentoImp.baseImponible = ds.Tables(1).Rows(0)("BaseNoi")
                                    impoRetencionDocSustentoImp.tarifa = ds.Tables(1).Rows(0)("TarifaNoi")
                                    impoRetencionDocSustentoImp.valorImpuesto = ds.Tables(1).Rows(0)("ValorImpuestoNoi")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If ds.Tables(1).Rows(0)("BaseExen") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoImpuestoDocSustento
                                    impoRetencionDocSustentoImp.codImpuestoDocSustento = ds.Tables(1).Rows(0)("CodImpDocSusExen")
                                    impoRetencionDocSustentoImp.codigoPorcentaje = ds.Tables(1).Rows(0)("CodPorExen")
                                    impoRetencionDocSustentoImp.baseImponible = ds.Tables(1).Rows(0)("BaseExen")
                                    impoRetencionDocSustentoImp.tarifa = ds.Tables(1).Rows(0)("TarifaExen")
                                    impoRetencionDocSustentoImp.valorImpuesto = ds.Tables(1).Rows(0)("ValorImpuestoExen")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                Dim listImpoRetencionDocSustentoRetencion As New List(Of Entidades.comprobanteRetencionDocsSustentoDocSustentoRetencion)

                                For Each r As DataRow In ds.Tables(1).Rows

                                    Dim impoRetencionDocSustentoRetencion As New Entidades.comprobanteRetencionDocsSustentoDocSustentoRetencion
                                    impoRetencionDocSustentoRetencion.codigo = r("Codigo")
                                    impoRetencionDocSustentoRetencion.codigoRetencion = r("CodigoRetencion")
                                    impoRetencionDocSustentoRetencion.baseImponible = r("BaseImponible")
                                    impoRetencionDocSustentoRetencion.porcentajeRetener = r("PorcentajeRetener")
                                    impoRetencionDocSustentoRetencion.valorRetenido = r("ValorRetenido")

                                    listImpoRetencionDocSustentoRetencion.Add(impoRetencionDocSustentoRetencion)

                                Next

                                Dim pagos As New Entidades.comprobanteRetencionDocsSustentoDocSustentoPagos
                                Dim pago As New Entidades.comprobanteRetencionDocsSustentoDocSustentoPagosPago

                                pago.formaPago = ds.Tables(1).Rows(0)("FormaPago")
                                pago.total = ds.Tables(1).Rows(0)("Total")

                                pagos.pago = pago
                                DocSustento.pagos = pagos
                                DocSustento.impuestosDocSustento = ListRetencionDocSustentoImp.ToArray
                                DocSustento.retenciones = listImpoRetencionDocSustentoRetencion.ToArray
                                DocSSustento.docSustento = DocSustento
                                oRetencion.docsSustento = DocSSustento

                            Else
                                oRetencion.docsSustento = Nothing
                            End If

                        End If


                    ElseIf i = 2 Then

                        If ds.Tables(2).Rows.Count > 0 Then

                            For Each r As DataRow In ds.Tables(2).Rows

                                Dim oRetencionDocSustentoReem As New Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalle
                                oRetencionDocSustentoReem.tipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                                oRetencionDocSustentoReem.identificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                                oRetencionDocSustentoReem.codPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                                oRetencionDocSustentoReem.tipoProveedorReembolso = r("TipoProveedorReembolso")
                                oRetencionDocSustentoReem.codDocReembolso = r("CodDocReembolso")
                                oRetencionDocSustentoReem.estabDocReembolso = r("EstabDocReembolso")
                                oRetencionDocSustentoReem.ptoEmiDocReembolso = r("PtoEmiDocReembolso")
                                oRetencionDocSustentoReem.secuencialDocReembolso = r("SecuencialDocReembolso")
                                oRetencionDocSustentoReem.fechaEmisionDocReembolso = CDate(r("FechaEmisionDocReembolso"))
                                oRetencionDocSustentoReem.numeroAutorizacionDocReemb = r("NumeroAutorizacionDocReem")

                                Dim listaoRetencionDocSustentoReemImp As New List(Of Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto)

                                If r("Base12") <> 0 Then
                                    Dim impoRetencionDocSustentoReemImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto
                                    impoRetencionDocSustentoReemImp.codigo = r("CodigoBase12")
                                    impoRetencionDocSustentoReemImp.codigoPorcentaje = r("CodigoPorcentajeBase12")
                                    impoRetencionDocSustentoReemImp.tarifa = r("TarifaBase12")
                                    impoRetencionDocSustentoReemImp.baseImponibleReembolso = r("Base12")
                                    impoRetencionDocSustentoReemImp.impuestoReembolso = r("ImpuestoReembolsoBase12")

                                    listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impoRetencionDocSustentoReemImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto
                                    impoRetencionDocSustentoReemImp.codigo = r("CodigoBase0")
                                    impoRetencionDocSustentoReemImp.codigoPorcentaje = r("CodigoPorcentajeBase0")
                                    impoRetencionDocSustentoReemImp.tarifa = r("TarifaBase0")
                                    impoRetencionDocSustentoReemImp.baseImponibleReembolso = r("Base0")
                                    impoRetencionDocSustentoReemImp.impuestoReembolso = r("ImpuestoReembolsoBase0")

                                    listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impoRetencionDocSustentoReemImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto
                                    impoRetencionDocSustentoReemImp.codigo = r("CodigoBaseNoi")
                                    impoRetencionDocSustentoReemImp.codigoPorcentaje = r("CodigoPorcentajeBaseNoi")
                                    impoRetencionDocSustentoReemImp.tarifa = r("TarifaBaseNoi")
                                    impoRetencionDocSustentoReemImp.baseImponibleReembolso = r("BaseNoi")
                                    impoRetencionDocSustentoReemImp.impuestoReembolso = r("ImpuestoReembolsoBaseNoi")

                                    listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impoRetencionDocSustentoReemImp As New Entidades.comprobanteRetencionDocsSustentoDocSustentoReembolsoDetalleDetalleImpuesto
                                    impoRetencionDocSustentoReemImp.codigo = r("CodigoBaseExen")
                                    impoRetencionDocSustentoReemImp.codigoPorcentaje = r("CodigoPorcentajeBaseExen")
                                    impoRetencionDocSustentoReemImp.tarifa = r("TarifaBaseExen")
                                    impoRetencionDocSustentoReemImp.baseImponibleReembolso = r("BaseExen")
                                    impoRetencionDocSustentoReemImp.impuestoReembolso = r("ImpuestoReembolsoBaseExen")

                                    listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                                End If

                                oRetencionDocSustentoReem.detalleImpuestos = listaoRetencionDocSustentoReemImp.ToArray
                                ListoRetencionDocSustentoReem.Add(oRetencionDocSustentoReem)

                            Next
                            oRetencion.docsSustento.docSustento.reembolsos = ListoRetencionDocSustentoReem.ToArray
                            'Else
                            '    oRetencion.docsSustento.docSustento.reembolsos = Nothing
                        End If

                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As New Entidades.comprobanteRetencionCampoAdicional

                            itemDatoAdicionalFac.nombre = r("Concepto")
                            itemDatoAdicionalFac.Value = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oRetencion.infoAdicional = listaDatosAdicional.ToArray
                    End If
                Next



            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'Dim sRuta As String = sRutaCarpeta & oLiquidacionCompra.infoTributaria.secuencial.ToString() + ".xml"
                Dim sRuta As String = sRutaCarpeta & "RT-" & oRetencion.infoTributaria.estab.ToString() & oRetencion.infoTributaria.ptoEmi.ToString() & oRetencion.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando RT...", "ManejoDeDocumentos")

                    'Dim ms As New MemoryStream

                    Dim xmlns As New XmlSerializerNamespaces()
                    xmlns.Add("", "")
                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.comprobanteRetencion))
                    Dim writer As TextWriter = New StreamWriter(sRuta)

                    x.Serialize(writer, oRetencion, xmlns)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado RT..." + sRuta, "ManejoDeDocumentos")

                    'Dim XMLbyte As Byte() = ms.ToArray

                    'Dim base64String As String = ""
                    'base64String = Convert.ToBase64String(XMLbyte)
                End If

                Dim ms As New MemoryStream

                Dim _xmlns As New XmlSerializerNamespaces()
                _xmlns.Add("", "")
                Dim _x As XmlSerializer = New XmlSerializer(GetType(Entidades.comprobanteRetencion))
                Dim _writer As TextWriter = New StreamWriter(ms)
                _x.Serialize(_writer, oRetencion, _xmlns)
                _writer.Close()


                Dim XMLbyte As Byte() = ms.ToArray

                Dim base64String As String = ""
                base64String = Convert.ToBase64String(XMLbyte)

                Utilitario.Util_Log.Escribir_Log("RT Serializado base 64..." + base64String, "ManejoDeDocumentos")

                Utilitario.Util_Log.Escribir_Log("RT CONSULTADA", "ManejoDeDocumentos")

                Base64 = base64String
                NombreXML = "RT-" + oRetencion.infoTributaria.estab + "-" + oRetencion.infoTributaria.ptoEmi + "-" + oRetencion.infoTributaria.secuencial + ".xml"



            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error NC: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try


            Return oRetencion
        Catch x As ArgumentException
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Retención en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "ArgumentException-Error al Consultar Retención con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la oRetencion en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Error al Consultar Retención con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: RetencionDL/ConsultarRetencion", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

    End Function

    Public Function ConsultarLiquidacionCompra_Ecuanexus(ByVal TipoRE As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oLiquidacionCompra As Entidades.liquidacionCompra = Nothing


        Dim listaDetalleLQ As List(Of Entidades.liquidacionCompraDetalle)
        Dim listaDatoAdicionalDetalleLQCompra As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra)
        Dim listaLiquidacionCompraImp As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto)
        Dim liquidacionCompraPagos As List(Of Entidades.liquidacionCompraInfoLiquidacionCompraPagosPago)
        Dim listaDatosAdicionalesLQ As List(Of Entidades.liquidacionCompraCampoAdicional)

        listaDetalleLQ = New List(Of Entidades.liquidacionCompraDetalle)
        listaDatoAdicionalDetalleLQCompra = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra)
        listaLiquidacionCompraImp = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto)
        liquidacionCompraPagos = New List(Of Entidades.liquidacionCompraInfoLiquidacionCompraPagosPago)

        listaDatosAdicionalesLQ = New List(Of Entidades.liquidacionCompraCampoAdicional)


        Dim listareembolsoLQ As New List(Of Entidades.liquidacionCompraReembolsoDetalle)
        Dim aplicadoDescuentoAdicional As Boolean = False
        Try

            Dim SP As String = ""
            If TipoRE = "LQE" Then
                If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "SS_SAP_FE_ObtenerLiquidacionCompra_Ecuanexus"
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "SS_SAP_FE_ONE_ObtenerLiquidacionCompra_Ecuanexus"
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "SS_SAP_FE_HEI_ObtenerLiquidacionCompra_Ecuanexus"
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "SS_SAP_FE_SYP_ObtenerLiquidacionCompra_Ecuanexus"
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "SS_SAP_FE_TM_ObtenerLiquidacionCompra_Ecuanexus"
                ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "SS_SAP_FE_SS_ObtenerLiquidacionCompra_Ecuanexus"
                End If

                If _tipoManejoEcua = "A" Then

                    oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Consultando Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)
            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")
            Dim tipoLiquidacion As Integer
            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                oLiquidacionCompra = New Entidades.liquidacionCompra

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try
                            For Each r As DataRow In ds.Tables(0).Rows

                                ' MANEJO DE FACTURAS DE EXPORTACION Y REEMBOLSO - 2018-02-18
                                ' Indica que tipo de factura es (0.- Normal, 1.- Exportadores, 2.- Reembolsos)
                                Try
                                    If r("Tipo").ToString() = "" Then
                                        tipoLiquidacion = 0
                                    Else
                                        'oLiquidacionCompra.Tipo = r("Tipo")
                                        tipoLiquidacion = r("Tipo")
                                    End If
                                    Utilitario.Util_Log.Escribir_Log(" (0.- Normal, 1.- Reembolsos)", "ManejoDeDocumentos")
                                    Utilitario.Util_Log.Escribir_Log("Tipo Factura : " & tipoLiquidacion.ToString(), "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                                    tipoLiquidacion = 0
                                End Try

                                oLiquidacionCompra.id = "comprobante"
                                oLiquidacionCompra.version = "1.0.0"

                                ' OFFLINE 14 NOVIEMBRE 2017
                                'FAMC 18/02/2019
                                'INICIO INFO TRIBUTARIA
                                Dim oLQInfoTributaria As New Entidades.liquidacionCompraInfoTributaria

                                'If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oLQInfoTributaria.claveAcceso = Nothing
                                'Dim clave = GenerarClave(r("FechaEmision").ToString.Replace("/", ""), r("CodigoDocumento"), r("RUC"), r("Ambiente"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento"), r("TipoEmision"))
                                'oLQInfoTributaria.claveAcceso = IIf(clave = "", Nothing, clave)
                                'Else
                                'oLQInfoTributaria.claveAcceso = r("ClaveAcceso")
                                'End If

                                oLQInfoTributaria.ambiente = r("Ambiente")
                                oLQInfoTributaria.tipoEmision = r("TipoEmision")
                                oLQInfoTributaria.razonSocial = r("RazonSocial")
                                If Not r("NombreComercial") = "" Then
                                    oLQInfoTributaria.nombreComercial = r("NombreComercial")
                                End If


                                oLQInfoTributaria.ruc = r("RUC")
                                'oLiquidacionCompra.Ruc = "0992737964001"
                                oLQInfoTributaria.codDoc = r("CodigoDocumento")
                                oLQInfoTributaria.estab = r("Establecimiento")
                                oLQInfoTributaria.ptoEmi = r("PuntoEmision")
                                oLQInfoTributaria.secuencial = r("SecuencialDocumento")
                                If Not oLQInfoTributaria.secuencial.ToString().Length.Equals("9") Then
                                    oLQInfoTributaria.secuencial = oLQInfoTributaria.secuencial.ToString().PadLeft(9, "0")
                                End If
                                Utilitario.Util_Log.Escribir_Log("oLiquidacionCompra.Secuencial : " & oLQInfoTributaria.secuencial.ToString(), "ManejoDeDocumentos")
                                oLQInfoTributaria.dirMatriz = r("DireccionMatriz")

                                If Not r("AgenteRetencion") = "0" Then
                                    oLQInfoTributaria.agenteRetencion = r("AgenteRetencion")
                                End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    oLQInfoTributaria.contribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                                End If

                                oLiquidacionCompra.infoTributaria = oLQInfoTributaria
                                'FIN INFO TRIBUTARIA

                                'INFO LIQUIDACION COMPRAS
                                Dim oLQInfoLiquidacion As New Entidades.liquidacionCompraInfoLiquidacionCompra

                                oLQInfoLiquidacion.fechaEmision = r("FechaEmision")
                                oLQInfoLiquidacion.dirEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oLQInfoLiquidacion.contribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oLQInfoLiquidacion.contribuyenteEspecial = Nothing
                                End If

                                oLQInfoLiquidacion.obligadoContabilidad = r("ObligadoContabilidad")
                                oLQInfoLiquidacion.tipoIdentificacionProveedor = r("TipoIdentificacionProveedor")

                                'If Not r("GuiaRemision") = "0" Then
                                '    oLiquidacionCompra.GuiaRemision = r("GuiaRemision")
                                'End If

                                oLQInfoLiquidacion.razonSocialProveedor = r("RazonSocialProveedor")
                                oLQInfoLiquidacion.identificacionProveedor = r("IdentificacionProveedor")

                                Try
                                    If Not r("DirProveedor") = "" Then
                                        oLQInfoLiquidacion.direccionProveedor = r("DirProveedor")
                                    End If
                                Catch ex As Exception
                                End Try

                                oLQInfoLiquidacion.totalSinImpuestos = r("TotalSinImpuesto")
                                oLQInfoLiquidacion.totalDescuento = r("TotalDescuento")

                                If Not r("CodDocReemb") = "" Then
                                    oLQInfoLiquidacion.codDocReembolso = r("CodDocReemb")
                                    oLQInfoLiquidacion.totalComprobantesReembolso = r("TotalComprobantesReembolso")
                                    oLQInfoLiquidacion.totalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                                    oLQInfoLiquidacion.totalImpuestoReembolso = r("TotalImpuestoReembolso")
                                End If

                                oLQInfoLiquidacion.importeTotal = r("ImporteTotal")
                                oLQInfoLiquidacion.moneda = r("Moneda")

                                oLiquidacionCompra.infoLiquidacionCompra = oLQInfoLiquidacion

                                Dim lstimpLQ As New List(Of Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impLQIVA As New Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto

                                    impLQIVA.codigo = r("Codigo8")
                                    impLQIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impLQIVA.tarifa = r("Tarifa8")
                                    impLQIVA.baseImponible = r("Base8")
                                    impLQIVA.valor = r("ValorIva8")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQIVA.descuentoAdicional = r("DescuentoAdicional")
                                        'aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impLQIVA As New Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto

                                    impLQIVA.codigo = r("Codigo12")
                                    impLQIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impLQIVA.tarifa = r("Tarifa12")
                                    impLQIVA.baseImponible = r("Base12")
                                    impLQIVA.valor = r("ValorIva12")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQIVA.descuentoAdicional = r("DescuentoAdicional")
                                        'aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impLQIVA As New Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto

                                    impLQIVA.codigo = r("Codigo13")
                                    impLQIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impLQIVA.tarifa = r("Tarifa13")
                                    impLQIVA.baseImponible = r("Base13")
                                    'impLQIVA.Valor = r("ImpuestoTotal")
                                    impLQIVA.valor = r("ValorIva13")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQIVA.descuentoAdicional = r("DescuentoAdicional")
                                        'aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base0") <> 0 Then

                                    Dim impLQNOIVA As New Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto

                                    impLQNOIVA.codigo = r("Codigo0")
                                    impLQNOIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impLQNOIVA.tarifa = r("Tarifa0")
                                    impLQNOIVA.baseImponible = r("Base0")
                                    impLQNOIVA.valor = r("ValorIva0")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQNOIVA.descuentoAdicional = r("DescuentoAdicional")
                                        'aplicadoDescuentoAdicional = True
                                    End If


                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("BaseNoi") <> 0 Then

                                    Dim impLQNOIVA As New Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto

                                    impLQNOIVA.codigo = r("CodigoNoi")
                                    impLQNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impLQNOIVA.tarifa = r("TarifaNoi")
                                    impLQNOIVA.baseImponible = r("BaseNoi")
                                    impLQNOIVA.valor = r("ValorIvaNoi")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQNOIVA.descuentoAdicional = r("DescuentoAdicional")
                                        'aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then

                                    Dim impLQNOIVA As New Entidades.liquidacionCompraInfoLiquidacionCompraTotalImpuesto

                                    impLQNOIVA.codigo = r("CodigoExen")
                                    impLQNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impLQNOIVA.tarifa = r("TarifaExen")
                                    impLQNOIVA.baseImponible = r("BaseExen")
                                    impLQNOIVA.valor = r("ValorIvaExen")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQNOIVA.descuentoAdicional = r("DescuentoAdicional")
                                        'aplicadoDescuentoAdicional = True
                                    End If


                                    lstimpLQ.Add(impLQNOIVA)
                                End If
                                oLiquidacionCompra.infoLiquidacionCompra.totalConImpuestos = lstimpLQ.ToArray
                            Next
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleLiquidacion As New Entidades.liquidacionCompraDetalle

                                itemDetalleLiquidacion.codigoPrincipal = r("CodigoPrincipal")
                                itemDetalleLiquidacion.codigoAuxiliar = r("CodigoAuxiliar")
                                itemDetalleLiquidacion.descripcion = r("Descripcion")
                                itemDetalleLiquidacion.cantidad = r("Cantidad")
                                itemDetalleLiquidacion.precioUnitario = r("PrecioUnitario")
                                itemDetalleLiquidacion.descuento = r("Descuento")
                                itemDetalleLiquidacion.precioTotalSinImpuesto = r("PrecioTotalSinImpuesto")

                                ''Datos adicionales de cada detalle del item                                     
                                Dim listaDetalleDatoAdicional As New List(Of Entidades.liquidacionCompraDetalleDetAdicional)
                                'Adicional1
                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As New Entidades.liquidacionCompraDetalleDetAdicional
                                    itemDetalleDatoAdicional.nombre = r("ConceptoAdicional1")
                                    itemDetalleDatoAdicional.valor = r("NombreAdicional1")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                                End If

                                'Adicional2
                                If Not r("ConceptoAdicional2") = "0" Then
                                    Dim itemDetalleDatoAdicional2 As New Entidades.liquidacionCompraDetalleDetAdicional
                                    itemDetalleDatoAdicional2.nombre = r("ConceptoAdicional2")
                                    itemDetalleDatoAdicional2.valor = r("NombreAdicional2")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional2)
                                End If

                                'Adicional3
                                If Not r("ConceptoAdicional3") = "0" Then
                                    Dim itemDetalleDatoAdicional3 As New Entidades.liquidacionCompraDetalleDetAdicional
                                    itemDetalleDatoAdicional3.nombre = r("ConceptoAdicional3")
                                    itemDetalleDatoAdicional3.valor = r("NombreAdicional3")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional3)
                                End If

                                itemDetalleLiquidacion.detallesAdicionales = listaDetalleDatoAdicional.ToArray


                                Dim lstimpdetalle As New List(Of Entidades.liquidacionCompraDetalleImpuesto)
                                'Detalle de impuesto de IVA
                                Dim impdetalleIVA As New Entidades.liquidacionCompraDetalleImpuesto

                                impdetalleIVA.codigo = r("Codigo")
                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%

                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA8" Then ' 12%

                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA" Then ' 12%

                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA13" Then ' 12%

                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_NOI" Then ' 12%

                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_EXEN" Then ' 12%

                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                End If

                                'impdetalleIVA.BaseImponible = r("PrecioTotalSinImpuesto") se comento 22/02/2022 porque ya se envia en cada indicador de impuesto
                                impdetalleIVA.valor = r("TotalIva")

                                'agrego impuesto a la lista
                                lstimpdetalle.Add(impdetalleIVA)

                                'agrego lista de impuesto al detalle
                                itemDetalleLiquidacion.impuestos = lstimpdetalle.ToArray

                                'agrego detalle a la lista
                                listaDetalleLQ.Add(itemDetalleLiquidacion)
                            Next
                            oLiquidacionCompra.detalles = listaDetalleLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows

                                Dim reembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalle

                                If Not r("TipoIdentificacionProveedorReembolso") = "" Then
                                    reembolsoLQ.tipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                                End If
                                If Not r("IdentificacionProveedorReembolso") = "" Then
                                    reembolsoLQ.identificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                                End If
                                If Not r("CodPaisPagoProveedorReembolso") = "" Then
                                    reembolsoLQ.codPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                                End If
                                If Not r("TipoProveedorReembolso") = "" Then
                                    reembolsoLQ.tipoProveedorReembolso = r("TipoProveedorReembolso")
                                End If
                                If Not r("CodDocReembolso") = "" Then
                                    reembolsoLQ.codDocReembolso = r("CodDocReembolso")
                                End If
                                If Not r("EstabDocReembolso") = "" Then
                                    reembolsoLQ.estabDocReembolso = r("EstabDocReembolso")
                                End If
                                If Not r("PtoEmiDocReembolso") = "" Then
                                    reembolsoLQ.ptoEmiDocReembolso = r("PtoEmiDocReembolso")
                                End If
                                If Not r("SecuencialDocReembolso") = "" Then
                                    reembolsoLQ.secuencialDocReembolso = r("SecuencialDocReembolso")
                                End If
                                If Not r("FechaEmisionDocReembolso") = "" Then
                                    reembolsoLQ.fechaEmisionDocReembolso = r("FechaEmisionDocReembolso")
                                End If
                                If Not r("NumeroAutorizacionDocReem") = "" Then
                                    reembolsoLQ.numeroautorizacionDocReemb = r("NumeroAutorizacionDocReem")
                                End If

                                Dim listaImpReembolsoLQ As New List(Of Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto)

                                If r("Base8") <> 0 Then

                                    Dim itemImpReembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto

                                    itemImpReembolsoLQ.codigo = r("Codigo8")
                                    itemImpReembolsoLQ.codigoPorcentaje = r("CodigoPorcentaje8")
                                    itemImpReembolsoLQ.tarifa = r("Tarifa8")
                                    itemImpReembolsoLQ.baseImponibleReembolso = r("Base8")
                                    itemImpReembolsoLQ.impuestoReembolso = r("ValorIvaReem8")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base12") <> 0 Then

                                    Dim itemImpReembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto

                                    itemImpReembolsoLQ.codigo = r("Codigo12")
                                    itemImpReembolsoLQ.codigoPorcentaje = r("CodigoPorcentaje12")
                                    itemImpReembolsoLQ.tarifa = r("Tarifa12")
                                    itemImpReembolsoLQ.baseImponibleReembolso = r("Base12")
                                    itemImpReembolsoLQ.impuestoReembolso = r("ValorIvaReem12")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)

                                End If

                                If r("Base13") <> 0 Then

                                    Dim itemImpReembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto

                                    itemImpReembolsoLQ.codigo = r("Codigo13")
                                    itemImpReembolsoLQ.codigoPorcentaje = r("CodigoPorcentaje13")
                                    itemImpReembolsoLQ.tarifa = r("Tarifa13")
                                    itemImpReembolsoLQ.baseImponibleReembolso = r("Base13")
                                    itemImpReembolsoLQ.impuestoReembolso = r("ValorIvaReem13")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)

                                End If

                                If r("Base0") <> 0 Then

                                    Dim itemImpReembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto

                                    itemImpReembolsoLQ.codigo = r("Codigo0")
                                    itemImpReembolsoLQ.codigoPorcentaje = r("CodigoPorcentaje0")
                                    itemImpReembolsoLQ.tarifa = r("Tarifa0")
                                    itemImpReembolsoLQ.baseImponibleReembolso = r("Base0")
                                    itemImpReembolsoLQ.impuestoReembolso = r("ValorIvaReem0")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)

                                End If

                                If r("BaseNoi") <> 0 Then

                                    Dim itemImpReembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto

                                    itemImpReembolsoLQ.codigo = r("CodigoNoi")
                                    itemImpReembolsoLQ.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    itemImpReembolsoLQ.tarifa = r("TarifaNoi")
                                    itemImpReembolsoLQ.baseImponibleReembolso = r("BaseNoi")
                                    itemImpReembolsoLQ.impuestoReembolso = r("ValorIvaReemNoi")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)

                                End If

                                If r("BaseExen") <> 0 Then

                                    Dim itemImpReembolsoLQ As New Entidades.liquidacionCompraReembolsoDetalleDetalleImpuesto

                                    itemImpReembolsoLQ.codigo = r("CodigoExen")
                                    itemImpReembolsoLQ.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    itemImpReembolsoLQ.tarifa = r("TarifaExen")
                                    itemImpReembolsoLQ.baseImponibleReembolso = r("BaseExen")
                                    itemImpReembolsoLQ.impuestoReembolso = r("ValorIvaReemExen")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)

                                End If
                                reembolsoLQ.detalleImpuestos = listaImpReembolsoLQ.ToArray

                                listareembolsoLQ.Add(reembolsoLQ)

                            Next
                            oLiquidacionCompra.reembolsos = listareembolsoLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Reembolso: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Reembolso: " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 3 Then
                        Try
                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim itemDatoAdicionalLQ As New Entidades.liquidacionCompraCampoAdicional

                                itemDatoAdicionalLQ.nombre = r("Concepto")
                                itemDatoAdicionalLQ.Value = r("Descripcion")
                                listaDatosAdicionalesLQ.Add(itemDatoAdicionalLQ)
                            Next
                            oLiquidacionCompra.infoAdicional = listaDatosAdicionalesLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Datos Adicionales: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Datos Adicionales: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 4 Then
                        Try
                            Dim Pagos As New Entidades.liquidacionCompraInfoLiquidacionCompraPagos
                            For Each r As DataRow In ds.Tables(4).Rows

                                Dim Pago As New Entidades.liquidacionCompraInfoLiquidacionCompraPagosPago

                                Pago.formaPago = r("FormaPago")
                                Pago.total = r("Total")
                                If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                    Pago.plazo = Nothing
                                Else
                                    Pago.plazo = r("Plazo")
                                End If
                                If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                    Pago.unidadTiempo = Nothing
                                Else
                                    Pago.unidadTiempo = r("UnidadTiempo")
                                End If
                                'liquidacionCompraPagos.Add(Pago)

                                Pagos.pago = Pago
                            Next
                            oLiquidacionCompra.infoLiquidacionCompra.pagos = Pagos
                        Catch ex As Exception
                            If _tipoManejoEcua = "A" Then
                                rsboAppEcua.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _ErrorEcua = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    End If

                Next

            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oLiquidacionCompra.infoTributaria.estab.ToString() & oLiquidacionCompra.infoTributaria.ptoEmi.ToString() & oLiquidacionCompra.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando LQ...", "ManejoDeDocumentos")

                    'Dim ms As New MemoryStream

                    Dim xmlns As New XmlSerializerNamespaces()
                    xmlns.Add("", "")
                    Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.liquidacionCompra))
                    Dim writer As TextWriter = New StreamWriter(sRuta)

                    x.Serialize(writer, oLiquidacionCompra, xmlns)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado LQ..." + sRuta, "ManejoDeDocumentos")

                    'Dim XMLbyte As Byte() = ms.ToArray

                    'Dim base64String As String = ""
                    'base64String = Convert.ToBase64String(XMLbyte)
                End If

                Dim ms As New MemoryStream

                Dim _xmlns As New XmlSerializerNamespaces()
                _xmlns.Add("", "")
                Dim _x As XmlSerializer = New XmlSerializer(GetType(Entidades.liquidacionCompra))
                Dim _writer As TextWriter = New StreamWriter(ms)
                _x.Serialize(_writer, oLiquidacionCompra, _xmlns)
                _writer.Close()


                Dim XMLbyte As Byte() = ms.ToArray

                Dim base64String As String = ""
                base64String = Convert.ToBase64String(XMLbyte)

                Utilitario.Util_Log.Escribir_Log("LQ Serializado base 64..." + base64String, "ManejoDeDocumentos")

                Utilitario.Util_Log.Escribir_Log("LQ CONSULTADA", "ManejoDeDocumentos")

                Base64 = base64String
                NombreXML = "LQ-" + oLiquidacionCompra.infoTributaria.estab + "-" + oLiquidacionCompra.infoTributaria.ptoEmi + "-" + oLiquidacionCompra.infoTributaria.secuencial + ".xml"



            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error LQE: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oLiquidacionCompra
            Utilitario.Util_Log.Escribir_Log("Liquidacion consultada", "ManejoDeDocumentos")
        Catch x As ArgumentException
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Liquidación de Compra en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "ArgumentException-Error al Consultar Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing

        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al consultar datos de la Liquidación de Compra en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejoEcua = "A" Then

                oFuncionesAddonEcua.GuardaLOG(TipoRE, DocEntry, "Error al Consultar Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        End Try

    End Function

    Public Function GrabaDatosAutorizacion(DocEntry As Integer, TipoDocumento As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim objectType As String = "" 'obtener el objtype del documento para la localizacion de topmanage
        Dim CodDoc As String = "" 'obtener el codigo del documento para la localizacion de topmanage
        Dim SerieDoc As String = ""
        Try
            Dim oDocumento As SAPbobsCOM.Documents = Nothing
            Dim oTransferencia As SAPbobsCOM.StockTransfer = Nothing

            If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Then  ' FACTURA DE CLIENTE
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "01"

            ElseIf TipoDocumento = "NDE" Then ''FACTURA DE ANTICIPO DE CLIENTES
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "05"

            ElseIf TipoDocumento = "FAE" Then ''FACTURA DE ANTICIPO DE CLIENTES
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDownPayments
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "01"

            ElseIf TipoDocumento = "NCE" Then 'NOTA DE CREDITO DE CLIENTES
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "04"

            ElseIf TipoDocumento = "GRE" Then 'GUIA DE REMISION - ENTREGA
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes
                ' objectType = oDocumento.DocObjectCode
                'CodDoc = "06"

            ElseIf TipoDocumento = "TRE" Then 'GUIA DE REMISION - TRANSFERENCIAS
                oTransferencia = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "06"

            ElseIf TipoDocumento = "TLE" Then 'GUIA DE REMISION - SOLICITUD TRANSLADOS
                oTransferencia = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
                oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
                ' objectType = oDocumento.DocObjectCode
                ' CodDoc = "06"

            ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then  'FACTURA DE PROVEEDOR/RETENCION                             
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "07"

            ElseIf TipoDocumento = "RDM" Then
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "07"

            End If

            If TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then
                If oTransferencia.GetByKey(DocEntry) Then
                    'oInvoice.Comments += "Procesada por la Plataforma de Integracion"
                    If _NumAutorizacionEcua <> "" Then

                        If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then

                            oTransferencia.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacionEcua.ToString()
                        ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                            oTransferencia.UserFields.Fields.Item("U_NO_AUTORI").Value = _NumAutorizacionEcua.ToString()
                        ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                            oTransferencia.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _NumAutorizacionEcua.ToString()
                            Try
                                If Not _NumAutorizacionEcua = Nothing Then
                                    oTransferencia.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacionEcua.ToString()
                                End If

                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("U_HBT_IdEnProveedor error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                            Try
                                If Not _NumAutorizacionEcua = Nothing Then
                                    oTransferencia.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAccesoEcua.ToString()
                                End If
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                            If GrabaDatosAutorizacion_HESION_GUIA(TipoDocumento, DocEntry) Then
                                Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")

                            End If
                            If GrabaDatosAutorizacion_HESION_GUIA_TRANSFERENCIAS(TipoDocumento, DocEntry) Then
                                Utilitario.Util_Log.Escribir_Log("Se guardaron los datos de autorizacion en las transferencias incluidas en la guia de remision ", "ManejoDeDocumentos")

                            End If

                        ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                            Try
                                If Not _NumAutorizacionEcua = Nothing Then
                                    oTransferencia.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _NumAutorizacionEcua.ToString()
                                End If

                            Catch ex As Exception

                                Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTO error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                        ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                            Try
                                If Not _NumAutorizacionEcua = Nothing Then
                                    oTransferencia.UserFields.Fields.Item("U_TM_NAUT").Value = _NumAutorizacionEcua.ToString()
                                End If
                                If Not _FechaAutorizacionEcua = Nothing Then
                                    If _EstadoAutorizacionEcua = "2" Then
                                        Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                        oTransferencia.UserFields.Fields.Item("U_TM_DATEA").Value = fechaaut
                                    End If

                                End If

                                If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                                    Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla Control de TM", "ManejoDeDocumentos")

                                End If
                            Catch ex As Exception

                                Utilitario.Util_Log.Escribir_Log("U_TM_NAUT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                        ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                            If Not _NumAutorizacionEcua = Nothing Then
                                oTransferencia.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacionEcua.ToString()
                            End If

                        End If


                        Try
                            If Not _NumAutorizacionEcua = Nothing Then
                                oTransferencia.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacionEcua.ToString()
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("U_NUM_AUTO_FAC error: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try


                        If _tipoManejoEcua = "A" Then
                            Try
                                If _EstadoAutorizacionEcua = "2" And _tipoManejoEcua = "A" Then
                                    rsboAppEcua.SetStatusBarMessage("(SS) N° Autorización: " + _NumAutorizacionEcua.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                End If

                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("(SS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try

                        End If

                    End If
                    '------------
                    Try
                        'oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                        If _FechaAutorizacionEcua <> "" Then
                            If _EstadoAutorizacionEcua = "2" Then
                                Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = CDate(fechaaut)
                            End If

                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    '-------------
                    Try
                        If _ObservacionEcua <> "" Then
                            oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _ObservacionEcua.ToString
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    '--------------
                    Try
                        If _EstadoAutorizacionEcua <> "" Then
                            oTransferencia.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacionEcua = "-1", "0", _EstadoAutorizacionEcua)
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_ESTADO_AUTORIZACIO error : " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    '--------------
                    If Not String.IsNullOrEmpty(_ClaveAccesoEcua) Then
                        Try
                            oTransferencia.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAccesoEcua.ToString()
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("U_CLAVE_ACCESO error: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try
                    End If
                    '---------------------

                    If Not String.IsNullOrEmpty(_FechaAutorizacionEcua.ToString) Then
                        Try

                            If _EstadoAutorizacionEcua = "2" Then
                                Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = fechaaut
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try
                    End If

                    Try
                        resultado = oTransferencia.Update()
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("error al ejecutar la funcion update : " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                End If
            Else
                If oDocumento.GetByKey(DocEntry) Then


                    'oInvoice.Comments += "Procesada por la Plataforma de Integracion"
                    If _NumAutorizacionEcua <> "" Then
                        If Not _NumAutorizacionEcua = Nothing Then
                            oDocumento.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacionEcua.ToString()
                        End If


                        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then

                            If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then


                                oDocumento.UserFields.Fields.Item("U_NUM_AUT_RET").Value = _NumAutorizacionEcua.ToString()


                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then


                                oDocumento.UserFields.Fields.Item("U_NA_RETENCION").Value = _NumAutorizacionEcua.ToString()

                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                Try

                                    oDocumento.UserFields.Fields.Item("U_HBT_AUT_RET").Value = _NumAutorizacionEcua.ToString()


                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_AUT_RET errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                Try

                                    oDocumento.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacionEcua.ToString()


                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_IdEnProveedor errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                                Try

                                    oDocumento.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAccesoEcua.ToString()


                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                Try

                                    oDocumento.UserFields.Fields.Item("U_SYP_NROAUTOC").Value = _NumAutorizacionEcua.ToString()


                                Catch ex As Exception
                                    oDocumento.UserFields.Fields.Item("U_SYP_NROAUTOO").Value = _NumAutorizacionEcua.ToString()
                                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTOO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                                If GrabaDatosAutorizacion_UDORT_TM(TipoDocumento, DocEntry) Then
                                    If _tipoManejoEcua = "A" Then
                                        rsboAppEcua.SetStatusBarMessage("N° Autorización grabada en el UDO TM_LE_RETCH: " + _NumAutorizacionEcua.ToString() + " Tipo Doc: " + TipoDocumento.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If
                                End If
                                If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                                    If _tipoManejoEcua = "A" Then
                                        rsboAppEcua.SetStatusBarMessage("N° Autorización grabada en la tabla Control Doc. Electrónicos: ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If
                                End If
                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                If Not _NumAutorizacionEcua = Nothing Then
                                    oDocumento.UserFields.Fields.Item("U_SS_NumAutRet").Value = _NumAutorizacionEcua.ToString()
                                End If

                            End If

                        Else
                            If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then

                                If _NumAutorizacionEcua <> "" Then
                                    oDocumento.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacionEcua.ToString()
                                End If

                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then

                                If _NumAutorizacionEcua <> "" Then
                                    oDocumento.UserFields.Fields.Item("U_NO_AUTORI").Value = _NumAutorizacionEcua.ToString()
                                End If

                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                Try
                                    If _NumAutorizacionEcua <> "" Then
                                        oDocumento.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _NumAutorizacionEcua.ToString()
                                    End If

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_AUT_FAC error: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                Try
                                    If _NumAutorizacionEcua <> "" Then
                                        oDocumento.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacionEcua.ToString()
                                    End If

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_IdEnProveedor errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                                Try
                                    If _ClaveAccesoEcua <> "" Then
                                        oDocumento.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAccesoEcua.ToString()
                                    End If

                                    Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso : " + _ClaveAccesoEcua.ToString, "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                                If TipoDocumento = "GRE" Then
                                    Utilitario.Util_Log.Escribir_Log("aantes de guardar en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")
                                    Utilitario.Util_Log.Escribir_Log("TipoDocumento" + TipoDocumento.ToString, "ManejoDeDocumentos")
                                    If GrabaDatosAutorizacion_HESION_GUIA(TipoDocumento, DocEntry) Then
                                        Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")
                                    End If
                                End If
                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                Try
                                    If _NumAutorizacionEcua <> "" Then
                                        oDocumento.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _NumAutorizacionEcua.ToString()
                                    End If

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                                Try
                                    If _NumAutorizacionEcua <> "" Then
                                        oDocumento.UserFields.Fields.Item("U_TM_NAUT").Value = _NumAutorizacionEcua.ToString()
                                    End If
                                    If _FechaAutorizacionEcua <> "" Then
                                        If _EstadoAutorizacionEcua = "2" Then
                                            Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                            oDocumento.UserFields.Fields.Item("U_TM_DATEA").Value = fechaaut
                                        End If

                                    End If

                                    If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                                        If _tipoManejoEcua = "A" Then
                                            rsboAppEcua.SetStatusBarMessage("N° Autorización grabada en la tabla Control Doc. Electrónicos: ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        End If
                                    End If

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_TM_NAUT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                If _NumAutorizacionEcua <> "" Then
                                    oDocumento.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacionEcua.ToString()
                                End If


                            End If

                            Try 'SI PARAMETRO ESTA ACTIVO, GUARDA EL NUMERO DE DOCUMENTO QUE SE ENVIÓ AL SRI EN EL CAMPO NUMATCARD
                                If Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = "Y" Then
                                    '_NumeroDeDocumentoSRI = ""
                                    '_NumeroDeDocumentoSRI = oObjeto.Establecimiento + "-" + oObjeto.PuntoEmision + "-" + oObjeto.Secuencial
                                    oDocumento.NumAtCard = _NumeroDeDocumentoSRIEcua
                                    Utilitario.Util_Log.Escribir_Log("NumeroDeDocumentoSRI: " + _NumeroDeDocumentoSRIEcua.ToString(), "ManejoDeDocumentos")
                                End If
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al setear NumeroDeDocumentoSRI: " + ex.Message.ToString(), "ManejoDeDocumentos")
                            End Try

                        End If

                        If _tipoManejoEcua = "A" Then
                            Try
                                If Not _FechaAutorizacionEcua = Nothing Then
                                    rsboAppEcua.SetStatusBarMessage("(SS) N° Autorización: " + _NumAutorizacionEcua.ToString() + " Tipo Doc: " + TipoDocumento.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                End If


                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("(SS) N° Autorización: errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try

                        End If


                    End If
                    Try
                        'oDocumento.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                        If _FechaAutorizacionEcua <> "" Then
                            If _EstadoAutorizacionEcua = "2" Then
                                Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                oDocumento.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = fechaaut
                            End If
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    Try
                        If _FechaAutorizacionEcua <> "" Then
                            If _EstadoAutorizacionEcua = "2" Then
                                Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                oDocumento.UserFields.Fields.Item("U_SYP_FECAUTOC").Value = fechaaut
                            End If

                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_SYP_FECAUTOC DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If _FechaAutorizacionEcua <> "" Then
                            If _EstadoAutorizacionEcua = "2" Then
                                Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                oDocumento.UserFields.Fields.Item("U_TM_DATEA").Value = fechaaut
                            End If

                        End If

                    End If
                    'If Len(_Observacion) > 250 Then
                    '    oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.Substring(1, 153).ToString()
                    'Else
                    '    oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                    'End If
                    Try
                        If _ObservacionEcua <> "" Then
                            oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _ObservacionEcua.ToString
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        If _EstadoAutorizacionEcua <> "" Then
                            oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacionEcua = "-1", "0", _EstadoAutorizacionEcua)
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    If _ClaveAccesoEcua <> "" Then
                        oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAccesoEcua.ToString()
                    End If

                    resultado = oDocumento.Update()
                End If
            End If


            If resultado = 0 Then
                result = True
            Else
                ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompanyEcua.GetLastError(ErrCode, ErrMsg)
                ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If
                If _tipoManejoEcua = "A" Then

                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                Utilitario.Util_Log.Escribir_Log("Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), "ManejoDeDocumentos")
                _ErrorEcua = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _ErrorEcua = ex.Message
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  " & _ErrorEcua.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_LiquidacionCompra(DocEntry As Integer, TipoDocumento As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents = Nothing


            oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
            If oDocumento.GetByKey(DocEntry) Then
                'If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                '    oDocumento.UserFields.Fields.Item("U_LQ_NUM_AUTO").Value = _NumAutorizacion.ToString()
                'ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                '    'oDocumento.UserFields.Fields.Item("U_NA_RETENCION").Value = _NumAutorizacion.ToString()
                'ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                '    'oDocumento.UserFields.Fields.Item("U_HBT_AUT_RET").Value = _NumAutorizacion.ToString()
                'ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

                'End If
                Try
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacionEcua.ToString()
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_NUM_AUTOR errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                'SEIDOR
                Try
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _NumAutorizacionEcua.ToString()
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_LQ_NUM_AUTO").Value = _NumAutorizacionEcua.ToString()
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_NUM_AUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_NO_AUTORI").Value = _NumAutorizacionEcua.ToString()
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_NO_AUTORI errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _NumAutorizacionEcua.ToString()
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_HBT_AUT_FAC errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    If Not String.IsNullOrEmpty(_FechaAutorizacionEcua) Then
                        If _EstadoAutorizacionEcua = "2" Then
                            Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                            oDocumento.UserFields.Fields.Item("U_LQ_FECHA_AUT").Value = fechaaut
                        End If

                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_FECHA_AUT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try

                    oDocumento.UserFields.Fields.Item("U_SYP_FECHAUTOR").Value = Date.Now
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FECHAUTOR DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_OBSERVACION").Value = _ObservacionEcua.ToString
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_OBSERVACION errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value = IIf(_EstadoAutorizacionEcua = "-1", "0", _EstadoAutorizacionEcua)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_ESTADO LQ: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                If Not String.IsNullOrEmpty(_ClaveAccesoEcua) Then

                    oDocumento.UserFields.Fields.Item("U_LQ_CLAVE").Value = _ClaveAccesoEcua.ToString()


                End If
                If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_TM_NAUT").Value = _NumAutorizacionEcua.ToString()
                    End If

                    If Not String.IsNullOrEmpty(_FechaAutorizacionEcua) Then
                        If _EstadoAutorizacionEcua = "2" Then
                            Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                            oDocumento.UserFields.Fields.Item("U_TM_DATEA").Value = fechaaut
                        End If

                    End If

                    If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                        Utilitario.Util_Log.Escribir_Log("Datos de autorizacion de Liquidacion grabados con éxito en la tabla Control Doc. Electronicos", "ManejoDeDocumentos")
                    End If
                End If
                If _Nombre_Proveedor_SAP_BOEcua = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oDocumento.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacionEcua.ToString()
                    End If

                End If
                resultado = oDocumento.Update()
            End If

            If _tipoManejoEcua = "A" Then
                Try
                    If _EstadoAutorizacionEcua = "2" And _tipoManejoEcua = "A" Then
                        rsboAppEcua.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacionEcua.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True
            Else

                rCompanyEcua.GetLastError(ErrCode, ErrMsg)

                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                If _tipoManejoEcua = "A" Then

                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _ErrorEcua = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _ErrorEcua = ex.Message
            If _tipoManejoEcua = "A" Then
                'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

                oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  " & _ErrorEcua.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                'End If
            End If

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry As Integer, TipoDocumento As String, MsgError As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            If oDocumento.GetByKey(DocEntry) Then
                Try
                    If Not String.IsNullOrEmpty(MsgError) Then
                        oDocumento.UserFields.Fields.Item("U_LQ_OBSERVACION").Value = MsgError
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_OBSERVACION error linea 4497: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    resultado = oDocumento.Update()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("error en linea 4503: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True
            Else
                rCompanyEcua.GetLastError(ErrCode, ErrMsg)
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  " & ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                _ErrorEcua = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _ErrorEcua = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_Error(DocEntry As Integer, TipoDocumento As String, MsgError As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            Dim oTransferencia As SAPbobsCOM.StockTransfer

            If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "NDE" Then  ' FACTURA DE CLIENTE
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                'oTipoTabla = "FCE"
            ElseIf TipoDocumento = "FAE" Then ''FACTURA DE ANTICIPO DE CLIENTES
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
            ElseIf TipoDocumento = "NCE" Then 'NOTA DE CREDITO DE CLIENTES
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            ElseIf TipoDocumento = "GRE" Then 'GUIA DE REMISION - ENTREGA
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            ElseIf TipoDocumento = "TRE" Then 'GUIA DE REMISION - TRANSFERENCIAS
                Try
                    oTransferencia = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("funcion guardar datos de autorizacion error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then  'FACTURA DE PROVEEDOR/RETENCION                             
                oDocumento = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            End If

            If TipoDocumento = "TRE" Then
                If oTransferencia.GetByKey(DocEntry) Then
                    Try
                        If Not String.IsNullOrEmpty(MsgError) Then
                            oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = MsgError
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error linea 4482 MD: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        resultado = oTransferencia.Update()
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("error funcion actualizar trnasferencia: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                End If
            Else

                If oDocumento.GetByKey(DocEntry) Then
                    Try
                        If Not String.IsNullOrEmpty(MsgError) Then
                            oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = MsgError
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error linea 4497: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        resultado = oDocumento.Update()
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("error en linea 4503: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                End If
            End If

            If resultado = 0 Then
                result = True
            Else
                rCompanyEcua.GetLastError(ErrCode, ErrMsg)
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  " & ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                _ErrorEcua = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _ErrorEcua = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_HESION_GUIA(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        'Dim DocEntryUdoRet As String = ""
        Dim DocNum As String = ""
        Dim _DocNum As String = ""
        'Dim listaTran As New List(Of Integer)

        If TipoDocumento = "TRE" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")
            If _tipoManejoEcua = "A" Then
                DocNum = oFuncionesB1Ecua.getRSvalue("SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                CODE = oFuncionesB1Ecua.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Transferencias""='Y' and ""U_HBT_NumeroDesde3"" = '" + DocNum.ToString() + "' ", "Code", "")
            Else
                DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Transferencias""='Y' and ""U_HBT_NumeroDesde3"" = '" + DocNum.ToString() + "' ", "Code", "")
            End If
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")
        ElseIf TipoDocumento = "GRE" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum " + DocEntryDoc.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("TipoDocumento " + TipoDocumento.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("_tipoManejo " + _tipoManejoEcua.ToString, "ManejoDeDocumentos")
            If _tipoManejoEcua = "A" Then
                Try
                    DocNum = oFuncionesB1Ecua.getRSvalue("SELECT ""DocNum"" FROM ""ODLN"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    CODE = oFuncionesB1Ecua.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Entregas""='Y' and ""U_HBT_NumeroDesde2"" = '" + DocNum.ToString() + "' ", "Code", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            Else
                Try
                    DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""ODLN"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Entregas""='Y' and ""U_HBT_NumeroDesde2"" = '" + DocNum.ToString() + "' ", "Code", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            End If
            'DocNum = getRSvalueGRHEISON(_DocNum, "DocNum", "")
            Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum query" + DocNum.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code query: " + CODE.ToString, "ManejoDeDocumentos")

        End If


        If CODE = "" Then
            CODE = "0"
        End If
        Try
            If CODE <> "0" Then
                Dim RetVal As Long
                Dim ErrCode As Long
                Dim ErrMsg As String

                Dim ActualizaSecuenc As Boolean = True

                Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
                Dim oUserTable As SAPbobsCOM.UserTable = Nothing
                GC.Collect()
                oUserObjectMD = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                Dim sCmp As SAPbobsCOM.CompanyService
                sCmp = rCompanyEcua.GetCompanyService

                oFuncionesAddonEcua.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @HBT_GUIAREMISION: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompanyEcua.UserTables.Item("HBT_GUIAREMISION")
                oUserTable.GetByKey(CODE)
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If

                If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                    oUserTable.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacionEcua
                End If
                If Not String.IsNullOrEmpty(_ClaveAccesoEcua) Then
                    oUserTable.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAccesoEcua
                End If


                RetVal = oUserTable.Update()
                If RetVal <> 0 Then
                    'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())

                    rCompanyEcua.GetLastError(ErrCode, ErrMsg)

                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                Else
                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_DOC_ELEC: " + CODE.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                End If
                Return True
            Else
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla HBT_GUIAREMISION: " + CODE.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try



        Return result
    End Function

    Public Function GrabaDatosAutorizacion_HESION_GUIA_TRANSFERENCIAS(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim oTransferencia As SAPbobsCOM.StockTransfer = Nothing
        Dim docentry As Integer
        Dim resultado As Integer = -1
        If TipoDocumento = "TRE" Then

            Dim recordset As SAPbobsCOM.Recordset = oFuncionesB1Ecua.getRecordSet("select distinct U_HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OWTR ON T1.U_HBT_NumeroDesde3=OWTR.DocNum where owtr.DocEntry =" + DocEntryDoc.ToString)
            If recordset.RecordCount > 1 Then
                While (recordset.EoF = False)
                    docentry = CInt(recordset.Fields.Item("U_HBT_DocEntry").Value)
                    If DocEntryDoc <> docentry Then
                        oTransferencia = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                        If oTransferencia.GetByKey(docentry) Then
                            If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                                oTransferencia.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacionEcua.ToString()
                            End If
                            If Not String.IsNullOrEmpty(_ClaveAccesoEcua) Then
                                oTransferencia.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAccesoEcua.ToString()
                            End If
                            If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                                oTransferencia.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacionEcua.ToString()
                            End If
                            If Not String.IsNullOrEmpty(_FechaAutorizacionEcua) Then
                                If _EstadoAutorizacionEcua = "2" Then
                                    Dim fechaaut As Date = CDate(_FechaAutorizacionEcua)
                                    oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = fechaaut
                                End If

                            End If
                            'oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                            If Not String.IsNullOrEmpty(_ObservacionEcua) Then
                                oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _ObservacionEcua.ToString
                            End If

                            'oTransferencia.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                            If Not String.IsNullOrEmpty(_ClaveAccesoEcua) Then
                                oTransferencia.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAccesoEcua.ToString()
                            End If
                            Try
                                resultado = oTransferencia.Update()
                            Catch ex As Exception
                                result = False
                                If _tipoManejoEcua = "A" Then
                                    rsboAppEcua.SetStatusBarMessage("Error al actualizar transferencia " + docentry.ToString() + " : " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Transferencia no actualizada: " + docentry.ToString() + " error: " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                End If
                            End Try
                            If resultado = 0 Then
                                If _tipoManejoEcua = "A" Then
                                    rsboAppEcua.SetStatusBarMessage("Transferencia: " + docentry.ToString() + " actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Transferencia actualizada: " + docentry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                End If
                                result = True

                            End If

                        End If
                    End If
                    recordset.MoveNext()
                End While
            End If
        End If
        Return result
    End Function

    Public Function getRSvalueGRHEISON(ByVal query As String, ByVal columnaRet As String, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Utilitario.Util_Log.Escribir_Log("getRSvalue-QUERY: " + query, "FuncionesB1")
            Dim r As SAPbobsCOM.Recordset = getRecordSetGRHEISON(query)
            Utilitario.Util_Log.Escribir_Log("getRSvalue-QUERY: " + query, "FuncionesB1")
            ret = nzStringGRHEISON(r.Fields.Item(columnaRet).Value, , valorNulo)
            ReleaseGRHEISON(r)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRSvalue Catch:" + ex.Message().ToString() + "-QUERY: " + query, "FuncionesB1")
        End Try
        Return ret
    End Function

    Public Function getRecordSetGRHEISON(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRecordSet " + ex.Message.ToString, "FuncionesB1")
        End Try
        Return fRS
    End Function

    Public Function nzStringGRHEISON(ByVal unString As String, Optional ByVal formatoSQL As Boolean = False, Optional ByVal valorSiNulo As String = "") As String
        Try
            If Not IsDBNull(unString) Then
                If formatoSQL Then
                    unString = unString.Replace("'", "' + CHAR(39) + '")
                End If
                'If unString = "0" Then
                '    unString = ""
                'End If
                valorSiNulo = unString
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("nzString Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
        Return valorSiNulo
    End Function

    Public Sub ReleaseGRHEISON(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Release Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
    End Sub

    Public Function GrabaDatosAutorizacion_TablaTM(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        Dim DocEntryUdoRet As String = ""
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    CODE = "SELECT IFNULL(""U_Estable"",'0') AS Establecimiento, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" =79 "
        '    _code = "SELECT IFNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = 79"
        'Else
        '    CODE = "SELECT ISNULL(""U_Estable"",'0') AS Establecimiento, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = 79"
        '    _code = "SELECT ISNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = 79"
        'End If
        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo DocEntry del UDO retencion " + DocEntryUdoRet.ToString, "ManejoDeDocumentos")
            DocEntryUdoRet = oFuncionesB1Ecua.getRSvalue("select T1.""DocEntry"" FROM ""OPCH"" T0 INNER JOIN ""@TM_LE_RETCH"" T1 ON T0.""U_TM_CRNUM""= T1.""DocEntry"" WHERE T0.""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocEntry", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo DocEntry del UDO retencion : " + DocEntryUdoRet.ToString, "ManejoDeDocumentos")
        End If
        'Dim Est As String = oFuncionesB1.getRSvalue(CODE, "Establecimiento")
        'Dim PuntoEmi As String = oFuncionesB1.getRSvalue(_code, "PuntoEmision")
        If TipoDocumento = "LQE" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
            CODE = oFuncionesB1Ecua.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_TipoDoc""='03' and ""U_TM_DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "Code", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
        ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
            CODE = oFuncionesB1Ecua.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_TipoDoc""='07' and ""U_TM_DocEntry"" = '" + DocEntryUdoRet.ToString() + "' ", "Code", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
        Else
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
            CODE = oFuncionesB1Ecua.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "Code", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
        End If

        '_code = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "Code", "")
        'Sql = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + oSerie
        'Dim LQELEC As String = oFuncionesB1.getRSvalue(Sql, "Code", "")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    CODE = "SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_DocEntry"" = " + DocEntryDoc.ToString
        'Else
        '    CODE = "SELECT Code FROM ""@TM_DOC_ELEC"" WHERE U_TM_DocEntry = " + DocEntryDoc.ToString
        'End If
        '_code = oFuncionesB1.getRSvalue(CODE, "Code", "")
        If CODE = "" Then
            CODE = "0"
        End If
        Try
            If CODE <> "0" Then
                Dim RetVal As Long
                Dim ErrCode As Long
                Dim ErrMsg As String

                Dim ActualizaSecuenc As Boolean = True

                Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
                Dim oUserTable As SAPbobsCOM.UserTable = Nothing
                GC.Collect()
                oUserObjectMD = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                Dim sCmp As SAPbobsCOM.CompanyService
                sCmp = rCompanyEcua.GetCompanyService

                oFuncionesAddonEcua.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @TM_DOC_ELEC: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompanyEcua.UserTables.Item("TM_DOC_ELEC")
                oUserTable.GetByKey(CODE)
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If

                If _EstadoAutorizacionEcua.ToString().Equals("2") Or _EstadoAutorizacionEcua.ToString().Equals("AUTORIZADO") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacionEcua.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "A"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_ObservacionEcua.ToString, 254)
                End If
                If _EstadoAutorizacionEcua.ToString().Equals("5") Or _EstadoAutorizacionEcua.ToString().Equals("EN PROCESO SRI") Or _EstadoAutorizacionEcua.ToString().Equals("7") Or _EstadoAutorizacionEcua.ToString().Equals("ERROR EN RECEPCION") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacionEcua.ToString
                    'oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "P"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_ObservacionEcua.ToString, 254)
                End If
                If _EstadoAutorizacionEcua.ToString().Equals("4") Or _EstadoAutorizacionEcua.ToString().Equals("ERROR AL FIRMAR") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacionEcua.ToString
                    'oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "P"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_ObservacionEcua.ToString, 254)
                End If
                If _EstadoAutorizacionEcua.ToString().Equals("3") Or _EstadoAutorizacionEcua.ToString().Equals("NO AUTORIZADA") Or _EstadoAutorizacionEcua.ToString().Equals("6") Or _EstadoAutorizacionEcua.ToString().Equals("DEVUELTA") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacionEcua.ToString
                    'oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "R"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_ObservacionEcua.ToString, 254)
                End If
                RetVal = oUserTable.Update()
                If RetVal <> 0 Then
                    rCompanyEcua.GetLastError(ErrCode, ErrMsg)
                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                Else
                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_DOC_ELEC: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                End If
                Return True
            Else
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla Control Doc. Electrónico", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejoEcua = "A" Then
                rsboAppEcua.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla TM_DOC_ELEC" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If

            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla TM_DOC_ELEC: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try


        Return result
    End Function
    Public Function GrabaDatosAutorizacion_UDORT_TM(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim _code As String = ""
        _code = oFuncionesB1Ecua.getRSvalue("select T1.""DocEntry"" from ""OPCH"" T0 inner join ""@TM_LE_RETCH"" T1 on T0.""U_TM_CRNUM""=T1.""DocEntry"" where T0.""DocEntry""= '" + DocEntryDoc.ToString() + "' ", "DocEntry", "")

        Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion GrabaDatosAutorizacion_UDORT_TM (antes del try)", "ManejoDeDocumentos")
        If _code <> "" Then

            Try
                Dim RetVal As Long
                Dim ErrCode As Long
                Dim ErrMsg As String

                Dim ActualizaSecuenc As Boolean = True
                Dim oGeneralService As SAPbobsCOM.GeneralService
                Dim oGeneralData As SAPbobsCOM.GeneralData
                Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

                Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
                Dim oUserTable As SAPbobsCOM.UserTable = Nothing
                GC.Collect()
                oUserObjectMD = rCompanyEcua.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                Dim sCmp As SAPbobsCOM.CompanyService
                sCmp = rCompanyEcua.GetCompanyService
                Utilitario.Util_Log.Escribir_Log("antes del if ", "ManejoDeDocumentos")
                If oUserObjectMD.GetByKey("TM_LE_RETCH") Then ' PREGUNTO SI ES UN UDO, YA QUE ALGUNOS CLIENTES NO TIENEN REGISTRADO EL UDO
                    'GuardaLOG(Tipotabla, DocEntry, "'EXX_DOCUM_LEG_INTER' es un UDO: ", Transaccion, TipoLog)
                    oGeneralService = sCmp.GetGeneralService("TM_LE_RETCH")
                    Utilitario.Util_Log.Escribir_Log("TM_LE_RETCH oGeneralService", "ManejoDeDocumentos")
                    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    oGeneralParams.SetProperty("Code", _code)
                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "Obteniendo Registro a actualizar en 'TM_LE_RETCH' por el Code: " + _code.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                    Utilitario.Util_Log.Escribir_Log("oGeneralData error", "ManejoDeDocumentos")
                    oGeneralData.SetProperty("U_TM_CASRI", _NumAutorizacionEcua.ToString)
                    oGeneralService.Update(oGeneralData)
                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "# RT AutorizacionSri actualizado en 'TM_LE_RETCH' por el Code: " + _code.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    Return True
                Else

                    oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "Obteniendo Registro a actualizar en 'TM_LE_RETCH' por el Code: " + _code.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    oUserTable = rCompanyEcua.UserTables.Item("TM_LE_RETCH")
                    oUserTable.GetByKey(_code)
                    If Not String.IsNullOrEmpty(_NumAutorizacionEcua) Then
                        oUserTable.UserFields.Fields.Item("U_TM_CASRI").Value = _NumAutorizacionEcua.ToString
                    End If

                    RetVal = oUserTable.Update()
                    If RetVal <> 0 Then ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        rCompanyEcua.GetLastError(ErrCode, ErrMsg)
                        ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "No se actualizaron datos en la tabla TM_LE_RETCH: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Else
                        oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_LE_RETCH: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If
                End If
                Return True
            Catch ex As Exception
                If _tipoManejoEcua = "A" Then
                    rsboAppEcua.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + "Error al actualizar el numero de autorizacion en TM_LE_RETCH" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If

                Return False
            End Try
        Else
            oFuncionesAddonEcua.GuardaLOG(TipoDocumento, DocEntryDoc, "TM_LE_RETCH No se encontro el code: " + _code.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
        End If
    End Function

    Public Function VerificarEstado(ByVal state As String, ByVal statusMsg As String, Optional ByVal code As String = "") As String
        Dim estado As String = ""
        Try
            If state = "10" And Not String.IsNullOrEmpty(statusMsg) Then
                estado = "6"
            ElseIf state = "100" And String.IsNullOrEmpty(statusMsg) Then
                estado = "2"
            ElseIf state = "100" And statusMsg = "SIN ENVIAR AL SRI" Then
                estado = "5"
            ElseIf String.IsNullOrEmpty(state) And String.IsNullOrEmpty(statusMsg) And code = "600" Then
                estado = "1"
            ElseIf state = "" And statusMsg = "" And code = "100" Then
                estado = "0"
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("nzString Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
        Return estado
    End Function

    Public Function Consulta_PDF_XML(clave As String, DocEntry As String, tipoDocumento As String, formato As String) As Boolean
        Try

            'Dim ObjetoRespuesta As Object = Nothing

            Dim Sincro_ruc As String = "", Sincro_Tipo_doc As String = "", Sincro_sec_ERP As String = "", Sincro_Num_Doc As String

            Dim info_company_numdoc() As String = Get_company_numdoc_by_proveedor(_Nombre_Proveedor_SAP_BOEcua, DocEntry, tipoDocumento)
            'RUC compania

            Sincro_ruc = info_company_numdoc(0) 'cero para ruc
            'NUM Doc

            Sincro_Num_Doc = info_company_numdoc(1) 'uno numero doc


            'tipo documento
            Sincro_Tipo_doc = ObtnerIdTipoDocumentoSRI(tipoDocumento)
            'secuencial
            Sincro_sec_ERP = DocEntry

            If Sincro_ruc = "" Or Sincro_Num_Doc = "" Then
                Return Nothing
            End If

            Dim _path = System.IO.Path.GetTempPath
            _path = _path & clave & "." & formato

            Dim _process As New System.Diagnostics.Process
            If System.IO.File.Exists(_path) Then

                _process.StartInfo.FileName = _path
            Else

                'Dim ws As Object
                Dim respuesta_WS As String = ""
                Dim ObjetoRespuesta As New Entidades.ConsultaDocRespuesta

                Dim ConsultarEstadoDoc As New Entidades.ConsultaDocumento
                ConsultarEstadoDoc.NombreWs = Functions.VariablesGlobales._NombreWsEcua
                ConsultarEstadoDoc.clave = Functions.VariablesGlobales._Token
                ConsultarEstadoDoc.ruc = Sincro_ruc
                ConsultarEstadoDoc.docType = Sincro_Tipo_doc
                ConsultarEstadoDoc.docNumber = Sincro_Num_Doc

                ObjetoRespuesta = CoreRest.ConsultaDocumento(ConsultarEstadoDoc, respuesta_WS)
                'ObjetoRespuesta = ws.ConsultarProcesoSincronizadorAX(Sincro_ruc, Sincro_Tipo_doc, Sincro_Num_Doc, Sincro_sec_ERP)

                If Not ObjetoRespuesta Is Nothing Then

                    Dim Archivobyte As Byte() = Nothing

                    Dim _nombreFile = ObjetoRespuesta.authorizationNumber.ToString

                    If formato = "pdf" Then
                        Archivobyte = Convert.FromBase64String(ObjetoRespuesta.pdf)
                        _path = _path & _nombreFile & ".pdf"

                    Else
                        Archivobyte = Convert.FromBase64String(ObjetoRespuesta.xml)
                        _path = _path & _nombreFile & ".xml"
                    End If

                    System.IO.File.WriteAllBytes(_path, Archivobyte)

                    _process.StartInfo.FileName = _path

                End If

            End If



            _process.Start()
            _process.Dispose()

            rsboAppEcua.SetStatusBarMessage(formato.ToUpper.ToString & " Abierto! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            rsboAppEcua.SetStatusBarMessage("Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try

    End Function

#End Region

#Region "Funciones ADO SQL"

    Public Function ConsultaParametro(ByVal Modulo As String, ByVal Tipo As String, ByVal Subtipo As String, ByVal Nombre As String) As String
        Try
            Dim valor As String = ""
            Dim sQueryPrefijo As String = ""
            If rCompanyEcua.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
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

            valor = oFuncionesAddonEcua.getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function EjecutarSP(SP As String, docentry As Integer) As DataSet

        Dim ds As New DataSet

        If rCompanyEcua.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            ds = ObtenerColeccion("CALL " + rCompanyEcua.CompanyDB + "." + SP + " ('" + docentry.ToString() + "')", False)
            Utilitario.Util_Log.Escribir_Log("Query Consulta Factura : " & "CALL " + SP + " ('" + docentry.ToString() + "')", "ManejoDeDocumentos")
        Else
            Try
                Using Cn As SqlConnection = GetSqlConnectionBase()
                    Using cm As New SqlCommand(SP, Cn)
                        Cn.Open()
                        cm.CommandType = CommandType.StoredProcedure
                        cm.Parameters.Add("@DocKey", SqlDbType.Int).Value = docentry
                        Dim da As New SqlDataAdapter
                        da.SelectCommand = cm
                        da.Fill(ds)
                    End Using
                End Using
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al ejecutar el SP:" + SP.ToString + " error: " + ex.Message.ToString + " DocEntry: " + docentry.ToString, "ManejoDeDocumentos")
                rsboAppEcua.SetStatusBarMessage("Ejecutar SP: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return Nothing
            End Try
        End If

        Return ds

    End Function

    Public Function ObtenerColeccion(ByVal Consulta As String, Optional ByVal KeepOpen As Boolean = False) As DataSet

        Dim ds As New DataSet
        Try
            If Consulta = String.Empty Then Return Nothing

            ConectaHANA()

            If CONEXION.State = ConnectionState.Closed Then
                CONEXION.Open()
            End If

            Dim DapTable As New Odbc.OdbcDataAdapter(Consulta, CONEXION)
            DapTable.SelectCommand.CommandTimeout = 0
            DapTable.Fill(ds)

            If Not KeepOpen Then
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If
            End If
            Return ds

        Catch ex As Odbc.OdbcException
            Utilitario.Util_Log.Escribir_Log("ObtenerColeccion: " + ex.Message + " QUERY: " + Consulta.ToString(), "ManejoDeDocumentos")
            Return Nothing
        End Try

    End Function

    Public Function GetSqlConnectionBase() As SqlConnection
        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        Dim cnBaseSAP As New SqlConnection
        Try

            If _tipoManejoEcua <> "A" Then
                BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")
            Else
                BD_User = Functions.VariablesGlobales._vgUserBD
            End If

            If BD_User = "" Then
                rsboAppEcua.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            If _tipoManejoEcua <> "A" Then
                BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")
            Else
                BD_Pass = Functions.VariablesGlobales._vgPassBD
            End If

            If BD_Pass = "" Then
                rsboAppEcua.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            Dim cadena As New SqlConnectionStringBuilder

            If _tipoManejoEcua = "A" Then

                If Not String.IsNullOrEmpty(Functions.VariablesGlobales._vgServerNode) Then
                    cadena.DataSource = Functions.VariablesGlobales._vgServerNode ' "S00SQL" 'rCompany.Server '
                    cadena.InitialCatalog = rCompanyEcua.CompanyDB
                    cadena.UserID = Functions.VariablesGlobales._vgUserBD
                    cadena.Password = Functions.VariablesGlobales._vgPassBD
                Else
                    cadena.DataSource = rCompanyEcua.Server ' "S00SQL" 'rCompany.Server '
                    cadena.InitialCatalog = rCompanyEcua.CompanyDB
                    cadena.UserID = Functions.VariablesGlobales._vgUserBD
                    cadena.Password = Functions.VariablesGlobales._vgPassBD
                End If
            Else
                cadena.DataSource = rCompanyEcua.Server ' "S00SQL" 'rCompany.Server '
                cadena.InitialCatalog = rCompanyEcua.CompanyDB
                cadena.UserID = BD_User
                cadena.Password = BD_Pass
                Utilitario.Util_Log.Escribir_Log("datos conexion sql User: " + BD_User + " Pass: " + BD_Pass + " tipo: " + _tipoManejoEcua, "ManejoDeDocumentos")
            End If

            cnBaseSAP.ConnectionString = cadena.ConnectionString
            Return cnBaseSAP

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("error GetSqlConnectionBase: " + ex.Message.ToString + " User: " + BD_User + " Pass: " + BD_Pass + " tipo: " + _tipoManejoEcua, "ManejoDeDocumentos")

            Return Nothing
        End Try
    End Function

    Public Function ConectaHANA(Optional ByRef mensaje As String = "") As Boolean
        Dim ConexionHana As String = String.Empty

        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        Dim _ServerNode As String = ""
        If _tipoManejoEcua = "S" Then
            _ServerNode = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "ServerNode")
            If String.IsNullOrEmpty(_ServerNode) Then
                _ServerNode = rCompanyEcua.Server
            End If
            BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")
            BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")
            Utilitario.Util_Log.Escribir_Log("_ServerNode: " + _ServerNode.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("BD_User: " + BD_User.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("BD_Pass: " + BD_Pass.ToString(), "ManejoDeDocumentos")
        End If


        Try


            If _tipoManejoEcua <> "A" Then


            Else
                BD_User = Functions.VariablesGlobales._vgUserBD
            End If
            'BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")

            If BD_User = "" Then
                rsboAppEcua.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If


            If _tipoManejoEcua <> "A" Then


            Else
                BD_Pass = Functions.VariablesGlobales._vgPassBD
            End If
            'BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")

            If BD_Pass = "" Then
                rsboAppEcua.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            If rCompanyEcua.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                If (IntPtr.Size = 8) Then
                    ConexionHana = String.Concat(ConexionHana, "Driver={HDBODBC};")
                Else
                    ConexionHana = String.Concat(ConexionHana, "Driver={HDBODBC32};")
                End If
                If _tipoManejoEcua = "A" Then
                    If Not String.IsNullOrEmpty(Functions.VariablesGlobales._vgServerNode) Then
                        ConexionHana = String.Concat(ConexionHana, "ServerNode=", Functions.VariablesGlobales._vgServerNode & ";")
                        ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._vgUserBD, ";")
                        ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._vgPassBD, ";")
                    Else
                        ConexionHana = String.Concat(ConexionHana, "ServerNode=", rCompanyEcua.Server & ";")
                        ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._vgUserBD, ";")
                        ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._vgPassBD, ";")
                    End If
                Else

                    ConexionHana = String.Concat(ConexionHana, "ServerNode=", _ServerNode & ";")
                    ConexionHana = String.Concat(ConexionHana, "UID=", BD_User, ";")
                    ConexionHana = String.Concat(ConexionHana, "PWD=", BD_Pass, ";")


                End If


                'pswBD_HANA

                CONEXION = New Odbc.OdbcConnection(ConexionHana)

                If CONEXION.State = ConnectionState.Closed Then
                    CONEXION.Open()
                End If
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If

                Return True

                'Else
                '    'CONEXION = New Odbc.OdbcConnection("DRIVER={SQL Server Native Client 10.0}; Server= " & serv & "; Database=" & bd & "; Uid=" & userdb & "; Pwd=" & passdb)
                '    CONEXION = New Odbc.OdbcConnection("DRIVER={" + _driversql + "}; Server= " & serv & "; Database=" & bd & "; Uid=" & userdb & "; Pwd=" & passdb)
                '    'CONEXION = New Odbc.OdbcConnection(GetSqlConnectionBaseString())
                '    If CONEXION.State = ConnectionState.Closed Then
                '        CONEXION.Open()
                '    End If
                '    If CONEXION.State = ConnectionState.Open Then
                '        CONEXION.Close()
                '    End If

                '    Return True

            End If
            Return False

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ConexionHana: " + ConexionHana.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Conecta_HANA: " + ex.Message, "ManejoDeDocumentos")
            Return False

        End Try

#Disable Warning BC42353 ' La función 'ConectaHANA' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    End Function

#End Region
    Private Function ValidaClave(ByVal clave As String, ByVal cod_comp As String, ByVal serie As String, ByVal secuencia As String) As String

        If clave.Length = 49 Then

            Dim secufull As String = secuencia
            Dim numero_documento As String = ""

            If Not secufull.Length = 9 Then

                secufull = secufull.PadLeft(9, "0")

            End If

            'construyo el numero del documento
            numero_documento = serie & secufull

            'verifico si el numero del documento y el cod del comprobante existe en la clave de acceso

            If clave.Substring(8, 2) = cod_comp And clave.Substring(24, 15) = numero_documento Then

                Return clave

            Else

                Return ""

            End If


        Else
            Return ""
        End If

    End Function

    Private Function GenerarClave(ByVal fechaemision As String, ByVal CodigoComprobante As String, ByVal ruc As String, ByVal ambiente As String, ByVal serie As String, ByVal secuencial As String, ByVal TipoEmision As String) As String

        'Dim sGUID As String
        'sGUID = System.Guid.NewGuid.ToString()
        Dim numeroaleatorio As String = ""
        Dim randon As New Random

        Do
            numeroaleatorio = randon.Next(0, 99999999).ToString
        Loop While numeroaleatorio.Length < 8

        Dim clave48d As String = fechaemision & CodigoComprobante & ruc & ambiente & serie & secuencial & numeroaleatorio.ToString & TipoEmision

        If clave48d.Length = 48 Then

            Dim clave1 = clave48d.ToCharArray()
            Dim suma As Integer = 0, factor As Integer = 7

            For Each item In clave1
                suma = suma + Convert.ToInt32(item.ToString()) * factor
                factor = factor - 1
                If factor = 1 Then factor = 7
            Next

            Dim digitoverificador = (suma Mod 11)
            digitoverificador = 11 - digitoverificador

            If digitoverificador = 11 Then
                digitoverificador = 0
            ElseIf digitoverificador = 10 Then
                digitoverificador = 1
            End If

            Return clave48d & digitoverificador.ToString()

        Else
            Return ""
        End If

    End Function

    Private Function ImprmirDOcAut(bytePDF As String) As Boolean

        Try

            'Dim ws As Object

            'ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

            Dim Stream = New MemoryStream(bytePDF)
            Using doc As New PdfDocument
                'doc.LoadFromXPS(arrb)
                doc.LoadFromStream(Stream)
                doc.PrintSettings.PrintController = New StandardPrintController
                doc.Print()
                Utilitario.Util_Log.Escribir_Log("Impresion Realizada: " + doc.ToString(), "ImpresionAutomatica")
            End Using


            Return True
        Catch ex As Exception
            mensajeDocAutEcua = ex.Message.ToString
            Return False
        End Try


    End Function

    Public Sub SetProtocolosdeSeguridad()



        'PARA TLS 1.2
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999



        ''PARA HTTPS



        'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)



    End Sub

End Class

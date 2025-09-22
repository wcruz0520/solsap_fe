Imports Functions

Public Class ProcesoEmision
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    Private _tipoManejo As String
    Private oFuncionesAddon As Functions.FuncionesAddon
    Dim DatabaseQueryManager_ As DatabaseQueryManager
    Dim ApiRequestManager_ As ApiRequestManager
    Dim ApiAuthManager_ As ApiAuthManager
    Dim FacturaManager_ As FacturaManager
    Dim NotaCreditoManager_ As NotaCreditoManager
    Dim NotaDebitoManager_ As NotaDebitoManager
    Dim LiquidacionManager_ As LiquidacionManager
    Dim RetencionManager_ As RetencionManager
    Dim DesencriptarQuery_ As DesencriptarQuery
    Dim FuncionesProcesoEmision_ As FuncionesProcesoEmision


    Public Sub New(company As SAPbobsCOM.Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones

        DatabaseQueryManager_ = New Negocio.DatabaseQueryManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon)
        ApiAuthManager_ = New Negocio.ApiAuthManager(_tipoManejo, rsboApp, rCompany)
        ApiRequestManager_ = New Negocio.ApiRequestManager(_tipoManejo, rsboApp, ApiAuthManager_)
        DesencriptarQuery_ = New Negocio.DesencriptarQuery(rCompany, rsboApp, _tipoManejo, oFuncionesAddon)
        FacturaManager_ = New Negocio.FacturaManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon, DatabaseQueryManager_, DesencriptarQuery_)
        NotaCreditoManager_ = New Negocio.NotaCreditoManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon, DatabaseQueryManager_, DesencriptarQuery_)
        NotaDebitoManager_ = New Negocio.NotaDebitoManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon, DatabaseQueryManager_, DesencriptarQuery_)
        LiquidacionManager_ = New Negocio.LiquidacionManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon, DatabaseQueryManager_, DesencriptarQuery_)
        RetencionManager_ = New Negocio.RetencionManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon, DatabaseQueryManager_, DesencriptarQuery_)
        FuncionesProcesoEmision_ = New Negocio.FuncionesProcesoEmision(rCompany, rsboApp, _tipoManejo, oFuncionesAddon)
    End Sub
    Public Function ProcesaEnvioDocumento(DocEntry As Integer,
                                          TipoDocumento As String,
                                          ByRef _Error As String,
                                          ByRef _MsgError As String,
                                          ByRef _EstadoAutorizacion As String,
                                          ByRef _ClaveAcceso As String,
                                          ByRef _NumAutorizacion As String,
                                          ByRef _Observacion As String,
                                          ByRef _FechaAutorizacion As DateTime,
                                          ByRef _NumeroDeDocumentoSRI As String,
                                          ByRef mensaje As String,
                                          ByRef _errorMensajeWSEnvío As String,
                                          ByRef oObjeto As Object,
                                          ByRef _campoNulo As String,
                                          ByVal _Nombre_Proveedor_SAP_BO As String,
                                          Optional ByVal sincronizado As Boolean = False) As String

        Try
            Dim result As Boolean = False
            Dim objetoRespuesta As Object = Nothing
            Dim TipoWS As String = "LOCAL"

            Dim BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo

            Dim sSQL As String = ""

            If _tipoManejo = "S" Then
                TipoWS = DatabaseQueryManager_.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
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
                'objetoRespuesta = SincronizarDocumentoEDOC(TipoDocumento, DocEntry, TipoWS)

                'SEGUIR VALIDANDO CUANDO SE CONSULTA O REENVIA DESDE LA PANTALLA DE DOCUMENTOS ENVIADOS
                Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta

                Dim u_clave_acceso As String = ""
                Dim oRs_ As SAPbobsCOM.Recordset = CType(rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" Then

                    Select Case TipoDocumento
                        Case "FAE"
                            oRs_.DoQuery($"SELECT ""U_CLAVE_ACCESO"" FROM ""ODPI"" WHERE ""DocEntry"" = {DocEntry}")
                        Case Else
                            oRs_.DoQuery($"SELECT ""U_CLAVE_ACCESO"" FROM ""OINV"" WHERE ""DocEntry"" = {DocEntry}")
                    End Select

                    If Not oRs_.EoF Then
                        u_clave_acceso = oRs_.Fields.Item("U_CLAVE_ACCESO").Value.ToString().Trim()
                    End If
                    RequestconsEst.claveAcceso = u_clave_acceso
                    objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                ElseIf TipoDocumento = "NCE" Then

                    oRs_.DoQuery($"SELECT ""U_CLAVE_ACCESO"" FROM ""ORIN"" WHERE ""DocEntry"" = {DocEntry}")

                    If Not oRs_.EoF Then
                        u_clave_acceso = oRs_.Fields.Item("U_CLAVE_ACCESO").Value.ToString().Trim()
                    End If
                    RequestconsEst.claveAcceso = u_clave_acceso
                    objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
                    oRs_.DoQuery($"SELECT ""U_CLAVE_ACCESO"" FROM ""OPCH"" WHERE ""DocEntry"" = {DocEntry}")

                    If Not oRs_.EoF Then
                        u_clave_acceso = oRs_.Fields.Item("U_CLAVE_ACCESO").Value.ToString().Trim()
                    End If
                    RequestconsEst.claveAcceso = u_clave_acceso
                    objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)

                ElseIf TipoDocumento = "LQE" Then
                    oRs_.DoQuery($"SELECT ""U_LQ_CLAVE"" FROM ""OPCH"" WHERE ""DocEntry"" = {DocEntry}")

                    If Not oRs_.EoF Then
                        u_clave_acceso = oRs_.Fields.Item("U_LQ_CLAVE").Value.ToString().Trim()
                    End If
                    RequestconsEst.claveAcceso = u_clave_acceso
                    objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)

                ElseIf TipoDocumento = "NDE" Then

                    oRs_.DoQuery($"SELECT ""U_CLAVE_ACCESO"" FROM ""OINV"" WHERE ""DocEntry"" = {DocEntry} AND ""DocSubType"" = 'DN'")

                    If Not oRs_.EoF Then
                        u_clave_acceso = oRs_.Fields.Item("U_CLAVE_ACCESO").Value.ToString().Trim()
                    End If
                    RequestconsEst.claveAcceso = u_clave_acceso
                    objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)

                End If

                If Not objetoRespuesta Is Nothing Then

                    'lbestadoSRI.Text = resp.Estado
                    _EstadoAutorizacion = objetoRespuesta.codigo
                    _ClaveAcceso = objetoRespuesta.claveAcceso

                    If _tipoManejo = "A" Then
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "SS - Sincronización Respuesta del SRI: " + _EstadoAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If

                    Dim mensajeError As String = ""
                    If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                        Try
                            _NumAutorizacion = objetoRespuesta.claveAcceso.ToString()
                            _Observacion = "Documento AUTORIZADO AUTORIZACION # " & _NumAutorizacion
                            Dim fechaTexto As String = objetoRespuesta.fechaAutorizacion.ToString
                            Dim fechaauto_convert = DateTime.Parse(fechaTexto, Nothing, Globalization.DateTimeStyles.RoundtripKind)
                            _FechaAutorizacion = fechaauto_convert.Date 'FormatearFechaSAP(fechaTexto)
                        Catch ex As Exception

                        End Try
                    Else
                        _NumAutorizacion = "0000000000"
                        _Observacion = "SS_SINCRO: " & objetoRespuesta.mensaje
                        mensajeError = _Observacion.ToString
                    End If

                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("SS_SINCRO_Grabando respuesta de SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "SS_SINCRO_Grabando Respuesta del SRI en Documento - " + TipoDocumento + " - " + DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If

                    _Observacion = String.Format("SINCRO Estado:{0} - # AUTORIZACION {1} - Mensaje - {2}", _EstadoAutorizacion.ToString, _NumAutorizacion.ToString, mensajeError)

                    ' Mando a Grabar a SAP
                    If TipoDocumento = "LQE" Then
                        result = FuncionesProcesoEmision_.GrabaDatosAutorizacion_LQ(DocEntry, TipoDocumento, _Nombre_Proveedor_SAP_BO, _NumAutorizacion, _FechaAutorizacion, _NumeroDeDocumentoSRI, _Observacion, _EstadoAutorizacion, _ClaveAcceso, _Error)
                    ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                    ElseIf TipoDocumento = "SSGR" Then

                    Else
                        result = FuncionesProcesoEmision_.GrabaDatosAutorizacion(DocEntry, TipoDocumento, _Nombre_Proveedor_SAP_BO, _NumAutorizacion, _FechaAutorizacion, _NumeroDeDocumentoSRI, _Observacion, _EstadoAutorizacion, _ClaveAcceso, _Error)
                    End If
                    If result Then
                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("Proceso terminado con exito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        End If

                    Else
                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("Ocurrio un Error al Guardar los datos de Autorización..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        End If

                    End If

                Else

                    'No se pudo sincronizar
                    _Observacion = "GS_SINCRO - ProcesaEnvioDocumento - ObjetoRespuesta vacio  " + DocEntry.ToString() + " - " + mensaje.ToString()
                    _Error = "GS_SINCRO-No se recibio respuesta inmediata del servicio : " + _errorMensajeWSEnvío.ToString()
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("GS_SINCRO-No se recibio respuesta, Presione nuevamente el boton de Consultar Autorizacion.", SAPbouiCOM.BoMessageTime.bmt_Short, True)

                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "GS_SINCRO-No se recibió respuesta de los Web Services", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                    End If

                    Try
                        If TipoDocumento = "LQE" Then

                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                        ElseIf TipoDocumento = "SSGR" Then

                        Else
                            FuncionesProcesoEmision_.GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _MsgError, _Error)
                        End If
                    Catch ex As Exception
                    End Try

                End If
            Else

                If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Seteando informacion a enviar..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" Then
                    oObjeto = FacturaManager_.ConsultarFactura(TipoDocumento, DocEntry, _Error, _campoNulo)
                ElseIf TipoDocumento = "NCE" Then
                    oObjeto = NotaCreditoManager_.ConsultarNotaCredito(TipoDocumento, DocEntry, _Error, _campoNulo)
                ElseIf TipoDocumento = "NDE" Then
                    oObjeto = NotaDebitoManager_.ConsultarNotaDebito(DocEntry, _Error)
                ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
                    oObjeto = RetencionManager_.ConsultarRetencion(DocEntry, TipoDocumento)
                ElseIf TipoDocumento = "LQE" Then
                    oObjeto = LiquidacionManager_.ConsultarLiquidacion(DocEntry, _Error)
                End If

                If Not oObjeto Is Nothing Then

                    Try
                        If Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = "Y" Then
                            _NumeroDeDocumentoSRI = ""
                            _NumeroDeDocumentoSRI = oObjeto.infoTributaria.estab + "-" + oObjeto.infoTributaria.ptoEmi + "-" + oObjeto.infoTributaria.secuencial
                            Utilitario.Util_Log.Escribir_Log("NumeroDeDocumentoSRI: " + _NumeroDeDocumentoSRI.ToString(), "ManejoDeDocumentos")
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error al setear NumeroDeDocumentoSRI: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    Utilitario.Util_Log.Escribir_Log("Enviando documento al SRI, por favor espere..!!", "ManejoDeDocumentos")

                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("Enviando documento al SRI, por favor espere..!!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Envíando Documento al SRI", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
                    End If

                    Utilitario.Util_Log.Escribir_Log($"Enviando documento al SRI, TipoDocumento: {TipoDocumento} DocEntry: {DocEntry} TipoWs: {TipoWS}", "ManejoDeDocumentos")

                    Dim respuesta_WS As String = ""
                    If Functions.VariablesGlobales._ActApiSS = "Y" AndAlso (TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE") Then
                        Dim oDocumento As SAPbobsCOM.Documents = Nothing
                        Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta
                        oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        Select Case oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value
                            Case "0", "3", "4", "6"
                                objetoRespuesta = ApiRequestManager_.EnviarFacturaSolsap(DirectCast(oObjeto, Entidades.RequestFactura))
                            Case "1", "5", "7"
                                RequestconsEst.claveAcceso = oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value
                                objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                        End Select

                    ElseIf Functions.VariablesGlobales._ActApiSS = "Y" AndAlso TipoDocumento = "NCE" Then
                        Dim oDocumento As SAPbobsCOM.Documents = Nothing
                        Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta
                        oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                        Select Case oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value
                            Case "0", "3", "4", "6"
                                objetoRespuesta = ApiRequestManager_.EnviarNCSolsap(DirectCast(oObjeto, Entidades.RequestNotaCredito))
                            Case "1", "5", "7"
                                RequestconsEst.claveAcceso = oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value
                                objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                        End Select

                    ElseIf Functions.VariablesGlobales._ActApiSS = "Y" AndAlso TipoDocumento = "NDE" Then
                        Dim oDocumento As SAPbobsCOM.Documents = Nothing
                        Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta
                        oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        Select Case oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value
                            Case "0", "3", "4", "6"
                                objetoRespuesta = ApiRequestManager_.EnviarNDSolsap(DirectCast(oObjeto, Entidades.RequestNotaDebito))
                            Case "1", "5", "7"
                                RequestconsEst.claveAcceso = oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value
                                objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                        End Select

                    ElseIf Functions.VariablesGlobales._ActApiSS = "Y" AndAlso TipoDocumento = "LQE" Then
                        Dim oDocumento As SAPbobsCOM.Documents = Nothing
                        Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta
                        oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        Select Case oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value
                            Case "0", "3", "4", "6"
                                objetoRespuesta = ApiRequestManager_.EnviarLiquidacionSolsap(DirectCast(oObjeto, Entidades.RequestLiquidacion))
                            Case "1", "5", "7"
                                RequestconsEst.claveAcceso = oDocumento.UserFields.Fields.Item("U_LQ_CLAVE").Value
                                objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                        End Select

                    ElseIf Functions.VariablesGlobales._ActApiSS = "Y" And (TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER") Then
                        Dim oDocumento As SAPbobsCOM.Documents = Nothing
                        Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta
                        oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        Select Case oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value
                            Case "0", "3", "4", "6"
                                objetoRespuesta = ApiRequestManager_.EnviarRetencionSolsap(DirectCast(oObjeto, Entidades.RequestRetencion))
                            Case "1", "5", "7"
                                RequestconsEst.claveAcceso = oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value
                                objetoRespuesta = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                        End Select
                    End If

                    If Not objetoRespuesta Is Nothing Then
                        Dim mensajesSRI As String = ""

                        If Functions.VariablesGlobales._ActApiSS = "Y" AndAlso TypeOf objetoRespuesta Is Entidades.ResponseDocuments Then
                            Dim respDoc As Entidades.ResponseDocuments = CType(objetoRespuesta, Entidades.ResponseDocuments)
                            _EstadoAutorizacion = respDoc.codigo
                            _Observacion = respDoc.mensaje
                            _ClaveAcceso = respDoc.claveAcceso
                            _NumAutorizacion = respDoc.claveAcceso
                        Else

                        End If

                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Respuesta del SRI: " + _EstadoAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                        If _tipoManejo = "A" Then
                            ' Seteo el Error recibido del API SOLSAP
                            rsboApp.SetStatusBarMessage("Recibiendo respuesta..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            'Dim respuestaaa = objetoRespuesta.Autorizaciones(0).Mensajes(0).mensaje1().ToString & "- " & objetoRespuesta.Autorizaciones(0).Mensajes(0).informacionAdicional().ToString
                        End If

                        If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" And (TipoWS = "NUBE_4_1" And Functions.VariablesGlobales._ActApiSS = "Y") Then
                            _Observacion = FuncionesProcesoEmision_.recorreError_Solsap(CType(objetoRespuesta, Entidades.ResponseDocuments), DocEntry.ToString())
                        ElseIf TipoDocumento = "NCE" And (TipoWS = "NUBE_4_1" And Functions.VariablesGlobales._ActApiSS = "Y") Then
                            _Observacion = FuncionesProcesoEmision_.recorreError_Solsap(CType(objetoRespuesta, Entidades.ResponseDocuments), DocEntry.ToString())
                        ElseIf TipoDocumento = "NDE" And (TipoWS = "NUBE_4_1" And Functions.VariablesGlobales._ActApiSS = "Y") Then
                            _Observacion = FuncionesProcesoEmision_.recorreError_Solsap(CType(objetoRespuesta, Entidades.ResponseDocuments), DocEntry.ToString())
                        ElseIf TipoDocumento = "LQE" And (TipoWS = "NUBE_4_1" And Functions.VariablesGlobales._ActApiSS = "Y") Then
                            _Observacion = FuncionesProcesoEmision_.recorreError_Solsap(CType(objetoRespuesta, Entidades.ResponseDocuments), DocEntry.ToString())
                        ElseIf (TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER") And (TipoWS = "NUBE_4_1" And Functions.VariablesGlobales._ActApiSS = "Y") Then
                            _Observacion = FuncionesProcesoEmision_.recorreError_Solsap(CType(objetoRespuesta, Entidades.ResponseDocuments), DocEntry.ToString())
                        End If

                        If _tipoManejo = "A" Then
                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Observación del SRI: " + _Observacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If
                        'oBackgroundWorker.ReportProgress(70)

                        If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                            Dim RequestconsEst As Entidades.RequestConsulta = New Entidades.RequestConsulta
                            Dim ResponseconsEst As Entidades.ResponseConsulta = New Entidades.ResponseConsulta
                            RequestconsEst.claveAcceso = _ClaveAcceso
                            ResponseconsEst = ApiRequestManager_.ConsultarDoc(RequestconsEst)
                            _NumAutorizacion = ResponseconsEst.claveAcceso

                            Try
                                'Dim fechaTexto As String = ResponseconsEst.fechaAutorizacion
                                Dim fecha_iso As String = ResponseconsEst.fechaAutorizacion.ToString
                                Dim fechaauto_convert = DateTime.Parse(fecha_iso, Nothing, Globalization.DateTimeStyles.RoundtripKind)
                                '_FechaAutorizacion = Nothing 'FormatearFechaSAP(fechaTexto)
                                _FechaAutorizacion = fechaauto_convert.Date
                            Catch ex As Exception

                            End Try
                        Else
                            _NumAutorizacion = "0000000000"
                            Try
                                mensajesSRI = objetoRespuesta.mensaje
                            Catch ex As Exception
                                mensajesSRI = " No se recibio la descripcion del Error "
                            End Try
                        End If

                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("Grabando respuesta de SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Respuesta SRI en Documento - " + TipoDocumento + " - DocEntry " + DocEntry.ToString() + " - # de Autorización - " + _NumAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If

                        _Observacion = String.Format("Estado:{0} - # AUTORIZACION {1} - RespuestaSRI - {2} - Error - {3} ", _EstadoAutorizacion.ToString, _NumAutorizacion.ToString, mensajesSRI, _Observacion.ToString)

                        ' Mando a Grabar a SAP
                        If TipoDocumento = "LQE" Then
                            result = FuncionesProcesoEmision_.GrabaDatosAutorizacion_LQ(DocEntry, TipoDocumento, _Nombre_Proveedor_SAP_BO, _NumAutorizacion, _FechaAutorizacion, _NumeroDeDocumentoSRI, _Observacion, _EstadoAutorizacion, _ClaveAcceso, _Error)
                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                        ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then

                        ElseIf TipoDocumento = "SSGR" Then

                        Else
                            result = FuncionesProcesoEmision_.GrabaDatosAutorizacion(DocEntry, TipoDocumento, _Nombre_Proveedor_SAP_BO, _NumAutorizacion, _FechaAutorizacion, _NumeroDeDocumentoSRI, _Observacion, _EstadoAutorizacion, _ClaveAcceso, _Error)
                        End If

                        If result Then
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Proceso terminado con exito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            End If

                        Else
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Ocurrio un Error al Guardar los datos de Autorización..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                        End If
                    Else
                        ' controlo error si no pude consumir el servicio del SRI
                        ' NO SE RECIBIO RESPUESTA DEL WEB SERVICE DE EDOC - ENVÍA FACTURA

                        _Observacion = "No se ha recibido respuesta del documento " + DocEntry.ToString() + " - Resp WS :" + respuesta_WS
                        _Error = respuesta_WS
                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("No se recibio respuesta inmediata del SRI, el documento será procesado nuevamente en 2 minutos, o use la opcion REENVIAR SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "No se recibió respuesta de los Web Services", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If

                        Try
                            If TipoDocumento = "LQE" Then
                                FuncionesProcesoEmision_.GrabaDatosAutorizacion_Error_LQ(DocEntry, TipoDocumento, _MsgError, _Error)
                            ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                            ElseIf TipoDocumento = "SSGR" Then

                            Else
                                FuncionesProcesoEmision_.GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _MsgError, _Error)
                            End If

                        Catch ex As Exception
                        End Try

                    End If
                Else
                    ' Controlo Error si no se seteo la factura con los datos de base 
                    _Observacion = "Ocurrio un error al Consultar los datos de la Factura: " & DocEntry.ToString() & " " & _campoNulo
                    _Error = _Observacion
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la factura en la Base, DocEntry:  " & DocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Ocurrio un error al consultar datos de la factura en la Base, DocEntry: " & DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If

                    Try
                        If TipoDocumento = "LQE" Then
                            FuncionesProcesoEmision_.GrabaDatosAutorizacion_Error_LQ(DocEntry, TipoDocumento, _MsgError, _Error)
                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                        ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then

                        ElseIf TipoDocumento = "SSGR" Then

                        Else
                            FuncionesProcesoEmision_.GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _MsgError, _Error)
                        End If
                    Catch ex As Exception
                    End Try

                End If

            End If

            Return _Observacion

        Catch ex As Exception
            _Error = ex.Message
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error:  " & ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return _Error + _errorMensajeWSEnvío
        End Try
    End Function
End Class

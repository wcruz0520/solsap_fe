Imports Functions

Public Class FuncionesProcesoEmision
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    Private _tipoManejo As String
    Private oFuncionesAddon As Functions.FuncionesAddon
    Private oFuncionesB1 As FuncionesB1

    Public Sub New(company As SAPbobsCOM.Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
    End Sub

    Public Function GrabaDatosAutorizacion_Error_LQ(DocEntry As Integer, TipoDocumento As String, ByVal MsgError As String, ByRef _Error As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            Dim oTransferencia As SAPbobsCOM.StockTransfer

            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            If oDocumento.GetByKey(DocEntry) Then
#Enable Warning BC42104 ' La variable 'oDocumento' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_OBSERVACION").Value = MsgError
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
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  " & ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_Error(DocEntry As Integer, TipoDocumento As String, ByVal MsgError As String, ByRef _Error As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            Dim oTransferencia As SAPbobsCOM.StockTransfer

            If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "NDE" Then  ' FACTURA DE CLIENTE
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                'oTipoTabla = "FCE"
            ElseIf TipoDocumento = "FAE" Then ''FACTURA DE ANTICIPO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
            ElseIf TipoDocumento = "NCE" Then 'NOTA DE CREDITO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            ElseIf TipoDocumento = "NDE" Then 'NOTA DE CREDITO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            ElseIf TipoDocumento = "GRE" Then 'GUIA DE REMISION - ENTREGA
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            ElseIf TipoDocumento = "TRE" Then 'GUIA DE REMISION - TRANSFERENCIAS
                Try
                    oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("funcion guardar datos de autorizacion error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then  'FACTURA DE PROVEEDOR/RETENCION                             
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            End If

            If TipoDocumento = "TRE" Then
#Disable Warning BC42104 ' La variable 'oTransferencia' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If oTransferencia.GetByKey(DocEntry) Then
#Enable Warning BC42104 ' La variable 'oTransferencia' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    Try
                        oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = MsgError
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
#Disable Warning BC42104 ' La variable 'oDocumento' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If oDocumento.GetByKey(DocEntry) Then
#Enable Warning BC42104 ' La variable 'oDocumento' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    Try
                        oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = MsgError
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
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  " & ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion(DocEntry As Integer, TipoDocumento As String, ByVal _Nombre_Proveedor_SAP_BO As String, ByVal _NumAutorizacion As String, ByVal _FechaAutorizacion As DateTime, ByVal _NumeroDeDocumentoSRI As String, ByVal _Observacion As String, ByVal _EstadoAutorizacion As String, ByVal _ClaveAcceso As String, ByRef _Error As String) As Boolean
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
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "01"

            ElseIf TipoDocumento = "FAE" Then ''FACTURA DE ANTICIPO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDownPayments
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "01"

            ElseIf TipoDocumento = "NCE" Then 'NOTA DE CREDITO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "04"
            ElseIf TipoDocumento = "NDE" Then 'NOTA DE DEBITO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "04"
            ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
            End If

            If TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then

            Else
                If oDocumento.GetByKey(DocEntry) Then

                    If _NumAutorizacion <> "" Then
                        oDocumento.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacion.ToString()

                        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then
                            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                oDocumento.UserFields.Fields.Item("U_SS_NumAutRet").Value = _NumAutorizacion.ToString()
                            End If
                        Else
                            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                oDocumento.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacion.ToString()
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                oDocumento.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacion.ToString()
                            End If

                            Try 'SI PARAMETRO ESTA ACTIVO, GUARDA EL NUMERO DE DOCUMENTO QUE SE ENVIÓ AL SRI EN EL CAMPO NUMATCARD
                                If Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = "Y" Then
                                    oDocumento.NumAtCard = _NumeroDeDocumentoSRI
                                    Utilitario.Util_Log.Escribir_Log("NumeroDeDocumentoSRI: " + _NumeroDeDocumentoSRI.ToString(), "ManejoDeDocumentos")
                                End If
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al setear NumeroDeDocumentoSRI: " + ex.Message.ToString(), "ManejoDeDocumentos")
                            End Try

                        End If

                        If _tipoManejo = "A" Then
                            Try
                                rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString() + " Tipo Doc: " + TipoDocumento.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try

                        End If


                    End If
                    Try
                        'Dim fecha_formateada = _FechaAutorizacion.ToString("yyyyMMdd")
                        oDocumento.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion '.ToString("yyyyMMdd")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_SYP_FECAUTOC").Value = Date.Now
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_SYP_FECAUTOC DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString + " Fecha y Hora Autorización " + _FechaAutorizacion.ToString("yyyyMMdd")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                        oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAcceso.ToString()
                    End If

                    resultado = oDocumento.Update()
                End If
            End If


            If resultado = 0 Then
                result = True
            Else
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If
                If _tipoManejo = "A" Then
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                End If
                Utilitario.Util_Log.Escribir_Log("Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), "ManejoDeDocumentos")
                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  " & _Error.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            End If
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_LQ(DocEntry As Integer, TipoDocumento As String, ByVal _Nombre_Proveedor_SAP_BO As String, ByVal _NumAutorizacion As String, ByVal _FechaAutorizacion As DateTime, ByVal _NumeroDeDocumentoSRI As String, ByVal _Observacion As String, ByVal _EstadoAutorizacion As String, ByVal _ClaveAcceso As String, ByRef _Error As String) As Boolean
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

            If TipoDocumento = "LQE" Then 'LIQUIDACION DE COMPRA
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
            End If

            If TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then

            Else
                If oDocumento.GetByKey(DocEntry) Then

                    Try
                        oDocumento.UserFields.Fields.Item("U_LQ_NUM_AUTO").Value = _NumAutorizacion.ToString()
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_LQ_NUM_AUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try


                    Try
                        'Dim fecha_formateada = _FechaAutorizacion.ToString("yyyyMMdd")
                        oDocumento.UserFields.Fields.Item("U_LQ_FECHA_AUT").Value = _FechaAutorizacion '.ToString("yyyyMMdd")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_LQ_FECHA_AUT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    'Try
                    '    oDocumento.UserFields.Fields.Item("U_SYP_FECAUTOC").Value = Date.Now
                    'Catch ex As Exception
                    '    Utilitario.Util_Log.Escribir_Log("U_SYP_FECAUTOC DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                    'End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_LQ_OBSERVACION").Value = _Observacion.ToString + " Fecha y Hora Autorización " + _FechaAutorizacion.ToString("yyyyMMdd")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_LQ_OBSERVACION errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_LQ_ESTADO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                        oDocumento.UserFields.Fields.Item("U_LQ_CLAVE").Value = _ClaveAcceso.ToString()
                    End If

                    resultado = oDocumento.Update()
                End If
            End If


            If resultado = 0 Then
                result = True
            Else
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If
                If _tipoManejo = "A" Then
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                End If
                Utilitario.Util_Log.Escribir_Log("Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), "ManejoDeDocumentos")
                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  " & _Error.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            End If
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutGuiasDesatendidas(DocEntry As Integer, TipoDocumento As String, ByVal _Nombre_Proveedor_SAP_BO As String, ByVal _NumAutorizacion As String, ByVal _FechaAutorizacion As DateTime, ByVal _NumeroDeDocumentoSRI As String, ByVal _Observacion As String, ByVal _EstadoAutorizacion As String, ByVal _ClaveAcceso As String, ByRef _Error As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim objectType As String = "" 'obtener el objtype del documento para la localizacion de topmanage
        Dim CodDoc As String = "" 'obtener el codigo del documento para la localizacion de topmanage
        Dim SerieDoc As String = ""
        Try

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService


            ' SI EXISTE ELIMINA PARA VOLVER A CREAR
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSGRNEW")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntry)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)


            If _NumAutorizacion <> "" Then


                oGeneralData.SetProperty("U_SS_NumAut", _NumAutorizacion.ToString())

                Try

                    oGeneralData.SetProperty("U_NUM_AUTO_FAC", _NumAutorizacion.ToString())
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_NUM_AUTO_FAC error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try


                If _tipoManejo = "A" Then
                    Try
                        rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                End If

            End If

            'campos normales

            Try
                'oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                ' oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion

                oGeneralData.SetProperty("U_FECHA_AUT_FACT", _FechaAutorizacion)

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            '-------------
            Try
                'oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                oGeneralData.SetProperty("U_OBSERVACION_FACT", _Observacion)

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            '--------------
            Try
                '  oTransferencia.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)

                oGeneralData.SetProperty("U_ESTADO_AUTORIZACIO", IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion))

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_ESTADO_AUTORIZACIO error : " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            '--------------
            If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                Try
                    ' oTransferencia.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAcceso.ToString()
                    oGeneralData.SetProperty("U_CLAVE_ACCESO", _ClaveAcceso.ToString())

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_CLAVE_ACCESO error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            End If
            '---------------------

            If Not String.IsNullOrEmpty(_FechaAutorizacion.ToString) Then
                Try
                    ' oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion

                    oGeneralData.SetProperty("U_FECHA_AUT_FACT", _FechaAutorizacion)

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            End If

            Try
                oGeneralService.Update(oGeneralData)
                resultado = 0
            Catch ex As Exception
                resultado = 1
            End Try



            If resultado = 0 Then
                result = True
            Else
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If
                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                Utilitario.Util_Log.Escribir_Log("Error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), "ManejoDeDocumentos")
                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización :  " & _Error.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_HESION_FACTURAGUIA(DocEntryDoc As Integer, TipoDocumento As String, ByVal _Nombre_Proveedor_SAP_BO As String, ByVal _NumAutorizacion As String, ByVal _FechaAutorizacion As DateTime, ByVal _NumeroDeDocumentoSRI As String, ByVal _Observacion As String, ByVal _EstadoAutorizacion As String, ByVal _ClaveAcceso As String, ByRef _Error As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        'Dim DocEntryUdoRet As String = ""
        Dim DocNum As String = ""
        Dim _DocNum As String = ""
        'Dim listaTran As New List(Of Integer)


        Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION GR: " + CODE.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum GR" + DocEntryDoc.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("TipoDocumento GR" + TipoDocumento.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("_tipoManejo GR" + _tipoManejo.ToString, "ManejoDeDocumentos")
        If _tipoManejo = "A" Then
            Try
                DocNum = oFuncionesB1.getRSvalue("SELECT ""DocNum"" FROM ""OINV"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            Try
                CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Facturas""='Y' and ""U_HBT_NumeroDesde1"" = '" + DocNum.ToString() + "' ", "Code", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
        Else
            Try
                DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""OINV"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            Try
                CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Facturas""='Y' and ""U_HBT_NumeroDesde1"" = '" + DocNum.ToString() + "' ", "Code", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
        End If

        Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum query" + DocNum.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("Obteniendo Code query: " + CODE.ToString, "ManejoDeDocumentos")




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
                oUserObjectMD = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                Dim sCmp As SAPbobsCOM.CompanyService
                sCmp = rCompany.GetCompanyService

                oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @HBT_GUIAREMISION: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompany.UserTables.Item("HBT_GUIAREMISION")
                oUserTable.GetByKey(CODE)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If
                oUserTable.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString
                oUserTable.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString

                RetVal = oUserTable.Update()
                If RetVal <> 0 Then

                    rCompany.GetLastError(ErrCode, ErrMsg)

                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                Else
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_DOC_ELEC: " + CODE.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    Dim oFacturaGuia As SAPbobsCOM.StockTransfer = Nothing
                    Dim docentryFG As Integer
                    Dim resultado As Integer = -1


                    Dim recordset As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select distinct U_HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OINV ON T1.U_HBT_NumeroDesde1=OINV.""DocNum"" where OINV.""DocEntry"" =" + DocEntryDoc.ToString)

                    If recordset.RecordCount > 1 Then

                        While (recordset.EoF = False)
                            docentryFG = CInt(recordset.Fields.Item("U_HBT_DocEntry").Value)
                            If DocEntryDoc <> docentryFG Then

                                oFacturaGuia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                oFacturaGuia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                                oFacturaGuia.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

                                If oFacturaGuia.GetByKey(docentryFG) Then

                                    oFacturaGuia.UserFields.Fields.Item("U_GR_CLAVE").Value = _ClaveAcceso.ToString()
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_NUM_AUTO").Value = _NumAutorizacion.ToString()
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_FECHA_AUT").Value = _FechaAutorizacion
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_OBSERVACION").Value = _Observacion.ToString
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)

                                    Try
                                        resultado = oFacturaGuia.Update()
                                    Catch ex As Exception
                                        result = False
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("Error al actualizar Factura " + docentryFG.ToString() + " : " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Factura no actualizada: " + docentryFG.ToString() + " error: " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End Try

                                    If resultado = 0 Then
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("Factura: " + docentryFG.ToString() + " actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Transferencia actualizada: " + docentryFG.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                        result = True

                                    End If

                                End If
                            End If
                            recordset.MoveNext()
                        End While

                    End If

                End If
                Return True
            Else
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla HBT_GUIAREMISION: " + CODE.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try



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

    Public Sub ReleaseGRHEISON(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Release Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
    End Sub

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

    Public Function getRecordSetGRHEISON(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRecordSet " + ex.Message.ToString, "FuncionesB1")
        End Try
        Return fRS
    End Function

    Public Function GrabaDatosAutGuiasDesatendidas_Error(DocEntry As Integer, TipoDocumento As String, ByVal MsgError As String, ByRef _Error As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService


            ' SI EXISTE ELIMINA PARA VOLVER A CREAR
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSGRNEW")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntry)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            Try
                'oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                oGeneralData.SetProperty("U_OBSERVACION_FACT", MsgError)

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try


            Try
                oGeneralService.Update(oGeneralData)
                resultado = 0
            Catch ex As Exception
                resultado = 1
            End Try


            If resultado = 0 Then
                result = True
            Else
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  " & ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_Error_FacturaGuiaRemision(DocEntry As Integer, TipoDocumento As String, ByVal MsgError As String, ByRef _Error As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String = ""

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            If oDocumento.GetByKey(DocEntry) Then
                Try
                    oDocumento.UserFields.Fields.Item("U_GR_OBSERVACION").Value = MsgError
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION error : " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    resultado = oDocumento.Update()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True
            Else

                rCompany.GetLastError(ErrCode, ErrMsg)

                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización :  " & ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function recorreError_Solsap(ByVal respuesta As Entidades.ResponseDocuments, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta Is Nothing Then
            Return mensaje
        End If

        estado = If(respuesta.codigo, "")

        If estado = "AUTORIZADO" Or estado = "2" Then
            mensaje = "Estado: AUTORIZADO"
            If Not String.IsNullOrEmpty(respuesta.mensaje) Then
                mensaje &= ", " & respuesta.mensaje
            End If
        Else
            mensaje = "Estado: " & estado
            If Not String.IsNullOrEmpty(respuesta.mensaje) Then
                mensaje &= " - " & respuesta.mensaje
            End If
            If respuesta.log IsNot Nothing AndAlso respuesta.log.Count > 0 Then
                mensaje &= " - Detalle: " & String.Join(" | ", respuesta.log)
            End If
        End If

        mensaje &= " - NÚMERO DEL DOCUMENTO: " & codigoDocumento
        Return mensaje
    End Function

End Class

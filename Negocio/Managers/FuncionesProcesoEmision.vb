Imports Functions

Public Class FuncionesProcesoEmision
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    Private _tipoManejo As String
    Private oFuncionesAddon As Functions.FuncionesAddon
    Public Sub New(company As SAPbobsCOM.Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
    End Sub
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
            End If

            If TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then

            Else
                If oDocumento.GetByKey(DocEntry) Then

                    If _NumAutorizacion <> "" Then
                        oDocumento.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacion.ToString()

                        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then

                        Else
                            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                oDocumento.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacion.ToString()
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

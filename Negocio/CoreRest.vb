Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft
Imports Newtonsoft.Json

Public Class CoreRest


    '-----SERVICIOS ECUANEXUS--------------

    Enum Servicio_Ecuanexus

        ECUA_EMISION = 1
        ' EDOC_RECEPCION = 2
        ECUA_CONSULTA = 3
        'EDOC_CONSULTA_RECEPCION = 4

        'EDOC_OPERACIONES = 5

        'EDOC_ENVIO_MAIL = 7

    End Enum

    Enum Tipo_Peticion
        GETT
        POST
        PUT
        DELETE
    End Enum

    Enum Tipo_Autenticacion
        Basic
        Bearer
        Ninguno
    End Enum

    Private _UsuarioGS As String
    Private _PwdGS As String
    Private _Credenciales64 As String

    '    '----------API AUTENTICACION-----------------------
    '    Public Property WS_SolicitarToken As String = ""

    '    '----------API EMISION-----------------------

    '    '1 f compra venta
    '    Public Property WS_EmisionFacturaCompraVenta As String = ""
    '    '2 f exportacion

    '    Public Property WS_EmisonFacturaComercialExportacion As String = ""

    '    '3 f zona franca 
    '    Public Property WS_EmisionFacturaZonaFranca As String = ""

    '    '4 f telecomunicaciones
    '    Public Property WS_EmisionFacturaTelecomunicaciones As String = ""

    '    '5 f alquileres

    '    Public Property WS_EmisionFacturasBienesInmuebles As String = ""

    '    '6 f libre consignacion
    '    Public Property WS_EmisionFacturaLibreConsignacion As String = ""

    '    '7 f servicio basico

    '    Public Property WS_EmisionFacturaServicioBasico As String = ""

    '    '8 f alcanzada x ice

    '    Public Property WS_EmisionFacturaAlcanzadaIce As String = ""

    '    '9 f exportacion de servicios

    '    Public Property WS_EmisionFacturaExportacionServicios As String = ""

    '    '10 nota credito/debito

    '    Public Property WS_EmisionNotaCreditoDebito As String = ""

    '    '11 nota de conciliacion

    '    Public Property WS_EmisionNotaConciliacion As String = ""

    '    '35 Bonificacion

    '    Public Property WS_FacturaBonificacion As String = ""

    '    '----------API CONSULTA-----------------------
    Public Property WS_ConsultaDocumento As String = ""

    Public Property WS_EnvioDocumento As String = ""

    '    Public Property WS_ConsultaCatalogosProducto As String = ""

    '    Public Property WS_ConsultaCataloLeyendas As String = ""

    '    Public Property WS_ConsultaCodigoControl As String = ""

    '    Public Property WS_ConsultaProductoPorId As String = ""

    '    Public Property WS_ArchivoDocumento As String = ""


    '    '----------API OPERACIONES-----------------------

    '    'registro p v add 01022022

    '    Public Property WS_RegistroPuntoVenta As String = ""

    '    'cerrar p v add 01022022

    '    Public Property WS_CerrarPuntoVenta As String = ""

    '    'registrar evento significativo add 01022022

    '    Public Property WS_RegistroEventoSignificativo As String = ""

    '    'solicitar CUIS add 01022022

    '    Public Property WS_SolicitarCUIS As String = ""

    '    'solicitar CUFD add 01022022

    '    Public Property WS_SolicitarCUFD As String = ""

    '    'Anular Documento add 11/2021
    '    Public Property WS_EmisionAnulacion As String = ""

    '    'Registrar Documento fuera Linea add 01022022
    '    Public Property WS_RegistrarDocumentoOffline As String = ""

    '    'add 11/01/2022
    '    Public Property WS_RecepcionAnexos As String = ""

    '    'Verificar Nit add 01022022
    '    Public Property WS_VerificarNit As String = ""

    '    '----------VARIABLES PARA ALMACEN DE TOKENS-----------------------

    '    Private _INFO_TOKEN_EDOC_EMISION As Entidades.RespToken = Nothing
    '    ' Private _INFO_TOKEN_EDOC_RECEPCION As Entidades.RespToken = Nothing
    '    'Private _INFO_TOKEN_EDOC_CONSULTA As Entidades.RespToken = Nothing
    '    Private _INFO_TOKEN_EDOC_CONSULTA As Entidades.RespToken = Nothing
    '    'Private _INFO_TOKEN_EDOC_CONSULTA_RECEPCION As Entidades.RespToken = Nothing
    '    'Private _INFO_TOKEN_EDOC_CLIENTE As Entidades.RespToken = Nothing
    '    Private _INFO_TOKEN_EDOC_ENVIO_MAIL As Entidades.RespToken = Nothing
    '    'Private _INFO_TOKEN_EDOC_ARCHIVO_EMISION As Entidades.RespToken = Nothing
    '    'Private _INFO_TOKEN_EDOC_ARCHIVO_RECEPCION As Entidades.RespToken = Nothing
    '    Private _INFO_TOKEN_EDOC_OPERACIONES As Entidades.RespToken = Nothing

    '    Sub New()

    '    End Sub

    Sub New(UsuarioGS As String, PwdGS As String)
        _UsuarioGS = UsuarioGS
        _PwdGS = PwdGS

        _Credenciales64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(_UsuarioGS & ":" & _PwdGS))
    End Sub

    Sub New()

    End Sub


    '#Region "Core Peticion"

    Public Function EnvioDocumento(ByVal EnvioDoc As Entidades.EnvioDocumento, Optional ByRef exerror As String = "") As Entidades.RespuestaEnvio
        Try

            Dim strjson = "", strjsonResp = ""
            Dim statuscod As Integer

            strjson = ECF_TOJSON(EnvioDoc, exerror)

            If String.IsNullOrEmpty(strjson) Then Return Nothing

            'If Not ComprobarTokenGS(Servicio_eDoc.EDOC_ENVIO_MAIL, exerror) Then Return Nothing

            strjsonResp = CrearPeticion(WS_EnvioDocumento, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Ninguno)

            'rsboAppEcua.SetStatusBarMessage("Enviando documento al SRI, por favor espere..!!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
            If strjsonResp = "No es posible conectar con el servidor remoto" Then
                exerror = strjsonResp
            Else
                If Not String.IsNullOrWhiteSpace(strjsonResp) Then

                    Return GetRespEnvioDocumento(strjsonResp, exerror)

                Else
                    exerror = strjsonResp

                End If
            End If



        Catch ex As Exception
            exerror = ex.Message
        End Try

        Return Nothing
    End Function

    Public Function ConsultaDocumento(ByVal consultaDoc As Entidades.ConsultaDocumento, Optional ByRef exerror As String = "") As Entidades.ConsultaDocRespuesta
        Try

            Dim strjson = "", strjsonResp = ""
            Dim statuscod As Integer

            'strjson = ECF_TOJSON(consultaDoc, exerror)

            'If String.IsNullOrEmpty(strjson) Then Return Nothing

            'If Not ComprobarTokenGS(Servicio_eDoc.EDOC_ENVIO_MAIL, exerror) Then Return Nothing

            Dim URLFormat As String = Functions.VariablesGlobales._WsEmisionConsultaEcua
            URLFormat = URLFormat.Replace("NombreWS", consultaDoc.NombreWs)
            URLFormat = URLFormat.Replace("clave", consultaDoc.clave)
            URLFormat = URLFormat.Replace("Ruc", consultaDoc.ruc)
            URLFormat = URLFormat.Replace("DocType", consultaDoc.docType)
            URLFormat = URLFormat.Replace("DocNumber", consultaDoc.docNumber)

            strjsonResp = CrearPeticion(URLFormat, statuscod, Tipo_Peticion.GETT, Nothing, Tipo_Autenticacion.Ninguno)

            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

                Return GetRespConsultaDoc(strjsonResp, exerror)

            Else
                exerror = strjsonResp

            End If


        Catch ex As Exception
            exerror = ex.Message
        End Try

        Return Nothing
    End Function


    '    Public Function SERVICIO_EDOC_CONSULTA_ARCHIVO(ByVal ARCHIVOPARAM As Entidades.FileEmision_REST, Optional ByRef exerror As String = "") As Entidades.RespFileAppResultDto
    '        Try

    '            Dim strjson = "", strjsonResp = ""
    '            Dim statuscod As Integer

    '            strjson = ECF_TOJSON(ARCHIVOPARAM, exerror)

    '            If String.IsNullOrEmpty(strjson) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_CONSULTA, exerror) Then Return Nothing

    '            strjsonResp = CrearPeticion(WS_ArchivoDocumento, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_CONSULTA)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespFile(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function


    '    Public Function SERVICIO_EDOC_CONSULTACATALOGO(ByVal NITEMISOR As String, ByVal IDCATALOGO As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoGeneralesAppResultDto)
    '        Try

    '            Dim strjsonResp = ""
    '            Dim statuscod As Integer

    '            If String.IsNullOrWhiteSpace(NITEMISOR) Or String.IsNullOrWhiteSpace(IDCATALOGO) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_CONSULTA, exerror) Then
    '                Utilitario.Util_Log.Escribir_Log("Token invalido para servicio Edoc : " + IDCATALOGO + " ex: " + exerror.ToString(), "CoreRest")
    '                Return Nothing
    '            End If


    '            'Para el servicio de consulta se hara replace por los siguientes datos
    '            '"RNCEmisor=GSIDENTIFICACION&NumDocumento=GSNDOCUMENTO" 'consulta
    '            ' Dim WSCONSULTAREPLACE As String = WS_ConsultaCatalogos.Replace("GSIDENTIFICACION", NITEMISOR).Replace("IDCATALOGO", IDCATALOGO)

    '            Dim WSCONSULTAREPLACE As String = WS_ConsultaCatalogos & "?" & "nit=" & NITEMISOR & "&catalogo=" & IDCATALOGO

    '            strjsonResp = CrearPeticion(WSCONSULTAREPLACE, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_CONSULTA)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespCatalogo(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_CONSULTACATALOGOPRODUCTO(ByVal NITEMISOR As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoProductoAppResultDto)
    '        Try

    '            Dim strjsonResp = ""
    '            Dim statuscod As Integer

    '            If String.IsNullOrWhiteSpace(NITEMISOR) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_CONSULTA, exerror) Then
    '                Utilitario.Util_Log.Escribir_Log("Token invalido para servicio Edoc : SERVICIO_EDOC_CONSULTACATALOGOPRODUCTO ex: " + exerror.ToString(), "CoreRest")
    '                Return Nothing
    '            End If


    '            'Para el servicio de consulta se hara replace por los siguientes datos
    '            '"RNCEmisor=GSIDENTIFICACION&NumDocumento=GSNDOCUMENTO" 'consulta
    '            Dim WSCONSULTAREPLACE As String = WS_ConsultaCatalogosProducto & "?" & "nit=" & NITEMISOR

    '            strjsonResp = CrearPeticion(WSCONSULTAREPLACE, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_CONSULTA)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespCatalogoProducto(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_CONSULTADOCUMENTO(ByVal NITEMISOR As String, ByVal CUFDOCUMENTO As String, Optional ByRef exerror As String = "") As Entidades.RespEmisionAppResultDto
    '        Try

    '            Dim strjsonResp = ""
    '            Dim statuscod As Integer

    '            If String.IsNullOrWhiteSpace(NITEMISOR) Or String.IsNullOrWhiteSpace(CUFDOCUMENTO) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_CONSULTA, exerror) Then
    '                Utilitario.Util_Log.Escribir_Log("Token invalido: " + exerror.ToString(), "CoreRest")
    '                Return Nothing
    '            End If


    '            'Para el servicio de consulta se hara replace por los siguientes datos
    '            '"RNCEmisor=GSIDENTIFICACION&NumDocumento=GSNDOCUMENTO" 'consulta
    '            Dim WSCONSULTAREPLACE As String = WS_ConsultaDocumento & "?" & "nit=" & NITEMISOR & "&cuf=" & CUFDOCUMENTO

    '            strjsonResp = CrearPeticion(WSCONSULTAREPLACE, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_CONSULTA)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespEdoc(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_EMISION(ByVal ECF As Object, ByVal CodDoc As String, Optional ByRef exerror As String = "") As Entidades.RespEmisionAppResultDto
    '        Try

    '            Dim strjson = "", strjsonResp = ""
    '            Dim statuscod As Integer
    '            Dim UrlEmision = ""

    '            strjson = ECF_TOJSON(ECF, exerror)

    '            If String.IsNullOrEmpty(strjson) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_EMISION, exerror) Then Return Nothing

    '            UrlEmision = GetWSEmisionByCodeSIN(CodDoc)

    '            strjsonResp = CrearPeticion(UrlEmision, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_EMISION)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespEdoc(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_EMISION_STR(ByVal ECF As String, ByVal CodDoc As String, Optional ByRef exerror As String = "") As Entidades.RespEmisionAppResultDto
    '        Try

    '            Dim strjson = "", strjsonResp = ""
    '            Dim statuscod As Integer
    '            Dim UrlEmision = ""

    '            strjson = ECF 'ECF_TOJSON(ECF, exerror)

    '            If String.IsNullOrEmpty(strjson) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_EMISION, exerror) Then Return Nothing

    '            UrlEmision = GetWSEmisionByCodeSIN(CodDoc)

    '            strjsonResp = CrearPeticion(UrlEmision, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_EMISION)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespEdoc(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function



    '    Public Function SERVICIO_EDOC_ANULACION(ByVal objAnulacion As Entidades.Anulacion_REST, Optional ByRef exerror As String = "") As Entidades.RespAnulacionAppResultDto
    '        Try

    '            Dim strjson = "", strjsonResp = ""
    '            Dim statuscod As Integer


    '            strjson = ECF_TOJSON(objAnulacion, exerror)

    '            If String.IsNullOrEmpty(strjson) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_OPERACIONES, exerror) Then Return Nothing

    '            strjsonResp = CrearPeticion(WS_EmisionAnulacion, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_OPERACIONES)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespAnulacion(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    ' Add 25/11/2021

    '    Public Function SERVICIO_EDOC_CONSULTACATALOGOLEYENDAS(ByVal NITEMISOR As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoLeyendasAppResultDto)
    '        Try

    '            Dim strjsonResp = ""
    '            Dim statuscod As Integer

    '            If String.IsNullOrWhiteSpace(NITEMISOR) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_CONSULTA, exerror) Then
    '                Utilitario.Util_Log.Escribir_Log("Token invalido para servicio Edoc : SERVICIO_EDOC_CONSULTACATALOGOLEYENDAS ex: " + exerror.ToString(), "CoreRest")
    '                Return Nothing
    '            End If


    '            'Para el servicio de consulta se hara replace por los siguientes datos
    '            '"RNCEmisor=GSIDENTIFICACION&NumDocumento=GSNDOCUMENTO" 'consulta
    '            Dim WSCONSULTAREPLACE As String = WS_ConsultaCataloLeyendas & "?" & "nit=" & NITEMISOR

    '            strjsonResp = CrearPeticion(WSCONSULTAREPLACE, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_CONSULTA)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespCatalogoLeyendas(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_CONSULTACODIGOCONTROL(ByVal NITEMISOR As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoProductoAppResultDto)
    '        Try

    '            Dim strjsonResp = ""
    '            Dim statuscod As Integer

    '            If String.IsNullOrWhiteSpace(NITEMISOR) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_CONSULTA, exerror) Then
    '                Utilitario.Util_Log.Escribir_Log("Token invalido para servicio Edoc : SERVICIO_EDOC_CONSULTACATALOGOPRODUCTO ex: " + exerror.ToString(), "CoreRest")
    '                Return Nothing
    '            End If


    '            'Para el servicio de consulta se hara replace por los siguientes datos
    '            '"RNCEmisor=GSIDENTIFICACION&NumDocumento=GSNDOCUMENTO" 'consulta
    '            Dim WSCONSULTAREPLACE As String = WS_ConsultaCatalogosProducto & "?" & "nit=" & NITEMISOR

    '            strjsonResp = CrearPeticion(WSCONSULTAREPLACE, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_CONSULTA)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespCatalogoProducto(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_RECEPCION_ANEXOS(ByVal objRecepcionAnexos As Entidades.RecepcionAnexos_REST, Optional ByRef exerror As String = "") As Entidades.RespRecepcionAnexos
    '        Try

    '            Dim strjson = "", strjsonResp = ""
    '            Dim statuscod As Integer


    '            strjson = ECF_TOJSON(objRecepcionAnexos, exerror)

    '            If String.IsNullOrEmpty(strjson) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_OPERACIONES, exerror) Then Return Nothing

    '            strjsonResp = CrearPeticion(WS_RecepcionAnexos, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_OPERACIONES)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespAnexos(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function

    '    Public Function SERVICIO_EDOC_CONSULTAR_CUFD(ByVal objConsultaCufd As Entidades.SolicitudCUFD, Optional ByRef exerror As String = "") As Entidades.RespSincronizarCufdAppResultDto
    '        Try

    '            Dim strjson = "", strjsonResp = ""
    '            Dim statuscod As Integer


    '            strjson = ECF_TOJSON(objConsultaCufd, exerror)

    '            If String.IsNullOrEmpty(strjson) Then Return Nothing

    '            If Not ComprobarTokenGS(Servicio_eDoc.EDOC_OPERACIONES, exerror) Then Return Nothing

    '            strjsonResp = CrearPeticion(WS_SolicitarCUFD, statuscod, Tipo_Peticion.POST, strjson, Tipo_Autenticacion.Bearer, Servicio_eDoc.EDOC_OPERACIONES)

    '            If Not String.IsNullOrWhiteSpace(strjsonResp) Then

    '                Return GetRespCufd(strjsonResp, exerror)

    '            Else
    '                exerror = strjsonResp

    '            End If


    '        Catch ex As Exception
    '            exerror = ex.Message
    '        End Try

    '        Return Nothing
    '    End Function


    '    Private Function ComprobarTokenGS(ByVal s As Servicio_eDoc, Optional ByRef exerror As String = "") As Boolean

    '        Dim resp As String = ""
    '        Dim statuscod As Integer
    '        Dim EndPoint As String = WS_SolicitarToken & CInt(s).ToString

    '        Try

    '            Select Case s

    '                Case Servicio_eDoc.EDOC_EMISION

    '                    If IsNothing(_INFO_TOKEN_EDOC_EMISION) Then

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_EMISION=NOTHING", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_EMISION = GetToken(resp)

    '                            Return True

    '                        Else

    '                            exerror = resp

    '                        End If

    '                    Else

    '                        If Not TokenExpirado(_INFO_TOKEN_EDOC_EMISION.expira) Then Return True

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_EMISION=CADUCADO", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_EMISION = GetToken(resp)

    '                            Return True

    '                        Else

    '                            exerror = resp

    '                        End If


    '                    End If


    '                Case Servicio_eDoc.EDOC_CONSULTA

    '                    If IsNothing(_INFO_TOKEN_EDOC_CONSULTA) Then

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_CONSULTA=NOTHING", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_CONSULTA = GetToken(resp)

    '                            Return True


    '                        Else
    '                            exerror = resp

    '                        End If

    '                    Else

    '                        If Not TokenExpirado(_INFO_TOKEN_EDOC_CONSULTA.Expira) Then Return True

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_CONSULTA=CADUCADO", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_CONSULTA = GetToken(resp)

    '                            Return True

    '                        Else

    '                            exerror = resp

    '                        End If


    '                    End If

    '                Case Servicio_eDoc.EDOC_OPERACIONES

    '                    If IsNothing(_INFO_TOKEN_EDOC_OPERACIONES) Then

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_OPERACIONES=NOTHING", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_OPERACIONES = GetToken(resp)

    '                            Return True

    '                        Else

    '                            exerror = resp
    '                        End If

    '                    Else

    '                        If Not TokenExpirado(_INFO_TOKEN_EDOC_OPERACIONES.Expira) Then Return True


    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_OPERACIONES=CADUCADO", "CoreRest")


    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_OPERACIONES = GetToken(resp)

    '                            Return True

    '                        Else

    '                            exerror = resp

    '                        End If


    '                    End If
    '                Case Servicio_eDoc.EDOC_ENVIO_MAIL
    '                    If IsNothing(_INFO_TOKEN_EDOC_ENVIO_MAIL) Then

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_ENVIO_MAIL=NOTHING", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_ENVIO_MAIL = GetToken(resp)

    '                            Return True

    '                        Else
    '                            exerror = resp

    '                        End If

    '                    Else

    '                        If Not TokenExpirado(_INFO_TOKEN_EDOC_ENVIO_MAIL.expira) Then Return True

    '                        Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS->_INFO_TOKEN_EDOC_ENVIO_MAIL=CADUCADO", "CoreRest")

    '                        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                            _INFO_TOKEN_EDOC_ENVIO_MAIL = GetToken(resp)

    '                            Return True

    '                        Else

    '                            exerror = resp

    '                        End If



    '                    End If

    '                    'Case Servicio_eDoc.EDOC_RECEPCION
    '                    '    If IsNothing(_INFO_TOKEN_EDOC_RECEPCION) Then

    '                    '        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                    '        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                    '            _INFO_TOKEN_EDOC_RECEPCION = GetToken(resp)

    '                    '            Return True

    '                    '        Else
    '                    '            exerror = resp
    '                    '        End If

    '                    '    Else

    '                    '        If Not TokenExpirado(_INFO_TOKEN_EDOC_RECEPCION.expira) Then Return True

    '                    '        resp = CrearPeticion(EndPoint, statuscod, Tipo_Peticion.GETT, , Tipo_Autenticacion.Basic)

    '                    '        If statuscod = 200 And Not String.IsNullOrWhiteSpace(resp) Then

    '                    '            _INFO_TOKEN_EDOC_RECEPCION = GetToken(resp)

    '                    '            Return True

    '                    '        Else

    '                    '            exerror = resp

    '                    '        End If



    '                    '    End If


    '                    'Catalogos 05/10/2021



    '            End Select

    '        Catch ex As Exception
    '            Utilitario.Util_Log.Escribir_Log("ComprobarTokenGS- Catch : " + ex.Message.ToString(), "CoreRest")

    '        End Try

    '        Return False
    '    End Function


    Private Function CrearPeticion(ByVal WSURL As String,
                                   ByRef Status_Code As Integer,
                                   Optional ByVal Peticion As Tipo_Peticion = Tipo_Peticion.GETT,
                                   Optional ByVal CuerpoPeticion As String = "",
                                   Optional ByVal Autenticacion As Tipo_Autenticacion = Tipo_Autenticacion.Basic) As String

        Try

            SeteaProtocoloSeguridad()

            Dim request As WebRequest = WebRequest.Create(WSURL)

            request.ContentType = "application/json"

            'Select Case Autenticacion

            '    Case Tipo_Autenticacion.Basic

            '        request.Headers.Add("Authorization", "Basic " + _Credenciales64)
            '        Utilitario.Util_Log.Escribir_Log("token " & ":" & _Credenciales64.ToString(), "CoreRest")
            '    Case Tipo_Autenticacion.Bearer

            '        request.Headers.Add("Authorization", "Bearer " & GetTokenByServicio(servicio))
            '        Utilitario.Util_Log.Escribir_Log("token " & ":" & GetTokenByServicio(servicio).ToString(), "CoreRest")

            'End Select

            Utilitario.Util_Log.Escribir_Log("WSURL " & ":" & WSURL, "CoreRest")
            Utilitario.Util_Log.Escribir_Log("Peticion " & ":" & Peticion.ToString(), "CoreRest")
            'Utilitario.Util_Log.Escribir_Log("CuerpoPeticion " & ":" & CuerpoPeticion.ToString(), "CoreRest")
            Utilitario.Util_Log.Escribir_Log("Autenticacion " & ":" & Autenticacion.ToString(), "CoreRest")
            'Utilitario.Util_Log.Escribir_Log("servicio " & ":" & servicio.ToString(), "CoreRest")


            Select Case Peticion
                Case Tipo_Peticion.PUT, Tipo_Peticion.POST

                    request.Method = Peticion.ToString

                    Dim data = Encoding.UTF8.GetBytes(CuerpoPeticion)

                    request.ContentLength = data.Length

                    Using _stream = request.GetRequestStream()

                        _stream.Write(data, 0, data.Length)

                    End Using

                Case Tipo_Peticion.GETT

                    request.Method = "GET"

                Case Tipo_Peticion.DELETE

                    request.Method = Peticion.ToString

            End Select


            Dim responseString As String = ""

            Try

                Dim response As HttpWebResponse = request.GetResponse()

                Status_Code = response.StatusCode

                responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

            Catch ex As WebException

                Dim response As HttpWebResponse = ex.Response

                Status_Code = response.StatusCode

                responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

            End Try

            Utilitario.Util_Log.Escribir_Log("Codigo Resp Servidor " & Status_Code.ToString & " ,responseString " & ":" & responseString.ToString(), "CoreRest")

            Return responseString

        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("Exception General en CrearPeticion() " & ex.Message, "CoreRest")

            Return ex.Message

        End Try

    End Function




    '#End Region

    '#Region "Conversion Formatos"

    Private Function GetRespEnvioDocumento(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.RespuestaEnvio
        If Not String.IsNullOrWhiteSpace(s) Then

            Try
                Return JsonConvert.DeserializeObject(Of Entidades.RespuestaEnvio)(s)

            Catch ex As Exception
                exerror = ex.Message
            End Try

        End If

        Return Nothing
    End Function

    Private Function GetRespConsultaDoc(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.ConsultaDocRespuesta
        If Not String.IsNullOrWhiteSpace(s) Then

            Try
                'Dim _lineas = File.ReadAllLines(s, Encoding.UTF7)
                Return JsonConvert.DeserializeObject(Of Entidades.ConsultaDocRespuesta)(s)

            Catch ex As Exception
                exerror = ex.Message
            End Try

        End If

        Return Nothing
    End Function

    '    Private Function GetRespCufd(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.RespSincronizarCufdAppResultDto
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of Entidades.RespSincronizarCufdAppResultDto)(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespAnexos(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.RespRecepcionAnexos
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of Entidades.RespRecepcionAnexos)(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespAnulacion(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.RespAnulacionAppResultDto
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of Entidades.RespAnulacionAppResultDto)(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespFile(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.RespFileAppResultDto
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of Entidades.RespFileAppResultDto)(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespEdoc(ByVal s As String, Optional ByRef exerror As String = "") As Entidades.RespEmisionAppResultDto
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of Entidades.RespEmisionAppResultDto)(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespCatalogo(ByVal s As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoGeneralesAppResultDto)
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of List(Of Entidades.RespCatalogoGeneralesAppResultDto))(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespCatalogoLeyendas(ByVal s As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoLeyendasAppResultDto)
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of List(Of Entidades.RespCatalogoLeyendasAppResultDto))(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetRespCatalogoProducto(ByVal s As String, Optional ByRef exerror As String = "") As List(Of Entidades.RespCatalogoProductoAppResultDto)
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of List(Of Entidades.RespCatalogoProductoAppResultDto))(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try

    '        End If

    '        Return Nothing
    '    End Function

    '    Private Function GetToken(s As String, Optional ByRef exerror As String = "") As Entidades.RespToken
    '        If Not String.IsNullOrWhiteSpace(s) Then

    '            Try
    '                Return JsonConvert.DeserializeObject(Of Entidades.RespToken)(s)

    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try


    '        End If

    '        Return Nothing
    '    End Function


    Private Function ECF_TOJSON(a As Object, Optional ByRef exerror As String = "") As String

        If Not IsNothing(a) Then

            Try
                Dim settings As New Json.JsonSerializerSettings()

                settings.NullValueHandling = Json.NullValueHandling.Ignore
                settings.Formatting = Formatting.Indented

                Dim ECFSTR As String = JsonConvert.SerializeObject(a, settings)

                Try
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"

                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log(ECFSTR, Guid.NewGuid.ToString("N") & "_" & DateTime.Now.ToString.Replace(":", ".").Replace("/", "-").Replace("\", "-"))
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "CoreRest_Serializacion")
                End Try

                Return ECFSTR
            Catch ex As Exception
                exerror = ex.Message
            End Try


        End If

        Return String.Empty

    End Function
    '#End Region


    '#Region "Utilidades"


    Private Sub SeteaProtocoloSeguridad()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999
    End Sub


    '    Private Function GetTokenByServicio(s As Servicio_eDoc) As String

    '        Select Case s

    '            Case Servicio_eDoc.EDOC_EMISION
    '                Return _INFO_TOKEN_EDOC_EMISION.Token
    '                'Case Servicio_eDoc.EDOC_ARCHIVO_EMISION
    '                'Return _INFO_TOKEN_EDOC_ARCHIVO_EMISION.token
    '            Case Servicio_eDoc.EDOC_CONSULTA
    '                Return _INFO_TOKEN_EDOC_CONSULTA.Token
    '            Case Servicio_eDoc.EDOC_OPERACIONES
    '                Return _INFO_TOKEN_EDOC_OPERACIONES.Token
    '            Case Servicio_eDoc.EDOC_ENVIO_MAIL
    '                Return _INFO_TOKEN_EDOC_ENVIO_MAIL.Token

    '        End Select

    '        Return ""

    '    End Function


    '    Private Function TokenExpirado(fexpiracion As String) As Boolean
    '        Try

    '            If DateTime.Compare(DateTimeOffset.Now.LocalDateTime, DateTimeOffset.Parse(fexpiracion).LocalDateTime) > 0 Then

    '                Utilitario.Util_Log.Escribir_Log("Funcion TokenExpirado->SI", "CoreRest")

    '                Return True
    '            End If

    '        Catch ex As Exception

    '        End Try

    '        Utilitario.Util_Log.Escribir_Log("Funcion TokenExpirado->NO", "CoreRest")

    '        Return False
    '    End Function


    '    '--add 03-05-2022
    '    Public Function DICCIONARIO_TOJSON(a As Dictionary(Of String, Double), Optional ByRef exerror As String = "") As String

    '        If Not IsNothing(a) Then

    '            Try
    '                Dim settings As New Json.JsonSerializerSettings()

    '                settings.NullValueHandling = Json.NullValueHandling.Ignore
    '                settings.Formatting = Formatting.Indented

    '                Dim ECFSTR As String = JsonConvert.SerializeObject(a, settings)

    '                Try
    '                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"

    '                    If System.IO.Directory.Exists(sRutaCarpeta) Then
    '                        Utilitario.Util_Log.Escribir_Log(ECFSTR, Guid.NewGuid.ToString("N") & "_" & DateTime.Now.ToString.Replace(":", ".").Replace("/", "-").Replace("\", "-"))
    '                    End If

    '                Catch ex As Exception
    '                    Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "CoreRest_DICCIONARIO_TOJSON")
    '                End Try

    '                Return ECFSTR
    '            Catch ex As Exception
    '                exerror = ex.Message
    '            End Try


    '        End If

    '        Return String.Empty

    '    End Function


    '    'add 05/10/2021

    '    Private Function GetWSEmisionByCodeSIN(ByVal Code As String) As String

    '        Dim _ws As String = ""

    '        Try

    '            Select Case Code
    '                Case "1"
    '                    Return WS_EmisionFacturaCompraVenta
    '                Case "3"
    '                    Return WS_EmisonFacturaComercialExportacion
    '                Case "5"
    '                    Return WS_EmisionFacturaZonaFranca
    '                Case "22"
    '                    Return WS_EmisionFacturaTelecomunicaciones
    '                Case "24"
    '                    Return WS_EmisionNotaCreditoDebito
    '                Case "14"
    '                    Return WS_EmisionFacturaAlcanzadaIce
    '                Case "4"
    '                    Return WS_EmisionFacturaLibreConsignacion
    '                Case "29"
    '                    Return WS_EmisionNotaConciliacion
    '                Case "13"
    '                    Return WS_EmisionFacturaServicioBasico
    '                Case "2"
    '                    Return WS_EmisionFacturasBienesInmuebles
    '                Case "28"
    '                    Return WS_EmisionFacturaExportacionServicios
    '                Case "35"
    '                    Return WS_FacturaBonificacion

    '            End Select


    '        Catch ex As Exception

    '        End Try

    '        Return _ws

    '    End Function

    '#End Region

End Class

Imports SAPbobsCOM
Imports Functions
Imports Documentos
Imports System.Globalization
Imports Negocio.Documentos

Public Class GuiaRemiManager

    Private rCompany As Company
    Private rsboApp As SAPbouiCOM.Application
    Private oFuncionesAddon As Functions.FuncionesAddon
    Private _tipoManejo As String
    Private dbManager As DatabaseQueryManager
    Private DesencriptarQuery_ As DesencriptarQuery

    Public Sub New(company As Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon, db As DatabaseQueryManager, dsQm As DesencriptarQuery)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
        dbManager = db
        DesencriptarQuery_ = dsQm
    End Sub

    Public Function ConsultarGuiaDeRemision(TipoGR As String, DocEntry As Integer) As Object

        Dim oGuiaRemision As Entidades.RequestGuiaRemision = Nothing
        Dim oDestinatario As Entidades.destinatarioGR
        Dim listaDestinatarios As List(Of Entidades.destinatarioGR)
        Dim listaDetalles As List(Of Entidades.detalleGR)
        Dim listaDatosAdicional As List(Of Entidades.infoAdicionalGR)
        Dim listaDatosAdicionalDetalle As List(Of Entidades.detalleAdicionalGR)

        oDestinatario = New Entidades.destinatarioGR
        listaDestinatarios = New List(Of Entidades.destinatarioGR)
        listaDetalles = New List(Of Entidades.detalleGR)
        listaDatosAdicional = New List(Of Entidades.infoAdicionalGR)
        listaDatosAdicionalDetalle = New List(Of Entidades.detalleAdicionalGR)

        Dim SP As String = ""
        Try
            If TipoGR = "GRE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(tipoDocumento.GuiaRemisionEntrega, DocEntry)
            ElseIf TipoGR = "TRE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(tipoDocumento.GuiaRemisionTraslado, DocEntry)
            ElseIf TipoGR = "TLE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(tipoDocumento.GuiaRemisionSolicitudTraslado, DocEntry)
            ElseIf TipoGR = "SSGR" Then
                SP = DesencriptarQuery_.GetQueryConsulta(tipoDocumento.GuiaRemisionDesatendida, DocEntry)
            End If

            Utilitario.Util_Log.Escribir_Log("Query Desencriptado " & SP.ToString(), "ManejoDeDocumentos")

            If SP.Contains("GSCODEEXCEPCION") Then
                Utilitario.Util_Log.Escribir_Log("EXCEPCION DETECTADA EN EL PROCESO DE OBTENER STRING QUERY - " & SP, "ManejoDeDocumentos")
                rsboApp.StatusBar.SetText(Functions.VariablesGlobales._gNombreAddOn + " - Ocurrio Un Error Favor falidar el Archivo de Log y Buscar el Codigo GSCODEEXCEPCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If

            If SP.Contains("El relleno entre caracteres no es válido y no se puede quitar.") Then 'DM 2024-06-14 se hace el replace debido a que al desencriptar esta concatenando el siguiente texto El relleno entre caracteres no es válido y no se puede quitar.
                Utilitario.Util_Log.Escribir_Log("Texto añadido al desencriptar", "ManejoDeDocumentos")
                SP = SP.Replace("El relleno entre caracteres no es válido y no se puede quitar.", "")
                Utilitario.Util_Log.Escribir_Log("Query Desencriptado con replace " & SP.ToString(), "ManejoDeDocumentos")
            End If

            Dim ds As DataSet

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                Dim SPs() As String = Split(SP, "--*")

                Dim ds1, ds2, ds3 As DataSet
                Dim dt1, dt2, dt3 As DataTable

                ds = dbManager.EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = dbManager.EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Destinatario"
                ds.Tables.Add(dt1)

                ds2 = dbManager.EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "Detalle"
                ds.Tables.Add(dt2)

                ds3 = dbManager.EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "Adicionales"
                ds.Tables.Add(dt3)

            Else
                ds = dbManager.EjecutarSP(SP, DocEntry)
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                oGuiaRemision = New Entidades.RequestGuiaRemision()
                oGuiaRemision.infoTributaria = New Entidades.infoTributariaGR
                oGuiaRemision.infoGuiaRemision = New Entidades.infoGuiaRemisionGR
                'oGuiaRemision.destinatarios = New List(Of Entidades.destinatarioGR)
                'oGuiaRemision.infoAdicional = New List(Of Entidades.infoAdicionalGR)

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            oGuiaRemision.infoTributaria.ambiente = r("Ambiente")

                            Dim claveAcceso As String = r("ClaveAcceso").ToString()
                            If Not String.IsNullOrEmpty(claveAcceso) AndAlso claveAcceso.Length = 49 Then
                                oGuiaRemision.infoTributaria.claveAcceso = claveAcceso
                            End If


                            oGuiaRemision.infoTributaria.tipoEmision = r("TipoEmision")
                            oGuiaRemision.infoTributaria.razonSocial = r("RazonSocial")

                            If Not r("NombreComercial") = "" Then
                                oGuiaRemision.infoTributaria.nombreComercial = r("NombreComercial")
                            End If
                            oGuiaRemision.infoTributaria.ruc = r("Ruc")
                            oGuiaRemision.infoTributaria.codDoc = r("codDoc") 'tenía como nombre CodigoDocumento
                            oGuiaRemision.infoTributaria.estab = r("Establecimiento")
                            oGuiaRemision.infoTributaria.ptoEmi = r("PuntoEmision")
                            oGuiaRemision.infoTributaria.secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemision.infoTributaria.secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemision.infoTributaria.secuencial = oGuiaRemision.infoTributaria.secuencial.PadLeft(9, "0")
                            End If
                            oGuiaRemision.infoTributaria.dirMatriz = r("DireccionMatriz")
                            oGuiaRemision.infoGuiaRemision.dirEstablecimiento = r("DireccionEstablecimiento")
                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemision.infoGuiaRemision.contribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemision.infoGuiaRemision.contribuyenteEspecial = Nothing
                            End If

                            'api no manda este campo
                            'If Not r("AgenteRetencion") = "0" Then
                            '    oGuiaRemision.infoGuiaRemision.agenteretencion = r("AgenteRetencion")
                            'End If

                            'If Not r("RegimenMicroempresas") = "0" Then
                            '    oGuiaRemision.infoGuiaRemision.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            'End If

                            'If Not r("ContribuyenteRimpe") = "0" Then
                            '    oGuiaRemision.infoGuiaRemision.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            'End If

                            oGuiaRemision.infoGuiaRemision.obligadoContabilidad = r("ObligadoContabilidad")

                            oGuiaRemision.infoTributaria.diaEmission = CDate(r("FechaEmision")).ToString("dd")
                            oGuiaRemision.infoTributaria.mesEmission = CDate(r("FechaEmision")).ToString("MM")
                            oGuiaRemision.infoTributaria.anioEmission = CDate(r("FechaEmision")).ToString("yyyy")

                            'api no manda fecha emision
                            'oGuiaRemision.infoGuiaRemision.FechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")

                            oGuiaRemision.infoGuiaRemision.dirPartida = r("DireccionPartida")
                            oGuiaRemision.infoGuiaRemision.razonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemision.infoGuiaRemision.tipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemision.infoGuiaRemision.rucTransportista = r("RucTranportista")

                            oGuiaRemision.infoGuiaRemision.fechaIniTransporte = CDate(r("FechaInicioTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.infoGuiaRemision.fechaFinTransporte = CDate(r("FechaFinTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.infoGuiaRemision.placa = r("Placa")

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows
                            oDestinatario = New Entidades.destinatarioGR

                            oDestinatario.identificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.razonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.dirDestinatario = r("DirDestinatario")

                            oDestinatario.motivoTraslado = r("MotivoTraslado")
                            oDestinatario.codEstabDestino = r("CodEstabDestino")

                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.ruta = r("Ruta")
                            End If

                            ' If oDestinatario.MotivoTraslado = "VENTA" Then
                            oDestinatario.codDocSustento = r("CodDocSustento")
                            oDestinatario.numDocSustento = r("NumDocSustento")
                            oDestinatario.numAutDocSustento = r("NumAutDocSustento")
                            '  oDestinatario.FechaEmisionDocSustentoSpecified = True
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.fechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            listaDestinatarios.Add(oDestinatario)

                        Next

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As New Entidades.detalleGR
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.codigoInterno = r("CodigoPrincipal")
                            itemDetalle.codigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.descripcion = r("Descripcion")
                            itemDetalle.cantidad = r("Cantidad")

                            Dim listaDetalleDatoAdicional As Object
                            listaDetalleDatoAdicional = New List(Of Entidades.detalleAdicionalGR)

                            'Adicional 1
                            If Not r("ConceptoAdicional1") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.detalleAdicionalGR
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional2") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.detalleAdicionalGR
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional3") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.detalleAdicionalGR
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            itemDetalle.detallesAdicionales = listaDetalleDatoAdicional

                            'agrego detalle a la lista
                            listaDetalles.Add(itemDetalle)
                        Next

                        oDestinatario.detalles = listaDetalles
                        oGuiaRemision.destinatarios = listaDestinatarios

                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As New Entidades.infoAdicionalGR
                            itemDatoAdicionalFac.nombre = r("Concepto")
                            itemDatoAdicionalFac.valor = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oGuiaRemision.infoAdicional = listaDatosAdicional
                    End If
                Next
            End If

            Return oGuiaRemision
            Utilitario.Util_Log.Escribir_Log($"GUIA CONSULTADA - {TipoGR}", "ManejoDeDocumentos")

        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Guia de Remisión en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then
                oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "ArgumentException-Error al Consultar Guia de Remisión con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la oGuiaRemision en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then
                oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Error al Consultar Guia de Remisión con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        End Try

    End Function
End Class

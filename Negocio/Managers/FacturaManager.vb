Imports SAPbobsCOM
Imports Newtonsoft.Json
Imports System.Net
Imports Functions
Imports Documentos
Imports System.Globalization

Public Class FacturaManager
    Private rCompany As Company
    Private rsboApp As SAPbouiCOM.Application
    Private oFuncionesAddon As Functions.FuncionesAddon
    Private _tipoManejo As String
    Private dbManager As DatabaseQueryManager
    'Private parent As ManejoDeDocumentoSolsap
    Private DesencriptarQuery_ As DesencriptarQuery

    Public Sub New(company As Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon, db As DatabaseQueryManager, dsQm As DesencriptarQuery)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
        dbManager = db
        'parent = owner
        DesencriptarQuery_ = dsQm
    End Sub

    Private Function ConsultaFacturaExportacion(TipoFactura As String, DocEntry As Integer, oFactura As Entidades.RequestFactura) As Object
        'Dim oFactura As Entidades.RequestFactura = Nothing

        Dim SP As String = ""
        Try

            If TipoFactura = "FAE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.FacturaAnticipo, DocEntry, "EXPORT")
            Else
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.Factura, DocEntry, "EXPORT")
            End If


            If SP.Contains("GSCODEEXCEPCION") Then

                Utilitario.Util_Log.Escribir_Log("EXCEPCION DETECTADA EN EL PROCESO DE OBTENER STRING QUERY - " & SP, "ManejoDeDocumentos")
                rsboApp.StatusBar.SetText(Functions.VariablesGlobales._gNombreAddOn + " - Ocurrio Un Error Favor falidar el Archivo de Log y Buscar el Codigo GSCODEEXCEPCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If

            Dim ds As DataSet

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")
                Dim SPs() As String = Split(SP, "--*")
                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020
                ds = dbManager.EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Exportacion"

            Else

                ds = dbManager.EjecutarSP(SP, DocEntry)

            End If

            Utilitario.Util_Log.Escribir_Log("Data Tables EXPORTACION : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then ' TABLA 0 ES EXPORTACION
                        For Each r As DataRow In ds.Tables(0).Rows

                            oFactura.infoFactura.comercioExterior = r("ComercioExterior")
                            oFactura.infoFactura.IncoTermFactura = r("IncoTermFactura")
                            oFactura.infoFactura.lugarIncoTerm = r("LugarIncoTerm")
                            oFactura.infoFactura.paisOrigen = r("PaisOrigen")
                            oFactura.infoFactura.puertoEmbarque = r("PuertoEmbarque")
                            oFactura.infoFactura.paisDestino = r("PaisDestino")
                            oFactura.infoFactura.paisAdquisicion = r("PaisAdquisicion")
                            oFactura.infoFactura.incoTermTotalSinImpuestos = r("IncoTermTotalSinImpuestos")
                            oFactura.infoFactura.fleteInternacional = FormatearNumero(r("FleteInternacional"))
                            oFactura.infoFactura.seguroInternacional = FormatearNumero(r("SeguroInternacional"))
                            oFactura.infoFactura.gastosAduaneros = FormatearNumero(r("GastosAduaneros"))
                            oFactura.infoFactura.gastosTransporteOtros = FormatearNumero(r("GastosTransporteOtros"))
                        Next
                    End If
                Next
            End If

            Return oFactura
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "ArgumentException-Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
        End Try
    End Function

    Private Function ConsultaFacturaReembolso(TipoFactura As String, DocEntry As Integer, oFactura As Entidades.RequestFactura) As Object
        'Dim oFactura As Entidades.RequestFactura = Nothing

        Dim SP As String = ""
        Try

            If TipoFactura = "FAE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.FacturaAnticipo, DocEntry, "REEMBOLSO")
            Else
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.Factura, DocEntry, "REEMBOLSO")
            End If


            If SP.Contains("GSCODEEXCEPCION") Then

                Utilitario.Util_Log.Escribir_Log("EXCEPCION DETECTADA EN EL PROCESO DE OBTENER STRING QUERY - " & SP, "ManejoDeDocumentos")
                rsboApp.StatusBar.SetText(Functions.VariablesGlobales._gNombreAddOn + " - Ocurrio Un Error Favor falidar el Archivo de Log y Buscar el Codigo GSCODEEXCEPCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If

            Dim ds As DataSet


            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")
                Dim SPs() As String = Split(SP, "--*")
                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

                Dim ds1 As DataSet
                Dim dt1 As DataTable

                ds = dbManager.EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Rembolso"

                ds1 = dbManager.EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "RembolsoDet"
                ds.Tables.Add(dt1)

            Else

                ds = dbManager.EjecutarSP(SP, DocEntry)

            End If

            Utilitario.Util_Log.Escribir_Log("Data Tables REEMBOLSO : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then ' TABLA 0 ES EXPORTACION
                        For Each r As DataRow In ds.Tables(0).Rows
                            ' Dim X As New Entidades.wsEDoc_Factura.ENTFactura
                            oFactura.infoFactura.codDocReembolso = r("CodDocReemb")
                            oFactura.infoFactura.totalComprobantesReembolso = r("TotalComprobantesReembolso")
                            oFactura.infoFactura.totalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                            oFactura.infoFactura.totalImpuestoReembolso = r("TotalImpuestoReembolso")

                        Next
                    ElseIf i = 1 Then

                        Dim Detalle As Object
                        Dim listaDetalle As Object
                        Dim LisimpdetalleIVA As Object
                        Dim impdetalleIVA As Object


                        'Detalle = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolso
                        'listaDetalle = New List(Of Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolso)
                        'LisimpdetalleIVA = New List(Of Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto)
                        'impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto

                        Detalle = New Entidades.reembolsoFE
                        listaDetalle = New List(Of Entidades.reembolsoFE)
                        LisimpdetalleIVA = New List(Of Entidades.detImpReemFE)
                        impdetalleIVA = New Entidades.detImpReemFE


                        For Each r As DataRow In ds.Tables(1).Rows
                            Detalle = New Entidades.reembolsoFE
                            LisimpdetalleIVA.Clear()

                            'Dim x As Entidades.wsEDoc_Factura.ENTFacturaReembolso
                            'TipoIdentificacionProveedorReembolso()
                            'TipoIdentificacionProveedorReembolso()
                            'TipoIdentificacionProveedorReembolso()

                            Detalle.TipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                            Detalle.IdentificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                            Detalle.CodPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                            Detalle.TipoProveedorReembolso = r("TipoProveedorReembolso")
                            Detalle.CodDocReembolso = r("CodDocReembolso")
                            Detalle.EstabDocReembolso = r("EstabDocReembolso")
                            Detalle.PtoEmiDocReembolso = r("PtoEmiDocReembolso")
                            Detalle.SecuencialDocReembolso = r("SecuencialDocReembolso")
                            'Detalle.FechaEmisionDocReembolso = r("FechaEmisionDocReembolso")
                            Detalle.FechaEmisionDocReembolso = CDate(r("FechaEmisionDocReembolso")).ToString("yyyy-MM-dd")
                            Detalle.NumeroAutorizacionDocReemb = r("NumeroAutorizacionDocReemb")

                            If r("BaseImponibleIVA0").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIVA0")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIVA0")
                                impdetalleIVA.tarifa = r("TarifaIVA0")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIVA0")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIVA0")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)

                            End If

                            If r("BaseImponibleIVA").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIVA")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIVA")
                                impdetalleIVA.tarifa = r("TarifaIVA")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIVA")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIVA")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleNoObjIVA").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoNoObjIVA")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeNoObjIVA")
                                impdetalleIVA.tarifa = r("TarifaNoObjIVA")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleNoObjIVA")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoNoObjIVA")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIvaExe").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIvaExe")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIvaExe")
                                impdetalleIVA.tarifa = r("TarifaIvaExe")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIvaExe")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIvaExe")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva8").ToString() <> 0 Then
                                listaDetalle = New List(Of Entidades.detImpReemFE)
                                listaDetalle.codigo = r("CodigoIva8")
                                listaDetalle.codigoPorcentaje = r("CodigoPorcentajeIva8")
                                listaDetalle.tarifa = r("TarifaIva8")
                                listaDetalle.BaseImponibleReembolso = r("BaseImponibleIva8")
                                listaDetalle.impuestoReembolso = r("ImpuestoReembolsoIva8")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If


                            If r("BaseImponibleIva5").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIva5")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIva5")
                                impdetalleIVA.tarifa = r("TarifaIva5")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva5")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIva5")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva15").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIva15")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIva15")
                                impdetalleIVA.tarifa = r("TarifaIva15")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva15")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIva15")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva14").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIva14")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIva14")
                                impdetalleIVA.tarifa = r("TarifaIva14")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva14")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIva14")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva13").ToString() <> 0 Then
                                impdetalleIVA = New Entidades.detImpReemFE
                                impdetalleIVA.codigo = r("CodigoIva13")
                                impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIva13")
                                impdetalleIVA.tarifa = r("TarifaIva13")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva13")
                                impdetalleIVA.impuestoReembolso = r("ImpuestoReembolsoIva13")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            'agrego lista de impuesto al detalle
                            Detalle.detalleImpuestos = LisimpdetalleIVA.ToList()

                            listaDetalle.Add(Detalle)

                        Next
                        oFactura.infoFactura.reembolsos = listaDetalle.ToArray()
                    End If
                Next
            End If

            Return oFactura
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "ArgumentException-Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
        End Try
    End Function

    Public Function ConsultarFactura(ByVal TipoFactura As String, ByVal DocEntry As Integer, ByRef _Error As String, ByRef _camponulo As String) As Object

        Dim oFactura As Entidades.RequestFactura = Nothing
        Dim listaDetalle As List(Of Entidades.detalleFE)
        Dim listaDatosAdicional As List(Of Entidades.infoAdicionalFE)
        Dim listaTotalesConImpuestos As List(Of Entidades.totalConImpuestosFE)
        Dim listaPagos As List(Of Entidades.pagosFE)
        Dim listaDatosAdicionalDetalle As List(Of Entidades.detallesAdicionalesFE)
        Dim listaImpuestos As List(Of Entidades.impuestosFE)

        listaDetalle = New List(Of Entidades.detalleFE)
        listaDatosAdicional = New List(Of Entidades.infoAdicionalFE)
        listaTotalesConImpuestos = New List(Of Entidades.totalConImpuestosFE)
        listaPagos = New List(Of Entidades.pagosFE)

        Try
            Dim SP As String = ""

            If Functions.VariablesGlobales._vgGuardarLog = "Y" Then
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, $"Tipo de factura = {TipoFactura}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, $"Consultando Factura con # DocEntry = {DocEntry}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If

            'Utilitario.Util_Log.Escribir_Log("SP: " + SP.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ANTES A CONSULTAR", "ManejoDeDocumentos")

            If TipoFactura = "FAE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.FacturaAnticipo, DocEntry)
            Else
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.Factura, DocEntry)
            End If

            Utilitario.Util_Log.Escribir_Log("Query Desencriptado " & SP.ToString(), "ManejoDeDocumentos")

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
                dt1.TableName = "Detalles"
                ds.Tables.Add(dt1)

                ds2 = dbManager.EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "InfoAdicionales"
                ds.Tables.Add(dt2)

                ds3 = dbManager.EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "FormaPago"
                ds.Tables.Add(dt3)
            Else
                ds = dbManager.EjecutarSP(SP, DocEntry)
            End If

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not dbManager.ValidarCamposNulos(ds, "2", _camponulo) Then Return Nothing
            End If

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("INGRESANDO A CONSULTAR", "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oFactura = New Entidades.RequestFactura
                oFactura.infoTributaria = New Entidades.infoTributariaFE()
                oFactura.infoFactura = New Entidades.infoFacturaFE()
                Dim tipo = ""
                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try
                            For Each r As DataRow In ds.Tables(0).Rows

                                oFactura.infoTributaria.ambiente = r("Ambiente")

                                Dim claveAcceso As String = r("ClaveAcceso").ToString()
                                If Not String.IsNullOrEmpty(claveAcceso) AndAlso claveAcceso.Length = 49 Then
                                    oFactura.infoTributaria.claveAcceso = claveAcceso
                                End If

                                oFactura.infoTributaria.razonSocial = r("RazonSocial")

                                oFactura.infoTributaria.nombreComercial = r("NombreComercial")

                                oFactura.infoTributaria.ruc = r("RUC")

                                oFactura.infoTributaria.tipoEmision = r("TipoEmision")

                                oFactura.infoTributaria.codDoc = r("CodigoDocumento")

                                oFactura.infoTributaria.estab = r("Establecimiento")

                                oFactura.infoTributaria.ptoEmi = r("PuntoEmision")

                                oFactura.infoTributaria.secuencial = r("SecuencialDocumento")
                                If Not oFactura.infoTributaria.secuencial.ToString().Length.Equals("9") Then oFactura.infoTributaria.secuencial = oFactura.infoTributaria.secuencial.ToString().PadLeft(9, "0")
                                Utilitario.Util_Log.Escribir_Log("oFactura.Secuencial : " & oFactura.infoTributaria.secuencial.ToString(), "ManejoDeDocumentos")

                                oFactura.infoTributaria.dirMatriz = r("DireccionMatriz")

                                oFactura.infoTributaria.diaEmission = CDate(r("FechaEmision")).ToString("dd")

                                oFactura.infoTributaria.mesEmission = CDate(r("FechaEmision")).ToString("MM")

                                oFactura.infoTributaria.anioEmission = CDate(r("FechaEmision")).ToString("yyyy")

                                Try
                                    'El servicio de facturación exige la fecha en formato dd/MM/yyyy
                                    oFactura.infoFactura.fechaEmision = CDate(r("FechaEmision")).ToString("dd/MM/yyyy")
                                    Utilitario.Util_Log.Escribir_Log("oFactura.FechaEmision : " & CDate(r("FechaEmision")).ToString("dd/MM/yyyy"), "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("oFactura.FechaEmision : " & ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                oFactura.infoFactura.dirEstablecimiento = r("DireccionEstablecimiento")

                                oFactura.infoFactura.contribuyenteEspecial = r("ContribuyenteEspecial")

                                oFactura.infoFactura.obligadoContabilidad = r("ObligadoContabilidad")

                                oFactura.infoFactura.tipoIdentificacionComprador = r("TipoIdentificadorComprador")

                                If Not r("GuiaRemision") = "0" Then oFactura.infoFactura.guiaRemision = r("GuiaRemision")

                                oFactura.infoFactura.razonSocialComprador = r("RazonSocialComprador")

                                oFactura.infoFactura.identificacionComprador = r("IdentificacionComprador")

                                oFactura.infoFactura.direccionComprador = r("DirComprador")

                                oFactura.infoFactura.totalSinImpuestos = FormatearNumero(r("TotalSinImpuesto"))

                                oFactura.infoFactura.totalDescuento = FormatearNumero(r("TotalDescuento"))

                                oFactura.infoFactura.propina = FormatearNumero(r("Propina"))

                                oFactura.infoFactura.importeTotal = FormatearNumero(r("ImporteTotal"))

                                oFactura.infoFactura.moneda = r("Moneda").ToString

                                tipo = r("TipoFactura")

                                'DATOS DE EXPORTACION

                                'If r("TipoFactura") = "1" Or r("TipoFactura") = "41" Then
                                'tipo = r("TipoFactura")
                                'oFactura.infoFactura.comercioExterior = r("ComercioExterior")
                                'oFactura.infoFactura.IncoTermFactura = r("IncoTermFactura")
                                'oFactura.infoFactura.lugarIncoTerm = r("LugarIncoTerm")
                                'oFactura.infoFactura.paisOrigen = r("PaisOrigen")
                                'oFactura.infoFactura.puertoEmbarque = r("PuertoEmbarque")
                                'oFactura.infoFactura.paisDestino = r("PaisDestino")
                                'oFactura.infoFactura.paisAdquisicion = r("PaisAdquisicion")
                                'oFactura.infoFactura.incoTermTotalSinImpuestos = r("IncoTermTotalSinImpuestos")
                                'oFactura.infoFactura.fleteInternacional = r("FleteInternacional")
                                'oFactura.infoFactura.seguroInternacional = r("SeguroInternacional")
                                'oFactura.infoFactura.gastosAduaneros = r("GastosAduaneros")
                                'oFactura.infoFactura.gastosTransporteOtros = r("GastosTransporteOtros")
                                'End If


                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo8")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base8"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva8"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo12")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base12"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva12"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo13")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base13"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva13"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo0")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base0"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva0"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoNoi")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("BaseNoi"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIvaNoi"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoExen")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("BaseExen"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIvaExen"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseIce") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoIce")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("BaseIce"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIvaIce"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base5") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo5")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje5")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("Base5"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIva5"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base15") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo15")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje15")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("Base15"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIva15"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base14") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo14")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje14")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("Base14"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIva14"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                Utilitario.Util_Log.Escribir_Log("Termina cabecera ", "ManejoDeDocumentos")

                                oFactura.infoFactura.totalConImpuestos = listaTotalesConImpuestos
                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleFactura As Entidades.detalleFE = New Entidades.detalleFE

                                itemDetalleFactura.codigoPrincipal = r("CodigoPrincipal").ToString

                                itemDetalleFactura.codigoAuxiliar = r("CodigoAuxiliar").ToString

                                itemDetalleFactura.descripcion = r("Descripcion").ToString

                                itemDetalleFactura.cantidad = CInt(r("Cantidad"))

                                itemDetalleFactura.precioUnitario = FormatearNumero(r("PrecioUnitario"))

                                itemDetalleFactura.descuento = FormatearNumero(r("Descuento"))

                                itemDetalleFactura.precioTotalSinImpuesto = FormatearNumero(r("PrecioTotalSinImpuesto"))

                                listaDatosAdicionalDetalle = New List(Of Entidades.detallesAdicionalesFE)

                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As Entidades.detallesAdicionalesFE = New Entidades.detallesAdicionalesFE
                                    itemDetalleDatoAdicional.nombre = r("ConceptoAdicional1").ToString
                                    itemDetalleDatoAdicional.valor = r("NombreAdicional1").ToString
                                    listaDatosAdicionalDetalle.Add(itemDetalleDatoAdicional)
                                End If

                                If Not r("ConceptoAdicional2") = "0" Then
                                    Dim itemDetalleDatoAdicional2 As Entidades.detallesAdicionalesFE = New Entidades.detallesAdicionalesFE
                                    itemDetalleDatoAdicional2.nombre = r("ConceptoAdicional2")
                                    itemDetalleDatoAdicional2.valor = r("NombreAdicional2")
                                    listaDatosAdicionalDetalle.Add(itemDetalleDatoAdicional2)
                                End If

                                If Not r("ConceptoAdicional3") = "0" Then
                                    Dim itemDetalleDatoAdicional3 As Entidades.detallesAdicionalesFE = New Entidades.detallesAdicionalesFE
                                    itemDetalleDatoAdicional3.nombre = r("ConceptoAdicional3")
                                    itemDetalleDatoAdicional3.valor = r("NombreAdicional3")
                                    listaDatosAdicionalDetalle.Add(itemDetalleDatoAdicional3)
                                End If

                                itemDetalleFactura.detallesAdicionales = listaDatosAdicionalDetalle

                                listaImpuestos = New List(Of Entidades.impuestosFE)

                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA8" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA13" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("CodigoIce").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentajeIce").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponibleIce"))
                                    impuesto.valor = FormatearNumero(r("TotalIvaIce"))
                                    impuesto.tarifa = FormatearNumero(r("TarifaIce"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA5" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA15" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA14" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                itemDetalleFactura.impuestos = listaImpuestos

                                listaDetalle.Add(itemDetalleFactura)
                            Next
                            Utilitario.Util_Log.Escribir_Log("Termina detalle", "ManejoDeDocumentos")
                            oFactura.detalles = listaDetalle
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalFac As Entidades.infoAdicionalFE = New Entidades.infoAdicionalFE
                                itemDatoAdicionalFac.nombre = r("Concepto")
                                itemDatoAdicionalFac.valor = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalFac)
                            Next
                            Utilitario.Util_Log.Escribir_Log("Termina info adicional ", "ManejoDeDocumentos")
                            oFactura.infoAdicional = listaDatosAdicional
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Cabecera Campo Adicional: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Informacion Adicional: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 3 Then
                        Try
                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim Pago As Entidades.pagosFE = New Entidades.pagosFE
                                Pago.formaPago = r("FormaPago").ToString
                                Pago.total = FormatearNumero(r("Total"))
                                Pago.plazo = r("Plazo").ToString
                                Pago.unidadTiempo = r("UnidadTiempo").ToString
                                listaPagos.Add(Pago)
                            Next
                            Utilitario.Util_Log.Escribir_Log("Termina forma de pagp", "ManejoDeDocumentos")
                            oFactura.infoFactura.pagos = listaPagos
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    End If

                    'EXPORTACION
                    If tipo = "1" And i = 0 Then
                        oFactura = ConsultaFacturaExportacion(TipoFactura, DocEntry, oFactura)
                    End If

                    If tipo = "2" And i = 0 Then
                        oFactura = ConsultaFacturaReembolso(TipoFactura, DocEntry, oFactura)
                    End If
                Next
            End If

            Return oFactura
            Utilitario.Util_Log.Escribir_Log("FACTURA CONSULTADA", "ManejoDeDocumentos")

        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"ArgumentException-Ocurrio un error al consultar datos de la factura en la Base, DocEntry: {DocEntry} Descr: {x.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, $"ArgumentException-Error al Consultar Factura con # DocEntry = {DocEntry}, Descr: {x.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"Ocurrio un error al consultar datos de la factura en la Base, DocEntry: {DocEntry} Descr: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, $"Error al Consultar Factura con # DocEntry = {DocEntry}, Descr: {ex.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        End Try
    End Function

    Private Function FormatearNumero(valor As Object) As String
        If valor Is Nothing Then Return Nothing

        Dim numero As Decimal
        If Decimal.TryParse(valor.ToString(), numero) Then
            Return numero.ToString(CultureInfo.InvariantCulture)
        End If

        Return valor.ToString()
    End Function
End Class

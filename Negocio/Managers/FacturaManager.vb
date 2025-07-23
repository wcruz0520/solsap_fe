Imports SAPbobsCOM
Imports Newtonsoft.Json
Imports System.Net

Public Class FacturaManager
    Private rCompany As Company
    Private rsboApp As SAPbouiCOM.Application
    Private oFuncionesAddon As Functions.FuncionesAddon
    Private _tipoManejo As String
    Private dbManager As DatabaseQueryManager
    Private parent As ManejoDeDocumentoSolsap

    Public Sub New(company As Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon, db As DatabaseQueryManager, owner As ManejoDeDocumentoSolsap)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
        dbManager = db
        parent = owner
    End Sub

    Public Function ConsultarFactura(ByVal TipoFactura As String, ByVal DocEntry As Integer) As Object

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
                SP = parent.GetQueryConsulta(tipoDocumento.FacturaAnticipo, DocEntry)
            Else
                SP = parent.GetQueryConsulta(tipoDocumento.Factura, DocEntry)
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
                If Not parent.ValidarCamposNulos(ds, "2") Then Return Nothing
            End If

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("INGRESANDO A CONSULTAR", "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oFactura = New Entidades.RequestFactura
                oFactura.infoTributaria = New Entidades.infoTributariaFE()
                oFactura.infoFactura = New Entidades.infoFacturaFE()

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

                                oFactura.infoFactura.totalSinImpuestos = parent.FormatearNumero(r("TotalSinImpuesto"))

                                oFactura.infoFactura.totalDescuento = parent.FormatearNumero(r("TotalDescuento"))

                                oFactura.infoFactura.propina = parent.FormatearNumero(r("Propina"))

                                oFactura.infoFactura.importeTotal = parent.FormatearNumero(r("ImporteTotal"))

                                oFactura.infoFactura.moneda = r("Moneda").ToString

                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo8")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impfaIVA.baseImponible = parent.FormatearNumero(r("Base8"))
                                    impfaIVA.valor = parent.FormatearNumero(r("ValorIva8"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo12")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impfaIVA.baseImponible = parent.FormatearNumero(r("Base12"))
                                    impfaIVA.valor = parent.FormatearNumero(r("ValorIva12"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo13")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impfaIVA.baseImponible = parent.FormatearNumero(r("Base13"))
                                    impfaIVA.valor = parent.FormatearNumero(r("ValorIva13"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo0")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impfaIVA.baseImponible = parent.FormatearNumero(r("Base0"))
                                    impfaIVA.valor = parent.FormatearNumero(r("ValorIva0"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoNoi")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impfaNOIVA.baseImponible = parent.FormatearNumero(r("BaseNoi"))
                                    impfaNOIVA.valor = parent.FormatearNumero(r("ValorIvaNoi"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoExen")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impfaNOIVA.baseImponible = parent.FormatearNumero(r("BaseExen"))
                                    impfaNOIVA.valor = parent.FormatearNumero(r("ValorIvaExen"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseIce") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoIce")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    impfaNOIVA.baseImponible = parent.FormatearNumero(r("BaseIce"))
                                    impfaNOIVA.valor = parent.FormatearNumero(r("ValorIvaIce"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base5") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo5")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje5")
                                    impfaNOIVA.baseImponible = parent.FormatearNumero(r("Base5"))
                                    impfaNOIVA.valor = parent.FormatearNumero(r("ValorIva5"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base15") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo15")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje15")
                                    impfaNOIVA.baseImponible = parent.FormatearNumero(r("Base15"))
                                    impfaNOIVA.valor = parent.FormatearNumero(r("ValorIva15"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base14") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo14")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje14")
                                    impfaNOIVA.baseImponible = parent.FormatearNumero(r("Base14"))
                                    impfaNOIVA.valor = parent.FormatearNumero(r("ValorIva14"))
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

                                itemDetalleFactura.precioUnitario = parent.FormatearNumero(r("PrecioUnitario"))

                                itemDetalleFactura.descuento = parent.FormatearNumero(r("Descuento"))

                                itemDetalleFactura.precioTotalSinImpuesto = parent.FormatearNumero(r("PrecioTotalSinImpuesto"))

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
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA8" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA13" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("CodigoIce").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentajeIce").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponibleIce"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIvaIce"))
                                    impuesto.tarifa = parent.FormatearNumero(r("TarifaIce"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA5" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA15" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = parent.FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = parent.FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = parent.FormatearNumero(r("Tarifa"))
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
                                Pago.total = parent.FormatearNumero(r("Total"))
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
End Class

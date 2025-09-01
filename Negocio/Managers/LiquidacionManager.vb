Imports SAPbobsCOM
Imports Functions
Imports System.Globalization

Public Class LiquidacionManager
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

    Public Function ConsultarLiquidacion(ByVal DocEntry As Integer, ByRef _Error As String) As Entidades.RequestLiquidacion
        Dim liquidacion As New Entidades.RequestLiquidacion
        liquidacion.infoTributaria = New Entidades.infoTributariaLQ
        liquidacion.infoLiquidacionCompra = New Entidades.infoLiquidacionCompraLQ
        liquidacion.detalles = New List(Of Entidades.detalleLQ)()
        liquidacion.reembolsos = New List(Of Entidades.reembolsoLQ)()
        liquidacion.infoAdicional = New List(Of Entidades.campoAdicionalLQ)()

        Try
            Dim SP As String = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.Liquidacion, DocEntry)
            If SP.Contains("El relleno entre caracteres no es válido y no se puede quitar.") Then
                SP = SP.Replace("El relleno entre caracteres no es válido y no se puede quitar.", "")
            End If

            Dim ds As DataSet
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Dim SPs() As String = Split(SP, "--*")
                ds = dbManager.EjecutarSP(SPs(0), DocEntry)
                ds.Tables(0).TableName = "Cabecera"
                If SPs.Length > 1 Then
                    Dim dt As DataTable = dbManager.EjecutarSP(SPs(1), DocEntry).Tables(0).Copy
                    dt.TableName = "Detalles"
                    ds.Tables.Add(dt)
                End If
                If SPs.Length > 2 Then
                    Dim dt As DataTable = dbManager.EjecutarSP(SPs(2), DocEntry).Tables(0).Copy
                    dt.TableName = "Reembolsos"
                    ds.Tables.Add(dt)
                End If
                If SPs.Length > 3 Then
                    Dim dt As DataTable = dbManager.EjecutarSP(SPs(3), DocEntry).Tables(0).Copy
                    dt.TableName = "InfoAdicional"
                    ds.Tables.Add(dt)
                End If
                If SPs.Length > 4 Then
                    Dim dt As DataTable = dbManager.EjecutarSP(SPs(4), DocEntry).Tables(0).Copy
                    dt.TableName = "Pagos"
                    ds.Tables.Add(dt)
                End If
            Else
                ds = dbManager.EjecutarSP(SP, DocEntry)
            End If

            If ds Is Nothing OrElse ds.Tables.Count = 0 Then Return Nothing

            If ds.Tables.Count > 0 Then
                For Each r As DataRow In ds.Tables(0).Rows
                    liquidacion.infoTributaria.ambiente = r("Ambiente").ToString()
                    liquidacion.infoTributaria.tipoEmision = r("TipoEmision").ToString()

                    Dim claveacceso As String = r("ClaveAcceso")
                    If claveacceso.Length >= 10 And claveacceso.Length <= 49 Then
                        liquidacion.infoTributaria.claveAcceso = r("ClaveAcceso").ToString()
                    End If

                    liquidacion.infoTributaria.razonSocial = r("RazonSocial").ToString()
                    liquidacion.infoTributaria.nombreComercial = r("NombreComercial").ToString()
                    liquidacion.infoTributaria.ruc = r("Ruc").ToString()
                    liquidacion.infoTributaria.codDoc = r("CodigoDocumento").ToString()
                    liquidacion.infoTributaria.estab = r("Establecimiento").ToString()
                    liquidacion.infoTributaria.ptoEmi = r("PuntoEmision").ToString()
                    liquidacion.infoTributaria.secuencial = r("SecuencialDocumento").ToString().PadLeft(9, "0"c)
                    liquidacion.infoTributaria.dirMatriz = r("DireccionMatriz").ToString()

                    Dim fecha As Date = Date.Parse(r("FechaEmision").ToString())
                    liquidacion.infoTributaria.diaEmission = fecha.ToString("dd")
                    liquidacion.infoTributaria.mesEmission = fecha.ToString("MM")
                    liquidacion.infoTributaria.anioEmission = fecha.ToString("yyyy")

                    liquidacion.infoLiquidacionCompra.fechaEmision = fecha.ToString("dd/MM/yyyy")

                    Try
                        liquidacion.infoTributaria.campoAdicional1 = r("campoAdicional1")
                        Utilitario.Util_Log.Escribir_Log(" liquidacion.infoTributaria.campoAdicional1 : " & r("campoAdicional1"), "ManejoDeDocumentos")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" liquidacion.infoTributaria.campoAdicional1 : " & ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        liquidacion.infoTributaria.campoAdicional2 = r("campoAdicional2")
                        Utilitario.Util_Log.Escribir_Log(" liquidacion.infoTributaria.campoAdicional2 : " & r("campoAdicional2"), "ManejoDeDocumentos")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" liquidacion.infoTributaria.campoAdicional2 : " & ex.Message.ToString, "ManejoDeDocumentos")
                    End Try


                    liquidacion.infoLiquidacionCompra.dirEstablecimiento = r("DireccionEstablecimiento").ToString()

                    Dim contri As String = r("ContribuyenteEspecial")

                    If contri <> "0" And contri.Length = 3 Then
                        liquidacion.infoLiquidacionCompra.contribuyenteEspecial = r("ContribuyenteEspecial").ToString()
                    End If

                    liquidacion.infoLiquidacionCompra.obligadoContabilidad = r("ObligadoContabilidad").ToString()
                    liquidacion.infoLiquidacionCompra.tipoIdentificacionProveedor = r("TipoIdentificacionProveedor").ToString()
                    liquidacion.infoLiquidacionCompra.razonSocialProveedor = r("RazonSocialProveedor").ToString()
                    liquidacion.infoLiquidacionCompra.identificacionProveedor = r("IdentificacionProveedor").ToString()
                    liquidacion.infoLiquidacionCompra.direccionProveedor = r("DirProveedor").ToString()
                    liquidacion.infoLiquidacionCompra.totalSinImpuestos = FormatearNumero(r("TotalSinImpuesto"))
                    liquidacion.infoLiquidacionCompra.totalDescuento = FormatearNumero(r("TotalDescuento"))

                    Dim codReemb As String = r("CodDocReemb").ToString()

                    If codReemb.Length > 0 Then
                        liquidacion.infoLiquidacionCompra.codDocReembolso = r("CodDocReemb").ToString()
                        liquidacion.infoLiquidacionCompra.totalComprobantesReembolso = r("TotalComprobantesReembolso").ToString()
                        liquidacion.infoLiquidacionCompra.totalBaseImponibleReembolso = r("TotalBaseImponibleReembolso").ToString()
                        liquidacion.infoLiquidacionCompra.totalImpuestoReembolso = r("TotalImpuestoReembolso").ToString()
                    End If

                    liquidacion.infoLiquidacionCompra.importeTotal = FormatearNumero(r("ImporteTotal"))
                    liquidacion.infoLiquidacionCompra.moneda = r("Moneda").ToString()

                    Dim totales As New List(Of Entidades.totalConImpuestoLQ)
                    Dim sufijos As String() = {"0", "5", "8", "12", "13", "14", "15", "Exen", "Ice", "Noi"}
                    For Each suf In sufijos
                        Dim baseCol As String = "Base" & suf
                        If r.Table.Columns.Contains(baseCol) AndAlso Convert.ToDecimal(r(baseCol)) <> 0 Then
                            Dim t As New Entidades.totalConImpuestoLQ
                            t.codigo = r("Codigo" & suf).ToString()
                            t.codigoPorcentaje = r("CodigoPorcentaje" & suf).ToString()
                            t.baseImponible = FormatearNumero(r("Base" & suf))
                            t.valor = FormatearNumero(r("ValorIva" & suf))
                            t.tarifa = FormatearNumero(r("Tarifa" & suf))
                            If r("DescuentoAdicional" & suf) > 0 Then
                                t.descuentoAdicional = FormatearNumero(r("DescuentoAdicional" & suf))
                            End If
                            totales.Add(t)
                        End If
                    Next
                    liquidacion.infoLiquidacionCompra.totalConImpuestos = totales
                Next
            End If

            If ds.Tables.Count > 1 Then
                For Each r As DataRow In ds.Tables(1).Rows
                    Dim det As New Entidades.detalleLQ
                    det.codigoPrincipal = r("CodigoPrincipal").ToString()
                    det.codigoAuxiliar = r("CodigoAuxiliar").ToString()
                    det.descripcion = r("Descripcion").ToString()
                    det.cantidad = CInt(r("Cantidad"))
                    det.precioUnitario = FormatearNumero(r("PrecioUnitario"))
                    det.descuento = FormatearNumero(r("Descuento"))
                    det.precioTotalSinImpuesto = FormatearNumero(r("PrecioTotalSinImpuesto"))
                    det.unidadMedida = r("UnidadMedida").ToString()

                    Dim lstImp As New List(Of Entidades.impuestoLQ)
                    Dim imp As New Entidades.impuestoLQ
                    imp.codigo = r("Codigo").ToString()
                    imp.codigoPorcentaje = r("CodigoPorcentaje").ToString()
                    imp.baseImponible = FormatearNumero(r("BaseImponible"))
                    imp.valor = FormatearNumero(r("TotalIva"))
                    imp.tarifa = CInt(r("Tarifa")).ToString()
                    lstImp.Add(imp)
                    det.impuestos = lstImp

                    Dim lstAd As New List(Of Entidades.detallesAdicionalLQ)
                    If r.Table.Columns.Contains("ConceptoAdicional1") Then
                        If r("ConceptoAdicional1").ToString().Trim() <> "-" Then
                            Dim da As New Entidades.detallesAdicionalLQ
                            da.nombre = r("ConceptoAdicional1").ToString()
                            da.valor = r("NombreAdicional1").ToString()
                            lstAd.Add(da)
                        End If
                    End If
                    det.detallesAdicionales = lstAd

                    liquidacion.detalles.Add(det)
                Next
            End If

            If ds.Tables.Count > 2 Then
                For Each r As DataRow In ds.Tables(2).Rows
                    Dim re As New Entidades.reembolsoLQ
                    re.tipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso").ToString()
                    re.identificacionProveedorReembolso = r("IdentificacionProveedorReembolso").ToString()
                    re.codPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso").ToString()
                    re.tipoProveedorReembolso = r("TipoProveedorReembolso").ToString()
                    re.codDocReembolso = r("CodDocReembolso").ToString()
                    re.estabDocReembolso = r("EstabDocReembolso").ToString()
                    re.ptoEmiDocReembolso = r("PtoEmiDocReembolso").ToString()
                    re.secuencialDocReembolso = r("SecuencialDocReembolso").ToString()
                    re.fechaEmisionDocReembolso = CDate(r("FechaEmisionDocReembolso")).ToString("dd/MM/yyyy")
                    re.numeroautorizacionDocReemb = r("NumeroAutorizacionDocReem").ToString()

                    Dim listaImp As New List(Of Entidades.detalleImpuestoReembolsoLQ)
                    Dim sufijos As String() = {"0", "5", "8", "12", "13", "14", "15", "Exen", "Ice", "Noi"}

                    For Each suf In sufijos
                        Dim baseCol As String = "Base" & suf
                        If r.Table.Columns.Contains(baseCol) AndAlso Convert.ToDecimal(r(baseCol)) <> 0 Then
                            Dim i As New Entidades.detalleImpuestoReembolsoLQ
                            i.codigo = r("Codigo" & suf).ToString()
                            i.codigoPorcentaje = r("CodigoPorcentaje" & suf).ToString()
                            i.tarifa = FormatearNumero(r("Tarifa" & suf))
                            i.baseImponibleReembolso = FormatearNumero(r("Base" & suf))
                            i.impuestoReembolso = FormatearNumero(r("ValorIvaReem" & suf))
                            listaImp.Add(i)
                        End If
                    Next

                    re.detalleImpuestos = listaImp
                    liquidacion.reembolsos.Add(re)
                Next
            End If

            If ds.Tables.Count > 3 Then
                For Each r As DataRow In ds.Tables(3).Rows
                    Dim ad As New Entidades.campoAdicionalLQ
                    ad.nombre = r("Concepto").ToString()
                    ad.valor = r("Descripcion").ToString()
                    liquidacion.infoAdicional.Add(ad)
                Next
            End If

            If ds.Tables.Count > 4 Then
                Dim lstPagos As New List(Of Entidades.pagoLQ)
                For Each r As DataRow In ds.Tables(4).Rows
                    Dim p As New Entidades.pagoLQ
                    p.formaPago = r("FormaPago").ToString()
                    p.total = FormatearNumero(r("Total"))
                    p.plazo = r("Plazo").ToString()
                    p.unidadTiempo = r("UnidadTiempo").ToString()
                    lstPagos.Add(p)
                Next
                liquidacion.infoLiquidacionCompra.pagos = lstPagos
            End If

            Return liquidacion

        Catch ex As Exception
            _Error = ex.Message
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error consultar liquidacion: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
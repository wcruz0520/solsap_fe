Imports SAPbobsCOM
Imports Functions
Imports System.Globalization
Imports Negocio.Documentos

Public Class RetencionManager
    Private rCompany As Company
    Private rsboApp As SAPbouiCOM.Application
    Private oFuncionesAddon As FuncionesAddon
    Private _tipoManejo As String
    Private dbManager As DatabaseQueryManager
    Private DesencriptarQuery_ As DesencriptarQuery

    Public Sub New(company As Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As FuncionesAddon, db As DatabaseQueryManager, dsQm As DesencriptarQuery)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
        dbManager = db
        DesencriptarQuery_ = dsQm
    End Sub

    Public Function ConsultarRetencion(DocEntry As Integer, ByVal tipoDoc As tipoDocumento) As Entidades.RequestRetencion
        Dim retencion As New Entidades.RequestRetencion
        retencion.infoTributaria = New Entidades.InfoTributariaRET
        retencion.infoCompRetencion = New Entidades.InfoCompRetencionRET
        retencion.docsSustento = New List(Of Entidades.DocSustentoRET)()
        retencion.infoAdicional = New List(Of Entidades.InfoAdicionalRET)()

        Dim SP As String = DesencriptarQuery_.GetQueryConsulta(tipoDoc, DocEntry)
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
                dt.TableName = "DocSustento"
                ds.Tables.Add(dt)
            End If
            If SPs.Length > 2 Then
                Dim dt As DataTable = dbManager.EjecutarSP(SPs(2), DocEntry).Tables(0).Copy
                dt.TableName = "Retenciones"
                ds.Tables.Add(dt)
            End If
            If SPs.Length > 3 Then
                Dim dt As DataTable = dbManager.EjecutarSP(SPs(3), DocEntry).Tables(0).Copy
                dt.TableName = "Reembolsos"
                ds.Tables.Add(dt)
            End If
            If SPs.Length > 4 Then
                Dim dt As DataTable = dbManager.EjecutarSP(SPs(4), DocEntry).Tables(0).Copy
                dt.TableName = "Adicionales"
                ds.Tables.Add(dt)
            End If
        Else
            ds = dbManager.EjecutarSP(SP, DocEntry)
        End If

        If ds Is Nothing OrElse ds.Tables.Count = 0 Then Return Nothing

        If ds.Tables.Count > 0 Then
            For Each r As DataRow In ds.Tables(0).Rows
                retencion.infoTributaria.ambiente = r("Ambiente").ToString
                retencion.infoTributaria.tipoEmision = r("TipoEmision").ToString
                retencion.infoTributaria.claveAcceso = r("ClaveAcceso").ToString
                retencion.infoTributaria.razonSocial = r("RazonSocial").ToString
                retencion.infoTributaria.nombreComercial = r("NombreComercial").ToString
                retencion.infoTributaria.ruc = r("Ruc").ToString
                retencion.infoTributaria.codDoc = r("CodigoDocumento").ToString
                retencion.infoTributaria.estab = r("Establecimiento").ToString
                retencion.infoTributaria.ptoEmi = r("PuntoEmision").ToString
                retencion.infoTributaria.secuencial = r("SecuencialDocumento").ToString.PadLeft(9, "0"c)
                retencion.infoTributaria.dirMatriz = r("DireccionMatriz").ToString

                Dim fecha As Date = Date.Parse(r("FechaEmision").ToString)
                retencion.infoTributaria.diaEmission = fecha.ToString("dd")
                retencion.infoTributaria.mesEmission = fecha.ToString("MM")
                retencion.infoTributaria.anioEmission = fecha.ToString("yyyy")

                retencion.infoCompRetencion.fechaEmision = fecha.ToString("dd/MM/yyyy")
                retencion.infoCompRetencion.dirEstablecimiento = r("DireccionEstablecimiento").ToString
                retencion.infoCompRetencion.contribuyenteEspecial = r("ContribuyenteEspecial").ToString
                retencion.infoCompRetencion.obligadoContabilidad = r("ObligadoContabilidad").ToString
                retencion.infoCompRetencion.tipoIdentificacionSujetoRetenido = r("TipoIdentificacionSujetoRetenido").ToString
                If r.Table.Columns.Contains("TipoSujetoRetenido") Then retencion.infoCompRetencion.tipoSujetoRetenido = r("TipoSujetoRetenido").ToString
                If r.Table.Columns.Contains("ParteRel") Then retencion.infoCompRetencion.parteRel = r("ParteRel").ToString
                retencion.infoCompRetencion.razonSocialSujetoRetenido = r("RazonSocialSujetoRetenido").ToString
                retencion.infoCompRetencion.identificacionSujetoRetenido = r("IdentificacionSujetoRetenido").ToString
                retencion.infoCompRetencion.periodoFiscal = r("PeriodoFiscal").ToString
            Next
        End If

        If ds.Tables.Count > 1 Then
            For Each r As DataRow In ds.Tables(1).Rows
                Dim doc As New Entidades.DocSustentoRET
                If r.Table.Columns.Contains("CodSustento") Then doc.codSustento = r("CodSustento").ToString
                doc.codDocSustento = r("CodDocRetener").ToString
                doc.numDocSustento = r("NumDocRetener").ToString
                doc.fechaEmisionDocSustento = CDate(r("FechaEmisionDocRetener")).ToString("yyyy-MM-dd")
                If r.Table.Columns.Contains("FechaRegistroContable") Then doc.fechaRegistroContable = CDate(r("FechaRegistroContable")).ToString("yyyy-MM-dd")
                If r.Table.Columns.Contains("NumAutDocSustento") Then doc.numAutDocSustento = r("NumAutDocSustento").ToString
                If r.Table.Columns.Contains("PagoLocExt") Then doc.pagoLocExt = r("PagoLocExt").ToString
                If r.Table.Columns.Contains("TipoRegi") Then doc.tipoRegi = r("TipoRegi").ToString
                If r.Table.Columns.Contains("PaisEfecPago") Then doc.paisEfecPago = r("PaisEfecPago").ToString
                If r.Table.Columns.Contains("AplicConvDobTrib") Then doc.aplicConvDobTrib = r("AplicConvDobTrib").ToString
                If r.Table.Columns.Contains("PagExtSujRetNorLeg") Then doc.pagExtSujRetNorLeg = r("PagExtSujRetNorLeg").ToString
                If r.Table.Columns.Contains("PagoRegFis") Then doc.pagRegFis = r("PagoRegFis").ToString
                If r.Table.Columns.Contains("TotalComprobantesReembolso") Then doc.totalComprobantesReembolso = r("TotalComprobantesReembolso").ToString
                If r.Table.Columns.Contains("TotalBaseImponibleReembolso") Then doc.totalBaseImponibleReembolso = r("TotalBaseImponibleReembolso").ToString
                If r.Table.Columns.Contains("TotalSinImpuestos") Then doc.totalSinImpuestos = r("TotalSinImpuestos").ToString
                If r.Table.Columns.Contains("ImporteTotal") Then doc.importeTotal = r("ImporteTotal").ToString

                doc.impuestosDocSustento = New List(Of Entidades.ImpuestoDocSustentoRET)()
                Dim sufijos As String() = {"8", "12", "0", "Noi", "Exen", "5", "15", "14", "13"}
                For Each suf In sufijos
                    Dim baseCol As String = "Base" & suf
                    Dim codImp As String = "CodImpDocSus" & suf
                    Dim codPor As String = "CodPor" & suf
                    Dim tarifa As String = "Tarifa" & suf
                    Dim valor As String = "ValorImpuesto" & suf
                    If r.Table.Columns.Contains(baseCol) AndAlso Convert.ToDecimal(r(baseCol)) <> 0 Then
                        Dim imp As New Entidades.ImpuestoDocSustentoRET
                        imp.codImpuestoDocSustento = r(codImp).ToString
                        imp.codigoPorcentaje = r(codPor).ToString
                        imp.baseImponible = r(baseCol).ToString
                        imp.tarifa = r(tarifa).ToString
                        imp.valorImpuesto = r(valor).ToString
                        doc.impuestosDocSustento.Add(imp)
                    End If
                Next

                doc.retenciones = New List(Of Entidades.RetencionRET)()
                doc.pagos = New List(Of Entidades.PagoRET)()
                If r.Table.Columns.Contains("FormaPago") Then
                    Dim pg As New Entidades.PagoRET
                    pg.formaPago = r("FormaPago").ToString
                    pg.total = r("Total").ToString
                    doc.pagos.Add(pg)
                End If

                retencion.docsSustento.Add(doc)
            Next
        End If

        If ds.Tables.Count > 2 Then
            Dim idx As Integer = 0
            For Each r As DataRow In ds.Tables(2).Rows
                If idx >= retencion.docsSustento.Count Then Exit For
                Dim re As New Entidades.RetencionRET
                re.codigo = r("Codigo").ToString
                re.codigoRetencion = r("CodigoRetencion").ToString
                re.baseImponible = r("BaseImponible").ToString
                re.porcentajeRetener = r("PorcentajeRetener").ToString
                re.valorRetenido = r("ValorRetenido").ToString
                If r.Table.Columns.Contains("FechaPagoDiv") Then
                    re.dividendos = New Entidades.DividendosRET
                    re.dividendos.fechaPagoDiv = CDate(r("FechaPagoDiv")).ToString("yyyy-MM-dd")
                    re.dividendos.imRentaSoc = r("ImRentaSoc").ToString
                    re.dividendos.ejerFisUtDiv = r("EjerFisUtDiv").ToString
                End If
                retencion.docsSustento(idx).retenciones.Add(re)
                idx += 1
            Next
        End If

        If ds.Tables.Count > 4 Then
            For Each r As DataRow In ds.Tables(4).Rows
                Dim inf As New Entidades.InfoAdicionalRET
                inf.nombre = r("Concepto").ToString
                inf.valor = r("Descripcion").ToString
                retencion.infoAdicional.Add(inf)
            Next
        End If

        Return retencion
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
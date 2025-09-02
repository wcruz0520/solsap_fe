Imports SAPbobsCOM
Imports Functions
Imports System.Globalization

Public Class NotaDebitoManager
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

    Public Function ConsultarNotaDebito(ByVal DocEntry As Integer, ByRef _Error As String) As Entidades.RequestNotaDebito
        Dim notaDebito As New Entidades.RequestNotaDebito
        notaDebito.infoTributaria = New Entidades.infoTributariaND
        notaDebito.infoNotaDebito = New Entidades.infoNotaDebitoND
        notaDebito.motivos = New List(Of Entidades.motivoND)()
        notaDebito.infoAdicional = New List(Of Entidades.infoAdicionalND)()

        Try
            Dim SP As String = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.NotaDebito, DocEntry)
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
                    dt.TableName = "Motivos"
                    ds.Tables.Add(dt)
                End If
                If SPs.Length > 2 Then
                    Dim dt As DataTable = dbManager.EjecutarSP(SPs(2), DocEntry).Tables(0).Copy
                    dt.TableName = "Adicionales"
                    ds.Tables.Add(dt)
                End If
                If SPs.Length > 3 Then
                    Dim dt As DataTable = dbManager.EjecutarSP(SPs(3), DocEntry).Tables(0).Copy
                    dt.TableName = "Pagos"
                    ds.Tables.Add(dt)
                End If
            Else
                ds = dbManager.EjecutarSP(SP, DocEntry)
            End If

            If ds Is Nothing OrElse ds.Tables.Count = 0 Then Return Nothing

            If ds.Tables(0).Rows.Count > 0 Then
                For Each r As DataRow In ds.Tables(0).Rows
                    notaDebito.infoTributaria.ambiente = r("Ambiente").ToString()
                    notaDebito.infoTributaria.tipoEmision = r("TipoEmision").ToString()

                    Dim claveAcceso As String = r("ClaveAcceso").ToString()
                    If Not String.IsNullOrEmpty(claveAcceso) AndAlso claveAcceso.Length = 49 Then
                        notaDebito.infoTributaria.claveAcceso = r("ClaveAcceso").ToString()
                    End If

                    notaDebito.infoTributaria.razonSocial = r("RazonSocial").ToString()
                    notaDebito.infoTributaria.nombreComercial = r("NombreComercial").ToString()
                    notaDebito.infoTributaria.ruc = r("RUC").ToString()
                    notaDebito.infoTributaria.codDoc = r("CodigoDocumento").ToString()
                    notaDebito.infoTributaria.estab = r("Establecimiento").ToString()
                    notaDebito.infoTributaria.ptoEmi = r("PuntoEmision").ToString()
                    notaDebito.infoTributaria.secuencial = r("SecuencialDocumento").ToString().PadLeft(9, "0"c)
                    notaDebito.infoTributaria.dirMatriz = r("DireccionMatriz").ToString()

                    Dim fecha As Date = Date.Parse(r("FechaEmision").ToString())
                    notaDebito.infoTributaria.diaEmission = fecha.ToString("dd")
                    notaDebito.infoTributaria.mesEmission = fecha.ToString("MM")
                    notaDebito.infoTributaria.anioEmission = fecha.ToString("yyyy")

                    notaDebito.infoNotaDebito.fechaEmision = fecha.ToString("dd/MM/yyyy")

                    Try
                        notaDebito.campoAdicional1 = r("campoAdicional1")
                        Utilitario.Util_Log.Escribir_Log(" notaDebito.campoAdicional1 : " & r("campoAdicional1"), "ManejoDeDocumentos")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" notaDebito.campoAdicional1 : " & ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        notaDebito.campoAdicional2 = r("campoAdicional2")
                        Utilitario.Util_Log.Escribir_Log(" notaDebito.campoAdicional2 : " & r("campoAdicional2"), "ManejoDeDocumentos")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" notaDebito.campoAdicional2 : " & ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    notaDebito.infoNotaDebito.dirEstablecimiento = r("DireccionEstablecimiento").ToString()
                    notaDebito.infoNotaDebito.tipoIdentificacionComprador = r("TipoIdentificadorComprador").ToString()
                    notaDebito.infoNotaDebito.razonSocialComprador = r("RazonSocialComprador").ToString()
                    notaDebito.infoNotaDebito.identificacionComprador = r("IdentificacionComprador").ToString()

                    Dim contri As String = r("ContribuyenteEspecial")

                    If contri <> "0" And contri.Length = 3 Then
                        notaDebito.infoNotaDebito.contribuyenteEspecial = r("ContribuyenteEspecial").ToString()
                    End If

                    notaDebito.infoNotaDebito.obligadoContabilidad = r("ObligadoContabilidad").ToString()
                    notaDebito.infoNotaDebito.codDocModificado = r("codDocModificado").ToString()
                    notaDebito.infoNotaDebito.numDocModificado = r("numDocModificado").ToString()
                    notaDebito.infoNotaDebito.fechaEmisionDocSustento = CDate(r("FechaEmisionDocModificado")).ToString("dd/MM/yyyy")
                    notaDebito.infoNotaDebito.totalSinImpuestos = FormatearNumero(r("TotalSinImpuesto"))
                    notaDebito.infoNotaDebito.valorTotal = FormatearNumero(r("ImporteTotal"))

                    Dim lstImp As New List(Of Entidades.impuestoND)()
                    Dim sufijos As String() = {"0", "5", "8", "12", "13", "14", "15", "Exen", "Ice", "Noi"}
                    For Each suf In sufijos
                        Dim baseCol As String = "Base" & suf
                        If r.Table.Columns.Contains(baseCol) AndAlso Convert.ToDecimal(r(baseCol)) <> 0 Then
                            Dim imp As New Entidades.impuestoND
                            imp.codigo = r("Codigo" & suf).ToString()
                            imp.codigoPorcentaje = r("CodigoPorcentaje" & suf).ToString()
                            imp.baseImponible = FormatearNumero(r(baseCol))
                            imp.valor = FormatearNumero(r("ValorIva" & suf))
                            imp.tarifa = FormatearNumero(r("Tarifa" & suf))
                            lstImp.Add(imp)
                        End If
                    Next
                    notaDebito.infoNotaDebito.impuestos = lstImp
                Next
            End If

            If ds.Tables(1).Rows.Count > 0 Then
                For Each r As DataRow In ds.Tables(1).Rows
                    Dim mot As New Entidades.motivoND
                    mot.razon = r("Descripcion").ToString()
                    If r.Table.Columns.Contains("PrecioTotalSinImpuesto") Then
                        mot.valor = FormatearNumero(r("PrecioTotalSinImpuesto"))
                    End If
                    notaDebito.motivos.Add(mot)
                Next
            End If

            If ds.Tables(2).Rows.Count > 0 Then
                For Each r As DataRow In ds.Tables(2).Rows
                    Dim ad As New Entidades.infoAdicionalND
                    ad.nombre = r("Concepto").ToString()
                    ad.valor = r("Descripcion").ToString()
                    notaDebito.infoAdicional.Add(ad)
                Next
            End If

            If ds.Tables(3).Rows.Count > 0 Then
                Dim lstPagos As New List(Of Entidades.pagoND)()
                For Each r As DataRow In ds.Tables(3).Rows
                    Dim pg As New Entidades.pagoND
                    pg.formaPago = r("FormaPago").ToString()
                    pg.total = FormatearNumero(r("Total"))
                    pg.plazo = r("Plazo").ToString()
                    pg.unidadTiempo = r("UnidadTiempo").ToString()
                    lstPagos.Add(pg)
                Next
                notaDebito.infoNotaDebito.pagos = lstPagos
            End If

            Return notaDebito
        Catch ex As Exception
            _Error = ex.Message
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error consultar nota de débito: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
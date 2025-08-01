Imports SAPbobsCOM
Imports Newtonsoft.Json
Imports System.Net
Imports Functions
Imports Documentos
Imports System.Globalization

Public Class NotaCreditoManager
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

    Public Function ConsultarNotaCredito(ByVal TipoNC As String, ByVal DocEntry As Integer, ByRef _Error As String, ByRef _camponulo As String) As Object

        Dim oNotaCredito As Entidades.RequestNotaCredito = Nothing
        Dim listaDetalle As List(Of Entidades.detalleNCE)
        Dim listaDatosAdicional As List(Of Entidades.infoAdicionalNCE)
        Dim listaTotalesConImpuestos As List(Of Entidades.totalConImpuestosNCE)
        Dim listaPagos As List(Of Entidades.pagosFE)
        Dim listaDatosAdicionalDetalle As List(Of Entidades.detallesAdicionalesNCE)
        Dim listaImpuestos As List(Of Entidades.impuestosNCE)

        listaDetalle = New List(Of Entidades.detalleNCE)
        listaDatosAdicional = New List(Of Entidades.infoAdicionalNCE)
        listaTotalesConImpuestos = New List(Of Entidades.totalConImpuestosNCE)
        listaPagos = New List(Of Entidades.pagosFE)

        Try
            Dim SP As String = ""

            If Functions.VariablesGlobales._vgGuardarLog = "Y" Then
                oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"Tipo de Nota de Crédito = {TipoNC}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
                oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"Consultando Nota de Crédito con # DocEntry = {DocEntry}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If

            'Utilitario.Util_Log.Escribir_Log("SP: " + SP.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ANTES A CONSULTAR", "ManejoDeDocumentos")

            If TipoNC = "NCE" Then
                SP = DesencriptarQuery_.GetQueryConsulta(Documentos.tipoDocumento.NotaCredito, DocEntry)
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

                'ds3 = dbManager.EjecutarSP(SPs(3).ToString(), DocEntry)
                'dt3 = ds3.Tables(0).Copy
                'dt3.TableName = "FormaPago"
                'ds.Tables.Add(dt3)
            Else
                ds = dbManager.EjecutarSP(SP, DocEntry)
            End If

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not dbManager.ValidarCamposNulos(ds, "2", _camponulo) Then Return Nothing
            End If

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("INGRESANDO A CONSULTAR", "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oNotaCredito = New Entidades.RequestNotaCredito
                oNotaCredito.infoTributaria = New Entidades.infoTributariaNCE()
                oNotaCredito.infoNotaCredito = New Entidades.infoNotaCreditoNCE()

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try
                            For Each r As DataRow In ds.Tables(0).Rows

                                oNotaCredito.infoTributaria.ambiente = r("Ambiente")

                                Dim claveAcceso As String = r("ClaveAcceso").ToString()
                                If Not String.IsNullOrEmpty(claveAcceso) AndAlso claveAcceso.Length = 49 Then
                                    oNotaCredito.infoTributaria.claveAcceso = claveAcceso
                                End If

                                oNotaCredito.infoTributaria.razonSocial = r("RazonSocial")

                                oNotaCredito.infoTributaria.nombreComercial = r("NombreComercial")

                                oNotaCredito.infoTributaria.ruc = r("RUC")

                                oNotaCredito.infoTributaria.tipoEmision = r("TipoEmision")

                                oNotaCredito.infoTributaria.codDoc = r("CodigoDocumento")

                                oNotaCredito.infoTributaria.estab = r("Establecimiento")

                                oNotaCredito.infoTributaria.ptoEmi = r("PuntoEmision")

                                oNotaCredito.infoTributaria.secuencial = r("SecuencialDocumento")
                                If Not oNotaCredito.infoTributaria.secuencial.ToString().Length.Equals("9") Then oNotaCredito.infoTributaria.secuencial = oNotaCredito.infoTributaria.secuencial.ToString().PadLeft(9, "0")
                                Utilitario.Util_Log.Escribir_Log("oNotaCredito.Secuencial : " & oNotaCredito.infoTributaria.secuencial.ToString(), "ManejoDeDocumentos")

                                oNotaCredito.infoTributaria.dirMatriz = r("DireccionMatriz")

                                oNotaCredito.infoTributaria.diaEmission = CDate(r("FechaEmision")).ToString("dd")

                                oNotaCredito.infoTributaria.mesEmission = CDate(r("FechaEmision")).ToString("MM")

                                oNotaCredito.infoTributaria.anioEmission = CDate(r("FechaEmision")).ToString("yyyy")

                                'Comienza estructura infoNotaCredito
                                Try
                                    'El servicio de facturación exige la fecha en formato dd/MM/yyyy
                                    oNotaCredito.infoNotaCredito.fechaEmision = CDate(r("FechaEmision")).ToString("dd/MM/yyyy")
                                    Utilitario.Util_Log.Escribir_Log("oNotaCredito.FechaEmision : " & CDate(r("FechaEmision")).ToString("dd/MM/yyyy"), "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("oNotaCredito.FechaEmision : " & ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                oNotaCredito.infoNotaCredito.dirEstablecimiento = r("DireccionEstablecimiento")
                                oNotaCredito.infoNotaCredito.tipoIdentificacionComprador = r("TipoIdentificadorComprador")
                                oNotaCredito.infoNotaCredito.razonSocialComprador = r("RazonSocialComprador")
                                oNotaCredito.infoNotaCredito.identificacionComprador = r("IdentificacionComprador")
                                oNotaCredito.infoNotaCredito.contribuyenteEspecial = r("ContribuyenteEspecial")
                                oNotaCredito.infoNotaCredito.obligadoContabilidad = r("ObligadoContabilidad")
                                oNotaCredito.infoNotaCredito.rise = r("Rise")
                                oNotaCredito.infoNotaCredito.codDocModificado = r("codDocModificado")
                                oNotaCredito.infoNotaCredito.numDocModificado = r("numDocModificado")
                                Try
                                    'El servicio de facturación exige la fecha en formato dd/MM/yyyy
                                    oNotaCredito.infoNotaCredito.fechaEmisionDocSustento = CDate(r("FechaEmisionDocModificado")).ToString("dd/MM/yyyy")
                                    Utilitario.Util_Log.Escribir_Log("oNotaCredito.FechaEmision : " & CDate(r("FechaEmisionDocModificado")).ToString("dd/MM/yyyy"), "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("oNotaCredito.FechaEmision : " & ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                oNotaCredito.infoNotaCredito.totalSinImpuestos = FormatearNumero(r("TotalSinImpuesto"))
                                oNotaCredito.infoNotaCredito.valorModificacion = FormatearNumero(r("ValorModificacion"))
                                oNotaCredito.infoNotaCredito.moneda = r("Moneda").ToString
                                oNotaCredito.infoNotaCredito.motivo = r("Motivo").ToString

                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo8")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base8"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva8"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo12")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base12"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva12"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo13")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base13"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva13"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo0")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impfaIVA.baseImponible = FormatearNumero(r("Base0"))
                                    impfaIVA.valor = FormatearNumero(r("ValorIva0"))
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("CodigoNoi")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("BaseNoi"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIvaNoi"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("CodigoExen")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("BaseExen"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIvaExen"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseIce") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("CodigoIce")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("BaseIce"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIvaIce"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base5") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo5")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje5")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("Base5"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIva5"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base15") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo15")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje15")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("Base15"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIva15"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base14") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo14")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje14")
                                    impfaNOIVA.baseImponible = FormatearNumero(r("Base14"))
                                    impfaNOIVA.valor = FormatearNumero(r("ValorIva14"))
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                Utilitario.Util_Log.Escribir_Log("Termina cabecera ", "ManejoDeDocumentos")

                                oNotaCredito.infoNotaCredito.totalConImpuestos = listaTotalesConImpuestos
                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleFactura As Entidades.detalleNCE = New Entidades.detalleNCE

                                itemDetalleFactura.codigoInterno = r("CodigoPrincipal").ToString

                                itemDetalleFactura.codigoAdicional = r("CodigoAuxiliar").ToString

                                itemDetalleFactura.descripcion = r("Descripcion").ToString

                                itemDetalleFactura.cantidad = CInt(r("Cantidad"))

                                itemDetalleFactura.precioUnitario = FormatearNumero(r("PrecioUnitario"))

                                itemDetalleFactura.descuento = FormatearNumero(r("Descuento"))

                                itemDetalleFactura.precioTotalSinImpuesto = FormatearNumero(r("PrecioTotalSinImpuesto"))

                                listaDatosAdicionalDetalle = New List(Of Entidades.detallesAdicionalesNCE)

                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As Entidades.detallesAdicionalesNCE = New Entidades.detallesAdicionalesNCE
                                    itemDetalleDatoAdicional.nombre = r("ConceptoAdicional1").ToString
                                    itemDetalleDatoAdicional.valor = r("NombreAdicional1").ToString
                                    listaDatosAdicionalDetalle.Add(itemDetalleDatoAdicional)
                                End If

                                If Not r("ConceptoAdicional2") = "0" Then
                                    Dim itemDetalleDatoAdicional2 As Entidades.detallesAdicionalesNCE = New Entidades.detallesAdicionalesNCE
                                    itemDetalleDatoAdicional2.nombre = r("ConceptoAdicional2")
                                    itemDetalleDatoAdicional2.valor = r("NombreAdicional2")
                                    listaDatosAdicionalDetalle.Add(itemDetalleDatoAdicional2)
                                End If

                                If Not r("ConceptoAdicional3") = "0" Then
                                    Dim itemDetalleDatoAdicional3 As Entidades.detallesAdicionalesNCE = New Entidades.detallesAdicionalesNCE
                                    itemDetalleDatoAdicional3.nombre = r("ConceptoAdicional3")
                                    itemDetalleDatoAdicional3.valor = r("NombreAdicional3")
                                    listaDatosAdicionalDetalle.Add(itemDetalleDatoAdicional3)
                                End If

                                itemDetalleFactura.detallesAdicionales = listaDatosAdicionalDetalle

                                listaImpuestos = New List(Of Entidades.impuestosNCE)

                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA8" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA13" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("CodigoIce").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentajeIce").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponibleIce"))
                                    impuesto.valor = FormatearNumero(r("TotalIvaIce"))
                                    impuesto.tarifa = FormatearNumero(r("TarifaIce"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA5" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA15" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = FormatearNumero(r("BaseImponible"))
                                    impuesto.valor = FormatearNumero(r("TotalIva"))
                                    impuesto.tarifa = FormatearNumero(r("Tarifa"))
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA14" Then ' 12%
                                    Dim impuesto As Entidades.impuestosNCE = New Entidades.impuestosNCE
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
                            oNotaCredito.detalles = listaDetalle
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalFac As Entidades.infoAdicionalNCE = New Entidades.infoAdicionalNCE
                                itemDatoAdicionalFac.nombre = r("Concepto")
                                itemDatoAdicionalFac.valor = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalFac)
                            Next
                            Utilitario.Util_Log.Escribir_Log("Termina info adicional ", "ManejoDeDocumentos")
                            oNotaCredito.infoAdicional = listaDatosAdicional
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
                            'oNotaCredito.infoNotaCredito.pagos = listaPagos
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    End If

                Next
            End If

            Return oNotaCredito
            Utilitario.Util_Log.Escribir_Log("FACTURA CONSULTADA", "ManejoDeDocumentos")

        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"ArgumentException-Ocurrio un error al consultar datos de la factura en la Base, DocEntry: {DocEntry} Descr: {x.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"ArgumentException-Error al Consultar Factura con # DocEntry = {DocEntry}, Descr: {x.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"Ocurrio un error al consultar datos de la factura en la Base, DocEntry: {DocEntry} Descr: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"Error al Consultar Factura con # DocEntry = {DocEntry}, Descr: {ex.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
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

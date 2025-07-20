'Imports Entidades
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
'https
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Functions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SAPbobsCOM
Imports System.Globalization
Imports Spire.Pdf
Imports Spire.Pdf.AutomaticFields
Imports Spire.Pdf.Graphics

Public Class ManejoDeDocumentoSolsap
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    'Private usrBD As String = "usredoc"
    'Private pswBD As String = "usredoc"
    'Private pswBD_HANA As String = "B1Icesa$"
    ''' OBSERVACION, ESTE USUARIO Y CLAVE LO DEBE TOMAR DESDE LA TABLA DE CONFIGURACION QUE DEBE SER UN UDO,
    ''' NO HABRÍA PROBLEMA YA QUE LA CONSULTA LA HARÍA POR LA DIAPI,
    ''' YA QUE ESTE USUARIO Y CLAVE LO USA SOLO PARA EJECUTAR LOS QUERY
    ''' OJO CON EL SERVICIO.

    Private _EstadoAutorizacion As String = ""
    Private _ClaveAcceso As String = ""
    Private _Observacion As String = ""
    Private _CampoNulo As String = ""
    Private _NumAutorizacion As String = ""
    Private _FechaAutorizacion As Date
    Private _EstadoSAP As String = ""
    Private _Error As String = ""
    Dim mensaje As String = ""
    Dim oObjeto As Object
    'Dim ObjetoRespuesta As Object = Nothing
    Dim oFuncionesAddon As FuncionesAddon

    Dim oFuncionesB1 As FuncionesB1

    Dim _tipoManejo As String
    Dim _errorMensajeWSEnvío As String
    Public _Nombre_Proveedor_SAP_BO As String = ""

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential

    Dim oDocumento As SAPbobsCOM.Documents

    Dim _GuardarLog As String = "N"

    Private _NumeroDeDocumentoSRI As String = ""

    Dim mensajeDocAut As String = ""

    ''' <summary>
    ''' Tipo Manejo, A - Addon, S -  Servicio
    ''' </summary>
    ''' <param name="Company"></param>
    ''' <param name="sboApp"></param>
    ''' <param name="tipoManejo"></param>
    ''' <remarks>Tipo Manejo, A - Addon, S -  Servicio</remarks>
    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application, tipoManejo As String, ByVal ProveedorSAPBO As String)
        'Utilitario.Util_Log.Escribir_Log("SubNew Inicio", "ManejoDeDocumentos")
        rCompany = Company
        _tipoManejo = tipoManejo
        _Nombre_Proveedor_SAP_BO = ProveedorSAPBO
        If tipoManejo = "A" Then
            rsboApp = sboApp
            oFuncionesAddon = New Functions.FuncionesAddon(rCompany, rsboApp, True, False)
            oFuncionesB1 = New Functions.FuncionesB1(rCompany, rsboApp, True, False)
        Else
            ' SI ES SERVICIO INSTANCIO ESTA CLASE, YA QUE NO USA LA UIAPI
            oFuncionesAddon = New Functions.FuncionesAddon(rCompany, rsboApp, True, False)
        End If
    End Sub

#Region "Consulta de Documentos"
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
                SP = GetQueryConsulta(tipoDocumento.FacturaAnticipo, DocEntry)
            Else
                SP = GetQueryConsulta(tipoDocumento.Factura, DocEntry)
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

                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Detalles"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "InfoAdicionales"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "FormaPago"
                ds.Tables.Add(dt3)
            Else
                ds = EjecutarSP(SP, DocEntry)
            End If

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "2") Then Return Nothing
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

                                oFactura.infoTributaria.claveAcceso = r("ClaveAcceso")

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

    Public Function ProcesaEnvioDocumento(DocEntry As Integer, TipoDocumento As String, Optional ByVal sincronizado As Boolean = False) As String

        Try
            Dim result As Boolean = False
            Dim objetoRespuesta As Object = Nothing
            Dim TipoWS As String = "LOCAL"

            Dim BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo

            Dim sSQL As String = ""

            If _tipoManejo = "S" Then
                TipoWS = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
            Else
                TipoWS = Functions.VariablesGlobales._TipoWS
            End If

            Utilitario.Util_Log.Escribir_Log("TIPO WEB SERVICES: " + TipoWS, "ManejoDeDocumentos")
            'Se escribe el log

            If sincronizado = True Then

            Else

                If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Seteando informacion a enviar..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" Then
                    oObjeto = ConsultarFactura(TipoDocumento, DocEntry)
                ElseIf TipoDocumento = "NCE" Then

                ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
                End If

                If Not oObjeto Is Nothing Then

                    Try
                        If Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = "Y" Then
                            _NumeroDeDocumentoSRI = ""
                            _NumeroDeDocumentoSRI = oObjeto.infoTributaria.estab + "-" + oObjeto.infoTributaria.ptoEmi + "-" + oObjeto.infoTributaria.secuencial
                            Utilitario.Util_Log.Escribir_Log("NumeroDeDocumentoSRI: " + _NumeroDeDocumentoSRI.ToString(), "ManejoDeDocumentos")
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error al setear NumeroDeDocumentoSRI: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    Utilitario.Util_Log.Escribir_Log("Enviando documento al SRI, por favor espere..!!", "ManejoDeDocumentos")

                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("Enviando documento al SRI, por favor espere..!!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Envíando Documento al SRI", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
                    End If

                    Utilitario.Util_Log.Escribir_Log($"Enviando documento al SRI, TipoDocumento: {TipoDocumento} DocEntry: {DocEntry} TipoWs: {TipoWS}", "ManejoDeDocumentos")

                    Dim respuesta_WS As String = ""
                    If Functions.VariablesGlobales._ActApiSS = "Y" AndAlso (TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE") Then
                        objetoRespuesta = EnviarFacturaSolsap(DirectCast(oObjeto, Entidades.RequestFactura))
                    Else

                    End If

                    If Not objetoRespuesta Is Nothing Then
                        Dim mensajesSRI As String = ""

                        If Functions.VariablesGlobales._ActApiSS = "Y" AndAlso TypeOf objetoRespuesta Is Entidades.ResponseDocuments Then
                            Dim respDoc As Entidades.ResponseDocuments = CType(objetoRespuesta, Entidades.ResponseDocuments)
                            _EstadoAutorizacion = respDoc.type
                            _Observacion = respDoc.msg
                            _ClaveAcceso = respDoc.claveAcceso
                        Else

                        End If

                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Respuesta del SRI: " + _EstadoAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                        If _tipoManejo = "A" Then
                            ' Seteo el Error recibido del servicio EDOC
                            rsboApp.SetStatusBarMessage("Recibiendo respuesta..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            'Dim respuestaaa = objetoRespuesta.Autorizaciones(0).Mensajes(0).mensaje1().ToString & "- " & objetoRespuesta.Autorizaciones(0).Mensajes(0).informacionAdicional().ToString
                        End If

                        If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" And (TipoWS = "NUBE_4_1" And Functions.VariablesGlobales._ActApiSS = "Y") Then
                            _Observacion = recorreErrorFactura_Solsap(CType(objetoRespuesta, Entidades.ResponseDocuments), DocEntry.ToString())
                        End If

                        If _tipoManejo = "A" Then
                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Observación del SRI: " + _Observacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If
                        'oBackgroundWorker.ReportProgress(70)

                        If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                            Try
                                _NumAutorizacion = objetoRespuesta.autorizaciones(0).numeroAutorizacion().ToString()
                                _FechaAutorizacion = objetoRespuesta.autorizaciones(0).FechaAutorizacion()
                            Catch ex As Exception

                            End Try
                        Else
                            _NumAutorizacion = "0000000000"
                            Try
                                mensajesSRI = objetoRespuesta.ErrorEDOC
                            Catch ex As Exception
                                mensajesSRI = " No se recibio la descripcion del Error "
                            End Try
                        End If

                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("Grabando respuesta de SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Respuesta SRI en Documento - " + TipoDocumento + " - DocEntry " + DocEntry.ToString() + " - # de Autorización - " + _NumAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If

                        _Observacion = String.Format("Estado:{0} - # AUTORIZACION {1} - RespuestaSRI - {2} - Error - {3} ", _EstadoAutorizacion.ToString, _NumAutorizacion.ToString, mensajesSRI, _Observacion.ToString)

                        ' Mando a Grabar a SAP
                        If TipoDocumento = "LQE" Then

                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                        ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then

                        ElseIf TipoDocumento = "SSGR" Then

                        Else
                            result = GrabaDatosAutorizacion(DocEntry, TipoDocumento)
                        End If

                        If result Then
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Proceso terminado con exito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            End If

                        Else
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Ocurrio un Error al Guardar los datos de Autorización..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                        End If
                    Else
                        ' controlo error si no pude consumir el servicio del SRI
                        ' NO SE RECIBIO RESPUESTA DEL WEB SERVICE DE EDOC - ENVÍA FACTURA

                        _Observacion = "No se ha recibido respuesta del documento " + DocEntry.ToString() + " - Resp WS :" + respuesta_WS
                        _Error = respuesta_WS
                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("No se recibio respuesta inmediata del SRI, el documento será procesado nuevamente en 2 minutos, o use la opcion REENVIAR SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "No se recibió respuesta de los Web Services", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        End If

                        Try
                            If TipoDocumento = "LQE" Then

                            ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                            ElseIf TipoDocumento = "SSGR" Then

                            Else
                                GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _Error)
                            End If

                        Catch ex As Exception
                        End Try

                    End If
                Else
                    ' Controlo Error si no se seteo la factura con los datos de base 
                    _Observacion = "Ocurrio un error al Consultar los datos de la Factura: " & DocEntry.ToString() & " " & _CampoNulo
                    _Error = _Observacion
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la factura en la Base, DocEntry:  " & DocEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Ocurrio un error al consultar datos de la factura en la Base, DocEntry: " & DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If

                    Try
                        If TipoDocumento = "LQE" Then

                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then

                        ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then

                        ElseIf TipoDocumento = "SSGR" Then

                        Else
                            GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _Error)
                        End If
                    Catch ex As Exception
                    End Try

                End If

            End If

            Return _Observacion

        Catch ex As Exception
            _Error = ex.Message
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error:  " & ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return _Error + _errorMensajeWSEnvío
        End Try
    End Function

    Public Function GrabaDatosAutorizacion(DocEntry As Integer, TipoDocumento As String) As Boolean
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
                        oDocumento.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_SYP_FECAUTOC").Value = Date.Now
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_SYP_FECAUTOC DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString + " Fecha y Hora Autorización " + _FechaAutorizacion.ToString
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

    Public Function GrabaDatosAutorizacion_Error(DocEntry As Integer, TipoDocumento As String, MsgError As String) As Boolean
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
    Public Function GrabaDatosAutorizacion_TablaTM(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        Dim DocEntryUdoRet As String = ""

        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo DocEntry del UDO retencion " + DocEntryUdoRet.ToString, "ManejoDeDocumentos")
            DocEntryUdoRet = oFuncionesB1.getRSvalue("select T1.""DocEntry"" FROM ""OPCH"" T0 INNER JOIN ""@TM_LE_RETCH"" T1 ON T0.""U_TM_CRNUM""= T1.""DocEntry"" WHERE T0.""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocEntry", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo DocEntry del UDO retencion : " + DocEntryUdoRet.ToString, "ManejoDeDocumentos")
        End If

        If TipoDocumento = "LQE" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
            CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_TipoDoc""='03' and ""U_TM_DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "Code", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
        ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
            CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_TipoDoc""='07' and ""U_TM_DocEntry"" = '" + DocEntryUdoRet.ToString() + "' ", "Code", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
        Else
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
            CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "Code", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla TM_DOC_ELEC: " + CODE.ToString, "ManejoDeDocumentos")
        End If

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

                oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @TM_DOC_ELEC: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompany.UserTables.Item("TM_DOC_ELEC")
                oUserTable.GetByKey(CODE)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If

                If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacion.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "A"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_Observacion.ToString, 254)
                End If
                If _EstadoAutorizacion.ToString().Equals("5") Or _EstadoAutorizacion.ToString().Equals("EN PROCESO SRI") Or _EstadoAutorizacion.ToString().Equals("7") Or _EstadoAutorizacion.ToString().Equals("ERROR EN RECEPCION") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacion.ToString
                    'oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "P"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_Observacion.ToString, 254)
                End If
                If _EstadoAutorizacion.ToString().Equals("4") Or _EstadoAutorizacion.ToString().Equals("ERROR AL FIRMAR") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacion.ToString
                    'oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "P"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_Observacion.ToString, 254)
                End If
                If _EstadoAutorizacion.ToString().Equals("3") Or _EstadoAutorizacion.ToString().Equals("NO AUTORIZADA") Or _EstadoAutorizacion.ToString().Equals("6") Or _EstadoAutorizacion.ToString().Equals("DEVUELTA") Then
                    oUserTable.UserFields.Fields.Item("U_TM_NroAutorizacion").Value = _NumAutorizacion.ToString
                    'oUserTable.UserFields.Fields.Item("U_TM_FechaAutorizacion").Value = Date.Now.ToString
                    oUserTable.UserFields.Fields.Item("U_TM_Status").Value = "R"
                    oUserTable.UserFields.Fields.Item("U_TM_Motivo").Value = Left(_Observacion.ToString, 254)
                End If
                RetVal = oUserTable.Update()
                If RetVal <> 0 Then
                    'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                Else
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_DOC_ELEC: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                End If
                Return True
            Else
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla Control Doc. Electrónico", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla TM_DOC_ELEC" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla TM_DOC_ELEC: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try
        Return result
    End Function

    Public Function recorreErrorFactura_Solsap(ByVal respuesta As Entidades.ResponseDocuments, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta Is Nothing Then
            Return mensaje
        End If

        estado = If(respuesta.type, "")

        If estado = "AUTORIZADO" Or estado = "2" Then
            mensaje = "Estado: AUTORIZADO"
            If Not String.IsNullOrEmpty(respuesta.msg) Then
                mensaje &= ", " & respuesta.msg
            End If
        Else
            mensaje = "Estado: " & estado
            If Not String.IsNullOrEmpty(respuesta.msg) Then
                mensaje &= " - " & respuesta.msg
            End If
            If respuesta.log IsNot Nothing AndAlso respuesta.log.Count > 0 Then
                mensaje &= " - Detalle: " & String.Join(" | ", respuesta.log)
            End If
        End If

        mensaje &= " - NÚMERO DEL DOCUMENTO: " & codigoDocumento
        Return mensaje
    End Function

    Public Function ConsultaParametro(ByVal Modulo As String, ByVal Tipo As String, ByVal Subtipo As String, ByVal Nombre As String) As String
        Try
            Dim valor As String = ""
            Dim sQueryPrefijo As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryPrefijo = "SELECT A.""U_Valor"" "
                sQueryPrefijo += "FROM ""@GS_CONFD"" A INNER JOIN "
                sQueryPrefijo += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                sQueryPrefijo += " WHERE  B.""U_Modulo"" = '" + Modulo + "' AND B.""U_Tipo"" = '" + Tipo + "' "
                sQueryPrefijo += " AND B.""U_Subtipo"" = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.""U_Nombre"" = '" + Nombre + "'"
            Else
                sQueryPrefijo = "SELECT A.U_Valor "
                sQueryPrefijo += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                sQueryPrefijo += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                sQueryPrefijo += " WHERE B.U_Modulo = '" + Modulo + "' AND  B.U_Tipo = '" + Tipo + "' "
                sQueryPrefijo += " AND B.U_Subtipo = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.U_Nombre = '" + Nombre + "'"
            End If

            valor = oFuncionesAddon.getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

#Region "Funciones ADO SQL"

    Public Function EjecutarSP(SP As String, docentry As Integer) As DataSet

        Dim ds As New DataSet

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            ' ds = ObtenerColeccion("CALL " + rCompany.CompanyDB + "." + SP + " ('" + docentry.ToString() + "')", False)
            Utilitario.Util_Log.Escribir_Log("Query Consulta : " & SP, "ManejoDeDocumentos")
            ds = ObtenerColeccion(SP, False)
        Else
            Try
                Utilitario.Util_Log.Escribir_Log("Query Consulta : " & SP, "ManejoDeDocumentos")

                Using Cn As SqlConnection = GetSqlConnectionBase()
                    Using cm As New SqlCommand(SP, Cn)
                        Cn.Open()
                        cm.CommandType = CommandType.Text
                        ' cm.Parameters.Add("@DocKey", SqlDbType.Int).Value = docentry

                        Dim da As New SqlDataAdapter
                        ' da.ReturnProviderSpecificTypes = True

                        da.SelectCommand = cm
                        da.Fill(ds)

                    End Using
                End Using
            Catch ex As Exception
                rsboApp.SetStatusBarMessage("Ejecutar SP: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Utilitario.Util_Log.Escribir_Log("Catch Query Consulta : " & ex.Message().ToString(), "ManejoDeDocumentos")
                Return Nothing
            End Try
        End If

        Return ds

    End Function

    ''' <summary>
    ''' Obtiene Cadena de Conexión
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSqlConnectionBase() As SqlConnection
        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        Dim cnBaseSAP As New SqlConnection
        Try

            If _tipoManejo <> "A" Then
                BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")
            Else
                BD_User = Functions.VariablesGlobales._vgUserBD
            End If

            If BD_User = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            If _tipoManejo <> "A" Then
                BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")
            Else
                BD_Pass = Functions.VariablesGlobales._vgPassBD
            End If

            If BD_Pass = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            Dim cadena As New SqlConnectionStringBuilder

            If _tipoManejo = "A" Then

                If Not String.IsNullOrEmpty(Functions.VariablesGlobales._vgServerNode) Then
                    cadena.DataSource = Functions.VariablesGlobales._vgServerNode ' "S00SQL" 'rCompany.Server '
                    cadena.InitialCatalog = rCompany.CompanyDB
                    cadena.UserID = Functions.VariablesGlobales._vgUserBD
                    cadena.Password = Functions.VariablesGlobales._vgPassBD
                Else
                    cadena.DataSource = rCompany.Server ' "S00SQL" 'rCompany.Server '
                    cadena.InitialCatalog = rCompany.CompanyDB
                    cadena.UserID = Functions.VariablesGlobales._vgUserBD
                    cadena.Password = Functions.VariablesGlobales._vgPassBD
                End If
            Else
                cadena.DataSource = rCompany.Server ' "S00SQL" 'rCompany.Server '
                cadena.InitialCatalog = rCompany.CompanyDB
                cadena.UserID = BD_User
                cadena.Password = BD_Pass
                Utilitario.Util_Log.Escribir_Log("datos conexion sql User: " + BD_User + " Pass: " + BD_Pass + " tipo: " + _tipoManejo, "ManejoDeDocumentos")
            End If

            cnBaseSAP.ConnectionString = cadena.ConnectionString
            Return cnBaseSAP

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("error GetSqlConnectionBase: " + ex.Message.ToString + " User: " + BD_User + " Pass: " + BD_Pass + " tipo: " + _tipoManejo, "ManejoDeDocumentos")

            Return Nothing
        End Try

#Disable Warning BC42105 ' La función 'GetSqlConnectionBase' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'GetSqlConnectionBase' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

#End Region

#Region "FUNCIONES HANA"
    Public CONEXION As Odbc.OdbcConnection

    Public Function ConectaHANA(Optional ByRef mensaje As String = "") As Boolean
        Dim ConexionHana As String = String.Empty

        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        Dim _ServerNode As String = ""
        If _tipoManejo = "S" Then
            _ServerNode = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "ServerNode")
            If String.IsNullOrEmpty(_ServerNode) Then
                _ServerNode = rCompany.Server
            End If
            BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")
            BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")
            Utilitario.Util_Log.Escribir_Log("_ServerNode: " + _ServerNode.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("BD_User: " + BD_User.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("BD_Pass: " + BD_Pass.ToString(), "ManejoDeDocumentos")
        End If


        Try


            If _tipoManejo <> "A" Then


            Else
                BD_User = Functions.VariablesGlobales._vgUserBD
            End If
            'BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")

            If BD_User = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If


            If _tipoManejo <> "A" Then


            Else
                BD_Pass = Functions.VariablesGlobales._vgPassBD
            End If
            'BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")

            If BD_Pass = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                If (IntPtr.Size = 8) Then
                    ConexionHana = String.Concat(ConexionHana, "Driver={HDBODBC};")
                Else
                    ConexionHana = String.Concat(ConexionHana, "Driver={HDBODBC32};")
                End If
                If _tipoManejo = "A" Then
                    If Not String.IsNullOrEmpty(Functions.VariablesGlobales._vgServerNode) Then
                        ConexionHana = String.Concat(ConexionHana, "ServerNode=", Functions.VariablesGlobales._vgServerNode & ";")
                        ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._vgUserBD, ";")
                        ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._vgPassBD, ";")
                    Else
                        ConexionHana = String.Concat(ConexionHana, "ServerNode=", rCompany.Server & ";")
                        ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._vgUserBD, ";")
                        ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._vgPassBD, ";")
                    End If
                Else

                    ConexionHana = String.Concat(ConexionHana, "ServerNode=", _ServerNode & ";")
                    ConexionHana = String.Concat(ConexionHana, "UID=", BD_User, ";")
                    ConexionHana = String.Concat(ConexionHana, "PWD=", BD_Pass, ";")


                End If


                'pswBD_HANA

                CONEXION = New Odbc.OdbcConnection(ConexionHana)

                If CONEXION.State = ConnectionState.Closed Then
                    CONEXION.Open()
                End If
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If

                Return True

                'Else
                '    'CONEXION = New Odbc.OdbcConnection("DRIVER={SQL Server Native Client 10.0}; Server= " & serv & "; Database=" & bd & "; Uid=" & userdb & "; Pwd=" & passdb)
                '    CONEXION = New Odbc.OdbcConnection("DRIVER={" + _driversql + "}; Server= " & serv & "; Database=" & bd & "; Uid=" & userdb & "; Pwd=" & passdb)
                '    'CONEXION = New Odbc.OdbcConnection(GetSqlConnectionBaseString())
                '    If CONEXION.State = ConnectionState.Closed Then
                '        CONEXION.Open()
                '    End If
                '    If CONEXION.State = ConnectionState.Open Then
                '        CONEXION.Close()
                '    End If

                '    Return True

            End If
            Return False

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ConexionHana: " + ConexionHana.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Conecta_HANA: " + ex.Message, "ManejoDeDocumentos")
            Return False

        End Try

#Disable Warning BC42353 ' La función 'ConectaHANA' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    End Function
#Enable Warning BC42353 ' La función 'ConectaHANA' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?

    Public Function ObtenerValor(ByVal Consulta As String, Optional ByVal KeepOpen As Boolean = False, Optional ByRef mensaje As String = "") As String

        Try
            If Consulta = String.Empty Then Return ""

            If CONEXION.State = ConnectionState.Closed Then
                CONEXION.Open()
            End If

            Dim Comando As New Odbc.OdbcCommand(Consulta, CONEXION)
            Comando.CommandTimeout = 0
            Comando.CommandText = Consulta

            Dim Valor As String = IIf(IsDBNull(Comando.ExecuteScalar), "", Comando.ExecuteScalar)
            If Valor Is Nothing Then Valor = ""

            If Not KeepOpen Then
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If
            End If

            REM Retornar el valor.
            Return Valor

        Catch ex As Odbc.OdbcException
            addLogTxt("ObtenerValor: " + ex.Message)
            Return ""

        End Try
    End Function

    Public Function ObtenerColeccion(ByVal Consulta As String, Optional ByVal KeepOpen As Boolean = False) As DataSet

        Dim ds As New DataSet
        Try
            If Consulta = String.Empty Then Return Nothing

            ConectaHANA()

            If CONEXION.State = ConnectionState.Closed Then
                CONEXION.Open()
            End If

            Dim DapTable As New Odbc.OdbcDataAdapter(Consulta, CONEXION)
            DapTable.SelectCommand.CommandTimeout = 0
            DapTable.Fill(ds)

            If Not KeepOpen Then
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If
            End If
            Return ds

        Catch ex As Odbc.OdbcException
            Utilitario.Util_Log.Escribir_Log("ObtenerColeccion: " + ex.Message + " QUERY: " + Consulta.ToString(), "ManejoDeDocumentos")
            Return Nothing
        End Try

    End Function

#End Region

#Region "LOG"
#Disable Warning BC42307 ' El parámetro de comentario XML 'Contenido' no coincide con un parámetro de la instrucción 'function' correspondiente.
#Disable Warning BC42307 ' El parámetro de comentario XML 'FileName' no coincide con un parámetro de la instrucción 'function' correspondiente.
#Disable Warning BC42307 ' El parámetro de comentario XML 'oRuta' no coincide con un parámetro de la instrucción 'function' correspondiente.
    ''' <summary>
    ''' Agrega una línea al archivo txt del log.
    ''' </summary>
    ''' <param name="Contenido">Contenido de la línea de texto</param>
    ''' <param name="FileName">Nombre del archivo an el que se registra el log (sin extensión .txt)</param>
    ''' <param name="oRuta">Ruta en la que se guardará el archivo (Ejemplo: C:\Logs)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function addLogTxt(ByVal texto As String) As Boolean
#Enable Warning BC42307 ' El parámetro de comentario XML 'oRuta' no coincide con un parámetro de la instrucción 'function' correspondiente.
#Enable Warning BC42307 ' El parámetro de comentario XML 'FileName' no coincide con un parámetro de la instrucción 'function' correspondiente.
#Enable Warning BC42307 ' El parámetro de comentario XML 'Contenido' no coincide con un parámetro de la instrucción 'function' correspondiente.

        Dim sRuta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "MyLog.txt"
        If Not File.Exists(sRuta) Then
            Dim strStreamW As Stream = Nothing
            Dim strStreamWriter As StreamWriter = Nothing

            strStreamW = File.Create(sRuta) ' lo creamos
            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default) '
            strStreamWriter.Close() ' cerramos
        End If

        Dim sTexto As New StringBuilder

        sTexto.AppendLine("FECHA: " & Now)
        sTexto.AppendLine("----------------------------------------------------------")
        sTexto.AppendLine(texto.ToString())

        Try
            Dim oTextWriter As TextWriter = New StreamWriter(sRuta, True)
            oTextWriter.WriteLine(sTexto.ToString)
            oTextWriter.Close()
            oTextWriter.Flush()
            oTextWriter = Nothing

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

#End Region

    Public Sub SetProtocolosdeSeguridad()
        'PARA TLS 1.2
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999

        'PARA HTTPS
        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
    End Sub
#Region "Funciones Complementarias para funcion WS Sincronizacion"

    Private Function ObtnerTipoDocumentoEDOC(ByVal tipoDocumento As String) As String

        If tipoDocumento = "FCE" Or tipoDocumento = "FRE" Or tipoDocumento = "FAE" Then
            Return "1"
        ElseIf tipoDocumento = "NDE" Then
            Return "4"
        ElseIf tipoDocumento = "NCE" Then
            Return "3"
        ElseIf tipoDocumento = "TRE" Or tipoDocumento = "TLE" Or tipoDocumento = "GRE" Or tipoDocumento = "SSGR" Then
            Return "5"
        ElseIf tipoDocumento = "REE" Or tipoDocumento = "REA" Or tipoDocumento = "RER" Then
            Return "2"
        ElseIf tipoDocumento = "LQE" Then
            Return "6"
        End If

        Return ""

    End Function

    Private Function Get_company_numdoc_by_proveedor(ByVal nombreProveedor As String, ByVal DocEnty As String, ByVal tipoDocumento As String) As String()

        Dim tabla_SAP As String = ""
        Dim ruc_numdoc() As String = {"", ""}

        If tipoDocumento = "FCE" Or tipoDocumento = "FRE" Or tipoDocumento = "NDE" Then
            tabla_SAP = "OINV"
        ElseIf tipoDocumento = "FAE" Then
            tabla_SAP = "ODPI"
        ElseIf tipoDocumento = "NCE" Then
            tabla_SAP = "ORIN"
        ElseIf tipoDocumento = "TRE" Then
            tabla_SAP = "OWTR"
        ElseIf tipoDocumento = "GRE" Then
            tabla_SAP = "ODLN"
        ElseIf tipoDocumento = "TLE" Then
            tabla_SAP = "OWTQ"
        ElseIf tipoDocumento = "REE" Or tipoDocumento = "REA" Or tipoDocumento = "RER" Or tipoDocumento = "RDM" Or tipoDocumento = "LQE" Then
            tabla_SAP = "OPCH"
        ElseIf tipoDocumento = "SSGR" Then
            tabla_SAP = """@SS_GRCAB"""
        End If

        'obtener informacion de los textbox

        Dim querySincro As String = ""

        If _tipoManejo = "A" Then

            If tabla_SAP = "OPCH" Then
                If tipoDocumento = "LQE" Then
                    querySincro = Functions.VariablesGlobales._SINCRO_LQE
                Else
                    querySincro = Functions.VariablesGlobales._SINCRO_RT
                End If
            ElseIf tabla_SAP = """@SS_GRCAB""" Then
                querySincro = Functions.VariablesGlobales._SINCRO_GRUDO
            Else

                querySincro = Functions.VariablesGlobales._SINCRO_DOC

            End If

        Else
            If tabla_SAP = "OPCH" Then
                If tipoDocumento = "LQE" Then
                    querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SINCRO_LQE")
                Else
                    querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SINCRO_RET")
                End If
            ElseIf tabla_SAP = """@SS_GRCAB""" Then
                querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "QueryGRUdo")

            Else

                querySincro = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SINCRO_DOC")

            End If

        End If




        'hacer un replace de tabla actual por lo que esta en la plantilla

        querySincro = querySincro.Replace("TABLA", tabla_SAP)
        querySincro = querySincro.Replace("IDENTIFICADOR", DocEnty)

        Utilitario.Util_Log.Escribir_Log("QUERY SINCRO: " + querySincro.ToString + " Tipo Doc:" + tipoDocumento, "FuncionesB1")

        Try

            'Realizo Consulta
            Dim dir_est As String = "", dir_pe As String = "", secuencial As String = "", ruc_compania As String = ""
            Dim numeroDOC As String = ""


            Dim r As SAPbobsCOM.Recordset = oFuncionesAddon.getRecordSet(querySincro)

            If r.RecordCount > 0 Then

                dir_est = oFuncionesAddon.nzString(r.Fields.Item("A").Value)
                dir_pe = oFuncionesAddon.nzString(r.Fields.Item("B").Value)
                secuencial = oFuncionesAddon.nzString(r.Fields.Item("C").Value)
                ruc_compania = oFuncionesAddon.nzString(r.Fields.Item("R").Value)

                If Not secuencial.Length = 9 Then
                    secuencial = secuencial.PadLeft(9, "0")
                End If

                numeroDOC = dir_est & "-" & dir_pe & "-" & secuencial

                If numeroDOC.Length = 17 And String.IsNullOrEmpty(ruc_compania) = False Then

                    ruc_numdoc(0) = ruc_compania
                    ruc_numdoc(1) = numeroDOC

                    Return ruc_numdoc
                End If

            End If


        Catch ex As Exception

        End Try



        Return ruc_numdoc

    End Function

#End Region

    Shared Function customCertValidation(ByVal sender As Object,
                                             ByVal cert As X509Certificate,
                                             ByVal chain As X509Chain,
                                             ByVal errors As SslPolicyErrors) As Boolean
        Return True
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

    Public Sub ReleaseGRHEISON(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Release Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
    End Sub

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

    Private Function ValidarCamposNulos(dataset As DataSet, numTabla As String) As Boolean

        Try
            Dim nombretabla = "Table" & numTabla
            'Dim DescripcionConcepto As String = ""
            'Dim ListaInforAdicional As New List(Of String)
            Dim concepto As String = Nothing
            Dim descripcion As String = Nothing
            For Each table As DataTable In dataset.Tables

                If table.TableName.ToString() = nombretabla Then

                    For rowIndex As Integer = 0 To table.Rows.Count - 1
                        Dim row As DataRow = table.Rows(rowIndex)
                        For columnIndex As Integer = 0 To table.Columns.Count - 1
                            Dim currentColumn As DataColumn = table.Columns(columnIndex)
                            'InforAdicional(rowIndex, columnIndex) = row(currentColumn)
                            If currentColumn.ColumnName = "Descripcion" Then
                                'descripcion = row(currentColumn)
                                If IsDBNull(row(currentColumn)) Then
                                    descripcion = "Nulo"
                                End If
                            End If
                            If currentColumn.ColumnName = "Concepto" Then
                                concepto = row(currentColumn)
                            End If



                        Next
                        'ListaInforAdicional.Add(descripcion & "|" & concepto)
                        'DescripcionConcepto = ""
                        If descripcion = "Nulo" Then
                            _CampoNulo = concepto.ToString & " se encuentra en nulo, por favor validar"
                            rsboApp.SetStatusBarMessage(table.TableName.ToString() & " Comcepto: " & concepto.ToString & " se encuentra en nulo, por favor validar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return False
                        End If

                    Next

                    'For Each lista In ListaInforAdicional
                    '    If String.IsNullOrEmpty(lista.Split("|")(0)) Then
                    '        rsboApp.SetStatusBarMessage(table.TableName.ToString() & " Comcepto: " & lista.Split("|")(1).ToString & " se encuentra en nulo, por favor validar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    '        Return False
                    '    End If
                    'Next

                Else
                    For Each row As DataRow In table.Rows
                        For Each column As DataColumn In table.Columns
                            If IsDBNull(row(column)) Then
                                _CampoNulo = column.ColumnName.ToString & " se encuentra en nulo, por favor validar"
                                rsboApp.SetStatusBarMessage(table.TableName.ToString() & " Columna: " & column.ColumnName & " se encuentra en nulo, por favor validar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Return False
                            End If

                        Next
                        'Console.WriteLine()
                    Next

                End If



            Next
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error en funcion Validar Campos Nulos " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try

        Return True
    End Function

    Public Function AbrirEnlaceExterno(enlace As String) As Boolean

        Try


            If Not String.IsNullOrEmpty(enlace) Then

                Dim rn As New System.Diagnostics.Process
                rn.StartInfo.FileName = enlace

                rn.Start()
                rn.Dispose()

                Return True
            End If


        Catch ex As Exception

        End Try

        Return False
    End Function

    Public Function ExisterchivoLocal(ruta As String) As Boolean

        Try
            If File.Exists(ruta) Then

                Dim Proc As New Process()
                Proc.StartInfo.FileName = ruta
                Proc.Start()
                Proc.Dispose()

                Return True

            End If

            Utilitario.Util_Log.Escribir_Log("Archivo no encontrado " & ruta, "frmDocumento")

        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("Error al Abrir PDf integrado por XML " & ex.Message, "frmDocumento")

        End Try


        Return False

    End Function

    Enum tipoDocumento
        Factura
        NotaCredito
        NotaDebito
        FacturaAnticipo
        UDO
        GuiaRemisionEntrega
        GuiaRemisionTraslado
        GuiaRemisionSolicitudTraslado
        Retencion
        RetencionNotaDebito
        RetencionAnticipo
        Liquidacion
        GuiaRemisionDesatendida
    End Enum

    Public Const sKey As String = "S01s7p1" ' CLAVE DE ENCRIPTACION LICENCIA
    Dim QueryDesencriptado As String = ""
    Private Function GetQueryConsulta(ByVal tipodoc As tipoDocumento, ByVal docentry As Integer, Optional ByVal Seccion As String = "") As String

        Try

            Dim partesQuerys As New List(Of String)
            Select Case Seccion

                Case "EXPORT"
                    partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_CompleExportacion.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                Case "REEMBOLSO"
                    partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_CompleReembolso.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                Case "DOCSENV"
                    partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_DocumentosEnviados.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))


                Case Else

                    If tipodoc = tipoDocumento.Factura Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.FacturaAnticipo Then

                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaAnticipoSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaAnticipoSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.NotaCredito Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaCreditoSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaCreditoSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.NotaDebito Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaDebitoSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaDebitoSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf (tipodoc = tipoDocumento.GuiaRemisionEntrega) Or (tipodoc = tipoDocumento.GuiaRemisionTraslado) Or (tipodoc = tipoDocumento.GuiaRemisionSolicitudTraslado) Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiaRemisionSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiaRemisionSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf (tipodoc = tipoDocumento.Retencion) Or (tipodoc = tipoDocumento.RetencionNotaDebito) Or (tipodoc = tipoDocumento.RetencionAnticipo) Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_RetencionSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_RetencionSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))


                    ElseIf tipodoc = tipoDocumento.Liquidacion Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_LiquidacionSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_LiquidacionSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.GuiaRemisionDesatendida Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiasDesatendidas01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiasDesatendidas02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))



                    End If


            End Select


            'Procesamiento de Querys
            QueryDesencriptado = ""

            For Each querysession As String In partesQuerys

                QueryDesencriptado = QueryDesencriptado + querysession

            Next

            'remplazamos 3 parametros
            Dim midata As String = QueryDesencriptado

            Select Case tipodoc
                Case tipoDocumento.Factura
                   ' midata = midata.Replace("A.""Docentry""=@DocKey", "A.""Docentry""=@DocKey AND A.""DocSubType"" <> 'DN'")
                Case tipoDocumento.NotaCredito
                   ' midata = midata.Replace("INV", "RIN")
                Case tipoDocumento.NotaDebito
                   ' midata = midata.Replace("A.""Docentry""=@DocKey", "A.""Docentry""=@DocKey AND A.""DocSubType"" = 'DN'")
                Case tipoDocumento.Retencion
                   ' midata = midata.Replace("INV", "PCH")
                Case tipoDocumento.FacturaAnticipo
                    midata = midata.Replace("INV", "DPI")
                Case tipoDocumento.RetencionAnticipo
                    midata = midata.Replace("PCH", "DPO")
                Case tipoDocumento.RetencionNotaDebito
                   ' midata = midata.Replace("PCH", "DPO")

                Case tipoDocumento.GuiaRemisionEntrega
                    'midata = midata.Replace("INV", "DLN")
                Case tipoDocumento.GuiaRemisionTraslado
                    midata = midata.Replace("DLN", "WTR")
                Case tipoDocumento.GuiaRemisionSolicitudTraslado
                    midata = midata.Replace("DLN", "WTQ")
            End Select



            'EL DOcentry
            midata = midata.Replace("@DocKey", docentry.ToString)
            midata = midata.Replace("@TipoDoc", "'" + tipodoc.ToString() + "'")
            midata = midata.Replace("@GS_SS_NAMEDB", rCompany.CompanyDB)

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then


                midata = HanaTablesSapReplace(midata)


            End If



            Return midata

        Catch ex As Exception

            Return "GSCODEEXCEPCION :" & ex.Message


        End Try



    End Function

    Private Function HanaTablesSapReplace(s As String) As String

        Dim MYSP As String = ""

        MYSP = s.Replace("""OINV""", rCompany.CompanyDB + ".""OINV""")
        MYSP = MYSP.Replace("""OCRD""", rCompany.CompanyDB + ".""OCRD""")
        MYSP = MYSP.Replace("""OITM""", rCompany.CompanyDB + ".""OITM""")
        MYSP = MYSP.Replace("""OPLN""", rCompany.CompanyDB + ".""OPLN""")
        MYSP = MYSP.Replace("""OUSR""", rCompany.CompanyDB + ".""OUSR""")
        MYSP = MYSP.Replace("""OEXD""", rCompany.CompanyDB + ".""OEXD""")
        MYSP = MYSP.Replace("""OSLP""", rCompany.CompanyDB + ".""OSLP""")
        MYSP = MYSP.Replace("""OCTG""", rCompany.CompanyDB + ".""OCTG""")
        MYSP = MYSP.Replace("""OCRN""", rCompany.CompanyDB + ".""OCRN""")
        MYSP = MYSP.Replace("""ORTT""", rCompany.CompanyDB + ".""ORTT""")
        MYSP = MYSP.Replace("""OIBT""", rCompany.CompanyDB + ".""OIBT""")
        MYSP = MYSP.Replace("""OITB""", rCompany.CompanyDB + ".""OITB""")
        MYSP = MYSP.Replace("""ORIN""", rCompany.CompanyDB + ".""ORIN""")
        MYSP = MYSP.Replace("""ODLN""", rCompany.CompanyDB + ".""ODLN""")
        MYSP = MYSP.Replace("""OWHT""", rCompany.CompanyDB + ".""OWHT""")
        MYSP = MYSP.Replace("""OWHS""", rCompany.CompanyDB + ".""OWHS""")
        MYSP = MYSP.Replace("""OCRG""", rCompany.CompanyDB + ".""OCRG""")
        MYSP = MYSP.Replace("""OPCH""", rCompany.CompanyDB + ".""OPCH""")
        MYSP = MYSP.Replace("""CUFD""", rCompany.CompanyDB + ".""CUFD""")
        MYSP = MYSP.Replace("""OSTA""", rCompany.CompanyDB + ".""OSTA""")
        MYSP = MYSP.Replace("""OWTR""", rCompany.CompanyDB + ".""OWTR""")
        MYSP = MYSP.Replace("""OITW""", rCompany.CompanyDB + ".""OITW""")
        MYSP = MYSP.Replace("""OCRY""", rCompany.CompanyDB + ".""OCRY""")
        MYSP = MYSP.Replace("""NNM1""", rCompany.CompanyDB + ".""NNM1""")
        MYSP = MYSP.Replace("""ODPI""", rCompany.CompanyDB + ".""ODPI""")
        MYSP = MYSP.Replace("""OADM""", rCompany.CompanyDB + ".""OADM""")
        MYSP = MYSP.Replace("""ODPO""", rCompany.CompanyDB + ".""ODPO""")
        MYSP = MYSP.Replace("""OCPR""", rCompany.CompanyDB + ".""OCPR""")
        MYSP = MYSP.Replace("""OBTN""", rCompany.CompanyDB + ".""OBTN""")
        MYSP = MYSP.Replace("""OITL""", rCompany.CompanyDB + ".""OITL""")

        'Logica para que dependiendo de una Opcion del Addon replace tablas que no se encuentren
        If Functions.VariablesGlobales._TablasNativasReplace <> "" Then

            Dim rtablas = Functions.VariablesGlobales._TablasNativasReplace.Split(";")

            For Each t In rtablas

                MYSP = MYSP.Replace($"""{t}""", rCompany.CompanyDB + $".""{t}""")

            Next

        End If

        'Remplazando sub tablas
        For i As Integer = 1 To 12

            MYSP = MYSP.Replace("""INV" & i.ToString & """", rCompany.CompanyDB + ".""INV" & i.ToString & """")
            MYSP = MYSP.Replace("""CRD" & i.ToString & """", rCompany.CompanyDB + ".""CRD" & i.ToString & """")
            MYSP = MYSP.Replace("""ITM" & i.ToString & """", rCompany.CompanyDB + ".""ITM" & i.ToString & """")
            MYSP = MYSP.Replace("""PLN" & i.ToString & """", rCompany.CompanyDB + ".""PLN" & i.ToString & """")
            MYSP = MYSP.Replace("""USR" & i.ToString & """", rCompany.CompanyDB + ".""USR" & i.ToString & """")
            MYSP = MYSP.Replace("""EXD" & i.ToString & """", rCompany.CompanyDB + ".""EXD" & i.ToString & """")
            MYSP = MYSP.Replace("""SLP" & i.ToString & """", rCompany.CompanyDB + ".""SLP" & i.ToString & """")
            MYSP = MYSP.Replace("""CTG" & i.ToString & """", rCompany.CompanyDB + ".""CTG" & i.ToString & """")
            MYSP = MYSP.Replace("""CRN" & i.ToString & """", rCompany.CompanyDB + ".""CRN" & i.ToString & """")
            MYSP = MYSP.Replace("""RTT" & i.ToString & """", rCompany.CompanyDB + ".""RTT" & i.ToString & """")
            MYSP = MYSP.Replace("""IBT" & i.ToString & """", rCompany.CompanyDB + ".""IBT" & i.ToString & """")
            MYSP = MYSP.Replace("""ITB" & i.ToString & """", rCompany.CompanyDB + ".""ITB" & i.ToString & """")
            MYSP = MYSP.Replace("""RIN" & i.ToString & """", rCompany.CompanyDB + ".""RIN" & i.ToString & """")
            MYSP = MYSP.Replace("""DLN" & i.ToString & """", rCompany.CompanyDB + ".""DLN" & i.ToString & """")
            MYSP = MYSP.Replace("""WHT" & i.ToString & """", rCompany.CompanyDB + ".""WHT" & i.ToString & """")
            MYSP = MYSP.Replace("""WHS" & i.ToString & """", rCompany.CompanyDB + ".""WHS" & i.ToString & """")
            MYSP = MYSP.Replace("""CRG" & i.ToString & """", rCompany.CompanyDB + ".""CRG" & i.ToString & """")
            MYSP = MYSP.Replace("""PCH" & i.ToString & """", rCompany.CompanyDB + ".""PCH" & i.ToString & """")
            MYSP = MYSP.Replace("""WTR" & i.ToString & """", rCompany.CompanyDB + ".""WTR" & i.ToString & """")
            MYSP = MYSP.Replace("""DPI" & i.ToString & """", rCompany.CompanyDB + ".""DPI" & i.ToString & """")
            MYSP = MYSP.Replace("""ADM" & i.ToString & """", rCompany.CompanyDB + ".""ADM" & i.ToString & """")
            MYSP = MYSP.Replace("""DPO" & i.ToString & """", rCompany.CompanyDB + ".""DPO" & i.ToString & """")
            MYSP = MYSP.Replace("""CPR" & i.ToString & """", rCompany.CompanyDB + ".""CPR" & i.ToString & """")
            MYSP = MYSP.Replace("""BTN" & i.ToString & """", rCompany.CompanyDB + ".""BTN" & i.ToString & """")
            MYSP = MYSP.Replace("""ITL" & i.ToString & """", rCompany.CompanyDB + ".""ITL" & i.ToString & """")


            'Logica para que dependiendo de una Opcion del Addon replace tablas que no se encuentren
            If Functions.VariablesGlobales._TablasNativasReplace <> "" Then

                Dim rtablas = Functions.VariablesGlobales._TablasNativasReplace.Split(";")

                For Each t In rtablas

                    MYSP = MYSP.Replace($"""{t.Substring(1)}" & i.ToString & """", rCompany.CompanyDB + $".""{t.Substring(1)}" & i.ToString & """")

                Next

            End If

        Next

        'tablas de usuario
        MYSP = MYSP.Replace("""@", String.Format("""{0}"".""@", rCompany.CompanyDB))


        Return MYSP

    End Function

    'CONSUMO DE API SOLSAP

    Public Function ObtenerTokenAutenticacion() As String
        Try
            Dim usuario As String = Functions.VariablesGlobales._ApiAutUser
            Dim password As String = Functions.VariablesGlobales._ApiAutPw
            Dim endpoint As String = Functions.VariablesGlobales._ApiAutSS

            If String.IsNullOrEmpty(usuario) OrElse String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(endpoint) Then
                If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Faltan datos de autenticación (usuario, clave o endpoint)", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return Nothing
            End If

            Dim jsonBody As String = $"{{""usuario"":""{usuario}"", ""password"":""{password}""}}"
            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonBody)
            End Using

            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Using reader As New StreamReader(response.GetResponseStream())
                Dim result As String = reader.ReadToEnd()
                Dim json As JObject = JObject.Parse(result)
                Dim token As String = json("token")?.ToString()

                If Not String.IsNullOrEmpty(token) Then
                    If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Autenticación exitosa", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Return token
                Else
                    If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("No se recibió token de autenticación", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return Nothing
                End If
            End Using

        Catch ex As Exception
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error al autenticar: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try
    End Function

    Public Function EnviarFacturaSolsap(factura As Entidades.RequestFactura) As Entidades.ResponseDocuments
        Try

            ActivarTLS()
            Dim token As String = ObtenerTokenAutenticacion()
            If String.IsNullOrEmpty(token) Then Return Nothing

            Dim endpoint As String = Functions.VariablesGlobales._ApiFactEmiSS
            If String.IsNullOrEmpty(endpoint) Then Return Nothing

            'comentar esta linea posterior a la correcion del endpoint de adicionales en el detalle
            If factura.detalles IsNot Nothing Then
                For Each det As Entidades.detalleFE In factura.detalles
                    det.detallesAdicionales = Nothing
                Next
            End If

            'Dim jsonBody As String = JsonConvert.SerializeObject(factura)
            Dim settings As New JsonSerializerSettings()
            settings.NullValueHandling = NullValueHandling.Ignore
            Dim jsonBody As String = JsonConvert.SerializeObject(factura, settings)

            Try
                Dim sRutaCarpeta As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\LOG_SAED\"
                Dim Secuencial = Right("000000000" & factura.infoTributaria.secuencial, 9)
                Dim sRuta As String = sRutaCarpeta & factura.infoTributaria.estab & factura.infoTributaria.ptoEmi & Secuencial & ".xml"
                'Dim sRuta As String = sRutaCarpeta & factura.infoTributaria.secuencial.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    Dim xml As XmlDocument = JsonConvert.DeserializeXmlNode(jsonBody, "factura")
                    xml.Save(sRuta)
                    Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Headers.Add("Authorization", $"Bearer {token}")

            Using sw As New StreamWriter(request.GetRequestStream())
                sw.Write(jsonBody)
            End Using

            Using resp As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(resp.GetResponseStream())
                    Dim result As String = reader.ReadToEnd()
                    Dim json As JObject = JObject.Parse(result)
                    Dim estado As String = json.SelectToken("data.result.estado")?.ToString()
                    Dim mensaje As String = json.SelectToken("data.result.mensaje")?.ToString()
                    Dim clave As String = json.SelectToken("data.result.claveAcceso")?.ToString()
                    Dim identificador As String = json.SelectToken("data.result.identificador")?.ToString()

                    Dim respuesta As New Entidades.ResponseDocuments()
                    respuesta.type = estado
                    respuesta.msg = mensaje
                    respuesta.claveAcceso = clave
                    respuesta.identificador = identificador
                    Return respuesta
                End Using
            End Using

        Catch webEx As WebException
            Dim mensajeError As String = webEx.Message
            Dim respErr As HttpWebResponse = TryCast(webEx.Response, HttpWebResponse)
            If respErr IsNot Nothing Then
                Using reader As New StreamReader(respErr.GetResponseStream())
                    mensajeError &= " - " & reader.ReadToEnd()
                End Using
            End If
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error enviando factura: " & mensajeError, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing

        Catch ex As Exception
            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Error enviando factura: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try
    End Function

    Public Sub ActivarTLS()
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or 768 Or 3072
    End Sub

    Private Function FormatearNumero(valor As Object) As String
        If valor Is Nothing Then Return Nothing

        Dim numero As Decimal
        If Decimal.TryParse(valor.ToString(), numero) Then
            Return numero.ToString(CultureInfo.InvariantCulture)
        End If

        Return valor.ToString()
    End Function
End Class

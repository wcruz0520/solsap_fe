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
Imports System.Xml.Serialization
Imports Functions
Imports Newtonsoft.Json.Linq
Imports SAPbobsCOM
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

            If TipoFactura = "FAE" Then
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "GS_SAP_FE_ObtenerFacturadeVentaAnticipo "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "GS_SAP_FE_ONE_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "GS_SAP_FE_HEI_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "GS_SAP_FE_SYP_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "GS_SAP_FE_TM_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "GS_SAP_FE_SS_ObtenerFacturadeVentaAnticipo "
                End If
            Else
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "GS_SAP_FE_ObtenerFacturadeVenta_4_3 "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "GS_SAP_FE_ONE_OBTENERFACTURADEVENTA_4_3 "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "GS_SAP_FE_HEI_OBTENERFACTURADEVENTA_4_3 "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "GS_SAP_FE_SYP_OBTENERFACTURADEVENTA_4_3 "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "GS_SAP_FE_TM_OBTENERFACTURADEVENTA_4_3 "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "GS_SAP_FE_SS_ObtenerFacturadeVenta_4_3 "
                End If
            End If

            If Functions.VariablesGlobales._vgGuardarLog = "Y" Then
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, $"Tipo de factura = {TipoFactura}, SP: {SP}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, $"Consultando Factura con # DocEntry = {DocEntry}, SP: {SP}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If

            Utilitario.Util_Log.Escribir_Log("SP: " + SP.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ANTES A CONSULTAR", "ManejoDeDocumentos")

            If TipoFactura = "FAE" Then
                SP = GetQueryConsulta(tipoDocumento.FacturaAnticipo, DocEntry)
            Else
                SP = GetQueryConsulta(tipoDocumento.Factura, DocEntry)
            End If

            Utilitario.Util_Log.Escribir_Log("Query Desencriptado " & SP.ToString(), "ManejoDeDocumentos")

            If SP.Contains("GSCODEEXCEPCION") Then
                Utilitario.Util_Log.Escribir_Log("EXCEPCION DETECTADA EN EL PROCESO DE OBTENER STRING QUERY - " & SP, "ManejoDeDocumentos")
                rsboApp.StatusBar.SetText(VariablesGlobales._gNombreAddOn + " - Ocurrio Un Error Favor falidar el Archivo de Log y Buscar el Codigo GSCODEEXCEPCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                                    oFactura.infoFactura.fechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")
                                    Utilitario.Util_Log.Escribir_Log("oFactura.FechaEmision : " + CDate(r("FechaEmision")).ToString("yyyy-MM-dd"), "ManejoDeDocumentos")
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

                                oFactura.infoFactura.totalSinImpuestos = r("TotalSinImpuesto").ToString

                                oFactura.infoFactura.totalDescuento = r("TotalDescuento").ToString

                                oFactura.infoFactura.propina = r("Propina").ToString

                                oFactura.infoFactura.importeTotal = r("ImporteTotal").ToString

                                oFactura.infoFactura.moneda = r("Moneda").ToString

                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo8")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje8")
                                    impfaIVA.baseImponible = r("Base8")
                                    impfaIVA.valor = r("ValorIva8")
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo12")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje12")
                                    impfaIVA.baseImponible = r("Base12")
                                    impfaIVA.valor = r("ValorIva12")
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo13")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje13")
                                    impfaIVA.baseImponible = r("Base13")
                                    impfaIVA.valor = r("ValorIva13")
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaIVA.codigo = r("Codigo0")
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje0")
                                    impfaIVA.baseImponible = r("Base0")
                                    impfaIVA.valor = r("ValorIva0")
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoNoi")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impfaNOIVA.baseImponible = r("BaseNoi")
                                    impfaNOIVA.valor = r("ValorIvaNoi")
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoExen")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen")
                                    impfaNOIVA.baseImponible = r("BaseExen")
                                    impfaNOIVA.valor = r("ValorIvaExen")
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseIce") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("CodigoIce")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    impfaNOIVA.baseImponible = r("BaseIce")
                                    impfaNOIVA.valor = r("ValorIvaIce")
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base5") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo5")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje5")
                                    impfaNOIVA.baseImponible = r("Base5")
                                    impfaNOIVA.valor = r("ValorIva5")
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base15") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo15")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje15")
                                    impfaNOIVA.baseImponible = r("Base15")
                                    impfaNOIVA.valor = r("ValorIva15")
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base14") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosFE = New Entidades.totalConImpuestosFE
                                    impfaNOIVA.codigo = r("Codigo14")
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje14")
                                    impfaNOIVA.baseImponible = r("Base14")
                                    impfaNOIVA.valor = r("ValorIva14")
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

                                itemDetalleFactura.precioUnitario = r("PrecioUnitario").ToString

                                itemDetalleFactura.descuento = r("Descuento").ToString

                                itemDetalleFactura.precioTotalSinImpuesto = r("PrecioTotalSinImpuesto").ToString

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
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA8" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA13" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("CodigoIce").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentajeIce").ToString
                                    impuesto.baseImponible = r("BaseImponibleIce").ToString
                                    impuesto.valor = r("TotalIvaIce").ToString
                                    impuesto.tarifa = r("TarifaIce").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA5" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
                                    listaImpuestos.Add(impuesto)
                                End If

                                If r("TaxCodeAp") = "IVA15" Then ' 12%
                                    Dim impuesto As Entidades.impuestosFE = New Entidades.impuestosFE
                                    impuesto.codigo = r("Codigo").ToString
                                    impuesto.codigoPorcentaje = r("CodigoPorcentaje").ToString
                                    impuesto.baseImponible = r("BaseImponible").ToString
                                    impuesto.valor = r("TotalIva").ToString
                                    impuesto.tarifa = r("Tarifa").ToString
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
                                Pago.total = r("Total").ToString
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

                    '' FACTURA DE EXPORTACIÓN
                    'If oFactura.Tipo = 1 Then  oFactura = ConsultaFacturaExportacion(TipoFactura, TipoWS, DocEntry, oFactura)

                    '' FACTURA DE REEMBOLSO
                    'If oFactura.Tipo = 2 Then  oFactura = ConsultaFacturaReembolso(TipoFactura, TipoWS, DocEntry, oFactura)
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
    Public Function ConsultarFactura_NUBE_4_3(ByVal TipoFactura As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object
        Dim oFactura As Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFactura = Nothing
        Dim listaDetalle As List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDetalleFactura)

        Dim listaFacCompensacion As List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTCompensacion)

        Dim listaDatosAdicional As List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalFactura)
        Dim FormasdePago As List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaPagos)

        'If TipoWS = "NUBE_4_1" Then
        listaDetalle = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDetalleFactura)
        listaFacCompensacion = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTCompensacion)
        listaDatosAdicional = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalFactura)
        FormasdePago = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaPagos)
        'End If

        Dim aplicadoDescuentoAdicional As Boolean = False

        Try
            Dim SP As String = ""

            If TipoFactura = "FAE" Then
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "GS_SAP_FE_ObtenerFacturadeVentaAnticipo "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "GS_SAP_FE_ONE_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "GS_SAP_FE_HEI_OBTENERFACTURADEVENTAANTICIPO "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "GS_SAP_FE_SYP_OBTENERFACTURADEVENTAANTICIPO "
                End If

                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Consultando Factura con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            Else
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "GS_SAP_FE_ObtenerFacturadeVenta4_1 "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "GS_SAP_FE_ONE_OBTENERFACTURADEVENTA "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "GS_SAP_FE_HEI_OBTENERFACTURADEVENTA "
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "GS_SAP_FE_SYP_OBTENERFACTURADEVENTA "
                End If

                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Consultando Factura con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oFactura = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFactura

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try

                            For Each r As DataRow In ds.Tables(0).Rows

                                ' MANEJO DE FACTURAS DE EXPORTACION Y REEMBOLSO - 2018-02-18
                                ' Indica que tipo de factura es (0.- Normal, 1.- Exportadores, 2.- Reembolsos)
                                Try
                                    If r("TipoFactura").ToString() = "" Then
                                        oFactura.Tipo = 0
                                    Else
                                        oFactura.Tipo = r("TipoFactura")
                                    End If
                                    Utilitario.Util_Log.Escribir_Log(" (0.- Normal, 1.- Exportadores, 2.- Reembolsos)", "ManejoDeDocumentos")
                                    Utilitario.Util_Log.Escribir_Log("Tipo Factura : " & oFactura.Tipo.ToString(), "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                                    oFactura.Tipo = 0
                                End Try

                                ' OFFLINE 14 NOVIEMBRE 2017
                                'FAMC 18/02/2019
                                If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                    oFactura.ClaveAcceso = Nothing
                                Else
                                    oFactura.ClaveAcceso = r("ClaveAcceso")
                                End If

                                oFactura.Ambiente = r("Ambiente")
                                oFactura.TipoEmision = r("TipoEmision")
                                oFactura.RazonSocial = r("RazonSocial")
                                oFactura.NombreComercial = r("NombreComercial")

                                oFactura.Ruc = r("RUC")
                                'oFactura.Ruc = "0992737964001"
                                oFactura.CodigoDocumento = r("CodigoDocumento")
                                oFactura.Establecimiento = r("Establecimiento")
                                oFactura.PuntoEmision = r("PuntoEmision")
                                oFactura.Secuencial = r("SecuencialDocumento")
                                If Not oFactura.Secuencial.ToString().Length.Equals("9") Then
                                    oFactura.Secuencial = oFactura.Secuencial.ToString().PadLeft(9, "0")
                                End If
                                Utilitario.Util_Log.Escribir_Log("oFactura.Secuencial : " & oFactura.Secuencial.ToString(), "ManejoDeDocumentos")
                                oFactura.DireccionMatriz = r("DireccionMatriz")
                                oFactura.FechaEmision = r("FechaEmision")
                                oFactura.DireccionEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oFactura.ContribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oFactura.ContribuyenteEspecial = Nothing
                                End If

                                oFactura.ObligadoContabilidad = r("ObligadoContabilidad")

                                'exportacion
                                If r("Tipo") = "1" Then
                                    oFactura.ComercioExterior = r("ComercioExterior")
                                    oFactura.IncoTermFactura = r("IncoTermFactura")
                                    oFactura.LugarIncoTerm = r("LugarIncoTerm")
                                    oFactura.PaisOrigen = r("PaisOrigen")
                                    oFactura.PaisAdquisicion = r("PaisAdquisicion")
                                    oFactura.IncoTermTotalSinImpuestos = r("IncoTermTotalSinImpuestos")
                                    oFactura.FleteInternacional = r("FleteInternacional")
                                    oFactura.SeguroInternacional = r("SeguroInternacional")
                                    oFactura.GastosAduaneros = r("GastosAduaneros")
                                    oFactura.GastosTransporteOtros = r("GastosTransporteOtros")
                                End If
                                'fin exportacion

                                'REEMBOLSO
                                If r("Tipo") = "2" Then
                                    oFactura.CodDocReemb = r("CodDocReemb")
                                    oFactura.TotalComprobantesReembolso = r("TotalComprobantesReembolso")
                                    oFactura.TotalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                                    oFactura.TotalImpuestoReembolso = r("TotalImpuestoReembolso")
                                End If
                                'FIN REEMBOLSO

                                'guia remision
                                If r("Tipo") = "5" Then
                                    oFactura.GuiaRemision = r("GuiaRemision")
                                    oFactura.DireccionPartida = r("DireccionPartida")
                                    oFactura.DireccionDestinatario = r("DireccionDestinatario")
                                    oFactura.FechaInicioTransporte = r("FechaInicioTransporte")
                                    oFactura.FechaFinTransporte = r("FechaFinTransporte")
                                    oFactura.RazonSocialTransportista = r("RazonSocialTransportista")
                                    oFactura.TipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                                    oFactura.IdentificacionTransportista = r("IdentificacionTransportista")
                                End If
                                'fin guia

                                oFactura.TipoIdentificacionComprador = r("TipoIdentificadorComprador")



                                oFactura.RazonSocialComprador = r("RazonSocialComprador")
                                oFactura.IdentificacionComprador = r("IdentificacionComprador")

                                Try
                                    If Not r("DirComprador") = "" Then
                                        oFactura.DirComprador = r("DirComprador")
                                    End If
                                Catch ex As Exception
                                End Try

                                oFactura.TotalSinImpuesto = r("TotalSinImpuesto")
                                oFactura.TotalDescuento = r("TotalDescuento")

                                oFactura.Propina = r("Propina")
                                oFactura.ImporteTotal = r("ImporteTotal")
                                oFactura.Moneda = r("Moneda")
                                oFactura.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                                oFactura.EmailResponsable = r("EmailResponsable")
                                oFactura.SecuencialERP = r("SecuencialERP")
                                oFactura.CodigoTransaccionERP = r("CodigoTransaccionERP")
                                oFactura.Estado = r("Estado")
                                oFactura.FechaCarga = r("FechaCarga")
                                oFactura.Campo1 = r("Campo1")
                                oFactura.Campo2 = r("Campo2")
                                oFactura.Campo3 = r("Campo3")

                                'IMPUESTO FACTURA
                                'Impuestos totalizados en la factura.
                                Dim lstimpfact As Object
                                lstimpfact = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaImpuesto)

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As Object
                                    impfaIVA = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaImpuesto

                                    impfaIVA.Codigo = "2"
                                    impfaIVA.CodigoPorcentaje = "2"
                                    impfaIVA.Tarifa = "12"
                                    impfaIVA.BaseImponible = r("Base12")
                                    impfaIVA.Valor = r("ImpuestoTotal")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impfaIVA.DescuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpfact.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As Object
                                    impfaIVA = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaImpuesto

                                    impfaIVA.Codigo = "2"
                                    impfaIVA.CodigoPorcentaje = "3"
                                    impfaIVA.Tarifa = "14"
                                    impfaIVA.BaseImponible = r("Base13")
                                    impfaIVA.Valor = r("ImpuestoTotal")
                                    If aplicadoDescuentoAdicional = False Then
                                        If r("DescuentoAdicional") <> "0" Then
                                            impfaIVA.DescuentoAdicional = r("DescuentoAdicional")
                                            aplicadoDescuentoAdicional = True
                                        End If
                                    End If


                                    lstimpfact.Add(impfaIVA)
                                End If
                                If r("Base0") <> 0 Then

                                    Dim impfaNOIVA As Object
                                    impfaNOIVA = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaImpuesto

                                    impfaNOIVA.Codigo = "2"
                                    impfaNOIVA.CodigoPorcentaje = "0"
                                    impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.BaseImponible = r("Base0")
                                    impfaNOIVA.Valor = 0
                                    If aplicadoDescuentoAdicional = False Then
                                        If r("DescuentoAdicional") <> "0" Then
                                            impfaNOIVA.DescuentoAdicional = r("DescuentoAdicional")
                                            aplicadoDescuentoAdicional = True
                                        End If
                                    End If

                                    lstimpfact.Add(impfaNOIVA)
                                End If

                                If r("BaseExo") <> 0 Then

                                    Dim impfaNOIVA As Object
                                    impfaNOIVA = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaImpuesto

                                    impfaNOIVA.Codigo = "2"
                                    impfaNOIVA.CodigoPorcentaje = "6"
                                    impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.BaseImponible = r("BaseExo")
                                    impfaNOIVA.Valor = 0
                                    If aplicadoDescuentoAdicional = False Then
                                        If r("DescuentoAdicional") <> "0" Then
                                            impfaNOIVA.DescuentoAdicional = r("DescuentoAdicional")
                                            aplicadoDescuentoAdicional = True
                                        End If
                                    End If

                                    lstimpfact.Add(impfaNOIVA)
                                End If
                                oFactura.ENTFacturaImpuesto = lstimpfact.ToArray
                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleFactura As Object
                                itemDetalleFactura = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDetalleFactura

                                itemDetalleFactura.CodigoPrincipal = r("CodigoPrincipal")
                                itemDetalleFactura.CodigoAuxiliar = r("CodigoAuxiliar")
                                itemDetalleFactura.Descripcion = r("Descripcion")
                                itemDetalleFactura.Cantidad = r("Cantidad")
                                itemDetalleFactura.PrecioUnitario = r("PrecioUnitario")
                                itemDetalleFactura.Descuento = r("Descuento")
                                itemDetalleFactura.PrecioTotalSinImpuesto = r("PrecioTotalSinImpuesto")

                                ''Datos adicionales de cada detalle del item                                     
                                Dim listaDetalleDatoAdicional As Object
                                listaDetalleDatoAdicional = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalDetalleFactura)
                                'Adicional1
                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As Object
                                    itemDetalleDatoAdicional = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalDetalleFactura
                                    itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                    itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                                End If

                                'Adicional2
                                If Not r("ConceptoAdicional2") = "0" Then
                                    Dim itemDetalleDatoAdicional2 As Object

                                    itemDetalleDatoAdicional2 = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalDetalleFactura
                                    itemDetalleDatoAdicional2.Nombre = r("ConceptoAdicional2")
                                    itemDetalleDatoAdicional2.Descripcion = r("NombreAdicional2")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional2)
                                End If

                                'Adicional3
                                If Not r("ConceptoAdicional3") = "0" Then
                                    Dim itemDetalleDatoAdicional3 As Object
                                    itemDetalleDatoAdicional3 = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalDetalleFactura
                                    itemDetalleDatoAdicional3.Nombre = r("ConceptoAdicional3")
                                    itemDetalleDatoAdicional3.Descripcion = r("NombreAdicional3")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional3)
                                End If

                                itemDetalleFactura.ENTDatoAdicionalDetalleFactura = listaDetalleDatoAdicional.ToArray

                                'IMPUESTOS DEL DETALLE
                                'Puede Tener IVA y/0 ICE
                                Dim lstimpdetalle As Object
                                'Detalle de impuesto de IVA
                                Dim impdetalleIVA As Object

                                lstimpdetalle = New List(Of Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDetalleFacturaImpuesto)
                                impdetalleIVA = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDetalleFacturaImpuesto

                                impdetalleIVA.Codigo = "2" '2 de que tabla debo verlo tabla 15 SRI
                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                    impdetalleIVA.CodigoPorcentaje = 0  '2 de que tabla debo verlo tabla 16
                                    impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA" Then ' 12%
                                    impdetalleIVA.CodigoPorcentaje = 2 '2 de que tabla debo verlo tabla 16
                                    impdetalleIVA.Tarifa = 12
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA13" Then ' 12%
                                    impdetalleIVA.CodigoPorcentaje = 3 '2 de que tabla debo verlo tabla 16
                                    impdetalleIVA.Tarifa = 14
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_EXO" Then ' 12%
                                    impdetalleIVA.CodigoPorcentaje = 6 '2 de que tabla debo verlo tabla 16
                                    impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                End If

                                impdetalleIVA.BaseImponible = r("PrecioTotalSinImpuesto")
                                impdetalleIVA.Valor = r("TotalIva")

                                'agrego impuesto a la lista
                                lstimpdetalle.Add(impdetalleIVA)

                                'agrego lista de impuesto al detalle
                                itemDetalleFactura.ENTDetalleFacturaImpuesto = lstimpdetalle.ToArray

                                'agrego detalle a la lista
                                listaDetalle.Add(itemDetalleFactura)
                            Next
                            oFactura.ENTDetalleFactura = listaDetalle.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemCompensancionFac As New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTCompensacion
                                If r("Codigo") <> "" Then
                                    itemCompensancionFac.Codigo = r("Codigo")
                                Else
                                    itemCompensancionFac.Codigo = Nothing
                                End If
                                If r("Tarifa") <> "" Then
                                    itemCompensancionFac.Tarifa = r("Tarifa")
                                Else
                                    itemCompensancionFac.Tarifa = Nothing
                                End If
                                If r("Valor") <> "" Then
                                    itemCompensancionFac.Valor = r("Valor")
                                Else
                                    itemCompensancionFac.Valor = Nothing
                                End If
                                listaFacCompensacion.Add(itemCompensancionFac)
                            Next
                            oFactura.ENTFacturaCompensacion = listaFacCompensacion.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("ENTCompensacion: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "ENTCompensacion: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalFac As Object
                                itemDatoAdicionalFac = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTDatoAdicionalFactura

                                itemDatoAdicionalFac.Nombre = r("Concepto")
                                itemDatoAdicionalFac.Descripcion = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalFac)
                            Next
                            oFactura.ENTDatoAdicionalFactura = listaDatosAdicional.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Cabecera Campo Adicional: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Cabecera Campo Adicional: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 3 Then
                        Try

                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim Pago As Object
                                Pago = New Entidades.WSEDOCNUBE_FACTURAS_v4_3.ENTFacturaPagos

                                Pago.FormaPago = r("FormaPago")
                                Pago.Total = r("Total")
                                If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                    Pago.Plazo = Nothing
                                Else
                                    Pago.Plazo = r("Plazo")
                                End If
                                If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                    Pago.UnidadTiempo = Nothing
                                Else
                                    Pago.UnidadTiempo = r("UnidadTiempo")
                                End If
                                FormasdePago.Add(Pago)
                            Next
                            oFactura.ENTFacturaPagos = FormasdePago.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    End If

                    ' FACTURA DE EXPORTACIÓN
                    If oFactura.Tipo = 1 Then
                        oFactura = ConsultaFacturaExportacion(TipoFactura, TipoWS, DocEntry, oFactura)
                    End If
                    ' FACTURA DE REEMBOLSO
                    If oFactura.Tipo = 2 Then
                        oFactura = ConsultaFacturaReembolso(TipoFactura, TipoWS, DocEntry, oFactura)
                    End If
                Next

            End If

            Return oFactura

        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

            oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "ArgumentException-Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la factura en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

            oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Error al Consultar Factura con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try
    End Function
    Private Function ConsultaFacturaExportacion(TipoFactura As String, TipoWS As String, DocEntry As Integer, oFactura As Object) As Object
        Dim SP As String = ""
        Try

            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "GS_SAP_FE_ObtenerFacturadeVenta_EXPORTACION "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "GS_SAP_FE_ONE_OBTENERFACTURADEVENTA_EXPORTACION "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "GS_SAP_FE_HEI_OBTENERFACTURADEVENTA_EXPORTACION"
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "GS_SAP_FE_SYP_OBTENERFACTURADEVENTA_EXPORTACION"
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "GS_SAP_FE_TM_OBTENERFACTURADEVENTA_EXPORTACION"
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                SP = "GS_SAP_FE_SS_OBTENERFACTURADEVENTA_EXPORTACION"
            End If

            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Favor espere... Consultando Datos de Factura de Exportación, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Consultando Datos de Factura de Exportación, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP




            If TipoFactura = "FAE" Then
                SP = GetQueryConsulta(tipoDocumento.FacturaAnticipo, DocEntry, "EXPORT")
            Else
                SP = GetQueryConsulta(tipoDocumento.Factura, DocEntry, "EXPORT")
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


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Exportacion"



            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS







            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables EXPORTACION : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then ' TABLA 0 ES EXPORTACION
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' Dim X As New Entidades.wsEDoc_Factura.ENTFactura
                            oFactura.ComercioExterior = r("ComercioExterior")
                            oFactura.IncoTermFactura = r("IncoTermFactura")
                            oFactura.LugarIncoTerm = r("LugarIncoTerm")
                            oFactura.PaisOrigen = r("PaisOrigen")
                            oFactura.PuertoEmbarque = r("PuertoEmbarque")
                            oFactura.PuertoDestino = r("PuertoDestino")
                            oFactura.PaisDestino = r("PaisDestino")
                            oFactura.PaisAdquisicion = r("PaisAdquisicion")
                            oFactura.IncoTermTotalSinImpuestos = r("IncoTermTotalSinImpuestos")

                            oFactura.FleteInternacional = r("FleteInternacional")
                            oFactura.SeguroInternacional = r("SeguroInternacional")
                            oFactura.GastosAduaneros = r("GastosAduaneros")
                            oFactura.GastosTransporteOtros = r("GastosTransporteOtros")
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
    Private Function ConsultaFacturaReembolso(TipoFactura As String, TipoWS As String, DocEntry As Integer, oFactura As Object) As Object
        Dim SP As String = ""
        Try

            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "GS_SAP_FE_ObtenerFacturadeVenta_REEMBOLSO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "GS_SAP_FE_ONE_OBTENERFACTURADEVENTA_REEMBOLSO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "GS_SAP_FE_HEI_OBTENERFACTURADEVENTA_REEMBOLSO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "GS_SAP_FE_SYP_OBTENERFACTURADEVENTA_REEMBOLSO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "GS_SAP_FE_TM_OBTENERFACTURADEVENTA_REEMBOLSO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                SP = "GS_SAP_FE_SS_OBTENERFACTURADEVENTA_REEMBOLSO "
            End If

            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Favor espere... Consultando Datos de Factura de Reembolso, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Consultando Datos de Factura de Reembolso, # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP



            If TipoFactura = "FAE" Then
                SP = GetQueryConsulta(tipoDocumento.FacturaAnticipo, DocEntry, "REEMBOLSO")
            Else
                SP = GetQueryConsulta(tipoDocumento.Factura, DocEntry, "REEMBOLSO")
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


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Rembolso"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "RembolsoDet"
                ds.Tables.Add(dt1)


            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS





            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables REEMBOLSO: " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then ' TABLA 0 ES EXPORTACION
                        For Each r As DataRow In ds.Tables(0).Rows
                            ' Dim X As New Entidades.wsEDoc_Factura.ENTFactura
                            oFactura.CodDocReemb = r("CodDocReemb")
                            oFactura.TotalComprobantesReembolso = r("TotalComprobantesReembolso")
                            oFactura.TotalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                            oFactura.TotalImpuestoReembolso = r("TotalImpuestoReembolso")

                        Next
                    ElseIf i = 1 Then

                        Dim Detalle As Object
                        Dim listaDetalle As Object
                        Dim LisimpdetalleIVA As Object
                        Dim impdetalleIVA As Object


                        If TipoWS = "LOCAL" Then
                            Detalle = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolso
                            listaDetalle = New List(Of Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolso)
                            LisimpdetalleIVA = New List(Of Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto)
                            impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                        ElseIf TipoWS = "NUBE_4_1" Then
                            Detalle = New Entidades.wsEDoc_Factura41.ENTFacturaReembolso
                            listaDetalle = New List(Of Entidades.wsEDoc_Factura41.ENTFacturaReembolso)
                            LisimpdetalleIVA = New List(Of Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto)
                            impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                        Else
                            Detalle = New Entidades.wsEDoc_Factura.ENTFacturaReembolso
                            listaDetalle = New List(Of Entidades.wsEDoc_Factura.ENTFacturaReembolso)
                            LisimpdetalleIVA = New List(Of Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto)
                            impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                        End If

                        For Each r As DataRow In ds.Tables(1).Rows
                            If TipoWS = "LOCAL" Then
                                Detalle = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolso
                            ElseIf TipoWS = "NUBE_4_1" Then
                                Detalle = New Entidades.wsEDoc_Factura41.ENTFacturaReembolso
                            Else
                                Detalle = New Entidades.wsEDoc_Factura.ENTFacturaReembolso
                            End If
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
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIVA0")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIVA0")
                                impdetalleIVA.Tarifa = r("TarifaIVA0")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIVA0")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIVA0")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)

                            End If

                            If r("BaseImponibleIVA").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIVA")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIVA")
                                impdetalleIVA.Tarifa = r("TarifaIVA")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIVA")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIVA")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleNoObjIVA").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoNoObjIVA")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeNoObjIVA")
                                impdetalleIVA.Tarifa = r("TarifaNoObjIVA")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleNoObjIVA")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoNoObjIVA")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIvaExe").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIvaExe")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIvaExe")
                                impdetalleIVA.Tarifa = r("TarifaIvaExe")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIvaExe")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIvaExe")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva8").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIva8")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIva8")
                                impdetalleIVA.Tarifa = r("TarifaIva8")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva8")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIva8")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If


                            If r("BaseImponibleIva5").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIva5")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIva5")
                                impdetalleIVA.Tarifa = r("TarifaIva5")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva5")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIva5")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva15").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIva15")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIva15")
                                impdetalleIVA.Tarifa = r("TarifaIva15")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva15")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIva15")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva14").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIva14")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIva14")
                                impdetalleIVA.Tarifa = r("TarifaIva14")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva14")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIva14")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            If r("BaseImponibleIva13").ToString() <> 0 Then
                                If TipoWS = "LOCAL" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura_LOCAL.ENTFacturaReembolsoImpuesto
                                ElseIf TipoWS = "NUBE_4_1" Then
                                    impdetalleIVA = New Entidades.wsEDoc_Factura41.ENTFacturaReembolsoImpuesto
                                Else
                                    impdetalleIVA = New Entidades.wsEDoc_Factura.ENTFacturaReembolsoImpuesto
                                End If
                                impdetalleIVA.Codigo = r("CodigoIva13")
                                impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentajeIva13")
                                impdetalleIVA.Tarifa = r("TarifaIva13")
                                impdetalleIVA.BaseImponibleReembolso = r("BaseImponibleIva13")
                                impdetalleIVA.ImpuestoReembolso = r("ImpuestoReembolsoIva13")
                                'agrego impuesto a la lista
                                LisimpdetalleIVA.Add(impdetalleIVA)
                            End If

                            'agrego lista de impuesto al detalle
                            Detalle.ENTFacturaReembolsoImpuestos = LisimpdetalleIVA.ToArray

                            listaDetalle.Add(Detalle)

                        Next
                        oFactura.ENTFacturaReembolso = listaDetalle.ToArray()
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


    Public Function ConsultarNotadeCredito(ByVal TipoNC As String, ByVal DocEntry As Integer) As Object

        Dim oNC As Entidades.RequestNotaCredito = Nothing
        Dim listaDetalle As List(Of Entidades.detalleNCE)
        Dim listaDatosAdicional As List(Of Entidades.infoAdicionalNCE)
        Dim listaTotalesConImpuestos As List(Of Entidades.totalConImpuestosNCE)
        Dim listaDatosAdicionalDetalle As List(Of Entidades.detallesAdicionalesNCE)
        Dim listaImpuestos As List(Of Entidades.impuestosNCE)

        listaDetalle = New List(Of Entidades.detalleNCE)
        listaDatosAdicional = New List(Of Entidades.infoAdicionalNCE)
        listaTotalesConImpuestos = New List(Of Entidades.totalConImpuestosNCE)

        Try
            Dim SP As String = ""

            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "GS_SAP_FE_ObtenerNotaDeCredito_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "GS_SAP_FE_ONE_OBTENERNOTADECREDITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "GS_SAP_FE_HEI_OBTENERNOTADECREDITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "GS_SAP_FE_SYP_OBTENERNOTADECREDITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "GS_SAP_FE_TM_OBTENERNOTADECREDITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                SP = "GS_SAP_FE_SS_OBTENERNOTADECREDITO_4_3 "
            End If

            If _tipoManejo = "A" Then oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"Consultando Nota de Crédito # DocEntry: {DocEntry} SP: {SP}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)

            SP = GetQueryConsulta(tipoDocumento.NotaCredito, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Query Desencriptado " & SP.ToString(), "ManejoDeDocumentos")

            If SP.Contains("GSCODEEXCEPCION") Then
                Utilitario.Util_Log.Escribir_Log("EXCEPCION DETECTADA EN EL PROCESO DE OBTENER STRING QUERY - " & SP, "ManejoDeDocumentos")
                rsboApp.StatusBar.SetText(VariablesGlobales._gNombreAddOn + " - Ocurrio Un Error Favor falidar el Archivo de Log y Buscar el Codigo GSCODEEXCEPCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            Else
                ds = EjecutarSP(SP, DocEntry)
            End If

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")
            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "2") Then Return Nothing
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows
                            Try
                                oNC.infoTributaria.ambiente = r("Ambiente").ToString

                                oNC.infoTributaria.claveAcceso = r("ClaveAcceso").ToString

                                oNC.infoTributaria.razonSocial = r("RazonSocial").ToString

                                oNC.infoTributaria.nombreComercial = r("NombreComercial").ToString

                                oNC.infoTributaria.ruc = r("Ruc").ToString

                                oNC.infoTributaria.tipoEmision = r("TipoEmision").ToString

                                oNC.infoTributaria.codDoc = r("CodigoDocumento").ToString

                                oNC.infoTributaria.estab = r("Establecimiento").ToString

                                oNC.infoTributaria.ptoEmi = r("PuntoEmision").ToString

                                oNC.infoTributaria.secuencial = r("SecuencialDocumento").ToString
                                If Not oNC.infoTributaria.secuencial.ToString().Length.Equals("9") Then oNC.infoTributaria.secuencial = oNC.infoTributaria.secuencial.PadLeft(9, "0")
                                Utilitario.Util_Log.Escribir_Log("oNC.Secuencial : " & oNC.infoTributaria.secuencial.ToString(), "ManejoDeDocumentos")

                                oNC.infoTributaria.dirMatriz = r("DireccionMatriz").ToString

                                oNC.infoTributaria.diaEmission = CDate(r("FechaEmision")).ToString("dd")

                                oNC.infoTributaria.mesEmission = CDate(r("FechaEmision")).ToString("MM")

                                oNC.infoTributaria.anioEmission = CDate(r("FechaEmision")).ToString("yyyy")

                                oNC.infoNotaCredito.fechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")

                                oNC.infoNotaCredito.dirEstablecimiento = r("DireccionEstablecimiento").ToString

                                oNC.infoNotaCredito.tipoIdentificacionComprador = r("TipoIdentificadorComprador").ToString

                                oNC.infoNotaCredito.razonSocialComprador = r("RazonSocialComprador").ToString

                                oNC.infoNotaCredito.identificacionComprador = r("IdentificacionComprador").ToString

                                If Not r("ContribuyenteEspecial") = "0" Then oNC.infoNotaCredito.contribuyenteEspecial = r("ContribuyenteEspecial").ToString

                                oNC.infoNotaCredito.obligadoContabilidad = r("ObligadoContabilidad").ToString

                                oNC.infoNotaCredito.rise = r("Rise").ToString

                                oNC.infoNotaCredito.codDocModificado = r("codDocModificado").ToString

                                oNC.infoNotaCredito.numDocModificado = r("numDocModificado").ToString

                                oNC.infoNotaCredito.fechaEmisionDocSustento = CDate(r("FechaEmisionDocModificado")).ToString("yyyy-MM-dd")

                                oNC.infoNotaCredito.totalSinImpuestos = r("TotalSinImpuesto").ToString

                                oNC.infoNotaCredito.valorModificacion = r("ValorModificacion").ToString

                                oNC.infoNotaCredito.moneda = r("Moneda").ToString

                                oNC.infoNotaCredito.motivo = r("Motivo").ToString

                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo8").ToString
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje8").ToString
                                    impfaIVA.baseImponible = r("Base8").ToString
                                    impfaIVA.valor = r("ValorIva8").ToString
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo12").ToString
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje12").ToString
                                    impfaIVA.baseImponible = r("Base12").ToString
                                    impfaIVA.valor = r("ValorIva12").ToString
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaIVA.codigo = r("Codigo13").ToString
                                    impfaIVA.codigoPorcentaje = r("CodigoPorcentaje13").ToString
                                    impfaIVA.baseImponible = r("Base13").ToString
                                    impfaIVA.valor = r("ValorIva13").ToString
                                    listaTotalesConImpuestos.Add(impfaIVA)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo0").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje0").ToString
                                    impfaNOIVA.baseImponible = r("Base0").ToString
                                    impfaNOIVA.valor = r("ValorIva0").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("CodigoNoi").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeNoi").ToString
                                    impfaNOIVA.baseImponible = r("BaseNoi").ToString
                                    impfaNOIVA.valor = r("ValorIvaNoi").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("CodigoExen").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeExen").ToString
                                    impfaNOIVA.baseImponible = r("BaseExen").ToString
                                    impfaNOIVA.valor = r("ValorIvaExen").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("BaseIce") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("CodigoIce").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentajeIce").ToString
                                    impfaNOIVA.baseImponible = r("BaseIce").ToString
                                    impfaNOIVA.valor = r("ValorIvaIce").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base5") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo5").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje5").ToString
                                    impfaNOIVA.baseImponible = r("Base5").ToString
                                    impfaNOIVA.valor = r("ValorIva5").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base15") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo15").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje15").ToString
                                    impfaNOIVA.baseImponible = r("Base15").ToString
                                    impfaNOIVA.valor = r("ValorIva15").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                If r("Base14") <> 0 Then
                                    Dim impfaNOIVA As Entidades.totalConImpuestosNCE = New Entidades.totalConImpuestosNCE
                                    impfaNOIVA.codigo = r("Codigo14").ToString
                                    impfaNOIVA.codigoPorcentaje = r("CodigoPorcentaje14").ToString
                                    impfaNOIVA.baseImponible = r("Base14").ToString
                                    impfaNOIVA.valor = r("ValorIva14").ToString
                                    listaTotalesConImpuestos.Add(impfaNOIVA)
                                End If

                                oNC.infoNotaCredito.totalConImpuestos = listaTotalesConImpuestos
                            Catch ex As Exception
                                If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Cabecera nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                _Error = "Cabecera: " + ex.Message.ToString()
                                Utilitario.Util_Log.Escribir_Log("Cabcera nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                Return Nothing
                            End Try
                        Next
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleNC As New Entidades.detalleNCE

                                itemDetalleNC.codigoPrincipal = r("CodigoPrincipal").ToString

                                itemDetalleNC.codigoAuxiliar = r("CodigoAuxiliar").ToString

                                itemDetalleNC.descripcion = r("Descripcion").ToString

                                itemDetalleNC.cantidad = CInt(r("Cantidad"))

                                itemDetalleNC.precioUnitario = r("PrecioUnitario").ToString

                                itemDetalleNC.descuento = r("Descuento").ToString

                                itemDetalleNC.precioTotalSinImpuesto = r("PrecioTotalSinImpuesto").ToString

                                listaDatosAdicionalDetalle = New List(Of Entidades.detallesAdicionalesNCE)

                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As Entidades.detallesAdicionalesNCE = New Entidades.detallesAdicionalesNCE
                                    itemDetalleDatoAdicional.nombre = r("ConceptoAdicional1")
                                    itemDetalleDatoAdicional.valor = r("NombreAdicional1")
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

                                itemDetalleNC.detallesAdicionales = listaDatosAdicionalDetalle

                                listaImpuestos = New List(Of Entidades.impuestosNCE)

                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA8" Then ' 0%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA13" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")

                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeIce") = "IVA_ICE" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("CodigoIce")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentajeIce")
                                    impdetalleIVA.tarifa = r("TarifaIce")
                                    impdetalleIVA.baseImponible = r("BaseImponibleIce")
                                    impdetalleIVA.valor = r("TotalIvaIce")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA5" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA15" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                If r("TaxCodeAp") = "IVA14" Then ' 12%
                                    Dim impdetalleIVA As Entidades.impuestosNCE = New Entidades.impuestosNCE
                                    impdetalleIVA.codigo = r("Codigo")
                                    impdetalleIVA.codigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.tarifa = r("Tarifa")
                                    impdetalleIVA.baseImponible = r("BaseImponible")
                                    impdetalleIVA.valor = r("TotalIva")
                                    listaImpuestos.Add(impdetalleIVA)
                                End If

                                itemDetalleNC.impuestos = listaImpuestos

                                listaDetalle.Add(itemDetalleNC)
                            Next
                            oNC.detalles = listaDetalle
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("detalle nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Utilitario.Util_Log.Escribir_Log("detalle nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            _Error = "detalle nota de credito error : " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalNce As Entidades.infoAdicionalNCE = New Entidades.infoAdicionalNCE
                                itemDatoAdicionalNce.nombre = r("Concepto").ToString
                                itemDatoAdicionalNce.valor = r("Descripcion").ToString
                                listaDatosAdicional.Add(itemDatoAdicionalNce)
                            Next
                            oNC.infoAdicional = listaDatosAdicional
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("informacion adicional nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Utilitario.Util_Log.Escribir_Log("informacion adicional nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            _Error = "adicionales de nota de credito error : " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    End If
                Next
            End If

            Return oNC

        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"ArgumentException-Ocurrio un error al consultar datos de la Nota de Credito en la Base, DocEntry: {DocEntry}, Descr: {x.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"ArgumentException-Error al Consultar Nota de Credito # DocEntry: {DocEntry}, Descr: {x.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"Ocurrio un error al consultar datos de la oNotaCredito en la Base, DocEntry: {DocEntry}, Descr: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoNC, DocEntry, $"Error al Consultar Nota de Credito # DocEntry: {DocEntry}, Descr: {ex.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        End Try
    End Function

    Public Function ConsultarNotadeDebito(ByVal TipoND As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object
        Dim oNotaDebito As Object = Nothing
        If TipoWS = "NUBE_4_1" Then
            oNotaDebito = ConsultarNotadeDebito_NUBE_4_1(TipoND, DocEntry, TipoWS)
        Else
            oNotaDebito = ConsultarNotadeDebito_LOCAL_NUBE(TipoND, DocEntry, TipoWS)
        End If

        Return oNotaDebito

    End Function
    Public Function ConsultarNotadeDebito_LOCAL_NUBE(ByVal TipoND As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oNotaDebito As Object = Nothing
        Dim listaDetalle As Object
        Dim listaDatosAdicional As Object
        Dim FormasdePago As Object

        If TipoWS = "LOCAL" Then
            listaDetalle = New List(Of Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTDetalleNotaDebito)
            listaDatosAdicional = New List(Of Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTDatoAdicionalNotaDebito)
            FormasdePago = New List(Of Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoPagos)
        Else
            listaDetalle = New List(Of Entidades.wsEDoc_NotaDeDebito.ENTDetalleNotaDebito)
            listaDatosAdicional = New List(Of Entidades.wsEDoc_NotaDeDebito.ENTDatoAdicionalNotaDebito)
            FormasdePago = New List(Of Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoPagos)
        End If

        Try

            Dim SP As String = ""

            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "GS_SAP_FE_ObtenerNotaDeDebito "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "GS_SAP_FE_ONE_OBTENERNOTADEDEBITO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "GS_SAP_FE_HEI_OBTENERNOTADEDEBITO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "GS_SAP_FE_SYP_OBTENERNOTADEDEBITO "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "GS_SAP_FE_TM_OBTENERNOTADEDEBITO "
            End If

            oFuncionesAddon.GuardaLOG(TipoND, DocEntry, "Consultando Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)



            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                If TipoWS = "LOCAL" Then
                    oNotaDebito = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebito
                Else
                    oNotaDebito = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebito
                End If

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oNotaDebito.ClaveAcceso = Nothing
                            Else
                                oNotaDebito.ClaveAcceso = r("ClaveAcceso")
                            End If


                            oNotaDebito.Ambiente = r("Ambiente")
                            oNotaDebito.TipoEmision = r("TipoEmision")

                            oNotaDebito.RazonSocial = r("RazonSocial")
                            oNotaDebito.NombreComercial = r("NombreComercial")
                            oNotaDebito.Ruc = r("Ruc")
                            'oNotaDebito.Ruc = "0992737964001"

                            oNotaDebito.CodigoDocumento = r("CodigoDocumento")
                            oNotaDebito.Establecimiento = r("Establecimiento")
                            oNotaDebito.PuntoEmision = r("PuntoEmision")
                            oNotaDebito.Secuencial = r("SecuencialDocumento")
                            If Not oNotaDebito.Secuencial.ToString().Length.Equals("9") Then
                                oNotaDebito.Secuencial = oNotaDebito.Secuencial.PadLeft(9, "0")
                            End If
                            oNotaDebito.DireccionMatriz = r("DireccionMatriz")
                            oNotaDebito.FechaEmision = r("FechaEmision")
                            oNotaDebito.DireccionEstablecimiento = r("DireccionEstablecimiento")

                            If Not r("ContribuyenteEspecial") = "0" Then
                                oNotaDebito.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oNotaDebito.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oNotaDebito.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oNotaDebito.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oNotaDebito.AgenteRetencion = r("ContribuyenteRimpe")
                            End If

                            oNotaDebito.ObligadoContabilidad = r("ObligadoContabilidad")

                            oNotaDebito.CodDocModificado = r("codDocModificado")
                            oNotaDebito.NumDocModificado = r("numDocModificado")
                            oNotaDebito.FechaEmisionDocModificado = r("FechaEmisionDocModificado")

                            oNotaDebito.TipoIdentificacionComprador = r("TipoIdentificadorComprador")

                            oNotaDebito.RazonSocialComprador = r("RazonSocialComprador")
                            oNotaDebito.IdentificacionComprador = r("IdentificacionComprador")

                            oNotaDebito.TotalSinImpuesto = r("TotalSinImpuesto")
                            ' oNotaDebito.TotalDescuento = r("TotalDescuento")

                            '   oNotaDebito.Propina = r("Propina")
                            oNotaDebito.ValorTotal = r("ImporteTotal")
                            'oNotaDebito.Moneda = r("Moneda")
                            'oNotaDebito.UsuarioCreador = r("UsuarioCreador")
                            oNotaDebito.UsuarioTransaccionERP = r("UsuarioCreador")
                            oNotaDebito.EmailResponsable = r("EmailResponsable")
                            'oNotaDebito.Estado = r("Telefono")
                            'oNotaDebito. = r("Telefono2")
                            oNotaDebito.SecuencialERP = r("SecuencialERP")
                            oNotaDebito.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oNotaDebito.Estado = r("Estado")
                            '  oNotaDebito.FechaCarga = r("FechaCarga")
                            oNotaDebito.Campo1 = r("Campo1")
                            oNotaDebito.Campo2 = r("Campo2")
                            oNotaDebito.Campo3 = r("Campo3")

                            ' oNotaDebito.MotivoModificacion = r("Motivo")

                            'IMPUESTO FACTURA
                            'Impuestos totalizados en la factura.
                            Dim lstimpfact As Object
                            If TipoWS = "LOCAL" Then
                                lstimpfact = New List(Of Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto)
                            Else
                                lstimpfact = New List(Of Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto)
                            End If
                            If r("Base12") <> 0 Then
                                Dim impfaIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaIVA.Codigo = "2"
                                'impfaIVA.CodigoPorcentaje = "2"
                                'impfaIVA.Tarifa = "12"
                                'impfaIVA.BaseImponible = r("Base12")
                                'impfaIVA.Valor = r("ImpuestoTotal")
                                impfaIVA.Codigo = r("Codigo12")
                                impfaIVA.CodigoPorcentaje = r("CodigoPorcentaje12")
                                impfaIVA.Tarifa = r("Tarifa12")
                                impfaIVA.BaseImponible = r("Base12")
                                impfaIVA.Valor = r("ValorIva12")

                                lstimpfact.Add(impfaIVA)
                            End If
                            If r("Base13") <> 0 Then
                                Dim impfaIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaIVA.Codigo = "2"
                                'impfaIVA.CodigoPorcentaje = "3"
                                'impfaIVA.Tarifa = "14"
                                'impfaIVA.BaseImponible = r("Base13")
                                'impfaIVA.Valor = r("ImpuestoTotal")
                                impfaIVA.Codigo = r("Codigo13")
                                impfaIVA.CodigoPorcentaje = r("CodigoPorcentaje13")
                                impfaIVA.Tarifa = r("Tarifa13")
                                impfaIVA.BaseImponible = r("Base13")
                                impfaIVA.Valor = r("ValorIva13")

                                lstimpfact.Add(impfaIVA)
                            End If

                            If r("Base0") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "0"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("Base0")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("Codigo0")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje0")
                                impfaNOIVA.Tarifa = r("Tarifa0")
                                impfaNOIVA.BaseImponible = r("Base0")
                                impfaNOIVA.Valor = r("ValorIva0")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            If r("BaseNoi") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "6"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("BaseExo")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("CodigoNoi")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentajeNoi")
                                impfaNOIVA.Tarifa = r("TarifaNoi")
                                impfaNOIVA.BaseImponible = r("BaseNoi")
                                impfaNOIVA.Valor = r("ValorIvaNoi")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            If r("BaseExen") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "6"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("BaseExo")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("CodigoExen")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentajeExen")
                                impfaNOIVA.Tarifa = r("TarifaExen")
                                impfaNOIVA.BaseImponible = r("BaseExen")
                                impfaNOIVA.Valor = r("ValorIvaExen")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            If r("BaseIce") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "6"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("BaseExo")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("CodigoIce")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentajeIce")
                                impfaNOIVA.Tarifa = r("TarifaIce")
                                impfaNOIVA.BaseImponible = r("BaseIce")
                                impfaNOIVA.Valor = r("ValorIvaIce")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            If r("Base5") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "6"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("BaseExo")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("Codigo5")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje5")
                                impfaNOIVA.Tarifa = r("Tarifa5")
                                impfaNOIVA.BaseImponible = r("Base5")
                                impfaNOIVA.Valor = r("ValorIva5")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            If r("Base15") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "6"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("BaseExo")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("Codigo15")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje15")
                                impfaNOIVA.Tarifa = r("Tarifa15")
                                impfaNOIVA.BaseImponible = r("Base15")
                                impfaNOIVA.Valor = r("ValorIva15")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            If r("Base14") <> 0 Then
                                Dim impfaNOIVA As Object
                                If TipoWS = "LOCAL" Then
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoImpuesto
                                Else
                                    impfaNOIVA = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoImpuesto
                                End If
                                'impfaNOIVA.Codigo = "2"
                                'impfaNOIVA.CodigoPorcentaje = "6"
                                'impfaNOIVA.Tarifa = "0"
                                'impfaNOIVA.BaseImponible = r("BaseExo")
                                'impfaNOIVA.Valor = 0
                                impfaNOIVA.Codigo = r("Codigo14")
                                impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje14")
                                impfaNOIVA.Tarifa = r("Tarifa14")
                                impfaNOIVA.BaseImponible = r("Base14")
                                impfaNOIVA.Valor = r("ValorIva14")

                                lstimpfact.Add(impfaNOIVA)
                            End If

                            oNotaDebito.ENTNotaDebitoImpuesto = lstimpfact.ToArray
                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows
                            Dim itemDetalleND As Object
                            If TipoWS = "LOCAL" Then
                                itemDetalleND = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTDetalleNotaDebito
                            Else
                                itemDetalleND = New Entidades.wsEDoc_NotaDeDebito.ENTDetalleNotaDebito
                            End If
                            'itemDetalleND.ValorSpecified = True
                            itemDetalleND.Razon = r("Descripcion")
                            itemDetalleND.Valor = r("PrecioTotalSinImpuesto")

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalleND)
                        Next
                        oNotaDebito.DetalleNotaDebito = listaDetalle.ToArray
                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDatoAdicionalFac As Object
                            If TipoWS = "LOCAL" Then
                                itemDatoAdicionalFac = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTDatoAdicionalNotaDebito
                            Else
                                itemDatoAdicionalFac = New Entidades.wsEDoc_NotaDeDebito.ENTDatoAdicionalNotaDebito
                            End If
                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oNotaDebito.ENTDatoAdicionalNotaDebito = listaDatosAdicional.ToArray
                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim Pago As Object
                            If TipoWS = "LOCAL" Then
                                Pago = New Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebitoPagos
                            Else
                                Pago = New Entidades.wsEDoc_NotaDeDebito.ENTNotaDebitoPagos
                            End If
                            Pago.FormaPago = r("FormaPago")
                            Pago.Total = r("Total")
                            If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                Pago.Plazo = Nothing
                            Else
                                Pago.Plazo = r("Plazo")
                            End If
                            If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                Pago.UnidadTiempo = Nothing
                            Else
                                Pago.UnidadTiempo = r("UnidadTiempo")
                            End If
                            FormasdePago.Add(Pago)
                        Next
                        oNotaDebito.ENTNotaDebitoPagos = FormasdePago.ToArray
                    End If
                Next

            End If


            'SERIALIZACION OBJETO XML
            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oNotaDebito.Secuencial.ToString() + oNotaDebito.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")

                    Dim x As XmlSerializer = Nothing

                    If TipoWS = "LOCAL" Then
                        x = New XmlSerializer(GetType(Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebito))
                    Else
                        x = New XmlSerializer(GetType(Entidades.wsEDoc_NotaDeDebito.ENTNotaDebito))
                    End If

                    Dim writer As TextWriter = New StreamWriter(sRuta)
                    x.Serialize(writer, oNotaDebito)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try


            Return oNotaDebito
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Nota de Debito7 en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

            oFuncionesAddon.GuardaLOG(TipoND, DocEntry, "ArgumentException-Error al Consultar Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la oNotaDebito en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

            oFuncionesAddon.GuardaLOG(TipoND, DocEntry, "Error al Consultar Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

    End Function
    Public Function ConsultarNotadeDebito_NUBE_4_1(ByVal TipoND As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oNotaDebito As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebito
        Dim listaDetalle As New List(Of Entidades.wsEDoc_NotaDeDebito41.ENTDetalleNotaDebito)
        Dim listaDatosAdicional As New List(Of Entidades.wsEDoc_NotaDeDebito41.ENTDatoAdicionalNotaDebito)
        Dim FormasdePago As New List(Of Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoPagos)

        Try

            Dim SP As String = ""

            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                SP = "GS_SAP_FE_ObtenerNotaDeDebito_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                SP = "GS_SAP_FE_ONE_OBTENERNOTADEDEBITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                SP = "GS_SAP_FE_HEI_OBTENERNOTADEDEBITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                SP = "GS_SAP_FE_SYP_OBTENERNOTADEDEBITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                SP = "GS_SAP_FE_TM_OBTENERNOTADEDEBITO_4_3 "
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                SP = "GS_SAP_FE_SS_OBTENERNOTADEDEBITO_4_3 "
            End If

            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoND, DocEntry, "Consultando Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP


            SP = GetQueryConsulta(tipoDocumento.NotaDebito, DocEntry)

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


                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")

                Dim SPs() As String = Split(SP, "--*")


                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

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



            '------------------------FIN QUERYS ENCRIPTADOS





            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "2") Then
                    Return Nothing
                End If
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then


                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            Try
                                If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                    oNotaDebito.ClaveAcceso = Nothing
                                Else
                                    oNotaDebito.ClaveAcceso = r("ClaveAcceso")
                                End If

                                oNotaDebito.Ambiente = r("Ambiente")
                                oNotaDebito.TipoEmision = r("TipoEmision")

                                oNotaDebito.RazonSocial = r("RazonSocial")

                                If Not r("NombreComercial") = "" Then
                                    oNotaDebito.NombreComercial = r("NombreComercial")
                                End If

                                oNotaDebito.Ruc = r("Ruc")
                                'oNotaDebito.Ruc = "0992737964001"

                                oNotaDebito.CodigoDocumento = r("CodigoDocumento")
                                oNotaDebito.Establecimiento = r("Establecimiento")
                                oNotaDebito.PuntoEmision = r("PuntoEmision")
                                oNotaDebito.Secuencial = r("SecuencialDocumento")
                                If Not oNotaDebito.Secuencial.ToString().Length.Equals("9") Then
                                    oNotaDebito.Secuencial = oNotaDebito.Secuencial.PadLeft(9, "0")
                                End If
                                oNotaDebito.DireccionMatriz = r("DireccionMatriz")
                                'oNotaDebito.FechaEmision = r("FechaEmision")
                                oNotaDebito.FechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")

                                oNotaDebito.DireccionEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oNotaDebito.ContribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oNotaDebito.ContribuyenteEspecial = Nothing
                                End If

                                If Not r("AgenteRetencion") = "0" Then
                                    oNotaDebito.AgenteRetencion = r("AgenteRetencion")
                                End If

                                If Not r("RegimenMicroempresas") = "0" Then
                                    oNotaDebito.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                                End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    oNotaDebito.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                                End If

                                oNotaDebito.ObligadoContabilidad = r("ObligadoContabilidad")

                                oNotaDebito.CodDocModificado = r("codDocModificado")
                                oNotaDebito.NumDocModificado = r("numDocModificado")
                                'oNotaDebito.FechaEmisionDocModificado = r("FechaEmisionDocModificado")
                                oNotaDebito.FechaEmisionDocModificado = CDate(r("FechaEmisionDocModificado")).ToString("yyyy-MM-dd")

                                oNotaDebito.TipoIdentificacionComprador = r("TipoIdentificadorComprador")

                                oNotaDebito.RazonSocialComprador = r("RazonSocialComprador")
                                oNotaDebito.IdentificacionComprador = r("IdentificacionComprador")

                                oNotaDebito.TotalSinImpuesto = r("TotalSinImpuesto")
                                ' oNotaDebito.TotalDescuento = r("TotalDescuento")

                                '   oNotaDebito.Propina = r("Propina")
                                oNotaDebito.ValorTotal = r("ImporteTotal")
                                'oNotaDebito.Moneda = r("Moneda")
                                'oNotaDebito.UsuarioCreador = r("UsuarioCreador")
                                oNotaDebito.UsuarioTransaccionERP = r("UsuarioCreador")
                                oNotaDebito.EmailResponsable = r("EmailResponsable")
                                'oNotaDebito.Estado = r("Telefono")
                                'oNotaDebito. = r("Telefono2")
                                oNotaDebito.SecuencialERP = r("SecuencialERP")
                                oNotaDebito.CodigoTransaccionERP = r("CodigoTransaccionERP")
                                oNotaDebito.Estado = r("Estado")
                                '  oNotaDebito.FechaCarga = r("FechaCarga")
                                oNotaDebito.Campo1 = r("Campo1")
                                oNotaDebito.Campo2 = r("Campo2")
                                oNotaDebito.Campo3 = r("Campo3")

                                'ADD DM 08012025
                                oNotaDebito.Campo4 = r("Campo4")
                                oNotaDebito.Campo5 = r("Campo5")
                                oNotaDebito.Campo6 = r("Campo6")
                                oNotaDebito.Campo7 = r("Campo7")
                                oNotaDebito.Campo8 = r("Campo8")
                                oNotaDebito.Campo9 = r("Campo9")
                                oNotaDebito.Campo10 = r("Campo10")
                                oNotaDebito.Campo11 = r("Campo11")
                                oNotaDebito.Campo12 = r("Campo12")
                                oNotaDebito.Campo13 = r("Campo13")
                                oNotaDebito.Campo14 = r("Campo14")
                                oNotaDebito.Campo15 = r("Campo15")

                                ' oNotaDebito.MotivoModificacion = r("Motivo")

                                'IMPUESTO FACTURA
                                'Impuestos totalizados en la factura.
                                Dim lstimpfact As New List(Of Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impfaIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto

                                    'impfaIVA.Codigo = "2"
                                    'impfaIVA.CodigoPorcentaje = "2"
                                    'impfaIVA.Tarifa = "12"
                                    impfaIVA.Codigo = r("Codigo8")
                                    impfaIVA.CodigoPorcentaje = r("CodigoPorcentaje8")
                                    impfaIVA.Tarifa = r("Tarifa8")
                                    impfaIVA.BaseImponible = r("Base8")
                                    impfaIVA.Valor = r("ValorIva8")

                                    lstimpfact.Add(impfaIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impfaIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaIVA.Codigo = "2"
                                    'impfaIVA.CodigoPorcentaje = "2"
                                    'impfaIVA.Tarifa = "12"
                                    impfaIVA.Codigo = r("Codigo12")
                                    impfaIVA.CodigoPorcentaje = r("CodigoPorcentaje12")
                                    impfaIVA.Tarifa = r("Tarifa12")
                                    impfaIVA.BaseImponible = r("Base12")
                                    'impfaIVA.Valor = r("ImpuestoTotal")
                                    impfaIVA.Valor = r("ValorIva12")
                                    lstimpfact.Add(impfaIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impfaIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaIVA.Codigo = "2"
                                    'impfaIVA.CodigoPorcentaje = "3"
                                    'impfaIVA.Tarifa = "14"
                                    impfaIVA.Codigo = r("Codigo13")
                                    impfaIVA.CodigoPorcentaje = r("CodigoPorcentaje13")
                                    impfaIVA.Tarifa = r("Tarifa13")
                                    impfaIVA.BaseImponible = r("Base13")
                                    'impfaIVA.Valor = r("ImpuestoTotal")
                                    impfaIVA.Valor = r("ValorIva13")
                                    lstimpfact.Add(impfaIVA)
                                End If


                                If r("Base0") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "0"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.Codigo = r("Codigo0")
                                    impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje0")
                                    impfaNOIVA.Tarifa = r("Tarifa0")
                                    impfaNOIVA.BaseImponible = r("Base0")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.Valor = r("ValorIva0")
                                    lstimpfact.Add(impfaNOIVA)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "6"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.Codigo = r("CodigoNoi")
                                    impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impfaNOIVA.Tarifa = r("TarifaNoi")
                                    impfaNOIVA.BaseImponible = r("BaseNoi")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.Valor = r("ValorIvaNoi")
                                    lstimpfact.Add(impfaNOIVA)

                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "7"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.Codigo = r("CodigoExen")
                                    impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentajeExen")
                                    impfaNOIVA.Tarifa = r("TarifaExen")
                                    impfaNOIVA.BaseImponible = r("BaseExen")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.Valor = r("ValorIvaExen")
                                    lstimpfact.Add(impfaNOIVA)

                                End If

                                If r("Base5") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "7"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.Codigo = r("Codigo5")
                                    impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje5")
                                    impfaNOIVA.Tarifa = r("Tarifa5")
                                    impfaNOIVA.BaseImponible = r("Base5")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.Valor = r("ValorIva5")
                                    lstimpfact.Add(impfaNOIVA)

                                End If

                                If r("Base15") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "7"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.Codigo = r("Codigo15")
                                    impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje15")
                                    impfaNOIVA.Tarifa = r("Tarifa15")
                                    impfaNOIVA.BaseImponible = r("Base15")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.Valor = r("ValorIva15")
                                    lstimpfact.Add(impfaNOIVA)

                                End If

                                If r("Base14") <> 0 Then
                                    Dim impfaNOIVA As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoImpuesto
                                    'impfaNOIVA.Codigo = "2"
                                    'impfaNOIVA.CodigoPorcentaje = "7"
                                    'impfaNOIVA.Tarifa = "0"
                                    impfaNOIVA.Codigo = r("Codigo14")
                                    impfaNOIVA.CodigoPorcentaje = r("CodigoPorcentaje14")
                                    impfaNOIVA.Tarifa = r("Tarifa14")
                                    impfaNOIVA.BaseImponible = r("Base14")
                                    'impfaNOIVA.Valor = 0
                                    impfaNOIVA.Valor = r("ValorIva14")
                                    lstimpfact.Add(impfaNOIVA)

                                End If

                                oNotaDebito.ENTNotaDebitoImpuesto = lstimpfact.ToArray
                            Catch ex As Exception
                                If _tipoManejo = "A" Then
                                    rsboApp.SetStatusBarMessage("CabEcera nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                                'Utilitario.Util_Log.Escribir_Log("DETALLE nota de credito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                Utilitario.Util_Log.Escribir_Log("CabEcera nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                                Return Nothing
                            End Try

                        Next
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleND As New Entidades.wsEDoc_NotaDeDebito41.ENTDetalleNotaDebito
                                'itemDetalleND.ValorSpecified = True
                                itemDetalleND.Razon = r("Descripcion")
                                itemDetalleND.Valor = r("PrecioTotalSinImpuesto")


                                'agrego detalle a la lista
                                listaDetalle.Add(itemDetalleND)
                            Next
                            oNotaDebito.DetalleNotaDebito = listaDetalle.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("detalle nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("detalle nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            Return Nothing
                        End Try


                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows
                                Dim itemDatoAdicionalFac As New Entidades.wsEDoc_NotaDeDebito41.ENTDatoAdicionalNotaDebito
                                itemDatoAdicionalFac.Nombre = r("Concepto")
                                itemDatoAdicionalFac.Descripcion = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalFac)
                            Next
                            oNotaDebito.ENTDatoAdicionalNotaDebito = listaDatosAdicional.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("informacion adicional nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("informacion adicional nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            Return Nothing
                        End Try


                    ElseIf i = 3 Then
                        Try
                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim Pago As New Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebitoPagos
                                Pago.FormaPago = r("FormaPago")
                                Pago.Total = r("Total")
                                If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                    Pago.Plazo = Nothing
                                Else
                                    Pago.Plazo = r("Plazo")
                                End If
                                If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                    Pago.UnidadTiempo = Nothing
                                Else
                                    Pago.UnidadTiempo = r("UnidadTiempo")
                                End If
                                FormasdePago.Add(Pago)
                            Next
                            oNotaDebito.ENTNotaDebitoPagos = FormasdePago.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("FORMA PAGO nota de credito error " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            Utilitario.Util_Log.Escribir_Log("FORMA PAGO  nota de debito error : " & ex.Message.ToString(), "ManejoDeDocumentos")
                            Return Nothing
                        End Try

                    End If
                Next

            End If

            Try

                'Dim sRutaCarpeta As String = ""
                'If _tipoManejo = "A" Then
                '    sRutaCarpeta = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
                'Else
                '    sRutaCarpeta = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'End If
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oNotaDebito.Secuencial.ToString() + oNotaDebito.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oNotaDebito.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oNotaDebito)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oNotaDebito.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oNotaDebito)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oNotaDebito
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Nota de Debito7 en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoND, DocEntry, "ArgumentException-Error al Consultar Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la oNotaDebito en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoND, DocEntry, "Error al Consultar Nota de Debito con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

    End Function

    Public Function ConsultarGuiaDeRemision(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object
        Dim oGuiaRemision As Object = Nothing
        If TipoWS = "NUBE_4_1" Then
            oGuiaRemision = ConsultarGuiaDeRemision_NUBE_4_1(TipoGR, DocEntry, TipoWS)
        Else
            oGuiaRemision = ConsultarGuiaDeRemision_LOCAL_NUBE(TipoGR, DocEntry, TipoWS)
        End If

        Return oGuiaRemision

    End Function
    Public Function ConsultarGuiaDeRemision_LOCAL_NUBE(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oGuiaRemision As Object = Nothing
        Dim oDestinatario As Object
        Dim listaDetinatarios As Object
        Dim listaDetalle As Object
        Dim listaDatosAdicional As Object

        If TipoWS = "LOCAL" Then
            oDestinatario = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemisionDestinatario
            listaDetinatarios = New List(Of Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemisionDestinatario)
            listaDetalle = New List(Of Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemisionDetalle)
            listaDatosAdicional = New List(Of Entidades.wsEDoc_GuiaRemision_LOCAL.ENTDatoAdicionalGuiaRemision)
        Else
            oDestinatario = New Entidades.wsEDoc_GuiaRemision.ENTGuiaRemisionDestinatario
            listaDetinatarios = New List(Of Entidades.wsEDoc_GuiaRemision.ENTGuiaRemisionDestinatario)
            listaDetalle = New List(Of Entidades.wsEDoc_GuiaRemision.ENTGuiaRemisionDetalle)
            listaDatosAdicional = New List(Of Entidades.wsEDoc_GuiaRemision.ENTDatoAdicionalGuiaRemision)
        End If

        Try
            Dim SP As String = ""
            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_ObtenerGuiaDeRemisionEntrega"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_ObtenerGuiaDeRemisionTransferencia"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONENTREGA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONTRANSFERENCIA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONSOLICITUDTRASLADO"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONENTREGA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONTRANSFERENCIA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONENTREGA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONTRANSFERENCIA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONENTREGA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONTRANSFERENCIA"

                    oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                End If
            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                If TipoWS = "LOCAL" Then
                    oGuiaRemision = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemision
                Else
                    oGuiaRemision = New Entidades.wsEDoc_GuiaRemision.ENTGuiaRemision
                End If

                'Dim x As New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemision
                'x.FechaEmision

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oGuiaRemision.ClaveAcceso = Nothing
                            Else
                                oGuiaRemision.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oGuiaRemision.Ambiente = r("Ambiente")
                            oGuiaRemision.TipoEmision = r("TipoEmision")
                            oGuiaRemision.RazonSocial = r("RazonSocial")
                            oGuiaRemision.NombreComercial = r("NombreComercial")
                            oGuiaRemision.Ruc = r("Ruc")
                            oGuiaRemision.CodigoDocumento = r("CodigoDocumento")
                            oGuiaRemision.Establecimiento = r("Establecimiento")
                            oGuiaRemision.PuntoEmision = r("PuntoEmision")
                            oGuiaRemision.Secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemision.Secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemision.Secuencial = oGuiaRemision.Secuencial.PadLeft(9, "0")
                            End If
                            oGuiaRemision.DireccionMatriz = r("DireccionMatriz")
                            oGuiaRemision.FechaEmision = r("FechaEmision")

                            If Not r("AgenteRetencion") = "0" Then
                                oGuiaRemision.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oGuiaRemision.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oGuiaRemision.ContribuyenteRimpe = r("ContribuyenteRimpe")
                            End If

                            oGuiaRemision.DireccionEstablecimiento = r("DireccionEstablecimiento")
                            oGuiaRemision.DireccionPartida = r("DireccionPartida")
                            oGuiaRemision.RazonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemision.TipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemision.RucTranportista = r("RucTranportista")
                            oGuiaRemision.ObligadoContabilidad = r("ObligadoContabilidad")

                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemision.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemision.ContribuyenteEspecial = Nothing
                            End If

                            oGuiaRemision.FechaInicioTransporte = r("FechaInicioTransporte")
                            oGuiaRemision.FechaFinTransporte = r("FechaFinTransporte")
                            oGuiaRemision.Placa = r("Placa")

                            'oGuiaRemision.UsuarioProceso = r("UsuarioCreador")
                            oGuiaRemision.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                            oGuiaRemision.EmailResponsable = r("EmailResponsable")
                            oGuiaRemision.SecuencialERP = r("SecuencialERP")
                            oGuiaRemision.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oGuiaRemision.Estado = r("Estado")
                            oGuiaRemision.Campo1 = r("Campo1")
                            oGuiaRemision.Campo2 = r("Campo2")
                            oGuiaRemision.Campo3 = r("Campo3")

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows

                            If TipoWS = "LOCAL" Then
                                oDestinatario = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemisionDestinatario
                            Else
                                oDestinatario = New Entidades.wsEDoc_GuiaRemision.ENTGuiaRemisionDestinatario
                            End If
                            oDestinatario.IdentificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.RazonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.DirDestinatario = r("DirDestinatario")

                            oDestinatario.MotivoTraslado = r("MotivoTraslado")
                            ' If oDestinatario.MotivoTraslado = "VENTA" Then
                            oDestinatario.CodDocSustento = r("CodDocSustento")
                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.Ruta = r("Ruta")
                            End If
                            oDestinatario.NumDocSustento = r("NumDocSustento")
                            oDestinatario.NumAutDocSustento = r("NumAutDocSustento")
                            '  oDestinatario.FechaEmisionDocSustentoSpecified = True
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.FechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            'Else
                            '    oDestinatario.FechaEmisionDocSustentoSpecified = False
                            'End If
                            listaDetinatarios.Add(oDestinatario)

                        Next

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As Object
                            If TipoWS = "LOCAL" Then
                                itemDetalle = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemisionDetalle
                            Else
                                itemDetalle = New Entidades.wsEDoc_GuiaRemision.ENTGuiaRemisionDetalle
                            End If
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.CodigoInterno = r("CodigoPrincipal")
                            itemDetalle.CodigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.Descripcion = r("Descripcion")
                            itemDetalle.Cantidad = r("Cantidad")

                            Dim listaDetalleDatoAdicional As Object
                            'Adicional1

                            If TipoWS = "LOCAL" Then
                                listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision_LOCAL.ENTDatoAdicionalGuiaRemisionDetalle)
                            Else
                                listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision.ENTDatoAdicionalGuiaRemisionDetalle)
                            End If
                            'Adicional 1
                            If Not r("ConceptoAdicional1") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                If TipoWS = "LOCAL" Then
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTDatoAdicionalGuiaRemisionDetalle
                                Else
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision.ENTDatoAdicionalGuiaRemisionDetalle
                                End If
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional2") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                If TipoWS = "LOCAL" Then
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTDatoAdicionalGuiaRemisionDetalle
                                Else
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision.ENTDatoAdicionalGuiaRemisionDetalle
                                End If
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional3") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                If TipoWS = "LOCAL" Then
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTDatoAdicionalGuiaRemisionDetalle
                                Else
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision.ENTDatoAdicionalGuiaRemisionDetalle
                                End If
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            itemDetalle.ENTDatoAdicionalGuiaRemisionDetalle = listaDetalleDatoAdicional.ToArray

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalle)
                        Next
                        oDestinatario.ENTGuiaRemisionDetalle = listaDetalle.ToArray
                        oGuiaRemision.ENTGuiaRemisionDestinatario = listaDetinatarios.ToArray()
                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As Object
                            If TipoWS = "LOCAL" Then
                                itemDatoAdicionalFac = New Entidades.wsEDoc_GuiaRemision_LOCAL.ENTDatoAdicionalGuiaRemision
                            Else
                                itemDatoAdicionalFac = New Entidades.wsEDoc_GuiaRemision.ENTDatoAdicionalGuiaRemision
                            End If
                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oGuiaRemision.ENTDatoAdicionalGuiaRemision = listaDatosAdicional.ToArray
                    End If
                Next
            End If


            'SERIALIZACION OBJETO XML
            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oGuiaRemision.Secuencial.ToString() + oGuiaRemision.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")

                    Dim x As XmlSerializer = Nothing

                    If TipoWS = "LOCAL" Then
                        x = New XmlSerializer(GetType(Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemision))
                    Else
                        x = New XmlSerializer(GetType(Entidades.wsEDoc_GuiaRemision.ENTGuiaRemision))
                    End If

                    Dim writer As TextWriter = New StreamWriter(sRuta)
                    x.Serialize(writer, oGuiaRemision)
                    writer.Close()
                    Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try


            Return oGuiaRemision
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Guia de Remisión en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

            oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "ArgumentException-Error al Consultar Guia de Remisión con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la oGuiaRemision en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If

            oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Error al Consultar Guia de Remisión con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try


    End Function
    Public Function ConsultarGuiaDeRemision_NUBE_4_1(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oGuiaRemision As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemision
        Dim oDestinatario As Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario
        Dim listaDetinatarios As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario)
        Dim listaDetalle As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle)
        Dim listaDatosAdicional As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision)

        Try
            Dim SP As String = ""
            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_ObtenerGuiaDeRemisionEntrega_4_3 "
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_ObtenerGuiaDeRemisionTransferencia_4_3 "
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_OBTENERGUIAREMISIONSOLICITUDTRASLADO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If

            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONENTREGA_4_3 "
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONTRANSFERENCIA_4_3 "
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_ONE_OBTENERGUIAREMISIONSOLICITUDTRASLADO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONENTREGA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONTRANSFERENCIA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_HEI_OBTENERGUIAREMISIONSOLICITUDTRASLADO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONENTREGA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONTRANSFERENCIA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_SYP_OBTENERGUIAREMISIONSOLICITUDTRASLADO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONENTREGA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONTRANSFERENCIA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_TM_OBTENERGUIAREMISIONSOLICITUDTRASLADO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                End If

            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If TipoGR = "GRE" Then
                    SP = "GS_SAP_FE_SS_OBTENERGUIAREMISIONENTREGA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If
                ElseIf TipoGR = "TRE" Then
                    SP = "GS_SAP_FE_SS_OBTENERGUIAREMISIONTRANSFERENCIA_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                ElseIf TipoGR = "TLE" Then
                    SP = "GS_SAP_FE_SS_OBTENERGUIAREMISIONSOLICITUDTRASLADO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If


            End If


            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP



            If TipoGR = "GRE" Then
                SP = GetQueryConsulta(tipoDocumento.GuiaRemisionEntrega, DocEntry)
            ElseIf TipoGR = "TRE" Then
                SP = GetQueryConsulta(tipoDocumento.GuiaRemisionTraslado, DocEntry)
            ElseIf TipoGR = "TLE" Then
                SP = GetQueryConsulta(tipoDocumento.GuiaRemisionSolicitudTraslado, DocEntry)
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


                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")

                Dim SPs() As String = Split(SP, "--*")


                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

                Dim ds1, ds2, ds3 As DataSet
                Dim dt1, dt2, dt3 As DataTable


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Destinatario"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "Detalle"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "Adicionales"
                ds.Tables.Add(dt3)


            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS



            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "3") Then
                    Return Nothing
                End If
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oGuiaRemision.ClaveAcceso = Nothing
                            Else
                                oGuiaRemision.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oGuiaRemision.Ambiente = r("Ambiente")
                            oGuiaRemision.TipoEmision = r("TipoEmision")
                            oGuiaRemision.RazonSocial = r("RazonSocial")

                            If Not r("NombreComercial") = "" Then
                                oGuiaRemision.NombreComercial = r("NombreComercial")
                            End If
                            oGuiaRemision.Ruc = r("Ruc")
                            oGuiaRemision.CodigoDocumento = r("CodigoDocumento")
                            oGuiaRemision.Establecimiento = r("Establecimiento")
                            oGuiaRemision.PuntoEmision = r("PuntoEmision")
                            oGuiaRemision.Secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemision.Secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemision.Secuencial = oGuiaRemision.Secuencial.PadLeft(9, "0")
                            End If
                            oGuiaRemision.DireccionMatriz = r("DireccionMatriz")
                            oGuiaRemision.DireccionEstablecimiento = r("DireccionEstablecimiento")
                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemision.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemision.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oGuiaRemision.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oGuiaRemision.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oGuiaRemision.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            End If

                            oGuiaRemision.ObligadoContabilidad = r("ObligadoContabilidad")

                            'oGuiaRemision.FechaEmision = r("FechaEmision")
                            oGuiaRemision.FechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")
                            oGuiaRemision.DireccionPartida = r("DireccionPartida")
                            oGuiaRemision.RazonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemision.TipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemision.RucTranportista = r("RucTranportista")


                            'oGuiaRemision.FechaInicioTransporte = r("FechaInicioTransporte")
                            'oGuiaRemision.FechaFinTransporte = r("FechaFinTransporte")
                            oGuiaRemision.FechaInicioTransporte = CDate(r("FechaInicioTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.FechaFinTransporte = CDate(r("FechaFinTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.Placa = r("Placa")


                            oGuiaRemision.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                            oGuiaRemision.EmailResponsable = r("EmailResponsable")
                            oGuiaRemision.SecuencialERP = r("SecuencialERP")
                            oGuiaRemision.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oGuiaRemision.Estado = r("Estado")
                            oGuiaRemision.Campo1 = r("Campo1")
                            oGuiaRemision.Campo2 = r("Campo2")
                            oGuiaRemision.Campo3 = r("Campo3")

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows
                            oDestinatario = New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario

                            oDestinatario.IdentificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.RazonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.DirDestinatario = r("DirDestinatario")

                            oDestinatario.MotivoTraslado = r("MotivoTraslado")
                            oDestinatario.CodEstabDestino = r("CodEstabDestino")

                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.Ruta = r("Ruta")
                            End If

                            ' If oDestinatario.MotivoTraslado = "VENTA" Then
                            oDestinatario.CodDocSustento = r("CodDocSustento")
                            oDestinatario.NumDocSustento = r("NumDocSustento")
                            oDestinatario.NumAutDocSustento = r("NumAutDocSustento")
                            '  oDestinatario.FechaEmisionDocSustentoSpecified = True
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.FechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            listaDetinatarios.Add(oDestinatario)

                        Next

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.CodigoInterno = r("CodigoPrincipal")
                            itemDetalle.CodigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.Descripcion = r("Descripcion")
                            itemDetalle.Cantidad = r("Cantidad")

                            Dim listaDetalleDatoAdicional As Object
                            listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle)

                            'Adicional 1
                            If Not r("ConceptoAdicional1") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional2") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional3") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            itemDetalle.ENTDatoAdicionalGuiaRemisionDetalle = listaDetalleDatoAdicional.ToArray

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalle)
                        Next

                        oDestinatario.ENTGuiaRemisionDetalle = listaDetalle.ToArray

                        oGuiaRemision.ENTGuiaRemisionDestinatario = listaDetinatarios.ToArray()
                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision
                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oGuiaRemision.ENTDatoAdicionalGuiaRemision = listaDatosAdicional.ToArray
                    End If
                Next
            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'Dim sRutaCarpeta As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
                Dim sRuta As String = sRutaCarpeta & oGuiaRemision.Secuencial.ToString() + oGuiaRemision.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oGuiaRemision
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
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try


    End Function

    Public Function ConsultarGuiaDesatendida_NUBE_4_1(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oGuiaRemision As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemision
        Dim oDestinatario As Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario
        Dim listaDetinatarios As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario)
        Dim listaDetalle As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle)
        Dim listaDatosAdicional As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision)

        Try
            Dim SP As String = ""

            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP

            SP = GetQueryConsulta(tipoDocumento.GuiaRemisionDesatendida, DocEntry)

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


                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")

                Dim SPs() As String = Split(SP, "--*")


                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

                Dim ds1, ds2, ds3 As DataSet
                Dim dt1, dt2, dt3 As DataTable


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Destinatario"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "Detalle"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "Adicionales"
                ds.Tables.Add(dt3)


            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS



            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "3") Then
                    Return Nothing
                End If
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oGuiaRemision.ClaveAcceso = Nothing
                            Else
                                oGuiaRemision.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oGuiaRemision.Ambiente = r("Ambiente")
                            oGuiaRemision.TipoEmision = r("TipoEmision")
                            oGuiaRemision.RazonSocial = r("RazonSocial")

                            If Not r("NombreComercial") = "" Then
                                oGuiaRemision.NombreComercial = r("NombreComercial")
                            End If
                            oGuiaRemision.Ruc = r("Ruc")
                            oGuiaRemision.CodigoDocumento = r("CodigoDocumento")
                            oGuiaRemision.Establecimiento = r("Establecimiento")
                            oGuiaRemision.PuntoEmision = r("PuntoEmision")
                            oGuiaRemision.Secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemision.Secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemision.Secuencial = oGuiaRemision.Secuencial.PadLeft(9, "0")
                            End If
                            oGuiaRemision.DireccionMatriz = r("DireccionMatriz")
                            oGuiaRemision.DireccionEstablecimiento = r("DireccionEstablecimiento")
                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemision.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemision.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oGuiaRemision.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oGuiaRemision.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oGuiaRemision.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            End If

                            oGuiaRemision.ObligadoContabilidad = r("ObligadoContabilidad")

                            'oGuiaRemision.FechaEmision = r("FechaEmision")
                            oGuiaRemision.FechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")
                            oGuiaRemision.DireccionPartida = r("DireccionPartida")
                            oGuiaRemision.RazonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemision.TipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemision.RucTranportista = r("RucTranportista")


                            'oGuiaRemision.FechaInicioTransporte = r("FechaInicioTransporte")
                            'oGuiaRemision.FechaFinTransporte = r("FechaFinTransporte")
                            oGuiaRemision.FechaInicioTransporte = CDate(r("FechaInicioTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.FechaFinTransporte = CDate(r("FechaFinTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.Placa = r("Placa")


                            oGuiaRemision.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                            oGuiaRemision.EmailResponsable = r("EmailResponsable")
                            oGuiaRemision.SecuencialERP = r("SecuencialERP")
                            oGuiaRemision.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oGuiaRemision.Estado = r("Estado")
                            oGuiaRemision.Campo1 = r("Campo1")
                            oGuiaRemision.Campo2 = r("Campo2")
                            oGuiaRemision.Campo3 = r("Campo3")

                            'NEW DM 08012025
                            oGuiaRemision.Campo4 = r("Campo4")
                            oGuiaRemision.Campo5 = r("Campo5")
                            oGuiaRemision.Campo6 = r("Campo6")
                            oGuiaRemision.Campo7 = r("Campo7")
                            oGuiaRemision.Campo8 = r("Campo8")
                            oGuiaRemision.Campo9 = r("Campo9")
                            oGuiaRemision.Campo10 = r("Campo10")
                            oGuiaRemision.Campo11 = r("Campo11")
                            oGuiaRemision.Campo12 = r("Campo12")
                            oGuiaRemision.Campo13 = r("Campo13")
                            oGuiaRemision.Campo14 = r("Campo14")
                            oGuiaRemision.Campo15 = r("Campo15")

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows
                            oDestinatario = New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario

                            oDestinatario.IdentificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.RazonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.DirDestinatario = r("DirDestinatario")

                            oDestinatario.MotivoTraslado = r("MotivoTraslado")
                            oDestinatario.CodEstabDestino = r("CodEstabDestino")

                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.Ruta = r("Ruta")
                            End If

                            ' If oDestinatario.MotivoTraslado = "VENTA" Then
                            oDestinatario.CodDocSustento = r("CodDocSustento")
                            oDestinatario.NumDocSustento = r("NumDocSustento")
                            oDestinatario.NumAutDocSustento = r("NumAutDocSustento")
                            '  oDestinatario.FechaEmisionDocSustentoSpecified = True
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.FechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            listaDetinatarios.Add(oDestinatario)

                        Next

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.CodigoInterno = r("CodigoPrincipal")
                            itemDetalle.CodigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.Descripcion = r("Descripcion")
                            itemDetalle.Cantidad = r("Cantidad")

                            Dim listaDetalleDatoAdicional As Object
                            listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle)

                            'Adicional 1
                            If Not r("ConceptoAdicional1") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional2") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional3") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            itemDetalle.ENTDatoAdicionalGuiaRemisionDetalle = listaDetalleDatoAdicional.ToArray

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalle)
                        Next

                        oDestinatario.ENTGuiaRemisionDetalle = listaDetalle.ToArray

                        oGuiaRemision.ENTGuiaRemisionDestinatario = listaDetinatarios.ToArray()
                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision
                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oGuiaRemision.ENTDatoAdicionalGuiaRemision = listaDatosAdicional.ToArray
                    End If
                Next
            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'Dim sRutaCarpeta As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
                Dim sRuta As String = sRutaCarpeta & oGuiaRemision.Secuencial.ToString() + oGuiaRemision.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oGuiaRemision
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
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try


    End Function



    Public Function Consultar_Factura_GuiaDeRemision(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oGuiaRemision As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemision
        Dim oDestinatario As Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario
        Dim listaDetinatarios As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario)
        Dim listaDetalle As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle)
        Dim listaDatosAdicional As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision)

        Try
            Dim SP As String = ""

            SP = "GS_SAP_FE_HEI_OBTENERFACTURAGUIAREMISIONENTREGA_4_3"
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If



            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP



            'If TipoGR = "GRE" Then
            '    SP = GetQueryConsulta(tipoDocumento.GuiaRemisionEntrega, DocEntry)
            'ElseIf TipoGR = "TRE" Then
            '    SP = GetQueryConsulta(tipoDocumento.GuiaRemisionTraslado, DocEntry)
            'ElseIf TipoGR = "TLE" Then
            '    SP = GetQueryConsulta(tipoDocumento.GuiaRemisionSolicitudTraslado, DocEntry)
            'End If

            SP = GetQueryConsulta(tipoDocumento.GuiaRemisionEntrega, DocEntry)

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


                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")

                Dim SPs() As String = Split(SP, "--*")


                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

                Dim ds1, ds2, ds3 As DataSet
                Dim dt1, dt2, dt3 As DataTable


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Destinatario"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "Detalle"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "Adicionales"
                ds.Tables.Add(dt3)


            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS


            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oGuiaRemision.ClaveAcceso = Nothing
                            Else
                                oGuiaRemision.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oGuiaRemision.Ambiente = r("Ambiente")
                            oGuiaRemision.TipoEmision = r("TipoEmision")
                            oGuiaRemision.RazonSocial = r("RazonSocial")

                            If Not r("NombreComercial") = "" Then
                                oGuiaRemision.NombreComercial = r("NombreComercial")
                            End If
                            oGuiaRemision.Ruc = r("Ruc")
                            oGuiaRemision.CodigoDocumento = r("CodigoDocumento")
                            oGuiaRemision.Establecimiento = r("Establecimiento")
                            oGuiaRemision.PuntoEmision = r("PuntoEmision")
                            oGuiaRemision.Secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemision.Secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemision.Secuencial = oGuiaRemision.Secuencial.PadLeft(9, "0")
                            End If
                            oGuiaRemision.DireccionMatriz = r("DireccionMatriz")
                            oGuiaRemision.DireccionEstablecimiento = r("DireccionEstablecimiento")
                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemision.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemision.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oGuiaRemision.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oGuiaRemision.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oGuiaRemision.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            End If

                            oGuiaRemision.ObligadoContabilidad = r("ObligadoContabilidad")

                            'oGuiaRemision.FechaEmision = r("FechaEmision")
                            oGuiaRemision.FechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")

                            oGuiaRemision.DireccionPartida = r("DireccionPartida")
                            oGuiaRemision.RazonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemision.TipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemision.RucTranportista = r("RucTranportista")


                            'oGuiaRemision.FechaInicioTransporte = r("FechaInicioTransporte")
                            'oGuiaRemision.FechaFinTransporte = r("FechaFinTransporte")
                            oGuiaRemision.FechaInicioTransporte = CDate(r("FechaInicioTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.FechaFinTransporte = CDate(r("FechaFinTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.Placa = r("Placa")


                            oGuiaRemision.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                            oGuiaRemision.EmailResponsable = r("EmailResponsable")
                            oGuiaRemision.SecuencialERP = r("SecuencialERP")
                            oGuiaRemision.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oGuiaRemision.Estado = r("Estado")
                            oGuiaRemision.Campo1 = r("Campo1")
                            oGuiaRemision.Campo2 = r("Campo2")
                            oGuiaRemision.Campo3 = r("Campo3")

                            'ADD DM 08012025
                            oGuiaRemision.Campo4 = r("Campo4")
                            oGuiaRemision.Campo5 = r("Campo5")
                            oGuiaRemision.Campo6 = r("Campo6")
                            oGuiaRemision.Campo7 = r("Campo7")
                            oGuiaRemision.Campo8 = r("Campo8")
                            oGuiaRemision.Campo9 = r("Campo9")
                            oGuiaRemision.Campo10 = r("Campo10")
                            oGuiaRemision.Campo11 = r("Campo11")
                            oGuiaRemision.Campo12 = r("Campo12")
                            oGuiaRemision.Campo13 = r("Campo13")
                            oGuiaRemision.Campo14 = r("Campo14")
                            oGuiaRemision.Campo15 = r("Campo15")
                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows
                            oDestinatario = New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario

                            oDestinatario.IdentificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.RazonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.DirDestinatario = r("DirDestinatario")

                            oDestinatario.MotivoTraslado = r("MotivoTraslado")
                            oDestinatario.CodEstabDestino = r("CodEstabDestino")

                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.Ruta = r("Ruta")
                            End If

                            ' If oDestinatario.MotivoTraslado = "VENTA" Then
                            oDestinatario.CodDocSustento = r("CodDocSustento")
                            oDestinatario.NumDocSustento = r("NumDocSustento")
                            oDestinatario.NumAutDocSustento = r("NumAutDocSustento")
                            '  oDestinatario.FechaEmisionDocSustentoSpecified = True
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.FechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            listaDetinatarios.Add(oDestinatario)

                        Next

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.CodigoInterno = r("CodigoPrincipal")
                            itemDetalle.CodigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.Descripcion = r("Descripcion")
                            itemDetalle.Cantidad = r("Cantidad")

                            Dim listaDetalleDatoAdicional As Object
                            listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle)

                            'Adicional 1
                            If Not r("ConceptoAdicional1") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional2") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional3") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            itemDetalle.ENTDatoAdicionalGuiaRemisionDetalle = listaDetalleDatoAdicional.ToArray

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalle)
                        Next

                        oDestinatario.ENTGuiaRemisionDetalle = listaDetalle.ToArray
                        oGuiaRemision.ENTGuiaRemisionDestinatario = listaDetinatarios.ToArray()

                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision
                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oGuiaRemision.ENTDatoAdicionalGuiaRemision = listaDatosAdicional.ToArray
                    End If
                Next
            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oGuiaRemision.Secuencial.ToString() + oGuiaRemision.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oGuiaRemision
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
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try


    End Function

    Public Function Consultar_SalidaMercancias_GuiaDeRemision(ByVal TipoGR As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oGuiaRemision As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemision
        Dim oDestinatario As Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario
        Dim listaDetinatarios As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario)
        Dim listaDetalle As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle)
        Dim listaDatosAdicional As New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision)

        Try
            Dim SP As String = ""

            SP = "GS_SAP_FE_HEI_OBTENERSALIDAMERCANCIASGUIAREMISIONENTREGA_4_3"
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoGR, DocEntry, "Consultando Guía de Remisión con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If


            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP



            'If TipoGR = "GRE" Then
            '    SP = GetQueryConsulta(tipoDocumento.GuiaRemisionEntrega, DocEntry)
            'ElseIf TipoGR = "TRE" Then
            '    SP = GetQueryConsulta(tipoDocumento.GuiaRemisionTraslado, DocEntry)
            'ElseIf TipoGR = "TLE" Then
            '    SP = GetQueryConsulta(tipoDocumento.GuiaRemisionSolicitudTraslado, DocEntry)
            'End If


            SP = GetQueryConsulta(tipoDocumento.GuiaRemisionEntrega, DocEntry)

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


                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")

                Dim SPs() As String = Split(SP, "--*")


                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

                Dim ds1, ds2, ds3 As DataSet
                Dim dt1, dt2, dt3 As DataTable


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Destinatario"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "Detalle"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "Adicionales"
                ds.Tables.Add(dt3)


            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS


            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oGuiaRemision.ClaveAcceso = Nothing
                            Else
                                oGuiaRemision.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oGuiaRemision.Ambiente = r("Ambiente")
                            oGuiaRemision.TipoEmision = r("TipoEmision")
                            oGuiaRemision.RazonSocial = r("RazonSocial")

                            If Not r("NombreComercial") = "" Then
                                oGuiaRemision.NombreComercial = r("NombreComercial")
                            End If
                            oGuiaRemision.Ruc = r("Ruc")
                            oGuiaRemision.CodigoDocumento = r("CodigoDocumento")
                            oGuiaRemision.Establecimiento = r("Establecimiento")
                            oGuiaRemision.PuntoEmision = r("PuntoEmision")
                            oGuiaRemision.Secuencial = r("SecuencialDocumento")
                            If Not oGuiaRemision.Secuencial.ToString().Length.Equals("9") Then
                                oGuiaRemision.Secuencial = oGuiaRemision.Secuencial.PadLeft(9, "0")
                            End If
                            oGuiaRemision.DireccionMatriz = r("DireccionMatriz")
                            oGuiaRemision.DireccionEstablecimiento = r("DireccionEstablecimiento")
                            If Not r("ContribuyenteEspecial") = "0" Then
                                oGuiaRemision.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oGuiaRemision.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oGuiaRemision.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oGuiaRemision.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oGuiaRemision.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            End If

                            oGuiaRemision.ObligadoContabilidad = r("ObligadoContabilidad")

                            'oGuiaRemision.FechaEmision = r("FechaEmision")
                            oGuiaRemision.FechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")

                            oGuiaRemision.DireccionPartida = r("DireccionPartida")
                            oGuiaRemision.RazonSocialTransportista = r("RazonSocialTransportista")
                            oGuiaRemision.TipoIdentificacionTransportista = r("TipoIdentificacionTransportista")
                            oGuiaRemision.RucTranportista = r("RucTranportista")


                            'oGuiaRemision.FechaInicioTransporte = r("FechaInicioTransporte")
                            'oGuiaRemision.FechaFinTransporte = r("FechaFinTransporte")
                            oGuiaRemision.FechaInicioTransporte = CDate(r("FechaInicioTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.FechaFinTransporte = CDate(r("FechaFinTransporte")).ToString("yyyy-MM-dd")
                            oGuiaRemision.Placa = r("Placa")


                            oGuiaRemision.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                            oGuiaRemision.EmailResponsable = r("EmailResponsable")
                            oGuiaRemision.SecuencialERP = r("SecuencialERP")
                            oGuiaRemision.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oGuiaRemision.Estado = r("Estado")
                            oGuiaRemision.Campo1 = r("Campo1")
                            oGuiaRemision.Campo2 = r("Campo2")
                            oGuiaRemision.Campo3 = r("Campo3")

                            'ADD DM 08012025
                            oGuiaRemision.Campo4 = r("Campo4")
                            oGuiaRemision.Campo5 = r("Campo5")
                            oGuiaRemision.Campo6 = r("Campo6")
                            oGuiaRemision.Campo7 = r("Campo7")
                            oGuiaRemision.Campo8 = r("Campo8")
                            oGuiaRemision.Campo9 = r("Campo9")
                            oGuiaRemision.Campo10 = r("Campo10")
                            oGuiaRemision.Campo11 = r("Campo11")
                            oGuiaRemision.Campo12 = r("Campo12")
                            oGuiaRemision.Campo13 = r("Campo13")
                            oGuiaRemision.Campo14 = r("Campo14")
                            oGuiaRemision.Campo15 = r("Campo15")

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows
                            oDestinatario = New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDestinatario

                            oDestinatario.IdentificacionDestinatario = r("IdentificacionDestinatario")
                            oDestinatario.RazonSocialDestinatario = r("RazonSocialDestinatario")
                            oDestinatario.DirDestinatario = r("DirDestinatario")

                            oDestinatario.MotivoTraslado = r("MotivoTraslado")
                            oDestinatario.CodEstabDestino = r("CodEstabDestino")

                            If Not r("Ruta").ToString() = "" Then
                                oDestinatario.Ruta = r("Ruta")
                            End If

                            ' If oDestinatario.MotivoTraslado = "VENTA" Then
                            oDestinatario.CodDocSustento = r("CodDocSustento")
                            oDestinatario.NumDocSustento = r("NumDocSustento")
                            oDestinatario.NumAutDocSustento = r("NumAutDocSustento")
                            '  oDestinatario.FechaEmisionDocSustentoSpecified = True
                            If Not r("FechaEmisionDocSustento").ToString() = "" Then
                                oDestinatario.FechaEmisionDocSustento = r("FechaEmisionDocSustento")
                            End If

                            listaDetinatarios.Add(oDestinatario)

                        Next

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDetalle As New Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemisionDetalle
                            'itemDetalle.CantidadSpecified = True

                            itemDetalle.CodigoInterno = r("CodigoPrincipal")
                            itemDetalle.CodigoAdicional = r("CodigoAuxiliar")
                            itemDetalle.Descripcion = r("Descripcion")
                            itemDetalle.Cantidad = r("Cantidad")

                            Dim listaDetalleDatoAdicional As Object
                            listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle)

                            'Adicional 1
                            If Not r("ConceptoAdicional1") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional2") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional2")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional2")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            If Not r("ConceptoAdicional3") = "0" Then
                                Dim itemDetalleDatoAdicional As Object
                                itemDetalleDatoAdicional = New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemisionDetalle
                                itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional3")
                                itemDetalleDatoAdicional.Descripcion = r("NombreAdicional3")
                                listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                            End If

                            itemDetalle.ENTDatoAdicionalGuiaRemisionDetalle = listaDetalleDatoAdicional.ToArray

                            'agrego detalle a la lista
                            listaDetalle.Add(itemDetalle)
                        Next

                        oDestinatario.ENTGuiaRemisionDetalle = listaDetalle.ToArray
                        oGuiaRemision.ENTGuiaRemisionDestinatario = listaDetinatarios.ToArray()

                    ElseIf i = 3 Then
                        For Each r As DataRow In ds.Tables(3).Rows
                            Dim itemDatoAdicionalFac As New Entidades.wsEDoc_GuiaRemision41.ENTDatoAdicionalGuiaRemision
                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oGuiaRemision.ENTDatoAdicionalGuiaRemision = listaDatosAdicional.ToArray
                    End If
                Next
            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oGuiaRemision.Secuencial.ToString() + oGuiaRemision.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oGuiaRemision.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oGuiaRemision)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oGuiaRemision
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
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_MainDocumentos/agregarControl", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try


    End Function

    Public Function ConsultarRetencionND(ByVal TipoRE As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oRetencion As Object = Nothing
        Dim lstimpfact As Object
        Dim listaDatosAdicional As Object


        Dim listaRetDocSus As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustento)
        Dim listaRetPago As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoPago)
        lstimpfact = New List(Of Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion)
        listaDatosAdicional = New List(Of Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion)
        Dim listaRetDocSusRet As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoRetencion)

        Dim oRetencionDocSustento As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustento
        Dim ListoRetencionDocSustentoReem As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolso)

        Try
            Dim SP As String = ""
            If TipoRE = "RDM" Then
                SP = "GS_SAP_FE_ObtenerRetencionNotaDebito_4_3"
                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención ND con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
            End If

            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP

            'If Functions.VariablesGlobales._vgGuardarLog = "Y" Then
            '    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

            '    oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Tipo de factura = " + TipoFactura.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '    oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Consultando Factura con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            '    'End If
            'End If


            If TipoRE = "RDM" Then

                SP = GetQueryConsulta(tipoDocumento.RetencionNotaDebito, DocEntry)

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

                Dim ds1, ds2, ds3, ds4 As DataSet
                Dim dt1, dt2, dt3, dt4 As DataTable


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Sustento"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "SustentoRetencion"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "SustentoReembolso"
                ds.Tables.Add(dt3)

                ds4 = EjecutarSP(SPs(4).ToString(), DocEntry)
                dt4 = ds4.Tables(0).Copy
                dt4.TableName = "Adicionales"
                ds.Tables.Add(dt4)


            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS



            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oRetencion = New Entidades.wsEDoc_Retencion41.ENTRetencion

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oRetencion.ClaveAcceso = Nothing
                            Else
                                oRetencion.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oRetencion.Ambiente = r("Ambiente")
                            oRetencion.TipoEmision = r("TipoEmision")

                            oRetencion.RazonSocial = r("RazonSocial")
                            If Not r("NombreComercial") = "" Then
                                oRetencion.NombreComercial = r("NombreComercial")
                            End If
                            oRetencion.Ruc = r("Ruc")

                            oRetencion.CodigoDocumento = r("CodigoDocumento")
                            oRetencion.Establecimiento = r("Establecimiento")
                            oRetencion.PuntoEmision = r("PuntoEmision")
                            oRetencion.Secuencial = r("SecuencialDocumento")

                            If Not oRetencion.Secuencial.ToString().Length.Equals("9") Then
                                oRetencion.Secuencial = oRetencion.Secuencial.ToString().PadLeft(9, "0")
                            End If
                            Utilitario.Util_Log.Escribir_Log("oRetencion.Secuencial : " & oRetencion.Secuencial.ToString(), "ManejoDeDocumentos")

                            oRetencion.DireccionMatriz = r("DireccionMatriz")

                            oRetencion.FechaEmision = r("FechaEmision")
                            oRetencion.DireccionEstablecimiento = r("DireccionEstablecimiento")

                            If Not r("ContribuyenteEspecial") = "0" Then
                                oRetencion.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oRetencion.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oRetencion.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oRetencion.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oRetencion.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            End If

                            oRetencion.ObligadoContabilidad = r("ObligadoContabilidad")

                            oRetencion.TipoIdentificacionSujetoRetenido = r("TipoIdentificacionSujetoRetenido")
                            oRetencion.RazonSocialSujetoRetenido = r("RazonSocialSujetoRetenido")
                            oRetencion.IdentificacionSujetoRetenido = r("IdentificacionSujetoRetenido")
                            oRetencion.PeriodoFiscal = r("PeriodoFiscal")

                            oRetencion.BaseImponible = r("TotalBaseImponible")
                            oRetencion.TotalRetencion = r("TotalRetencion")

                            oRetencion.UsuarioTransaccionERP = r("UsuarioCreador")
                            oRetencion.EmailResponsable = r("EmailResponsable")
                            oRetencion.SecuencialERP = r("SecuencialERP")
                            oRetencion.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oRetencion.Estado = r("Estado")
                            oRetencion.Campo1 = r("Campo1")
                            oRetencion.Campo2 = r("Campo2")
                            oRetencion.Campo3 = r("Campo3")

                            If Not r("TipoRetencion") = "0" Then
                                Utilitario.Util_Log.Escribir_Log("TipoRetencion : " & r("TipoRetencion"), "ManejoDeDocumentos")
                                oRetencion.Tipo = r("TipoRetencion")
                                oRetencion.TipoSujetoRetenido = r("TipoSujetoRetenido")
                                oRetencion.ParteRel = r("ParteRel")
                            End If

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows

                            oRetencionDocSustento = New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustento

                            oRetencionDocSustento.CodDocSustento = r("CodDocRetener")
                            oRetencionDocSustento.NumDocSustento = r("NumDocRetener")
                            oRetencionDocSustento.FechaEmisionDocSustento = DirectCast(r("FechaEmisionDocRetener"), Date)

                            If oRetencion.Tipo = 1 Then

                                oRetencionDocSustento.CodSustento = r("CodSustento")
                                oRetencionDocSustento.FechaRegistroContable = r("FechaRegistroContable")
                                oRetencionDocSustento.NumAutDocSustento = r("NumAutDocSustento")
                                oRetencionDocSustento.PagoLocExt = r("PagoLocExt")
                                If oRetencionDocSustento.PagoLocExt = "02" Then
                                    oRetencionDocSustento.TipoRegi = r("TipoRegi")
                                    oRetencionDocSustento.PaisEfecPago = r("PaisEfecPago")
                                    oRetencionDocSustento.AplicConvDobTrib = r("AplicConvDobTrib")
                                    If oRetencionDocSustento.AplicConvDobTrib = "NO" Then
                                        oRetencionDocSustento.PagExtSujRetNorLeg = r("PagExtSujRetNorLeg")
                                    End If
                                    oRetencionDocSustento.PagoRegFis = r("PagoRegFis")
                                End If

                                If r("TotalComprobantesReembolso") <> 0 Then
                                    oRetencionDocSustento.TotalComprobantesReembolso = r("TotalComprobantesReembolso")
                                End If
                                If r("TotalBaseImponibleReembolso") <> 0 Then
                                    oRetencionDocSustento.TotalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                                End If
                                If r("TotalImpuestoReembolso") <> 0 Then
                                    oRetencionDocSustento.TotalImpuestoReembolso = r("TotalImpuestoReembolso")
                                End If

                                oRetencionDocSustento.TotalSinImpuestos = r("TotalSinImpuestos")
                                oRetencionDocSustento.ImporteTotal = r("ImporteTotal")

                                Dim ListRetencionDocSustentoImp As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSus8")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPor8")
                                    impoRetencionDocSustentoImp.BaseImponible = r("Base8")
                                    impoRetencionDocSustentoImp.Tarifa = r("Tarifa8")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuesto8")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSus12")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPor12")
                                    impoRetencionDocSustentoImp.BaseImponible = r("Base12")
                                    impoRetencionDocSustentoImp.Tarifa = r("Tarifa12")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuesto12")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("Base0") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSus0")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPor0")
                                    impoRetencionDocSustentoImp.BaseImponible = r("Base0")
                                    impoRetencionDocSustentoImp.Tarifa = r("Tarifa0")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuesto0")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("BaseNoi") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSusNoi")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPorNoi")
                                    impoRetencionDocSustentoImp.BaseImponible = r("BaseNoi")
                                    impoRetencionDocSustentoImp.Tarifa = r("TarifaNoi")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuestoNoi")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("BaseExen") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSusExen")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPorExen")
                                    impoRetencionDocSustentoImp.BaseImponible = r("BaseExen")
                                    impoRetencionDocSustentoImp.Tarifa = r("TarifaExen")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuestoExen")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("Base5") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSus5")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPor5")
                                    impoRetencionDocSustentoImp.BaseImponible = r("Base5")
                                    impoRetencionDocSustentoImp.Tarifa = r("Tarifa5")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuesto5")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("Base15") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSus15")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPor15")
                                    impoRetencionDocSustentoImp.BaseImponible = r("Base15")
                                    impoRetencionDocSustentoImp.Tarifa = r("Tarifa15")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuesto15")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                If r("Base14") <> 0 Then
                                    Dim impoRetencionDocSustentoImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoImpuesto
                                    impoRetencionDocSustentoImp.CodImpuestoDocSustento = r("CodImpDocSus14")
                                    impoRetencionDocSustentoImp.CodigoPorcentaje = r("CodPor14")
                                    impoRetencionDocSustentoImp.BaseImponible = r("Base14")
                                    impoRetencionDocSustentoImp.Tarifa = r("Tarifa14")
                                    impoRetencionDocSustentoImp.ValorImpuesto = r("ValorImpuesto14")

                                    ListRetencionDocSustentoImp.Add(impoRetencionDocSustentoImp)
                                End If

                                oRetencionDocSustento.ENTRetencionDocSustentoImpuesto = ListRetencionDocSustentoImp.ToArray

                            End If

                            listaRetDocSus.Add(oRetencionDocSustento)

                            Dim retPago As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoPago
                            retPago.FormaPago = r("FormaPago")
                            retPago.Total = r("Total")
                            listaRetPago.Add(retPago)


                        Next
                        oRetencion.ENTRetencionDocSustento = listaRetDocSus.ToArray

                        oRetencionDocSustento.ENTRetencionDocSustentoPago = listaRetPago.ToArray

                    ElseIf i = 2 Then

                        For Each r As DataRow In ds.Tables(2).Rows

                            Dim RetDocSusRet As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoRetencion

                            RetDocSusRet.Codigo = r("Codigo")
                            RetDocSusRet.CodigoRetencion = r("CodigoRetencion")
                            RetDocSusRet.BaseImponible = r("BaseImponible")
                            RetDocSusRet.PorcentajeRetener = r("PorcentajeRetener")
                            RetDocSusRet.ValorRetenido = r("ValorRetenido")

                            If oRetencionDocSustento.CodSustento = "10" Then
                                Utilitario.Util_Log.Escribir_Log("oRetencion.FechaPagoDiv: " & r("FechaPagoDiv").ToString(), "ManejoDeDocumentos")
                                RetDocSusRet.FechaPagoDiv = CDate(r("FechaPagoDiv"))
                                Utilitario.Util_Log.Escribir_Log("oRetencion.ImRentaSoc: " & r("ImRentaSoc").ToString(), "ManejoDeDocumentos")
                                RetDocSusRet.ImRentaSoc = r("ImRentaSoc")
                                Utilitario.Util_Log.Escribir_Log("oRetencion.EjerFisUtDiv: " & r("EjerFisUtDiv").ToString(), "ManejoDeDocumentos")
                                RetDocSusRet.EjerFisUtDiv = r("EjerFisUtDiv")
                            End If

                            listaRetDocSusRet.Add(RetDocSusRet)

                        Next
                        oRetencionDocSustento.ENTRetencionDocSustentoRetencion = listaRetDocSusRet.ToArray


                    ElseIf i = 3 Then

                        For Each r As DataRow In ds.Tables(3).Rows

                            Dim oRetencionDocSustentoReem As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolso
                            oRetencionDocSustentoReem.TipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                            oRetencionDocSustentoReem.IdentificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                            oRetencionDocSustentoReem.CodPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                            oRetencionDocSustentoReem.TipoProveedorReembolso = r("TipoProveedorReembolso")
                            oRetencionDocSustentoReem.CodDocReembolso = r("CodDocReembolso")
                            oRetencionDocSustentoReem.EstabDocReembolso = r("EstabDocReembolso")
                            oRetencionDocSustentoReem.PtoEmiDocReembolso = r("PtoEmiDocReembolso")
                            oRetencionDocSustentoReem.SecuencialDocReembolso = r("SecuencialDocReembolso")
                            oRetencionDocSustentoReem.FechaEmisionDocReembolso = CDate(r("FechaEmisionDocReembolso"))
                            oRetencionDocSustentoReem.NumeroAutorizacionDocReemb = r("NumeroAutorizacionDocReem")

                            Dim listaoRetencionDocSustentoReemImp As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto)

                            If r("Base12") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBase12")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBase12")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBase12")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("Base12")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBase12")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            If r("Base0") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBase0")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBase0")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBase0")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("Base0")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBase0")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            If r("BaseNoi") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBaseNoi")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBaseNoi")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBaseNoi")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("BaseNoi")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBaseNoi")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            If r("BaseExen") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBaseExen")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBaseExen")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBaseExen")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("BaseExen")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBaseExen")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            If r("Base5") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBase5")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBase5")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBase5")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("Base5")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBase5")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            If r("Base15") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBase15")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBase15")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBase15")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("Base15")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBase15")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            If r("Base14") <> 0 Then
                                Dim impoRetencionDocSustentoReemImp As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoReembolsoImpuesto
                                impoRetencionDocSustentoReemImp.Codigo = r("CodigoBase14")
                                impoRetencionDocSustentoReemImp.CodigoPorcentaje = r("CodigoPorcentajeBase14")
                                impoRetencionDocSustentoReemImp.Tarifa = r("TarifaBase14")
                                impoRetencionDocSustentoReemImp.BaseImponibleReembolso = r("Base14")
                                impoRetencionDocSustentoReemImp.ImpuestoReembolso = r("ImpuestoReembolsoBase14")

                                listaoRetencionDocSustentoReemImp.Add(impoRetencionDocSustentoReemImp)
                            End If

                            oRetencionDocSustentoReem.ENTRetencionDocSustentoReembolsoImpuesto = listaoRetencionDocSustentoReemImp.ToArray
                            ListoRetencionDocSustentoReem.Add(oRetencionDocSustentoReem)

                        Next
                        oRetencion.ENTRetencionDocSustento(0).ENTRetencionDocSustentoReembolso = ListoRetencionDocSustentoReem.ToArray

                    ElseIf i = 4 Then
                        For Each r As DataRow In ds.Tables(4).Rows
                            Dim itemDatoAdicionalFac As Object
                            itemDatoAdicionalFac = New Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion

                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oRetencion.ENTDatoAdicionalRetencion = listaDatosAdicional.ToArray
                    End If
                Next



            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'Dim sRutaCarpeta As String = ""
                'If _tipoManejo = "A" Then
                '    sRutaCarpeta = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
                'Else
                '    sRutaCarpeta = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'End If
                Dim sRuta As String = sRutaCarpeta & oRetencion.Secuencial.ToString() + oRetencion.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_Retencion41.ENTRetencion))
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oRetencion)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_Retencion.ENTRetencion))
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oRetencion)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oRetencion
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Retención en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "ArgumentException-Error al Consultar Retención con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la oRetencion en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Error al Consultar Retención con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: RetencionDL/ConsultarRetencion", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

    End Function

    Public Function ConsultarRetencion_NUBE_4_1_ATS(ByVal TipoRE As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oRetencion As Object = Nothing
        Dim lstimpfact As Object
        Dim listaDatosAdicional As Object


        Dim listaRetDocSus As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustento)
        Dim listaRetPago As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoPago)
        lstimpfact = New List(Of Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion)
        listaDatosAdicional = New List(Of Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion)
        Dim listaRetDocSusRet As New List(Of Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoRetencion)

        Dim oRetencionDocSustento As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustento


        Try
            Dim SP As String = ""
            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_ObtenerRetencionAnticipo_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "GS_SAP_FE_ObtenerRetencion_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_ONE_OBTENERRETENCIONANTICIPO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "GS_SAP_FE_ONE_OBTENERRETENCION_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_HEI_OBTENERRETENCIONANTICIPO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "GS_SAP_FE_HEI_OBTENERRETENCION_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_SYP_OBTENERRETENCIONANTICIPO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "GS_SAP_FE_SYP_OBTENERRETENCION_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_TM_OBTENERRETENCIONANTICIPO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "GS_SAP_FE_TM_OBTENERRETENCION_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_SS_OBTENERRETENCIONANTICIPO_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                Else
                    SP = "GS_SAP_FE_SS_OBTENERRETENCION_4_3"
                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Retención con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                End If
            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                oRetencion = New Entidades.wsEDoc_Retencion41.ENTRetencion

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        For Each r As DataRow In ds.Tables(0).Rows

                            ' OFFLINE 14 NOVIEMBRE 2017
                            If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                oRetencion.ClaveAcceso = Nothing
                            Else
                                oRetencion.ClaveAcceso = r("ClaveAcceso")
                            End If

                            oRetencion.Ambiente = r("Ambiente")
                            oRetencion.TipoEmision = r("TipoEmision")

                            oRetencion.RazonSocial = r("RazonSocial")
                            If Not r("NombreComercial") = "" Then
                                oRetencion.NombreComercial = r("NombreComercial")
                            End If

                            oRetencion.Ruc = r("Ruc")

                            oRetencion.CodigoDocumento = r("CodigoDocumento")
                            oRetencion.Establecimiento = r("Establecimiento")
                            oRetencion.PuntoEmision = r("PuntoEmision")
                            oRetencion.Secuencial = r("SecuencialDocumento")

                            If Not oRetencion.Secuencial.ToString().Length.Equals("9") Then
                                oRetencion.Secuencial = oRetencion.Secuencial.ToString().PadLeft(9, "0")
                            End If
                            Utilitario.Util_Log.Escribir_Log("oRetencion.Secuencial : " & oRetencion.Secuencial.ToString(), "ManejoDeDocumentos")

                            oRetencion.DireccionMatriz = r("DireccionMatriz")

                            oRetencion.FechaEmision = r("FechaEmision")
                            oRetencion.DireccionEstablecimiento = r("DireccionEstablecimiento")

                            If Not r("ContribuyenteEspecial") = "0" Then
                                oRetencion.ContribuyenteEspecial = r("ContribuyenteEspecial")
                            Else
                                oRetencion.ContribuyenteEspecial = Nothing
                            End If

                            If Not r("AgenteRetencion") = "0" Then
                                oRetencion.AgenteRetencion = r("AgenteRetencion")
                            End If

                            If Not r("RegimenMicroempresas") = "0" Then
                                oRetencion.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                            End If

                            If Not r("ContribuyenteRimpe") = "0" Then
                                oRetencion.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                            End If

                            oRetencion.ObligadoContabilidad = r("ObligadoContabilidad")

                            oRetencion.TipoIdentificacionSujetoRetenido = r("TipoIdentificacionSujetoRetenido")
                            oRetencion.RazonSocialSujetoRetenido = r("RazonSocialSujetoRetenido")
                            oRetencion.IdentificacionSujetoRetenido = r("IdentificacionSujetoRetenido")
                            oRetencion.PeriodoFiscal = r("PeriodoFiscal")

                            oRetencion.BaseImponible = r("TotalBaseImponible")
                            oRetencion.TotalRetencion = r("TotalRetencion")

                            oRetencion.UsuarioTransaccionERP = r("UsuarioCreador")
                            oRetencion.EmailResponsable = r("EmailResponsable")
                            oRetencion.SecuencialERP = r("SecuencialERP")
                            oRetencion.CodigoTransaccionERP = r("CodigoTransaccionERP")
                            oRetencion.Estado = r("Estado")
                            oRetencion.Campo1 = r("Campo1")
                            oRetencion.Campo2 = r("Campo2")
                            oRetencion.Campo3 = r("Campo3")

                        Next
                    ElseIf i = 1 Then
                        For Each r As DataRow In ds.Tables(1).Rows

                            oRetencionDocSustento = New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustento

                            oRetencionDocSustento.CodDocSustento = r("CodDocRetener")
                            oRetencionDocSustento.NumDocSustento = r("NumDocRetener")
                            oRetencionDocSustento.FechaEmisionDocSustento = DirectCast(r("FechaEmisionDocRetener"), Date)
                            'impRet.FechaEmisionDocRetener = DirectCast(r("FechaEmisionDocRetener"), Date)
                            oRetencionDocSustento.CodSustento = r("CodSustento")
                            oRetencionDocSustento.FechaRegistroContable = r("FechaRegistroContable")
                            oRetencionDocSustento.NumAutDocSustento = r("NumAutDocSustento")
                            oRetencionDocSustento.TotalSinImpuestos = r("TotalSinImpuestos")
                            oRetencionDocSustento.ImporteTotal = r("ImporteTotal")

                            listaRetDocSus.Add(oRetencionDocSustento)


                            Dim RetDocSusRet As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoRetencion


                            RetDocSusRet.Codigo = r("Codigo")
                            RetDocSusRet.CodigoRetencion = r("CodigoRetencion")
                            RetDocSusRet.BaseImponible = r("BaseImponible")
                            RetDocSusRet.PorcentajeRetener = r("PorcentajeRetener")
                            RetDocSusRet.ValorRetenido = r("ValorRetenido")
                            listaRetDocSusRet.Add(RetDocSusRet)



                            Dim retPago As New Entidades.wsEDoc_Retencion41.ENTRetencionDocSustentoPago
                            retPago.FormaPago = r("FormaPago")
                            retPago.Total = r("Total")
                            listaRetPago.Add(retPago)




                        Next
                        oRetencion.ENTRetencionDocSustento = listaRetDocSus.ToArray
                        oRetencionDocSustento.ENTRetencionDocSustentoRetencion = listaRetDocSusRet.ToArray
                        oRetencionDocSustento.ENTRetencionDocSustentoPago = listaRetPago.ToArray

                    ElseIf i = 2 Then
                        For Each r As DataRow In ds.Tables(2).Rows
                            Dim itemDatoAdicionalFac As Object
                            itemDatoAdicionalFac = New Entidades.wsEDoc_Retencion41.ENTDatoAdicionalRetencion

                            itemDatoAdicionalFac.Nombre = r("Concepto")
                            itemDatoAdicionalFac.Descripcion = r("Descripcion")
                            listaDatosAdicional.Add(itemDatoAdicionalFac)
                        Next
                        oRetencion.ENTDatoAdicionalRetencion = listaDatosAdicional.ToArray
                    End If
                Next



            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oRetencion.Secuencial.ToString() + oRetencion.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_Retencion41.ENTRetencion))
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oRetencion)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_Retencion.ENTRetencion))
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oRetencion)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try


            Return oRetencion
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Retención en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "ArgumentException-Error al Consultar Retención con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la oRetencion en la Base, DocEntry:  " & DocEntry.ToString() & "Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Error al Consultar Retención con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: RetencionDL/ConsultarRetencion", ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

    End Function

    Public Function ConsultarRetencion(ByVal TipoRE As String, ByVal DocEntry As Integer) As Object

        Dim oRetencion As New Entidades.RequestRetencion
        Dim listaImpuestos As List(Of Entidades.impuestosRET)
        Dim listaDatosAdicional As List(Of Entidades.infoAdicionalRET)

        listaImpuestos = New List(Of Entidades.impuestosRET)
        listaDatosAdicional = New List(Of Entidades.infoAdicionalRET)

        Dim listaDocsSustento As List(Of Tuple(Of String, String, String))

        Try
            Dim SP As String = ""
            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_ObtenerRetencionAnticipo_4_3"
                Else
                    SP = "GS_SAP_FE_ObtenerRetencion_4_3"
                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_ONE_OBTENERRETENCIONANTICIPO_4_3"
                Else
                    SP = "GS_SAP_FE_ONE_OBTENERRETENCION_4_3"
                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_HEI_OBTENERRETENCIONANTICIPO_4_3"
                Else
                    SP = "GS_SAP_FE_HEI_OBTENERRETENCION_4_3"
                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_SYP_OBTENERRETENCIONANTICIPO_4_3"
                Else
                    SP = "GS_SAP_FE_SYP_OBTENERRETENCION_4_3"
                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_TM_OBTENERRETENCIONANTICIPO_4_3"
                Else
                    SP = "GS_SAP_FE_TM_OBTENERRETENCION_4_3"
                End If
            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If TipoRE = "REA" Then
                    SP = "GS_SAP_FE_SS_OBTENERRETENCIONANTICIPO_4_3"
                Else
                    SP = "GS_SAP_FE_SS_OBTENERRETENCION_4_3"
                End If
            End If

            If _tipoManejo = "A" Then oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, $"Consultando Retención # DocEntry: {DocEntry} SP: {SP}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)

            If TipoRE = "REA" Then
                SP = GetQueryConsulta(tipoDocumento.RetencionAnticipo, DocEntry)
            Else
                SP = GetQueryConsulta(tipoDocumento.Retencion, DocEntry)
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

                Dim ds1, ds2, ds3, ds4 As DataSet
                Dim dt1, dt2, dt3, dt4 As DataTable

                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "DocSustento"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "DocSustentoRetencion"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "DocSustentoReembolso"
                ds.Tables.Add(dt3)

                ds4 = EjecutarSP(SPs(4).ToString(), DocEntry)
                dt4 = ds4.Tables(0).Copy
                dt4.TableName = "Adicionales"
                ds.Tables.Add(dt4)
            Else
                ds = EjecutarSP(SP, DocEntry)
            End If

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "4") Then Return Nothing
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try
                            For Each r As DataRow In ds.Tables(0).Rows

                                oRetencion.infoTributaria.ambiente = r("Ambiente").ToString

                                oRetencion.infoTributaria.claveAcceso = r("ClaveAcceso").ToString

                                oRetencion.infoTributaria.razonSocial = r("RazonSocial").ToString

                                If Not r("NombreComercial") = "" Then oRetencion.infoTributaria.nombreComercial = r("NombreComercial").ToString

                                oRetencion.infoTributaria.ruc = r("Ruc").ToString

                                oRetencion.infoTributaria.tipoEmision = r("TipoEmision").ToString

                                oRetencion.infoTributaria.codDoc = r("CodigoDocumento").ToString

                                oRetencion.infoTributaria.estab = r("Establecimiento").ToString

                                oRetencion.infoTributaria.ptoEmi = r("PuntoEmision").ToString

                                oRetencion.infoTributaria.secuencial = r("SecuencialDocumento").ToString
                                If Not oRetencion.infoTributaria.secuencial.ToString().Length.Equals("9") Then oRetencion.infoTributaria.secuencial = oRetencion.infoTributaria.secuencial.ToString().PadLeft(9, "0")
                                Utilitario.Util_Log.Escribir_Log("oRetencion.Secuencial : " & oRetencion.infoTributaria.secuencial, "ManejoDeDocumentos")

                                oRetencion.infoTributaria.dirMatriz = r("DireccionMatriz").ToString

                                oRetencion.infoTributaria.diaEmission = CDate(r("FechaEmision")).ToString("dd")

                                oRetencion.infoTributaria.mesEmission = CDate(r("FechaEmision")).ToString("MM")

                                oRetencion.infoTributaria.anioEmission = CDate(r("FechaEmision")).ToString("yyyy")

                                oRetencion.infoCompRetencion.fechaEmision = CDate(r("FechaEmision")).ToString("yyyy-MM-dd")

                                oRetencion.infoCompRetencion.dirEstablecimiento = r("DireccionEstablecimiento").ToString

                                If Not r("ContribuyenteEspecial") = "0" Then oRetencion.infoCompRetencion.contribuyenteEspecial = r("ContribuyenteEspecial").ToString

                                oRetencion.infoCompRetencion.obligadoContabilidad = r("ObligadoContabilidad").ToString

                                oRetencion.infoCompRetencion.tipoIdentificacionSujetoRetenido = r("TipoIdentificacionSujetoRetenido").ToString

                                oRetencion.infoCompRetencion.razonSocialSujetoRetenido = r("RazonSocialSujetoRetenido").ToString

                                oRetencion.infoCompRetencion.identificacionSujetoRetenido = r("IdentificacionSujetoRetenido").ToString

                                oRetencion.infoCompRetencion.periodoFiscal = r("PeriodoFiscal").ToString

                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Cabecera Retencion" & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Cabecera Retencion: " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 1 Then

                        Try
                            listaDocsSustento = New List(Of Tuple(Of String, String, String))

                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim CodDocSustento As String = r("CodDocRetener").ToString
                                Dim NumDocSustento As String = r("NumDocRetener").ToString
                                Dim FechaEmisionDocSustento As String = CDate(r("FechaEmisionDocRetener")).ToString("yyyy-MM-dd")

                                listaDocsSustento.Add(Tuple.Create(CodDocSustento, NumDocSustento, FechaEmisionDocSustento))
                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Informacion documento retener " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Informacion documento retener: " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 2 Then

                        Try
                            Dim y As Integer = 0
                            For Each r As DataRow In ds.Tables(2).Rows

                                Dim RetDocSusRet As New Entidades.impuestosRET
                                RetDocSusRet.codigo = r("Codigo").ToString
                                RetDocSusRet.codigoRetencion = r("CodigoRetencion").ToString
                                RetDocSusRet.baseImponible = r("BaseImponible").ToString
                                RetDocSusRet.porcentajeRetener = r("PorcentajeRetener").ToString
                                RetDocSusRet.valorRetenido = r("ValorRetenido").ToString

                                RetDocSusRet.codDocSustento = listaDocsSustento(y).Item1
                                RetDocSusRet.numDocSustento = listaDocsSustento(y).Item2
                                RetDocSusRet.fechaEmisionDocSustento = listaDocsSustento(y).Item3

                                listaImpuestos.Add(RetDocSusRet)
                                y += 1
                            Next
                            oRetencion.impuestos = listaImpuestos
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Informacion Retenciones " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Informacion Retenciones: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 3 Then

                    ElseIf i = 4 Then
                        Try
                            For Each r As DataRow In ds.Tables(4).Rows
                                Dim itemDatoAdicionalRet As Entidades.infoAdicionalRET = New Entidades.infoAdicionalRET
                                itemDatoAdicionalRet.nombre = r("Concepto")
                                itemDatoAdicionalRet.valor = r("Descripcion")
                                listaDatosAdicional.Add(itemDatoAdicionalRet)
                            Next
                            Utilitario.Util_Log.Escribir_Log("Termina info adicional ", "ManejoDeDocumentos")
                            oRetencion.infoAdicional = listaDatosAdicional
                        Catch ex As Exception
                            If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Retencion Informacion Adicional " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            _Error = "Retencion Informacion Adicional: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    End If
                Next

            End If

            Return oRetencion
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"ArgumentException-Ocurrio un error al consultar datos de la Retención en la BD, DocEntry: {DocEntry} Descr: {x.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, $"ArgumentException-Error al Consultar Retención # DocEntry: {DocEntry} Descr: {x.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage($"Ocurrio un error al consultar datos de la oRetencion en la Base, DocEntry: {DocEntry} Descr: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, $"Error al Consultar Retención # DocEntry: {DocEntry}, Descr: {ex.Message}", FuncionesAddon.Transacciones.Creacion, FuncionesAddon.TipoLog.Emision)
            End If
            Return Nothing
        End Try
    End Function

    Public Function ConsultarLiquidacion(ByVal TipoLQ As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object
        Dim oLiquidacion As Object = Nothing

        If TipoWS = "NUBE_4_1" Then
            oLiquidacion = ConsultarLiquidacionCompra(TipoLQ, DocEntry, TipoWS)
        Else
            oLiquidacion = ConsultarLiquidacionCompra_LOCAL_NUBE(TipoLQ, DocEntry, TipoWS)
        End If

        Return oLiquidacion

    End Function

    Public Function ConsultarLiquidacionCompra(ByVal TipoRE As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oLiquidacionCompra As Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompra = Nothing


        Dim listaDetalleLQ As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompra)
        Dim listaDatoAdicionalDetalleLQCompra As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra)
        Dim listaLiquidacionCompraImp As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto)
        Dim liquidacionCompraPagos As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraPagos)
        Dim listaDatosAdicionalesLQ As List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalLiquidacionCompra)

        listaDetalleLQ = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompra)
        listaDatoAdicionalDetalleLQCompra = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra)
        listaLiquidacionCompraImp = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto)
        liquidacionCompraPagos = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraPagos)

        listaDatosAdicionalesLQ = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalLiquidacionCompra)


        Dim listareembolsoLQ As New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolso)
        Dim aplicadoDescuentoAdicional As Boolean = False
        Try

            Dim SP As String = ""
            If TipoRE = "LQE" Then
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "GS_SAP_FE_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "GS_SAP_FE_ONE_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "GS_SAP_FE_HEI_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "GS_SAP_FE_SYP_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "GS_SAP_FE_TM_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "GS_SAP_FE_SS_ObtenerLiquidacionCompra_4_3"
                End If

                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
            End If


            '------------------------------------NUEVA LOGICA QUERYS ENCRYPTADOS---------
            ' 1 RECUPERO QUERY ENCRYPTADOS Y EJECUTO SP

            'If Functions.VariablesGlobales._vgGuardarLog = "Y" Then
            '    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

            '    oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Tipo de factura = " + TipoFactura.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '    oFuncionesAddon.GuardaLOG(TipoFactura, DocEntry, "Consultando Factura con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


            '    'End If
            'End If



            SP = GetQueryConsulta(tipoDocumento.Liquidacion, DocEntry)

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


                'Utilitario.Util_Log.Escribir_Log("Empezando Split..", "ManejoDeDocumentos")

                Dim SPs() As String = Split(SP, "--*")


                ' Este codigo se optimizo mara recuperar los datatables  Arturo 24.03.2020

                Dim ds1, ds2, ds3, ds4 As DataSet
                Dim dt1, dt2, dt3, dt4 As DataTable


                ds = EjecutarSP(SPs(0).ToString(), DocEntry)
                ds.Tables(0).TableName = "Cabecera"

                ds1 = EjecutarSP(SPs(1).ToString(), DocEntry)
                dt1 = ds1.Tables(0).Copy
                dt1.TableName = "Detalles"
                ds.Tables.Add(dt1)

                ds2 = EjecutarSP(SPs(2).ToString(), DocEntry)
                dt2 = ds2.Tables(0).Copy
                dt2.TableName = "Reembolso"
                ds.Tables.Add(dt2)

                ds3 = EjecutarSP(SPs(3).ToString(), DocEntry)
                dt3 = ds3.Tables(0).Copy
                dt3.TableName = "Adicionales"
                ds.Tables.Add(dt3)

                ds4 = EjecutarSP(SPs(4).ToString(), DocEntry)
                dt4 = ds4.Tables(0).Copy
                dt4.TableName = "Pagos"
                ds.Tables.Add(dt4)

            Else

                ds = EjecutarSP(SP, DocEntry)

            End If



            '------------------------FIN QUERYS ENCRIPTADOS



            'Dim ds As DataSet = EjecutarSP(SP, DocEntry)


            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Functions.VariablesGlobales._ValidarCamposNulos = "Y" And _tipoManejo = "A" Then
                If Not ValidarCamposNulos(ds, "3") Then
                    Return Nothing
                End If
            End If

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                oLiquidacionCompra = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompra

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try
                            For Each r As DataRow In ds.Tables(0).Rows

                                ' MANEJO DE FACTURAS DE EXPORTACION Y REEMBOLSO - 2018-02-18
                                ' Indica que tipo de factura es (0.- Normal, 1.- Exportadores, 2.- Reembolsos)
                                'Try
                                '    If r("TipoLiquidacionCompra").ToString() = "" Then
                                '        oLiquidacionCompra.Tipo = 0
                                '    Else
                                '        'oLiquidacionCompra.Tipo = r("TipoLiquidacionCompra")
                                '        oLiquidacionCompra.Tipo = 1
                                '    End If
                                '    Utilitario.Util_Log.Escribir_Log(" (0.- Normal, 1.- Exportadores, 2.- Reembolsos)", "ManejoDeDocumentos")
                                '    Utilitario.Util_Log.Escribir_Log("Tipo Factura : " & oLiquidacionCompra.Tipo.ToString(), "ManejoDeDocumentos")
                                'Catch ex As Exception
                                '    Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                                '    oLiquidacionCompra.Tipo = 0
                                'End Try

                                ' OFFLINE 14 NOVIEMBRE 2017
                                'FAMC 18/02/2019
                                If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                    oLiquidacionCompra.ClaveAcceso = Nothing
                                Else
                                    oLiquidacionCompra.ClaveAcceso = r("ClaveAcceso")
                                End If

                                oLiquidacionCompra.Ambiente = r("Ambiente")
                                oLiquidacionCompra.TipoEmision = r("TipoEmision")
                                oLiquidacionCompra.RazonSocial = r("RazonSocial")
                                If Not r("NombreComercial") = "" Then
                                    oLiquidacionCompra.NombreComercial = r("NombreComercial")
                                End If


                                oLiquidacionCompra.Ruc = r("RUC")
                                'oLiquidacionCompra.Ruc = "0992737964001"
                                oLiquidacionCompra.CodigoDocumento = r("CodigoDocumento")
                                oLiquidacionCompra.Establecimiento = r("Establecimiento")
                                oLiquidacionCompra.PuntoEmision = r("PuntoEmision")
                                oLiquidacionCompra.Secuencial = r("SecuencialDocumento")
                                If Not oLiquidacionCompra.Secuencial.ToString().Length.Equals("9") Then
                                    oLiquidacionCompra.Secuencial = oLiquidacionCompra.Secuencial.ToString().PadLeft(9, "0")
                                End If
                                Utilitario.Util_Log.Escribir_Log("oLiquidacionCompra.Secuencial : " & oLiquidacionCompra.Secuencial.ToString(), "ManejoDeDocumentos")
                                oLiquidacionCompra.DireccionMatriz = r("DireccionMatriz")
                                oLiquidacionCompra.FechaEmision = r("FechaEmision")
                                oLiquidacionCompra.DireccionEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oLiquidacionCompra.ContribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oLiquidacionCompra.ContribuyenteEspecial = Nothing
                                End If

                                If Not r("AgenteRetencion") = "0" Then
                                    oLiquidacionCompra.AgenteRetencion = r("AgenteRetencion")
                                End If

                                If Not r("RegimenMicroempresas") = "0" Then
                                    oLiquidacionCompra.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                                End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    oLiquidacionCompra.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                                End If

                                oLiquidacionCompra.ObligadoContabilidad = r("ObligadoContabilidad")
                                oLiquidacionCompra.TipoIdentificacionProveedor = r("TipoIdentificacionProveedor")

                                'If Not r("GuiaRemision") = "0" Then
                                '    oLiquidacionCompra.GuiaRemision = r("GuiaRemision")
                                'End If

                                oLiquidacionCompra.RazonSocialProveedor = r("RazonSocialProveedor")
                                oLiquidacionCompra.IdentificacionProveedor = r("IdentificacionProveedor")

                                Try
                                    If Not r("DirProveedor") = "" Then
                                        oLiquidacionCompra.DirProveedor = r("DirProveedor")
                                    End If
                                Catch ex As Exception
                                End Try

                                oLiquidacionCompra.TotalSinImpuesto = r("TotalSinImpuesto")
                                oLiquidacionCompra.TotalDescuento = r("TotalDescuento")

                                If Not r("CodDocReemb") = "" Then
                                    oLiquidacionCompra.CodDocReemb = r("CodDocReemb")
                                    oLiquidacionCompra.TotalComprobantesReembolso = r("TotalComprobantesReembolso")
                                    oLiquidacionCompra.TotalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                                    oLiquidacionCompra.TotalImpuestoReembolso = r("TotalImpuestoReembolso")
                                End If




                                oLiquidacionCompra.ImporteTotal = r("ImporteTotal")
                                oLiquidacionCompra.Moneda = r("Moneda")
                                oLiquidacionCompra.Tipo = r("Tipo")
                                'oLiquidacionCompra.UsuarioCreador = r("UsuarioCreador") ' LOCAL
                                'Try
                                '    oLiquidacionCompra.UsuarioProceso = r("UsuarioCreador") 'NUBE
                                'Catch ex As Exception
                                'End Try

                                oLiquidacionCompra.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                                oLiquidacionCompra.EmailResponsable = r("EmailResponsable")
                                oLiquidacionCompra.SecuencialERP = r("SecuencialERP")
                                oLiquidacionCompra.CodigoTransaccionERP = r("CodigoTransaccionERP")

                                'oLiquidacionCompra.FechaCarga = r("FechaCarga")
                                oLiquidacionCompra.Campo1 = r("Campo1")
                                oLiquidacionCompra.Campo2 = r("Campo2")
                                oLiquidacionCompra.Campo3 = r("Campo3")
                                oLiquidacionCompra.Campo4 = r("Campo4")
                                oLiquidacionCompra.Campo5 = r("Campo5")
                                oLiquidacionCompra.Campo6 = r("Campo6")
                                oLiquidacionCompra.Campo7 = r("Campo7")
                                oLiquidacionCompra.Campo8 = r("Campo8")
                                oLiquidacionCompra.Campo9 = r("Campo9")
                                oLiquidacionCompra.Campo10 = r("Campo10")

                                'ADD DM 08012025
                                oLiquidacionCompra.Campo11 = r("Campo11")
                                oLiquidacionCompra.Campo12 = r("Campo12")
                                oLiquidacionCompra.Campo13 = r("Campo13")
                                oLiquidacionCompra.Campo14 = r("Campo14")
                                oLiquidacionCompra.Campo15 = r("Campo15")

                                'IMPUESTO FACTURA
                                'Impuestos totalizados en la factura.
                                Dim lstimpLQ As Object
                                lstimpLQ = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impLQIVA As Object
                                    impLQIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQIVA.Codigo = "2"
                                    'impLQIVA.CodigoPorcentaje = "2"
                                    'impLQIVA.Tarifa = "12"
                                    impLQIVA.Codigo = r("Codigo8")
                                    impLQIVA.CodigoPorcentaje = r("CodigoPorcentaje8")
                                    impLQIVA.Tarifa = r("Tarifa8")
                                    impLQIVA.BaseImponible = r("Base8")
                                    impLQIVA.Valor = r("ValorIva8")

                                    If r("DescuentoAdicional8") <> "0" Then
                                        impLQIVA.DescuentoAdicional = r("DescuentoAdicional8")
                                        'aplicadoDescuentoAdicional = False
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impLQIVA As Object
                                    impLQIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQIVA.Codigo = "2"
                                    'impLQIVA.CodigoPorcentaje = "2"
                                    'impLQIVA.Tarifa = "12"
                                    impLQIVA.Codigo = r("Codigo12")
                                    impLQIVA.CodigoPorcentaje = r("CodigoPorcentaje12")
                                    impLQIVA.Tarifa = r("Tarifa12")
                                    impLQIVA.BaseImponible = r("Base12")
                                    impLQIVA.Valor = r("ValorIva12")

                                    If r("DescuentoAdicional12") <> "0" Then
                                        impLQIVA.DescuentoAdicional = r("DescuentoAdicional12")
                                        'aplicadoDescuentoAdicional = False
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                'If r("Base13") <> 0 Then
                                '    Dim impLQIVA As Object
                                '    impLQIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto

                                '    'impLQIVA.Codigo = "2"
                                '    'impLQIVA.CodigoPorcentaje = "3"
                                '    'impLQIVA.Tarifa = "14"
                                '    impLQIVA.Codigo = r("Codigo13")
                                '    impLQIVA.CodigoPorcentaje = r("CodigoPorcentaje13")
                                '    impLQIVA.Tarifa = r("Tarifa13")
                                '    impLQIVA.BaseImponible = r("Base13")
                                '    'impLQIVA.Valor = r("ImpuestoTotal")
                                '    impLQIVA.Valor = r("ValorIva13")
                                '    'If aplicadoDescuentoAdicional = False Then
                                '    If r("DescuentoAdicional13") <> "0" Then
                                '        impLQIVA.DescuentoAdicional = r("DescuentoAdicional13")
                                '        'aplicadoDescuentoAdicional = True
                                '    End If
                                '    'End If


                                '    lstimpLQ.Add(impLQIVA)
                                'End If

                                If r("Base0") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "0"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo0")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje0")
                                    impLQNOIVA.Tarifa = r("Tarifa0")
                                    impLQNOIVA.BaseImponible = r("Base0")
                                    impLQNOIVA.Valor = r("ValorIva0")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional0") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional0")
                                        'aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("BaseNoi") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "6"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("CodigoNoi")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impLQNOIVA.Tarifa = r("TarifaNoi")
                                    impLQNOIVA.BaseImponible = r("BaseNoi")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIvaNoi")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicionalNoi") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicionalNoi")
                                        'aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("CodigoExen")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentajeExen")
                                    impLQNOIVA.Tarifa = r("TarifaExen")
                                    impLQNOIVA.BaseImponible = r("BaseExen")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIvaExen")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicionalExen") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicionalExen")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base5") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo5")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje5")
                                    impLQNOIVA.Tarifa = r("Tarifa5")
                                    impLQNOIVA.BaseImponible = r("Base5")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva5")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional5") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional5")
                                        'aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base15") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo15")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje15")
                                    impLQNOIVA.Tarifa = r("Tarifa15")
                                    impLQNOIVA.BaseImponible = r("Base15")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva15")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional15") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional15")
                                        'aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base14") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo14")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje14")
                                    impLQNOIVA.Tarifa = r("Tarifa14")
                                    impLQNOIVA.BaseImponible = r("Base14")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva14")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional14") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional14")
                                        'aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                oLiquidacionCompra.ENTLiquidacionCompraImpuesto = lstimpLQ.ToArray
                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleLiquidacion As Object
                                itemDetalleLiquidacion = New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompra

                                itemDetalleLiquidacion.CodigoPrincipal = r("CodigoPrincipal")
                                itemDetalleLiquidacion.CodigoAuxiliar = r("CodigoAuxiliar")
                                itemDetalleLiquidacion.Descripcion = r("Descripcion")
                                Try
                                    If Not r("UnidadMedida") = "" Then
                                        itemDetalleLiquidacion.UnidadMedida = r("UnidadMedida")
                                    End If
                                Catch ex As Exception
                                End Try
                                itemDetalleLiquidacion.Cantidad = r("Cantidad")
                                itemDetalleLiquidacion.PrecioUnitario = r("PrecioUnitario")
                                itemDetalleLiquidacion.Descuento = r("Descuento")
                                itemDetalleLiquidacion.PrecioTotalSinImpuesto = r("PrecioTotalSinImpuesto")

                                ''Datos adicionales de cada detalle del item                                     
                                Dim listaDetalleDatoAdicional As Object
                                listaDetalleDatoAdicional = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra)
                                'Adicional1
                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As Object
                                    itemDetalleDatoAdicional = New Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra
                                    itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                    itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                                End If

                                'Adicional2
                                If Not r("ConceptoAdicional2") = "0" Then
                                    Dim itemDetalleDatoAdicional2 As Object

                                    itemDetalleDatoAdicional2 = New Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra
                                    itemDetalleDatoAdicional2.Nombre = r("ConceptoAdicional2")
                                    itemDetalleDatoAdicional2.Descripcion = r("NombreAdicional2")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional2)
                                End If

                                'Adicional3
                                If Not r("ConceptoAdicional3") = "0" Then
                                    Dim itemDetalleDatoAdicional3 As Object
                                    itemDetalleDatoAdicional3 = New Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalDetalleLiquidacionCompra
                                    itemDetalleDatoAdicional3.Nombre = r("ConceptoAdicional3")
                                    itemDetalleDatoAdicional3.Descripcion = r("NombreAdicional3")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional3)
                                End If

                                itemDetalleLiquidacion.ENTDatoAdicionalDetalleLiquidacionCompra = listaDetalleDatoAdicional.ToArray

                                'IMPUESTOS DEL DETALLE
                                'Puede Tener IVA y/0 ICE
                                Dim lstimpdetalle As Object
                                'Detalle de impuesto de IVA
                                Dim impdetalleIVA As Object

                                lstimpdetalle = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto)
                                impdetalleIVA = New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto

                                'impdetalleIVA.Codigo = "2" '2 de que tabla debo verlo tabla 15 SRI
                                impdetalleIVA.Codigo = r("Codigo")
                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                    'impdetalleIVA.CodigoPorcentaje = 0  '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA8" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 2 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 12
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 2 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 12
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA13" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 3 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 14
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 6 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 7 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")

                                ElseIf r("TaxCodeAp") = "IVA5" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 7 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")

                                ElseIf r("TaxCodeAp") = "IVA15" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 7 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")

                                ElseIf r("TaxCodeAp") = "IVA14" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 7 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")

                                End If

                                'impdetalleIVA.BaseImponible = r("PrecioTotalSinImpuesto") se comento 22/02/2022 porque ya se envia en cada indicador de impuesto
                                impdetalleIVA.Valor = r("TotalIva")

                                'agrego impuesto a la lista
                                lstimpdetalle.Add(impdetalleIVA)

                                'agrego lista de impuesto al detalle
                                itemDetalleLiquidacion.ENTDetalleLiquidacionCompraImpuesto = lstimpdetalle.ToArray

                                'agrego detalle a la lista
                                listaDetalleLQ.Add(itemDetalleLiquidacion)
                            Next
                            oLiquidacionCompra.ENTDetalleLiquidacionCompra = listaDetalleLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows


                                'Dim reembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolso
                                Dim reembolsoLQ As Object
                                reembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolso
                                If Not r("TipoIdentificacionProveedorReembolso") = "" Then
                                    reembolsoLQ.TipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                                End If
                                If Not r("IdentificacionProveedorReembolso") = "" Then
                                    reembolsoLQ.IdentificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                                End If
                                If Not r("CodPaisPagoProveedorReembolso") = "" Then
                                    reembolsoLQ.CodPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                                End If
                                If Not r("TipoProveedorReembolso") = "" Then
                                    reembolsoLQ.TipoProveedorReembolso = r("TipoProveedorReembolso")
                                End If
                                If Not r("CodDocReembolso") = "" Then
                                    reembolsoLQ.CodDocReembolso = r("CodDocReembolso")
                                End If
                                If Not r("EstabDocReembolso") = "" Then
                                    reembolsoLQ.EstabDocReembolso = r("EstabDocReembolso")
                                End If
                                If Not r("PtoEmiDocReembolso") = "" Then
                                    reembolsoLQ.PtoEmiDocReembolso = r("PtoEmiDocReembolso")
                                End If
                                If Not r("SecuencialDocReembolso") = "" Then
                                    reembolsoLQ.SecuencialDocReembolso = r("SecuencialDocReembolso")
                                End If
                                If Not r("FechaEmisionDocReembolso") = "" Then
                                    reembolsoLQ.FechaEmisionDocReembolso = r("FechaEmisionDocReembolso")
                                End If
                                If Not r("NumeroAutorizacionDocReem") = "" Then
                                    reembolsoLQ.NumeroAutorizacionDocReemb = r("NumeroAutorizacionDocReem")
                                End If

                                'Dim listaImpReembolsoLQ As New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto)
                                'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                Dim listaImpReembolsoLQ As Object
                                listaImpReembolsoLQ = New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto)

                                If r("Base8") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "2"
                                    'itemImpReembolsoLQ.Tarifa = "12"
                                    itemImpReembolsoLQ.Codigo = r("Codigo8")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje8")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa8")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base8")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem8")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base12") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "2"
                                    'itemImpReembolsoLQ.Tarifa = "12"
                                    itemImpReembolsoLQ.Codigo = r("Codigo12")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje12")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa12")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base12")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem12")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("Base13") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "2"
                                    'itemImpReembolsoLQ.Tarifa = "14"
                                    itemImpReembolsoLQ.Codigo = r("Codigo13")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje13")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa13")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base13")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem13")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("Base0") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "0"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo0")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje0")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa0")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base0")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem0")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("BaseNoi") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "6"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("CodigoNoi")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentajeNoi")
                                    itemImpReembolsoLQ.Tarifa = r("TarifaNoi")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("BaseNoi")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReemNoi")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("BaseExen") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("CodigoExen")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentajeExen")
                                    itemImpReembolsoLQ.Tarifa = r("TarifaExen")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("BaseExen")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReemExen")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base5") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo5")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje5")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa5")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base5")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem5")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base15") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo15")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje15")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa15")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base15")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem15")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base14") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo14")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje14")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa14")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base14")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem14")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                reembolsoLQ.ENTLiquidacionCompraReembolsoImpuestos = listaImpReembolsoLQ.ToArray

                                listareembolsoLQ.Add(reembolsoLQ)

                            Next
                            oLiquidacionCompra.ENTLiquidacionCompraReembolso = listareembolsoLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Reembolso: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Reembolso: " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 3 Then
                        Try
                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim itemDatoAdicionalLQ As Object
                                itemDatoAdicionalLQ = New Entidades.wsEDoc_LiquidacionCompra.ENTDatoAdicionalLiquidacionCompra

                                itemDatoAdicionalLQ.Nombre = r("Concepto")
                                itemDatoAdicionalLQ.Descripcion = r("Descripcion")
                                listaDatosAdicionalesLQ.Add(itemDatoAdicionalLQ)
                            Next
                            oLiquidacionCompra.ENTDatoAdicionalLiquidacionCompra = listaDatosAdicionalesLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Datos Adicionales: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Datos Adicionales: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 4 Then
                        Try

                            For Each r As DataRow In ds.Tables(4).Rows
                                Dim Pago As Object
                                Pago = New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraPagos

                                Pago.FormaPago = r("FormaPago")
                                Pago.Total = r("Total")
                                If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                    Pago.Plazo = Nothing
                                Else
                                    Pago.Plazo = r("Plazo")
                                End If
                                If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                    Pago.UnidadTiempo = Nothing
                                Else
                                    Pago.UnidadTiempo = r("UnidadTiempo")
                                End If
                                liquidacionCompraPagos.Add(Pago)
                            Next
                            oLiquidacionCompra.ENTLiquidacionCompraPagos = liquidacionCompraPagos.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    End If

                Next

            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'Dim sRutaCarpeta As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
                'Dim sRutaCarpeta As String = ""
                'If _tipoManejo = "A" Then
                '    sRutaCarpeta = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
                'Else
                '    sRutaCarpeta = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                'End If

                Dim sRuta As String = sRutaCarpeta & oLiquidacionCompra.Secuencial.ToString() + oLiquidacionCompra.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oLiquidacionCompra.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oLiquidacionCompra)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oLiquidacionCompra.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oLiquidacionCompra)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oLiquidacionCompra
            Utilitario.Util_Log.Escribir_Log("Liquidacion consultada", "ManejoDeDocumentos")
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Liquidación de Compra en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "ArgumentException-Error al Consultar Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing

        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la Liquidación de Compra en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Error al Consultar Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        End Try

    End Function

    Public Function ConsultarLiquidacionCompra_LOCAL_NUBE(ByVal TipoRE As String, ByVal DocEntry As Integer, ByVal TipoWS As String) As Object

        Dim oLiquidacionCompra As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompra = Nothing


        Dim listaDetalleLQ As List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompra)
        Dim listaDatoAdicionalDetalleLQCompra As List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalDetalleLiquidacionCompra)
        Dim listaLiquidacionCompraImp As List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompraImpuesto)
        Dim liquidacionCompraPagos As List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraPagos)
        Dim listaDatosAdicionalesLQ As List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalLiquidacionCompra)

        listaDetalleLQ = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompra)
        listaDatoAdicionalDetalleLQCompra = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalDetalleLiquidacionCompra)
        listaLiquidacionCompraImp = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompraImpuesto)
        liquidacionCompraPagos = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraPagos)

        listaDatosAdicionalesLQ = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalLiquidacionCompra)


        Dim listareembolsoLQ As New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolso)
        Dim aplicadoDescuentoAdicional As Boolean = False
        Try

            Dim SP As String = ""
            If TipoRE = "LQE" Then
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    SP = "GS_SAP_FE_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    SP = "GS_SAP_FE_ONE_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    SP = "GS_SAP_FE_HEI_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    SP = "GS_SAP_FE_SYP_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    SP = "GS_SAP_FE_TM_ObtenerLiquidacionCompra_4_3"
                ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    SP = "GS_SAP_FE_SS_ObtenerLiquidacionCompra_4_3"
                End If

                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Consultando Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", SP: " + SP.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
            End If

            Dim ds As DataSet = EjecutarSP(SP, DocEntry)
            Utilitario.Util_Log.Escribir_Log("Data Tables : " & ds.Tables.Count.ToString(), "ManejoDeDocumentos")

            If Not ds Is Nothing And Not ds.Tables.Count = 0 Then
                oLiquidacionCompra = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompra

                For i As Integer = 0 To ds.Tables.Count - 1
                    If i = 0 Then
                        Try
                            For Each r As DataRow In ds.Tables(0).Rows

                                ' MANEJO DE FACTURAS DE EXPORTACION Y REEMBOLSO - 2018-02-18
                                ' Indica que tipo de factura es (0.- Normal, 1.- Exportadores, 2.- Reembolsos)
                                'Try
                                '    If r("TipoLiquidacionCompra").ToString() = "" Then
                                '        oLiquidacionCompra.Tipo = 0
                                '    Else
                                '        'oLiquidacionCompra.Tipo = r("TipoLiquidacionCompra")
                                '        oLiquidacionCompra.Tipo = 1
                                '    End If
                                '    Utilitario.Util_Log.Escribir_Log(" (0.- Normal, 1.- Exportadores, 2.- Reembolsos)", "ManejoDeDocumentos")
                                '    Utilitario.Util_Log.Escribir_Log("Tipo Factura : " & oLiquidacionCompra.Tipo.ToString(), "ManejoDeDocumentos")
                                'Catch ex As Exception
                                '    Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                                '    oLiquidacionCompra.Tipo = 0
                                'End Try

                                ' OFFLINE 14 NOVIEMBRE 2017
                                'FAMC 18/02/2019
                                If ValidaClave(r("ClaveAcceso"), r("CodigoDocumento"), r("Establecimiento") & r("PuntoEmision"), r("SecuencialDocumento")) = "" Then
                                    oLiquidacionCompra.ClaveAcceso = Nothing
                                Else
                                    oLiquidacionCompra.ClaveAcceso = r("ClaveAcceso")
                                End If

                                oLiquidacionCompra.Ambiente = r("Ambiente")
                                oLiquidacionCompra.TipoEmision = r("TipoEmision")
                                oLiquidacionCompra.RazonSocial = r("RazonSocial")
                                If Not r("NombreComercial") = "" Then
                                    oLiquidacionCompra.NombreComercial = r("NombreComercial")
                                End If


                                oLiquidacionCompra.Ruc = r("RUC")
                                'oLiquidacionCompra.Ruc = "0992737964001"
                                oLiquidacionCompra.CodigoDocumento = r("CodigoDocumento")
                                oLiquidacionCompra.Establecimiento = r("Establecimiento")
                                oLiquidacionCompra.PuntoEmision = r("PuntoEmision")
                                oLiquidacionCompra.Secuencial = r("SecuencialDocumento")
                                If Not oLiquidacionCompra.Secuencial.ToString().Length.Equals("9") Then
                                    oLiquidacionCompra.Secuencial = oLiquidacionCompra.Secuencial.ToString().PadLeft(9, "0")
                                End If
                                Utilitario.Util_Log.Escribir_Log("oLiquidacionCompra.Secuencial : " & oLiquidacionCompra.Secuencial.ToString(), "ManejoDeDocumentos")
                                oLiquidacionCompra.DireccionMatriz = r("DireccionMatriz")
                                oLiquidacionCompra.FechaEmision = r("FechaEmision")
                                oLiquidacionCompra.DireccionEstablecimiento = r("DireccionEstablecimiento")

                                If Not r("ContribuyenteEspecial") = "0" Then
                                    oLiquidacionCompra.ContribuyenteEspecial = r("ContribuyenteEspecial")
                                Else
                                    oLiquidacionCompra.ContribuyenteEspecial = Nothing
                                End If

                                If Not r("AgenteRetencion") = "0" Then
                                    oLiquidacionCompra.AgenteRetencion = r("AgenteRetencion")
                                End If

                                If Not r("RegimenMicroempresas") = "0" Then
                                    oLiquidacionCompra.RegimenMicroempresas = Convert.ToBoolean(r("RegimenMicroempresas"))
                                End If

                                If Not r("ContribuyenteRimpe") = "0" Then
                                    oLiquidacionCompra.ContribuyenteRimpe = Convert.ToBoolean(r("ContribuyenteRimpe"))
                                End If

                                oLiquidacionCompra.ObligadoContabilidad = r("ObligadoContabilidad")
                                oLiquidacionCompra.TipoIdentificacionProveedor = r("TipoIdentificacionProveedor")

                                'If Not r("GuiaRemision") = "0" Then
                                '    oLiquidacionCompra.GuiaRemision = r("GuiaRemision")
                                'End If

                                oLiquidacionCompra.RazonSocialProveedor = r("RazonSocialProveedor")
                                oLiquidacionCompra.IdentificacionProveedor = r("IdentificacionProveedor")

                                Try
                                    If Not r("DirProveedor") = "" Then
                                        oLiquidacionCompra.DirProveedor = r("DirProveedor")
                                    End If
                                Catch ex As Exception
                                End Try

                                oLiquidacionCompra.TotalSinImpuesto = r("TotalSinImpuesto")
                                oLiquidacionCompra.TotalDescuento = r("TotalDescuento")

                                If Not r("CodDocReemb") = "" Then
                                    oLiquidacionCompra.CodDocReemb = r("CodDocReemb")
                                    oLiquidacionCompra.TotalComprobantesReembolso = r("TotalComprobantesReembolso")
                                    oLiquidacionCompra.TotalBaseImponibleReembolso = r("TotalBaseImponibleReembolso")
                                    oLiquidacionCompra.TotalImpuestoReembolso = r("TotalImpuestoReembolso")
                                End If




                                oLiquidacionCompra.ImporteTotal = r("ImporteTotal")
                                oLiquidacionCompra.Moneda = r("Moneda")
                                oLiquidacionCompra.Tipo = r("Tipo")
                                'oLiquidacionCompra.UsuarioCreador = r("UsuarioCreador") ' LOCAL
                                'Try
                                '    oLiquidacionCompra.UsuarioProceso = r("UsuarioCreador") 'NUBE
                                'Catch ex As Exception
                                'End Try

                                oLiquidacionCompra.UsuarioTransaccionERP = r("UsuarioTransaccionERP")
                                oLiquidacionCompra.EmailResponsable = r("EmailResponsable")
                                oLiquidacionCompra.SecuencialERP = r("SecuencialERP")
                                oLiquidacionCompra.CodigoTransaccionERP = r("CodigoTransaccionERP")

                                'oLiquidacionCompra.FechaCarga = r("FechaCarga")
                                oLiquidacionCompra.Campo1 = r("Campo1")
                                oLiquidacionCompra.Campo2 = r("Campo2")
                                oLiquidacionCompra.Campo3 = r("Campo3")
                                oLiquidacionCompra.Campo4 = r("Campo4")
                                oLiquidacionCompra.Campo5 = r("Campo5")
                                oLiquidacionCompra.Campo6 = r("Campo6")
                                oLiquidacionCompra.Campo7 = r("Campo7")
                                oLiquidacionCompra.Campo8 = r("Campo8")
                                oLiquidacionCompra.Campo9 = r("Campo9")
                                oLiquidacionCompra.Campo10 = r("Campo10")
                                'IMPUESTO FACTURA
                                'Impuestos totalizados en la factura.
                                Dim lstimpLQ As Object
                                lstimpLQ = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto)

                                If r("Base8") <> 0 Then
                                    Dim impLQIVA As Object
                                    impLQIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQIVA.Codigo = "2"
                                    'impLQIVA.CodigoPorcentaje = "2"
                                    'impLQIVA.Tarifa = "12"
                                    impLQIVA.Codigo = r("Codigo8")
                                    impLQIVA.CodigoPorcentaje = r("CodigoPorcentaje8")
                                    impLQIVA.Tarifa = r("Tarifa8")
                                    impLQIVA.BaseImponible = r("Base8")
                                    impLQIVA.Valor = r("ValorIva8")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQIVA.DescuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base12") <> 0 Then
                                    Dim impLQIVA As Object
                                    impLQIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQIVA.Codigo = "2"
                                    'impLQIVA.CodigoPorcentaje = "2"
                                    'impLQIVA.Tarifa = "12"
                                    impLQIVA.Codigo = r("Codigo12")
                                    impLQIVA.CodigoPorcentaje = r("CodigoPorcentaje12")
                                    impLQIVA.Tarifa = r("Tarifa12")
                                    impLQIVA.BaseImponible = r("Base12")
                                    impLQIVA.Valor = r("ValorIva12")

                                    If r("DescuentoAdicional") <> "0" Then
                                        impLQIVA.DescuentoAdicional = r("DescuentoAdicional")
                                        aplicadoDescuentoAdicional = True
                                    End If

                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base13") <> 0 Then
                                    Dim impLQIVA As Object
                                    impLQIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompraImpuesto

                                    'impLQIVA.Codigo = "2"
                                    'impLQIVA.CodigoPorcentaje = "3"
                                    'impLQIVA.Tarifa = "14"
                                    impLQIVA.Codigo = r("Codigo13")
                                    impLQIVA.CodigoPorcentaje = r("CodigoPorcentaje13")
                                    impLQIVA.Tarifa = r("Tarifa13")
                                    impLQIVA.BaseImponible = r("Base13")
                                    'impLQIVA.Valor = r("ImpuestoTotal")
                                    impLQIVA.Valor = r("ValorIva13")
                                    If aplicadoDescuentoAdicional = False Then
                                        If r("DescuentoAdicional") <> "0" Then
                                            impLQIVA.DescuentoAdicional = r("DescuentoAdicional")
                                            aplicadoDescuentoAdicional = True
                                        End If
                                    End If


                                    lstimpLQ.Add(impLQIVA)
                                End If

                                If r("Base0") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "0"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo0")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje0")
                                    impLQNOIVA.Tarifa = r("Tarifa0")
                                    impLQNOIVA.BaseImponible = r("Base0")
                                    impLQNOIVA.Valor = r("ValorIva0")
                                    If aplicadoDescuentoAdicional = False Then
                                        If r("DescuentoAdicional") <> "0" Then
                                            impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional")
                                            aplicadoDescuentoAdicional = True
                                        End If
                                    End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("BaseNoi") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "6"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("CodigoNoi")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentajeNoi")
                                    impLQNOIVA.Tarifa = r("TarifaNoi")
                                    impLQNOIVA.BaseImponible = r("BaseNoi")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIvaNoi")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicionalNoi") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicionalNoi")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("BaseExen") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("CodigoExen")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentajeExen")
                                    impLQNOIVA.Tarifa = r("TarifaExen")
                                    impLQNOIVA.BaseImponible = r("BaseExen")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIvaExen")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicionalExen") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicionalExen")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    ' End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base5") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo5")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje5")
                                    impLQNOIVA.Tarifa = r("Tarifa5")
                                    impLQNOIVA.BaseImponible = r("Base5")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva5")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional5") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional5")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base15") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo15")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje15")
                                    impLQNOIVA.Tarifa = r("Tarifa15")
                                    impLQNOIVA.BaseImponible = r("Base15")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva15")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional15") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional15")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base14") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo14")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje14")
                                    impLQNOIVA.Tarifa = r("Tarifa14")
                                    impLQNOIVA.BaseImponible = r("Base14")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva14")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional14") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional14")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                If r("Base13") <> 0 Then

                                    Dim impLQNOIVA As Object
                                    impLQNOIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraImpuesto

                                    'impLQNOIVA.Codigo = "2"
                                    'impLQNOIVA.CodigoPorcentaje = "7"
                                    'impLQNOIVA.Tarifa = "0"
                                    impLQNOIVA.Codigo = r("Codigo13")
                                    impLQNOIVA.CodigoPorcentaje = r("CodigoPorcentaje13")
                                    impLQNOIVA.Tarifa = r("Tarifa13")
                                    impLQNOIVA.BaseImponible = r("Base13")
                                    'impLQNOIVA.Valor = 0
                                    impLQNOIVA.Valor = r("ValorIva13")
                                    'If aplicadoDescuentoAdicional = False Then
                                    If r("DescuentoAdicional13") <> "0" Then
                                        impLQNOIVA.DescuentoAdicional = r("DescuentoAdicional13")
                                        aplicadoDescuentoAdicional = True
                                    End If
                                    'End If

                                    lstimpLQ.Add(impLQNOIVA)
                                End If

                                oLiquidacionCompra.ENTLiquidacionCompraImpuesto = lstimpLQ.ToArray
                            Next
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Cabecera " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Cabecera: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 1 Then
                        Try
                            For Each r As DataRow In ds.Tables(1).Rows

                                Dim itemDetalleLiquidacion As Object
                                itemDetalleLiquidacion = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompra

                                itemDetalleLiquidacion.CodigoPrincipal = r("CodigoPrincipal")
                                itemDetalleLiquidacion.CodigoAuxiliar = r("CodigoAuxiliar")
                                itemDetalleLiquidacion.Descripcion = r("Descripcion")
                                Try
                                    If Not r("UnidadMedida") = "" Then
                                        itemDetalleLiquidacion.UnidadMedida = r("UnidadMedida")
                                    End If
                                Catch ex As Exception
                                End Try
                                itemDetalleLiquidacion.Cantidad = r("Cantidad")
                                itemDetalleLiquidacion.PrecioUnitario = r("PrecioUnitario")
                                itemDetalleLiquidacion.Descuento = r("Descuento")
                                itemDetalleLiquidacion.PrecioTotalSinImpuesto = r("PrecioTotalSinImpuesto")

                                ''Datos adicionales de cada detalle del item                                     
                                Dim listaDetalleDatoAdicional As Object
                                listaDetalleDatoAdicional = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalDetalleLiquidacionCompra)
                                'Adicional1
                                If Not r("ConceptoAdicional1") = "0" Then
                                    Dim itemDetalleDatoAdicional As Object
                                    itemDetalleDatoAdicional = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalDetalleLiquidacionCompra
                                    itemDetalleDatoAdicional.Nombre = r("ConceptoAdicional1")
                                    itemDetalleDatoAdicional.Descripcion = r("NombreAdicional1")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional)
                                End If

                                'Adicional2
                                If Not r("ConceptoAdicional2") = "0" Then
                                    Dim itemDetalleDatoAdicional2 As Object

                                    itemDetalleDatoAdicional2 = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalDetalleLiquidacionCompra
                                    itemDetalleDatoAdicional2.Nombre = r("ConceptoAdicional2")
                                    itemDetalleDatoAdicional2.Descripcion = r("NombreAdicional2")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional2)
                                End If

                                'Adicional3
                                If Not r("ConceptoAdicional3") = "0" Then
                                    Dim itemDetalleDatoAdicional3 As Object
                                    itemDetalleDatoAdicional3 = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalDetalleLiquidacionCompra
                                    itemDetalleDatoAdicional3.Nombre = r("ConceptoAdicional3")
                                    itemDetalleDatoAdicional3.Descripcion = r("NombreAdicional3")
                                    listaDetalleDatoAdicional.Add(itemDetalleDatoAdicional3)
                                End If

                                itemDetalleLiquidacion.ENTDatoAdicionalDetalleLiquidacionCompra = listaDetalleDatoAdicional.ToArray

                                'IMPUESTOS DEL DETALLE
                                'Puede Tener IVA y/0 ICE
                                Dim lstimpdetalle As Object
                                'Detalle de impuesto de IVA
                                Dim impdetalleIVA As Object

                                lstimpdetalle = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompraImpuesto)
                                impdetalleIVA = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDetalleLiquidacionCompraImpuesto

                                'impdetalleIVA.Codigo = "2" '2 de que tabla debo verlo tabla 15 SRI
                                impdetalleIVA.Codigo = r("Codigo")
                                If r("TaxCodeAp") = "IVA_EXE" Then ' 0%
                                    'impdetalleIVA.CodigoPorcentaje = 0  '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA8" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 2 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 12
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 2 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 12
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA13" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 3 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 14
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_NOI" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 6 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_EXEN" Then ' 12%
                                    'impdetalleIVA.CodigoPorcentaje = 7 '2 de que tabla debo verlo tabla 16
                                    'impdetalleIVA.Tarifa = 0
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_5" Then
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_15" Then
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")
                                ElseIf r("TaxCodeAp") = "IVA_14" Then
                                    impdetalleIVA.CodigoPorcentaje = r("CodigoPorcentaje")
                                    impdetalleIVA.Tarifa = r("Tarifa")
                                    impdetalleIVA.BaseImponible = r("BaseImponible")

                                End If


                                'impdetalleIVA.BaseImponible = r("PrecioTotalSinImpuesto") se comento 22/02/2022 porque ya se envia en cada indicador de impuesto
                                impdetalleIVA.Valor = r("TotalIva")

                                'agrego impuesto a la lista
                                lstimpdetalle.Add(impdetalleIVA)

                                'agrego lista de impuesto al detalle
                                itemDetalleLiquidacion.ENTDetalleLiquidacionCompraImpuesto = lstimpdetalle.ToArray

                                'agrego detalle a la lista
                                listaDetalleLQ.Add(itemDetalleLiquidacion)
                            Next
                            oLiquidacionCompra.ENTDetalleLiquidacionCompra = listaDetalleLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("DETALLE: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "DETALLE: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 2 Then
                        Try
                            For Each r As DataRow In ds.Tables(2).Rows


                                'Dim reembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolso
                                Dim reembolsoLQ As Object
                                reembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolso
                                If Not r("TipoIdentificacionProveedorReembolso") = "" Then
                                    reembolsoLQ.TipoIdentificacionProveedorReembolso = r("TipoIdentificacionProveedorReembolso")
                                End If
                                If Not r("IdentificacionProveedorReembolso") = "" Then
                                    reembolsoLQ.IdentificacionProveedorReembolso = r("IdentificacionProveedorReembolso")
                                End If
                                If Not r("CodPaisPagoProveedorReembolso") = "" Then
                                    reembolsoLQ.CodPaisPagoProveedorReembolso = r("CodPaisPagoProveedorReembolso")
                                End If
                                If Not r("TipoProveedorReembolso") = "" Then
                                    reembolsoLQ.TipoProveedorReembolso = r("TipoProveedorReembolso")
                                End If
                                If Not r("CodDocReembolso") = "" Then
                                    reembolsoLQ.CodDocReembolso = r("CodDocReembolso")
                                End If
                                If Not r("EstabDocReembolso") = "" Then
                                    reembolsoLQ.EstabDocReembolso = r("EstabDocReembolso")
                                End If
                                If Not r("PtoEmiDocReembolso") = "" Then
                                    reembolsoLQ.PtoEmiDocReembolso = r("PtoEmiDocReembolso")
                                End If
                                If Not r("SecuencialDocReembolso") = "" Then
                                    reembolsoLQ.SecuencialDocReembolso = r("SecuencialDocReembolso")
                                End If
                                If Not r("FechaEmisionDocReembolso") = "" Then
                                    reembolsoLQ.FechaEmisionDocReembolso = r("FechaEmisionDocReembolso")
                                End If
                                If Not r("NumeroAutorizacionDocReem") = "" Then
                                    reembolsoLQ.NumeroAutorizacionDocReemb = r("NumeroAutorizacionDocReem")
                                End If

                                'Dim listaImpReembolsoLQ As New List(Of Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto)
                                'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                Dim listaImpReembolsoLQ As Object
                                listaImpReembolsoLQ = New List(Of Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto)

                                If r("Base8") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "2"
                                    'itemImpReembolsoLQ.Tarifa = "12"
                                    itemImpReembolsoLQ.Codigo = r("Codigo8")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje8")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa8")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base8")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem8")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base12") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "2"
                                    'itemImpReembolsoLQ.Tarifa = "12"
                                    itemImpReembolsoLQ.Codigo = r("Codigo12")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje12")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa12")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base12")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem12")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("Base13") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "2"
                                    'itemImpReembolsoLQ.Tarifa = "14"
                                    itemImpReembolsoLQ.Codigo = r("Codigo13")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje13")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa13")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base13")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem13")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("Base0") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "0"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo0")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje0")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa0")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base0")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem0")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("BaseNoi") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "6"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("CodigoNoi")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentajeNoi")
                                    itemImpReembolsoLQ.Tarifa = r("TarifaNoi")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("BaseNoi")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReemNoi")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If
                                If r("BaseExen") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("CodigoExen")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentajeExen")
                                    itemImpReembolsoLQ.Tarifa = r("TarifaExen")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("BaseExen")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReemExen")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base5") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo5")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje5")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa5")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base5")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem5")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If


                                If r("Base15") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo15")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje15")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa15")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base15")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem15")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                If r("Base14") <> 0 Then
                                    'Dim impLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTDetalleLiquidacionCompraImpuesto
                                    'Dim itemImpReembolsoLQ As New Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompraReembolsoImpuesto
                                    Dim itemImpReembolsoLQ As Object
                                    itemImpReembolsoLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraReembolsoImpuesto
                                    'itemImpReembolsoLQ.Codigo = "2"
                                    'itemImpReembolsoLQ.CodigoPorcentaje = "7"
                                    'itemImpReembolsoLQ.Tarifa = "0"
                                    itemImpReembolsoLQ.Codigo = r("Codigo14")
                                    itemImpReembolsoLQ.CodigoPorcentaje = r("CodigoPorcentaje14")
                                    itemImpReembolsoLQ.Tarifa = r("Tarifa14")
                                    itemImpReembolsoLQ.BaseImponibleReembolso = r("Base14")
                                    itemImpReembolsoLQ.ImpuestoReembolso = r("ValorIvaReem14")

                                    listaImpReembolsoLQ.Add(itemImpReembolsoLQ)
                                End If

                                reembolsoLQ.ENTLiquidacionCompraReembolsoImpuestos = listaImpReembolsoLQ.ToArray

                                listareembolsoLQ.Add(reembolsoLQ)

                            Next
                            oLiquidacionCompra.ENTLiquidacionCompraReembolso = listareembolsoLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Reembolso: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Reembolso: " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    ElseIf i = 3 Then
                        Try
                            For Each r As DataRow In ds.Tables(3).Rows
                                Dim itemDatoAdicionalLQ As Object
                                itemDatoAdicionalLQ = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTDatoAdicionalLiquidacionCompra

                                itemDatoAdicionalLQ.Nombre = r("Concepto")
                                itemDatoAdicionalLQ.Descripcion = r("Descripcion")
                                listaDatosAdicionalesLQ.Add(itemDatoAdicionalLQ)
                            Next
                            oLiquidacionCompra.ENTDatoAdicionalLiquidacionCompra = listaDatosAdicionalesLQ.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Datos Adicionales: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Datos Adicionales: " + ex.Message.ToString()
                            Return Nothing
                        End Try
                    ElseIf i = 4 Then
                        Try

                            For Each r As DataRow In ds.Tables(4).Rows
                                Dim Pago As Object
                                Pago = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompraPagos

                                Pago.FormaPago = r("FormaPago")
                                Pago.Total = r("Total")
                                If IsNothing(r("Plazo").ToString()) Or r("Plazo").ToString() = "0" Then
                                    Pago.Plazo = Nothing
                                Else
                                    Pago.Plazo = r("Plazo")
                                End If
                                If IsNothing(r("UnidadTiempo").ToString()) Or r("UnidadTiempo").ToString() = "" Then
                                    Pago.UnidadTiempo = Nothing
                                Else
                                    Pago.UnidadTiempo = r("UnidadTiempo")
                                End If
                                liquidacionCompraPagos.Add(Pago)
                            Next
                            oLiquidacionCompra.ENTLiquidacionCompraPagos = liquidacionCompraPagos.ToArray
                        Catch ex As Exception
                            If _tipoManejo = "A" Then
                                rsboApp.SetStatusBarMessage("Forma de Pago : " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            _Error = "Forma de Pago : " + ex.Message.ToString()
                            Return Nothing
                        End Try

                    End If

                Next

            End If

            Try
                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                Dim sRuta As String = sRutaCarpeta & oLiquidacionCompra.Secuencial.ToString() + oLiquidacionCompra.SecuencialERP.ToString() + ".xml"
                If System.IO.Directory.Exists(sRutaCarpeta) Then
                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                    If TipoWS = "NUBE_4_1" Then
                        Dim x As XmlSerializer = New XmlSerializer(oLiquidacionCompra.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oLiquidacionCompra)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    Else
                        Dim x As XmlSerializer = New XmlSerializer(oLiquidacionCompra.GetType)
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        x.Serialize(writer, oLiquidacionCompra)
                        writer.Close()
                        Utilitario.Util_Log.Escribir_Log("Serializado..." + sRuta, "ManejoDeDocumentos")
                    End If

                End If

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
            End Try

            Return oLiquidacionCompra
            Utilitario.Util_Log.Escribir_Log("Liquidacion consultada", "ManejoDeDocumentos")
        Catch x As ArgumentException
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("ArgumentException-Ocurrio un error al consultar datos de la Liquidación de Compra en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & x.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "ArgumentException-Error al Consultar Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", Descr: " + x.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing

        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Ocurrio un error al consultar datos de la Liquidación de Compra en la Base, DocEntry :  " & DocEntry.ToString() & " Descr: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(TipoRE, DocEntry, "Error al Consultar Liquidación de Compra con # DocEntry = " + DocEntry.ToString() + ", Descr: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        End Try

    End Function

#End Region

#Region "Envío de Documentos"
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
                'Se valida el parametro sincronizado 
                'en caso de ser verdadero solo se llamara al metodo sincronizar
                'caso contrario se Enviara el documento
                'llamar metodo sincronizador
                objetoRespuesta = SincronizarDocumentoEDOC(TipoDocumento, DocEntry, TipoWS)
                'objetoRespuesta = EnviaDocumentoSRI(oObjeto, TipoDocumento, DocEntry, TipoWS)

                If Not objetoRespuesta Is Nothing Then

                    'lbestadoSRI.Text = resp.Estado
                    _EstadoAutorizacion = objetoRespuesta.EstadoEDOC
                    _ClaveAcceso = objetoRespuesta.ClaveAcceso

                    If _tipoManejo = "A" Then

                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "GS_SINCRO_Respuesta del SRI: " + _EstadoAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    End If

                    Dim mensajeError As String = ""
                    If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                        Try
                            _NumAutorizacion = objetoRespuesta.AutorizacionSRI.ToString()
                            _Observacion = "GS_SINCRO: Documento AUTORIZADO AUTORIZACION # " & _NumAutorizacion
                            _FechaAutorizacion = objetoRespuesta.fechaAutorizacion
                        Catch ex As Exception

                        End Try
                    Else
                        _NumAutorizacion = "0000000000"
                        _Observacion = "GS_SINCRO: " & objetoRespuesta.ErrorEDOC
                        mensajeError = _Observacion.ToString
                    End If



                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("GS_SINCRO_Grabando respuesta de SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "GS_SINCRO_Grabando Respuesta del SRI en Documento - " + TipoDocumento + " - " + DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                    End If

                    _Observacion = String.Format("SINCRO Estado:{0} - # AUTORIZACION {1} - Mensaje - {2}", _EstadoAutorizacion.ToString, _NumAutorizacion.ToString, mensajeError)

                    ' Mando a Grabar a SAP
                    If TipoDocumento = "LQE" Then
                        result = GrabaDatosAutorizacion_LiquidacionCompra(DocEntry, TipoDocumento)

                    ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                        result = GrabaDatosAutorizacion_HESION_FACTURAGUIA(DocEntry, TipoDocumento)

                    ElseIf TipoDocumento = "SSGR" Then

                        result = GrabaDatosAutGuiasDesatendidas(DocEntry, TipoDocumento)
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

                    'AQUI FUNCION IMPRIMIR DOCUMENTO
                    If _tipoManejo = "A" Then
                        If Functions.VariablesGlobales._ImpDocAut = "Y" Then
                            If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                                rsboApp.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + " - Imprimiendo Documento por favor esperar... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                mensajeDocAut = ""
                                If ImprmirDOcAut(_NumAutorizacion) Then
                                    rsboApp.SetStatusBarMessage("El documento se imprimió con éxito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                Else
                                    rsboApp.SetStatusBarMessage("Error al imprimir PDF: " + mensajeDocAut.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End If
                            End If
                        End If
                    End If
                Else
                    'No se pudo sincronizar
                    _Observacion = "GS_SINCRO - ProcesaEnvioDocumento - ObjetoRespuesta vacio  " + DocEntry.ToString() + " - " + mensaje.ToString()
                    _Error = "GS_SINCRO-No se recibio respuesta inmediata del servicio : " + _errorMensajeWSEnvío.ToString()
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("GS_SINCRO-No se recibio respuesta, Presione nuevamente el boton de Consultar Autorizacion.", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "GS_SINCRO-No se recibió respuesta de los Web Services", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If

                    Try
                        If TipoDocumento = "LQE" Then
                            GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry, TipoDocumento, _Error)
                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                            GrabaDatosAutorizacion_Error_FacturaGuiaRemision(DocEntry, _Error)
                        ElseIf TipoDocumento = "SSGR" Then
                            GrabaDatosAutGuiasDesatendidas_Error(DocEntry, TipoDocumento, _Error)
                        Else
                            GrabaDatosAutorizacion_Error(DocEntry, TipoDocumento, _Error)
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Else

                If _tipoManejo = "A" Then rsboApp.SetStatusBarMessage("Seteando informacion a enviar..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" Then
                    oObjeto = ConsultarFactura(TipoDocumento, DocEntry)
                    'ElseIf TipoDocumento = "NDE" Then
                    '    oObjeto = ConsultarNotadeDebito(TipoDocumento, DocEntry, TipoWS)
                ElseIf TipoDocumento = "NCE" Then
                    oObjeto = ConsultarNotadeCredito(TipoDocumento, DocEntry)
                    'ElseIf TipoDocumento = "GRE" Or TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then 'AGREGADO TLE solicitud de traslado
                    '    If Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                    '        oObjeto = Consultar_Factura_GuiaDeRemision(TipoDocumento, DocEntry, TipoWS)
                    '    ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then
                    '        oObjeto = Consultar_SalidaMercancias_GuiaDeRemision(TipoDocumento, DocEntry, TipoWS)
                    '    Else
                    '        oObjeto = ConsultarGuiaDeRemision(TipoDocumento, DocEntry, TipoWS)
                    '    End If
                ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
                    oObjeto = ConsultarRetencion(TipoDocumento, DocEntry)
                    'ElseIf TipoDocumento = "LQE" Then
                    '    oObjeto = ConsultarLiquidacion(TipoDocumento, DocEntry, TipoWS)
                    'ElseIf TipoDocumento = "RDM" Then
                    '    oObjeto = ConsultarRetencionND(TipoDocumento, DocEntry, TipoWS)
                    'ElseIf TipoDocumento = "SSGR" Then
                    '    oObjeto = ConsultarGuiaDesatendida_NUBE_4_1(TipoDocumento, DocEntry, TipoWS)
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
                    'Se envia  a procesar el documento
                    objetoRespuesta = EnviaDocumentoSRI(oObjeto, TipoDocumento, DocEntry, TipoWS, respuesta_WS)


                    If Not objetoRespuesta Is Nothing Then
                        ' oBackgroundWorker.ReportProgress(60)
                        Dim mensajesSRI As String = ""
                        'lbestadoSRI.Text = resp.Estado
                        _EstadoAutorizacion = objetoRespuesta.Estado
                        _ClaveAcceso = objetoRespuesta.ClaveAcceso

                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Respuesta del SRI: " + _EstadoAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)



                        If _tipoManejo = "A" Then
                            ' Seteo el Error recibido del servicio EDOC
                            rsboApp.SetStatusBarMessage("Recibiendo respuesta..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            'Dim respuestaaa = objetoRespuesta.Autorizaciones(0).Mensajes(0).mensaje1().ToString & "- " & objetoRespuesta.Autorizaciones(0).Mensajes(0).informacionAdicional().ToString
                        End If

                        'If objetoRespuesta.Estado = 7 Or objetoRespuesta.Estado = 5 Then
                        '    _Observacion = "Estado: AUTORIZADO, # Autorizacion: " + objetoRespuesta.autorizaciones(0).numeroAutorizacion().ToString() + " - Ambiente: " + objetoRespuesta.autorizaciones(0).ambiente().ToString()

                        'ElseIf objetoRespuesta.Estado = "AUTORIZADO" Or objetoRespuesta.Estado = "2" Then

                        '    _Observacion = "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA" & " - NÚMERO DEL DOCUMENTO: " + DocEntry.ToString()

                        'Else

                        '    _Observacion = ""

                        'End If

                        If TipoDocumento = "FCE" Or TipoDocumento = "FRE" Or TipoDocumento = "FAE" Then
                            If TipoWS = "LOCAL" Then
                                _Observacion = recorreErrorFactura_LOCAL(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE" Then
                                _Observacion = recorreErrorFactura(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE_4_1" Then
                                _Observacion = recorreErrorFactura_NUBE41(objetoRespuesta, DocEntry.ToString())
                            End If

                        ElseIf TipoDocumento = "NDE" Then
                            If TipoWS = "LOCAL" Then
                                _Observacion = recorreErrorNotaDebito_LOCAL(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE" Then
                                _Observacion = recorreErrorNotaDebito(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE_4_1" Then
                                _Observacion = recorreErrorNotaDebito_NUBE41(objetoRespuesta, DocEntry.ToString())
                            End If

                        ElseIf TipoDocumento = "NCE" Then
                            If TipoWS = "LOCAL" Then
                                _Observacion = recorreErrorNotaCredito_LOCAL(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE" Then
                                _Observacion = recorreErrorNotaCredito(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE_4_1" Then
                                _Observacion = recorreErrorNotaCredito_NUBE41(objetoRespuesta, DocEntry.ToString())
                            End If

                        ElseIf TipoDocumento = "GRE" Or TipoDocumento = "TRE" Or TipoDocumento = "TLE" Or TipoDocumento = "SSGR" Then
                            If TipoWS = "LOCAL" Then
                                _Observacion = recorreErrorGuiaRemision_LOCAL(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE" Then
                                _Observacion = recorreErrorGuiaRemision(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE_4_1" Then
                                _Observacion = recorreErrorGuiaRemision_NUBE41(objetoRespuesta, DocEntry.ToString())
                            End If

                        ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
                            If TipoWS = "LOCAL" Then
                                _Observacion = recorreErrorRetencion_LOCAL(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE" Then
                                _Observacion = recorreErrorRetencion(objetoRespuesta, DocEntry.ToString())
                            ElseIf TipoWS = "NUBE_4_1" Then
                                _Observacion = recorreErrorRetencion_NUBE41(objetoRespuesta, DocEntry.ToString())
                            End If

                        End If
                        If TipoDocumento = "LQE" Then
                            '_Observacion = recorreErrorLiquidacionCompra(objetoRespuesta, DocEntry.ToString())
                            If TipoWS = "LOCAL" Then
                                _Observacion = recorreErrorLiquidacionCompra_LOCAL(objetoRespuesta, DocEntry.ToString())
                            Else
                                _Observacion = recorreErrorLiquidacionCompra(objetoRespuesta, DocEntry.ToString())
                            End If
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
                                'mensajesSRI = _Observacion.ToString
                            Catch ex As Exception
                                mensajesSRI = " No se recibio la descripcion del Error "
                            End Try
                        End If

                        If _tipoManejo = "A" Then
                            rsboApp.SetStatusBarMessage("Grabando respuesta de SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            '  oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Grabando Respuesta del SRI en Documento - " + TipoDocumento + " - " + DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Respuesta SRI en Documento - " + TipoDocumento + " - DocEntry " + DocEntry.ToString() + " - # de Autorización - " + _NumAutorizacion.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                        End If

                        _Observacion = String.Format("Estado:{0} - # AUTORIZACION {1} - RespuestaSRI - {2} - Error - {3} ", _EstadoAutorizacion.ToString, _NumAutorizacion.ToString, mensajesSRI, _Observacion.ToString)

                        ' Mando a Grabar a SAP
                        If TipoDocumento = "LQE" Then
                            result = GrabaDatosAutorizacion_LiquidacionCompra(DocEntry, TipoDocumento)

                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                            result = GrabaDatosAutorizacion_Factura_GuiaRemision(DocEntry, TipoDocumento)

                        ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then
                            result = GrabaDatosAutorizacion_SalidaMercancias_GuiaRemision(DocEntry, TipoDocumento)

                        ElseIf TipoDocumento = "SSGR" Then

                            result = GrabaDatosAutGuiasDesatendidas(DocEntry, TipoDocumento)

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

                        If _tipoManejo = "A" Then
                            If Functions.VariablesGlobales._ImpDocAut = "Y" Then
                                If _EstadoAutorizacion.ToString().Equals("2") Or _EstadoAutorizacion.ToString().Equals("AUTORIZADO") Then
                                    rsboApp.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + " - Imprimiendo Documento por favor esperar... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    mensajeDocAut = ""
                                    If ImprmirDOcAut(_NumAutorizacion) Then
                                        rsboApp.SetStatusBarMessage("El documento se imprimió con éxito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    Else
                                        rsboApp.SetStatusBarMessage("Error al imprimir PDF: " + mensajeDocAut.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End If
                                End If
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
                                GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry, TipoDocumento, _Error)

                            ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                                GrabaDatosAutorizacion_Error_FacturaGuiaRemision(DocEntry, _Error)

                            ElseIf TipoDocumento = "SSGR" Then

                                GrabaDatosAutGuiasDesatendidas_Error(DocEntry, TipoDocumento, _Error)

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
                            GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry, TipoDocumento, _Error)

                        ElseIf Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                            GrabaDatosAutorizacion_Error_FacturaGuiaRemision(DocEntry, _Error)

                        ElseIf Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then
                            GrabaDatosAutorizacion_Error_SalidaMercanciasGuiaRemision(DocEntry, _Error)

                        ElseIf TipoDocumento = "SSGR" Then

                            GrabaDatosAutGuiasDesatendidas_Error(DocEntry, TipoDocumento, _Error)

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

    Public Function SincronizarDocumentoEDOC(ByVal tipoDocumento As String, DocEntry As String, ByVal TipoWS As String) As Object
        ' ws.Url = Url ' Seteo la URL en el servicio web
        ' Entorno 2- en Linea, 1- en Batch

        Dim ObjetoRespuesta As Object = Nothing
        Dim mensajeRespuesta As String = ""
        _errorMensajeWSEnvío = ""
        Dim url As String = ""

        Dim SALIDA_POR_PROXY As String = ""
        SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
        Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
        Dim Proxy_puerto As String = ""
        Dim Proxy_IP As String = ""
        Dim Proxy_Usuario As String = ""
        Dim Proxy_Clave As String = ""

        'Dim wsauto As New WSAutorizacionComp.AutorizacionComprobantesService
        'wsauto.Url = url
        'wsauto.Timeout = 10000

        Try


            Dim WS_EmisionConsul As String

            If _tipoManejo = "S" Then
                WS_EmisionConsul = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionConsulta")
            Else
                WS_EmisionConsul = Functions.VariablesGlobales._wsConsultaEmision
            End If



            'OBTENER INFORMACION COMPANIA PARA SINCRONIZAR
            'RUC COMPANIA
            'TIPO DOC
            'NUMDOC xxx-xxx-xxxxxxxxx  en este formato
            'SECERP

            Dim Sincro_ruc As String = "", Sincro_Tipo_doc As String = "", Sincro_sec_ERP As String = "", Sincro_Num_Doc As String



            '--------------------------
            'numero de documento xxx-xxx-xxxxxxxxx  en este formato y ruc
            Dim info_company_numdoc() As String = Get_company_numdoc_by_proveedor(_Nombre_Proveedor_SAP_BO, DocEntry, tipoDocumento)
            'RUC compania

            Sincro_ruc = info_company_numdoc(0) 'cero para ruc
            'NUM Doc

            Sincro_Num_Doc = info_company_numdoc(1) 'uno numero doc


            'tipo documento
            Sincro_Tipo_doc = ObtnerTipoDocumentoEDOC(tipoDocumento)
            'secuencial
            Sincro_sec_ERP = DocEntry

            If Sincro_ruc = "" Or Sincro_Num_Doc = "" Then
                Return Nothing
            End If


            If WS_EmisionConsul = "" Then
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                Return Nothing
            End If

            Dim ws As Object

            If TipoWS = "LOCAL" Then
                ObjetoRespuesta = New Entidades.wsEDoc_ConsultaEmision_LOCAL.RespuestaConProcesoEDOC
                ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
                ws.Url = WS_EmisionConsul
                If SALIDA_POR_PROXY = "Y" Then

                    Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                    Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                    Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                    Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                    Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                    If Not Proxy_puerto = "" Then
                        proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                    Else
                        proxyobject = New System.Net.WebProxy(Proxy_IP)
                    End If
                    cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                    proxyobject.Credentials = cred

                    ws.Proxy = proxyobject
                    ws.Credentials = cred

                End If

                SetProtocolosdeSeguridad()
                'OBJETO CON LA RESPUESTA DEL ESTADO EDOC
                ObjetoRespuesta = ws.ConsultarProcesoSincronizadorAX(Sincro_ruc, Sincro_Tipo_doc, Sincro_Num_Doc, Sincro_sec_ERP)

            ElseIf TipoWS = "NUBE" Then

                ObjetoRespuesta = New Entidades.wsEDoc_ConsultaEmision.RespuestaConProcesoEDOC
                ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                'Dim x As New Entidades.wsEDoc_Factura.WSEDOCNUBE_FACTURAS
                ws.Url = WS_EmisionConsul

                If SALIDA_POR_PROXY = "Y" Then

                    Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                    Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                    Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                    Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                    Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                    If Not Proxy_puerto = "" Then
                        proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                    Else
                        proxyobject = New System.Net.WebProxy(Proxy_IP)
                    End If
                    cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                    proxyobject.Credentials = cred

                    ws.Proxy = proxyobject
                    ws.Credentials = cred

                End If
                'OBJETO CON LA RESPUESTA DEL ESTADO EDOC
                ObjetoRespuesta = ws.ConsultarProcesoSincronizadorAX(Sincro_ruc, Sincro_Tipo_doc, Sincro_Num_Doc, Sincro_sec_ERP)

            ElseIf TipoWS = "NUBE_4_1" Then
                'ObjetoRespuesta = New Entidades.wsEDoc_ConsultaEmision.RespuestaConProcesoEDOC
                'ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                'ws.Url = WS_EmisionConsul

                ObjetoRespuesta = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.RespuestaConProcesoEDOC
                ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA
                ws.Url = WS_EmisionConsul


                If SALIDA_POR_PROXY = "Y" Then

                    Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                    Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                    Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                    Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                    Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                    Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                    If Not Proxy_puerto = "" Then
                        proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                    Else
                        proxyobject = New System.Net.WebProxy(Proxy_IP)
                    End If
                    cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                    proxyobject.Credentials = cred

                    ws.Proxy = proxyobject
                    ws.Credentials = cred

                End If
                'OBJETO CON LA RESPUESTA DEL ESTADO EDOC
                'oObjeto
                SetProtocolosdeSeguridad()

                ObjetoRespuesta = ws.ConsultarProcesoSincronizadorAX(Sincro_ruc, Sincro_Tipo_doc, Sincro_Num_Doc, Sincro_sec_ERP)

            End If




            If Not mensajeRespuesta = "" Then
                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, mensajeRespuesta, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _errorMensajeWSEnvío = mensajeRespuesta
            End If

            If Not ObjetoRespuesta Is Nothing Then
                _NumeroDeDocumentoSRI = Sincro_Num_Doc.ToString
                Return ObjetoRespuesta
            Else
                Return Nothing
            End If

        Catch tx As TimeoutException
            'resp.Estado = "7 En Espera SRI"
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("(GS) WS : " + tx.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, "WS : " + tx.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return ObjetoRespuesta
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("(GS) WS : " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, "WS : " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If
            Return Nothing
        End Try

    End Function

    Public Function EnviaDocumentoSRI(oObjeto As Object, tipoDocumento As String, DocEntry As String, ByVal TipoWS As String, ByRef mensajeRespuesta As String) As Object
        ' ws.Url = Url ' Seteo la URL en el servicio web
        ' Entorno 2- en Linea, 1- en Batch

        Dim ObjetoRespuesta As Object = Nothing
        Utilitario.Util_Log.Escribir_Log("Objeto respuesta ", "ManejoDeDocumentos")
        ' Dim mensajeRespuesta As String = ""
        _errorMensajeWSEnvío = ""
        Dim url As String = ""

        Dim SALIDA_POR_PROXY As String = ""
        SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
        Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
        Dim Proxy_puerto As String = ""
        Dim Proxy_IP As String = ""
        Dim Proxy_Usuario As String = ""
        Dim Proxy_Clave As String = ""

        'Dim wsauto As New WSAutorizacionComp.AutorizacionComprobantesService
        'wsauto.Url = url
        'wsauto.Timeout = 10000
        Utilitario.Util_Log.Escribir_Log("TIPO DOCUMENTO : " + tipoDocumento.ToString, "ManejoDeDocumentos")
        Try

            Dim ClaveWS As String = ""
            Dim TipoEmisionWS As String = ""
            If _tipoManejo = "S" Then
                ClaveWS = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EmisionClave")
                TipoEmisionWS = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EmisionTipo")
            Else
                ClaveWS = Functions.VariablesGlobales._wsClaveEmision
                TipoEmisionWS = Functions.VariablesGlobales._TipoEmision
            End If

            If ClaveWS = "" Then
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No existe informacion de la clave de los Web Services, revisar parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

                    oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, "No existe informacion de la clave de los Web Services, revisar parametrización, DocEntry: " & DocEntry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                Exit Function
            End If

            If tipoDocumento = "FCE" Or tipoDocumento = "FRE" Or tipoDocumento = "FAE" Then

                Dim WS_EmisionFC As String = ""
                If _tipoManejo = "S" Then
                    WS_EmisionFC = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionFC")
                Else
                    WS_EmisionFC = Functions.VariablesGlobales._wsEmisionFactura
                End If

                If WS_EmisionFC = "" Then
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Exit Function
                End If

                Dim ws As Object
                If TipoWS = "LOCAL" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_Factura_LOCAL.RespuestaEDOC
                    ws = New Entidades.wsEDoc_Factura_LOCAL.WSEDOC_FACTURAS
                    ws.Url = WS_EmisionFC
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarFacturaSRI(ClaveWS _
                                                          , TipoEmisionWS _
                                                          , DirectCast(oObjeto, Entidades.wsEDoc_Factura_LOCAL.ENTFactura), mensajeRespuesta)

                ElseIf TipoWS = "NUBE" Then

                    ObjetoRespuesta = New Entidades.wsEDoc_Factura.RespuestaEDOC
                    ws = New Entidades.wsEDoc_Factura.WSEDOCNUBE_FACTURAS
                    'Dim x As New Entidades.wsEDoc_Factura.WSEDOCNUBE_FACTURAS
                    ws.Url = WS_EmisionFC

                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    ' clave = gsedoc
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    If Functions.VariablesGlobales._vgHttps = "Y" Then
                        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    End If


                    ObjetoRespuesta = ws.EnviarFacturaSRI(ClaveWS _
                                                          , TipoEmisionWS _
                                                          , DirectCast(oObjeto, Entidades.wsEDoc_Factura.ENTFactura), mensajeRespuesta)
                ElseIf TipoWS = "NUBE_4_1" Then
                    ' oFuncionesAddon.GuardaLOG(TipoWS, DocEntry, "Factura 4_1 consulta web service= ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Utilitario.Util_Log.Escribir_Log("Factura 4_1 consulta web service" + TipoWS, "ManejoDeDocumentos")
                    ObjetoRespuesta = New Entidades.wsEDoc_Factura41.RespuestaEDOC
                    ws = New Entidades.wsEDoc_Factura41.WSEDOCNUBE_FACTURAS
                    ws.Url = WS_EmisionFC
                    ' oFuncionesAddon.GuardaLOG(TipoWS, DocEntry, "Factura 4_1 obtiene ws= " + WS_EmisionFC.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Utilitario.Util_Log.Escribir_Log("Factura 4_1 obtiene ws= " + WS_EmisionFC.ToString, "ManejoDeDocumentos")
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                        Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                        Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                        Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    ' clave = gsedoc
                    'oFuncionesAddon.GuardaLOG(TipoWS, DocEntry, "Factura 4_1 antes de consultar salida https= ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Utilitario.Util_Log.Escribir_Log("Factura 4_1 antes de consultar salida https ", "ManejoDeDocumentos")
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    Utilitario.Util_Log.Escribir_Log("Factura 4_1 despues de consultar salida https ", "ManejoDeDocumentos")
                    'oFuncionesAddon.GuardaLOG(TipoWS, DocEntry, "Factura 4_1 consultar salida https= ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Utilitario.Util_Log.Escribir_Log("envio documento sri clave ws: " + ClaveWS.ToString + " tipo emiion: " + TipoEmisionWS.ToString + " objeto: " + oObjeto.ToString, "ManejoDeDocumentos")
                    ActivarTLS()
                    SetProtocolosdeSeguridad()

                    Try

                        ObjetoRespuesta = ws.EnviarFacturaSRI(ClaveWS _
                                                          , TipoEmisionWS _
                                                          , DirectCast(oObjeto, Entidades.wsEDoc_Factura41.ENTFactura), mensajeRespuesta)


                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("error envio documento sri " + ex.Message + IIf(String.IsNullOrEmpty(mensajeRespuesta), "", mensajeRespuesta), "ManejoDeDocumentos")
                    End Try

                End If

            ElseIf tipoDocumento = "NCE" Then
                Dim WS_EmisionNC As String = ""

                If _tipoManejo = "S" Then
                    WS_EmisionNC = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionNC")
                Else
                    WS_EmisionNC = Functions.VariablesGlobales._wsEmisionNotaCredito
                End If

                If WS_EmisionNC = "" Then
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Exit Function
                End If

                Dim ws As Object
                If TipoWS = "LOCAL" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_NotaDeCredito_LOCAL.RespuestaEDOC
                    ws = New Entidades.wsEDoc_NotaDeCredito_LOCAL.WSEDOC_NOTAS_CREDITO
                    ws.Url = WS_EmisionNC
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarNotaCreditoSRI(ClaveWS _
                                                              , TipoEmisionWS _
                                                              , DirectCast(oObjeto, Entidades.wsEDoc_NotaDeCredito_LOCAL.ENTNotaCredito), mensajeRespuesta)
                ElseIf TipoWS = "NUBE" Then

                    ObjetoRespuesta = New Entidades.wsEDoc_NotaDeCredito.RespuestaEDOC
                    ws = New Entidades.wsEDoc_NotaDeCredito.WSEDOCNUBE_NOTAS_CREDITO
                    ws.Url = WS_EmisionNC
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    If Functions.VariablesGlobales._vgHttps = "Y" Then
                        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    End If
                    ObjetoRespuesta = ws.EnviarNotaCreditoSRI(ClaveWS _
                                                              , TipoEmisionWS _
                                                              , DirectCast(oObjeto, Entidades.wsEDoc_NotaDeCredito.ENTNotaCredito), mensajeRespuesta)
                ElseIf TipoWS = "NUBE_4_1" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_NotaDeCredito41.RespuestaEDOC
                    ws = New Entidades.wsEDoc_NotaDeCredito41.WSEDOCNUBE_NOTAS_CREDITO
                    ws.Url = WS_EmisionNC
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                        Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                        Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                        Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarNotaCreditoSRI(ClaveWS _
                                                              , TipoEmisionWS _
                                                              , DirectCast(oObjeto, Entidades.wsEDoc_NotaDeCredito41.ENTNotaCredito), mensajeRespuesta)
                End If


            ElseIf tipoDocumento = "NDE" Then
                Dim WS_EmisionND As String = ""
                If _tipoManejo = "S" Then
                    WS_EmisionND = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionND")
                Else
                    WS_EmisionND = Functions.VariablesGlobales._wsEmisionNotaDebito
                End If

                If WS_EmisionND = "" Then
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Exit Function
                End If

                Dim ws As Object
                If TipoWS = "LOCAL" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_NotaDeDebito_LOCAL.RespuestaEDOC
                    ws = New Entidades.wsEDoc_NotaDeDebito_LOCAL.WSEDOC_NOTAS_DEBITO
                    ws.Url = WS_EmisionND
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarNotaDebitoSRI(ClaveWS _
                                                             , TipoEmisionWS _
                                                             , DirectCast(oObjeto, Entidades.wsEDoc_NotaDeDebito_LOCAL.ENTNotaDebito), mensajeRespuesta)
                ElseIf TipoWS = "NUBE" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_NotaDeDebito.RespuestaEDOC
                    ws = New Entidades.wsEDoc_NotaDeDebito.WSEDOCNUBE_NOTAS_DEBITO
                    ws.Url = WS_EmisionND
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    If Functions.VariablesGlobales._vgHttps = "Y" Then
                        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    End If
                    ObjetoRespuesta = ws.EnviarNotaDebitoSRI(ClaveWS _
                                                             , TipoEmisionWS _
                                                             , DirectCast(oObjeto, Entidades.wsEDoc_NotaDeDebito.ENTNotaDebito), mensajeRespuesta)
                ElseIf TipoWS = "NUBE_4_1" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_NotaDeDebito41.RespuestaEDOC
                    ws = New Entidades.wsEDoc_NotaDeDebito41.WSEDOCNUBE_NOTAS_DEBITO
                    ws.Url = WS_EmisionND
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                        Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                        Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                        Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarNotaDebitoSRI(ClaveWS _
                                                             , TipoEmisionWS _
                                                             , DirectCast(oObjeto, Entidades.wsEDoc_NotaDeDebito41.ENTNotaDebito), mensajeRespuesta)
                End If


            ElseIf tipoDocumento = "GRE" Or tipoDocumento = "TRE" Or tipoDocumento = "TLE" Or tipoDocumento = "SSGR" Then
                Dim ws As Object

                Dim WS_EmisionGuia As String = ""
                If _tipoManejo = "S" Then
                    WS_EmisionGuia = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionGuia")
                Else
                    WS_EmisionGuia = Functions.VariablesGlobales._wsEmisionGuiaRemision
                End If

                If WS_EmisionGuia = "" Then
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Exit Function
                End If

                If TipoWS = "LOCAL" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_GuiaRemision_LOCAL.RespuestaEDOC
                    ws = New Entidades.wsEDoc_GuiaRemision_LOCAL.WSEDOC_GUIAS_REMISION
                    ws.Url = WS_EmisionGuia
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarGuiaRemisionSRI(ClaveWS _
                                                               , TipoEmisionWS _
                                                               , DirectCast(oObjeto, Entidades.wsEDoc_GuiaRemision_LOCAL.ENTGuiaRemision), mensajeRespuesta)
                ElseIf TipoWS = "NUBE" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_GuiaRemision.RespuestaEDOC
                    ws = New Entidades.wsEDoc_GuiaRemision.WSEDOCNUBE_GUIAS_REMISION
                    ws.Url = WS_EmisionGuia
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    If Functions.VariablesGlobales._vgHttps = "Y" Then
                        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    End If
                    ObjetoRespuesta = ws.EnviarGuiaRemisionSRI(ClaveWS _
                                                           , TipoEmisionWS _
                                                           , DirectCast(oObjeto, Entidades.wsEDoc_GuiaRemision.ENTGuiaRemision), mensajeRespuesta)
                ElseIf TipoWS = "NUBE_4_1" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_GuiaRemision41.RespuestaEDOC
                    ws = New Entidades.wsEDoc_GuiaRemision41.WSEDOCNUBE_GUIAS_REMISION
                    ws.Url = WS_EmisionGuia
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                        Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                        Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                        Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)
                        proxyobject.Credentials = cred
                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarGuiaRemisionSRI(ClaveWS _
                                                               , TipoEmisionWS _
                                                               , DirectCast(oObjeto, Entidades.wsEDoc_GuiaRemision41.ENTGuiaRemision), mensajeRespuesta)
                End If


            ElseIf tipoDocumento = "RDM" Or tipoDocumento = "REE" Or tipoDocumento = "REA" Or tipoDocumento = "RER" Then
                Dim ws As Object

                Dim WS_EmisionRetencion As String = ""
                If _tipoManejo = "S" Then
                    WS_EmisionRetencion = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionRetencion")
                Else
                    WS_EmisionRetencion = Functions.VariablesGlobales._wsEmisionRetencion
                End If

                If WS_EmisionRetencion = "" Then
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Exit Function
                End If

                If TipoWS = "LOCAL" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_Retencion_LOCAL.RespuestaEDOC
                    ws = New Entidades.wsEDoc_Retencion_LOCAL.WSEDOC_RETENCIONES
                    ws.Url = WS_EmisionRetencion
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If

                    SetProtocolosdeSeguridad()

                    ObjetoRespuesta = ws.EnviarRetencionSRI(ClaveWS _
                                                            , TipoEmisionWS _
                                                            , DirectCast(oObjeto, Entidades.wsEDoc_Retencion_LOCAL.ENTRetencion), mensajeRespuesta)
                ElseIf TipoWS = "NUBE" Then

                    ObjetoRespuesta = New Entidades.wsEDoc_Retencion.RespuestaEDOC
                    ws = New Entidades.wsEDoc_Retencion.WSEDOCNUBE_RETENCIONES
                    ws.Url = WS_EmisionRetencion
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    If Functions.VariablesGlobales._vgHttps = "Y" Then
                        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    End If

                    ObjetoRespuesta = ws.EnviarRetencionSRI(ClaveWS _
                                                            , TipoEmisionWS _
                                                            , DirectCast(oObjeto, Entidades.wsEDoc_Retencion.ENTRetencion), mensajeRespuesta)

                ElseIf TipoWS = "NUBE_4_1" Then
                    Utilitario.Util_Log.Escribir_Log("tipo ws : " + TipoWS.ToString + " tipo documento: " + tipoDocumento.ToString, "ManejoDeDocumentos")
                    ObjetoRespuesta = New Entidades.wsEDoc_Retencion41.RespuestaEDOC
                    ws = New Entidades.wsEDoc_Retencion41.WSEDOCNUBE_RETENCIONES
                    ws.Url = WS_EmisionRetencion
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                        Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                        Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                        Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred
                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    Utilitario.Util_Log.Escribir_Log("enviando por https", "ManejoDeDocumentos")
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarRetencionSRI(ClaveWS _
                                                            , TipoEmisionWS _
                                                            , DirectCast(oObjeto, Entidades.wsEDoc_Retencion41.ENTRetencion), mensajeRespuesta)

                End If

            ElseIf tipoDocumento = "LQE" Then
                Dim ws As Object 'New Entidades.wsEDoc_LiquidacionCompra.WSEDOCNUBE_LIQUIDACIONES_COMPRA

                Dim WS_EmisionLiquidacionCompra As String = ""
                If _tipoManejo = "S" Then
                    WS_EmisionLiquidacionCompra = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_LiquidacionCompra")
                Else
                    WS_EmisionLiquidacionCompra = Functions.VariablesGlobales._wsEmisionLiquidacionCompra
                End If

                If WS_EmisionLiquidacionCompra = "" Then
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    Exit Function
                End If

                If TipoWS = "LOCAL" Then
                    ObjetoRespuesta = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.RespuestaEDOC
                    ws = New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.WSEDOC_LIQUIDACIONES_COMPRA
                    ws.Url = WS_EmisionLiquidacionCompra
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If

                    SetProtocolosdeSeguridad()

                    ObjetoRespuesta = ws.EnviarLiquidacionCompraSRI(ClaveWS _
                                                            , TipoEmisionWS _
                                                            , DirectCast(oObjeto, Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.ENTLiquidacionCompra), mensajeRespuesta)
                ElseIf TipoWS = "NUBE" Then

                    ObjetoRespuesta = New Entidades.wsEDoc_LiquidacionCompra.RespuestaEDOC
                    ws = New Entidades.wsEDoc_LiquidacionCompra.WSEDOCNUBE_LIQUIDACIONES_COMPRA
                    ws.Url = WS_EmisionLiquidacionCompra
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred

                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    End If
                    ObjetoRespuesta = ws.EnviarLiquidacionCompraSRI(ClaveWS _
                                                            , TipoEmisionWS _
                                                            , DirectCast(oObjeto, Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompra), mensajeRespuesta)

                ElseIf TipoWS = "NUBE_4_1" Then
                    ObjetoRespuesta = New Entidades.wsEDoc_LiquidacionCompra.RespuestaEDOC
                    ws = New Entidades.wsEDoc_LiquidacionCompra.WSEDOCNUBE_LIQUIDACIONES_COMPRA
                    ws.Url = WS_EmisionLiquidacionCompra
                    If SALIDA_POR_PROXY = "Y" Then

                        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                        If Not Proxy_puerto = "" Then
                            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                        Else
                            proxyobject = New System.Net.WebProxy(Proxy_IP)
                        End If
                        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                        proxyobject.Credentials = cred
                        ws.Proxy = proxyobject
                        ws.Credentials = cred

                    End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    Utilitario.Util_Log.Escribir_Log("ENVIO", "ManejoDeDocumentos")
                    SetProtocolosdeSeguridad()
                    ObjetoRespuesta = ws.EnviarLiquidacionCompraSRI(ClaveWS _
                                                            , TipoEmisionWS _
                                                            , DirectCast(oObjeto, Entidades.wsEDoc_LiquidacionCompra.ENTLiquidacionCompra), mensajeRespuesta)
                    Utilitario.Util_Log.Escribir_Log("Retorno", "ManejoDeDocumentos")

                End If

            End If
            If _tipoManejo = "A" Then
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            End If
            If Not mensajeRespuesta = "" Then
                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, mensajeRespuesta, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _errorMensajeWSEnvío = mensajeRespuesta
            End If

            If Not ObjetoRespuesta Is Nothing Then
                _NumeroDeDocumentoSRI = oObjeto.Establecimiento + "-" + oObjeto.PuntoEmision + "-" + oObjeto.Secuencial
                Return ObjetoRespuesta
            Else
                Return Nothing
            End If

        Catch tx As TimeoutException
            'resp.Estado = "7 En Espera SRI"
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("(GS)" + tx.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, "WS : " + tx.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return ObjetoRespuesta
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("(GS)" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
            If _tipoManejo = "A" Then

                oFuncionesAddon.GuardaLOG(tipoDocumento, DocEntry, "WS : " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

            End If

            Return Nothing
        End Try

#Disable Warning BC42105 ' La función 'EnviaDocumentoSRI' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'EnviaDocumentoSRI' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Public Function GrabaDatosAutGuiasDesatendidas_Error(DocEntry As Integer, TipoDocumento As String, MsgError As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService


            ' SI EXISTE ELIMINA PARA VOLVER A CREAR
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSGRNEW")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntry)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            Try
                'oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                oGeneralData.SetProperty("U_OBSERVACION_FACT", _Observacion)

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try


            Try
                oGeneralService.Update(oGeneralData)
                resultado = 0
            Catch ex As Exception
                resultado = 1
            End Try


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

    Public Function GrabaDatosAutGuiasDesatendidas(DocEntry As Integer, TipoDocumento As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim objectType As String = "" 'obtener el objtype del documento para la localizacion de topmanage
        Dim CodDoc As String = "" 'obtener el codigo del documento para la localizacion de topmanage
        Dim SerieDoc As String = ""
        Try

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService


            ' SI EXISTE ELIMINA PARA VOLVER A CREAR
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSGRNEW")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntry)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)


            If _NumAutorizacion <> "" Then


                oGeneralData.SetProperty("U_SS_NumAut", _NumAutorizacion.ToString())

                Try

                    oGeneralData.SetProperty("U_NUM_AUTO_FAC", _NumAutorizacion.ToString())
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_NUM_AUTO_FAC error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try


                If _tipoManejo = "A" Then
                    Try
                        rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                End If

            End If

            'campos normales

            Try
                'oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                ' oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion

                oGeneralData.SetProperty("U_FECHA_AUT_FACT", _FechaAutorizacion)

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            '-------------
            Try
                'oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                oGeneralData.SetProperty("U_OBSERVACION_FACT", _Observacion)

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            '--------------
            Try
                '  oTransferencia.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)

                oGeneralData.SetProperty("U_ESTADO_AUTORIZACIO", IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion))

            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("U_ESTADO_AUTORIZACIO error : " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            '--------------
            If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                Try
                    ' oTransferencia.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAcceso.ToString()
                    oGeneralData.SetProperty("U_CLAVE_ACCESO", _ClaveAcceso.ToString())

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_CLAVE_ACCESO error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            End If
            '---------------------

            If Not String.IsNullOrEmpty(_FechaAutorizacion.ToString) Then
                Try
                    ' oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion

                    oGeneralData.SetProperty("U_FECHA_AUT_FACT", _FechaAutorizacion)

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            End If

            Try
                oGeneralService.Update(oGeneralData)
                resultado = 0
            Catch ex As Exception
                resultado = 1
            End Try



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

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
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

            ElseIf TipoDocumento = "NDE" Then ''FACTURA DE ANTICIPO DE CLIENTES
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "05"

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

            ElseIf TipoDocumento = "GRE" Then 'GUIA DE REMISION - ENTREGA
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes
                ' objectType = oDocumento.DocObjectCode
                'CodDoc = "06"

            ElseIf TipoDocumento = "TRE" Then 'GUIA DE REMISION - TRANSFERENCIAS
                oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "06"

            ElseIf TipoDocumento = "TLE" Then 'GUIA DE REMISION - SOLICITUD TRANSLADOS
                oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
                oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
                ' objectType = oDocumento.DocObjectCode
                ' CodDoc = "06"

            ElseIf TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then  'FACTURA DE PROVEEDOR/RETENCION                             
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "07"

            ElseIf TipoDocumento = "RDM" Then
                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo
                'objectType = oDocumento.DocObjectCode
                'CodDoc = "07"

            End If

            If TipoDocumento = "TRE" Or TipoDocumento = "TLE" Then
                If oTransferencia.GetByKey(DocEntry) Then
                    'oInvoice.Comments += "Procesada por la Plataforma de Integracion"
                    If _NumAutorizacion <> "" Then

                        If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                            oTransferencia.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacion.ToString()
                        ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                            oTransferencia.UserFields.Fields.Item("U_NO_AUTORI").Value = _NumAutorizacion.ToString()
                        ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                            oTransferencia.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _NumAutorizacion.ToString()
                            Try
                                oTransferencia.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString()
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("U_HBT_IdEnProveedor error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                            Try
                                oTransferencia.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString()
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                            If GrabaDatosAutorizacion_HESION_GUIA(TipoDocumento, DocEntry) Then
                                Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")

                            End If
                            If GrabaDatosAutorizacion_HESION_GUIA_TRANSFERENCIAS(TipoDocumento, DocEntry) Then
                                Utilitario.Util_Log.Escribir_Log("Se guardaron los datos de autorizacion en las transferencias incluidas en la guia de remision ", "ManejoDeDocumentos")

                            End If

                        ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                            Try
                                oTransferencia.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _NumAutorizacion.ToString()
                            Catch ex As Exception

                                Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTO error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                        ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                            Try
                                oTransferencia.UserFields.Fields.Item("U_TM_NAUT").Value = _NumAutorizacion.ToString()
                                oTransferencia.UserFields.Fields.Item("U_TM_DATEA").Value = Date.Now
                                If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                                    Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla Control de TM", "ManejoDeDocumentos")

                                End If
                            Catch ex As Exception

                                Utilitario.Util_Log.Escribir_Log("U_TM_NAUT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try
                        ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                            oTransferencia.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacion.ToString()
                        End If


                        Try
                            oTransferencia.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacion.ToString()
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("U_NUM_AUTO_FAC error: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try


                        If _tipoManejo = "A" Then
                            Try
                                rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                            End Try

                        End If

                    End If
                    '------------
                    Try
                        'oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                        oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    '-------------
                    Try
                        oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_OBSERVACION_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    '--------------
                    Try
                        oTransferencia.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_ESTADO_AUTORIZACIO error : " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    '--------------
                    If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                        Try
                            oTransferencia.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAcceso.ToString()
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("U_CLAVE_ACCESO error: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try
                    End If
                    '---------------------

                    If Not String.IsNullOrEmpty(_FechaAutorizacion.ToString) Then
                        Try
                            oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try
                    End If

                    Try
                        resultado = oTransferencia.Update()
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("error al ejecutar la funcion update : " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                End If
            Else
                If oDocumento.GetByKey(DocEntry) Then


                    'oInvoice.Comments += "Procesada por la Plataforma de Integracion"
                    If _NumAutorizacion <> "" Then
                        oDocumento.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacion.ToString()

                        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Or TipoDocumento = "RDM" Then

                            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                oDocumento.UserFields.Fields.Item("U_NUM_AUT_RET").Value = _NumAutorizacion.ToString()
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                oDocumento.UserFields.Fields.Item("U_NA_RETENCION").Value = _NumAutorizacion.ToString()
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                Try
                                    oDocumento.UserFields.Fields.Item("U_HBT_AUT_RET").Value = _NumAutorizacion.ToString()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_AUT_RET errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                Try
                                    oDocumento.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_IdEnProveedor errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                                Try
                                    oDocumento.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                Try
                                    oDocumento.UserFields.Fields.Item("U_SYP_NROAUTOC").Value = _NumAutorizacion.ToString()
                                Catch ex As Exception
                                    oDocumento.UserFields.Fields.Item("U_SYP_NROAUTOO").Value = _NumAutorizacion.ToString()
                                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTOO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                                If GrabaDatosAutorizacion_UDORT_TM(TipoDocumento, DocEntry) Then
                                    If _tipoManejo = "A" Then
                                        rsboApp.SetStatusBarMessage("N° Autorización grabada en el UDO TM_LE_RETCH: " + _NumAutorizacion.ToString() + " Tipo Doc: " + TipoDocumento.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If
                                End If
                                If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                                    If _tipoManejo = "A" Then
                                        rsboApp.SetStatusBarMessage("N° Autorización grabada en la tabla Control Doc. Electrónicos: ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If
                                End If
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                oDocumento.UserFields.Fields.Item("U_SS_NumAutRet").Value = _NumAutorizacion.ToString()
                            End If

                        Else
                            If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                oDocumento.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacion.ToString()
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                oDocumento.UserFields.Fields.Item("U_NO_AUTORI").Value = _NumAutorizacion.ToString()
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                Try
                                    oDocumento.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _NumAutorizacion.ToString()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_AUT_FAC error: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try

                                Try
                                    oDocumento.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_IdEnProveedor errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                                Try
                                    oDocumento.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString()
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso : " + _ClaveAcceso.ToString, "ManejoDeDocumentos")
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_HBT_ClaveAcceso errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                                If TipoDocumento = "GRE" Then
                                    Utilitario.Util_Log.Escribir_Log("aantes de guardar en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")
                                    Utilitario.Util_Log.Escribir_Log("TipoDocumento" + TipoDocumento.ToString, "ManejoDeDocumentos")
                                    If GrabaDatosAutorizacion_HESION_GUIA(TipoDocumento, DocEntry) Then
                                        Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")
                                    End If
                                End If
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                Try
                                    oDocumento.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _NumAutorizacion.ToString()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                                Try
                                    oDocumento.UserFields.Fields.Item("U_TM_NAUT").Value = _NumAutorizacion.ToString()
                                    oDocumento.UserFields.Fields.Item("U_TM_DATEA").Value = Date.Now
                                    If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("N° Autorización grabada en la tabla Control Doc. Electrónicos: ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        End If
                                    End If

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("U_TM_NAUT error: " + ex.Message.ToString, "ManejoDeDocumentos")
                                End Try
                            ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                oDocumento.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacion.ToString()

                            End If

                            Try 'SI PARAMETRO ESTA ACTIVO, GUARDA EL NUMERO DE DOCUMENTO QUE SE ENVIÓ AL SRI EN EL CAMPO NUMATCARD
                                If Functions.VariablesGlobales._AsignarNumeroDocEnNumAtCard = "Y" Then
                                    '_NumeroDeDocumentoSRI = ""
                                    '_NumeroDeDocumentoSRI = oObjeto.Establecimiento + "-" + oObjeto.PuntoEmision + "-" + oObjeto.Secuencial
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
                        'oDocumento.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                        oDocumento.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try

                    'Try
                    '    'NUEVO CAMPO SOLICITADO POR CRECOS QUE GURDE FECHA Y HORA
                    '    oDocumento.UserFields.Fields.Item("U_GSF_NUM_COM").Value = _FechaAutorizacion.ToString
                    'Catch ex As Exception
                    '    Utilitario.Util_Log.Escribir_Log("U_FECHA_AUT_FACT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                    'End Try

                    Try
                        oDocumento.UserFields.Fields.Item("U_SYP_FECAUTOC").Value = Date.Now
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("U_SYP_FECAUTOC DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                    End Try
                    If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        oDocumento.UserFields.Fields.Item("U_TM_DATEA").Value = Date.Now
                    End If
                    'If Len(_Observacion) > 250 Then
                    '    oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.Substring(1, 153).ToString()
                    'Else
                    '    oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                    'End If
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

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_LiquidacionCompra(DocEntry As Integer, TipoDocumento As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents = Nothing


            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
            If oDocumento.GetByKey(DocEntry) Then
                'If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                '    oDocumento.UserFields.Fields.Item("U_LQ_NUM_AUTO").Value = _NumAutorizacion.ToString()
                'ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                '    'oDocumento.UserFields.Fields.Item("U_NA_RETENCION").Value = _NumAutorizacion.ToString()
                'ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                '    'oDocumento.UserFields.Fields.Item("U_HBT_AUT_RET").Value = _NumAutorizacion.ToString()
                'ElseIf _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

                'End If
                Try
                    oDocumento.UserFields.Fields.Item("U_NUM_AUTOR").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_NUM_AUTOR errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                'SEIDOR
                Try
                    oDocumento.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_NUM_AUTO").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_NUM_AUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    oDocumento.UserFields.Fields.Item("U_NO_AUTORI").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_NO_AUTORI errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    oDocumento.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_HBT_AUT_FAC errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try

                    oDocumento.UserFields.Fields.Item("U_LQ_FECHA_AUT").Value = _FechaAutorizacion
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_FECHA_AUT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try

                    oDocumento.UserFields.Fields.Item("U_SYP_FECHAUTOR").Value = _FechaAutorizacion
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FECHAUTOR DIBEAL: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_OBSERVACION").Value = _Observacion.ToString
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_OBSERVACION errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_ESTADO LQ: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                    oDocumento.UserFields.Fields.Item("U_LQ_CLAVE").Value = _ClaveAcceso.ToString()
                End If
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    oDocumento.UserFields.Fields.Item("U_TM_NAUT").Value = _NumAutorizacion.ToString()
                    oDocumento.UserFields.Fields.Item("U_TM_DATEA").Value = Date.Now
                    If GrabaDatosAutorizacion_TablaTM(TipoDocumento, DocEntry) Then
                        Utilitario.Util_Log.Escribir_Log("Datos de autorizacion de Liquidacion grabados con éxito en la tabla Control Doc. Electronicos", "ManejoDeDocumentos")
                    End If
                End If
                If _Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    oDocumento.UserFields.Fields.Item("U_SS_NumAut").Value = _NumAutorizacion.ToString()

                End If
                resultado = oDocumento.Update()
            End If

            If _tipoManejo = "A" Then
                Try
                    rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True
            Else
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            If _tipoManejo = "A" Then
                'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

                oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  " & _Error.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                'End If
            End If

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_Factura_GuiaRemision(DocEntry As Integer, TipoDocumento As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents = Nothing


            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
            If oDocumento.GetByKey(DocEntry) Then

                Try
                    oDocumento.UserFields.Fields.Item("U_GR_CLAVE").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_CLAVE errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_GR_NUM_AUTO").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_NUM_AUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try

                    oDocumento.UserFields.Fields.Item("U_GR_FECHA_AUT").Value = _FechaAutorizacion
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_FECHA_AUT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try


                Try
                    oDocumento.UserFields.Fields.Item("U_GR_OBSERVACION").Value = _Observacion.ToString
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_GR_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_ESTADO: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try



                resultado = oDocumento.Update()


            End If

            If _tipoManejo = "A" Then
                Try
                    rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True

                If GrabaDatosAutorizacion_HESION_FACTURAGUIA(TipoDocumento, DocEntry) Then
                    Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")

                End If

            Else

                rCompany.GetLastError(ErrCode, ErrMsg)

                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            If _tipoManejo = "A" Then
                'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

                oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  " & _Error.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                'End If
            End If

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_SalidaMercancias_GuiaRemision(DocEntry As Integer, TipoDocumento As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents = Nothing


            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            oDocumento.DocObjectCode = BoObjectTypes.oInventoryGenExit
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
            If oDocumento.GetByKey(DocEntry) Then

                Try
                    oDocumento.UserFields.Fields.Item("U_GR_CLAVE").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_CLAVE errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_GR_NUM_AUTO").Value = _NumAutorizacion.ToString()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_NUM_AUTO errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try

                    oDocumento.UserFields.Fields.Item("U_GR_FECHA_AUT").Value = _FechaAutorizacion
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_FECHA_AUT errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try


                Try
                    oDocumento.UserFields.Fields.Item("U_GR_OBSERVACION").Value = _Observacion.ToString
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION errorgetbykey: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    oDocumento.UserFields.Fields.Item("U_GR_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_ESTADO: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try



                resultado = oDocumento.Update()


            End If

            If _tipoManejo = "A" Then
                Try
                    rsboApp.SetStatusBarMessage("(GS) N° Autorización: " + _NumAutorizacion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("(GS) N° Autorización: error " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True

                If GrabaDatosAutorizacion_HESION_SALIDAGUIA(TipoDocumento, DocEntry) Then
                    Utilitario.Util_Log.Escribir_Log("Se guardaron los datos exitosamente en la tabla HBT_GUIAREMISION", "ManejoDeDocumentos")

                End If

            Else

                rCompany.GetLastError(ErrCode, ErrMsg)

                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Ocurrio un error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End If

                If _tipoManejo = "A" Then

                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  #Error: " + ErrCode.ToString() + " Mensaje: " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                End If
                _Error = ErrCode.ToString() + "-" + ErrMsg
            End If

        Catch ex As Exception
            result = False
            _Error = ex.Message
            If _tipoManejo = "A" Then
                'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "GuardarLog") = "Y" Then

                oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntry, "Error al grabar datos de Autorización LQ :  " & _Error.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                'End If
            End If

            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_Factura/GrabaDatosAutorizacion Usuario: " + _ConexionSAP.SBO_Application.Company.DatabaseName.ToString() + " - " + _ConexionSAP.SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
        End Try

        Return result
    End Function

    Public Function GrabaDatosAutorizacion_Error_LiquidacionCompra(DocEntry As Integer, TipoDocumento As String, MsgError As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            If oDocumento.GetByKey(DocEntry) Then
                Try
                    oDocumento.UserFields.Fields.Item("U_LQ_OBSERVACION").Value = MsgError
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_LQ_OBSERVACION error linea 4497: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    resultado = oDocumento.Update()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("error en linea 4503: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

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

    Public Function GrabaDatosAutorizacion_Error_FacturaGuiaRemision(DocEntry As Integer, MsgError As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String = ""

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            If oDocumento.GetByKey(DocEntry) Then
                Try
                    oDocumento.UserFields.Fields.Item("U_GR_OBSERVACION").Value = MsgError
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION error : " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    resultado = oDocumento.Update()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True
            Else

                rCompany.GetLastError(ErrCode, ErrMsg)

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

    Public Function GrabaDatosAutorizacion_Error_SalidaMercanciasGuiaRemision(DocEntry As Integer, MsgError As String) As Boolean
        Dim result As Boolean = False
        Dim resultado As Integer = -1

        Dim ErrCode As Long
        Dim ErrMsg As String = ""

        Try
            Dim oDocumento As SAPbobsCOM.Documents
            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenExit
            oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            If oDocumento.GetByKey(DocEntry) Then
                Try
                    oDocumento.UserFields.Fields.Item("U_GR_OBSERVACION").Value = MsgError
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION error : " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

                Try
                    resultado = oDocumento.Update()
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_GR_OBSERVACION error: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try

            End If

            If resultado = 0 Then
                result = True
            Else

                rCompany.GetLastError(ErrCode, ErrMsg)

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
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    CODE = "SELECT IFNULL(""U_Estable"",'0') AS Establecimiento, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" =79 "
        '    _code = "SELECT IFNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = 79"
        'Else
        '    CODE = "SELECT ISNULL(""U_Estable"",'0') AS Establecimiento, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = 79"
        '    _code = "SELECT ISNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = 79"
        'End If
        If TipoDocumento = "REE" Or TipoDocumento = "REA" Or TipoDocumento = "RER" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo DocEntry del UDO retencion " + DocEntryUdoRet.ToString, "ManejoDeDocumentos")
            DocEntryUdoRet = oFuncionesB1.getRSvalue("select T1.""DocEntry"" FROM ""OPCH"" T0 INNER JOIN ""@TM_LE_RETCH"" T1 ON T0.""U_TM_CRNUM""= T1.""DocEntry"" WHERE T0.""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocEntry", "")
            Utilitario.Util_Log.Escribir_Log("Obteniendo DocEntry del UDO retencion : " + DocEntryUdoRet.ToString, "ManejoDeDocumentos")
        End If
        'Dim Est As String = oFuncionesB1.getRSvalue(CODE, "Establecimiento")
        'Dim PuntoEmi As String = oFuncionesB1.getRSvalue(_code, "PuntoEmision")
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

        '_code = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "Code", "")
        'Sql = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + oSerie
        'Dim LQELEC As String = oFuncionesB1.getRSvalue(Sql, "Code", "")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    CODE = "SELECT ""Code"" FROM ""@TM_DOC_ELEC"" WHERE ""U_TM_DocEntry"" = " + DocEntryDoc.ToString
        'Else
        '    CODE = "SELECT Code FROM ""@TM_DOC_ELEC"" WHERE U_TM_DocEntry = " + DocEntryDoc.ToString
        'End If
        '_code = oFuncionesB1.getRSvalue(CODE, "Code", "")
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
    Public Function GrabaDatosAutorizacion_UDORT_TM(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim _code As String = ""
        _code = oFuncionesB1.getRSvalue("select T1.""DocEntry"" from ""OPCH"" T0 inner join ""@TM_LE_RETCH"" T1 on T0.""U_TM_CRNUM""=T1.""DocEntry"" where T0.""DocEntry""= '" + DocEntryDoc.ToString() + "' ", "DocEntry", "")

        Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion GrabaDatosAutorizacion_UDORT_TM (antes del try)", "ManejoDeDocumentos")
        If _code <> "" Then

            Try
                Dim RetVal As Long
                Dim ErrCode As Long
                Dim ErrMsg As String

                Dim ActualizaSecuenc As Boolean = True
                Dim oGeneralService As SAPbobsCOM.GeneralService
                Dim oGeneralData As SAPbobsCOM.GeneralData
                Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

                Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
                Dim oUserTable As SAPbobsCOM.UserTable = Nothing
                GC.Collect()
                oUserObjectMD = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                Dim sCmp As SAPbobsCOM.CompanyService
                sCmp = rCompany.GetCompanyService
                Utilitario.Util_Log.Escribir_Log("antes del if ", "ManejoDeDocumentos")
                If oUserObjectMD.GetByKey("TM_LE_RETCH") Then ' PREGUNTO SI ES UN UDO, YA QUE ALGUNOS CLIENTES NO TIENEN REGISTRADO EL UDO
                    'GuardaLOG(Tipotabla, DocEntry, "'EXX_DOCUM_LEG_INTER' es un UDO: ", Transaccion, TipoLog)
                    oGeneralService = sCmp.GetGeneralService("TM_LE_RETCH")
                    Utilitario.Util_Log.Escribir_Log("TM_LE_RETCH oGeneralService", "ManejoDeDocumentos")
                    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    oGeneralParams.SetProperty("Code", _code)
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Obteniendo Registro a actualizar en 'TM_LE_RETCH' por el Code: " + _code.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                    Utilitario.Util_Log.Escribir_Log("oGeneralData error", "ManejoDeDocumentos")
                    oGeneralData.SetProperty("U_TM_CASRI", _NumAutorizacion.ToString)
                    oGeneralService.Update(oGeneralData)
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "# RT AutorizacionSri actualizado en 'TM_LE_RETCH' por el Code: " + _code.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    Return True
                Else

                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Obteniendo Registro a actualizar en 'TM_LE_RETCH' por el Code: " + _code.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    oUserTable = rCompany.UserTables.Item("TM_LE_RETCH")
                    oUserTable.GetByKey(_code)
                    oUserTable.UserFields.Fields.Item("U_TM_CASRI").Value = _NumAutorizacion.ToString
                    RetVal = oUserTable.Update()
                    If RetVal <> 0 Then
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "No se actualizaron datos en la tabla TM_LE_RETCH: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Else
                        oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_LE_RETCH: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End If
                End If
                Return True
            Catch ex As Exception
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + "Error al actualizar el numero de autorizacion en TM_LE_RETCH" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia" + ex.Message.ToString(), Transaccion, TipoLog)
                ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
                Return False
            End Try
        Else
            oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "TM_LE_RETCH No se encontro el code: " + _code.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
        End If

#Disable Warning BC42353 ' La función 'GrabaDatosAutorizacion_UDORT_TM' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    End Function
#Enable Warning BC42353 ' La función 'GrabaDatosAutorizacion_UDORT_TM' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    Public Function GrabaDatosAutorizacion_HESION_GUIA(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        'Dim DocEntryUdoRet As String = ""
        Dim DocNum As String = ""
        Dim _DocNum As String = ""
        'Dim listaTran As New List(Of Integer)

        If TipoDocumento = "TRE" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")
            If _tipoManejo = "A" Then
                DocNum = oFuncionesB1.getRSvalue("SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Transferencias""='Y' and ""U_HBT_NumeroDesde3"" = '" + DocNum.ToString() + "' ", "Code", "")
            Else
                DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""OWTR"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Transferencias""='Y' and ""U_HBT_NumeroDesde3"" = '" + DocNum.ToString() + "' ", "Code", "")
            End If
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")

            'ElseIf TipoDocumento = "TLE" Then
            '    Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")
            '    If _tipoManejo = "A" Then
            '        DocNum = oFuncionesB1.getRSvalue("SELECT ""DocNum"" FROM ""OWTQ"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            '        CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Transferencias""='Y' and ""U_HBT_NumeroDesde3"" = '" + DocNum.ToString() + "' ", "Code", "")
            '    Else
            '        DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""OWTQ"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            '        CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Transferencias""='Y' and ""U_HBT_NumeroDesde3"" = '" + DocNum.ToString() + "' ", "Code", "")
            '    End If
            '    Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")

        ElseIf TipoDocumento = "GRE" Then
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION: " + CODE.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum " + DocEntryDoc.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("TipoDocumento " + TipoDocumento.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("_tipoManejo " + _tipoManejo.ToString, "ManejoDeDocumentos")
            If _tipoManejo = "A" Then
                Try
                    DocNum = oFuncionesB1.getRSvalue("SELECT ""DocNum"" FROM ""ODLN"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Entregas""='Y' and ""U_HBT_NumeroDesde2"" = '" + DocNum.ToString() + "' ", "Code", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            Else
                Try
                    DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""ODLN"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
                Try
                    CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Entregas""='Y' and ""U_HBT_NumeroDesde2"" = '" + DocNum.ToString() + "' ", "Code", "")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
                End Try
            End If
            'DocNum = getRSvalueGRHEISON(_DocNum, "DocNum", "")
            Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum query" + DocNum.ToString, "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Obteniendo Code query: " + CODE.ToString, "ManejoDeDocumentos")

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

                oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @HBT_GUIAREMISION: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompany.UserTables.Item("HBT_GUIAREMISION")
                oUserTable.GetByKey(CODE)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If
                oUserTable.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString
                oUserTable.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString

                RetVal = oUserTable.Update()
                If RetVal <> 0 Then
                    'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                Else
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_DOC_ELEC: " + CODE.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                End If
                Return True
            Else
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla HBT_GUIAREMISION: " + CODE.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try



        Return result
    End Function

    Public Function GrabaDatosAutorizacion_HESION_FACTURAGUIA(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        'Dim DocEntryUdoRet As String = ""
        Dim DocNum As String = ""
        Dim _DocNum As String = ""
        'Dim listaTran As New List(Of Integer)


        Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION GR: " + CODE.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum GR" + DocEntryDoc.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("TipoDocumento GR" + TipoDocumento.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("_tipoManejo GR" + _tipoManejo.ToString, "ManejoDeDocumentos")
        If _tipoManejo = "A" Then
            Try
                DocNum = oFuncionesB1.getRSvalue("SELECT ""DocNum"" FROM ""OINV"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            Try
                CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Facturas""='Y' and ""U_HBT_NumeroDesde1"" = '" + DocNum.ToString() + "' ", "Code", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
        Else
            Try
                DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""OINV"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            Try
                CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Facturas""='Y' and ""U_HBT_NumeroDesde1"" = '" + DocNum.ToString() + "' ", "Code", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
        End If

        Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum query" + DocNum.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("Obteniendo Code query: " + CODE.ToString, "ManejoDeDocumentos")




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

                oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @HBT_GUIAREMISION: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompany.UserTables.Item("HBT_GUIAREMISION")
                oUserTable.GetByKey(CODE)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If
                oUserTable.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString
                oUserTable.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString

                RetVal = oUserTable.Update()
                If RetVal <> 0 Then

                    rCompany.GetLastError(ErrCode, ErrMsg)

                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                Else
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla TM_DOC_ELEC: " + CODE.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    Dim oFacturaGuia As SAPbobsCOM.StockTransfer = Nothing
                    Dim docentryFG As Integer
                    Dim resultado As Integer = -1


                    Dim recordset As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select distinct U_HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OINV ON T1.U_HBT_NumeroDesde1=OINV.""DocNum"" where OINV.""DocEntry"" =" + DocEntryDoc.ToString)

                    If recordset.RecordCount > 1 Then

                        While (recordset.EoF = False)
                            docentryFG = CInt(recordset.Fields.Item("U_HBT_DocEntry").Value)
                            If DocEntryDoc <> docentryFG Then

                                oFacturaGuia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                oFacturaGuia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                                oFacturaGuia.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

                                If oFacturaGuia.GetByKey(docentryFG) Then

                                    oFacturaGuia.UserFields.Fields.Item("U_GR_CLAVE").Value = _ClaveAcceso.ToString()
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_NUM_AUTO").Value = _NumAutorizacion.ToString()
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_FECHA_AUT").Value = _FechaAutorizacion
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_OBSERVACION").Value = _Observacion.ToString
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)

                                    Try
                                        resultado = oFacturaGuia.Update()
                                    Catch ex As Exception
                                        result = False
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("Error al actualizar Factura " + docentryFG.ToString() + " : " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Factura no actualizada: " + docentryFG.ToString() + " error: " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End Try

                                    If resultado = 0 Then
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("Factura: " + docentryFG.ToString() + " actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Transferencia actualizada: " + docentryFG.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                        result = True

                                    End If

                                End If
                            End If
                            recordset.MoveNext()
                        End While

                    End If

                End If
                Return True
            Else
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla HBT_GUIAREMISION: " + CODE.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try



        Return result
    End Function

    Public Function GrabaDatosAutorizacion_HESION_SALIDAGUIA(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim CODE As String = ""
        Dim _code As String = ""
        'Dim DocEntryUdoRet As String = ""
        Dim DocNum As String = ""
        Dim _DocNum As String = ""
        'Dim listaTran As New List(Of Integer)


        Utilitario.Util_Log.Escribir_Log("Obteniendo Code de la tabla HBT_GUIAREMISION SM GR: " + CODE.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum SM GR" + DocEntryDoc.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("TipoDocumento SM GR" + TipoDocumento.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("_tipoManejo SM GR" + _tipoManejo.ToString, "ManejoDeDocumentos")
        If _tipoManejo = "A" Then
            Try
                DocNum = oFuncionesB1.getRSvalue("SELECT ""DocNum"" FROM ""OIGE"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA SM ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            Try
                CODE = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Salidas""='Y' and ""U_HBT_NumeroDesde4"" = '" + DocNum.ToString() + "' ", "Code", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA SM ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
        Else
            Try
                DocNum = getRSvalueGRHEISON("SELECT ""DocNum"" FROM ""OIGE"" WHERE ""DocEntry"" = '" + DocEntryDoc.ToString() + "' ", "DocNum", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA SM ERROR DocNum: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
            Try
                CODE = getRSvalueGRHEISON("SELECT ""Code"" FROM ""@HBT_GUIAREMISION"" WHERE ""U_HBT_Salidas""='Y' and ""U_HBT_NumeroDesde4"" = '" + DocNum.ToString() + "' ", "Code", "")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GrabaDatosAutorizacion_HESION_GUIA SM ERROR CODE: " + ex.Message.ToString, "ManejoDeDocumentos")
            End Try
        End If

        Utilitario.Util_Log.Escribir_Log("ObteniendoDocNum query" + DocNum.ToString, "ManejoDeDocumentos")
        Utilitario.Util_Log.Escribir_Log("Obteniendo Code query: " + CODE.ToString, "ManejoDeDocumentos")




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

                oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Obteniendo Informacion de la tabla @HBT_GUIAREMISION: ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                oUserTable = rCompany.UserTables.Item("HBT_GUIAREMISION")
                oUserTable.GetByKey(CODE)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("Actualizando datos de autorizacion en la tabla Control de Doc. Electrónicos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                End If
                oUserTable.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString
                oUserTable.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString

                RetVal = oUserTable.Update()
                If RetVal <> 0 Then

                    rCompany.GetLastError(ErrCode, ErrMsg)

                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Datos no actualizados en la tabla TM_DOC_ELEC: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                    If _tipoManejo = "A" Then
                        rsboApp.SetStatusBarMessage("Datos no actualizados en la tabla HBT_GUIAREMISION: " + ErrCode.ToString + " - " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    End If
                    Return False
                Else
                    oFuncionesAddon.GuardaLOG(TipoDocumento, DocEntryDoc, "Datos actualizados en la tabla HBT_GUIAREMISION: " + CODE.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                    Dim oFacturaGuia As SAPbobsCOM.StockTransfer = Nothing
                    Dim docentryFG As Integer
                    Dim resultado As Integer = -1


                    Dim recordset As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select distinct U_HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OIGE ON T1.U_HBT_NumeroDesde4=OIGE.""DocNum"" where OIGE.""DocEntry"" =" + DocEntryDoc.ToString)

                    If recordset.RecordCount > 1 Then

                        While (recordset.EoF = False)
                            docentryFG = CInt(recordset.Fields.Item("U_HBT_DocEntry").Value)
                            If DocEntryDoc <> docentryFG Then

                                oFacturaGuia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                                oFacturaGuia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenExit
                                oFacturaGuia.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

                                If oFacturaGuia.GetByKey(docentryFG) Then

                                    oFacturaGuia.UserFields.Fields.Item("U_GR_CLAVE").Value = _ClaveAcceso.ToString()
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_NUM_AUTO").Value = _NumAutorizacion.ToString()
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_FECHA_AUT").Value = _FechaAutorizacion
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_OBSERVACION").Value = _Observacion.ToString
                                    oFacturaGuia.UserFields.Fields.Item("U_GR_ESTADO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)

                                    Try
                                        resultado = oFacturaGuia.Update()
                                    Catch ex As Exception
                                        result = False
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("Error al actualizar Salida de Mercancia " + docentryFG.ToString() + " : " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Salida de Mercancias no actualizada: " + docentryFG.ToString() + " error: " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End Try

                                    If resultado = 0 Then
                                        If _tipoManejo = "A" Then
                                            rsboApp.SetStatusBarMessage("Salida de Mercancias: " + docentryFG.ToString() + " actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Salida de Mercancias actualizada: " + docentryFG.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                        result = True

                                    End If

                                End If
                            End If
                            recordset.MoveNext()
                        End While

                    End If

                End If
                Return True
            Else
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Code del documento creado en la Tabla HBT_GUIAREMISION: " + CODE.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return False
            End If
        Catch ex As Exception
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("SAED - Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            Utilitario.Util_Log.Escribir_Log("Error al actualizar datos de autorizacion en la tabla HBT_GUIAREMISION: " + ex.Message.ToString, "ManejoDeDocumentos")
            Return False
        End Try



        Return result
    End Function


    Public Function GrabaDatosAutorizacion_HESION_GUIA_TRANSFERENCIAS(TipoDocumento As String, DocEntryDoc As String) As Boolean
        Dim result As Boolean = False
        Dim oTransferencia As SAPbobsCOM.StockTransfer = Nothing
        Dim docentry As Integer
        Dim resultado As Integer = -1
        If TipoDocumento = "TRE" Then

            Dim recordset As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select distinct U_HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OWTR ON T1.U_HBT_NumeroDesde3=OWTR.DocNum where owtr.DocEntry =" + DocEntryDoc.ToString)
            If recordset.RecordCount > 1 Then
                While (recordset.EoF = False)
                    docentry = CInt(recordset.Fields.Item("U_HBT_DocEntry").Value)
                    If DocEntryDoc <> docentry Then
                        oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                        If oTransferencia.GetByKey(docentry) Then
                            oTransferencia.UserFields.Fields.Item("U_HBT_IdEnProveedor").Value = _NumAutorizacion.ToString()
                            oTransferencia.UserFields.Fields.Item("U_HBT_ClaveAcceso").Value = _ClaveAcceso.ToString()
                            oTransferencia.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = _NumAutorizacion.ToString()
                            'oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = Date.Now
                            oTransferencia.UserFields.Fields.Item("U_FECHA_AUT_FACT").Value = _FechaAutorizacion
                            oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = _Observacion.ToString
                            oTransferencia.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = IIf(_EstadoAutorizacion = "-1", "0", _EstadoAutorizacion)
                            If Not String.IsNullOrEmpty(_ClaveAcceso) Then
                                oTransferencia.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = _ClaveAcceso.ToString()
                            End If
                            Try
                                resultado = oTransferencia.Update()
                            Catch ex As Exception
                                result = False
                                If _tipoManejo = "A" Then
                                    rsboApp.SetStatusBarMessage("Error al actualizar transferencia " + docentry.ToString() + " : " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Transferencia no actualizada: " + docentry.ToString() + " error: " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                End If
                            End Try
                            If resultado = 0 Then
                                If _tipoManejo = "A" Then
                                    rsboApp.SetStatusBarMessage("Transferencia: " + docentry.ToString() + " actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddon.GuardaLOG(TipoDocumento.ToString, DocEntryDoc.ToString, "Transferencia actualizada: " + docentry.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                End If
                                result = True

                            End If

                        End If
                    End If
                    recordset.MoveNext()
                End While
            End If
        End If
        Return result
    End Function
    Public Function ReenvioMail(sCorreo As String, sClaveAcceso As String) As Boolean

        Try
            Dim url As String = ""
            'url = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionReenvioMail")
            url = Functions.VariablesGlobales._wsReenvioMail
            If url = "" Then
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                Exit Function
            End If

            Dim ws As New Entidades.wsEDoc_ReEnvioMail.WSEDOC_ENVIARMAIL
            ws.Url = url

            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""
            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                If Not Proxy_puerto = "" Then
                    proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                Else
                    proxyobject = New System.Net.WebProxy(Proxy_IP)
                End If
                cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                proxyobject.Credentials = cred

                ws.Proxy = proxyobject
                ws.Credentials = cred

            End If
            'If Functions.VariablesGlobales._vgHttps = "Y" Then
            '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
            'End If
            SetProtocolosdeSeguridad()
            Dim mensa As String = ""
            ws.EnviarCorreoDocumentoEmitido(sClaveAcceso, sCorreo, mensa, True, True)
            Utilitario.Util_Log.Escribir_Log("mail enviado: " + mensa, "ManejoDeDocumentos")
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("GS, Error:" + ex.Message.ToString())
            Return False
        End Try

        Return True
    End Function


    Public Function ConsultaPDF(sClaveAcceso As String) As Boolean
        Try
            Dim TipoWebServices As String = "LOCAL"
            'TipoWebServices = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
            TipoWebServices = Functions.VariablesGlobales._TipoWS
            Dim url As String = ""
            'url = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionConsulta")
            url = Functions.VariablesGlobales._wsConsultaEmision
            If url = "" Then
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                Exit Function
            End If

            Utilitario.Util_Log.Escribir_Log("VisualizaPDF_Bytes :  " + sClaveAcceso, "ManejoDeDocumentos")

            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            Dim ruta As String = ""
            Dim ws As Object

            Utilitario.Util_Log.Escribir_Log("VER PDF WS : " + TipoWebServices, "ManejoDeDocumentos")

            If TipoWebServices = "LOCAL" Then
                ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
            ElseIf TipoWebServices = "NUBE" Then
                ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                'ElseIf TipoWebServices = "NUBE_4_1" Then
                '    ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA
            ElseIf TipoWebServices = "NUBE_4_1" Then
                ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

            End If

            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                If Not Proxy_puerto = "" Then
                    proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                Else
                    proxyobject = New System.Net.WebProxy(Proxy_IP)
                End If
                cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                proxyobject.Credentials = cred


                ws.Proxy = proxyobject

                ws.Credentials = cred

            End If

            ws.Url = url

            Utilitario.Util_Log.Escribir_Log("VER PDF URL : " + url, "ManejoDeDocumentos")

            Dim VisualizaPDF_Bytes As String = "N"
            'VisualizaPDF_Bytes = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "VisualizaPDF_Bytes")
            VisualizaPDF_Bytes = Functions.VariablesGlobales._VisualizaPDFByte

            Utilitario.Util_Log.Escribir_Log("VisualizaPDF_Bytes :  " + VisualizaPDF_Bytes, "ManejoDeDocumentos")
            If VisualizaPDF_Bytes = "Y" Then

                'BYTES
                Dim filepath As String = Path.GetTempPath()
                filepath += sClaveAcceso + ".pdf"
                'If IsNotFileInUse(filepath) Then
                '    ' El archivo no está en uso, puedes proceder con tu lógica aquí
                '    Console.WriteLine("El archivo no está en uso.")
                'Else
                '    Console.WriteLine("El archivo está en uso por otro proceso.")
                'End If
                ' SI NO EXISTE EN LA CARPETA TEMPORAL, LO CONSULTO AL WS
                If Not File.Exists(filepath) Then
                    rsboApp.SetStatusBarMessage("Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Dim FS As FileStream = Nothing
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    Dim dbbyte As Byte() = Nothing
                    mensaje = ""
                    If TipoWebServices = "LOCAL" Then
                        dbbyte = ws.ConsultarDocumento(sClaveAcceso, "PDF")
                    ElseIf TipoWebServices = "NUBE" Then
                        dbbyte = ws.ConsultarDocumento(sClaveAcceso, "PDF")
                    ElseIf TipoWebServices = "NUBE_4_1" Then
                        dbbyte = ws.ConsultarDocumento(sClaveAcceso, "PDF", mensaje)
                    End If
                    If dbbyte Is Nothing Then
                        rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Else

                        FS = New FileStream(filepath, System.IO.FileMode.Create)
                        FS.Write(dbbyte, 0, dbbyte.Length)
                        FS.Close()
                    End If
                End If
                ''BYTES

                Dim Proc As New Process()
                Proc.StartInfo.FileName = filepath
                Proc.Start()
                Proc.Dispose()

                'funciona pero no coloca en medio
                'Dim pdf As PdfDocument = New PdfDocument()
                'pdf.LoadFromFile(filepath)

                'Dim font As PdfTrueTypeFont = New PdfTrueTypeFont(New System.Drawing.Font("Arial", 10.0F, FontStyle.Regular), True)

                'Dim pageNumber As PdfPageNumberField = New PdfPageNumberField()
                'Dim pageCount As PdfPageCountField = New PdfPageCountField()
                'Dim compositeField As PdfCompositeField = New PdfCompositeField(font, PdfBrushes.Black, "Documento generado y autorizado desde SAP B1")

                'compositeField.StringFormat = New PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle)

                'For i As Integer = 0 To pdf.Pages.Count - 1
                '    compositeField.Draw(pdf.Pages(i).Canvas, pdf.Pages(i).Size.Width / 2 - 20, pdf.Pages(i).Size.Height - pdf.PageSettings.Margins.Bottom
                '                        )
                'Next

                'pdf.SaveToFile(filepath)


            Else

                '' RUTA
                rsboApp.SetStatusBarMessage("Consultando url: " + url.ToString() + " Clave Acceso: " + sClaveAcceso, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                'If Functions.VariablesGlobales._vgHttps = "Y" Then
                '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                'End If
                SetProtocolosdeSeguridad()
                mensaje = ""
                If TipoWebServices = "LOCAL" Then
                    ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "PDF")
                ElseIf TipoWebServices = "NUBE" Then
                    ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "PDF")
                ElseIf TipoWebServices = "NUBE_4_1" Then
                    ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "PDF", mensaje)
                End If

                If ruta Is Nothing Then
                    rsboApp.SetStatusBarMessage("El ws NO devolvio la ruta " + mensaje.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                If ruta.Contains("win-u8ppvmocuel") Then
                    ruta = ruta.Replace("win-u8ppvmocuel", "gurusoft-lab.com")
                End If
                'ruta
                rsboApp.SetStatusBarMessage("Abriendo la siguiente ruta: " + ruta.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                Dim Proc As New Process()
                Proc.StartInfo.FileName = ruta
                Proc.Start()
                Proc.Dispose()
                '' END RUTA


            End If

            rsboApp.SetStatusBarMessage("PDF Abierto! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try

    End Function



    Public Function ConsultaXML(sClaveAcceso As String) As Boolean
        Try
            Dim TipoWebServices As String = "LOCAL"
            'TipoWebServices = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
            TipoWebServices = Functions.VariablesGlobales._TipoWS
            Dim url As String = ""
            'url = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionConsulta")
            url = Functions.VariablesGlobales._wsConsultaEmision
            If url = "" Then
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                Exit Function
            End If

            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            Dim ruta As String = ""
            Dim ws As Object
            If TipoWebServices = "LOCAL" Then
                ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
            ElseIf TipoWebServices = "NUBE" Then
                ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
            ElseIf TipoWebServices = "NUBE_4_1" Then
                ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA
            End If

            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                If Not Proxy_puerto = "" Then
                    proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                Else
                    proxyobject = New System.Net.WebProxy(Proxy_IP)
                End If
                cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                proxyobject.Credentials = cred

#Disable Warning BC42104 ' La variable 'ws' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                ws.Proxy = proxyobject
#Enable Warning BC42104 ' La variable 'ws' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                ws.Credentials = cred

            End If

            ws.Url = url


            Dim VisualizaPDF_Bytes As String = "N"
            'VisualizaPDF_Bytes = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "VisualizaPDF_Bytes")
            VisualizaPDF_Bytes = Functions.VariablesGlobales._VisualizaPDFByte
            If VisualizaPDF_Bytes = "Y" Then

                'BYTES
                Dim filepath As String = Path.GetTempPath()
                filepath += sClaveAcceso + ".xml"
                ' SI NO EXISTE EN LA CARPETA TEMPORAL, LO CONSULTO AL WS
                If Not File.Exists(filepath) Then
                    rsboApp.SetStatusBarMessage("Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Dim FS As FileStream = Nothing
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    SetProtocolosdeSeguridad()
                    'Dim dbbyte As Byte() = ws.ConsultarDocumento(sClaveAcceso, "XML")
                    Dim dbbyte As Byte() = Nothing
                    mensaje = ""
                    If TipoWebServices = "LOCAL" Then
                        dbbyte = ws.ConsultarDocumento(sClaveAcceso, "XML")
                    ElseIf TipoWebServices = "NUBE" Then
                        dbbyte = ws.ConsultarDocumento(sClaveAcceso, "XML")
                    ElseIf TipoWebServices = "NUBE_4_1" Then
                        dbbyte = ws.ConsultarDocumento(sClaveAcceso, "XML", mensaje)
                    End If
                    If dbbyte Is Nothing Then
                        rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Else
                        FS = New FileStream(filepath, System.IO.FileMode.Create)
                        FS.Write(dbbyte, 0, dbbyte.Length)
                        FS.Close()
                    End If
                End If
                'BYTES
                Dim Proc As New Process()
                Proc.StartInfo.FileName = filepath
                Proc.Start()
                Proc.Dispose()

            Else

                '' RUTA
                rsboApp.SetStatusBarMessage("Consultando url: " + url.ToString() + " Clave Acceso: " + sClaveAcceso, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                'If Functions.VariablesGlobales._vgHttps = "Y" Then
                '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                'End If
                SetProtocolosdeSeguridad()
                'ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "XML")
                mensaje = ""
                If TipoWebServices = "LOCAL" Then
                    ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "XML")
                ElseIf TipoWebServices = "NUBE" Then
                    ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "XML")
                ElseIf TipoWebServices = "NUBE_4_1" Then
                    ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "XML", mensaje)
                End If
                If ruta Is Nothing Then
                    rsboApp.SetStatusBarMessage("El ws NO devolvio la ruta " + mensaje.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                If ruta.Contains("win-u8ppvmocuel") Then
                    ruta = ruta.Replace("win-u8ppvmocuel", "gurusoft-lab.com")
                End If
                'ruta
                rsboApp.SetStatusBarMessage("Abriendo la siguiente ruta: " + ruta.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                Dim Proc As New Process()
                Proc.StartInfo.FileName = ruta
                Proc.Start()
                Proc.Dispose()
                '' END RUTA

            End If
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("XML Abierto! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If

        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
#Disable Warning BC42353 ' La función 'ConsultaXML' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    End Function


#Region "Recorre Errores"

    Public Function recorreErrorNotaCredito(ByVal respuesta As Entidades.wsEDoc_NotaDeCredito.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_NotaDeCredito.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_NotaDeCredito.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeCredito.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_NotaDeCredito.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeCredito.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_NotaDeCredito.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorFactura(ByVal respuesta As Entidades.wsEDoc_Factura.RespuestaEDOC, ByVal codigoDocumento As String) As String

        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_Factura.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_Factura.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Factura.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_Factura.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Factura.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_Factura.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorNotaDebito(ByVal respuesta As Entidades.wsEDoc_NotaDeDebito.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_NotaDeDebito.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_NotaDeDebito.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeDebito.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_NotaDeDebito.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeDebito.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_NotaDeDebito.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorGuiaRemision(ByVal respuesta As Entidades.wsEDoc_GuiaRemision.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_GuiaRemision.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_GuiaRemision.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_GuiaRemision.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_GuiaRemision.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_GuiaRemision.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_GuiaRemision.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorRetencion(ByVal respuesta As Entidades.wsEDoc_Retencion.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_Retencion.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_Retencion.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Retencion.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_Retencion.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Retencion.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_Retencion.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function

    Public Function recorreErrorNotaCredito_LOCAL(ByVal respuesta As Entidades.wsEDoc_NotaDeCredito_LOCAL.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_NotaDeCredito_LOCAL.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_NotaDeCredito_LOCAL.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeCredito_LOCAL.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_NotaDeCredito_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeCredito_LOCAL.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_NotaDeCredito_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorFactura_LOCAL(ByVal respuesta As Entidades.wsEDoc_Factura_LOCAL.RespuestaEDOC, ByVal codigoDocumento As String) As String

        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_Factura_LOCAL.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_Factura_LOCAL.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Factura_LOCAL.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_Factura_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Factura_LOCAL.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_Factura_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorNotaDebito_LOCAL(ByVal respuesta As Entidades.wsEDoc_NotaDeDebito_LOCAL.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_NotaDeDebito_LOCAL.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_NotaDeDebito_LOCAL.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeDebito_LOCAL.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_NotaDeDebito_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeDebito_LOCAL.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_NotaDeDebito_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorGuiaRemision_LOCAL(ByVal respuesta As Entidades.wsEDoc_GuiaRemision_LOCAL.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_GuiaRemision_LOCAL.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_GuiaRemision_LOCAL.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_GuiaRemision_LOCAL.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_GuiaRemision_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_GuiaRemision_LOCAL.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_GuiaRemision_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorRetencion_LOCAL(ByVal respuesta As Entidades.wsEDoc_Retencion_LOCAL.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_Retencion_LOCAL.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_Retencion_LOCAL.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Retencion_LOCAL.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_Retencion_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Retencion_LOCAL.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_Retencion_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function

    Public Function recorreErrorNotaCredito_NUBE41(ByVal respuesta As Entidades.wsEDoc_NotaDeCredito41.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_NotaDeCredito41.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_NotaDeCredito41.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeCredito41.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_NotaDeCredito41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeCredito41.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_NotaDeCredito41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorFactura_NUBE41(ByVal respuesta As Entidades.wsEDoc_Factura41.RespuestaEDOC, ByVal codigoDocumento As String) As String

        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_Factura41.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_Factura41.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Factura41.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_Factura41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Factura41.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_Factura41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorNotaDebito_NUBE41(ByVal respuesta As Entidades.wsEDoc_NotaDeDebito41.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_NotaDeDebito41.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_NotaDeDebito41.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeDebito41.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_NotaDeDebito41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_NotaDeDebito41.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_NotaDeDebito41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorGuiaRemision_NUBE41(ByVal respuesta As Entidades.wsEDoc_GuiaRemision41.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_GuiaRemision41.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_GuiaRemision41.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_GuiaRemision41.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_GuiaRemision41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_GuiaRemision41.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_GuiaRemision41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorRetencion_NUBE41(ByVal respuesta As Entidades.wsEDoc_Retencion41.RespuestaEDOC, ByVal codigoDocumento As String) As String
        Dim mensaje As String = ""
        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_Retencion41.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_Retencion41.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = 7 Or respuesta.Estado = "7" Or respuesta.Estado = 5 Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Retencion41.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_Retencion41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_Retencion41.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_Retencion41.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function

    Public Function recorreErrorLiquidacionCompra_LOCAL(ByVal respuesta As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.RespuestaEDOC, ByVal codigoDocumento As String) As String

        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = "7" Or respuesta.Estado = "5" Or respuesta.Estado = 5 Or respuesta.Estado = 7 Then
                'If respuesta.Estado = "ERROR EN RECEPION" Or respuesta.Estado = "EN PROCESO SRI" Or respuesta.Estado = "7" Or respuesta.Estado = "5" Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.WSEDOC_LIQUIDACIONES_COMPRA_LOCAL.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
    Public Function recorreErrorLiquidacionCompra(ByVal respuesta As Entidades.wsEDoc_LiquidacionCompra.RespuestaEDOC, ByVal codigoDocumento As String) As String

        Dim estado As String = ""

        If respuesta.Estado = "AUTORIZADO" Or respuesta.Estado = "2" Then
            For Each item As Entidades.wsEDoc_LiquidacionCompra.EAutorizacion In respuesta.autorizaciones
                If item.estado = "AUTORIZADO" And item.mensajes.Count <= 2 Then
                    mensaje = "Estado: AUTORIZADO, # Autorizacion: " + respuesta.autorizaciones(0).numeroAutorizacion
                    mensaje = mensaje & " - Ambiente: " + item.ambiente
                    Exit For
                End If
            Next
            mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
            estado = respuesta.Estado
            Return mensaje
        Else
            Dim Errores As String = ""
            Dim men As New Entidades.wsEDoc_LiquidacionCompra.EMensaje
            estado = respuesta.Estado
            men.identificador = "E00"
            If respuesta.Estado = "7" Or respuesta.Estado = "5" Or respuesta.Estado = 5 Or respuesta.Estado = 7 Then
                'If respuesta.Estado = "ERROR EN RECEPION" Or respuesta.Estado = "EN PROCESO SRI" Or respuesta.Estado = "7" Or respuesta.Estado = "5" Then

                mensaje += "Estado: EN ESPERA DE AUTORIZACIÓN DEL SRI," + "VERIFICAR LA AUTORIZACIÓN DE ESTE COMPROBANTE ELECTRÓNICO DURANTE EL DÍA"
                mensaje = mensaje & " - NÚMERO DEL DOCUMENTO: " + codigoDocumento
                Return mensaje
            Else
                If respuesta.autorizaciones.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_LiquidacionCompra.EAutorizacion In respuesta.autorizaciones
                        For Each item1 As Entidades.wsEDoc_LiquidacionCompra.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += item1.mensaje1 + ",  "
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                mensaje = item1.identificador & ": " & item1.mensaje1 & " - " & item1.informacionAdicional
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                        Errores += " * AMBIENTE :" + item.ambiente
                    Next
                End If
                If respuesta.Comprobantes.Count <> 0 Then
                    For Each item As Entidades.wsEDoc_LiquidacionCompra.EComprobante In respuesta.Comprobantes
                        For Each item1 As Entidades.wsEDoc_LiquidacionCompra.EMensaje In item.mensajes
                            If item1.tipo = "ERROR" Then
                                Errores += " * " + item1.mensaje1 + ": " + item1.informacionAdicional + " - " + "Codigo Error: " + item1.identificador + " - " + item1.informacionAdicional
                                men.tipo = item1.tipo
                                men.informacionAdicional = item1.informacionAdicional
                                men.identificador = item1.identificador
                                men.mensaje1 = mensaje
                                Exit For
                            End If
                        Next
                    Next

                End If
                mensaje += " Estado: NO AUTORIZADO " + Errores
                mensaje = mensaje & ",NÚMERO DEL DOCUMENTO: " + codigoDocumento
            End If
            Return mensaje
        End If

    End Function
#End Region

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

    Public Function EjecutarQuery(Query As String) As DataSet

        Dim ds As New DataSet
        Try
            Using Cn As SqlConnection = GetSqlConnectionBase()
                Using cm As New SqlCommand()
                    cm.CommandText = Query
                    cm.CommandType = CommandType.Text
                    cm.Connection = Cn
                    Cn.Open()

                    Dim da As New SqlDataAdapter
                    da.SelectCommand = cm
                    da.Fill(ds)

                End Using
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
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

    Public Sub ValidarConexionADO()
        Dim cnBaseSAP As New SqlConnection
        Try

            cnBaseSAP = GetSqlConnectionBase()
            cnBaseSAP.Open()

            If cnBaseSAP.State = ConnectionState.Open Then
                cnBaseSAP.Close()
            End If

        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            rsboApp.SetStatusBarMessage("Cadena Conexión " + cnBaseSAP.ConnectionString.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub


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

    Private Function ValidaClave(ByVal clave As String, ByVal cod_comp As String, ByVal serie As String, ByVal secuencia As String) As String

        If clave.Length = 49 Then

            Dim secufull As String = secuencia
            Dim numero_documento As String = ""

            If Not secufull.Length = 9 Then

                secufull = secufull.PadLeft(9, "0")

            End If

            'construyo el numero del documento
            numero_documento = serie & secufull

            'verifico si el numero del documento y el cod del comprobante existe en la clave de acceso

            If clave.Substring(8, 2) = cod_comp And clave.Substring(24, 15) = numero_documento Then

                Return clave

            Else

                Return ""

            End If


        Else
            Return ""
        End If

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

    Public Function AnularDocumento(ByVal AidForm As String) As Boolean
        'Dim _AidForm As String = AidForm
        'If AidForm = "133" Or AidForm = "60091" Or AidForm = "60090" Then

        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

        '    Dim _oDocumentoSAP As SAPbobsCOM.
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        'ElseIf AidForm = "65303" Then
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
        '    oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
        'End If
        'oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

        'If BusinessObjectInfo.FormTypeEx = "133" Or BusinessObjectInfo.FormTypeEx = "60090" Then  ' FACTURA DE CLIENTE - FACTURA DEUDOR + PAGO
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oTipoTabla = "FCE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "60091" Then ' FACTURA DE RESERVA
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oTipoTabla = "FRE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "65303" Then ' NOTA DE DEBITO
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oTipoTabla = "NDE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "65300" Then ''FACTURA DE ANTICIPO DE CLIENTES
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
        '    oTipoTabla = "FAE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "179" Then 'NOTA DE CREDITO DE CLIENTES
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
        '    oTipoTabla = "NCE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "140" Then 'GUIA DE REMISION - ENTREGA
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
        '    oTipoTabla = "GRE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "940" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
        '    oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
        '    oTipoTabla = "TRE"

        'ElseIf BusinessObjectInfo.FormTypeEx = "1250000940" Then 'GUIA DE REMISION - SOLICITUD TRANSLADO                                            
        '    oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
        '    oTipoTabla = "TLE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "141" Then  'FACTURA DE PROVEEDOR/RETENCION                             
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "REE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "65306" Then  'NOTA DE DEBITO PROVEEDOR/RETENCION                             
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "RDM"
        '    Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
        'ElseIf BusinessObjectInfo.FormTypeEx = "65301" Then  'FACTURA DE ANTICIPO DE PROVEEDOR/RETENCION                             
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "REA"
        '    Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
        'ElseIf BusinessObjectInfo.FormTypeEx = "60092" Then  'FACTURA DE RESERVA PROVEEDOR/RETENCION                           
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "RER"
        '    Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")



        Return True
    End Function

    Private Function ImprmirDOcAut(numaut As String) As Boolean

        Try
            Dim TipoWebServicesImpDocAut As String = Functions.VariablesGlobales._TipoWS
            Utilitario.Util_Log.Escribir_Log("funcion imprimir tipo Web Service: " + TipoWebServicesImpDocAut.ToString, "frmImpresionPorBloque")

            Dim url As String = Functions.VariablesGlobales._wsConsultaEmision

            If url = "" Then
                rsboApp.SetStatusBarMessage(Functions.VariablesGlobales._vgNombreAddOn + " No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                mensajeDocAut = "No existe informacion del Web Service, revisar Parametrización"
                Return False
            End If

            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmImpresionPorBloque")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            Dim ruta As String = ""
            Dim ws As Object

            ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmImpresionPorBloque")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmImpresionPorBloque")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmImpresionPorBloque")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmImpresionPorBloque")

                If Not Proxy_puerto = "" Then
                    proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                Else
                    proxyobject = New System.Net.WebProxy(Proxy_IP)
                End If
                cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                proxyobject.Credentials = cred
                ws.Proxy = proxyobject
                ws.Credentials = cred

            End If

            ws.Url = url

            Dim path = System.IO.Path.GetTempPath + numaut + ".pdf"

            SetProtocolosdeSeguridad()
            Dim FS As FileStream = Nothing
            Dim dbbyte As Byte() = Nothing
            If Not File.Exists(path) Then
                dbbyte = ws.ConsultarDocumento(numaut, "PDF", mensaje)


                If dbbyte Is Nothing Then
                    mensajeDocAut = " - Arreglo de bytes vacío...! "
                    Return False
                Else
                    FS = New FileStream(path, System.IO.FileMode.Create)
                    FS.Write(dbbyte, 0, dbbyte.Length)
                    FS.Close()

                End If
            Else
                dbbyte = File.ReadAllBytes(path)
            End If


            If File.Exists(path) Then
                'Dim Esperas As Integer = 0
                'Using p As New Process
                '    p.StartInfo.FileName = path
                '    p.StartInfo.Verb = "Print"

                '    p.Start()

                '    Threading.Thread.Sleep(3000) ' tiempo X para que el programa cliente se active he imprima

                '    p.CloseMainWindow() ' Cierre ventana cliente
                '    ' si la ventana sigue abierta, se encicla hasta cerrarla.
                '    While Not p.HasExited
                '        Threading.Thread.Sleep(1000)
                '        Esperas += 1
                '        p.CloseMainWindow()
                '    End While
                'End Using
                Dim Stream = New MemoryStream(dbbyte)
                Using doc As New PdfDocument
                    'doc.LoadFromXPS(arrb)
                    doc.LoadFromStream(Stream)
                    doc.PrintSettings.PrintController = New StandardPrintController
                    doc.Print()
                    Utilitario.Util_Log.Escribir_Log("Impresion Realizada: " + doc.ToString(), "ImpresionAutomatica")
                End Using

            End If
            Return True
        Catch ex As Exception
            mensajeDocAut = ex.Message.ToString
            Return False
        End Try


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

    Function IsNotFileInUse(filePath As String) As Boolean
        Dim fileInUse As Boolean = False
        Dim fileStream As FileStream = Nothing

        Try
            ' Intenta abrir el archivo en modo exclusivo (FileShare.None)
            fileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None)
        Catch ex As IOException
            ' Si hay una excepción, el archivo está en uso por otro proceso
            fileInUse = True
        Finally
            ' Asegúrate de cerrar el stream después de usarlo
            If fileStream IsNot Nothing Then
                fileStream.Close()
            End If
        End Try

        ' Devuelve True si el archivo no está en uso, False si está en uso
        Return Not fileInUse
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

    Public Sub ActivarTLS()
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or 768 Or 3072
    End Sub
End Class

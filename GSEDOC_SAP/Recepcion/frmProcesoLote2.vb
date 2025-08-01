Imports Entidades
Imports System.Threading
Imports System.Globalization

Public Class frmProcesoLote2
    Public oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    '
    Private listaDocumentoPorUsuario As List(Of Entidades.DocumentoTipo)
    Private listaFCs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
    Private listaREs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
    Private listaNCs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
    Private listaNDs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaDebito)
    Private listaGRs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTGuiaRemision)
    Public listaDetalleArtiulos As New List(Of Entidades.DetalleArticulo)
    Private oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura
    Private oNotaDeCredito As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito
    Private oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion

    Dim _RUC As String = ""
    Dim sCardCode As String = ""
    Dim _WS_Recepcion As String = ""
    Dim _WS_RecepcionCambiarEstado As String = ""
    Dim _WS_RecepcionClave As String = ""
    Dim _WS_Recepcion_NombreCampoOC As String = ""
    Dim mensaje As String = ""
    Dim i As Integer
    Dim ofila As Integer
    Dim residuo As Integer = 0
    Dim RegistrosXPaginas As Integer = 0
    Dim TotalDocs As Integer = 0
    '
    Dim odt As SAPbouiCOM.DataTable
    Dim oUserDataSource As SAPbouiCOM.UserDataSource

    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition

    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential

    'FAMC carga de estados
    Dim _WS_RecepcionCargaEstados As String = ""

    Dim DocEntryFacturaRecibida_UDO As String = 0
    ''
    Dim _oDocumento As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion



    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormularioPL2()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        rsboApp.SetStatusBarMessage(NombreAddon + " - Consultando, Espere Por favor!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        If RecorreFormulario(rsboApp, "frmProcesoLote2") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmProcesoLote2.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmProcesoLote2").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmProcesoLote2")

            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO

            oForm.Freeze(True)

            ' CHOOSE FROM LIST
            oCFLs = oForm.ChooseFromLists
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            'oCFLCreationParams.ObjectType = "Exx_DEPOTRANS"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            ' END CHOOSE FROM LIST

            Dim txtRuc As SAPbouiCOM.EditText
            txtRuc = oForm.Items.Item("txtRuc").Specific
            oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtRuc.DataBind.SetBound(True, "", "EditDS")
            'txtRuc.ChooseFromListUID = "CFL1"
            'txtRuc.ChooseFromListAlias = "CardType"
            'txtRuc.SetChooseFromList("MyCFL", "AcctCode")

            Dim cmbTipo As SAPbouiCOM.ComboBox
            cmbTipo = oForm.Items.Item("cbxTipo").Specific
            ''cmbTipo.ValidValues.Add("0", "Todos")
            cmbTipo.ValidValues.Add("01", "Factura")
            cmbTipo.ValidValues.Add("04", "Nota de Crédito")
            ''cmbTipo.ValidValues.Add("05", "Nota de Débito")
            ''cmbTipo.ValidValues.Add("06", "Guía de Remisión")

            cmbTipo.ValidValues.Add("07", "Comp. de Retención")
            cmbTipo.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim lnkPr As SAPbouiCOM.LinkedButton
            lnkPr = oForm.Items.Item("lnkPr").Specific
            lnkPr.LinkedObjectType = 2
            lnkPr.Item.LinkTo = "txtRuc"

            Dim focus As SAPbouiCOM.EditText
            focus = oForm.Items.Item("focus").Specific
            focus.Item.Visible = False




            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Tipo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Fecha", SAPbouiCOM.BoFieldsType.ft_Date, 100)

            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("FechaAutorizacion", SAPbouiCOM.BoFieldsType.ft_Date, 100)

            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Folio", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("RUC", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("RazonSocial", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Valor", SAPbouiCOM.BoFieldsType.ft_Price, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("ClaveAcceso", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("NumAutorizacion", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)


            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("NumDocRetener", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)

            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("OC", SAPbouiCOM.BoFieldsType.ft_Integer, 20)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Mapeado", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Borrador", SAPbouiCOM.BoFieldsType.ft_Integer, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Sucursal", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)


            cargarDocumentos()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub
    Private Sub campoRET()
        Dim cbxTipo As SAPbouiCOM.ComboBox
        cbxTipo = oForm.Items.Item("cbxTipo").Specific
        Dim btnCRT As SAPbouiCOM.Button
        btnCRT = oForm.Items.Item("btnCreaeRT").Specific
        If cbxTipo.Value = "07" Then
            btnCRT.Item.Visible = True
        Else
            btnCRT.Item.Visible = False
        End If


    End Sub
    Private Sub CrearPRLote()

        Try
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oDatable As SAPbouiCOM.DataTable
            'pintar filas
            Dim gcss As SAPbouiCOM.CommonSetting
            gcss = oGrid.CommonSetting
            ' oDatable = oForm.DataSources.DataTables.Item("dtDocs")
            oDatable = oGrid.DataTable
            Dim x As Integer, y As Integer
            Dim nombre_estado As String = ""
            Dim ss_tipotabla As String = ""
            Dim identificador As Integer = 0
            Dim indexgrid As Integer = 0
            For x = 0 To oDatable.Rows.Count - 1
                nombre_estado = oDatable.GetValue("Folio", x)
                'If nombre_estado = Estados_docenviados.EN_PROCESO_SRI Or nombre_estado = Estados_docenviados.ERROR_EN_RECEPCION Then
                '    ss_tipotabla = obtenerTipoTabla(oDatable.GetValue("ObjType", x), oDatable.GetValue("DocSubType", x))
                '    identificador = CInt(oDatable.GetValue("DocEntry", x))
                For y = 1 To oGrid.Rows.Count
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    If indexgrid = x Then
                        'oForm.Freeze(True)
                        'gcss.GetCellBackColor(y, 3)
                        '255, 255, 0  255000
                        nombre_estado = oDatable.GetValue("Folio", x)
                        gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                        Exit For
                    End If

                Next

                ''rsboApp.MessageBox("ok")
                ''oForm.Freeze(False)
                'Try
                '    ' oManejoDocumentos.ProcesaEnvioDocumento(identificador, ss_tipotabla, True)
                rsboApp.MessageBox(oDatable.GetValue("Folio", x))
                If Guarda_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO) Then  ' GUARDO EL DOCUMENTO RECIBIDO EN EL UDO FACTURA RECIBIDA
                    rsboApp.MessageBox("POSI POSI")
                    ' Exitoso = CrearFacturaPreliminarRelacionada(sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                End If
                '    gcss.SetRowBackColor(y, 255000)
                'Catch ex As Exception
                '    gcss.SetRowBackColor(y, RGB(255, 0, 0))
                'End Try


                'End If
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next

            rsboApp.StatusBar.SetText("(SAED) El estado de los documentos han sido actualizados correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)



        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Ocurrio un error al llamar la funcion ConsultarEstados " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    Public Shared Function formatDecimal(ByVal numero As String) As Decimal

        Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
        Dim result As Double = 0
        Try
            If numero = "" Then
                numero = "0"
            End If
            If numero IsNot Nothing Then
                If Not numero.Contains(",") Then
                    result = Double.Parse(numero, CultureInfo.InvariantCulture)
                Else
                    result = Convert.ToDouble(numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
                    'result = Double.Parse((numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString())), CultureInfo.InvariantCulture)
                End If
            End If
        Catch e As Exception
            Try
                'result = Convert.ToDouble(numero)
                result = Double.Parse(numero, CultureInfo.InvariantCulture)
            Catch
                Try
                    'result = Convert.ToDouble(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."))
                    result = Double.Parse(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."), CultureInfo.InvariantCulture)
                Catch
                    Throw New Exception("Wrong string-to-double format")
                End Try
            End Try
        End Try
        Return result

    End Function

    Public Function Guarda_DocumentoRecibido_Factura(ByRef DocEntryFacturaRecibida_UDO As String) As Boolean

        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote")
        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        Dim _ClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim numDoc As String = ""
        Dim resultsFC As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
        resultsFC = listaFCs.FindAll(Function(column) column.ClaveAcceso = _ClaveAcceso)
        'Dim ofactura As Object
        For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In resultsFC
            Try
                oFuncionesAddon.GuardaLOG("FACTURA", _ClaveAcceso, "Creando registro de Factura Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

                Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + sRUC + "'", "CardCode", "")
                numDoc = oFac.Establecimiento + "-" + oFac.PuntoEmision + "-" + oFac.Secuencial

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                'oGeneralData.SetProperty("Code", conta)

                'oGeneralData.SetProperty("U_Ruta_pdf", rutaFC.ToString())
                oGeneralData.SetProperty("U_RUC", oFac.Ruc.ToString())
                oGeneralData.SetProperty("U_Nombre", oFac.RazonSocial.ToString())
                oGeneralData.SetProperty("U_CardCode", sCardCode.ToString())
                oGeneralData.SetProperty("U_Mapeado", "PL")
                oGeneralData.SetProperty("U_ClaAcc", oFac.ClaveAcceso.ToString())
                oGeneralData.SetProperty("U_NumAut", oFac.AutorizacionSRI.ToString())
                oGeneralData.SetProperty("U_FecAut", oFac.FechaAutorizacion.ToString())
                oGeneralData.SetProperty("U_NumDoc", numDoc.ToString())
                oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString()) 'sDocEntryPreliminar
                oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(formatDecimal(oFac.TotalSinImpuesto.ToString())))
                oGeneralData.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(oFac.TotalDescuento.ToString())))
                oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(formatDecimal(oFac.ImporteTotal)))
                oGeneralData.SetProperty("U_rGast", "0")
                oGeneralData.SetProperty("U_rImp", "0")
                oGeneralData.SetProperty("U_rTotal", "0")
                oGeneralData.SetProperty("U_IdGS", oFac.IdFactura.ToString())
                oGeneralData.SetProperty("U_Sincro", 0)
                oGeneralData.SetProperty("U_Estado", "docPreliminar")

                For Each impFAC As Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto In oFac.ENTFacturaImpuesto
                    Dim BaseImponibleIVA As Decimal = 0
                    Dim BaseImponible0 As Decimal = 0
                    Dim BaseImponibleNoObjeto As Decimal = 0
                    Dim BaseImponibleExento As Decimal = 0
                    Dim Iva As Decimal = 0
                    Dim ICE As Decimal = 0
                    Dim BaseImponibleICE As Decimal = 0
                    If impFAC.Codigo = 2 Then
                        If impFAC.CodigoPorcentaje = 2 Or impFAC.CodigoPorcentaje = 3 Then
                            BaseImponibleIVA += impFAC.BaseImponible
                            Iva += impFAC.Valor
                        ElseIf impFAC.CodigoPorcentaje = 0 Then
                            BaseImponible0 += impFAC.BaseImponible
                        ElseIf impFAC.CodigoPorcentaje = 6 Then
                            BaseImponibleNoObjeto += impFAC.BaseImponible
                        ElseIf impFAC.CodigoPorcentaje = 7 Then
                            BaseImponibleExento += impFAC.BaseImponible
                        End If
                    ElseIf impFAC.Codigo = 3 Then
                        ICE += impFAC.Valor
                    End If

                    oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(formatDecimal(BaseImponibleIVA.ToString())))
                    oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(formatDecimal(BaseImponible0.ToString())))
                    oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(formatDecimal(BaseImponibleNoObjeto.ToString())))
                    oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(formatDecimal(BaseImponibleExento.ToString())))
                    oGeneralData.SetProperty("U_ICE", Convert.ToDouble(formatDecimal(ICE.ToString())))
                    oGeneralData.SetProperty("U_IVA", Convert.ToDouble(formatDecimal(Iva.ToString())))
                    oGeneralData.SetProperty("U_rTades", Convert.ToDouble(formatDecimal(impFAC.BaseImponible.ToString())))
                    oGeneralData.SetProperty("U_rPDesc", Convert.ToDouble(formatDecimal(impFAC.Tarifa.ToString())))
                    oGeneralData.SetProperty("U_rDesc", Convert.ToDouble(formatDecimal(impFAC.Valor.ToString())))
                Next
                oGeneralData.SetProperty("U_Tipo", "Factura Tipo Servicio PL")

                oChildren = oGeneralData.Child("GS0_FVR")
                For Each detalleFac As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFac.ENTDetalleFactura
                    oChild = oChildren.Add
                    oChild.SetProperty("U_CodPrin", Left(detalleFac.CodigoPrincipal.ToString(), 99))
                    oChild.SetProperty("U_CodAuxi", Left(detalleFac.CodigoAuxiliar.ToString(), 99))
                    oChild.SetProperty("U_CodSAP", "Servicio PL")
                    oChild.SetProperty("U_Descripc", Left(detalleFac.Descripcion.ToString(), 99))
                    oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(detalleFac.Cantidad.ToString())))
                    oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(detalleFac.PrecioUnitario.ToString())))
                    oChild.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(detalleFac.Descuento.ToString())))
                    oChild.SetProperty("U_Total", Convert.ToDouble(formatDecimal(detalleFac.PrecioTotalSinImpuesto.ToString())))
                Next

                'oChildren = oGeneralData.Child("GS1_FVR")
                'oChild = oChildren.Add
                'oChild.SetProperty("U_DocEntr", Integer.Parse(odt.GetValue(0, i).ToString()))
                'oChild.SetProperty("U_LineNu", Integer.Parse(odt.GetValue(1, i).ToString()))
                'oChild.SetProperty("U_ItemCode", odt.GetValue(2, i).ToString())
                'oChild.SetProperty("U_Descripc", odt.GetValue(3, i).ToString())
                'oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(odt.GetValue(4, i).ToString())))
                'oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(odt.GetValue(5, i).ToString())))
                'oChild.SetProperty("U_DiscPr", Convert.ToDouble(formatDecimal(odt.GetValue(6, i).ToString())))
                'oChild.SetProperty("U_TaxCode", odt.GetValue(7, i).ToString())
                'oChild.SetProperty("U_lTotal", Convert.ToDouble(formatDecimal(odt.GetValue(8, i).ToString())))
                'oChild.SetProperty("U_ObjType", odt.GetValue(9, i).ToString())


                oGeneralParams = oGeneralService.Add(oGeneralData)
                DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
                oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Se creo registro de Factura Recibida UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return True
                '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
                'sDocEntry = oGeneralParams.GetProperty("Code")
            Catch ex As Exception
                oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Ocurrior un error al crear registro de Factura Recibida UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Factura Recibida en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        Next


    End Function


    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            'Dim typeEx, idForm As String
            'typeEx = oFuncionesB1.FormularioActivo(idForm)
            If pVal.FormTypeEx = "frmProcesoLote2" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal

                        'If oCFLEvento.BeforeAction = False Then
                        'ChooseFromList(pVal, FormUID)
                        'End If

                        If oCFLEvento.BeforeAction = False Then
                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty

                            If Not oDataTable Is Nothing Then
                                val = oDataTable.GetValue(0, 0)
                                val1 = oDataTable.GetValue(1, 0)
                                Try

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                                    oUserDataSource.ValueEx = val

                                Catch ex As Exception
                                End Try

                                Try
                                    Dim txtRaz As SAPbouiCOM.EditText
                                    txtRaz = oForm.Items.Item("txtRaz").Specific
                                    txtRaz.Value = val1
                                Catch ex As Exception
                                End Try
                            Else
                                Dim txtRaz As SAPbouiCOM.EditText
                                txtRaz = oForm.Items.Item("txtRaz").Specific
                                txtRaz.Value = ""
                            End If

                        Else
                            'BubbleEvent = False
                        End If


                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If Not pVal.Before_Action Then
                            Try
                                oForm = rsboApp.Forms.Item("frmProcesoLote2")
                                Dim cbxTipo As SAPbouiCOM.ComboBox
                                cbxTipo = oForm.Items.Item("cbxTipo").Specific

                                oCons = oCFL.GetConditions()

                                Dim lbSocio As SAPbouiCOM.StaticText
                                lbSocio = oForm.Items.Item("lbSocio").Specific



                                If oCons.Count > 0 Then 'If there are already user conditions.
                                    If cbxTipo.Value = "07" Then ' SI ES 07, SIGNIFICA QUE ES RETENCION, POR ENDE PAGO RECIBIDO DE CLIENTE
                                        oCons.Item(oCons.Count - 1).CondVal = "C"
                                        lbSocio.Caption = "Cliente :"

                                    Else
                                        oCons.Item(oCons.Count - 1).CondVal = "S"
                                        lbSocio.Caption = "Proveedor :"

                                    End If
                                End If

                                oCFL.SetConditions(oCons)
                                Dim txtRuc As SAPbouiCOM.EditText
                                txtRuc = oForm.Items.Item("txtRuc").Specific
                                txtRuc.ChooseFromListUID = "CFL1"
                                txtRuc.ChooseFromListAlias = "CardType"

                            Catch ex As Exception

                            End Try


                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                'Case "btnCreaeRT"
                                '    CrearPRLote()

                                '    ' Dim resp As Integer = rsboApp.MessageBox("Desea actualizar los Estados ?", 1, "SI", "NO")



                                Case "obtnBuscar"
                                    'rsboApp.SetStatusBarMessage("Estará Disponible en una proxima version!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    cargarDocumentos()
                                Case "obtnCrear"
                                    CrearPRLote()

                                Case "oGrid"
                                    ofila = pVal.Row
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLote2")
                                    Dim oGrid As SAPbouiCOM.Grid = oFor.Items.Item("oGrid").Specific
                                    oGrid.Rows.SelectedRows.Add(ofila)
                                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                        ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                                        ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                                    Next
                                Case "btnAnt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLote2")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) - 1, TotalDocs)
                                Case "btnSig"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLote2")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) + 1, TotalDocs)
                                Case "btnPri"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLote2")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), 1, TotalDocs)
                                Case "btnUlt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLote2")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbNP.Caption), TotalDocs)
                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pVal.ColUID = "RUC" And pVal.BeforeAction = True Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                            Dim sCardCode As String = ""
                            Dim sLicTradNum As String = ""
                            Dim tipoDocumento As String = ""
                            ofila = pVal.Row
                            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                            oGrid.Rows.SelectedRows.Add(ofila)
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                'ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder))
                                sLicTradNum = oDataTable.GetValue(4, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                tipoDocumento = oDataTable.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                ' sLicTradNum = odt.GetValue("RUC", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                            Next
                            Try
                                Dim QueryExisteProveedor As String = ""
                                Dim QueryExisteCliente As String = ""
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                    Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then

                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sLicTradNum + "'"
                                        QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sLicTradNum + "'"
                                    Else
                                        QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sLicTradNum + "'"
                                        QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sLicTradNum + "'"
                                    End If

                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""U_DOCUMENTO"" = '" + sLicTradNum + "'"
                                        QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""U_DOCUMENTO"" = '" + sLicTradNum + "'"
                                    Else
                                        QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND U_DOCUMENTO = '" + sLicTradNum + "'"
                                        QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND U_DOCUMENTO = '" + sLicTradNum + "'"
                                    End If
                                End If

                                sCardCode = ""

                                If tipoDocumento = "Retención de Cliente" Then
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                                Else
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                                End If

                            Catch ex As Exception
                            End Try
                            If sCardCode <> "" Then
                                oForm.Items.Item("txtOpch").Specific.Value = sCardCode
                                oForm.Items.Item("lnkOpch").Click()
                                'rsboApp.SendKeys("^+U")
                            Else
                                Dim respuesta = rsboApp.MessageBox(NombreAddon + " - No existe el proveedor con el RUC/Cedula seleccionado: " + sLicTradNum + ", Desea Crearlo ?", 1, "OK", "Cancelar")
                                If respuesta = 1 Then
                                    rsboApp.ActivateMenuItem("2561")
                                    oForm = rsboApp.Forms.GetFormByTypeAndCount(134, -1)
                                    oForm.Select()
                                    rsboApp.ActivateMenuItem("1282") 'NUEVO
                                End If
                            End If
                            BubbleEvent = False
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.BeforeAction = False Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

                            'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                            'oGrid.Rows.SelectedRows.Add(ofila)

                            Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                            Dim QueryExisteProveedor As String = ""
                            Dim QueryExisteCliente As String = ""
                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sRUC + "'"
                                    QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
                                Else
                                    QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sRUC + "'"
                                    QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
                                End If
                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""U_DOCUMENTO"" = '" + sRUC + "'"
                                    QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""U_DOCUMENTO"" = '" + sRUC + "'"
                                Else
                                    QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND U_DOCUMENTO = '" + sRUC + "'"
                                    QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND U_DOCUMENTO = '" + sRUC + "'"
                                End If
                            End If

                            sCardCode = ""
                            Dim tipoDocumento As String = oDataTable.GetValue(0, ofila).ToString()
                            If tipoDocumento = "Retención de Cliente" Then
                                sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                            Else
                                sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                            End If

                            If sCardCode = "" Then
                                Dim respuesta = rsboApp.MessageBox(NombreAddon + " - No existe el Socio de Negocio con el RUC/Cedula seleccionado: " + sRUC + ", Desea Crearlo ?", 1, "OK", "Cancelar")
                                If respuesta = 1 Then
                                    rsboApp.ActivateMenuItem("2561")
                                    oForm = rsboApp.Forms.GetFormByTypeAndCount(134, -1)
                                    oForm.Select()
                                    rsboApp.ActivateMenuItem("1282") 'NUEVO
                                End If
                            Else
                                Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                                Dim sNombre As String = oDataTable.GetValue(5, ofila).ToString()
                                Dim sMapeado As String = oDataTable.GetValue(11, ofila).ToString()
                                Dim iBorrador As Integer = Integer.Parse(oDataTable.GetValue(12, ofila).ToString())

                                rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Documento de " + sNombre + ", por favor espere..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                If oDataTable.GetValue(0, ofila).ToString() = "Factura" Then
                                    Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                    results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                        oFactura = oFac
                                    Next
                                    Dim sQueryIdDocumento As String = ""
                                    Dim idDocumentoRecibido_UDO As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sQueryIdDocumento = "SELECT ""U_SSIDDOCUMENTO"" FROM ""ODRF"" WHERE ""ObjType"" = 18 and ""DocEntry"" = " + iBorrador.ToString()
                                    Else
                                        sQueryIdDocumento = "select U_SSIDDOCUMENTO from ODRF Where ObjType = 18 and DocEntry = " + iBorrador.ToString()
                                    End If
                                    If iBorrador = 0 Then
                                        ofrmDocumento.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oFactura, ofila)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumento.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If

                                ElseIf oDataTable.GetValue(0, ofila).ToString() = "Nota de Crédito" Then
                                    Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                                    results = listaNCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results
                                        oNotaDeCredito = oNC
                                    Next
                                    Dim sQueryIdDocumento As String = ""
                                    Dim idDocumentoRecibido_UDO As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sQueryIdDocumento = "SELECT ""U_SSIDDOCUMENTO"" FROM ""ODRF"" WHERE ""ObjType"" = 19 and ""DocEntry"" = " + iBorrador.ToString()
                                    Else
                                        sQueryIdDocumento = "select U_SSIDDOCUMENTO from ODRF Where ObjType = 19 and DocEntry = " + iBorrador.ToString()
                                    End If
                                    If iBorrador = 0 Then
                                        ofrmDocumentoNC.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oNotaDeCredito, ofila)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumentoNC.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If

                                ElseIf oDataTable.GetValue(0, ofila).ToString() = "Retención de Cliente" Then
                                    Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                                    results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                                        oRetencion = oRE
                                    Next
                                    Dim sQueryIdDocumento As String = ""
                                    Dim idDocumentoRecibido_UDO As String = ""
                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            sQueryIdDocumento += " SELECT B.""U_SSIDDOCUMENTO"""
                                            sQueryIdDocumento += " FROM ""OPDF"" B "
                                            sQueryIdDocumento += " WHERE B.""DocEntry"" = " + iBorrador.ToString()
                                            sQueryIdDocumento += " AND B.""U_SSCREADAR"" = 'SI'"
                                            sQueryIdDocumento += " AND B.""ObjType"" = 24"
                                        Else
                                            sQueryIdDocumento += " SELECT B.U_SSIDDOCUMENTO"
                                            sQueryIdDocumento += " FROM OPDF B "
                                            sQueryIdDocumento += " WHERE B.DocEntry = " + iBorrador.ToString()
                                            sQueryIdDocumento += " AND B.U_SSCREADAR = 'SI'"
                                            sQueryIdDocumento += " AND B.ObjType = 24"
                                        End If
                                    Else
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            sQueryIdDocumento += " SELECT A.""U_SSIDDOCUMENTO"""
                                            sQueryIdDocumento += " FROM ""PDF3"" A INNER JOIN"
                                            sQueryIdDocumento += " ""OPDF"" B ON A.""DocNum"" = B.""DocEntry"" AND A.""U_SSCREADAR"" = 'SI'"
                                            sQueryIdDocumento += " WHERE B.""DocEntry"" = " + iBorrador.ToString()
                                            sQueryIdDocumento += " AND B.""ObjType"" = 24"
                                        Else
                                            sQueryIdDocumento += " SELECT A.U_SSIDDOCUMENTO"
                                            sQueryIdDocumento += " FROM PDF3 A INNER JOIN"
                                            sQueryIdDocumento += " OPDF B ON A.DocNum = B.DocEntry AND A.U_SSCREADAR = 'SI'"
                                            sQueryIdDocumento += " WHERE B.DocEntry = " + iBorrador.ToString()
                                            sQueryIdDocumento += " AND B.ObjType = 24"
                                        End If
                                    End If

                                    If iBorrador = 0 Then
                                        ofrmDocumentoRE.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oRetencion, ofila)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumentoRE.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If
                                End If
                                rsboApp.StatusBar.SetText(NombreAddon + " - Documento de " + sNombre + ", Cargado!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If

                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cargarDocumentos()
        Try

            ' COMENTAMOS ESTA PARTE XQ NO SE PUEDE INGRESAR A SAP TODOS LOS DOCUMENTOS, SOLO FACTURA POR AHORA
            'Dim oUsuarioBL As New BL.UsuarioBL
            '' Obtengo los documentos que tiene asociado el departamento de usario que esta logoneado
            'listaDocumentoPorUsuario = oUsuarioBL.ConsultarDepartamentoUsuario(_SBO_Application.Company.UserName, oCadenaConexion.GetCadena_WebConfigArmada().ConnectionString, _BaseSAP, "R")

            ' AGEGO MANUALMENTE DEBIDO A QUE ESTA COMENTADA LA CONSULTA
            listaDocumentoPorUsuario = New List(Of Entidades.DocumentoTipo)
            listaDocumentoPorUsuario.Add(New Entidades.DocumentoTipo("FC", "Factura", "13", "SI"))
            rsboApp.SetStatusBarMessage(NombreAddon + " - Cargando Documentos Recibidos, Por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
            oForm.Freeze(True)
            If Not listaDocumentoPorUsuario Is Nothing Then

                Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                oDataTable.Rows.Clear()

                If CargarDocumento() Then
                    rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Documentos Recibidos, Listo..!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If

            End If

            'CargaDocumentosFormato()

            oForm.Freeze(False)
        Catch ex As Exception
            'lbError.Visible = True
            'lbError.Text = ex.Message
            'oUtilitario_Email = New Utilitario.UtilManejador_Email("Error: UserControl_RecepcionDocumentos/CargarDocumentoPorUsuario Usuario: " + _SBO_Application.Company.UserName.ToString(), ConfigurationManager.AppSettings("CorreoResponsable"), ex.Message)
            'oUtilitario_Email.Enviar()
            rsboApp.StatusBar.SetText("Error al Cargar los documentos: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Function CargarDocumento() As Boolean
        Try
            'CONSULTO EL RUC DE LA BASE ACTUAL            
            ' _RUC = oFuncionesB1.getRSvalue("SELECT TAXIDNUM FROM OADM", "TAXIDNUM", "")

            ' OBTENGO URL DEL SERVICIO DE RECEPCION
            'RegistrosXPaginas = oFuncionesAddon.getRSvalue("SELECT TOP 1 ""U_Valor"" from ""@GS_CONFD"" where U_Nombre='Registros_por_paginas'", "U_Valor", "")
            Dim nRegistros As String = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Registros_por_paginas")
            If nRegistros = "" Then
                RegistrosXPaginas = 10
            Else
                RegistrosXPaginas = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Registros_por_paginas")
            End If


            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")
            'FAMC cargo los estados parametrizados
            _WS_RecepcionCargaEstados = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Estados_docs")

            If String.IsNullOrEmpty(_WS_RecepcionCargaEstados) Then _WS_RecepcionCargaEstados = "01"
            'RegistrosXPaginas =  oUserTable.UserFields.Fields.Item("U_ws_RP").Value

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
            WS.Url = _WS_Recepcion

            'MANEJO DE PROXY
            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""
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

                WS.Proxy = proxyobject
                WS.Credentials = cred

            End If
            ' END MANEJO DE PROXY

            'rsboApp.StatusBar.SetText("Consultando url: " + _WS_Recepcion.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'rsboApp.StatusBar.SetText("Clave Url: " + _WS_RecepcionClave.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            TotalDocs = 0
            Dim NumeroPaginas As Integer = 0

            '******* PARA FILTRO, A LA ESPERA SE LE AGREGE EL NOMBRE DE PROVEEDOR LIKE AL WS
            Dim cbxTipo As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipo").Specific
            Dim txtRUC As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific
            Dim txtRaz As SAPbouiCOM.EditText = oForm.Items.Item("txtRaz").Specific

            Dim LicTradNum As String = ""

            '******* PARA FILTRO
            Select Case cbxTipo.Value
                Case "01"
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""U_DOCUMENTO"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT U_DOCUMENTO FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    Dim z = WS.ConsultarFactura(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    If Not z Is Nothing Then
                        i = z.Count
                        listaFCs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura))
                        'ORDENA LA LISTA DE FACTURAS POR FECHA DE EMISION
                        'listaFCs = (From M In listaFCs Order By M.FechaEmision Descending Select M).ToList

                        'listaFCs = (From l In listaFCs Order By l.FechaEmision.Month Descending Select l).ToList

                        TotalDocs = listaFCs.Count
                        'CALCULO EL NUMERO DE PAGINAS EN BASE A LA CANTIDAD DE REGISTROS
                        If TotalDocs <= RegistrosXPaginas Then
                            NumeroPaginas = 1
                        Else
                            NumeroPaginas = Int(TotalDocs / RegistrosXPaginas)
                            residuo = (TotalDocs Mod RegistrosXPaginas)
                            If residuo > 0 Then
                                NumeroPaginas += 1
                            End If
                        End If
                        llenarGrid("01", RegistrosXPaginas, NumeroPaginas, 1, TotalDocs)
                        lbInfo.Caption = "Nº Total Facturas Recibidas :" + TotalDocs.ToString()

                    Else
                        lbInfo.Caption = "Nº Total Documentos Recibidos :" + TotalDocs.ToString()
                        Return False
                    End If

                Case "04"
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""U_DOCUMENTO"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT U_DOCUMENTO FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    Dim z = WS.ConsultarNC(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    If Not z Is Nothing Then
                        i = z.Count
                        listaNCs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito))
                        TotalDocs = listaNCs.Count
                        'CALCULO EL NUMERO DE PAGINAS EN BASE A LA CANTIDAD DE REGISTROS
                        If TotalDocs <= RegistrosXPaginas Then
                            NumeroPaginas = 1

                        Else
                            NumeroPaginas = Int(TotalDocs / RegistrosXPaginas)
                            residuo = (TotalDocs Mod RegistrosXPaginas)
                            If residuo > 0 Then
                                NumeroPaginas += 1
                            End If
                        End If
                        llenarGrid("04", RegistrosXPaginas, NumeroPaginas, 1, TotalDocs)
                        lbInfo.Caption = "Nº Total Nota de Crédito Recibidas :" + TotalDocs.ToString()

                    Else
                        lbInfo.Caption = "Nº Total Documentos Recibidos :" + TotalDocs.ToString()
                        Return False
                    End If

                Case "07"
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '    sQuery = "SELECT TOP 1 ""County"" from ""CRD1"" where ""LicTradNum"" = '" + sRUC + "' ORDER BY 1 DESC"
                        'Else
                        '    sQuery = "select top 1 County from CRD1 where LicTradNum ='" + sRUC + "' ORDER BY 1 DESC"
                        'End If

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'C' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""U_DOCUMENTO"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT U_DOCUMENTO FROM OCRD WITH(NOLOCK) where CardType = 'C' AND CardCode = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'C' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific

                    Dim z = WS.ConsultarRetencion(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    If Not z Is Nothing Then
                        i = z.Count
                        listaREs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))
                        TotalDocs = listaREs.Count
                        'CALCULO EL NUMERO DE PAGINAS EN BASE A LA CANTIDAD DE REGISTROS
                        If TotalDocs <= RegistrosXPaginas Then
                            NumeroPaginas = 1

                        Else
                            NumeroPaginas = Int(TotalDocs / RegistrosXPaginas)
                            residuo = (TotalDocs Mod RegistrosXPaginas)
                            If residuo > 0 Then
                                NumeroPaginas += 1
                            End If
                        End If
                        llenarGrid("07", RegistrosXPaginas, NumeroPaginas, 1, TotalDocs)
                        lbInfo.Caption = "Nº Total Retenciones Recibidas :" + TotalDocs.ToString()

                    Else
                        lbInfo.Caption = "Nº Total Documentos Recibidos :" + TotalDocs.ToString()
                        Return False
                    End If

                    'Case "05"
                    '    Dim z = WS.ConsultarND(_WS_RecepcionClave, LicTradNum, "01", mensaje).ToList
                    '    If Not z Is Nothing Then
                    '        i = z.Count
                    '        listaNDs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaDebito))
                    '    Else
                    '        Return False
                    '    End If
                    'Case "06"
                    '    Dim z = WS.ConsultarGR(_WS_RecepcionClave, LicTradNum, "01", mensaje).ToList
                    '    If Not z Is Nothing Then
                    '        i = z.Count
                    '        listaGRs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTGuiaRemision))
                    '    Else
                    '        Return False
                    '    End If
            End Select
            rsboApp.SetStatusBarMessage(NombreAddon + " - Mensaje WS: " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            Return True
        Catch ex As Exception
            'mensaje += " - Catch: " + ex.Message
            ' rsboApp.SetStatusBarMessage("Catch: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            rsboApp.StatusBar.SetText(NombreAddon + "Mensaje: " + mensaje + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub llenarGrid(tipoDoc As String, RegistrosXPaginas As Integer, NumeroPaginas As Integer, PaginaActual As Integer, TotalDocs As Integer)
        Dim i As Integer = 0
        Dim PendienteMapear As Boolean = False
        Dim NumIni As Integer = 0
        Dim NumFin As Integer = 0
        Dim RangoHasta As Integer = 0
        Dim m_oProgBar As SAPbouiCOM.ProgressBar
        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
            oForm.Freeze(True)

            Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
            Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific
            lbNP.Caption = NumeroPaginas.ToString()
            lbPA.Caption = PaginaActual.ToString()
            Dim btnAnt As SAPbouiCOM.Button = oForm.Items.Item("btnAnt").Specific
            Dim btnSig As SAPbouiCOM.Button = oForm.Items.Item("btnSig").Specific
            Dim btnPri As SAPbouiCOM.Button = oForm.Items.Item("btnPri").Specific
            Dim btnUlt As SAPbouiCOM.Button = oForm.Items.Item("btnUlt").Specific



            If PaginaActual = 1 Then
                btnAnt.Item.Enabled = False
                btnPri.Item.Enabled = False
            Else
                btnAnt.Item.Enabled = True
                btnPri.Item.Enabled = True
            End If
            NumFin = PaginaActual * RegistrosXPaginas
            If PaginaActual = NumeroPaginas Then
                NumIni = Integer.Parse((PaginaActual - 1) * RegistrosXPaginas)
                If residuo > 0 Then
                    If TotalDocs <= RegistrosXPaginas Then
                        RegistrosXPaginas = TotalDocs
                    Else
                        RegistrosXPaginas = residuo
                    End If
                Else
                    If TotalDocs <= RegistrosXPaginas Then
                        RegistrosXPaginas = TotalDocs
                    End If
                End If

                btnSig.Item.Enabled = False
                btnUlt.Item.Enabled = False
            Else
                NumIni = NumFin - RegistrosXPaginas + 1
                btnSig.Item.Enabled = True
                btnUlt.Item.Enabled = True
            End If


            m_oProgBar = rsboApp.StatusBar.CreateProgressBar("My Progress Bar", RegistrosXPaginas, False)
            m_oProgBar.Value = 0

            '// Forward\Backward commands
            'Dim iPos As Integer
            Dim sQuery As String = ""
            Dim sSucursal As String = ""

            If tipoDoc = "01" Then
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()

                oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)
                'listaFCs = (From M In listaFCs Order By M.FechaEmision Descending Select M).ToList
                'obtener las fechas 

                For Each oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In listaFCs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    Dim sRUC As String = oFactura.Ruc
                    Dim sQuerySucursal = ""

                    'preguntar si oFactura.fechaautorizacion es mayor 
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        'sucursal
                        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '    sSucursal = oFuncionesB1.getRSvalue("SELECT ""CRD1"".""County"" FROM ""CRD1"" INNER JOIN ""OCRD"" ON ""CRD1"".""CardCode""=""OCRD"".""CardCode"" where ""OCRD"".""LicTradNum"" = '" + sRUC + "'", "County", "")
                        'Else
                        '    sSucursal = oFuncionesB1.getRSvalue("select  CRD1.County from CRD1 WITH(NOLOCK) INNER JOIN OCRD ON CRD1.CardCode=OCRD.CardCode where OCRD.LicTradNum ='" + sRUC + "'", "County", "")
                        'End If
                        '------------
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_FC")
                            sSucursal = sSucursal.Replace("RUC", sRUC.ToString)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        Else
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_FC")
                            sSucursal = sSucursal.Replace("RUC", sRUC.ToString)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        End If

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_NUM_AUTOR"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_NUM_AUTOR = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_NO_AUTORI"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_NO_AUTORI = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_SYP_NroAuto
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Try
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NROAUTO"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NroAuto"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            End Try

                        Else
                            Try
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SYP_NROAUTO = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SYP_NroAuto = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            End Try

                        End If
                    End If


                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, "DocEntry", "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Factura")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, oFactura.FechaEmision)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, oFactura.FechaAutorizacion)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oFactura.Establecimiento + "-" + oFactura.PuntoEmision + "-" + oFactura.Secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oFactura.Ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(oFactura.RazonSocial, 250))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oFactura.ImporteTotal))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oFactura.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oFactura.AutorizacionSRI)


                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)

                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Factura de " + oFactura.RazonSocial
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Next
                i = 0

            ElseIf tipoDoc = "04" Then
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)

                For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In listaNCs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    sQuery = ""
                    Dim sQuerySucursal = ""
                    Dim sRUC As String = oNC.Ruc
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_NC")
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "County", "")
                        Else
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_NC")
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        End If

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_NUM_AUTOR"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_NUM_AUTOR = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_NO_AUTORI"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_NO_AUTORI = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_SYP_NroAuto
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Try
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SYP_NROAUTO"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SYP_NroAuto"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                            End Try

                        Else
                            Try
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SYP_NROAUTO = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SYP_NroAuto = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                            End Try

                        End If
                    End If

                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, "DocEntry", "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Nota de Crédito")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, oNC.FechaEmision)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, oNC.FechaAutorizacion)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oNC.Establecimiento + "-" + oNC.PuntoEmision + "-" + oNC.Secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oNC.Ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(oNC.RazonSocial, 250))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oNC.ValorModificacion))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oNC.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oNC.AutorizacionSRI)
                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Nota de Crédito de " + oNC.RazonSocial

                Next
                i = 0
            ElseIf tipoDoc = "07" Then
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)


                For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In listaREs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    Dim sucursal As String = ""
                    sQuery = ""
                    sSucursal = ""
                    Dim sQuerySucursal = ""
                    Dim sRUC As String = oRE.Ruc
                    Dim IdCol As String = "DocEntry"
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        IdCol = "DocNum"
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_RET")
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "County", "")
                        Else
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_RET")
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        End If

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_CXS_NUM_AUTO_RETE"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_CXS_NUM_AUTO_RETE = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        IdCol = "DocNum"
                        'U_NUM_AUT
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_NUM_AUT"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_NUM_AUT = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_FX_AUTO_RETENCION
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            '
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""OPDF"" where ""U_FX_AUTO_RETENCION"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from OPDF WITH(NOLOCK) where U_FX_AUTO_RETENCION = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    End If

                    Dim numretener As String

                    Dim odetalle() As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion
                    odetalle = oRE.ENTDetalleRetencion
                    numretener = odetalle(0).NumDocRetener

                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, IdCol, "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If


                    'sucursal = oFuncionesB1.getRSvalue(sSucursal, "Sucursal", "")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Retención de Cliente")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, oRE.FechaEmision)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, oRE.FechaAutorizacion)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oRE.Establecimiento + "-" + oRE.PuntoEmision + "-" + oRE.Secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oRE.Ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, oRE.RazonSocial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oRE.TotalRetencion))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oRE.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oRE.AutorizacionSRI)

                    'oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRe", i, IIf(IsNothing(odetalle.NumDocRetener), "", odetalle.NumDocRetener))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRetener", i, IIf(String.IsNullOrEmpty(numretener), "0", numretener))

                    If Not String.IsNullOrEmpty(DocPreliminar) Then

                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Retención Recibida de " + oRE.RazonSocial
                    'rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Next
                i = 0
            End If

            CargaDocumentosFormato(IIf(tipoDoc = "07", "RE", "TD"))
            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + "Error - llenarGrid: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            '// Stop the progress bar
#Disable Warning BC42104 ' La variable 'm_oProgBar' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            m_oProgBar.Stop()
#Enable Warning BC42104 ' La variable 'm_oProgBar' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
            m_oProgBar = Nothing
        End Try
    End Sub

    Private Sub CargaDocumentosFormato(TipoDocumento As String)
        Try
            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            'oGrid.DataTable.ExecuteQuery(sQuery)

            oGrid.Columns.Item(0).Description = "Tipo Documento"
            oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "Fecha"
            oGrid.Columns.Item(1).TitleObject.Caption = "Fecha"
            oGrid.Columns.Item(1).Editable = False


            oGrid.Columns.Item(2).Description = "Fecha"
            oGrid.Columns.Item(2).TitleObject.Caption = "FechaAutorizacion"
            oGrid.Columns.Item(2).Editable = False
            Dim FchAut As String = ""
            FchAut = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "MostrarFechaAutorizacion")
            If FchAut = "Y" Then
                oGrid.Columns.Item(2).Visible = True
            Else
                oGrid.Columns.Item(2).Visible = False
            End If

            oGrid.Columns.Item(3).Description = "Folio"
            oGrid.Columns.Item(3).TitleObject.Caption = "Folio"
            oGrid.Columns.Item(3).Editable = False

            'Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            'oEditTextColumn = oGrid.Columns.Item(2)
            'oEditTextColumn.LinkedObjectType = 13

            oGrid.Columns.Item(4).Description = "RUC"
            oGrid.Columns.Item(4).TitleObject.Caption = "RUC"
            oGrid.Columns.Item(4).Editable = False
            Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            oEditTextColum = oGrid.Columns.Item(4)
            oEditTextColum.LinkedObjectType = 2

            oGrid.Columns.Item(5).Description = "RazonSocial"
            oGrid.Columns.Item(5).TitleObject.Caption = "RazonSocial"
            oGrid.Columns.Item(5).Editable = False

            oGrid.Columns.Item(6).Description = "Valor"
            oGrid.Columns.Item(6).TitleObject.Caption = "Valor"
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).RightJustified = True

            oGrid.Columns.Item(7).Description = "ClaveAcceso"
            oGrid.Columns.Item(7).TitleObject.Caption = "ClavedeAcceso"
            oGrid.Columns.Item(7).Editable = False
            '
            oGrid.Columns.Item(8).Description = "NumAutorizacion"
            oGrid.Columns.Item(8).TitleObject.Caption = "Numero de Autorizacion"
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(8).Visible = False

            oGrid.Columns.Item(9).Description = "NumDocRetener"
            oGrid.Columns.Item(9).TitleObject.Caption = "NumDocRetener"
            oGrid.Columns.Item(9).Editable = False
            If TipoDocumento = "RE" Then '
                oGrid.Columns.Item(9).Visible = True
            Else
                oGrid.Columns.Item(9).Visible = False
            End If

            oGrid.Columns.Item(10).Description = "OC"
            oGrid.Columns.Item(10).TitleObject.Caption = "# OC"
            oGrid.Columns.Item(10).Editable = False
            oGrid.Columns.Item(10).Visible = False

            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oEditTextColumn = oGrid.Columns.Item(10)
            oEditTextColumn.LinkedObjectType = 22

            oGrid.Columns.Item(11).Description = "Mapeado"
            oGrid.Columns.Item(11).TitleObject.Caption = "Mapeado"
            oGrid.Columns.Item(11).Editable = False
            oGrid.Columns.Item(11).ForeColor = ColorTranslator.ToOle(Color.White)
            oGrid.Columns.Item(11).Visible = False

            oGrid.Columns.Item(12).Description = "Borrador"
            oGrid.Columns.Item(12).TitleObject.Caption = "Documento Preliminar"
            oGrid.Columns.Item(12).Editable = False


            Dim NomCamAdi = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Nombre_CA")

            oGrid.Columns.Item(13).Description = "sucursal"
            oGrid.Columns.Item(13).TitleObject.Caption = NomCamAdi '"Sucursal"
            oGrid.Columns.Item(13).Editable = False
            'oGrid.Columns.Item(10).ForeColor = ColorTranslator.ToOle(Color.White)
            oGrid.Columns.Item(13).Visible = True

            Dim oEditTextColumn2 As SAPbouiCOM.EditTextColumn
            oEditTextColumn2 = oGrid.Columns.Item(12)
            If TipoDocumento = "RE" Then ' SI ES RETENCION - EL PAGO RECIBIDO BORRADOR ES OTRA TABLA Y OTRO OBJTYPE
                oEditTextColumn2.LinkedObjectType = 140
            Else
                oEditTextColumn2.LinkedObjectType = 112
            End If

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()

            Try
                For numfila As Integer = 0 To oGrid.Rows.Count - 1
                    Dim valorFila As Integer = oGrid.GetDataTableRowIndex(numfila)
                    If (valorFila <> -1) Then
                        If (oGrid.DataTable.GetValue("Mapeado", valorFila) = "SI") Then
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 11, ColorTranslator.ToOle(Color.LightGreen))
                        Else
                            oGrid.CommonSetting.SetCellBackColor(numfila + 1, 11, ColorTranslator.ToOle(Color.Red))
                        End If
                    End If
                Next
            Catch ex As Exception
            Finally
            End Try
            'campoRET()
            oForm.Freeze(False)
        Catch ex As Exception

        End Try
    End Sub

    Private Function ChooseFromList(ByRef pVal As SAPbouiCOM.ItemEvent, ByVal FormUID As String) As Boolean

        Dim bBubbleEvent As Boolean = True

        If FormUID = "frmProcesoLote2" Then
            oForm = rsboApp.Forms.Item("frmProcesoLote2")

            odt = oFuncionesB1.CargaChooseFromList(pVal, oForm)

            Dim Val, Val1 As String
            Try
                If Not odt Is Nothing Then

                    Val = odt.GetValue(0, 0)
                    Val1 = odt.GetValue(1, 0)

                    Try
                        'Dim focus As SAPbouiCOM.EditText
                        'focus = oForm.Items.Item("focus").Specific
                        'focus.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Try
                            'Dim txtRUC As SAPbouiCOM.EditText
                            'oForm.Items.Item("txtRuc").Specific.value = Val
                            'txtRuc.DataBind.SetBound(True, "", "EditDS")
                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                            oUserDataSource.ValueEx = Val
                        Catch ex As Exception
                        End Try

                        Try
                            Dim txtRaz As SAPbouiCOM.EditText
                            txtRaz = oForm.Items.Item("txtRaz").Specific
                            txtRaz.Value = Val1
                        Catch ex As Exception

                        End Try


                    Catch ex As Exception

                    End Try

                End If
            Catch ex As Exception
                rsboApp.MessageBox("Error: ChooseFromList" & vbCrLf & vbCrLf & ex.Message)
            End Try


        End If
        Return bBubbleEvent

    End Function

    Private Function RecorreFormulario(ByVal oApp As SAPbouiCOM.Application, ByVal Formulario As String) As Boolean
        Try
            For Each oForm In oApp.Forms
                Select Case oForm.UniqueID
                    Case Formulario
                        oForm.Visible = True
                        oForm.Select()
                        Return True
                End Select
            Next

            For Each oForm In oApp.Forms
                If oForm.UniqueID = Formulario Then
                    oForm.Visible = True
                    oForm.Select()
                    ' oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
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

    Public Function MarcarVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String) As Boolean
        Try
            Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
            WS.Url = _WS_RecepcionCambiarEstado

            ' MANEJO PROXY
            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

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

                WS.Proxy = proxyobject
                WS.Credentials = cred
            End If
            ' END  MANEJO PROXY

            If WS.MarcarVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                rsboApp.SetStatusBarMessage("Documento Marcado como Integrado : " + mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Return True
            Else
                rsboApp.SetStatusBarMessage("El documento NO se marco como Integrado : " + mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Return False
            End If
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("El documento NO se marco como Integrado : " + ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Return False
        End Try
    End Function

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

        Try
            If eventInfo.FormUID = "frmProcesoLote2" And eventInfo.ItemUID = "oGrid" Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                If eventInfo.BeforeAction Then
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    ofila = eventInfo.Row
                    Dim oFor As SAPbouiCOM.Form
                    oFor = rsboApp.Forms.Item("frmProcesoLote2")
                    Dim oGrid As SAPbouiCOM.Grid = oFor.Items.Item("oGrid").Specific
                    oGrid.Rows.SelectedRows.Add(ofila)

                    '  ofila = oGrid.GetDataTableRowIndex(eventInfo.Row)
                    ' ofila = eventInfo.Row
                    'oGrid.DataTable = oFor.DataSources.DataTables.Item("dtDocs")
                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                        ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))

                        ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                    Next


                    oMenuItem = rsboApp.Menus.Item("1280")
                    If oMenuItem.SubMenus.Exists("Mapear") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("Mapear"))
                    End If

                    If oMenuItem.SubMenus.Exists("CrearFactura") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("CrearFactura"))
                    End If

                    If oMenuItem.SubMenus.Exists("Marcar") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("Marcar"))
                    End If

                    Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                    Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

                    If oDataTable.GetValue(9, ofila).ToString() = "NO" Then
                        oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "Mapear"
                        oCreationPackage.String = "Mapear Codigos de Proveedor..."
                        oCreationPackage.Enabled = True
                        oCreationPackage.Position = 20
                        oMenuItem = rsboApp.Menus.Item("1280")
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)

                    End If

                    Dim OC As String = oDataTable.GetValue(8, ofila).ToString()
                    Dim Mapeado As String = oDataTable.GetValue(9, ofila).ToString()

                    If OC <> "0" Or Mapeado = "SI" Then
                        oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "CrearFactura"
                        oCreationPackage.String = "Crear Factura Preliminar..."
                        oCreationPackage.Enabled = True
                        oCreationPackage.Position = 21
                        oMenuItem = rsboApp.Menus.Item("1280")
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    End If
                    ' If OC > 0 Then
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "Marcar"
                    oCreationPackage.String = "Marcar Documento como Integrado..."
                    oCreationPackage.Enabled = True
                    oCreationPackage.Position = 21
                    oMenuItem = rsboApp.Menus.Item("1280")
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                    'End If
                Else
                    oMenuItem = rsboApp.Menus.Item("1280")
                    If oMenuItem.SubMenus.Exists("Mapear") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("Mapear"))
                    End If
                    If oMenuItem.SubMenus.Exists("CrearFactura") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("CrearFactura"))
                    End If
                    If oMenuItem.SubMenus.Exists("Marcar") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("Marcar"))
                    End If
                End If
            End If
        Catch ex As Exception
            rsboApp.MessageBox("Error: " & ex.Message)
        End Try

    End Sub

    Private Sub rSboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            If Not pVal.BeforeAction Then
                If pVal.MenuUID = "Mapear" Then
                    Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    If typeEx = "frmProcesoLote2" Then
                        If ofila > 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()

                            sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + sRUC + "'", "CardCode", "")
                            If String.IsNullOrEmpty(sCardCode) Then
                                Dim respuesta = rsboApp.MessageBox(NombreAddon + " - No existe el proveedor, Desea Crearlo ?", 1, "OK", "Cancelar")
                                If respuesta = 1 Then
                                    rsboApp.ActivateMenuItem("2561")
                                    oForm = rsboApp.Forms.GetFormByTypeAndCount(134, -1)
                                    oForm.Select()
                                    rsboApp.ActivateMenuItem("1282") 'NUEVO
                                End If
                            Else ' SI EXISTE EL PROVEEDOR
                                ' OBTNEGO EL OBJETO QUE LE DI CLICK DERECHO
                                Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                                Dim sNombre As String = oDataTable.GetValue(5, ofila).ToString()
                                If oDataTable.GetValue(0, ofila).ToString() = "Factura" Then
                                    Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                    results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    listaDetalleArtiulos.Clear()
                                    If results.Count > 0 Then
                                        For Each oFC As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFC.ENTDetalleFactura
                                                listaDetalleArtiulos.Add(New Entidades.DetalleArticulo(sClaveAcceso, oDetalle.CodigoPrincipal, _
                                                                                                  oDetalle.CodigoAuxiliar, oDetalle.Descripcion, oDetalle.Cantidad, oDetalle.PrecioUnitario, oDetalle.Descuento, oDetalle.PrecioTotalSinImpuesto))
                                            Next
                                        Next
                                    End If
                                    ofrmMapeo.CargaFormularioMapeo(sRUC, sCardCode, sNombre, listaDetalleArtiulos, ofila, "FV")
                                End If

                            End If
                        Else
                            rsboApp.MessageBox(NombreAddon + " - Por favor dar click en filas que tengan información..")
                        End If
                    End If
                ElseIf pVal.MenuUID = "Marcar" Then
                    Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    If typeEx = "frmProcesoLote2" Then
                        If ofila >= 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLote2")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                            Dim tipoDocumento As String = oDataTable.GetValue(0, ofila).ToString()
                            If tipoDocumento = "Factura" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If MarcarVisto(oFac.IdFactura, 1, mensaje) Then
                                            oDataTable.Rows.Remove(ofila)
                                            CargaDocumentosFormato("FE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                Next
                            ElseIf tipoDocumento = "Nota de Crédito" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                                results = listaNCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If MarcarVisto(oNC.IdNotaCredito, 3, mensaje) Then
                                            oDataTable.Rows.Remove(ofila)
                                            CargaDocumentosFormato("NE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                Next
                            ElseIf tipoDocumento = "Retención de Cliente" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                                results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If MarcarVisto(oRE.IdRetencion, 2, mensaje) Then
                                            oDataTable.Rows.Remove(ofila)
                                            CargaDocumentosFormato("RE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                Next
                            End If

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmProcesoLote2")

        End Try

    End Sub


#Region "Funciones Comentadas"

    Private Function CrearFacturaPreliminar(ByVal PO_DocEntry As Integer, oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura) As Boolean

        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String

        Dim iTotalPO_Line As Integer
        Dim iTotalFrgChg_Line As Integer

        Dim baseGRPO As SAPbobsCOM.Documents
        baseGRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

        Try

            If baseGRPO.GetByKey(PO_DocEntry) = True Then
                GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                GRPO.CardCode = baseGRPO.CardCode
                GRPO.DocDate = Today.Date
                GRPO.DocDueDate = Today.Date

                iTotalPO_Line = baseGRPO.Lines.Count
                iTotalFrgChg_Line = baseGRPO.Expenses.Count

                ' DATOS DE AUTORIZACION
                GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = oFactura.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SER_EST").Value = oFactura.Establecimiento
                GRPO.UserFields.Fields.Item("U_SER_PE").Value = oFactura.PuntoEmision

                'CREADAPORGSEDOC
                GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"

                GRPO.FolioNumber = oFactura.Secuencial

                Dim x As Integer
                For x = 0 To iTotalPO_Line - 1
                    baseGRPO.Lines.SetCurrentLine(x)
                    If baseGRPO.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close Then
                    Else
                        GRPO.Lines.ItemCode = baseGRPO.Lines.ItemCode
                        GRPO.Lines.WarehouseCode = baseGRPO.Lines.WarehouseCode
                        GRPO.Lines.Quantity = baseGRPO.Lines.Quantity
                        GRPO.Lines.BaseType = "22"
                        GRPO.Lines.BaseEntry = baseGRPO.DocEntry
                        GRPO.Lines.BaseLine = baseGRPO.Lines.LineNum
                        GRPO.Lines.Add()
                    End If
                Next

                ' Freight Charges
                'If iTotalFrgChg_Line > 0 Then
                '    Dim fcnt As Integer
                '    For fcnt = 0 To iTotalFrgChg_Line - 1
                '        GRPO.Expenses.SetCurrentLine(fcnt)
                '        GRPO.Expenses.ExpenseCode = baseGRPO.Expenses.ExpenseCode
                '        GRPO.Expenses.BaseDocType = "22"
                '        GRPO.Expenses.BaseDocLine = baseGRPO.Expenses.LineNum
                '        GRPO.Expenses.BaseDocEntry = baseGRPO.DocEntry
                '        GRPO.Expenses.Add()
                '    Next
                'End If

                GRPO.Comments += "Creado por el addon GSEDOC"

                'Add the Invoice
                RetVal = GRPO.Add

                'Check the result
                If RetVal <> 0 Then
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                    Return False
                Else
                    Return True
                End If

            End If

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un error al crear la factura preliminar", 1, "GSEDOC")
            Return False
        Finally
            baseGRPO = Nothing
            GRPO = Nothing

            GC.Collect()
        End Try


    End Function

    Private Function CrearFacturaPreliminarMapeada(sCardCode As String, oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura) As Boolean

        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String

#Disable Warning BC42024 ' Variable local sin usar: 'iTotalPO_Line'.
        Dim iTotalPO_Line As Integer
#Enable Warning BC42024 ' Variable local sin usar: 'iTotalPO_Line'.
#Disable Warning BC42024 ' Variable local sin usar: 'iTotalFrgChg_Line'.
        Dim iTotalFrgChg_Line As Integer
#Enable Warning BC42024 ' Variable local sin usar: 'iTotalFrgChg_Line'.

        'Dim baseGRPO As SAPbobsCOM.Documents
        'baseGRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

        Try

            ' If baseGRPO.GetByKey(PO_DocEntry) = True Then
            GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
            GRPO.CardCode = sCardCode
            GRPO.DocDate = oFactura.FechaEmision
            GRPO.DocDueDate = Today.Date

            'iTotalPO_Line = baseGRPO.Lines.Count
            'iTotalFrgChg_Line = baseGRPO.Expenses.Count

            ' DATOS DE AUTORIZACION
            GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = oFactura.AutorizacionSRI
            GRPO.UserFields.Fields.Item("U_SER_EST").Value = oFactura.Establecimiento
            GRPO.UserFields.Fields.Item("U_SER_PE").Value = oFactura.PuntoEmision

            'CREADAPORGSEDOC
            GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"

            GRPO.FolioNumber = oFactura.Secuencial
            Dim itemCode As String = ""
            Dim line As Integer = 0
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFactura.ENTDetalleFactura
                'SELECT DfltWH FROM OITM WITH(NOLOCK) WHERE ItemCode
                itemCode = oFuncionesB1.getRSvalue("SELECT ITemCode FROM OSCN WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "' AND Substitute = '" + oDetalle.CodigoPrincipal + "'", "ITemCode", "")
                GRPO.Lines.ItemCode = itemCode
                GRPO.Lines.WarehouseCode = oFuncionesB1.getRSvalue("SELECT DfltWH FROM OITM WITH(NOLOCK) WHERE ItemCode = '" + itemCode + "'", "DfltWH", "")
                GRPO.Lines.Quantity = oDetalle.Cantidad

                GRPO.Lines.Add()
                line += 1
            Next

            GRPO.Comments += "Creado por el addon GSEDOC"

            'Add the Invoice
            RetVal = GRPO.Add

            'Check the result
            If RetVal <> 0 Then
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                Return False
            Else
                Return True
            End If

            'End If

        Catch ex As Exception
            Return False
        Finally
            GRPO = Nothing
            GC.Collect()
        End Try


    End Function

#End Region

End Class

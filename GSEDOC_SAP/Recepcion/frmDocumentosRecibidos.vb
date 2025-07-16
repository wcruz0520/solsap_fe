Imports Entidades
'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Xml.Serialization
Imports System.IO

Public Class frmDocumentosRecibidos
    Public oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    'Estas Listas Contendran lo Cargado de Forma Manual



    '----fin de carga------------------------------

    '
    Private listaDocumentoPorUsuario As List(Of Entidades.DocumentoTipo)
    Private listaFCs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
    Private listaFCsFecha As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
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

    ''
    Dim _oDocumento As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion

    Dim EsRetCero As Boolean = False
    Dim ExistFactura As Boolean = False
    Dim Cuota As Integer = 0
    Dim _oFactura As SAPbobsCOM.Documents
    Dim Factura As String = ""
    Dim respN As Integer = 0, respS As String = ""
    Dim ListaFila As New List(Of Integer)
    Dim ListaDocEntryFact As New List(Of Integer)

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormularioDocumentosRecibidos()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        rsboApp.SetStatusBarMessage(NombreAddon + " - Consultando, Espere Por favor!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        If RecorreFormulario(rsboApp, "frmDocumentosRecibidos") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentosRecibidos.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentosRecibidos").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")

            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO

            oForm.Freeze(True)

            If Functions.VariablesGlobales._vgMostrarLogo = "Y" Then
                Dim ipLogoSS As SAPbouiCOM.PictureBox
                ipLogoSS = oForm.Items.Item("ipLogoSS").Specific
                ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            Else
                Dim ipLogoSS As SAPbouiCOM.PictureBox
                ipLogoSS = oForm.Items.Item("ipLogoSS").Specific
                ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
                ipLogoSS.Item.Visible = False
            End If

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
            txtRuc.ChooseFromListUID = "CFL1"
            txtRuc.ChooseFromListAlias = "CardCode"
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
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("IdDoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Seleccionar", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1)

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
                ' nombre_estado = oDatable.GetValue("EstadoDoc", x)
                'If nombre_estado = Estados_docenviados.EN_PROCESO_SRI Or nombre_estado = Estados_docenviados.ERROR_EN_RECEPCION Then
                '    ss_tipotabla = obtenerTipoTabla(oDatable.GetValue("ObjType", x), oDatable.GetValue("DocSubType", x))
                '    identificador = CInt(oDatable.GetValue("DocEntry", x))
                For y = 1 To oGrid.Rows.Count
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    If indexgrid = x Then
                        'oForm.Freeze(True)
                        'gcss.GetCellBackColor(y, 3)
                        '255, 255, 0  255000
                        gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                        Exit For
                    End If

                Next

                'rsboApp.MessageBox("ok")
                'oForm.Freeze(False)
                Try
                    ' oManejoDocumentos.ProcesaEnvioDocumento(identificador, ss_tipotabla, True)
                    '   rsboApp.MessageBox(ss_tipotabla & " " & oDatable.GetValue("DocEntry", x))
                    gcss.SetRowBackColor(y + 1, 255000)
                Catch ex As Exception
                    gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                End Try


                'End If
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next

            rsboApp.StatusBar.SetText("(SAED) El estado de los documentos han sido actualizados correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)



        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Ocurrio un error al llamar la funcion ConsultarEstados " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    'Private Function CrearPagoRecibido_E_ONormal(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
    '    Dim RetVal As Long
    '    Dim ErrCode As Long
    '    Dim ErrMsg As String
    '    Dim vPay As SAPbobsCOM.Payments

    '    Dim sQueryCodRetencionCB As String = ""
    '    Dim sQueryCuentaRetencionCB As String = ""
    '    Dim sQueryCrTypeCodeCB As String = ""
    '    Dim sQueryNombreCuentaRetencionCB As String = ""

    '    Dim CodRetencionCB As String = ""
    '    Dim CrTypeCodeCB As String = ""
    '    Dim CuentaRetencionCB As String = ""
    '    Dim NombreCuentaRetencionCB As String = ""

    '    Dim secuencialCB As Integer = 1
    '    Dim secuencialCBMP As Integer = 1
    '    Dim oCardCode As String = _oDocumento.Ruc
    '    Try

    '        'Dim vPay As SAPbobsCOM.Documents
    '        oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '        vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
    '        vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
    '        'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
    '        vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
    '        vPay.CardCode = oCardCode
    '        vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
    '        vPay.DocCurrency = "USD"
    '        'vPay.DocDate = DateSerisal(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now
    '        vPay.DocDate = _oDocumento.FechaAutorizacion
    '        vPay.TaxDate = _oDocumento.FechaEmision
    '        'vPay.DocDate = Date.Now
    '        'vPay.TaxDate = _oDocumento.FechaEmision
    '        vPay.DocRate = 0
    '        vPay.HandWritten = 0
    '        vPay.JournalRemarks = ""
    '        'vPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES
    '        vPay.Reference1 = ""
    '        '' vPay.Series = 0
    '        'vPay.TaxDate = DateSerial(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now            
    '        ' vPay.TaxDate = Date.Now


    '        '1 RENTA 2 IVA
    '        ' DETALLES
    '        Dim sQueryCodRetencion As String = ""
    '        Dim sQueryCuentaRetencion As String = ""
    '        Dim sQueryCrTypeCode As String = ""

    '        Dim CodRetencion As String = ""
    '        Dim CrTypeCode As String = ""
    '        Dim CuentaRetencion As String = ""

    '        Dim secuencial As Integer = 1
    '        For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In _oDocumento.ENTDetalleRetencion

    '            If oDetalle.ValorRetenido > 0 Then

    '                vPay.CreditCards.AdditionalPaymentSum = 0
    '                vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

    '                If oDetalle.Codigo = 1 Then ' RENTA

    '                    sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
    '                    CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
    '                    Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoRE")
    '                    If CodRetencion = "" Then
    '                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
    '                        Exit Function
    '                    End If
    '                ElseIf oDetalle.Codigo = 2 Then ' IVA

    '                    sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
    '                    CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
    '                    Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoRE")
    '                    If CodRetencion = "" Then
    '                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
    '                        Exit Function
    '                    End If
    '                ElseIf oDetalle.Codigo = 6 Then ' ISD

    '                    sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
    '                    CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
    '                    Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoRE")
    '                    If CodRetencion = "" Then
    '                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
    '                        Exit Function
    '                    End If
    '                End If

    '                sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
    '                CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
    '                Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmDocumentoRE")

    '                sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
    '                CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
    '                Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoRE")

    '                'vPay.CreditCards.CreditAcct = IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
    '                'vPay.CreditCards.CreditCard = IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
    '                'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)

    '                vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
    '                vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
    '                Try
    '                    vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
    '                Catch ex As Exception
    '                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + CrTypeCode.ToString(), "frmDocumentoRE")
    '                End Try



    '                vPay.CreditCards.CreditCardNumber = _oDocumento.Secuencial
    '                vPay.CreditCards.CreditSum = oDetalle.ValorRetenido ' _oDocumento.TotalRetencion ' formatDecimal(_oDocumento.TotalRetencion.ToString())
    '                ' vPay.CreditCards.CreditType = 1
    '                vPay.CreditCards.FirstPaymentSum = _oDocumento.TotalRetencion
    '                'vPay.CreditCards.NumOfCreditPayments = 1
    '                'vPay.CreditCards.NumOfPayments = 1

    '                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
    '                    vPay.CreditCards.FirstPaymentDue = _oDocumento.FechaAutorizacion

    '                    Try
    '                        If Not IsNothing(oDetalle.NumDocRetener) Then
    '                            'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
    '                            'Left(odt.GetValue(0, i).ToString(), 99))

    '                            'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Length)
    '                            vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
    '                            vPay.CreditCards.OwnerPhone = _oDocumento.Establecimiento + _oDocumento.PuntoEmision
    '                        End If

    '                    Catch ex As Exception
    '                    End Try


    '                    If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
    '                    End If
    '                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
    '                    End If
    '                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = _oDocumento.Secuencial
    '                    End If
    '                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = _oDocumento.AutorizacionSRI
    '                    End If
    '                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = _oDocumento.Establecimiento + _oDocumento.PuntoEmision
    '                    End If
    '                    If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = oCardCode
    '                    End If
    '                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
    '                    vPay.CreditCards.VoucherNum = _oDocumento.Establecimiento + _oDocumento.PuntoEmision + _oDocumento.Secuencial
    '                    If oFuncionesB1.checkCampoBD("RCT3", "NUM_AUT") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = _oDocumento.AutorizacionSRI
    '                    End If
    '                    If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
    '                        vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = _oDocumento.FechaAutorizacion
    '                    End If

    '                End If


    '                If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
    '                    vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
    '                End If
    '                If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
    '                    vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
    '                End If

    '                vPay.CreditCards.Add()
    '                vPay.CreditCards.SetCurrentLine(secuencial)
    '                secuencial += 1


    '            End If


    '        Next
    '        ' FACTURAS

    '        If CargaFacturaRelacionadas Then
    '            For Each o As Entidades.FacturaVenta In oListaFacturaVenta
    '                Utilitario.Util_Log.Escribir_Log("Datos Docentry: " & o.DocEntry.ToString & " valor a retener: " & o.ValorARetener.ToString, "datosfacturasrelacionadas")
    '                vPay.Invoices.DocEntry = o.DocEntry
    '                vPay.Invoices.SumApplied = o.ValorARetener
    '                vPay.Invoices.Add()
    '            Next
    '        End If



    '        RetVal = vPay.Add()
    '        'Dim xml As String = vPay.GetAsXML()
    '        'Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")



    '        If RetVal <> 0 Then
    '            'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
    '            Try
    '                Dim xml As String = vPay.GetAsXML()
    '                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
    '                Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoRE.oCardCode.ToString() + ofrmDocumentoRE._IdGS.ToString() + ".xml"
    '                'Dim xml As String = vPay.GetAsXML()
    '                If System.IO.Directory.Exists(sRutaCarpeta) Then
    '                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
    '                    Dim writer As TextWriter = New StreamWriter(sRuta)
    '                    writer.Write(xml)
    '                    writer.Close()
    '                End If
    '            Catch ex As Exception
    '                Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
    '            End Try

    '            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
    '            rCompany.GetLastError(ErrCode, ErrMsg)
    '            rsboApp.MessageBox(ErrCode & " " & ErrMsg)
    '            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '            Return False
    '        Else
    '            Try
    '                Dim xml As String = vPay.GetAsXML()
    '                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
    '                Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoRE.oCardCode.ToString() + ofrmDocumentoRE._IdGS.ToString() + ".xml"
    '                'Dim xml As String = vPay.GetAsXML()
    '                If System.IO.Directory.Exists(sRutaCarpeta) Then
    '                    Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
    '                    Dim writer As TextWriter = New StreamWriter(sRuta)
    '                    writer.Write(xml)
    '                    writer.Close()
    '                End If
    '            Catch ex As Exception
    '                Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
    '            End Try
    '            rCompany.GetNewObjectCode(sDocEntryPreliminar)
    '            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '            Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
    '            Return True
    '        End If
    '    Catch ex As Exception
    '        Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
    '        rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '        Return False
    '    Finally
    '        vPay = Nothing
    '        GC.Collect()
    '    End Try

    'End Function

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            'Dim typeEx, idForm As String
            'typeEx = oFuncionesB1.FormularioActivo(idForm)
            If pVal.FormTypeEx = "frmDocumentosRecibidos" Then
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
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
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
                                oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")
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

                                'arturo 10042024
                                Case "btnLodXML"

                                    'cargar Por XML
                                    CargarXML()

                                Case "obtnBuscar"
                                    'rsboApp.SetStatusBarMessage("Estará Disponible en una proxima version!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    cargarDocumentos()

                                Case "obtSel"
                                    Dim contadorSel As Integer = 0
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("obtSel").Specific
                                    If Seleccionar.Caption = "Seleccionar Todo" Then
                                        If SeleccionarDocumentosPendientes(contadorSel) Then
                                            Seleccionar.Caption = "Desmarcar Todo"
                                            rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se seleccionaron " + contadorSel.ToString + " registros para reenviar el correo con los archivos adjuntos.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    Else
                                        If DesmarcarDocumentosPendientes(contadorSel) Then
                                            Seleccionar.Caption = "Seleccionar Todo"
                                            rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se desmarcarón " + contadorSel.ToString + " registros para reenviar el correo con los archivos adjuntos.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            'oForm.Items.Item("btnCon").Enabled = False
                                            rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - No existen registros marcados.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    End If

                                Case "btnMarcar"
                                    MarcarDocIntegradoEnLote()

                                Case "oGrid"
                                    ofila = pVal.Row
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidos")
                                    Dim oGrid As SAPbouiCOM.Grid = oFor.Items.Item("oGrid").Specific
                                    oGrid.Rows.SelectedRows.Add(ofila)
                                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                        ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                                        ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                                        Dim isChecked As String = oGrid.DataTable.GetValue("Seleccionar", ofila)
                                        If isChecked = "Y" Then
                                            If Not ListaFila.Contains(ofila) Then
                                                ListaFila.Add(ofila)
                                            End If
                                        Else
                                            If ListaFila.Contains(ofila) Then
                                                ListaFila.Remove(ofila)
                                            End If
                                        End If
                                    Next

                                Case "btnAnt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidos")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) - 1, TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("obtSel").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                                Case "btnSig"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidos")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) + 1, TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("obtSel").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                                Case "btnPri"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidos")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), 1, TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("obtSel").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                                Case "btnUlt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidos")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbNP.Caption), TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("obtSel").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pVal.ColUID = "RUC" And pVal.BeforeAction = True Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
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
                                    Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                                    Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE _
                                    Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

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
                                    Dim sQueryProveedor As String = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardType"" = 'C' and ""LicTradNum""= '" + sLicTradNum.ToString + "' "
                                    Dim _sQueryProveedor As String = oFuncionesB1.getRSvalue(sQueryProveedor, "U_SSCLIENTEBANCO", "")
                                    Utilitario.Util_Log.Escribir_Log("Query Buscar SN Cliente/Proveedor: " + sQueryProveedor.ToString + " Resultado:" + _sQueryProveedor.ToString, "frmDocumentosRecibidos")
                                    If _sQueryProveedor = "PROVEEDOR" Then
                                        sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")

                                    Else
                                        sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                                        Utilitario.Util_Log.Escribir_Log("Query Buscar SN Cliente: " + QueryExisteCliente.ToString + " Resultado:" + sCardCode.ToString, "frmDocumentosRecibidos")
                                    End If

                                Else
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                                    Utilitario.Util_Log.Escribir_Log("Query Buscar SN Proveedor: " + QueryExisteProveedor.ToString + " Resultado:" + sCardCode.ToString, "frmDocumentosRecibidos")
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
                        If pVal.BeforeAction = False And pVal.ItemUID = "oGrid" Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

                            'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                            'oGrid.Rows.SelectedRows.Add(ofila)

                            Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                            Dim QueryExisteProveedor As String = ""
                            Dim QueryExisteCliente As String = ""
                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE _
                                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
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
                                Dim sQueryProveedor As String = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardType"" = 'C' and ""LicTradNum""= '" + sRUC.ToString + "' "
                                Dim _sQueryProveedor As String = oFuncionesB1.getRSvalue(sQueryProveedor, "U_SSCLIENTEBANCO", "")
                                Utilitario.Util_Log.Escribir_Log("Query Buscar SN Cliente/Proveedor: " + sQueryProveedor.ToString + " Resultado:" + _sQueryProveedor.ToString, "frmDocumentosRecibidos")
                                If _sQueryProveedor = "PROVEEDOR" Then
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                                Else
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                                End If

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
                                Dim iIdDocEdoc As Object = oDataTable.GetValue(14, ofila).ToString()

                                rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Documento de " + sNombre + ", por favor espere..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                WS.Url = Functions.VariablesGlobales._WS_Recepcion

                                'MANEJO DE PROXY
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

                                    WS.Proxy = proxyobject
                                    WS.Credentials = cred

                                End If
                                ' END MANEJO DE PROXY
                                If oDataTable.GetValue(0, ofila).ToString() = "Factura" Then
                                    If Functions.VariablesGlobales._TipoWS = "LOCAL" Then
                                        Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                        results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                        For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                            oFactura = oFac
                                        Next
                                    End If
                                    'Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                    'results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    'For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                    '    oFactura = oFac
                                    'Next
                                    Dim sQueryIdDocumento As String = ""
                                    Dim idDocumentoRecibido_UDO As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sQueryIdDocumento = "SELECT ""U_SSIDDOCUMENTO"" FROM ""ODRF"" WHERE ""ObjType"" = 18 and ""DocEntry"" = " + iBorrador.ToString()
                                    Else
                                        sQueryIdDocumento = "select U_SSIDDOCUMENTO from ODRF Where ObjType = 18 and DocEntry = " + iBorrador.ToString()
                                    End If
                                    If iBorrador = 0 Then
                                        mensaje = ""
                                        Dim documentroIntegradoPorXml As Boolean = False

                                        If Functions.VariablesGlobales._TipoWS <> "LOCAL" Then
                                            oFactura = Nothing

                                            'Logica Agregada para ver si el detalle del documento
                                            ' se debe consultar o ya esta en la lista de los cargados por XML

                                            If (listaFCs.Where(Function(f) f.ClaveAcceso = iIdDocEdoc).Count <= 0) Then
                                                'si no esta se consulta del WS y si lo esta se recupera de la lista
                                                SetProtocolosdeSeguridad()
                                                oFactura = WS.ConsultarFactura_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)
                                            Else

                                                oFactura = listaFCs.Where(Function(f) f.ClaveAcceso = iIdDocEdoc).FirstOrDefault()
                                                documentroIntegradoPorXml = True
                                            End If



                                        End If

                                        ofrmDocumento.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oFactura, ofila, documentroIntegradoPorXml)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumento.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If

                                ElseIf oDataTable.GetValue(0, ofila).ToString() = "Nota de Crédito" Then
                                    If Functions.VariablesGlobales._TipoWS = "LOCAL" Then
                                        Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                                        results = listaNCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                        For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results
                                            oNotaDeCredito = oNC
                                        Next
                                    End If
                                    'Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                                    'results = listaNCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    'For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results
                                    '    oNotaDeCredito = oNC
                                    'Next
                                    Dim sQueryIdDocumento As String = ""
                                    Dim idDocumentoRecibido_UDO As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sQueryIdDocumento = "SELECT ""U_SSIDDOCUMENTO"" FROM ""ODRF"" WHERE ""ObjType"" = 19 and ""DocEntry"" = " + iBorrador.ToString()
                                    Else
                                        sQueryIdDocumento = "select U_SSIDDOCUMENTO from ODRF Where ObjType = 19 and DocEntry = " + iBorrador.ToString()
                                    End If
                                    If iBorrador = 0 Then
                                        mensaje = ""
                                        Dim documentroIntegradoPorXml As Boolean = False

                                        If Functions.VariablesGlobales._TipoWS <> "LOCAL" Then
                                            oNotaDeCredito = Nothing
                                            SetProtocolosdeSeguridad()
                                            ' oNotaDeCredito = WS.ConsultarNotaCredito_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)

                                            'Logica Agregada para ver si el detalle del documento
                                            ' se debe consultar o ya esta en la lista de los cargados por XML

                                            If (listaNCs.Where(Function(f) f.ClaveAcceso = iIdDocEdoc).Count <= 0) Then
                                                'si no esta se consulta del WS y si lo esta se recupera de la lista
                                                SetProtocolosdeSeguridad()
                                                oNotaDeCredito = WS.ConsultarNotaCredito_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)
                                            Else

                                                oNotaDeCredito = listaNCs.Where(Function(f) f.ClaveAcceso = iIdDocEdoc).FirstOrDefault()
                                                documentroIntegradoPorXml = True
                                            End If


                                        End If

                                        ofrmDocumentoNC.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oNotaDeCredito, ofila, documentroIntegradoPorXml)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumentoNC.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If

                                ElseIf oDataTable.GetValue(0, ofila).ToString() = "Retención de Cliente" Then

                                    If Functions.VariablesGlobales._TipoWS = "LOCAL" Then
                                        Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                                        results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                        For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                                            oRetencion = oRE
                                        Next
                                    End If
                                    'Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                                    'results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    'For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                                    '    oRetencion = oRE
                                    'Next
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
                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            sQueryIdDocumento += " SELECT B.""U_SSIDDOCUMENTO"""
                                            sQueryIdDocumento += " FROM ""@TM_LE_RETVH"" B "
                                            sQueryIdDocumento += " WHERE B.""DocEntry"" = " + iBorrador.ToString()
                                            sQueryIdDocumento += " AND B.""U_SSCREADAR"" = 'SI'"
                                            sQueryIdDocumento += " AND B.""Object"" = 'TM_RETV'"
                                        Else
                                            sQueryIdDocumento += " SELECT B.U_SSIDDOCUMENTO"
                                            sQueryIdDocumento += " FROM @TM_LE_RETVH B "
                                            sQueryIdDocumento += " WHERE B.DocEntry = " + iBorrador.ToString()
                                            sQueryIdDocumento += " AND B.U_SSCREADAR = 'SI'"
                                            sQueryIdDocumento += " AND B.Object  = 'TM_RETV'"
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

                                        Dim documentroIntegradoPorXml As Boolean = False

                                        If Functions.VariablesGlobales._TipoWS <> "LOCAL" Then
                                            oRetencion = Nothing
                                            mensaje = ""

                                            'Logica Agregada para ver si el detalle del documento
                                            ' se debe consultar o ya esta en la lista de los cargados por XML

                                            If (listaREs.Where(Function(f) f.ClaveAcceso = iIdDocEdoc).Count <= 0) Then
                                                SetProtocolosdeSeguridad()
                                                oRetencion = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)
                                            Else

                                                oRetencion = listaREs.Where(Function(f) f.ClaveAcceso = iIdDocEdoc).FirstOrDefault()
                                                documentroIntegradoPorXml = True

                                            End If



                                        End If

                                        ofrmDocumentoRE.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oRetencion, ofila, documentroIntegradoPorXml)
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

                Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                oDataTable.Rows.Clear()

                MarcarVistosDocumentosPendientes()

                If CargarDocumento() Then
                    rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Documentos Recibidos, Listo..!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Else
                    ' listaFCs = New List(Of wsEDoc_ConsultaRecepcion.ENTFactura)

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
    Private Sub MarcarVistosDocumentosPendientes()

        Try
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")
            Dim cbxTipo As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipo").Specific

            Try
                oForm.DataSources.DataTables.Add("dtVIS")
            Catch ex As Exception
            End Try

            Dim dtVIS As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtVIS")
            Dim QueryV As String = ""
            Dim codDoc As Integer = 0
            dtVIS.Rows.Clear()

            Select Case cbxTipo.Value
                Case "01"
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        QueryV = " SELECT top 20 ""U_IdGS"", ""DocEntry"" As IdDocumentoRecibido, 1 AS TipoDoc, ""U_ClaAcc"" As ClaveAcceso "
                        QueryV += " FROM ""@GS_FVR"" where ""U_Estado"" = 'docFinal' and ""U_Sincro"" = 1 AND IFNULL(""U_SincroE"",0) = 0 "
                    Else
                        QueryV = " SELECT top 20 U_IdGS, DocEntry As IdDocumentoRecibido, 1 AS TipoDoc, U_ClaAcc As ClaveAcceso "
                        QueryV += " FROM ""@GS_FVR"" where U_Estado = 'docFinal' and U_Sincro = 1 AND ISNULL(U_SincroE,0) = 0 "
                    End If
                    codDoc = 1
                Case "04"
                    'Dim QueryV As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        QueryV = " SELECT top 20 ""U_IdGS"", ""DocEntry"" As IdDocumentoRecibido, 1 AS TipoDoc, ""U_ClaAcc"" As ClaveAcceso "
                        QueryV += " FROM ""@GS_NCR"" where ""U_Estado"" = 'docFinal' and ""U_Sincro"" = 1 AND IFNULL(""U_SincroE"",0) = 0 "
                    Else
                        QueryV = " SELECT top 20 U_IdGS, DocEntry As IdDocumentoRecibido, 1 AS TipoDoc, U_ClaAcc As ClaveAcceso "
                        QueryV += " FROM ""@GS_NCR"" where U_Estado = 'docFinal' and U_Sincro = 1 AND ISNULL(U_SincroE,0) = 0 "
                    End If
                    codDoc = 3
                Case "07"
                    'Dim QueryV As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        QueryV = " SELECT top 20 ""U_IdGS"", ""DocEntry"" As IdDocumentoRecibido, 1 AS TipoDoc, ""U_ClaAcc"" As ClaveAcceso "
                        QueryV += " FROM ""@GS_RER"" where ""U_Estado"" = 'docFinal' and ""U_Sincro"" = 1 AND IFNULL(""U_SincroE"",0) = 0 "
                    Else
                        QueryV = " SELECT top 20 U_IdGS, DocEntry As IdDocumentoRecibido, 1 AS TipoDoc, U_ClaAcc As ClaveAcceso "
                        QueryV += " FROM ""@GS_RER"" where U_Estado = 'docFinal' and U_Sincro = 1 AND ISNULL(U_SincroE,0) = 0 "
                    End If
                    codDoc = 2
            End Select

            'DesdeF.ToString("yyyyMMdd")

            Utilitario.Util_Log.Escribir_Log("QueryV: " + QueryV.ToString(), "frmDocumentosRecibidos")
            Try
                oForm.DataSources.DataTables.Item("dtVIS").ExecuteQuery(QueryV)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error QueryV : " + ex.Message.ToString(), "frmDocumentosRecibidos")
            End Try

            Dim U_IdGS As Integer = 0
            Dim IdDocumentoRecibido As String = ""
            Dim TipoDoc As Integer = 0
            Dim ClaveAcceso As String = ""
            Dim Mensaje As String = ""

            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos pendientes de marcar integrados : " + dtVIS.Rows.Count().ToString(), "frmDocumentosRecibidos")

            For i As Integer = 0 To dtVIS.Rows.Count - 1

                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    U_IdGS = dtVIS.GetValue("U_IdGS", i)
                    IdDocumentoRecibido = dtVIS.GetValue("IDDOCUMENTORECIBIDO", i).ToString.Trim
                    TipoDoc = dtVIS.GetValue("TIPODOC", i)
                    ClaveAcceso = dtVIS.GetValue("CLAVEACCESO", i).ToString().Trim
                Else
                    U_IdGS = dtVIS.GetValue("U_IdGS", i)
                    IdDocumentoRecibido = dtVIS.GetValue("IdDocumentoRecibido", i).ToString.Trim
                    TipoDoc = dtVIS.GetValue("TipoDoc", i)
                    ClaveAcceso = dtVIS.GetValue("ClaveAcceso", i).ToString().Trim
                End If

                Utilitario.Util_Log.Escribir_Log("U_IdGS : " + U_IdGS.ToString(), "frmDocumentosRecibidos")
                Utilitario.Util_Log.Escribir_Log("IdDocumentoRecibido : " + IdDocumentoRecibido.ToString(), "frmDocumentosRecibidos")
                Utilitario.Util_Log.Escribir_Log("TipoDoc : " + TipoDoc.ToString(), "frmDocumentosRecibidos")
                Utilitario.Util_Log.Escribir_Log("ClaveAcceso : " + ClaveAcceso.ToString(), "frmDocumentosRecibidos")

                If codDoc = 1 Then
                    Try
                        ofrmDocumento.MarcarVisto(U_IdGS, codDoc, Mensaje, IdDocumentoRecibido)
                        Utilitario.Util_Log.Escribir_Log("ReProceso Visto(Integrado FC) en EDOC: " + ClaveAcceso + Mensaje.ToString(), "frmDocumentosRecibidos")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error ofrmDocumento.MarcarVisto : " + ex.Message.ToString(), "frmDocumentosRecibidos")
                    End Try

                ElseIf codDoc = 3 Then
                    ofrmDocumentoNC.MarcarVisto(U_IdGS, codDoc, Mensaje, IdDocumentoRecibido)
                    Utilitario.Util_Log.Escribir_Log("ReProceso Visto(Integrado NC) en EDOC: " + ClaveAcceso + Mensaje.ToString(), "frmDocumentosRecibidos")
                ElseIf codDoc = 2 Then
                    ofrmDocumentoRE.MarcarVisto(U_IdGS, codDoc, Mensaje, IdDocumentoRecibido)
                    Utilitario.Util_Log.Escribir_Log("ReProceso Visto(Integrado RT) en EDOC: " + ClaveAcceso + Mensaje.ToString(), "frmDocumentosRecibidos")
                End If

                rsboApp.StatusBar.SetText(NombreAddon + " - Documento  " + ClaveAcceso + " Actualizado", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Next

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Reproceso Marcar Integrados eDoc: " + ex.Message.ToString(), "frmDocumentosRecibidos")
        End Try

    End Sub

    Public Function CargarDocumento() As Boolean
        Try
            'CONSULTO EL RUC DE LA BASE ACTUAL            
            ' _RUC = oFuncionesB1.getRSvalue("SELECT TAXIDNUM FROM OADM", "TAXIDNUM", "")

            ' OBTENGO URL DEL SERVICIO DE RECEPCION
            'RegistrosXPaginas = oFuncionesAddon.getRSvalue("SELECT TOP 1 ""U_Valor"" from ""@GS_CONFD"" where U_Nombre='Registros_por_paginas'", "U_Valor", "")
            Dim nRegistros As String = Functions.VariablesGlobales._nRegistros
            If nRegistros = "" Then
                RegistrosXPaginas = 10
            Else
                RegistrosXPaginas = nRegistros
            End If


            _WS_Recepcion = Functions.VariablesGlobales._WS_Recepcion
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = Functions.VariablesGlobales._WS_RecepcionCambiarEstado
            _WS_RecepcionClave = Functions.VariablesGlobales._WS_RecepcionClave
            'FAMC cargo los estados parametrizados
            _WS_RecepcionCargaEstados = Functions.VariablesGlobales._WS_RecepcionCargaEstados

            If String.IsNullOrEmpty(_WS_RecepcionCargaEstados) Then
                _WS_RecepcionCargaEstados = "1"
            End If
            'RegistrosXPaginas =  oUserTable.UserFields.Fields.Item("U_ws_RP").Value

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
            'Dim WS As New Entidades.WSRAD_KEY_CONSULTA4_3.WSRAD_KEY_CONSULTA
            WS.Url = _WS_Recepcion

            'MANEJO DE PROXY
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
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
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

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If
                    'If TipoWS = "NUBE_4_1" Then

                    'End If
                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    'If Functions.VariablesGlobales._MarcarContabiliadosManualFC = "Y" Then
                    '    rsboApp.SetStatusBarMessage(NombreAddon + " - Se procedera a verificar las facturas ya contabilizadas, por favor espere un momento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    '    'oManejoDocumentos.SetProtocolosdeSeguridad()
                    '    SetProtocolosdeSeguridad()
                    '    Dim x = WS.ConsultarFactura(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    '    MarcarDocumentosContabilizadosManualFC(x)
                    'End If
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'oManejoDocumentos.SetProtocolosdeSeguridad()

                    Dim oFiltrosRecepcionEC As New wsEDoc_ConsultaRecepcion.ClsBusqueda
                    oFiltrosRecepcionEC.CiaTipoAlojamientoKey = Functions.VariablesGlobales._WS_RecepcionClave
                    oFiltrosRecepcionEC.RucProveedor = LicTradNum
                    oFiltrosRecepcionEC.Estado = _WS_RecepcionCargaEstados

                    Dim txtNumDoc As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDoc").Specific
                    If Not txtNumDoc.Value = "" Then
                        oFiltrosRecepcionEC.NumDocumento = txtNumDoc.Value
                    End If

                    ' FILTRO: Rango de Fechas
                    Dim txtFechaD As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaD").Specific
                    Dim txtFechaH As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaH").Specific
                    Dim dfechaDesde As Date
                    Dim dfechaHasta As Date
                    If Not oFuncionesB1.BobStringToDate(txtFechaD.Value, dfechaDesde) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Desde es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not oFuncionesB1.BobStringToDate(txtFechaH.Value, dfechaHasta) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Hasta es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If dfechaDesde > dfechaHasta Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - La Fecha Desde no puede ser mayor que la Fecha Hasta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not txtFechaD.Value = "" Then
                        If txtFechaH.Value = "" Then
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Debe ingresar ambas fechas para realizar la consulta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oForm.Freeze(False)
                            Return False
                            Exit Function
                        End If
                    End If
                    If Not txtFechaH.Value = "" Then
                        If txtFechaD.Value = "" Then
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Debe ingresar ambas fechas para realizar la consulta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oForm.Freeze(False)
                            Return False
                            Exit Function
                        End If
                    End If
                    If Not txtFechaH.Value = "" Then
                        oFiltrosRecepcionEC.FechaEmisionDesde = dfechaDesde
                        oFiltrosRecepcionEC.FechaEmisionHasta = dfechaHasta
                    Else
                        oFiltrosRecepcionEC.FechaEmisionDesde = Nothing
                        oFiltrosRecepcionEC.FechaEmisionHasta = Nothing
                    End If

                    Try
                        Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                        Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Recepcion_FC" + ".xml"
                        If System.IO.Directory.Exists(sRutaCarpeta) Then
                            Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de Busqueda", "frmDocumentosRecibidos")

                            Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_ConsultaRecepcion.ClsBusqueda))
                            Dim writer As TextWriter = New StreamWriter(sRuta)
                            x.Serialize(writer, oFiltrosRecepcionEC)
                            writer.Close()
                            Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda" + sRuta, "frmDocumentosRecibidos")
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    SetProtocolosdeSeguridad()
                    Dim z As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)

                    If Functions.VariablesGlobales._MarcarContabiliadosManualFC = "Y" Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - Se procedera a verificar las facturas ya contabilizadas, por favor espere un momento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        Dim x = WS.ConsultarFactura_CabeceraBuscar(oFiltrosRecepcionEC, mensaje).ToList
                        MarcarDocumentosContabilizadosManualFC(x)
                    End If

                    If Functions.VariablesGlobales._TipoWS = "LOCAL" Then
                        z = WS.ConsultarFactura(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    Else
                        z = WS.ConsultarFactura_CabeceraBuscar(oFiltrosRecepcionEC, mensaje).ToList
                    End If
                    'Dim z = WS.ConsultarFactura_CabeceraBuscar(oFiltrosRecepcionEC, mensaje).ToList
                    'Dim z = WS.ConsultarFactura(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    'WS.con()
                    If Not z Is Nothing Then
                        i = z.Count
                        listaFCs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura))
                        'ORDENA LA LISTA DE FACTURAS POR FECHA DE EMISION
                        'listaFCs = (From M In listaFCs Order By M.FechaEmision Descending Select M).ToList
                        'listaFCs = (From l In listaFCs Order By l.FechaEmision.Month Descending Select l).ToList

                        'Se agregaran Las facturas con Estado docPrelXML- Artur
                        AgregarDocumentoEnEstadoDocPrelXML(cbxTipo.Value)


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
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
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
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    'If Functions.VariablesGlobales._MarcarContabiliadosManualNC = "Y" Then
                    '    rsboApp.SetStatusBarMessage(NombreAddon + " - Se procedera a verificar las notas de credito ya contabilizadas, por favor espere un momento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    '    'oManejoDocumentos.SetProtocolosdeSeguridad()
                    '    SetProtocolosdeSeguridad()
                    '    Dim x = WS.ConsultarNC(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    '    MarcarDocumentosContabilizadosManualNC(x)
                    'End If
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'oManejoDocumentos.SetProtocolosdeSeguridad()
                    Dim oFiltrosRecepcionECNC As New wsEDoc_ConsultaRecepcion.ClsBusqueda
                    oFiltrosRecepcionECNC.CiaTipoAlojamientoKey = Functions.VariablesGlobales._WS_RecepcionClave
                    oFiltrosRecepcionECNC.RucProveedor = LicTradNum
                    oFiltrosRecepcionECNC.Estado = _WS_RecepcionCargaEstados

                    Dim txtNumDoc As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDoc").Specific
                    If Not txtNumDoc.Value = "" Then
                        oFiltrosRecepcionECNC.NumDocumento = txtNumDoc.Value
                    End If

                    ' FILTRO: Rango de Fechas
                    Dim txtFechaD As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaD").Specific
                    Dim txtFechaH As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaH").Specific
                    Dim dfechaDesde As Date
                    Dim dfechaHasta As Date
                    If Not oFuncionesB1.BobStringToDate(txtFechaD.Value, dfechaDesde) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Desde es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not oFuncionesB1.BobStringToDate(txtFechaH.Value, dfechaHasta) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Hasta es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If dfechaDesde > dfechaHasta Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - La Fecha Desde no puede ser mayor que la Fecha Hasta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not txtFechaD.Value = "" Then
                        If txtFechaH.Value = "" Then
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Debe ingresar ambas fechas para realizar la consulta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oForm.Freeze(False)
                            Return False
                            Exit Function
                        End If
                    End If
                    If Not txtFechaH.Value = "" Then
                        If txtFechaD.Value = "" Then
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Debe ingresar ambas fechas para realizar la consulta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oForm.Freeze(False)
                            Return False
                            Exit Function
                        End If
                    End If
                    If Not txtFechaH.Value = "" Then
                        oFiltrosRecepcionECNC.FechaEmisionDesde = dfechaDesde
                        oFiltrosRecepcionECNC.FechaEmisionHasta = dfechaHasta
                    Else
                        oFiltrosRecepcionECNC.FechaEmisionDesde = Nothing
                        oFiltrosRecepcionECNC.FechaEmisionHasta = Nothing
                    End If

                    Try
                        Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                        Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Recepcion_NC" + ".xml"
                        If System.IO.Directory.Exists(sRutaCarpeta) Then
                            Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de Busqueda", "frmDocumentosRecibidos")

                            Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_ConsultaRecepcion.ClsBusqueda))
                            Dim writer As TextWriter = New StreamWriter(sRuta)
                            x.Serialize(writer, oFiltrosRecepcionECNC)
                            writer.Close()
                            Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda NC" + sRuta, "frmDocumentosRecibidos")
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    Dim z As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                    SetProtocolosdeSeguridad()
                    mensaje = ""
                    If Functions.VariablesGlobales._MarcarContabiliadosManualNC = "Y" Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - Se procedera a verificar las notas de credito ya contabilizadas, por favor espere un momento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        Dim x = WS.ConsultarNotaCredito_CabeceraBuscar(oFiltrosRecepcionECNC, mensaje).ToList
                        MarcarDocumentosContabilizadosManualNC(x)
                    End If
                    mensaje = ""
                    If Functions.VariablesGlobales._TipoWS = "LOCAL" Then
                        z = WS.ConsultarNC(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    Else
                        z = WS.ConsultarNotaCredito_CabeceraBuscar(oFiltrosRecepcionECNC, mensaje).ToList
                    End If
                    'Dim z = WS.ConsultarNC(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList

                    If Not z Is Nothing Then
                        i = z.Count
                        listaNCs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito))

                        'Se intenta agregar las NC cargadas desde XML

                        AgregarDocumentoEnEstadoDocPrelXML(cbxTipo.Value)

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
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
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

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'C' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    If Functions.VariablesGlobales._MarcarContabiliadosManualRT = "Y" Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - Se procedera a verificar las retenciones ya contabilizadas, por favor espere un momento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                        'oManejoDocumentos.SetProtocolosdeSeguridad()
                        SetProtocolosdeSeguridad()
                        Dim x = WS.ConsultarRetencion(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                        MarcarDocumentosContabilizadosManualRT(x)
                    End If
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'oManejoDocumentos.SetProtocolosdeSeguridad()
                    Dim oFiltrosRecepcionECRT As New wsEDoc_ConsultaRecepcion.ClsBusqueda
                    oFiltrosRecepcionECRT.CiaTipoAlojamientoKey = Functions.VariablesGlobales._WS_RecepcionClave
                    oFiltrosRecepcionECRT.RucProveedor = LicTradNum
                    oFiltrosRecepcionECRT.Estado = _WS_RecepcionCargaEstados

                    Dim txtNumDoc As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDoc").Specific
                    If Not txtNumDoc.Value = "" Then
                        oFiltrosRecepcionECRT.NumDocumento = txtNumDoc.Value
                    End If

                    ' FILTRO: Rango de Fechas
                    Dim txtFechaD As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaD").Specific
                    Dim txtFechaH As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaH").Specific
                    Dim dfechaDesde As Date
                    Dim dfechaHasta As Date
                    If Not oFuncionesB1.BobStringToDate(txtFechaD.Value, dfechaDesde) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Desde es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not oFuncionesB1.BobStringToDate(txtFechaH.Value, dfechaHasta) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Hasta es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If dfechaDesde > dfechaHasta Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - La Fecha Desde no puede ser mayor que la Fecha Hasta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not txtFechaD.Value = "" Then
                        If txtFechaH.Value = "" Then
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Debe ingresar ambas fechas para realizar la consulta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oForm.Freeze(False)
                            Return False
                            Exit Function
                        End If
                    End If
                    If Not txtFechaH.Value = "" Then
                        If txtFechaD.Value = "" Then
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Debe ingresar ambas fechas para realizar la consulta..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            oForm.Freeze(False)
                            Return False
                            Exit Function
                        End If
                    End If
                    If Not txtFechaH.Value = "" Then
                        oFiltrosRecepcionECRT.FechaEmisionDesde = dfechaDesde
                        oFiltrosRecepcionECRT.FechaEmisionHasta = dfechaHasta
                    Else
                        oFiltrosRecepcionECRT.FechaEmisionDesde = Nothing
                        oFiltrosRecepcionECRT.FechaEmisionHasta = Nothing
                    End If

                    Try
                        Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                        Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Recepcion_RT" + ".xml"
                        If System.IO.Directory.Exists(sRutaCarpeta) Then
                            Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de Busqueda", "frmDocumentosRecibidos")

                            Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_ConsultaRecepcion.ClsBusqueda))
                            Dim writer As TextWriter = New StreamWriter(sRuta)
                            x.Serialize(writer, oFiltrosRecepcionECRT)
                            writer.Close()
                            Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda RT" + sRuta, "frmDocumentosRecibidos")
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    Dim z As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                    SetProtocolosdeSeguridad()
                    mensaje = ""
                    If Functions.VariablesGlobales._MarcarContabiliadosManualRT = "Y" Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - Se procedera a verificar las retenciones ya contabilizadas, por favor espere un momento", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        Dim x = WS.ConsultarRetencion_CabeceraBuscar(oFiltrosRecepcionECRT, mensaje).ToList
                        MarcarDocumentosContabilizadosManualRT(x)
                    End If
                    mensaje = ""
                    If Functions.VariablesGlobales._TipoWS = "LOCAL" Then
                        z = WS.ConsultarRetencion(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    Else
                        z = WS.ConsultarRetencion_CabeceraBuscar(oFiltrosRecepcionECRT, mensaje).ToList
                    End If
                    'Dim z = WS.ConsultarRetencion(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList

                    If Not z Is Nothing Then
                        i = z.Count
                        listaREs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))

                        AgregarDocumentoEnEstadoDocPrelXML(cbxTipo.Value)

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

    Private Sub AgregarDocumentoEnEstadoDocPrelXML(tipoDoc As String)

        Select Case tipoDoc
            Case "01"

                Dim rs As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim query = "Select ""U_ClaAcc"" FROM ""@GS_FVR"" where ""U_Estado""='docPrelXML'"

                rs.DoQuery(query)

                If rs.RecordCount > 0 Then

                    Dim cabeceraDocumento As New Entidades.wsEDoc_ConsultaRecepcion.ENTFactura
                    Dim claveaccesoPrel As String = ""
                    Dim rutaArchivoFuente As String = ""
                    While Not rs.EoF

                        claveaccesoPrel = rs.Fields.Item("U_ClaAcc").Value

                        rutaArchivoFuente = $"{Functions.VariablesGlobales._RutaIntegracionXML}\{claveaccesoPrel}.xml"

                        If File.Exists(rutaArchivoFuente) Then

                            If (listaFCs.Where(Function(f) f.ClaveAcceso = claveaccesoPrel).Count <= 0) Then

                                cabeceraDocumento = Negocio.ssXML.OperacionesXML.LeerXMLFactura2(rutaArchivoFuente)

                                If Not IsNothing(cabeceraDocumento) Then

                                    listaFCs.Add(cabeceraDocumento)

                                End If

                            End If


                        End If

                        rs.MoveNext()

                    End While



                End If

                Release(rs)

            Case "04"

                Dim rs As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim query = "Select ""U_ClaAcc"" FROM ""@GS_NCR"" where ""U_Estado""='docPrelXML'"

                rs.DoQuery(query)

                If rs.RecordCount > 0 Then

                    Dim cabeceraDocumento As New Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito
                    Dim claveaccesoPrel As String = ""
                    Dim rutaArchivoFuente As String = ""
                    While Not rs.EoF

                        claveaccesoPrel = rs.Fields.Item("U_ClaAcc").Value

                        rutaArchivoFuente = $"{Functions.VariablesGlobales._RutaIntegracionXML}\{claveaccesoPrel}.xml"

                        If File.Exists(rutaArchivoFuente) Then

                            If (listaNCs.Where(Function(f) f.ClaveAcceso = claveaccesoPrel).Count <= 0) Then

                                cabeceraDocumento = Negocio.ssXML.OperacionesXML.LeerXMLNotaCredito2(rutaArchivoFuente)

                                If Not IsNothing(cabeceraDocumento) Then

                                    listaNCs.Add(cabeceraDocumento)

                                End If

                            End If


                        End If

                        rs.MoveNext()

                    End While

                End If

                Release(rs)

            Case "07"

                Dim rs As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim query = "Select ""U_ClaAcc"" FROM ""@GS_RER"" where ""U_Estado""='docPrelXML'"

                rs.DoQuery(query)

                If rs.RecordCount > 0 Then

                    Dim cabeceraDocumento As New Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                    Dim claveaccesoPrel As String = ""
                    Dim rutaArchivoFuente As String = ""
                    While Not rs.EoF

                        claveaccesoPrel = rs.Fields.Item("U_ClaAcc").Value

                        rutaArchivoFuente = $"{Functions.VariablesGlobales._RutaIntegracionXML}\{claveaccesoPrel}.xml"

                        If File.Exists(rutaArchivoFuente) Then

                            If (listaREs.Where(Function(f) f.ClaveAcceso = claveaccesoPrel).Count <= 0) Then

                                cabeceraDocumento = Negocio.ssXML.OperacionesXML.LeerXMLRetencion2(rutaArchivoFuente)

                                If Not IsNothing(cabeceraDocumento) Then

                                    listaREs.Add(cabeceraDocumento)

                                End If

                            End If


                        End If

                        rs.MoveNext()

                    End While

                End If

                Release(rs)

            Case Else

        End Select


    End Sub
    Public Sub Release(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub llenarGrid(tipoDoc As String, RegistrosXPaginas As Integer, NumeroPaginas As Integer, PaginaActual As Integer, TotalDocs As Integer)
        Dim i As Integer = 0
        Dim PendienteMapear As Boolean = False
        Dim NumIni As Integer = 0
        Dim NumFin As Integer = 0
        Dim RangoHasta As Integer = 0
        Dim m_oProgBar As SAPbouiCOM.ProgressBar
        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
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
                NumIni = NumFin - RegistrosXPaginas
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
                            sSucursal = Functions.VariablesGlobales._Adicional_FC
                            sSucursal = sSucursal.Replace("RUC", sRUC.ToString)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        Else
                            sSucursal = Functions.VariablesGlobales._Adicional_FC
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
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NROAUTOO"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
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
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_TM_NAUT"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_TM_NAUT = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_HBT_AUT_FAC"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_HBT_AUT_FAC = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SS_NumAut"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SS_NumAut = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
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
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(oFactura.RazonSocial, 250).Trim)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oFactura.ImporteTotal))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oFactura.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oFactura.AutorizacionSRI)


                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("IdDoc", i, oFactura.IdFactura.ToString)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Seleccionar", i, "N")
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
                            sSucursal = Functions.VariablesGlobales._Adicional_NC
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "County", "")
                        Else
                            sSucursal = Functions.VariablesGlobales._Adicional_NC
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
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_TM_NAUT"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_TM_NAUT = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_HBT_AUT_FAC"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_HBT_AUT_FAC = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SS_NumAut"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SS_NumAut = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
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
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Trim(Left(oNC.RazonSocial, 250).Trim))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oNC.ValorModificacion))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oNC.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oNC.AutorizacionSRI)
                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("IdDoc", i, oNC.IdNotaCredito.ToString)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Seleccionar", i, "N")
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
                            sSucursal = Functions.VariablesGlobales._Adicional_RET
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "County", "")
                        Else
                            sSucursal = Functions.VariablesGlobales._Adicional_RET
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
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        IdCol = "DocNum"
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""@TM_LE_RETVH"" where ""U_TM_CASRI"" = '" + oRE.AutorizacionSRI + "' and ""U_TM_STATUS""='Borrador' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from ""@TM_LE_RETVH"" WITH(NOLOCK) where U_TM_CASRI = '" + oRE.AutorizacionSRI + "' and U_TM_STATUS='Borrador' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        'U_FX_AUTO_RETENCION
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            '
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""OPDF"" where ""U_HBT_NUM_AUT"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from OPDF WITH(NOLOCK) where U_HBT_NUM_AUT = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        IdCol = "DocNum"
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_SS_AutRetRec"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_SS_AutRetRec = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    End If

                    Dim numretener As String

                    'Dim odetalle() As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion
                    'odetalle = oRE.ENTDetalleRetencion
                    'numretener = odetalle(0).NumDocRetener
                    Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                    WS.Url = Functions.VariablesGlobales._WS_Recepcion
                    Dim iIdDocEdoc As Long = Long.Parse(oRE.IdRetencion.ToString.ToString())

                    Dim odetalleRT As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                    odetalleRT = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)
                    If Not IsNothing(odetalleRT) Then
                        numretener = odetalleRT.ENTDetalleRetencion(0).NumDocRetener
                    End If
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
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Trim(Left(oRE.RazonSocial, 250)))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oRE.TotalRetencion))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oRE.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oRE.AutorizacionSRI)

                    'oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRetener", i, "")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRetener", i, IIf(String.IsNullOrEmpty(numretener), "0", numretener))

                    If Not String.IsNullOrEmpty(DocPreliminar) Then

                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("IdDoc", i, oRE.IdRetencion.ToString)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Seleccionar", i, "N")
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
            Utilitario.Util_Log.Escribir_Log("Error al llenar grid : " + ex.Message.ToString(), "frmDocumentoRecibido")
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
            FchAut = Functions.VariablesGlobales._MostrarFechaAutorizacion
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


            Dim NomCamAdi = Functions.VariablesGlobales._Nombre_CA

            oGrid.Columns.Item(13).Description = "sucursal"
            oGrid.Columns.Item(13).TitleObject.Caption = NomCamAdi '"Sucursal"
            oGrid.Columns.Item(13).Editable = False
            'oGrid.Columns.Item(10).ForeColor = ColorTranslator.ToOle(Color.White)
            oGrid.Columns.Item(13).Visible = True

            oGrid.Columns.Item(14).Description = "IdDoc"
            oGrid.Columns.Item(14).TitleObject.Caption = "IdDoc"
            oGrid.Columns.Item(14).Editable = False
            oGrid.Columns.Item(14).Visible = False

            oGrid.Columns.Item(15).Description = "Seleccionar"
            oGrid.Columns.Item(15).TitleObject.Caption = "Seleccionar"
            oGrid.Columns.Item(15).Editable = True
            oGrid.Columns.Item(15).Visible = True
            oGrid.Columns.Item(15).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            Dim oEditTextColumn2 As SAPbouiCOM.EditTextColumn
            oEditTextColumn2 = oGrid.Columns.Item(12)
            If TipoDocumento = "RE" Then ' SI ES RETENCION - EL PAGO RECIBIDO BORRADOR ES OTRA TABLA Y OTRO OBJTYPE
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    oEditTextColumn2.LinkedObjectType = "TM_RETV"
                Else
                    oEditTextColumn2.LinkedObjectType = 140
                End If

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

        If FormUID = "frmDocumentosRecibidos" Then
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")

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

                WS.Proxy = proxyobject
                WS.Credentials = cred
            End If
            ' END  MANEJO PROXY

            'If Functions.VariablesGlobales._vgHttps = "Y" Then
            '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
            'End If
            'oManejoDocumentos.SetProtocolosdeSeguridad()
            SetProtocolosdeSeguridad()
            If WS.MarcarVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                rsboApp.SetStatusBarMessage("Documento Marcado como Integrado : " + mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                Dim preliminar As String = oDataTable.GetValue(12, ofila).ToString()

                If Not preliminar = "0" Then


                    Dim idDocRec As String = ""
                    Dim QueryidDocRec As String = ""

                    If TipoDocumento = 1 Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QueryidDocRec = "SELECT ""DocEntry"" from ""@GS_FVR"" where ""U_IdGS"" ='" + IdDocumento.ToString + "'"
                        Else
                            QueryidDocRec = "SELECT DocEntry from ""@GS_FVR"" WITH(NOLOCK) where U_IdGS ='" + IdDocumento.ToString + "'"
                        End If
                        idDocRec = oFuncionesB1.getRSvalue(QueryidDocRec, "DocEntry", "")
                        ofrmDocumento.ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(idDocRec, 1)
                    ElseIf TipoDocumento = 3 Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QueryidDocRec = "SELECT ""DocEntry"" from ""@GS_NCR"" where ""U_IdGS"" ='" + IdDocumento.ToString + "'"
                        Else
                            QueryidDocRec = "SELECT DocEntry from ""@GS_NCR"" WITH(NOLOCK) where U_IdGS ='" + IdDocumento.ToString + "'"
                        End If
                        idDocRec = oFuncionesB1.getRSvalue(QueryidDocRec, "DocEntry", "")
                        ofrmDocumentoNC.ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(idDocRec, 1)

                    ElseIf TipoDocumento = 2 Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QueryidDocRec = "SELECT ""DocEntry"" from ""@GS_RER"" where ""U_IdGS"" ='" + IdDocumento.ToString + "'"
                        Else
                            QueryidDocRec = "SELECT DocEntry from ""@GS_RER"" WITH(NOLOCK) where U_IdGS ='" + IdDocumento.ToString + "'"
                        End If
                        idDocRec = oFuncionesB1.getRSvalue(QueryidDocRec, "DocEntry", "")
                        ofrmDocumentoRE.ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocRec, 1)
                    End If
                End If

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
            If eventInfo.FormUID = "frmDocumentosRecibidos" And eventInfo.ItemUID = "oGrid" Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                If eventInfo.BeforeAction Then
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    ofila = eventInfo.Row
                    Dim oFor As SAPbouiCOM.Form
                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidos")
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

                    Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
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

                    'If OC <> "0" Or Mapeado = "SI" Then ' se comento 20231114 debido a que no tiene ningua funcion es especifico
                    '    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    '    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    '    oCreationPackage.UniqueID = "CrearFactura"
                    '    oCreationPackage.String = "Crear Factura Preliminar..."
                    '    oCreationPackage.Enabled = True
                    '    oCreationPackage.Position = 21
                    '    oMenuItem = rsboApp.Menus.Item("1280")
                    '    oMenus = oMenuItem.SubMenus
                    '    oMenus.AddEx(oCreationPackage)
                    'End If
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
                    If typeEx = "frmDocumentosRecibidos" Then
                        If ofila > 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
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
                                                listaDetalleArtiulos.Add(New Entidades.DetalleArticulo(sClaveAcceso, oDetalle.CodigoPrincipal,
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
                    If typeEx = "frmDocumentosRecibidos" Then
                        If ofila >= 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                            Dim tipoDocumento As String = oDataTable.GetValue(0, ofila).ToString()
                            Dim QueryidDocUDO As String = ""
                            Dim idDocUDO As String = ""
                            Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                            WS.Url = Functions.VariablesGlobales._WS_Recepcion
                            Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                            sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + sRUC + "'", "CardCode", "")
                            If tipoDocumento = "Factura" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then

                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_FVR"" where ""U_IdGS"" ='" + oFac.IdFactura.ToString + "'"
                                        Else
                                            QueryidDocUDO = "SELECT DocEntry from ""@GS_FVR"" WITH(NOLOCK) where U_IdGS ='" + oFac.IdFactura.ToString + "'"
                                        End If
                                        idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")
                                        If idDocUDO <> "0" Then
                                            If ActEstadoMarcado_Factura(idDocUDO) Then
                                                If MarcarVisto(oFac.IdFactura, 1, mensaje) Then
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("FE")
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                End If
                                            End If
                                            'ofrmDocumento.MarcarVisto(oFac.IdFactura.ToString, 1, mensaje, idDocUDO)
                                            ''ofrmDocumento.MarcarVisto(idDocRec, 1, mensaje, oFac.IdFactura.ToString)
                                            'oDataTable.Rows.Remove(ofila)
                                            'CargaDocumentosFormato("FE")
                                            'rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                        Else
                                            If Guarda_DocumentoRecibido_Factura(oFac) Then
                                                If MarcarVisto(oFac.IdFactura, 1, mensaje) Then
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("FE")
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                Else
                                                    rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            ElseIf tipoDocumento = "Nota de Crédito" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                                results = listaNCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_NCR"" where ""U_IdGS"" ='" + oNC.IdNotaCredito.ToString + "'"
                                        Else
                                            QueryidDocUDO = "SELECT DocEntry from ""@GS_NCR"" WITH(NOLOCK) where U_IdGS ='" + oNC.IdNotaCredito.ToString + "'"
                                        End If
                                        idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")
                                        If idDocUDO <> "0" Then
                                            If ActEstadoMarcado_NotaCredito(idDocUDO) Then
                                                If MarcarVisto(oNC.IdNotaCredito, 3, mensaje) Then
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("NE")
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                End If
                                            End If
                                            'ofrmDocumentoNC.MarcarVisto(oNC.IdNotaCredito, 3, mensaje, idDocUDO)
                                            ''ofrmDocumentoNC.MarcarVisto(idDocUDO, 3, mensaje, oNC.IdNotaCredito)
                                            'oDataTable.Rows.Remove(ofila)
                                            'CargaDocumentosFormato("NE")
                                        Else
                                            If Guarda_DocumentoRecibido_NotaCredito(oNC) Then
                                                If MarcarVisto(oNC.IdNotaCredito, 3, mensaje) Then
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("NE")
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                Else
                                                    rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                End If
                                            End If
                                        End If
                                        'ofrmDocumentoNC.ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(idDocRec, 1)


                                    End If
                                Next
                            ElseIf tipoDocumento = "Retención de Cliente" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                                results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)

                                For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results

                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        SetProtocolosdeSeguridad()
                                        Dim _oRetencion = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, oRE.IdRetencion, mensaje)
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_RER"" where ""U_IdGS"" ='" + oRE.IdRetencion.ToString + "'"
                                        Else
                                            QueryidDocUDO = "SELECT DocEntry from ""@GS_RER"" WITH(NOLOCK) where U_IdGS ='" + oRE.IdRetencion.ToString + "'"
                                        End If
                                        idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")
                                        If idDocUDO <> "0" Then
                                            If ActEstadoMarcado_Retencion(idDocUDO) Then
                                                If MarcarVisto(oRE.IdRetencion, 2, mensaje) Then
                                                    If ValidarRetCero(idDocUDO.ToString, _oRetencion) Then
                                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Validando Retencion 0%..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                        If ListaDocEntryFact.Count > 0 And EsRetCero Then
                                                            If ActualizarFactura(ListaDocEntryFact, idDocUDO) Then
                                                                oDataTable.Rows.Remove(ofila)
                                                                CargaDocumentosFormato("RE")
                                                                rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                            End If

                                                        End If

                                                    End If
                                                End If

                                            End If

                                            If MarcarVisto(oRE.IdRetencion, 2, mensaje) Then
                                                If ListaDocEntryFact.Count > 0 And EsRetCero Then
                                                    If ActualizarFactura(ListaDocEntryFact, idDocUDO) Then

                                                    End If
                                                End If
                                                oDataTable.Rows.Remove(ofila)
                                                CargaDocumentosFormato("RE")
                                                rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                            End If
                                            'ofrmDocumentoRE.MarcarVisto(oRE.IdRetencion, 2, mensaje, idDocUDO)
                                            ''ofrmDocumentoNC.MarcarVisto(idDocUDO, 3, mensaje, oNC.IdNotaCredito)
                                            'oDataTable.Rows.Remove(ofila)
                                            'CargaDocumentosFormato("RE")
                                        Else
                                            Dim idudo As String = "0"
                                            If ValidarRetCero(idudo.ToString, _oRetencion) Then

                                                rsboApp.SetStatusBarMessage(NombreAddon + " - Validando Retencion 0%..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                            End If

                                            If Guarda_DocumentoRecibido_Retencion(idudo, _oRetencion) Then
                                                If MarcarVisto(oRE.IdRetencion, 2, mensaje) Then
                                                    If ListaDocEntryFact.Count > 0 And EsRetCero Then
                                                        ActualizarFactura(ListaDocEntryFact, idudo)
                                                    End If
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("RE")
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                Else
                                                    rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmDocumentosRecibidos")

        End Try

    End Sub

    Public Function Guarda_DocumentoRecibido_Factura(ByVal oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura) As Boolean

        'Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        'Dim claveAcceso As String = oDataTable.GetValue(5, ofila).ToString()
        'Dim claveAcceso As String = oFactura.ClaveAcceso.ToString
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim cardCode As String = ""
        Dim DocEntryFacturaRecibida_UDO As String = 0

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            cardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + oFactura.Ruc.ToString() + "'", "CardCode", "")
        Else
            cardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + oFactura.Ruc.ToString() + "'", "CardCode", "")
        End If
        If cardCode = "" Then
            cardCode = "N/A"
        End If
        Dim BaseImponibleIVA As Decimal = 0
        Dim BaseImponible0 As Decimal = 0
        Dim BaseImponibleNoObjeto As Decimal = 0
        Dim BaseImponibleExento As Decimal = 0
        Dim Iva As Decimal = 0
        Dim ICE As Decimal = 0
        Dim BaseImponibleICE As Decimal = 0
        For Each facImpuesto As Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto In oFactura.ENTFacturaImpuesto
            If facImpuesto.Codigo = 2 Then
                If facImpuesto.CodigoPorcentaje = 2 Or facImpuesto.CodigoPorcentaje = 3 Then
                    BaseImponibleIVA += facImpuesto.BaseImponible
                    Iva += facImpuesto.Valor
                ElseIf facImpuesto.CodigoPorcentaje = 0 Then
                    BaseImponible0 += facImpuesto.BaseImponible
                ElseIf facImpuesto.CodigoPorcentaje = 6 Then
                    BaseImponibleNoObjeto += facImpuesto.BaseImponible
                ElseIf facImpuesto.CodigoPorcentaje = 7 Then
                    BaseImponibleExento += facImpuesto.BaseImponible
                End If
            End If
        Next
        For Each facImpuesto As Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto In oFactura.ENTFacturaImpuesto
            If facImpuesto.Codigo = 3 Then
                '  BaseImponibleICE += facImpuesto.BaseImponible
                ICE += facImpuesto.Valor
            End If
        Next

        Try
            oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso.ToString(), "Registrando Factura Marcada", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("REE " + oFactura.ClaveAcceso.ToString() + " Registrando Factura Marcada", "frmDocumentosRecibidos")
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralData.SetProperty("U_RUC", oFactura.Ruc.ToString())
            'oGeneralData.SetProperty("U_Nombre", Left(oFactura.NombreComercial.ToString().Replace(Chr(10), ""), 99))
            oGeneralData.SetProperty("U_Nombre", Left(oFactura.RazonSocial.ToString().Replace(Chr(10), ""), 99))
            oGeneralData.SetProperty("U_CardCode", cardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", "")
            oGeneralData.SetProperty("U_ClaAcc", oFactura.ClaveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oFactura.AutorizacionSRI.ToString())
            oGeneralData.SetProperty("U_FecAut", oFactura.FechaAutorizacion.ToString())
            'oGeneralData.SetProperty("U_FechaS", Date.Now.ToString())
            oGeneralData.SetProperty("U_NumDoc", oFactura.Establecimiento.ToString() + "-" + oFactura.PuntoEmision.ToString() + "-" + oFactura.Secuencial.ToString())
            'oGeneralData.SetProperty("U_FPrelim", oForm.Items.Item("txtFPre").Specific.Value.ToString())
            oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleIVA, 2).ToString())))
            oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponible0, 2).ToString())))
            oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleNoObjeto, 2).ToString())))
            oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleExento, 2).ToString())))
            oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oFactura.TotalSinImpuesto, 2).ToString())))
            oGeneralData.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oFactura.TotalDescuento, 2).ToString())))
            oGeneralData.SetProperty("U_ICE", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleICE, 2).ToString())))
            oGeneralData.SetProperty("U_IVA", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(Iva, 2).ToString())))
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oFactura.ImporteTotal, 2).ToString())))
            oGeneralData.SetProperty("U_rTades", "0")
            oGeneralData.SetProperty("U_rPDesc", "0")
            oGeneralData.SetProperty("U_rDesc", "0")
            oGeneralData.SetProperty("U_rGast", "0")
            oGeneralData.SetProperty("U_rImp", "0")
            oGeneralData.SetProperty("U_rTotal", "0")
            oGeneralData.SetProperty("U_IdGS", oFactura.IdFactura.ToString())
            oGeneralData.SetProperty("U_Sincro", "0")
            oGeneralData.SetProperty("U_Tipo", "Factura de Servicio")
            oGeneralData.SetProperty("U_SincroE", "1")
            oGeneralData.SetProperty("U_Estado", "docMarcado")


            oChildren = oGeneralData.Child("GS0_FVR")
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            For Each facDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFactura.ENTDetalleFactura
                oChild = oChildren.Add
                Dim CodAux As String = ""
                Dim CodPrin As String = ""
                If facDetalle.CodigoAuxiliar = Nothing Then
                    CodAux = "N/A"
                Else
                    CodAux = facDetalle.CodigoAuxiliar.ToString()
                End If
                If facDetalle.CodigoPrincipal = Nothing Then
                    CodPrin = "N/A"
                Else
                    CodPrin = facDetalle.CodigoPrincipal.ToString()
                End If
                oChild.SetProperty("U_CodPrin", Left(CodPrin, 99))
                oChild.SetProperty("U_CodAuxi", CodAux)
                'oChild.SetProperty("U_CodSAP", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_Descripc", Left(facDetalle.Descripcion.ToString(), 100))
                oChild.SetProperty("U_Cantid", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle.Cantidad.ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle.PrecioUnitario.ToString())))
                oChild.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle.Descuento.ToString())))
                oChild.SetProperty("U_Total", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle.PrecioTotalSinImpuesto.ToString())))
            Next
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso.ToString(), "Se creo registro de Factura Marcada, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso.ToString(), "Ocurrior un error al crear registro de Factura marcada UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Factura marcada en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("REE " + oFactura.ClaveAcceso.ToString() + " Ocurrior un error al crear registro de Factura marcada UDO: " + ex.Message.ToString(), "frmDocumentosRecibidos")
            mensaje = ex.Message.ToString
            Return False
        End Try
    End Function

    Public Function ActEstadoMarcado_Factura(ByVal IdUdo As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        'Dim IdUdo = oFuncionesB1.getRSvalue("SELECT ""DocEntry"" FROM ""@GS_FVR"" where ""U_IdGS"" = 'S' AND ""LicTradNum"" = '" + oFactura.IdFactura.ToString() + "'", "DocEntry", "0")

        If Not IdUdo = "0" Then

            Try
                'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", IdUdo)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_Estado", "docMarcado")
                oGeneralData.SetProperty("U_Sincro", "0")

                oGeneralService.Update(oGeneralData)

                Return True
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al actualizar estado al preliminar generado " + ex.Message.ToString, "frmDocumentosRecibidos")
                Return False
            End Try

        End If


    End Function

    Public Function ActEstadoMarcado_NotaCredito(ByVal IdUdo As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        'Dim IdUdo = oFuncionesB1.getRSvalue("SELECT ""DocEntry"" FROM ""@GS_FVR"" where ""U_IdGS"" = 'S' AND ""LicTradNum"" = '" + oFactura.IdFactura.ToString() + "'", "DocEntry", "0")

        If Not IdUdo = "0" Then

            Try
                'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", IdUdo)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_Estado", "docMarcado")
                oGeneralData.SetProperty("U_Sincro", "0")

                oGeneralService.Update(oGeneralData)

                Return True
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al actualizar estado al preliminar generado " + ex.Message.ToString, "frmDocumentosRecibidos")
                Return False
            End Try

        End If


    End Function

    Public Function ActEstadoMarcado_Retencion(ByVal IdUdo As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        'Dim IdUdo = oFuncionesB1.getRSvalue("SELECT ""DocEntry"" FROM ""@GS_FVR"" where ""U_IdGS"" = 'S' AND ""LicTradNum"" = '" + oFactura.IdFactura.ToString() + "'", "DocEntry", "0")

        If Not IdUdo = "0" Then

            Try
                'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_RER")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", IdUdo)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_Estado", "docMarcado")
                oGeneralData.SetProperty("U_Sincro", "0")

                oGeneralService.Update(oGeneralData)

                Return True
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al actualizar estado al preliminar generado " + ex.Message.ToString, "frmDocumentosRecibidos")
                Return False
            End Try

        End If


    End Function

    Public Function Guarda_DocumentoRecibido_NotaCredito(ByVal oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito) As Boolean

        'Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        'Dim claveAcceso As String = oDataTable.GetValue(5, ofila).ToString()

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim cardCode As String = ""
        Dim DocEntryFacturaRecibida_UDO As String = 0

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            cardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + oNC.Ruc.ToString() + "'", "CardCode", "")
        Else
            cardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + oNC.Ruc.ToString() + "'", "CardCode", "")
        End If
        If cardCode = "" Then
            cardCode = "N/A"
        End If
        Dim BaseImponibleIVA As Decimal = 0
        Dim BaseImponible0 As Decimal = 0
        Dim BaseImponibleNoObjeto As Decimal = 0
        Dim BaseImponibleExento As Decimal = 0
        Dim Iva As Decimal = 0
        Dim ICE As Decimal = 0
        Dim BaseImponibleICE As Decimal = 0
        For Each NCImpuesto As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto In oNC.ENTNotaCreditoImpuesto
            If NCImpuesto.Codigo = 2 Then
                If NCImpuesto.CodigoPorcentaje = 2 Or NCImpuesto.CodigoPorcentaje = 3 Then
                    BaseImponibleIVA += NCImpuesto.BaseImponible
                    Iva += NCImpuesto.Valor
                ElseIf NCImpuesto.CodigoPorcentaje = 0 Then
                    BaseImponible0 += NCImpuesto.BaseImponible
                ElseIf NCImpuesto.CodigoPorcentaje = 6 Then
                    BaseImponibleNoObjeto += NCImpuesto.BaseImponible
                ElseIf NCImpuesto.CodigoPorcentaje = 7 Then
                    BaseImponibleExento += NCImpuesto.BaseImponible
                End If
            End If
        Next
        For Each NCImpuesto As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto In oNC.ENTNotaCreditoImpuesto
            If NCImpuesto.Codigo = 3 Then
                '  BaseImponibleICE += facImpuesto.BaseImponible
                ICE += NCImpuesto.Valor
            End If
        Next

        Try
            oFuncionesAddon.GuardaLOG("NCR", oNC.ClaveAcceso.ToString(), "Registrando NotaCredito Marcada", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralData.SetProperty("U_RUC", oNC.Ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oNC.RazonSocial.ToString().Replace(Chr(10), ""), 99))
            oGeneralData.SetProperty("U_CardCode", cardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", "")
            oGeneralData.SetProperty("U_ClaAcc", oNC.ClaveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oNC.AutorizacionSRI.ToString())
            oGeneralData.SetProperty("U_FecAut", oNC.FechaAutorizacion.ToString())
            'oGeneralData.SetProperty("U_FechaS", Date.Now.ToString())
            oGeneralData.SetProperty("U_NumDoc", oNC.Establecimiento.ToString() + "-" + oNC.PuntoEmision.ToString() + "-" + oNC.Secuencial.ToString())
            'oGeneralData.SetProperty("U_FPrelim", oForm.Items.Item("txtFPre").Specific.Value.ToString())
            oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleIVA, 2).ToString())))
            oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponible0, 2).ToString())))
            oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleNoObjeto, 2).ToString())))
            oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleExento, 2).ToString())))
            oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oNC.TotalSinImpuesto, 2).ToString())))
            oGeneralData.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oNC.Descuento, 2).ToString())))
            oGeneralData.SetProperty("U_ICE", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleICE, 2).ToString())))
            oGeneralData.SetProperty("U_IVA", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(Iva, 2).ToString())))
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oNC.ValorModificacion, 2).ToString())))
            oGeneralData.SetProperty("U_rTades", "0")
            oGeneralData.SetProperty("U_rPDesc", "0")
            oGeneralData.SetProperty("U_rDesc", "0")
            oGeneralData.SetProperty("U_rGast", "0")
            oGeneralData.SetProperty("U_rImp", "0")
            oGeneralData.SetProperty("U_rTotal", "0")
            oGeneralData.SetProperty("U_IdGS", oNC.IdNotaCredito.ToString())
            oGeneralData.SetProperty("U_Sincro", "0")
            oGeneralData.SetProperty("U_Tipo", "NC de Servicio")
            oGeneralData.SetProperty("U_SincroE", "1")
            oGeneralData.SetProperty("U_Estado", "docMarcado")


            oChildren = oGeneralData.Child("GS0_NCR")
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            For Each NCDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito In oNC.ENTDetalleNotaCredito
                oChild = oChildren.Add
                Dim CodAux As String = ""
                Dim CodPrin As String = ""
                If NCDetalle.CodigoAuxiliar = Nothing Then
                    CodAux = "N/A"
                Else
                    CodAux = NCDetalle.CodigoAuxiliar.ToString()
                End If
                If NCDetalle.CodigoPrincipal = Nothing Then
                    CodPrin = "N/A"
                Else
                    CodPrin = NCDetalle.CodigoPrincipal.ToString()
                End If
                oChild.SetProperty("U_CodPrin", Left(CodPrin, 99))
                oChild.SetProperty("U_CodAuxi", CodAux)
                'oChild.SetProperty("U_CodSAP", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_Descripc", Left(NCDetalle.Descripcion.ToString(), 100))
                oChild.SetProperty("U_Cantid", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle.Cantidad.ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle.PrecioUnitario.ToString())))
                oChild.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle.Descuento.ToString())))
                oChild.SetProperty("U_Total", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle.PrecioTotalSinImpuesto.ToString())))
            Next
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("REE", oNC.ClaveAcceso.ToString(), "Se creo registro de NotaCredito Marcada, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("REE", oNC.ClaveAcceso.ToString(), "Ocurrior un error al crear registro de nota de credito marcada UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar nota de credito marcada en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            mensaje = ex.Message.ToString
            Return False
        End Try
    End Function

    Public Function Guarda_DocumentoRecibido_Retencion(ByRef DocEntryFacturaRecibida_UDO As Integer, ByVal oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion) As Boolean
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService



        Try
            rsboApp.StatusBar.SetText(NombreAddon + "- Creando registro de Pago Recibido(Retencion) Recibida UDO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", oRE.ClaveAcceso, "Creando registro de Pago Recibido(Retencion) Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Try
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sCardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM OCRD where ""LicTradNum"" = '" + oRE.Ruc.ToString() + "' AND ""CardType"" = 'C' ", "CardCode", "")
                Else
                    sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + oRE.Ruc.ToString() + "' AND CardType = 'C' ", "CardCode", "")
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("REE - error al obtener CardCode DM : " + ex.ToString, "frmDocumentosRecibidos")
            End Try

            Dim numDoc As String = oRE.Establecimiento + "-" + oRE.PuntoEmision + "-" + oRE.Secuencial

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralData.SetProperty("Code", conta)
            oGeneralData.SetProperty("U_RUC", oRE.Ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oRE.RazonSocial.ToString(), 99))
            oGeneralData.SetProperty("U_CardCode", sCardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", oForm.Items.Item("lbMapp").Specific.Value.ToString())
            oGeneralData.SetProperty("U_ClaAcc", oRE.ClaveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oRE.AutorizacionSRI.ToString())
            oGeneralData.SetProperty("U_FecAut", oRE.FechaAutorizacion.ToString())
            oGeneralData.SetProperty("U_NumDoc", numDoc.ToString())
            'oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString())
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(frmDocumentoRE.formatDecimal(oRE.TotalRetencion.ToString())))
            oGeneralData.SetProperty("U_IdGS", oRE.IdRetencion.ToString())
            oGeneralData.SetProperty("U_Sincro", 0)
            oGeneralData.SetProperty("U_SincroE", 1)
            oGeneralData.SetProperty("U_Estado", "docMarcado")


            oChildren = oGeneralData.Child("GS0_RER")
            For Each detalleRet As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRE.ENTDetalleRetencion

                Dim ejeFiscal As String = oRE.PeriodoFiscal
                Dim _numDocRet As String = ""
                oChild = oChildren.Add
                oChild.SetProperty("U_CodRet", detalleRet.CodigoRetencion.ToString())
                If IsNothing(detalleRet.NumDocRetener) Then
                    _numDocRet = "0"
                Else
                    _numDocRet = detalleRet.NumDocRetener.ToString
                End If
                oChild.SetProperty("U_NumDocRe", _numDocRet)
                oChild.SetProperty("U_Fecha", detalleRet.FechaEmisionDocRetener.ToString())
                oChild.SetProperty("U_pFiscal", ejeFiscal.ToString())
                oChild.SetProperty("U_Base", Convert.ToDouble(frmDocumentoRE.formatDecimal(detalleRet.BaseImponible.ToString())))
                If detalleRet.Codigo = 1 Then
                    oChild.SetProperty("U_Impuesto", "RENTA")
                Else
                    oChild.SetProperty("U_Impuesto", "IVA")
                End If
                oChild.SetProperty("U_Porcent", Convert.ToDouble(frmDocumentoRE.formatDecimal(detalleRet.PorcentajeRetener.ToString())))
                'If detalleRet.PorcentajeRetener = 0 Then
                '    EsRetCero = True
                'End If
                oChild.SetProperty("U_valorR", Convert.ToDouble(frmDocumentoRE.formatDecimal(detalleRet.ValorRetenido.ToString())))

            Next
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("PRR", oRE.ClaveAcceso, "Se creo registro de Pago Recibido(Retencion) marcada UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Se creo registro de Pago Recibido(Retencion) marcada UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", oRE.ClaveAcceso, "Ocurrior un error al crear registro de Pago Recibido(Retencion) marcada UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Pago Recibido(Retencion) marcada en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function ValidarRetCero(ByRef DocEntryFacturaRecibida_UDO As String, ByVal oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion) As Boolean

        Try
            rsboApp.StatusBar.SetText(NombreAddon + "- Verificando Retencion 0%", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Factura = ""
            Dim contRetCero As Integer = 0
            Dim contExisteFact As Integer = 0
            'EsRetCero = False
            'If Not oRE.TotalRetencion = 0 Then

            '    EsRetCero = False
            '    Utilitario.Util_Log.Escribir_Log("No es retencion 0%" + oRE.ENTDetalleRetencion(0).CodDocRetener.ToString, "frmDocumentosRecibidos")
            '    Return False
            'Else
            '    EsRetCero = True
            'End If

            For Each detalleRet As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRE.ENTDetalleRetencion

                'If Not IsNothing(oRetencion) Then
                Dim numdocretenr = detalleRet.NumDocRetener
                Dim sQueryFactura = ""
                Dim sQueryFacturaAnt = ""
                If Not IsNothing(numdocretenr) Then
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND ""U_SER_EST"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND ""U_SER_PE"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND ""CANCELED""='N' "

                            sQueryFacturaAnt = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""ODPI"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ""ODPI"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND ""U_SER_EST"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND ""U_SER_PE"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND ""FolioNum"" = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND ""CANCELED""='N' "

                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(InstlmntID) from INV6 where DocEntry=OINV.DocEntry and Status='O') as ""Cuota"" FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND U_SER_EST = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND U_SER_PE = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND FolioNum = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND CANCELED='N' "

                            sQueryFacturaAnt = " SELECT DocEntry,(select min(InstlmntID) from INV6 where DocEntry=OINV.DocEntry and Status='O') as ""Cuota"" FROM ODPI WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND U_SER_EST = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND U_SER_PE = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND FolioNum = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND CANCELED='N' "

                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM ""OINV"" A INNER JOIN "
                            sQueryFactura += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND B.""BeginStr"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.""EndStr"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND A.""FolioNum"" =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "

                            sQueryFacturaAnt = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFacturaAnt += " FROM ""ODPI"" A INNER JOIN "
                            sQueryFacturaAnt += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFacturaAnt += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND B.""BeginStr"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND B.""EndStr"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND A.""FolioNum"" =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "

                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM OINV A INNER JOIN "
                            sQueryFactura += " NNM1 B ON A.Series = B.Series "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND B.BeginStr = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.EndStr = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND A.FolioNum =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "

                            sQueryFacturaAnt = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFacturaAnt += " FROM ODPI A INNER JOIN "
                            sQueryFacturaAnt += " NNM1 B ON A.Series = B.Series "
                            sQueryFacturaAnt += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND B.BeginStr = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND B.EndStr = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND A.FolioNum =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND A.DocStatus='O' AND A.CANCELED='N' "

                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM ""OINV"" A "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND right(A.""NumAtCard"",17) = '" & numdocretenr.ToString().Substring(0, 3) _
                                    & "-" & numdocretenr.ToString().Substring(3, 3) & "-" & numdocretenr.ToString().Substring(6, 9) & "'"
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "

                            sQueryFacturaAnt = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFacturaAnt += " FROM ""ODPI"" A "
                            sQueryFacturaAnt += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND right(A.""NumAtCard"",17) = '" & numdocretenr.ToString().Substring(0, 3) _
                                    & "-" & numdocretenr.ToString().Substring(3, 3) & "-" & numdocretenr.ToString().Substring(6, 9) & "'"
                            sQueryFacturaAnt += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "

                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM OINV A  "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND right(A.NumAtCard,17) = '" & numdocretenr.ToString().Substring(0, 3) _
                                    & "-" & numdocretenr.ToString().Substring(3, 3) & "-" & numdocretenr.ToString().Substring(6, 9) & "'"
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "

                            sQueryFacturaAnt = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFacturaAnt += " FROM ODPI A  "
                            sQueryFacturaAnt += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND right(A.NumAtCard,17) = '" & numdocretenr.ToString().Substring(0, 3) _
                                    & "-" & numdocretenr.ToString().Substring(3, 3) & "-" & numdocretenr.ToString().Substring(6, 9) & "'"
                            sQueryFacturaAnt += " AND A.DocStatus='O' AND A.CANCELED='N' "

                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM ""OINV"" A INNER JOIN "
                            sQueryFactura += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND B.""BeginStr"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.""EndStr"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.""DocNum"",7),'0',' ')),' ','0') =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "

                            sQueryFacturaAnt = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFacturaAnt += " FROM ""ODPI"" A INNER JOIN "
                            sQueryFacturaAnt += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFacturaAnt += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND B.""BeginStr"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND B.""EndStr"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.""DocNum"",7),'0',' ')),' ','0') =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "

                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFactura += " FROM OINV A INNER JOIN "
                            sQueryFactura += " NNM1 B ON A.Series = B.Series "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND B.BeginStr = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.EndStr = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.DocNum,7),'0',' ')),' ','0') =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "

                            sQueryFacturaAnt = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""A"".""DocEntry"" and ""Status""='O') as ""Cuota"" "
                            sQueryFacturaAnt += " FROM ODPI A INNER JOIN "
                            sQueryFacturaAnt += " NNM1 B ON A.Series = B.Series "
                            sQueryFacturaAnt += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND B.BeginStr = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND B.EndStr = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.DocNum,7),'0',' ')),' ','0') =" & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND A.DocStatus='O' AND A.CANCELED='N' "

                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND ""U_SS_Est"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND ""U_SS_Pemi"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "

                            sQueryFacturaAnt = " SELECT ""DocEntry"",(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ""ODPI"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND ""U_SS_Est"" = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND ""U_SS_Pemi"" = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND ""FolioNum"" = " & Integer.Parse(numdocretenr.Substring(6, 9))
                            sQueryFacturaAnt += " AND ""DocStatus""='O' AND ""CANCELED""='N' "
                        Else
                            sQueryFactura = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND U_SS_Est = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND U_SS_Pemi = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND FolioNum = " & Integer.Parse(numdocretenr.ToString().Substring(6, 9))
                            sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "

                            sQueryFacturaAnt = " SELECT DocEntry,(select min(""InstlmntID"") from INV6 where ""DocEntry""=""OINV"".""DocEntry"" and ""Status""='O') as ""Cuota"" FROM ODPI WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFacturaAnt += " AND U_SS_Est = '" & numdocretenr.ToString().Substring(0, 3) & "'"
                            sQueryFacturaAnt += " AND U_SS_Pemi = '" & numdocretenr.ToString().Substring(3, 3) & "'"
                            sQueryFacturaAnt += " AND FolioNum = " & Integer.Parse(numdocretenr.ToString().Substring(6, 9))
                            sQueryFacturaAnt += " AND DocStatus='O' AND CANCELED='N' "

                        End If
                    End If

                Else
                    contExisteFact = +1
                End If

                Try
                    Factura = oFuncionesB1.getRSvalue(sQueryFactura, "DocEntry", "")
                    Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + sQueryFactura.ToString + " Resultado: " + Factura.ToString, "frmDocumentosRecibidos")
                    Utilitario.Util_Log.Escribir_Log("Query Factura Anticipo Relacionada:" + sQueryFacturaAnt.ToString + " Resultado: " + Factura.ToString, "frmDocumentosRecibidos")
                    If Factura = "" Or Factura = "0" Then
                        contExisteFact = +1

                    Else
                        ListaDocEntryFact.Add(CInt(Factura))
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + ex.Message().ToString() + "-QUERY: " + sQueryFactura, "frmDocumentosRecibidos")
                End Try

                If detalleRet.ValorRetenido = 0 Then
                    contRetCero = +1
                End If

            Next

            If contRetCero > 0 Then

                EsRetCero = True

            End If

            If contExisteFact > 0 Then
                ExistFactura = False
            Else
                ExistFactura = True
            End If

            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al validar retencion 0%:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function ActualizarFactura(ByVal ListaDocEntrys As List(Of Integer), ByRef DocEntryFacturaRecibida_UDO As String) As Boolean

        Try

            Dim DocEntryFac As Integer = 0
            respN = 0
            respS = ""
            _oFactura = Nothing

            _oFactura = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            _oFactura.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
            _oFactura.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            For Each DocEntryFac In ListaDocEntrys.Distinct

                _oFactura.GetByKey(DocEntryFac)
                _oFactura.UserFields.Fields.Item("U_SS_IDRETCERO").Value = CInt(DocEntryFacturaRecibida_UDO)

                Dim resp As Integer = _oFactura.Update()
                If resp = 0 Then
                    Utilitario.Util_Log.Escribir_Log("Factura con DocEntry: " + DocEntryFac.ToString + "actualizado correctamente", "MarcadoEnLote")
                    'Return True
                Else
                    rCompany.GetLastError(respN, respS)
                    Utilitario.Util_Log.Escribir_Log("Factura con DocEntry: " + DocEntryFac.ToString + " no se pudo actualizar, codigo error: " + respN.ToString + " mensaje: " + respS, "MarcadoEnLote")
                    'Return False
                End If

            Next

            Return True

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("No se pudo actualizar id udo a factura, codigo error: " + respN.ToString + " mensaje: " + respS, "MarcadoEnLote")
            Return False
        End Try




    End Function
    Public Function ActualizadoEstadoMarcado_DocumentoRecibido(ByVal tipoDoc As String, ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            Utilitario.Util_Log.Escribir_Log("Actualizando estado docMarcado: " + tipoDoc.ToString, "ManejoDeDocumentos")
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", "docMarcado")

            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al actualizar el estado a docMarcado :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try
    End Function

    Public Function MarcarDocumentosContabilizadosManualFC(ByVal Lista As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura))
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
        WS.Url = _WS_RecepcionCambiarEstado
        Try
            For Each xoFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In Lista
                Dim query As String = "", rsult As String = ""
                'query = "select ""DocEntry"" from OPCH where ""Canceled""<>'Y' and ""U_NUM_AUTOR""='" + xoFactura.AutorizacionSRI.ToString + "'"
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    query = "select ""DocEntry"" from OPCH where ""CANCELED""='N' and ""U_NUM_AUTOR""='" + xoFactura.AutorizacionSRI.ToString + "'"
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    query = "select ""DocEntry"" from OPCH where ""CANCELED""='N' and ""U_SS_NumAut""='" + xoFactura.AutorizacionSRI.ToString + "'"
                End If
                rsult = oFuncionesB1.getRSvalue(query, "DocEntry", "")
                If Not rsult = "0" And Not rsult = "" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    SetProtocolosdeSeguridad()
                    If WS.MarcarVisto(Functions.VariablesGlobales._WS_RecepcionClave, xoFactura.IdFactura, 1, mensaje) Then
                        rsboApp.SetStatusBarMessage("La factura con # de autorizacion: " + xoFactura.AutorizacionSRI.ToString + " ya se encuentra contabilizada", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Else
                        rsboApp.SetStatusBarMessage("No se pudo marcar : " + mensaje.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    End If
                End If

            Next
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("No se pudo marcar fc: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
#Disable Warning BC42105 ' La función 'MarcarDocumentosContabilizadosManualFC' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function

    Public Function MarcarDocumentosContabilizadosManualNC(ByVal Lista As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito))
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
        WS.Url = _WS_RecepcionCambiarEstado
        Try
            For Each xoNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In Lista
                Dim query As String = "", rsult As String = ""
                'query = "select ""DocEntry"" from ORPC where ""Canceled""<>'Y' and ""U_NUM_AUTOR""='" + xoNC.AutorizacionSRI.ToString + "'"
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    query = "select ""DocEntry"" from ORPC where ""CANCELED""='N' and ""U_NUM_AUTOR""='" + xoNC.AutorizacionSRI.ToString + "'"
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    query = "select ""DocEntry"" from ORPC where ""CANCELED""='N' and ""U_SS_NumAut""='" + xoNC.AutorizacionSRI.ToString + "'"
                End If
                rsult = oFuncionesB1.getRSvalue(query, "DocEntry", "")
                If Not rsult = "0" And Not rsult = "" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    SetProtocolosdeSeguridad()
                    If WS.MarcarVisto(Functions.VariablesGlobales._WS_RecepcionClave, xoNC.IdNotaCredito, 3, mensaje) Then
                        rsboApp.SetStatusBarMessage("La NC con # de autorizacion: " + xoNC.AutorizacionSRI.ToString + " ya se encuentra contabilizada", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Else
                        rsboApp.SetStatusBarMessage("No se pudo marcar : " + mensaje.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    End If
                End If

            Next
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("No se pudo marcar nc: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
#Disable Warning BC42105 ' La función 'MarcarDocumentosContabilizadosManualNC' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function


    Public Function MarcarDocumentosContabilizadosManualRT(ByVal Lista As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
        WS.Url = _WS_RecepcionCambiarEstado
        Try
            For Each xoRT As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In Lista
                Dim query As String = "", rsult As String = ""
                'query = "select TOP 1 T0.""DocEntry"" from ORCT T0 INNER JOIN RCT3 T1 ON T1.""DocNum""=T0.""DocEntry"" WHERE T0.""Canceled""<>'Y' and T1.""U_CXS_NUM_AUTO_RETE""='" + xoRT.AutorizacionSRI.ToString + "'"
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    query = "select TOP 1 T0.""DocEntry"" from ORCT T0 INNER JOIN RCT3 T1 ON T1.""DocNum""=T0.""DocEntry"" WHERE T0.""Canceled""='N' and T1.""U_CXS_NUM_AUTO_RETE""='" + xoRT.AutorizacionSRI.ToString + "'"
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    query = "select TOP 1 T0.""DocEntry"" from ORCT T0 INNER JOIN RCT3 T1 ON T1.""DocNum""=T0.""DocEntry"" WHERE T0.""Canceled""='N' and T1.""U_SS_AutRetRec""='" + xoRT.AutorizacionSRI.ToString + "'"
                End If
                rsult = oFuncionesB1.getRSvalue(query, "DocEntry", "")
                If Not rsult = "0" And Not rsult = "" Then
                    'ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    SetProtocolosdeSeguridad()
                    If WS.MarcarVisto(Functions.VariablesGlobales._WS_RecepcionClave, xoRT.IdRetencion, 2, mensaje) Then
                        rsboApp.SetStatusBarMessage("La RT con # de autorizacion: " + xoRT.AutorizacionSRI.ToString + " ya se encuentra contabilizada", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Else
                        rsboApp.SetStatusBarMessage("No se pudo marcar : " + mensaje.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    End If
                End If

            Next
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("No se pudo marcar RT: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
#Disable Warning BC42105 ' La función 'MarcarDocumentosContabilizadosManualRT' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'MarcarDocumentosContabilizadosManualRT' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

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

    Shared Function customCertValidation(ByVal sender As Object,
                                             ByVal cert As X509Certificate,
                                             ByVal chain As X509Chain,
                                             ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function

    Public Sub SetProtocolosdeSeguridad()



        'PARA TLS 1.2
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999



        'PARA HTTPS



        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)



    End Sub

    Private Function SeleccionarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                'estado = oGridDet.GetValue("EstadoDoc", i)
                '  If estado = "AUTORIZADA" Then
                oGridDet.SetValue("Seleccionar", i, "Y")
                If Not ListaFila.Contains(i) Then
                    ListaFila.Add(i)
                End If
                contador += 1
                resul = True
                ' End If
            Next
            oForm.Freeze(False)
            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmDocumentosRecibidos")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al SeleccionarDocumentosPendientes:" + ex.Message().ToString(), "frmDocumentosRecibidos")
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function

    Private Function DesmarcarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                'estado = oGridDet.GetValue("EstadoDoc", i)
                '  If estado = "AUTORIZADA" Then
                oGridDet.SetValue("Seleccionar", i, "N")
                If ListaFila.Contains(i) Then
                    ListaFila.Remove(i)
                End If
                contador += 1
                resul = True
                '  End If
            Next
            oForm.Freeze(False)
            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmDocumentosRecibidos")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al SeleccionarDocumentosPendientes:" + ex.Message().ToString(), "frmDocumentosRecibidos")
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function

    Private Function MarcarDocIntegradoEnLote() As Boolean
        Dim resul As Boolean = False
        Try

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
            WS.Url = Functions.VariablesGlobales._WS_Recepcion
            ' Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim filasAEliminar As New List(Of Integer)
            Dim SN_Folio As New List(Of String)
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")
            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim IdEdocDoc As String = ""
            Dim clave As String = ""
            Dim Seleccionar As String = ""
            Dim Folio As String = ""
            Dim SN As String = ""
            Dim filasss As Integer = 0
            Dim cbxTipo As SAPbouiCOM.ComboBox
            cbxTipo = oForm.Items.Item("cbxTipo").Specific

            Dim k As Integer = 0
            If ListaFila.Count > 0 Then
                ListaFila.Sort()
            End If

            If ListaFila.Count > 0 Then

                'For k As Integer = 0 To oGridDet.Rows.Count - 1
                For Each k In ListaFila

                    Seleccionar = oGridDet.GetValue("Seleccionar", k)
                    IdEdocDoc = oGridDet.GetValue("ClaveAcceso", k)
                    Dim idDocUDO As String = ""
                    Dim QueryidDocUDO As String = ""

                    If Seleccionar = "Y" Then

                        IdEdocDoc = oGridDet.GetValue("IdDoc", k)
                        clave = oGridDet.GetValue("ClaveAcceso", k)
                        SN = oGridDet.GetValue("RazonSocial", k)
                        Folio = oGridDet.GetValue("Folio", k)

                        If cbxTipo.Value = "01" Then

                            Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                            results = listaFCs.FindAll(Function(column) column.ClaveAcceso = clave)

                            For Each FacMarcada As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results

                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_FVR"" where ""U_IdGS"" ='" + FacMarcada.IdFactura.ToString + "'"
                                Else
                                    QueryidDocUDO = "SELECT DocEntry from ""@GS_FVR"" WITH(NOLOCK) where U_IdGS ='" + FacMarcada.IdFactura.ToString + "'"
                                End If
                                idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")

                                If idDocUDO <> "0" Then
                                    If ActEstadoMarcado_Factura(idDocUDO) Then
                                        If MarcarVisto(FacMarcada.IdFactura, 1, mensaje) Then
                                            filasAEliminar.Add(k)
                                            oGridDet.SetValue("Seleccionar", k, "N")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Folio + " Estado actuializado y Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    End If
                                Else
                                    If Guarda_DocumentoRecibido_Factura(FacMarcada) Then
                                        If MarcarVisto(FacMarcada.IdFactura, 1, mensaje) Then
                                            filasAEliminar.Add(k)
                                            oGridDet.SetValue("Seleccionar", k, "N")
                                            'SN_Folio.Add(SN + ";" + Folio)
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Folio + " Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    End If
                                End If


                            Next

                        ElseIf cbxTipo.Value = "04" Then

                            Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                            results = listaNCs.FindAll(Function(column) column.ClaveAcceso = clave)

                            For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results

                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_NCR"" where ""U_IdGS"" ='" + oNC.IdNotaCredito.ToString + "'"
                                Else
                                    QueryidDocUDO = "SELECT DocEntry from ""@GS_NCR"" WITH(NOLOCK) where U_IdGS ='" + oNC.IdNotaCredito.ToString + "'"
                                End If
                                idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")

                                If idDocUDO <> "0" Then

                                    If ActEstadoMarcado_NotaCredito(idDocUDO) Then
                                        If MarcarVisto(oNC.IdNotaCredito, 3, mensaje) Then
                                            filasAEliminar.Add(k)
                                            oGridDet.SetValue("Seleccionar", k, "N")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Folio + " Estado actualizado y Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    End If

                                Else
                                    If Guarda_DocumentoRecibido_NotaCredito(oNC) Then
                                        If MarcarVisto(oNC.IdNotaCredito, 3, mensaje) Then
                                            filasAEliminar.Add(k)
                                            oGridDet.SetValue("Seleccionar", k, "N")
                                            'SN_Folio.Add(SN + ";" + Folio)
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Folio + " Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    End If
                                End If


                            Next

                        ElseIf cbxTipo.Value = "07" Then

                            Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                            results = listaREs.FindAll(Function(column) column.ClaveAcceso = clave)
                            Dim DocEntryFacturaRecibida_UDO As Integer = 0
                            Dim sRUC As String = oGridDet.GetValue(4, ofila).ToString()
                            sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + sRUC + "'", "CardCode", "")
                            Factura = ""
                            For Each oRt As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                                EsRetCero = False
                                ExistFactura = False
                                SetProtocolosdeSeguridad()
                                Dim _oRetencion = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, oRt.IdRetencion, mensaje)

                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_RER"" where ""U_IdGS"" ='" + _oRetencion.IdRetencion.ToString + "'"
                                Else
                                    QueryidDocUDO = "SELECT DocEntry from ""@GS_RER"" WITH(NOLOCK) where U_IdGS ='" + _oRetencion.IdRetencion.ToString + "'"
                                End If
                                idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")


                                If idDocUDO <> "0" Then

                                    If ActEstadoMarcado_Retencion(idDocUDO) Then
                                        If ValidarRetCero(DocEntryFacturaRecibida_UDO.ToString, _oRetencion) Then
                                            If ListaDocEntryFact.Count > 0 And EsRetCero Then
                                                If ActualizarFactura(ListaDocEntryFact, idDocUDO) Then
                                                    rsboApp.SetStatusBarMessage(NombreAddon + "Facturas con Ret 0% actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                End If
                                            End If

                                        End If

                                        If MarcarVisto(_oRetencion.IdRetencion, 2, mensaje) Then
                                            filasAEliminar.Add(k)
                                            oGridDet.SetValue("Seleccionar", k, "N")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Folio + " Estado actualizado y Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                        End If
                                    End If

                                Else

                                    If Guarda_DocumentoRecibido_Retencion(DocEntryFacturaRecibida_UDO, _oRetencion) Then
                                        If ValidarRetCero(DocEntryFacturaRecibida_UDO.ToString, _oRetencion) Then
                                            'If ExistFactura And EsRetCero Then
                                            If ListaDocEntryFact.Count > 0 And EsRetCero Then
                                                If ActualizarFactura(ListaDocEntryFact, DocEntryFacturaRecibida_UDO) Then
                                                    rsboApp.SetStatusBarMessage(NombreAddon + "Facturas con Ret 0% actualizada correctamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                End If
                                            End If

                                        End If

                                        If MarcarVisto(_oRetencion.IdRetencion, 2, mensaje) Then
                                            filasAEliminar.Add(k)
                                            oGridDet.SetValue("Seleccionar", k, "N")
                                            'SN_Folio.Add(SN + ";" + Folio)
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Folio + " Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If


                                    End If

                                End If
                                ListaDocEntryFact.Clear()

                            Next

                        End If

                    End If
                Next

            End If

            ' Eliminar las filas seleccionadas
            Dim rowsToDelete As List(Of Integer) = filasAEliminar
            rowsToDelete.Sort()
            rowsToDelete.Reverse()

            For Each rowIndex As Integer In rowsToDelete
                If rowIndex >= 0 And rowIndex < oGridDet.Rows.Count Then
                    oGridDet.Rows.Remove(rowIndex)
                End If
            Next

            If cbxTipo.Value = "01" Then
                CargaDocumentosFormato("FE") 'se reordena las filas por eso tengo que validar otros datos como sn y folio ya que al eliminar las filas se cambian de orden
            ElseIf cbxTipo.Value = "04" Then
                CargaDocumentosFormato("NE")
            ElseIf cbxTipo.Value = "07" Then
                CargaDocumentosFormato("RE")
            End If

            'For j As Integer = 0 To SN_Folio.Count - 1
            '    Dim ValoresLista = SN_Folio(j)
            '    For k As Integer = 0 To oGridDet.Rows.Count - 1
            '        SN = oGridDet.GetValue("RazonSocial", k)
            '        Folio = oGridDet.GetValue("Folio", k)
            '        If SN = ValoresLista.Split(";")(0) And Folio = ValoresLista.Split(";")(1) Then
            '            oGridDet.Rows.Remove(k)
            '            If cbxTipo.Value = "01" Then
            '                CargaDocumentosFormato("FE") 'se reordena las filas por eso tengo que validar otros datos como sn y folio ya que al eliminar las filas se cambian de orden
            '            ElseIf cbxTipo.Value = "04" Then
            '                CargaDocumentosFormato("NE")
            '            ElseIf cbxTipo.Value = "07" Then
            '                CargaDocumentosFormato("RE")
            '            End If
            '            oForm.Freeze(False)
            '            Exit For
            '        End If

            '    Next
            'Next
            Dim BtnSeleccionar = oForm.Items.Item("obtSel").Specific
            If BtnSeleccionar.Caption = "Desmarcar Todo" Then
                BtnSeleccionar.Caption = "Seleccionar Todo"
            End If
            filasAEliminar.Clear()
            ListaDocEntryFact.Clear()
            ListaFila.Clear()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al Marcar los documentos pendientes:" + ex.Message().ToString(), "frmDocumentosRecibidos")
            resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function



    Private Sub AgregarFilaXMLGrid(xobjeto As Object, tipoDoc As String, docPreliminar As String)
        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidos")


        Try

            oForm.Freeze(True)

            oForm.DataSources.DataTables.Item("dtDocs").Rows.Add()
            Dim i As Integer = oForm.DataSources.DataTables.Item("dtDocs").Rows.Count - 1
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, tipoDoc)
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, xobjeto.FechaEmision)

            oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, xobjeto.FechaAutorizacion)

            oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, xobjeto.Establecimiento + "-" + xobjeto.PuntoEmision + "-" + xobjeto.Secuencial)
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, xobjeto.Ruc)
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(xobjeto.RazonSocial, 250).Trim)

            If tipoDoc = "Nota de Crédito" Then
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(xobjeto.ValorModificacion))
            ElseIf tipoDoc = "Retención de Cliente" Then
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(xobjeto.TotalRetencion))
            Else
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(xobjeto.ImporteTotal))
            End If


            oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, xobjeto.ClaveAcceso)
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, xobjeto.AutorizacionSRI)

            If tipoDoc = "Retención de Cliente" Then

                Dim numretener As String = ""

                If Not IsNothing(xobjeto.ENTDetalleRetencion) Then
                    numretener = xobjeto.ENTDetalleRetencion(0).NumDocRetener

                End If

                oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRetener", i, IIf(String.IsNullOrEmpty(numretener), "0", numretener))


            End If


            If Not String.IsNullOrEmpty(docPreliminar) Then
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, docPreliminar)
            End If

            oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, "")
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("IdDoc", i, xobjeto.ClaveAcceso)
            oForm.DataSources.DataTables.Item("dtDocs").SetValue("Seleccionar", i, "N")
        Catch ex As Exception

        Finally

            oForm.Freeze(False)

        End Try


    End Sub

    Private Function CopiarDocumentos(ArchivoXML As String) As Boolean

        Try

            If File.Exists(ArchivoXML) Then

                Dim archivoPDF As String = ArchivoXML.ToLower.Replace(".xml", ".pdf")


                File.Copy(ArchivoXML, Functions.VariablesGlobales._RutaIntegracionXML & "\" & Path.GetFileName(ArchivoXML), True)

                File.Copy(archivoPDF, Functions.VariablesGlobales._RutaIntegracionXML & "\" & Path.GetFileName(archivoPDF), True)

                Return True

            End If


        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("No se Pudo el Archivo " & ArchivoXML, "frmDocumentosRecibidos")

        End Try

        Return False
    End Function

    Private Sub CargarXML()

        Dim selectFileDialog As New SelectFileDialog("C:\", "", "|*.xml", DialogType.OPEN)
        selectFileDialog.PermitirMultiFile = True
        selectFileDialog.Open()


        If selectFileDialog.SelectedFiles.Count > 0 Then

            For Each _rta As String In selectFileDialog.SelectedFiles

                Try
                    ' oForm.Items.Item("txtb64cer").Specific.Value = Convert.ToBase64String(File.ReadAllBytes(selectFileDialog.SelectedFile))

                    Dim rutaXML = _rta

                    If File.Exists(rutaXML) Then
                        rsboApp.SetStatusBarMessage("Cargando El documento: " & rutaXML,, False)

                        Dim todoXML As String = File.ReadAllText(rutaXML, System.Text.Encoding.UTF8)

                        If todoXML.Contains("infoFactura") Then

                            rsboApp.SetStatusBarMessage("Cargando Factura " & rutaXML,, False)

                            Dim xdoc As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura = Negocio.ssXML.OperacionesXML.LeerXMLFactura2(rutaXML)

                            If Not IsNothing(xdoc) Then

                                If (listaFCs.Where(Function(f) f.ClaveAcceso = xdoc.ClaveAcceso).Count <= 0) Then

                                    listaFCs.Add(xdoc)

                                    AgregarFilaXMLGrid(xdoc, "Factura", "")


                                    CopiarDocumentos(rutaXML)

                                End If


                            End If


                        ElseIf todoXML.Contains("infoNotaCredito") Then

                            rsboApp.SetStatusBarMessage("Cargando Nota Credito " & rutaXML,, False)

                            Dim xdoc As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito = Negocio.ssXML.OperacionesXML.LeerXMLNotaCredito2(rutaXML)

                            If Not IsNothing(xdoc) Then

                                If (listaNCs.Where(Function(f) f.ClaveAcceso = xdoc.ClaveAcceso).Count <= 0) Then

                                    listaNCs.Add(xdoc)

                                    AgregarFilaXMLGrid(xdoc, "Nota de Crédito", "")

                                    CopiarDocumentos(rutaXML)

                                End If

                            End If

                        ElseIf todoXML.Contains("infoCompRetencion") Then

                            rsboApp.SetStatusBarMessage("Cargando Retencion " & rutaXML,, False)

                            Dim xdoc As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion = Negocio.ssXML.OperacionesXML.LeerXMLRetencion2(rutaXML)

                            If Not IsNothing(xdoc) Then

                                If (listaREs.Where(Function(f) f.ClaveAcceso = xdoc.ClaveAcceso).Count <= 0) Then

                                    listaREs.Add(xdoc)

                                    AgregarFilaXMLGrid(xdoc, "Retención de Cliente", "")

                                    CopiarDocumentos(rutaXML)

                                End If

                            End If

                        Else

                            rsboApp.SetStatusBarMessage("No se Detecto el tipo de Documento " & rutaXML)

                        End If

                    Else

                        rsboApp.SetStatusBarMessage("No se encontro la Ruta " & rutaXML)

                    End If


                Catch ex As Exception

                    rsboApp.SetStatusBarMessage($"Error al Cargar el Archivo  XML {_rta} , ex: {ex.Message}")

                End Try


            Next


        End If








    End Sub




End Class

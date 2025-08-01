Imports System.IO
Imports System.Threading
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Linq

'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Public Class frmDocumentoREXML
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable
    Dim oCardCode As String = ""
    Dim _oDocumento As Retencion
    Public listaDetalleArtiulos As New List(Of Entidades.DetalleArticulo)
    Dim _fila As Integer
    Dim ObjTypeRelacionado As Integer = 0
    Dim oDocumentoSAP As SAPbobsCOM.Payments

    Dim oDocumentoSAPCB As SAPbobsCOM.Payments

    Dim _WS_Recepcion As String = ""
    Dim _WS_RecepcionCambiarEstado As String = ""
    Dim _WS_RecepcionClave As String = ""
    Dim _WS_RecepcionArchivo As String = ""
    Dim _ClaveAcceso As String = ""
    Dim _IdGS As Long = 0
    Dim _TipoDocumento As String = ""
    Dim CargaFacturaRelacionadas As Boolean = False
    Dim oListaFacturaVenta As New List(Of Entidades.FacturaVenta)
    Dim oFacturaVenta As Entidades.FacturaVenta
    Dim oGroupFolder As SAPbouiCOM.Item

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential

    Public vPay2 As SAPbobsCOM.IPayments



    Dim _sRUC As String = ""

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioDocumento(sRUC As String, sCardCode As String, sNombre As String, oDocumento As Retencion, ofila As Integer)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmDocumentoREXML") Then
            Exit Sub
        End If
        oCardCode = sCardCode
        _fila = ofila
        _sRUC = sRUC
        oListaFacturaVenta.Clear()

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentoREXML.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
                _oDocumento = oDocumento
                _ClaveAcceso = oDocumento.RetCabecera._claveAcceso
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentoREXML").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            ' RECUPERAR PARAMETRO, "RELACIONAR FACTURAS DE CLIENTES"
            Dim BFR As String = ""
            BFR = Functions.VariablesGlobales._BFR
            If BFR = "Y" Then
                CargaFacturaRelacionadas = True
            Else
                CargaFacturaRelacionadas = False
            End If
            'END RECUPERAR PARAMETRO, "RELACIONAR FACTURAS DE CLIENTES"

            oForm = rsboApp.Forms.Item("frmDocumentoREXML")

            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO
            oForm.Freeze(True)
            '

            ' CHECK "RELACIONAR FACTURAS DE CLIENTES"
            Dim chkDesc As SAPbouiCOM.CheckBox
            chkDesc = oForm.Items.Item("chkPa").Specific
            chkDesc.Item.Visible = True
            oForm.DataSources.UserDataSources.Add("chkPa", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkDesc.ValOn = "Y"
            chkDesc.ValOff = "N"
            chkDesc.DataBind.SetBound(True, "", "chkPa")
            chkDesc.Checked = CargaFacturaRelacionadas
            ' END CHECK "RELACIONAR FACTURAS DE CLIENTES"

            oForm.Items.Item("objR").Visible = False ' Guardo el ObjType del documento relacionado, lo lleno desde frm consultaordenes
            oForm.Items.Item("docR").Visible = False ' Guardo los docEntrys de los documentos relacionados, lo lleno desde frm consultaordenes

            _IdGS = _oDocumento.RetCabecera._DocEntry

            oForm.Items.Item("txtRUC").Enabled = False
            Dim txtRUC As SAPbouiCOM.EditText
            txtRUC = oForm.Items.Item("txtRUC").Specific
            txtRUC.Value = sRUC

            oForm.Items.Item("txtNombre").Enabled = False
            Dim txtNommbre As SAPbouiCOM.EditText
            txtNommbre = oForm.Items.Item("txtNombre").Specific
            txtNommbre.Value = sNombre

            oForm.Items.Item("txtF").Enabled = True
            Dim txtF As SAPbouiCOM.EditText
            txtF = oForm.Items.Item("txtF").Specific
            'txtF.Item.RightJustified = True
            txtF.Value = "0"

            'oForm.Items.Item("txtCodigo").Enabled = False
            Dim txtCodigo As SAPbouiCOM.EditText
            txtCodigo = oForm.Items.Item("txtCodigo").Specific
            txtCodigo.Value = sCardCode
            'txtCodigo.Item.Enabled = False
            Dim lnkCuentCN As SAPbouiCOM.LinkedButton
            lnkCuentCN = oForm.Items.Item("lnkCuentC").Specific
            lnkCuentCN.LinkedObjectType = 2
            lnkCuentCN.Item.LinkTo = "txtCodigo"

            oForm.Items.Item("txtClaAcc").Enabled = False
            Dim txtClaAcc As SAPbouiCOM.EditText
            txtClaAcc = oForm.Items.Item("txtClaAcc").Specific
            txtClaAcc.Value = _oDocumento.RetCabecera._claveAcceso

            oForm.Items.Item("txtNumAut").Enabled = False
            Dim txtNumAut As SAPbouiCOM.EditText
            txtNumAut = oForm.Items.Item("txtNumAut").Specific
            txtNumAut.Value = _oDocumento.RetCabecera._NumeroAutorizacion

            oForm.Items.Item("txtFecAut").Enabled = False
            Dim txtFecAut As SAPbouiCOM.EditText
            txtFecAut = oForm.Items.Item("txtFecAut").Specific
            txtFecAut.Value = _oDocumento.RetCabecera._FechaAutorizacion

            oForm.Items.Item("txtNumDoc").Enabled = False
            Dim txtNumDoc As SAPbouiCOM.EditText
            txtNumDoc = oForm.Items.Item("txtNumDoc").Specific
            txtNumDoc.Value = _oDocumento.RetCabecera._estab + "-" + _oDocumento.RetCabecera._ptoEmi + "-" + _oDocumento.RetCabecera._secuencial

            Dim lbEstAut As SAPbouiCOM.StaticText
            lbEstAut = oForm.Items.Item("lbEstAut").Specific
            lbEstAut.Caption = "DOCUMENTO AUTORIZADO POR EL SRI"
            lbEstAut.Item.ForeColor = RGB(7, 118, 10)

            Dim lnkPDF As SAPbouiCOM.LinkedButton
            lnkPDF = oForm.Items.Item("lnkPDF").Specific
            lnkPDF.Item.Visible = True
            Dim lnkXML As SAPbouiCOM.LinkedButton
            lnkXML = oForm.Items.Item("lnkXML").Specific
            lnkXML.Item.Visible = False
            Dim Item_16 As SAPbouiCOM.StaticText
            Item_16 = oForm.Items.Item("Item_16").Specific
            Item_16.Item.Visible = True
            Dim Item_18 As SAPbouiCOM.StaticText
            Item_18 = oForm.Items.Item("Item_18").Specific
            Item_18.Item.Visible = False

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodRet", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            'oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Cod", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("NumDocRe", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Fecha", SAPbouiCOM.BoFieldsType.ft_Date, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("pFiscal", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Base", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Impuesto", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Codigo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Porcent", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("valorR", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Factura", SAPbouiCOM.BoFieldsType.ft_Integer, 25)
            'If CargaFacturaRelacionadas Then
            '    oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Factura", SAPbouiCOM.BoFieldsType.ft_Integer, 25)
            'End If

            Dim PendienteMapear As Boolean = False

            oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(oDocumento.RetDetalleImp.Count)
            Dim i As Integer = 0
            Dim TotalValorRetenido As Decimal = 0
            For Each odetalle As RetDetalleImpuestos In oDocumento.RetDetalleImp

                If IIf(IsNothing(odetalle._codDocSustento), "", odetalle._codDocSustento).ToString().Length = 0 Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodRet", i, "")
                Else
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodRet", i, IIf(IsNothing(IIf(odetalle._codDocSustento.ToString().Equals("01"), "FACTURA", "")), "", IIf(odetalle._codDocSustento.ToString().Equals("01"), "FACTURA", "")))
                End If

                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Cod", i, odetalle.Codigo.ToString())
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRe", i, IIf(IsNothing(odetalle._numDocSustento), "", odetalle._numDocSustento))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, IIf(IsNothing(odetalle._fechaEmisionDocSustento), "", Convert.ToDateTime(odetalle._fechaEmisionDocSustento)))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("pFiscal", i, _oDocumento.RetCabecera._periodoFiscal)
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Base", i, Convert.ToDouble(odetalle._baseImponible))
                If odetalle._codigo = 1 Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Impuesto", i, "RENTA")
                ElseIf odetalle._codigo = 2 Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Impuesto", i, "IVA")
                ElseIf odetalle._codigo = 6 Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Impuesto", i, "ISD")
                End If
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Codigo", i, odetalle._codigoRetencion.ToString())

                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Porcent", i, Convert.ToDouble(odetalle._porcentajeRetener))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("valorR", i, Convert.ToDouble(odetalle._valorRetenido))

                If Not IsNothing(odetalle._numDocSustento) Then
                    'If CargaFacturaRelacionadas Then
                    Dim sQueryFactura As String = ""
                    Dim Factura As String = ""

                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND ""U_SER_EST"" = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND ""U_SER_PE"" = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "
                        Else
                            sQueryFactura = " SELECT DocEntry FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND U_SER_EST = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND U_SER_PE = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND FolioNum = " & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"" "
                            sQueryFactura += " FROM ""OINV"" A INNER JOIN "
                            sQueryFactura += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND B.""BeginStr"" = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.""EndStr"" = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND A.""FolioNum"" =" & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "
                        Else
                            sQueryFactura = " SELECT DocEntry "
                            sQueryFactura += " FROM OINV A INNER JOIN "
                            sQueryFactura += " NNM1 B ON A.Series = B.Series "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND B.BeginStr = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.EndStr = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND A.FolioNum =" & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "

                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"" "
                            sQueryFactura += " FROM ""OINV"" A "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND right(A.""NumAtCard"",17) = '" & odetalle._numDocSustento.ToString().Substring(0, 3) _
                                & "-" & odetalle._numDocSustento.ToString().Substring(3, 3) & "-" & odetalle._numDocSustento.ToString().Substring(6, 9) & "'"
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "
                        Else
                            sQueryFactura = " SELECT DocEntry "
                            sQueryFactura += " FROM OINV A  "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND right(A.NumAtCard,17) = '" & odetalle._numDocSustento.ToString().Substring(0, 3) _
                                & "-" & odetalle._numDocSustento.ToString().Substring(3, 3) & "-" & odetalle._numDocSustento.ToString().Substring(6, 9) & "'"
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"" "
                            sQueryFactura += " FROM ""OINV"" A INNER JOIN "
                            sQueryFactura += " ""NNM1"" B ON A.""Series"" = B.""Series"" "
                            sQueryFactura += " WHERE A.""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND B.""BeginStr"" = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.""EndStr"" = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.""DocNum"",7),'0',' ')),' ','0') =" & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND A.""DocStatus""='O' AND A.""CANCELED""='N' "
                        Else
                            sQueryFactura = " SELECT DocEntry "
                            sQueryFactura += " FROM OINV A INNER JOIN "
                            sQueryFactura += " NNM1 B ON A.Series = B.Series "
                            sQueryFactura += " WHERE A.CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND B.BeginStr = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND B.EndStr = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND REPLACE(LTRIM(REPLACE(RIGHT(A.DocNum,7),'0',' ')),' ','0') =" & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND A.DocStatus='O' AND A.CANCELED='N' "

                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQueryFactura = " SELECT ""DocEntry"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                            sQueryFactura += " AND ""U_SS_Est"" = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND ""U_SS_Pemi"" = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "
                        Else
                            sQueryFactura = " SELECT DocEntry FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                            sQueryFactura += " AND U_SS_Est = '" & odetalle._numDocSustento.ToString().Substring(0, 3) & "'"
                            sQueryFactura += " AND U_SS_Pemi = '" & odetalle._numDocSustento.ToString().Substring(3, 3) & "'"
                            sQueryFactura += " AND FolioNum = " & Integer.Parse(odetalle._numDocSustento.Substring(6, 9))
                            sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "
                        End If
                    End If
                    Try
                        Factura = oFuncionesB1.getRSvalue(sQueryFactura, "DocEntry", "")
                        Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + sQueryFactura.ToString + " Resultado: " + Factura.ToString, "frmDocumentoREXML")
                        If Factura = "" Then
                            Factura = "0"
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + ex.Message().ToString() + "-QUERY: " + sQueryFactura, "frmDocumentoREXML")
                    End Try
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Factura", i, Integer.Parse(IIf(Factura = "", 0, Factura)))
                    ' GUARDO LA INFO DE LA FACTURA Y EL VALOR A RETENER, PARA AL CREAR EL PAGO DESCONTARLE EL VALOR DE LA RETENCIÓN A CADA FACTURA
                    Utilitario.Util_Log.Escribir_Log(String.Format("Query:{0}..DocEntry-Respuesta:{1}", sQueryFactura, Factura), "frmDocumentoREXML")
                    If Not Factura = "0" Then
                        oFacturaVenta = New Entidades.FacturaVenta
                        oFacturaVenta.DocEntry = Factura
                        oFacturaVenta.ValorARetener = Convert.ToDouble(odetalle._valorRetenido)

                        Dim query As System.Collections.Generic.IEnumerable(Of Entidades.FacturaVenta)
                        query = oListaFacturaVenta.Where(Function(q As Entidades.FacturaVenta) q.DocEntry = Factura)
                        If query.Count() > 0 Then
                            query.Single().ValorARetener += Convert.ToDouble(odetalle._valorRetenido)

                        Else
                            oListaFacturaVenta.Add(oFacturaVenta)
                        End If

                        'END GUARDO LA INFO DE LA FACTURA Y EL VALOR A RETENER, PARA AL CREAR EL PAGO DESCONTARLE EL VALOR DE LA RETENCIÓN A CADA FACTURA
                    End If
                    'End If 'CargaFacturaRelacionadas 
                End If
                TotalValorRetenido += Convert.ToDouble(odetalle._valorRetenido)

                i += 1
            Next

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.Item.Enabled = False
            oGrid.Item.FromPane = 0
            oGrid.Item.ToPane = 0

            ' Comprobante
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Caption = "Comprobante"
            ' N. Comprobante
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "N. Comprobante"
            ' Fecha Emisión
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Caption = "Fecha Emisión"
            ' Ejercicio Fiscal
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Caption = "Ejercicio Fiscal"
            ' Base Imponible
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Caption = "Base Imponible"
            oGrid.Columns.Item(4).RightJustified = True
            ' Impuesto
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Impuesto"
            ' Codigo
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Caption = "Codigo"
            ' % Retención
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Caption = "% Retención"
            oGrid.Columns.Item(7).RightJustified = True
            ' Valor Retenido
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(8).TitleObject.Caption = "Valor Retenido"
            oGrid.Columns.Item(8).RightJustified = True

            If CargaFacturaRelacionadas = True Then
                ' DocEntry Factura de Venta Relacionada
                oGrid.Columns.Item(9).Editable = False
                oGrid.Columns.Item(9).TitleObject.Caption = "Factura"
                oGrid.Columns.Item(9).RightJustified = True
                Dim oEditTextColum As SAPbouiCOM.EditTextColumn
                oEditTextColum = oGrid.Columns.Item(9)
                oEditTextColum.LinkedObjectType = 13
            Else
                oGrid.Columns.Item(9).Visible = False
            End If

            'oGrid.Columns.Item(2).Description = "CodSAP"
            'oGrid.Columns.Item(2).TitleObject.Caption = "CodSAP"
            'oGrid.Columns.Item(2).Editable = False
            'Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            'oEditTextColum = oGrid.Columns.Item(2)
            'oEditTextColum.LinkedObjectType = 4


            oForm.Items.Item("txtSub").Enabled = False
            Dim txtSub As SAPbouiCOM.EditText
            txtSub = oForm.Items.Item("txtSub").Specific
            txtSub.Item.RightJustified = True
            'txtSub.Value = formatDecimal(Math.Round(TotalValorRetenido, 2).ToString())
            txtSub.Value = formatDecimal(Math.Round(TotalValorRetenido, 2).ToString())
            txtSub.Item.FromPane = 0
            txtSub.Item.ToPane = 0

            ' LINK FACTURA PRELIMINAR
            oForm.Items.Item("txtFPre").Enabled = False
            Dim txtFPre As SAPbouiCOM.EditText
            txtFPre = oForm.Items.Item("txtFPre").Specific

            'If iBorrador > 0 Then
            '    txtFPre.Value = iBorrador
            '    Dim Query As String = ""
            '    Query = "SELECT U_Tipo FROM ""@GS_DocumentosRec"" "
            '    Query += " Where U_DocEntry =  " + iBorrador.ToString()
            '    Query += " AND U_Estado = 'docPreliminar'"
            '    Query += " AND U_ObjType = '19'"
            '    Dim TipoDocumento As String = oFuncionesB1.getRSvalue(Query, "U_Tipo", "")
            '    cbxTipo.Select(TipoDocumento, SAPbouiCOM.BoSearchKey.psk_ByValue)
            '    cbxTipo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            '    OcultarObjetosPorTipo(TipoDocumento)
            '    If TipoDocumento = "NC Inventariable Relacionada" Then
            '        ' Recupero el ObjType Relacionados y los docEntrys de la tabla logs, que alimente cuando se creo el premilimar
            '        Dim Query2 As String = ""
            '        Query2 = "SELECT U_ObjTypeR, U_DocEntryR FROM ""@GS_DocumentosRec"" "
            '        Query2 += " Where U_DocEntry =  " + iBorrador.ToString()
            '        Query2 += " AND U_Estado = 'docPreliminar'"
            '        Query2 += " AND U_Tipo = '" + TipoDocumento + "'"
            '        Dim ObjTypeRelacionado As String = oFuncionesB1.getRSvalue(Query2, "U_ObjTypeR", "")
            '        Dim DocEntrysRelacionado As String = oFuncionesB1.getRSvalue(Query2, "U_DocEntryR", "")
            '        rsboApp.Forms.Item("frmDocumento").Items.Item("objR").Specific.value = ObjTypeRelacionado
            '        rsboApp.Forms.Item("frmDocumento").Items.Item("docR").Specific.value = DocEntrysRelacionado
            '        CargaDocumentoRelacionados()
            '    End If
            'Else
            '    cbxTipo.Select("NC Inventariable", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '    cbxTipo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            '    OcultarObjetosPorTipo("NC Inventariable")
            'End If

            Dim lnkP As SAPbouiCOM.LinkedButton
            lnkP = oForm.Items.Item("lnkP").Specific
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                lnkP.LinkedObjectType = "TM_RETV"
            Else

                lnkP.LinkedObjectType = 140
            End If
            lnkP.Item.LinkTo = "txtFPre"

            Dim flDetalle As SAPbouiCOM.Folder
            flDetalle = oForm.Items.Item("flDetalle").Specific
            'flDetalle.Select()


            ' END GRID DE DOCUMENTOS RELACIONADO

            'oForm.Width = 750
            'oForm.Height = 482

            oForm.Visible = True
            oForm.Select()

            'If iBorrador > 0 Then
            '    Try
            '        cbxTipo.Active = False
            '        cbxTipo.Item.Enabled = False
            '        oForm.Items.Item("obtnGrabar").Visible = False
            '        oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
            '        Dim oB As SAPbouiCOM.Button
            '        oB = oForm.Items.Item("2").Specific
            '        oB.Caption = "OK"
            '    Catch ex As Exception
            '    End Try
            'End If

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un error al cargar el documento :" + ex.Message().ToString(), 1, NombreAddon)
        End Try

    End Sub

    ''' <summary>
    ''' Metodo para cargar el formulario cuando ya esta creado el documento, siendo los estados: docPreliminar,docFinal
    ''' </summary>
    ''' <param name="IdDocumentoRecibido_UDO"></param>
    ''' <param name="Estado"></param>
    ''' <remarks></remarks>
    Public Sub CargaFormularioDocumentoExistente(IdDocumentoRecibido_UDO As String, Estado As String)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmDocumentoREXML") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentoREXML.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentoREXML").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmDocumentoREXML")
            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO
            oForm.Freeze(True)

            Dim chkDesc As SAPbouiCOM.CheckBox
            chkDesc = oForm.Items.Item("chkPa").Specific
            chkDesc.Item.Visible = False

            oForm.Items.Item("objR").Visible = False
            oForm.Items.Item("docR").Visible = False

            ' DATA TABLE CABECERA
            Try
                oForm.DataSources.DataTables.Add("dtDocCAB")
            Catch ex As Exception
            End Try
            Dim QueryCabecera As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCabecera = " SELECT ""U_RUC"" ,""U_Nombre"" ,""U_CardCode"""
                QueryCabecera += " ,""U_ClaAcc"" ,""U_NumAut"" "
                QueryCabecera += " ,""U_FecAut"" ,""U_NumDoc"" ,""U_FPrelim"""
                QueryCabecera += " ,""U_vTotal"""
                QueryCabecera += " ,""U_Tipo"" ,""U_IdGS"" ,""U_Sincro"""
                QueryCabecera += " ,""U_SincroE"" ,""U_Estado"" ,""U_FechaS"""
                QueryCabecera += "  FROM ""@GS_RER"" "
                QueryCabecera += "  WHERE ""DocEntry"" =  " + IdDocumentoRecibido_UDO
            Else
                QueryCabecera = " SELECT U_RUC ,U_Nombre ,U_CardCode"
                QueryCabecera += " ,U_ClaAcc ,U_NumAut "
                QueryCabecera += " ,U_FecAut ,U_NumDoc ,U_FPrelim"
                QueryCabecera += " ,U_vTotal"
                QueryCabecera += " ,U_Tipo ,U_IdGS ,U_Sincro"
                QueryCabecera += " ,U_SincroE ,U_Estado ,U_FechaS"
                QueryCabecera += "  FROM ""@GS_RER"" A WITH(NOLOCK)"
                QueryCabecera += "  WHERE A.DocEntry =  " + IdDocumentoRecibido_UDO
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocCAB").ExecuteQuery(QueryCabecera)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString() + " - Query: " + QueryCabecera, "frmDocumentoREXML")
            End Try

            odt = oForm.DataSources.DataTables.Item("dtDocCAB")
            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                _IdGS = Long.Parse(odt.GetValue("U_IdGS", i).ToString())

                Dim txtRUC As SAPbouiCOM.EditText
                txtRUC = oForm.Items.Item("txtRUC").Specific
                txtRUC.Value = odt.GetValue("U_RUC", i).ToString()

                oForm.Items.Item("txtNombre").Enabled = False
                Dim txtNommbre As SAPbouiCOM.EditText
                txtNommbre = oForm.Items.Item("txtNombre").Specific
                txtNommbre.Value = odt.GetValue("U_Nombre", i).ToString()

                oForm.Items.Item("txtF").Enabled = True
                Dim txtF As SAPbouiCOM.EditText
                txtF = oForm.Items.Item("txtF").Specific
                'txtF.Item.RightJustified = True
                txtF.Value = "0"

                'oForm.Items.Item("txtCodigo").Enabled = False
                Dim txtCodigo As SAPbouiCOM.EditText
                txtCodigo = oForm.Items.Item("txtCodigo").Specific
                txtCodigo.Value = odt.GetValue("U_CardCode", i).ToString()
                'txtCodigo.Item.Enabled = False
                Dim lnkCuentCN As SAPbouiCOM.LinkedButton
                lnkCuentCN = oForm.Items.Item("lnkCuentC").Specific
                lnkCuentCN.LinkedObjectType = 2
                lnkCuentCN.Item.LinkTo = "txtCodigo"

                oForm.Items.Item("txtClaAcc").Enabled = False
                Dim txtClaAcc As SAPbouiCOM.EditText
                txtClaAcc = oForm.Items.Item("txtClaAcc").Specific
                txtClaAcc.Value = odt.GetValue("U_ClaAcc", i).ToString()
                _ClaveAcceso = odt.GetValue("U_ClaAcc", i).ToString()

                oForm.Items.Item("txtNumAut").Enabled = False
                Dim txtNumAut As SAPbouiCOM.EditText
                txtNumAut = oForm.Items.Item("txtNumAut").Specific
                txtNumAut.Value = odt.GetValue("U_NumAut", i).ToString()

                oForm.Items.Item("txtFecAut").Enabled = False
                Dim txtFecAut As SAPbouiCOM.EditText
                txtFecAut = oForm.Items.Item("txtFecAut").Specific
                txtFecAut.Value = odt.GetValue("U_FecAut", i).ToString()

                oForm.Items.Item("txtNumDoc").Enabled = False
                Dim txtNumDoc As SAPbouiCOM.EditText
                txtNumDoc = oForm.Items.Item("txtNumDoc").Specific
                txtNumDoc.Value = odt.GetValue("U_NumDoc", i).ToString() ' _oDocumento.RetCabecera._estab + "-" + _oDocumento.RetCabecera._ptoEmi + "-" + _oDocumento.RetCabecera._secuencial

                Dim lbEstAut As SAPbouiCOM.StaticText
                lbEstAut = oForm.Items.Item("lbEstAut").Specific
                lbEstAut.Caption = "DOCUMENTO AUTORIZADO POR EL SRI"
                lbEstAut.Item.ForeColor = RGB(7, 118, 10)

                Dim lnkPDF As SAPbouiCOM.LinkedButton
                lnkPDF = oForm.Items.Item("lnkPDF").Specific
                lnkPDF.Item.Visible = True
                Dim lnkXML As SAPbouiCOM.LinkedButton
                lnkXML = oForm.Items.Item("lnkXML").Specific
                lnkXML.Item.Visible = False
                Dim Item_16 As SAPbouiCOM.StaticText
                Item_16 = oForm.Items.Item("Item_16").Specific
                Item_16.Item.Visible = True
                Dim Item_18 As SAPbouiCOM.StaticText
                Item_18 = oForm.Items.Item("Item_18").Specific
                Item_18.Item.Visible = False

                oForm.Items.Item("txtSub").Enabled = False
                Dim txtSub As SAPbouiCOM.EditText
                txtSub = oForm.Items.Item("txtSub").Specific
                txtSub.Item.RightJustified = True
                'txtSub.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_vTotal", i).ToString()), 2))
                txtSub.Value = Math.Round(Convert.ToDouble(odt.GetValue("U_vTotal", i).ToString()), 2)
                txtSub.Item.FromPane = 0
                txtSub.Item.ToPane = 0

                oForm.Items.Item("txtF").Enabled = True
                Dim Focus As SAPbouiCOM.EditText
                Focus = oForm.Items.Item("txtF").Specific
                'txtF.Item.RightJustified = True
                Focus.Value = "0"

                oForm.Items.Item("txtFPre").Enabled = False
                Dim txtFPre As SAPbouiCOM.EditText
                txtFPre = oForm.Items.Item("txtFPre").Specific
                txtFPre.Value = odt.GetValue("U_FPrelim", i).ToString()

            Next

            ' DATA TABLE DETALLE
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try
            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodRet", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("NumDocRe", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Fecha", SAPbouiCOM.BoFieldsType.ft_Date, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("pFiscal", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Base", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Impuesto", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Porcent", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("valorR", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            Dim QueryDetalle As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryDetalle = "SELECT  A.""U_CodRet"", A.""U_NumDocRe"", A.""U_Fecha"""
                QueryDetalle += ", A.""U_pFiscal"", A.""U_Base"", A.""U_Impuesto"""
                QueryDetalle += ", A.""U_Porcent"", A.""U_valorR"" "
                QueryDetalle += "  FROM ""@GS0_RER"" A "
                QueryDetalle += "  WHERE A.""DocEntry"" =  " + IdDocumentoRecibido_UDO
            Else
                QueryDetalle = "SELECT  A.U_CodRet, A.U_NumDocRe, A.U_Fecha"
                QueryDetalle += ", A.U_pFiscal, A.U_Base, A.U_Impuesto"
                QueryDetalle += ", A.U_Porcent, A.U_valorR "
                QueryDetalle += "  FROM ""@GS0_RER"" A WITH(NOLOCK)"
                QueryDetalle += "  WHERE A.DocEntry =  " + IdDocumentoRecibido_UDO
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocs").ExecuteQuery(QueryDetalle)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString() + " - Query: " + QueryDetalle, "frmDocumentoREXML")
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.Item.Enabled = False
            oGrid.Item.FromPane = 0
            oGrid.Item.ToPane = 0

            ' Comprobante
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Caption = "Comprobante"
            ' N. Comprobante
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "N. Comprobante"
            ' Fecha Emisión
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Caption = "Fecha Emisión"
            ' Ejercicio Fiscal
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Caption = "Ejercicio Fiscal"
            ' Base Imponible
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Caption = "Base Imponible"
            oGrid.Columns.Item(4).RightJustified = True
            ' Impuesto
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Impuesto"
            ' % Retención
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Caption = "% Retención"
            oGrid.Columns.Item(6).RightJustified = True
            ' Valor Retenido
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Caption = "Valor Retenido"
            oGrid.Columns.Item(7).RightJustified = True

            Dim lnkP As SAPbouiCOM.LinkedButton
            lnkP = oForm.Items.Item("lnkP").Specific
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                lnkP.LinkedObjectType = "TM_RETV"
            Else
                lnkP.LinkedObjectType = 140
            End If
            lnkP.Item.LinkTo = "txtFPre"

            Dim flDetalle As SAPbouiCOM.Folder
            flDetalle = oForm.Items.Item("flDetalle").Specific
            flDetalle.Select()

            oForm.Items.Item("obtnGrabar").Visible = False
            oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
            Dim oB As SAPbouiCOM.Button
            oB = oForm.Items.Item("2").Specific
            oB.Caption = "OK"

            'oForm.Width = 750
            'oForm.Height = 482

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - Error al cargar formulario recibido existente: " + ex.Message.ToString())
        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Private Function CrearPagoRecibido(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean

        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

            Return CrearPagoRecibido_E_O(sDocEntryPreliminar, DocEntryRERecibida_UDO)

        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

            Return CrearPagoRecibido_Exxis_seidor(sDocEntryPreliminar, DocEntryRERecibida_UDO)
            'Return CrearPagoRecibido_S(sDocEntryPreliminar, DocEntryRERecibida_UDO)
        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
            Return CrearPagoRecibido_TopManage(sDocEntryPreliminar, DocEntryRERecibida_UDO)
        End If

    End Function

    ''' <summary>
    ''' Creacion de Pago Recibido - Retención ( EXXIS - ONE SOLUTIONS )
    ''' </summary>
    ''' <param name="sDocEntryPreliminar"></param>
    ''' <param name="DocEntryRERecibida_UDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CrearPagoRecibido_E_O(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        'Dim RetVal As Long
        'Dim ErrCode As Long
        'Dim ErrMsg As String
        'Dim vPay As SAPbobsCOM.Payments
        'CLIENTE BANCARIO 
        Dim sQueryCB As String = ""
        'sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardCode""= '" + oCardCode.ToString + "' " se comento ya que el codigo cambia para el caso de crear un pago tipo proveedor
        sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardType"" = 'C' and ""LicTradNum""= '" + _sRUC.ToString + "' "
        Dim clienteBancario As String = oFuncionesB1.getRSvalue(sQueryCB, "U_SSCLIENTEBANCO", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo campo cliente - QUERY: " + sQueryCB + "Resultado :" + clienteBancario.ToString(), "frmDocumentoREXML")
        If clienteBancario = "SI" Then
            Dim estado As Boolean
            estado = CrearPagoRecibido_E_OCB(sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado

            'FINCLIENFTEBANCARIO
        ElseIf clienteBancario = "PROVEEDOR" Then
            Dim estado As Boolean
            estado = CrearPagoRecibido_E_OProveedor(sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado
        Else
            Dim estado As Boolean
            estado = CrearPagoRecibido_E_ONormal(sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado
        End If



    End Function

    ''' <summary>
    ''' Creacion de Pago Recibido - Retención ( SYPSOFT)
    ''' </summary>
    ''' <param name="sDocEntryPreliminar"></param>
    ''' <param name="DocEntryRERecibida_UDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CrearPagoRecibido_Exxis_seidor(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim Exxis_Seidor As String = ""
        Exxis_Seidor = Functions.VariablesGlobales._PagoRecibido_Seidor_exxis
        If Exxis_Seidor = "Y" Then
            'Crear pago recibido por tarjeta de credito
            Dim estado As Boolean
            estado = CrearPagoRecibido_E_S(sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado
        Else
            'crear pago normal para SYP
            Dim estado As Boolean
            estado = CrearPagoRecibido_S(sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado
        End If
    End Function
    Private Function CrearPagoRecibido_E_S(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCB As Integer = 1
        Dim secuencialCBMP As Integer = 1

        Try

            'Dim vPay As SAPbobsCOM.Documents
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            vPay.CardCode = oCardCode
            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES

            'Dim moneda As String = " SELECT TOP 1 ""Currency"" from OCRD where ""CardCode""= '" + oCardCode + "' "
            'Dim CodMoneda As String = oFuncionesB1.getRSvalue(moneda, "Currency", "")
            'If CodMoneda = "##" Then
            '    vPay.DocCurrency = ""
            'Else
            '    vPay.DocCurrency = CodMoneda.ToString
            'End If

            'vPay.DocCurrency = "USD"

            'vPay.DocDate = DateSerisal(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now
            vPay.DocDate = _oDocumento.RetCabecera._FechaAutorizacion
            vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            'vPay.DocDate = Date.Now
            'vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            'vPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.Reference1 = ""
            '' vPay.Series = 0
            'vPay.TaxDate = DateSerial(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now            
            ' vPay.TaxDate = Date.Now
            Try
                If oFuncionesB1.checkCampoBD("ORCT", "SYP_TIPOOPERACION") Then
                    vPay.UserFields.Fields.Item("U_SYP_TIPOOPERACION").Value = "C-101"
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("error SYP_TIPOOPERACION: " + ex.Message.ToString, "frmDocumentoREXML")
            End Try

            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1
            Dim TotalValorRetenido As Decimal = 0

            For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                If oDetalle._valorRetenido > 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle._codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString())
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmDocumentoREXML")

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Dim TypeCode As Integer = CrTypeCode
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End If
                    'vPay.CreditCards.CreditAcct = IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    'vPay.CreditCards.CreditCard = IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)

                    vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("CrTypeCode Asignado RENTA - QUERY: " + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End Try



                    vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                    vPay.CreditCards.CreditSum = oDetalle._valorRetenido ' TotalValorRetenido ' formatDecimal(TotalValorRetenido.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    vPay.CreditCards.FirstPaymentSum = TotalValorRetenido
                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    'If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion

                    Try
                        If Not IsNothing(oDetalle._numDocSustento) Then
                            'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Substring(6, 9)).ToString()
                            'Left(odt.GetValue(0, i).ToString(), 99))

                            'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Length)
                            vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                            vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                        End If

                    Catch ex As Exception
                    End Try


                    If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = _oDocumento.RetCabecera._secuencial
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = oCardCode
                    End If

                    vPay.CreditCards.VoucherNum = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi + _oDocumento.RetCabecera._secuencial
                    If oFuncionesB1.checkCampoBD("RCT3", "U_SYP_NROAUTO") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = _oDocumento.RetCabecera._FechaAutorizacion
                    End If

                    'End If


                    If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                    End If

                    Try
                        If oFuncionesB1.checkCampoBD("ORCT", "SYP_FECHARET") Then
                            Try
                                vPay.UserFields.Fields.Item("U_SYP_FECHARET").Value = _oDocumento.RetCabecera._fechaEmision
                            Catch ex As Exception
                            End Try
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "SYP_TipoOperacion") Then
                            Try
                                vPay.UserFields.Fields.Item("U_SYP_TipoOperacion").Value = "A-015"
                            Catch ex As Exception
                            End Try
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "SYP_DETIPO") Then
                            Try
                                vPay.UserFields.Fields.Item("U_SYP_DETIPO").Value = "RETENCION DE CLIENTES"
                            Catch ex As Exception
                            End Try
                        End If

                        If oFuncionesB1.checkCampoBD("ORCT", "FX_AUTO_RETENCION") Then
                            vPay.UserFields.Fields.Item("U_FX_AUTO_RETENCION").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If

                        If oFuncionesB1.checkCampoBD("ORCT", "SSCREADAR") Then
                            vPay.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "SSIDDOCUMENTO") Then
                            vPay.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("error seccion syp: " + ex.ToString(), "frmDocumentoREXML")
                    End Try
                    vPay.CreditCards.Add()
                    vPay.CreditCards.SetCurrentLine(secuencial)
                    secuencial += 1


                End If


            Next
            ' FACTURAS

            If CargaFacturaRelacionadas Then
                For Each o As Entidades.FacturaVenta In oListaFacturaVenta
                    Utilitario.Util_Log.Escribir_Log("Datos Docentry: " & o.DocEntry.ToString & " valor a retener: " & o.ValorARetener.ToString, "datosfacturasrelacionadas")
                    vPay.Invoices.DocEntry = o.DocEntry
                    vPay.Invoices.SumApplied = o.ValorARetener
                    vPay.Invoices.Add()
                Next
            End If



            RetVal = vPay.Add()
            'Dim xml As String = vPay.GetAsXML()
            'Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")



            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function
    Private Function CrearPagoRecibido_S(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments
        Try

            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            vPay.CardCode = oCardCode

            'vPay.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES

            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            'vPay.DocCurrency = "USD"

            If Functions.VariablesGlobales._vgFechaEmisionRetencion = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.DueDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionRetencionP = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            Else
                vPay.DocDate = Date.Now
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            End If

            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            vPay.Reference1 = ""

            Try
                vPay.UserFields.Fields.Item("U_SYP_PTSC").Value = _oDocumento.RetCabecera._ptoEmi
            Catch ex As Exception
                vPay.UserFields.Fields.Item("U_BPP_PTSC").Value = _oDocumento.RetCabecera._ptoEmi
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: 1037 frmDocumentoREXML " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            End Try
            Try
                vPay.UserFields.Fields.Item("U_SYP_SUCRET").Value = _oDocumento.RetCabecera._estab
            Catch ex As Exception
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: 1042 frmDocumentoREXML " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            End Try
            Try
                vPay.UserFields.Fields.Item("U_SYP_PTCC").Value = _oDocumento.RetCabecera._secuencial
            Catch ex As Exception
                vPay.UserFields.Fields.Item("U_BPP_PTCC").Value = _oDocumento.RetCabecera._secuencial
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: 1048 frmDocumentoREXML " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            End Try


            If oFuncionesB1.checkCampoBD("ORCT", "SYP_FECHARET") Then
                Try
                    vPay.UserFields.Fields.Item("U_SYP_FECHARET").Value = _oDocumento.RetCabecera._fechaEmision
                Catch ex As Exception
                End Try
            End If
            If oFuncionesB1.checkCampoBD("ORCT", "SYP_TipoOperacion") Then
                Try
                    vPay.UserFields.Fields.Item("U_SYP_TipoOperacion").Value = "A-015"
                Catch ex As Exception
                End Try
            End If
            If oFuncionesB1.checkCampoBD("ORCT", "SYP_DETIPO") Then
                Try
                    vPay.UserFields.Fields.Item("U_SYP_DETIPO").Value = "RETENCION DE CLIENTES"
                Catch ex As Exception
                End Try
            End If

            If oFuncionesB1.checkCampoBD("ORCT", "FX_AUTO_RETENCION") Then
                vPay.UserFields.Fields.Item("U_FX_AUTO_RETENCION").Value = _oDocumento.RetCabecera._NumeroAutorizacion
            End If

            If oFuncionesB1.checkCampoBD("ORCT", "SSCREADAR") Then
                vPay.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
            End If
            If oFuncionesB1.checkCampoBD("ORCT", "SSIDDOCUMENTO") Then
                vPay.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
            End If

            'vPay.AccountPayment.AccountCode = sCuentaRetencion
            'vPay.AccountPayments.AccountName = sCuentaRetencionNombre
            'vPay.AccountPayments.SumPaid = TotalValorRetenido
            'vPay.AccountPayments.Add()

            Dim CodRetencion As String = ""
            CodRetencion = Functions.VariablesGlobales._CodigoRetencion
            If CodRetencion = "" Then
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "ERROR - Revisar la configuracion en la opcion de parametrizaciones, debe tener registrado una Cuenta Contable Retención!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
                rsboApp.StatusBar.SetText(NombreAddon + " - Revisar la configuracion en la opcion de parametrizaciones, debe tener registrado una Cuenta Contable Retención", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            'oFuncionesB1.getRSvalue
            Dim sQueryCuentaRetencion As String = ""
            Dim CuentaRetencion As String = ""
            If rCompany.DbServerType = 9 Then
                sQueryCuentaRetencion = "select ""AcctCode"" from ""OACT"" where ""FormatCode"" = '" & CodRetencion & "'"
            Else
                sQueryCuentaRetencion = "select AcctCode from OACT WITH(NOLOCK) where FormatCode  = '" & CodRetencion & "'"
            End If
            CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")


            Dim CodRetencionRENTA As String = ""
            CodRetencionRENTA = Functions.VariablesGlobales._CodigoRetencionR
            Dim sQueryCuentaRetencionRENTA As String = ""
            Dim CuentaRetencionRENTA As String = ""
            If rCompany.DbServerType = 9 Then
                sQueryCuentaRetencionRENTA = "select ""AcctCode"" from ""OACT"" where ""FormatCode"" = '" & CodRetencionRENTA & "'"
            Else
                sQueryCuentaRetencionRENTA = "select AcctCode from OACT WITH(NOLOCK) where FormatCode  = '" & CodRetencionRENTA & "'"
            End If
            CuentaRetencionRENTA = oFuncionesB1.getRSvalue(sQueryCuentaRetencionRENTA, "AcctCode", "")

            '1 RENTA 2 IVA
            ' DETALLES
            Dim ValorRenta As Decimal = 0
            Dim ValorIva As Decimal = 0
            For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp
                If oDetalle._codigo = 1 Then
                    ValorRenta += oDetalle._valorRetenido
                ElseIf oDetalle._codigo = 2 Then
                    ValorIva += oDetalle._valorRetenido
                End If
            Next

            ' EN LA PESTAÑA TRANSFERENCIA VA EL VALOR RETENIDO DE IVA
            If ValorIva > 0 Then
                vPay.TransferAccount = CuentaRetencion
                vPay.TransferDate = Date.Now
                vPay.TransferSum = ValorIva
                vPay.TransferReference = "RTE " + _oDocumento.RetCabecera._secuencial
            End If

            ' EN LA PESTAÑA EFECTIVO VA EL VALOR DE RETENIDO DE LA FUENTE
            If ValorRenta > 0 Then
                vPay.CashAccount = IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA)
                vPay.CashSum = ValorRenta
            End If

            ' FACTURAS
            If CargaFacturaRelacionadas Then
                For Each o As Entidades.FacturaVenta In oListaFacturaVenta
                    vPay.Invoices.DocEntry = o.DocEntry
                    vPay.Invoices.SumApplied = o.ValorARetener
                    vPay.Invoices.Add()
                Next
            End If

            RetVal = vPay.Add()
            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If

        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try
    End Function

    Private Function CrearPagoRecibido_E_OCB(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCBMP As Integer = 1
        Try

            'CLIENTE BANCARIO 
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Creando Pago Recibido Tipo Cuenta(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido Tipo Cuenta(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rAccount
            'vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.DocCurrency = "USD"
            'vPay.DocDate = Date.Now
            'vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            If Functions.VariablesGlobales._vgFechaEmisionRetencion = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
                vPay.DueDate = _oDocumento.RetCabecera._fechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionRetencionP = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            Else

                vPay.DocDate = _oDocumento.RetCabecera._FechaAutorizacion
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            End If

            vPay.DocRate = 0

            'AGREGAR DETALLE DEL PAGO
            ' OBTENCION CUENTA CONTABLE
            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" Where ""CardCode"" =  '" + oCardCode + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardCode =  '" + oCardCode + "'"
            End If
            FormatCodeProveedor = oFuncionesB1.getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            Dim FormatCode As String = ""
            Dim sQueryAcctCode As String = ""
            If FormatCodeProveedor = "" Then
                FormatCode = Functions.VariablesGlobales._Cuenta_RE
            Else
                FormatCode = FormatCodeProveedor
            End If

            If FormatCode = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oFuncionesAddon.GuardaLOG("REE", _oDocumento.RetCabecera._claveAcceso, "ERROR - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode + "'"
            End If
            Dim Cuenta As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")
            ' END OBTENCION CUENTA CONTABLE
            Try
                vPay.AccountPayments.AccountCode = Cuenta
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("agregando cuenta" + ex.Message.ToString + Cuenta.ToString(), "frmDocumentoREXML")
            End Try

            'vPay.AccountPayments.AccountName = NombreCuentaRetencionCB

            Dim TotalValorRetenido As Decimal = 0

            For Each detalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                TotalValorRetenido += detalle._valorRetenido


            Next

            Try
                vPay.AccountPayments.SumPaid = TotalValorRetenido
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("agregando total retencion " + ex.Message.ToString + TotalValorRetenido.ToString(), "frmDocumentoREXML")
            End Try

            Try
                vPay.AccountPayments.Add()
                vPay.AccountPayments.SetCurrentLine(1)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("error al agregar lineas pago recibido " + ex.Message.ToString, "frmDocumentoREXML")
            End Try

            'END AGREGAR DETALLE DEL PAGO


            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1
            For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                If oDetalle._valorRetenido > 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle._codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString())
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmDocumentoREXML")

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    Dim TypeCode As Integer = CrTypeCode
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End If
                    'vPay.CreditCards.CreditAcct = IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    'vPay.CreditCards.CreditCard = IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Try
                        vPay.CreditCards.CreditAcct = CuentaRetencion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando cuenta retencion: " + CuentaRetencion.ToString(), "frmDocumentoREXML")
                    End Try
                    'IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    Try
                        vPay.CreditCards.CreditCard = CodRetencion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando codigo retencion: " + CodRetencion.ToString(), "frmDocumentoREXML")
                    End Try
                    ' IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando codigo tipo codigo: " + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End Try

                    Try
                        vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando numero de tarjeta de credito : " + _oDocumento.RetCabecera._secuencial.ToString(), "frmDocumentoREXML")
                    End Try

                    Try
                        vPay.CreditCards.CreditSum = oDetalle._valorRetenido
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando valor retenido : " + oDetalle._valorRetenido.ToString(), "frmDocumentoREXML")
                    End Try
                    ' TotalValorRetenido ' formatDecimal(TotalValorRetenido.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    Try
                        vPay.CreditCards.FirstPaymentSum = TotalValorRetenido
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando total retencion : " + TotalValorRetenido.ToString(), "frmDocumentoREXML")
                    End Try

                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        Try
                            vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + _oDocumento.RetCabecera._FechaAutorizacion.ToString(), "frmDocumentoREXML")
                        End Try


                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Substring(6, 9)).ToString()
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + oDetalle._numDocSustento.ToString(), "frmDocumentoREXML")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando monto base : " + oDetalle._baseImponible.ToString(), "frmDocumentoREXML")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando monto base : " + oDetalle._baseImponible.ToString(), "frmDocumentoREXML")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = _oDocumento.RetCabecera._secuencial
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando numero de retencion : " + _oDocumento.RetCabecera._secuencial.ToString(), "frmDocumentoREXML")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando numero de autorizacion de retencion : " + _oDocumento.RetCabecera._NumeroAutorizacion.ToString(), "frmDocumentoREXML")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando est y punto de emision : " + _oDocumento.RetCabecera._estab.ToString() + _oDocumento.RetCabecera._ptoEmi, "frmDocumentoREXML")
                        End Try
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = oCardCode
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando est y punto de emision : " + oCardCode.ToString, "frmDocumentoREXML")
                        End Try

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi + _oDocumento.RetCabecera._secuencial
                        If oFuncionesB1.checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = _oDocumento.RetCabecera._FechaAutorizacion
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        Try
                            vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + _oDocumento.RetCabecera._FechaAutorizacion.ToString(), "frmDocumentoREXML")
                        End Try


                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Substring(6, 9)).ToString()
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + oDetalle._numDocSustento.ToString(), "frmDocumentoREXML")
                        End Try
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = _oDocumento.RetCabecera._secuencial
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = oCardCode
                        End If

                        If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = _oDocumento.RetCabecera._secuencial
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If
                        Catch ex As Exception
                        End Try
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_TIP_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_TIP_RET").Value = "E"
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_SERIE_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_SERIE_RET").Value = _oDocumento.RetCabecera._estab.ToString + "-" + _oDocumento.RetCabecera._ptoEmi.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_RET").Value = _oDocumento.RetCabecera._secuencial.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_FEC_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_FEC_RET").Value = CDate(_oDocumento.RetCabecera._FechaAutorizacion)
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_CADRET") Then
                            vPay.UserFields.Fields.Item("U_HBT_CADRET").Value = CDate(_oDocumento.RetCabecera._FechaAutorizacion)
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "HBT_Depositado") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_HBT_Depositado").Value = "SI"
                        End If

                    End If


                    If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    Try
                        If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando docentry udo : " + DocEntryRERecibida_UDO.ToString, "frmDocumentoREXML")
                    End Try

                    Try
                        vPay.CreditCards.Add()
                        vPay.CreditCards.SetCurrentLine(secuencial)
                        secuencial += 1
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando lineas medio de pago : " + ex.Message.ToString, "frmDocumentoREXML")
                    End Try


                End If
            Next
            Try
                RetVal = vPay.Add()
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("erro al agregar pago recibido tipo cuenta : " + ex.Message.ToString, "frmDocumentoREXML")
            End Try

            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function

    Private Function CrearPagoRecibido_E_ONormal(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCB As Integer = 1
        Dim secuencialCBMP As Integer = 1
        'Dim fechaVencRtMP As Date
        Try

            'Dim vPay As SAPbobsCOM.Documents
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            vPay.CardCode = oCardCode
            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.DocCurrency = "USD"
            'vPay.DocDate = DateSerisal(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now
            'vPay.DocDate = _oDocumento.RetCabecera._FechaAutorizacion
            'vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            If Functions.VariablesGlobales._vgFechaEmisionRetencion = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.DueDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionRetencionP = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            Else
                vPay.DocDate = Date.Now
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            End If


            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            'vPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.Reference1 = ""
            '' vPay.Series = 0
            'vPay.TaxDate = DateSerial(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now            
            ' vPay.TaxDate = Date.Now


            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1

            Dim TotalValorRetenido As Decimal = 0

            For Each detalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                TotalValorRetenido += detalle._valorRetenido


            Next

            For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                If oDetalle._valorRetenido >= 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle._codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString())
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmDocumentoREXML")

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Dim TypeCode As Integer = CrTypeCode
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End If
                    'vPay.CreditCards.CreditAcct = IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    'vPay.CreditCards.CreditCard = IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)

                    vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("CrTypeCode Asignado RENTA - QUERY: " + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End Try




                    vPay.CreditCards.CreditSum = oDetalle._valorRetenido ' TotalValorRetenido ' formatDecimal(TotalValorRetenido.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    vPay.CreditCards.FirstPaymentSum = TotalValorRetenido
                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        'fechaVencRtMP = _oDocumento.RetCabecera._FechaAutorizacion
                        'vPay.CreditCards.CardValidUntil = fechaVencRtMP

                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Substring(6, 9)).ToString()
                                'Left(odt.GetValue(0, i).ToString(), 99))

                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Length)
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If

                        Catch ex As Exception
                        End Try


                        If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = _oDocumento.RetCabecera._secuencial
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = oCardCode
                        End If
                        'If oFuncionesB1.checkCampoBD("RCT3", "REPL_NUM_RETE") Then
                        '    vPay.CreditCards.UserFields.Fields.Item("U_REPL_NUM_RETE").Value = _oDocumento.RetCabecera._secuencial
                        'End If
                        If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = _oDocumento.RetCabecera._secuencial
                        End If


                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi + _oDocumento.RetCabecera._secuencial
                        If oFuncionesB1.checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = _oDocumento.RetCabecera._FechaAutorizacion
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If
                        Catch ex As Exception
                        End Try
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_TIP_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_TIP_RET").Value = "E"
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_SERIE_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_SERIE_RET").Value = _oDocumento.RetCabecera._estab.ToString + "-" + _oDocumento.RetCabecera._ptoEmi.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_RET").Value = _oDocumento.RetCabecera._secuencial.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_FEC_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_FEC_RET").Value = CDate(_oDocumento.RetCabecera._FechaAutorizacion)
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_CADRET") Then
                            vPay.UserFields.Fields.Item("U_HBT_CADRET").Value = CDate(_oDocumento.RetCabecera._FechaAutorizacion)
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "HBT_Depositado") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_HBT_Depositado").Value = "SI"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion

                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If

                        Catch ex As Exception
                        End Try


                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = _oDocumento.RetCabecera._secuencial
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = oCardCode
                        End If

                        If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = _oDocumento.RetCabecera._secuencial
                        End If

                    End If


                    If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                    End If

                    vPay.CreditCards.Add()
                    vPay.CreditCards.SetCurrentLine(secuencial)
                    secuencial += 1


                End If


            Next
            'dibeal
            Try
                If oFuncionesB1.checkCampoBD("ORCT", "DIB_TipoOperacion") Then
                    vPay.UserFields.Fields.Item("U_DIB_TipoOperacion").Value = "A-0015"
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("DIB_TipoOperacion error: " + ex.Message.ToString, "frmDocumentoREXML")
            End Try
            'Try
            '    If oFuncionesB1.checkCampoBD("ORCT", "DIB_TipoOperacion") Then
            '        vPay.UserFields.Fields.Item("U_DIB_TipoOperacion").Value = "A-0015"
            '    End If
            'Catch ex As Exception
            '    Utilitario.Util_Log.Escribir_Log("DIB_TipoOperacion error: " + ex.Message.ToString, "frmDocumentoREXML")
            'End Try

            ' FACTURAS
            If CargaFacturaRelacionadas Then
                For Each o As Entidades.FacturaVenta In oListaFacturaVenta
                    Utilitario.Util_Log.Escribir_Log("Datos Docentry: " & o.DocEntry.ToString & " valor a retener: " & o.ValorARetener.ToString, "datosfacturasrelacionadas")
                    vPay.Invoices.DocEntry = o.DocEntry
                    vPay.Invoices.SumApplied = o.ValorARetener
                    vPay.Invoices.Add()
                Next
            End If



            RetVal = vPay.Add()
            'Dim xml As String = vPay.GetAsXML()
            'Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")



            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function

    Private Function CrearPagoRecibido_E_OProveedor(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCB As Integer = 1
        Dim secuencialCBMP As Integer = 1
        'Dim fechaVencRtMP As Date
        Try

            'Dim vPay As SAPbobsCOM.Documents
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" WHERE ""CardType"" = 'C' and ""LicTradNum""= '" + _sRUC.ToString + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardType = 'C' and LicTradNum =  '" + _sRUC.ToString + "'"
            End If


            FormatCodeProveedor = oFuncionesB1.getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            Dim CUENTA As String = oFuncionesB1.getRSvalue("SELECT ""AcctCode"" FROM OACT WHERE  ""ActId""= '" + FormatCodeProveedor + "'", "AcctCode", "")

            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rSupplier
            vPay.CardCode = oCardCode
            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.DocCurrency = "USD"
            'vPay.DocDate = DateSerisal(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now
            'vPay.DocDate = _oDocumento.RetCabecera._FechaAutorizacion
            'vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            If Functions.VariablesGlobales._vgFechaEmisionRetencion = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.DueDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionRetencionP = "Y" Then
                vPay.DocDate = _oDocumento.RetCabecera._fechaEmision
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            Else
                vPay.DocDate = Date.Now
                vPay.TaxDate = _oDocumento.RetCabecera._fechaEmision
            End If


            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            'vPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.Reference1 = ""
            '' vPay.Series = 0
            'vPay.TaxDate = DateSerial(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now            
            ' vPay.TaxDate = Date.Now


            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1

            Dim TotalValorRetenido As Decimal = 0

            For Each detalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                TotalValorRetenido += detalle._valorRetenido


            Next
            For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp

                If oDetalle._valorRetenido > 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle._codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString())
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    ElseIf oDetalle._codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle._codigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmDocumentoREXML")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle._codigoRetencion.ToString() + " - " + oDetalle._porcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmDocumentoREXML")

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Dim TypeCode As Integer = CrTypeCode
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    If CrTypeCode = 0 Then
                        sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                        CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End If
                    'vPay.CreditCards.CreditAcct = IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    'vPay.CreditCards.CreditCard = IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)

                    vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle._codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle._codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle._codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("CrTypeCode Asignado RENTA - QUERY: " + CrTypeCode.ToString(), "frmDocumentoREXML")
                    End Try




                    vPay.CreditCards.CreditSum = oDetalle._valorRetenido ' TotalValorRetenido ' formatDecimal(TotalValorRetenido.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    vPay.CreditCards.FirstPaymentSum = TotalValorRetenido
                    'vPay.CreditCards.NumOfCreditPayments = 1
                    'vPay.CreditCards.NumOfPayments = 1

                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        'fechaVencRtMP = _oDocumento.RetCabecera._FechaAutorizacion
                        'vPay.CreditCards.CardValidUntil = fechaVencRtMP

                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Substring(6, 9)).ToString()
                                'Left(odt.GetValue(0, i).ToString(), 99))

                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle._numDocSustento.Length)
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If

                        Catch ex As Exception
                        End Try


                        If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = _oDocumento.RetCabecera._secuencial
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = oCardCode
                        End If
                        'If oFuncionesB1.checkCampoBD("RCT3", "REPL_NUM_RETE") Then
                        '    vPay.CreditCards.UserFields.Fields.Item("U_REPL_NUM_RETE").Value = _oDocumento.RetCabecera._secuencial
                        'End If
                        If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = _oDocumento.RetCabecera._secuencial
                        End If


                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi + _oDocumento.RetCabecera._secuencial
                        If oFuncionesB1.checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = _oDocumento.RetCabecera._FechaAutorizacion
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion
                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If
                        Catch ex As Exception
                        End Try
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_TIP_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_TIP_RET").Value = "E"
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_SERIE_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_SERIE_RET").Value = _oDocumento.RetCabecera._estab.ToString + "-" + _oDocumento.RetCabecera._ptoEmi.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_RET").Value = _oDocumento.RetCabecera._secuencial.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_NUM_AUT") Then
                            vPay.UserFields.Fields.Item("U_HBT_NUM_AUT").Value = _oDocumento.RetCabecera._NumeroAutorizacion.ToString
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_FEC_RET") Then
                            vPay.UserFields.Fields.Item("U_HBT_FEC_RET").Value = CDate(_oDocumento.RetCabecera._FechaAutorizacion)
                        End If
                        If oFuncionesB1.checkCampoBD("ORCT", "HBT_CADRET") Then
                            vPay.UserFields.Fields.Item("U_HBT_CADRET").Value = CDate(_oDocumento.RetCabecera._FechaAutorizacion)
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "HBT_Depositado") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_HBT_Depositado").Value = "SI"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                        vPay.CreditCards.FirstPaymentDue = _oDocumento.RetCabecera._FechaAutorizacion

                        Try
                            If Not IsNothing(oDetalle._numDocSustento) Then
                                vPay.CreditCards.VoucherNum = Left(oDetalle._numDocSustento.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                            End If

                        Catch ex As Exception
                        End Try


                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle._baseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = _oDocumento.RetCabecera._secuencial
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = _oDocumento.RetCabecera._NumeroAutorizacion
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = _oDocumento.RetCabecera._estab + _oDocumento.RetCabecera._ptoEmi
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = oCardCode
                        End If

                        If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = _oDocumento.RetCabecera._secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = _oDocumento.RetCabecera._secuencial
                        End If

                    End If


                    If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                    End If

                    vPay.CreditCards.Add()
                    vPay.CreditCards.SetCurrentLine(secuencial)
                    secuencial += 1


                End If


            Next
            'dibeal
            Try
                If oFuncionesB1.checkCampoBD("ORCT", "DIB_TipoOperacion") Then
                    vPay.UserFields.Fields.Item("U_DIB_TipoOperacion").Value = "A-0015"
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("DIB_TipoOperacion error: " + ex.Message.ToString, "frmDocumentoREXML")
            End Try


            vPay.ControlAccount = CUENTA

            RetVal = vPay.Add()

            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & ofrmDocumentoREXML.oCardCode.ToString() + ofrmDocumentoREXML._IdGS.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "ManejoDeDocumentos")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "ManejoDeDocumentos")
                End Try
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function
    Public Function CrearPagoRecibido_TopManage(ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
#Disable Warning BC42024 ' Variable local sin usar: 'RetVal'.
        Dim RetVal As Long
#Enable Warning BC42024 ' Variable local sin usar: 'RetVal'.
#Disable Warning BC42024 ' Variable local sin usar: 'ErrCode'.
        Dim ErrCode As Long
#Enable Warning BC42024 ' Variable local sin usar: 'ErrCode'.
#Disable Warning BC42024 ' Variable local sin usar: 'ErrMsg'.
        Dim ErrMsg As String
#Enable Warning BC42024 ' Variable local sin usar: 'ErrMsg'.
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim NumAtCard As String = ""

        Dim totalRT As Decimal = 0
        Dim totalIVA As Decimal = 0
        Dim _NumDocRet As String = _oDocumento.RetDetalleImp(0)._numDocSustento.ToString().Substring(0, 3) _
                                & "-" & _oDocumento.RetDetalleImp(0)._numDocSustento.ToString().Substring(3, 3) & "-" & _oDocumento.RetDetalleImp(0)._numDocSustento.ToString().Substring(6, 9)

        For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp
            If oDetalle._codigo = 1 Then ' RENTA
                totalRT = totalRT + oDetalle._valorRetenido
            ElseIf oDetalle._codigo = 2 Then ' IVA
                totalIVA = totalIVA + oDetalle._valorRetenido
            End If
        Next
        Try
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Comprobante de Retencion Venta estatus Borrador", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            'oForm = rsboApp.Forms.Item("frmDocumentoREXML")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("TM_RETV")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralData.SetProperty("U_TM_CARDCODE", Left(oCardCode.ToString, 15))
            'oGeneralData.SetProperty("U_TM_CARDNAME", Left(_oDocumento.NombreComercial.ToString, 15))
            If CargaFacturaRelacionadas Then
                oGeneralData.SetProperty("U_TM_NUMATCARD", _NumDocRet.ToString)
            End If
            oGeneralData.SetProperty("U_TM_DOCNUM_CR", _oDocumento.RetCabecera._estab.ToString + "-" + _oDocumento.RetCabecera._ptoEmi.ToString + "-" + _oDocumento.RetCabecera._secuencial.ToString)
            oGeneralData.SetProperty("U_TM_FAUT", _oDocumento.RetCabecera._FechaAutorizacion)
            oGeneralData.SetProperty("U_TM_CASRI", _oDocumento.RetCabecera._NumeroAutorizacion.ToString)
            oGeneralData.SetProperty("U_TM_CACC", _oDocumento.RetCabecera._claveAcceso.ToString())
            oGeneralData.SetProperty("U_TM_TAXDATE", _oDocumento.RetCabecera._FechaAutorizacion)
            oGeneralData.SetProperty("U_TM_TOTALRIVA", Convert.ToDouble(formatDecimal(totalIVA.ToString)))
            oGeneralData.SetProperty("U_TM_TOTALRIR", Convert.ToDouble(formatDecimal(totalRT.ToString)))
            oGeneralData.SetProperty("U_SSIDDOCUMENTO", DocEntryRERecibida_UDO.ToString)
            oGeneralData.SetProperty("U_SSCREADAR", "SI")


            oChildren = oGeneralData.Child("TM_LE_RETVL")
            For Each oDetalle As RetDetalleImpuestos In _oDocumento.RetDetalleImp
                oChild = oChildren.Add
                oChild.SetProperty("U_TM_BASERET", Convert.ToDouble(formatDecimal(oDetalle._baseImponible.ToString)))
                oChild.SetProperty("U_TM_IMPRET", Convert.ToDouble(formatDecimal(oDetalle._valorRetenido.ToString)))
                oChild.SetProperty("U_TM_PORCRET", Convert.ToDouble(formatDecimal(oDetalle._porcentajeRetener.ToString)))
            Next

            oGeneralParams = oGeneralService.Add(oGeneralData)
            sDocEntryPreliminar = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, "Comprobante Retencion Ventas Borrador, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
            'rCompany.GetNewObjectCode(sDocEntryPreliminar)
            'DocEntryNCRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            'oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryNCRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            'rsboApp.StatusBar.SetText(NombreAddon + " - Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryNCRecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "ERROR al crear registro en el udo comprobante retencion ventas: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Pago Recibido(Retencion) Recibida en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("EEROR al crear registro en el udo comprobante retencion ventas: " + ex.Message, "frmDocumentoREXML")
            Return False
        End Try
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

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            If typeEx = "frmDocumentoREXML" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        If pVal.BeforeAction = False Then
                            If pVal.ActionSuccess Then
                                'oForm.Freeze(True)
                                Try
                                    'Dim dFontHeightRatio As Double = Math.Round(rsboApp.GetFormItemDefaultHeight(SAPbouiCOM.BoFormSizeableItemTypes.fsit_EDIT) / 14.0, 2)          'Ratio is based on Edit text item. 14.00 is the reference Height that i created the forms in
                                    'Dim dFontWidthRatio As Double = Math.Round(rsboApp.GetFormItemDefaultWidth(SAPbouiCOM.BoFormSizeableItemTypes.fsit_EDIT) / 80.0, 2)    'Ratio is based on Edit text item. 80.00 is the reference Width that i created the forms in

                                    oGroupFolder = oForm.Items.Item("Item_21")
                                    oGroupFolder.Width = oForm.Width - 65
                                    oGroupFolder.Height = oForm.Height - 240
                                Catch ex As Exception

                                End Try
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK

                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID

                                Case "lnkPDF"
                                    ConsutarPDFRecibido(_IdGS, 2)

                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.Before_Action Then
                            Select Case pVal.ItemUID

                                Case "obtnGrabar"
                                    Dim rl As String = ""
                                    rl = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionLite")
                                    If rl = "Y" Then
                                        GuardarArchivoRETRecibido(_IdGS, 2)
                                    End If

                                    Dim obtnGrabar As SAPbouiCOM.Button
                                    obtnGrabar = oForm.Items.Item("obtnGrabar").Specific

                                    'Dim Exitoso As Boolean = False
                                    Dim sDocEntryPreliminar As String = "0"
                                    Dim iReturnValue As Integer
                                    iReturnValue = rsboApp.MessageBox(NombreAddon + " - Se Creará el Pago Recibido con la Retención Recibida, Desea Continuar", 1, "&SI", "&NO")
                                    If iReturnValue = 1 Then
                                        rsboApp.StatusBar.SetText(NombreAddon + "- Creando Pago Recibido con la Retención Recibida, por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                        'ANTES DE MANDAR A CREAR EL PAGO RECIBIDO GUARDO EL DOCUMENTO EN EL UDO
                                        ' GUARDO EL DOCUMENTO RECIBIDO EN EL UDO FACTURA RECIBIDA
                                        Dim DocEntryRERecibida_UDO As String = 0
                                        If Guarda_DocumentoRecibido_RE(DocEntryRERecibida_UDO) Then
                                            If CrearPagoRecibido(sDocEntryPreliminar, DocEntryRERecibida_UDO) Then
                                                ' BLOQUEO BOTON GRABAR
                                                oForm.Items.Item("obtnGrabar").Visible = False
                                                oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                                Dim oB As SAPbouiCOM.Button
                                                oB = oForm.Items.Item("2").Specific
                                                oB.Caption = "OK"
                                                ' MUESTRO EL LINK BUTTON DE PAGO RECIBIDO(RETENCION) PRELIMINAR 
                                                ' Asigno el docentry de la factura preliminar guardada.
                                                ' se busca por el numero de autorizacion top 1 descendente
                                                oForm.Items.Item("Item_23").Visible = True
                                                oForm.Items.Item("lnkP").Visible = True
                                                oForm.Items.Item("txtFPre").Visible = True

                                                Dim txtFPre As SAPbouiCOM.EditText
                                                txtFPre = oForm.Items.Item("txtFPre").Specific
                                                txtFPre.Value = sDocEntryPreliminar

                                                ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                                Try
                                                    rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(True)
                                                    Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidosXML").DataSources.DataTables.Item("dtDocs")
                                                    odt.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                    rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(False)
                                                Catch ex As Exception
                                                    Utilitario.Util_Log.Escribir_Log("obtnGrabar - Try Catch Remove Fila:" + ex.Message().ToString(), "frmDocumentoREXML")
                                                Finally
                                                    rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(False)
                                                End Try


                                                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.RetCabecera._claveAcceso, " Proceso terminado Exitosamente!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                            End If
                                        End If
                                    End If
                            End Select
                        Else
                            Select Case pVal.ItemUID
                                Case "chkPa"  ' RELACIONAR A FACTURAS
                                    RelacionarFacturas()
                            End Select
                        End If

                End Select
            End If

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + "Error en: rSboApp_ItemEvent," + ex.Message().ToString())
        End Try
    End Sub

    Private Sub RelacionarFacturas()

        Try
            oForm = rsboApp.Forms.Item("frmDocumentoREXML")
            oForm.Freeze(True)

            Dim oCheckbox As SAPbouiCOM.CheckBox = oForm.Items.Item("chkPa").Specific
            If oCheckbox.Checked = True Then
                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                oGrid.Columns.Item(9).Visible = True
                ' DocEntry Factura de Venta Relacionada
                oGrid.Columns.Item(9).Editable = False
                oGrid.Columns.Item(9).TitleObject.Caption = "Factura"
                oGrid.Columns.Item(9).RightJustified = True
                Dim oEditTextColum As SAPbouiCOM.EditTextColumn
                oEditTextColum = oGrid.Columns.Item(9)
                oEditTextColum.LinkedObjectType = 13
                CargaFacturaRelacionadas = True
            Else
                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                oGrid.Columns.Item(9).Visible = False
                CargaFacturaRelacionadas = False
            End If

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error OcultarObjetosPorTipo Cath: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            If oForm IsNot Nothing Then
                'Descongelamos el formulario haya habido o no excepción
                oForm.Freeze(False)
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Agrego logica a este evento para insertar registros en la tabla de usuario GS_DocumentosRec, la cual
    ''' contedra todos los documentos que pasaron por la creacion del preliminar por medio del addon Recepcion
    ''' </summary>
    ''' <param name="BusinessObjectInfo"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    Private Sub rSboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.FormDataEvent
        Try

            If BusinessObjectInfo.FormTypeEx = "170" Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If Not BusinessObjectInfo.BeforeAction Then
                            Select Case BusinessObjectInfo.ActionSuccess
                                Case True
                                    oDocumentoSAP = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                                    oDocumentoSAP.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                                    'oDocumentoSAPCB = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                                    'oDocumentoSAPCB.DocType = SAPbobsCOM.BoRcptTypes.rAccount
                                    'oDocumentoSAPCB.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)


                                    If oDocumentoSAP.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                        If Not oDocumentoSAP.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then

                                            If oDocumentoSAP.DocType = SAPbobsCOM.BoRcptTypes.rCustomer Or oDocumentoSAP.DocType = SAPbobsCOM.BoRcptTypes.rAccount Or oDocumentoSAP.DocType = SAPbobsCOM.BoRcptTypes.rSupplier Then

                                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
                                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                    ' SI EL PROVEEDOR ES EXXIS O ONE SOLUTIONS SE RECORRE LAS TARJETAS, POR QUE ES AHI DONDE REGISTRAN LA RETENCIÓN
                                                    Dim count As Integer = 0
                                                    Dim CREADAPORGSEDOC As String = oDocumentoSAP.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value.ToString
                                                    For count = 0 To oDocumentoSAP.CreditCards.Count - 1
                                                        oDocumentoSAP.CreditCards.SetCurrentLine(count)

                                                        'If oDocumentoSAP.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                                        If CREADAPORGSEDOC = "SI" Then
                                                            Dim idDocumentoRecibido_UDO As String = ""
                                                            Try
                                                                idDocumentoRecibido_UDO = oDocumentoSAP.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                                            Catch ex As Exception
                                                            End Try

                                                            ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                                            Dim idFacturaGS As String = ""
                                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                                            Else
                                                                idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                                            End If

                                                            ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                                            '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                                            Else
                                                                _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                                            End If
                                                            ' LE CAMBIA EL ESTADO A LA FACTURA UDO A DOCFINAL
                                                            ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docFinal")
                                                            ' ACTUALIZA EL CAMPO SINCRO A 1, ESTE CAMPO IDENTIFICA QUE YA ESTA SINCRONIZADA EN SAP
                                                            ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                                                            ' MARCA EL DOCUMENTO COMO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                                            ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)

                                                            ActualizadoEstadoUdoRetencionXML(idFacturaGS, "Contabilizado")
                                                            'MarcarVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO)
                                                            ' EL WINDOWS SERVICE DEBE SIEMPRE TOMAR COMO REFERENCIA EL CAMPO SINCRO, Y ENVIAR A EDOC LO QUE TENGA EL CAMPO SINCRO.
                                                            ' ES DECIR SI EL CAMPO SINCRO ES IGUAL A 1, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO TRUE
                                                            ' SI EL CAMPO ES IGUAL A 0, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO FALSE


                                                        End If
                                                    Next

                                                    ' SI LA PANTALLA DE DOCUMENTOS RECIBIDOS ESTA ABIERTA ELIMINO LA LINEA DE LA FACTURA RECIBIDA
                                                    ' YA QUE YA ESTA INTEGRADA
                                                    If CREADAPORGSEDOC = "SI" Then
                                                        Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                                                            If rsboApp.Forms.Item("frmDocumentosRecibidosXML").Visible = True Then
                                                                rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(True)
                                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidosXML").DataSources.DataTables.Item("dtDocs")
                                                                odt.Rows.Remove(_fila)
                                                                rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(False)
                                                            End If
                                                        Catch ex As Exception
                                                            Utilitario.Util_Log.Escribir_Log("et_FORM_DATA_ADD - Try Catch Remove Fila:" + ex.Message().ToString(), "frmDocumentoREXML")
                                                        Finally
                                                            rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(False)
                                                        End Try
                                                    End If
                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                    ' CUANDO EL PROVEEDOR ES SYPSOFT SE REVISA EN LA CABECERA DEL PAGO RECIBIDO
                                                    If oDocumentoSAP.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then

                                                        Dim idDocumentoRecibido_UDO As String = ""
                                                        Try
                                                            idDocumentoRecibido_UDO = oDocumentoSAP.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                                        Catch ex As Exception
                                                        End Try

                                                        ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                                        Dim idFacturaGS As String = ""
                                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                            idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                                        Else
                                                            idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                                        End If

                                                        ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                                        '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                            _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                                        Else
                                                            _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                                        End If
                                                        ' LE CAMBIA EL ESTADO A LA FACTURA UDO A DOCFINAL
                                                        ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docFinal")
                                                        ' ACTUALIZA EL CAMPO SINCRO A 1, ESTE CAMPO IDENTIFICA QUE YA ESTA SINCRONIZADA EN SAP
                                                        ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                                                        ' MARCA EL DOCUMENTO COMO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                                        ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                                                        ActualizadoEstadoUdoRetencionXML(idFacturaGS, "Contabilizado")
                                                        'MarcarVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO)
                                                        ' EL WINDOWS SERVICE DEBE SIEMPRE TOMAR COMO REFERENCIA EL CAMPO SINCRO, Y ENVIAR A EDOC LO QUE TENGA EL CAMPO SINCRO.
                                                        ' ES DECIR SI EL CAMPO SINCRO ES IGUAL A 1, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO TRUE
                                                        ' SI EL CAMPO ES IGUAL A 0, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO FALSE

                                                        ' SI LA PANTALLA DE DOCUMENTOS RECIBIDOS ESTA ABIERTA ELIMINO LA LINEA DE LA FACTURA RECIBIDA
                                                        ' YA QUE YA ESTA INTEGRADA
                                                        Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                                                            If rsboApp.Forms.Item("frmDocumentosRecibidosXML").Visible = True Then
                                                                rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(True)
                                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidosXML").DataSources.DataTables.Item("dtDocs")
                                                                odt.Rows.Remove(_fila)
                                                                rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(False)
                                                            End If
                                                        Catch ex As Exception
                                                        End Try
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If

                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If Not BusinessObjectInfo.BeforeAction Then
                            Select Case BusinessObjectInfo.ActionSuccess
                                Case True
                                    oDocumentoSAP = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                                    oDocumentoSAP.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                                    If oDocumentoSAP.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then ' SI ES UNA CANCELACION CAMBIO EL ESTADO EN EDOC A NO SINCRONIZADO
                                        If oDocumentoSAP.DocType = SAPbobsCOM.BoRcptTypes.rCustomer Or oDocumentoSAP.DocType = SAPbobsCOM.BoRcptTypes.rAccount Or oDocumentoSAP.DocType = SAPbobsCOM.BoRcptTypes.rSupplier Then

                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
                                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                ' SI EL PROVEEDOR ES EXXIS O ONE SOLUTIONS SE RECORRE LAS TARJETAS, POR QUE ES AHI DONDE REGISTRAN LA RETENCIÓN

                                                Dim count As Integer = 0
                                                For count = 0 To oDocumentoSAP.CreditCards.Count - 1
                                                    oDocumentoSAP.CreditCards.SetCurrentLine(count)

                                                    If oDocumentoSAP.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                                        Dim idDocumentoRecibido_UDO As String = ""
                                                        Try
                                                            idDocumentoRecibido_UDO = oDocumentoSAP.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                                        Catch ex As Exception
                                                        End Try
                                                        ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                                        Dim idFacturaGS As String = ""
                                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                            idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                                        Else
                                                            idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                                        End If
                                                        ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                                        '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                            _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                                        Else
                                                            _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                                        End If
                                                        ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docCancelado")

                                                        ' ACTUALIZA EL CAMPO SINCRO A 0, AL CANCELARLO EL SE DEBE MARCAR COMO NO SINCRONIZADO
                                                        ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 0)

                                                        ' MARCA EL DOCUMENTO COMO NO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                                        ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 0)
                                                        ActualizadoEstadoUdoRetencionXML(idFacturaGS, "Cancelado")
                                                        'MarcarNOVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO)

                                                        ' EL WINDOWS SERVICE DEBE SIEMPRE TOMAR COMO REFERENCIA EL CAMPO SINCRO, Y ENVIAR A EDOC LO QUE TENGA EL CAMPO SINCRO.
                                                        ' ES DECIR SI EL CAMPO SINCRO ES IGUAL A 1, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO TRUE
                                                        ' SI EL CAMPO ES IGUAL A 0, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO FALSE
                                                    End If

                                                Next

                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

                                                If oDocumentoSAP.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                                    Dim idDocumentoRecibido_UDO As String = ""
                                                    Try
                                                        idDocumentoRecibido_UDO = oDocumentoSAP.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                                    Catch ex As Exception
                                                    End Try
                                                    ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                                    Dim idFacturaGS As String = ""
                                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                        idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                                    Else
                                                        idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                                    End If
                                                    ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                                    '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                        _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                                    Else
                                                        _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                                    End If
                                                    ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docCancelado")

                                                    ' ACTUALIZA EL CAMPO SINCRO A 0, AL CANCELARLO EL SE DEBE MARCAR COMO NO SINCRONIZADO
                                                    ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 0)
                                                    ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 0)
                                                    ActualizadoEstadoUdoRetencionXML(idFacturaGS, "Cancelado")
                                                    ' MARCA EL DOCUMENTO COMO NO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                                    'MarcarNOVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO)

                                                    ' EL WINDOWS SERVICE DEBE SIEMPRE TOMAR COMO REFERENCIA EL CAMPO SINCRO, Y ENVIAR A EDOC LO QUE TENGA EL CAMPO SINCRO.
                                                    ' ES DECIR SI EL CAMPO SINCRO ES IGUAL A 1, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO TRUE
                                                    ' SI EL CAMPO ES IGUAL A 0, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO FALSE
                                                End If

                                            End If

                                        End If

                                    End If
                            End Select

                        End If
                End Select
            ElseIf BusinessObjectInfo.FormTypeEx = "UDO_FT_TM_RETV" Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If BusinessObjectInfo.BeforeAction = False Then
                            Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(BusinessObjectInfo.FormUID)
                            Dim CreadoPorAddon As String = mForm.DataSources.DBDataSources.Item("@TM_LE_RETVH").GetValue("U_SSCREADAR", 0).ToString.Trim()
                            Dim estado As String = mForm.DataSources.DBDataSources.Item("@TM_LE_RETVH").GetValue("U_TM_STATUS", 0).ToString.Trim()
                            Dim idDocumentoRecibido_UDO As String = mForm.DataSources.DBDataSources.Item("@TM_LE_RETVH").GetValue("U_SSIDDOCUMENTO", 0).ToString.Trim()
                            Dim DocContabilizado As String = ""
                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                DocContabilizado = oFuncionesB1.getRSvalue("SELECT ""U_SincroE"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_SincroE", "")
                            Else
                                DocContabilizado = oFuncionesB1.getRSvalue(" select U_SincroE from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_SincroE", "")
                            End If
                            If CreadoPorAddon = "SI" And DocContabilizado <> "1" Then
                                If estado <> "Borrador" Then
                                    ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                    Dim idFacturaGS As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                    Else
                                        idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                    End If

                                    ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                    '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                    Else
                                        _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                    End If
                                    ' LE CAMBIA EL ESTADO A LA FACTURA UDO A DOCFINAL
                                    Try
                                        ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docFinal")
                                        rsboApp.SetStatusBarMessage("Se actualizo estado del documento recibido ..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                                        rsboApp.SetStatusBarMessage("Se actualizo estado sincornizado del documento recibido ..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                                        ActualizadoEstadoUdoRetencionXML(idFacturaGS, "Contabilizado")
                                        'MarcarVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO)
                                        rsboApp.SetStatusBarMessage("Se marco como contabilizado el documento recibido ..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                                    Catch ex As Exception
                                    End Try


                                    Try
                                        If rsboApp.Forms.Item("frmDocumentosRecibidosXML").Visible = True Then
                                            rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(True)
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidosXML").DataSources.DataTables.Item("dtDocs")
                                            odt.Rows.Remove(_fila)
                                            rsboApp.Forms.Item("frmDocumentosRecibidosXML").Freeze(False)
                                        End If
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If
                        End If
                End Select
            End If

        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " ERROR - CATH DATA EVENT :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - rSboApp_FormDataEvent: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function Guarda_DocumentoRecibido_RE(ByRef DocEntryNCRecibida_UDO As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Utilitario.Util_Log.Escribir_Log("DocEntryNCRecibida_UDO" + DocEntryNCRecibida_UDO.ToString, "frmDocumentoREXML")
        Try
            rsboApp.StatusBar.SetText(NombreAddon + "- Creando registro de Pago Recibido(Retencion) Recibida UDO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Creando registro de Pago Recibido(Retencion) Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oForm = rsboApp.Forms.Item("frmDocumentoREXML")
            Try
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_RER")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                'oGeneralData.SetProperty("Code", conta)
                oGeneralData.SetProperty("U_RUC", oForm.Items.Item("txtRUC").Specific.Value.ToString())
                oGeneralData.SetProperty("U_Nombre", Left(oForm.Items.Item("txtNombre").Specific.Value.ToString(), 99))
                oGeneralData.SetProperty("U_CardCode", oForm.Items.Item("txtCodigo").Specific.Value.ToString())
                'oGeneralData.SetProperty("U_Mapeado", oForm.Items.Item("lbMapp").Specific.Value.ToString())
                oGeneralData.SetProperty("U_ClaAcc", oForm.Items.Item("txtClaAcc").Specific.Value.ToString())
                oGeneralData.SetProperty("U_NumAut", oForm.Items.Item("txtNumAut").Specific.Value.ToString())
                oGeneralData.SetProperty("U_FecAut", oForm.Items.Item("txtFecAut").Specific.Value.ToString())
                oGeneralData.SetProperty("U_NumDoc", oForm.Items.Item("txtNumDoc").Specific.Value.ToString())
                oGeneralData.SetProperty("U_FPrelim", oForm.Items.Item("txtFPre").Specific.Value.ToString())
                oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSub").Specific.Value.ToString())))
                oGeneralData.SetProperty("U_IdGS", _oDocumento.RetCabecera._DocEntry.ToString())
                oGeneralData.SetProperty("U_Sincro", 0)
                oGeneralData.SetProperty("U_Estado", "docPreliminar")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("cabecera UDO:" + ex.Message, "frmDocumentoREXML")
            End Try



#Disable Warning BC42104 ' La variable 'oGeneralData' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            oChildren = oGeneralData.Child("GS0_RER")
#Enable Warning BC42104 ' La variable 'oGeneralData' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1
                oChild = oChildren.Add
                oChild.SetProperty("U_CodRet", odt.GetValue(0, i).ToString())
                oChild.SetProperty("U_NumDocRe", odt.GetValue(1, i).ToString())
                oChild.SetProperty("U_Fecha", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_pFiscal", odt.GetValue(3, i).ToString())
                oChild.SetProperty("U_Base", Convert.ToDouble(formatDecimal(odt.GetValue(4, i).ToString())))
                oChild.SetProperty("U_Impuesto", odt.GetValue(5, i).ToString())
                oChild.SetProperty("U_Porcent", Convert.ToDouble(formatDecimal(odt.GetValue(7, i).ToString())))
                oChild.SetProperty("U_valorR", Convert.ToDouble(formatDecimal(odt.GetValue(8, i).ToString())))
            Next

#Disable Warning BC42104 ' La variable 'oGeneralService' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            oGeneralParams = oGeneralService.Add(oGeneralData)
#Enable Warning BC42104 ' La variable 'oGeneralService' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            DocEntryNCRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryNCRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryNCRecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Ocurrior un error al crear registro de Pago Recibido(Retencion) Recibida UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Pago Recibido(Retencion) Recibida en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al guardar Pago Recibido(Retencion) Recibida en el UDO:" + ex.Message, "frmDocumentoREXML")
            Return False
        End Try
    End Function

    Public Sub Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String, DocEntryPreliminar As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            Dim query As String
            Dim CodeExist As String = "0"
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@GS_RER"" Where ""DocEntry"" = '" + DocEntryRERecibida_UDO + "' "
            Else
                query = "Select DocEntry From [@GS_RER] Where DocEntry = '" + DocEntryRERecibida_UDO + "' "
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            If CodeExist = "0" Or CodeExist = "" Then ' 
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "No existen registros coincidentes, la consulta no trajo registro: " + query.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Else
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_RER")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

                oGeneralService.Update(oGeneralData)

            End If

            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: Actualizando Numero de Documento Preliminar en Documento Recibido UDO: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub
    Public Sub Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try

            rsboApp.StatusBar.SetText(NombreAddon + " - Eliminando Documento Recibido UDO Retención # " + DocEntryRERecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Eliminando Documento Recibido UDO Retención # " + DocEntryRERecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryRERecibida_UDO)

            'oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            'oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

            oGeneralService.Delete(oGeneralParams)


            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO Retención..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub

    Public Function ActualizadoEstadoUdoRetencionXML(ByRef IdUdoFcXML As String, Sincronizado As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Actualizando a Sincronizado = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RT")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", IdUdoFcXML)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", Sincronizado)
            oGeneralData.SetProperty("U_FechaFin", DateTime.Now.ToString)
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Se usa para enviar cambiar el estado a sincronizado en SAP BO
    ''' </summary>
    ''' <param name="DocEntryFacturaRecibida_UDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ActualizadoEstadoSincronizado_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Actualizando a Sincronizado = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Sincro", Sincronizado)
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Cambia el estado a Sincronizado EDOC, cuando ya el web service edoc me haya dicho que fue marcado como sincronizado en EDOC
    ''' Ya que caso contrario significa que cuando se envión a marcar como sincronizado en EDOC, no hubo conexion al ws, en ese caso se 
    ''' debe crear un windows service para que valide el campo Sincro y SincroE, si exite un registro con Sincro = 1 y SincroE = 0, el servicio debe tomar
    ''' ese registro y mandar a sincronizar EDOC hasta tener respuesta
    ''' </summary>
    ''' <param name="DocEntryFacturaRecibida_UDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_SincroE", Sincronizado)
            'oGeneralData.SetProperty("U_Estado", "docPreliminar")
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE : " + ex.Message.ToString, "frmDocumento")
            'rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Actualiza el estado del UDO Factura, cuando pasa de Preliminar a Documento Definitivo o a Cancelado
    ''' </summary>
    ''' <param name="DocEntryFacturaRecibida_UDO"></param>
    ''' <param name="Estado"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ActualizadoEstado_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Estado As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", Estado)
            oGeneralData.SetProperty("U_FechaS", Integer.Parse(Date.Now.ToString("yyyyMMdd")))
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Function MarcarVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String) As Boolean
        Try
            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

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
            'If Functions.VariablesGlobales._vgHttps = "Y" Then
            '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
            'End If
            ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
            If WS.MarcarVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                'oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Documento Marcado como Visto(Integrado) en EDOC ", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1)
                Return True
            Else
                'oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Error al marcar documento como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS : " + mensaje, Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error MarcarVisto : " + ex.Message.ToString, "frmDocumentoREXML")
            Return False
        End Try
    End Function

    Public Function MarcarNOVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String) As Boolean
        Try
            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

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
            'If Functions.VariablesGlobales._vgHttps = "Y" Then
            '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
            'End If
            ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
            If WS.MarcarNoVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 0)
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Documento Marcado como NO Visto(Integrado) en EDOC Satisfactoriamente! ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return True
            Else
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, " Error al marcar documento NO como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub ConsutarPDFRecibido(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer)
        Try

            'Dim rl As String = ""
            'rl = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionLite")
            'If rl = "Y" Then
            '    'Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO
            '    Dim rutarl = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Ruta_Compartida")
            '    Dim fechacarpeta As String = Date.Today.ToString("dd-MM-yyyy")
            '    Dim fechacreacion As String = ""
            '    Dim rutaFC As String = ""
            '    rutaFC = rutarl & "\" & "RETENCIONES" & "\" & fechacarpeta & "\"
            '    Dim filepath As String = rutaFC
            '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '        fechacreacion = oFuncionesB1.getRSvalue("SELECT ""CreateDate"" FROM ""@GS_RER"" where ""U_IdGS"" = " + IdDocumento.ToString, "CreateDate", "")
            '    Else
            '        fechacreacion = oFuncionesB1.getRSvalue("SELECT CreateDate FROM ""@GS_RER"" WITH(NOLOCK) where U_IdGS = " + IdDocumento.ToString, "CreateDate", "")
            '    End If
            '    fechacreacion = CDate(fechacreacion).Date.ToString("dd-MM-yyyy")
            '    If fechacarpeta = fechacreacion Then
            '        filepath += _ClaveAcceso + ".pdf"
            '    Else
            '        rutaFC = rutarl & "\" & "RETENCIONES" & "\" & fechacreacion & "\"
            '        filepath = rutaFC
            '        filepath += _ClaveAcceso + ".pdf"
            '    End If

            '    Dim Proc As New Process()
            '    Proc.StartInfo.FileName = filepath
            '    Proc.Start()
            '    Proc.Dispose()

            'Else
            '    _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            '    If _WS_Recepcion = "" Then
            '        rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            '    End If
            '    _WS_RecepcionArchivo = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsultaArchivo")
            '    _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

            '    rsboApp.StatusBar.SetText(NombreAddon + " - Ruta Recepcion: " + _WS_RecepcionArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '    rsboApp.StatusBar.SetText(NombreAddon + " - Clave Recepcion: " + _WS_RecepcionClave, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            '    Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO
            '    WS.Url = _WS_RecepcionArchivo
            '    ' MANEJO PROXY
            '    Dim SALIDA_POR_PROXY As String = ""
            '    SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            '    Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            '    Dim Proxy_puerto As String = ""
            '    Dim Proxy_IP As String = ""
            '    Dim Proxy_Usuario As String = ""
            '    Dim Proxy_Clave As String = ""

            '    If SALIDA_POR_PROXY = "Y" Then
            '        Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
            '        Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
            '        Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
            '        Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

            '        Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
            '        Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
            '        Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
            '        Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

            '        If Not Proxy_puerto = "" Then
            '            proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
            '        Else
            '            proxyobject = New System.Net.WebProxy(Proxy_IP)
            '        End If
            '        cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

            '        proxyobject.Credentials = cred

            '        WS.Proxy = proxyobject
            '        WS.Credentials = cred
            '    End If
            '    ' END  MANEJO PROXY

            '    Dim filepath As String = Path.GetTempPath()
            '    filepath += _ClaveAcceso + ".pdf"

            '    ' SI NO EXISTE EN LA CARPETA TEMPORAL, LO CONSULTO AL WS
            '    If Not File.Exists(filepath) Then
            '        rsboApp.SetStatusBarMessage(NombreAddon + " - Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
            '        Dim FS As FileStream = Nothing
            '        'If Functions.VariablesGlobales._vgHttps = "Y" Then
            '        '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
            '        'End If
            '        'oManejoDocumentos.SetProtocolosdeSeguridad()
            '        ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
            '        Dim dbbyte As Byte() = WS.ConsultaArchivoProveedor_PDF(_WS_RecepcionClave, TipoDocumento, IdDocumento, mensaje)
            '        If dbbyte Is Nothing Then
            '            rsboApp.SetStatusBarMessage(NombreAddon + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
            '        Else
            '            FS = New FileStream(filepath, System.IO.FileMode.Create)
            '            FS.Write(dbbyte, 0, dbbyte.Length)
            '            FS.Close()
            '        End If

            '    End If

            '    Dim Proc As New Process()
            '    Proc.StartInfo.FileName = filepath
            '    Proc.Start()
            '    Proc.Dispose()
            'End If

            Dim numeracion = _ClaveAcceso.ToString.Substring(24, 15)
            Dim fecha = _ClaveAcceso.ToString.Substring(0, 8)
            Dim ruc = _ClaveAcceso.ToString.Substring(10, 13)

            Dim año = Right(fecha, 4)
            Dim mes = fecha.ToString.Substring(2, 2)
            Dim dia = fecha.ToString.Substring(0, 2)

            Dim nameFile = "07_" & ruc & "_" & año & mes & dia & "_" & numeracion
            Utilitario.Util_Log.Escribir_Log("Nombre archivo" & nameFile, "frmDocumentoREXML")



            Dim filepath As String = Functions.VariablesGlobales._RutaProRT & nameFile & ".pdf"
            If File.Exists(filepath) Then

                Dim Proc As New Process()
                Proc.StartInfo.FileName = filepath
                Proc.Start()
                Proc.Dispose()

            ElseIf Not File.Exists(filepath) Then

                rsboApp.StatusBar.SetText(NombreAddon + " -  No se encuentra PDF con clave : " + _ClaveAcceso.ToString + " verificar en la ruta: " + Functions.VariablesGlobales._RutaProRT.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If



        Catch ex As Exception
            rsboApp.SetStatusBarMessage(NombreAddon + " - Ocurrio un error al generar el PDF recibido! " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
        End Try
    End Sub

    Public Sub GuardarArchivoRETRecibido(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer)
        'obtengo la ruta compartida
        Dim rutarl = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Ruta_Compartida")
        'ontegno la fecha actual
        Dim fechacarpeta As String = Date.Today.ToString("dd-MM-yyyy")
        Dim rutaFC As String = ""
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO

        Try
            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionArchivo = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsultaArchivo")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

            rsboApp.StatusBar.SetText(NombreAddon + " - Ruta Recepcion: " + _WS_RecepcionArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            rsboApp.StatusBar.SetText(NombreAddon + " - Clave Recepcion: " + _WS_RecepcionClave, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Dim pdf As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO

            rsboApp.StatusBar.SetText(NombreAddon + " - Seteando Entidad ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            pdf.Url = _WS_RecepcionArchivo

            'verifico si la carpeta NOTADECREDITO existe en la ruta compartida
            If Not Directory.Exists(rutarl & "\" & "RETENCIONES") Then

                Directory.CreateDirectory(rutarl & "\" & "RETENCIONES")
                Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + rutarl & "\" & "RETENCIONES".ToString, "frmDocumento")
            End If
            'verifico si la carpeta con la fecha actual existe dentro de la carpeta NOTADECREDITOA
            If Not Directory.Exists(rutarl & "\" & "RETENCIONES" & "\" & fechacarpeta) Then
                Directory.CreateDirectory(rutarl & "\" & "RETENCIONES" & "\" & fechacarpeta)
                Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + rutarl & "\" & "RETENCIONES" & "\" & fechacarpeta, "frmDocumento")

            End If

            rutaFC = rutarl & "\" & "RETENCIONES" & "\" & fechacarpeta & "\"
            WS.Url = _WS_RecepcionArchivo

            Dim filepath As String = rutaFC
            'obtengo la ruta completa añadiendo la clave de acceso
            filepath += _ClaveAcceso + ".pdf"

            'verifico si la clave de acceso existe dentro de la carpeta NOTADECREDITO
            If Not File.Exists(filepath) Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - Guardando pdf, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim FS As FileStream = Nothing
                Dim dbbyte As Byte() = WS.ConsultaArchivoProveedor_PDF(_WS_RecepcionClave, TipoDocumento, _oDocumento.RetCabecera._DocEntry, mensaje)
                If dbbyte Is Nothing Then
                    rsboApp.SetStatusBarMessage(NombreAddon + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Else
                    FS = New FileStream(filepath, System.IO.FileMode.Create)
                    FS.Write(dbbyte, 0, dbbyte.Length)
                    FS.Close()
                    rsboApp.SetStatusBarMessage(NombreAddon + " - PDF guardado exitosamente..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                End If
            Else
                rsboApp.SetStatusBarMessage(NombreAddon + " Ya existe un pdf con esta clave de acceso: " + _ClaveAcceso.ToString() + " verificar en la ruta compartida", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            End If
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(NombreAddon + " - Ocurrio un error al guardar el PDF recibido! " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
        End Try


    End Sub

    Private Sub InsertaRegistroDocumentoRecibidoLOG(oDocumento As SAPbobsCOM.Documents, Estado As String, log As String, tipoDocumento As String)
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim oUserTable As SAPbobsCOM.UserTable
        Try
            ' ESTADOS
            ' docPreliminar - Cuando se crea el premilimar
            ' docFinal      - Cuando ya pasa a un documento real
            ' docSincronizado Cuando ya esta sincronizado con EDOC
            ' docReAbierto  - Cuando vuelve a cambiar el estado a EDOC para volverlo a ingresar
            ' docCancelado  - Cuando al documento real le hacen una cancelacion SAP
            ' Error         - Para describir los errores por el catch

            oUserTable = rCompany.UserTables.Item("GS_DocumentosRec")
            Dim Secuencia As Integer = oFuncionesB1.getCorrelativo("code", "[@GS_DocumentosRec]")
            '// set the two default fields 
            oUserTable.Code = Secuencia
            oUserTable.Name = Secuencia

            oUserTable.UserFields.Fields.Item("U_ObjType").Value = oDocumento.DocObjectCodeEx '"18"
            oUserTable.UserFields.Fields.Item("U_DocSubType").Value = oDocumento.DocumentSubType '"--"
            oUserTable.UserFields.Fields.Item("U_DocEntry").Value = oDocumento.DocEntry.ToString()
            oUserTable.UserFields.Fields.Item("U_Folio").Value = (oDocumento.UserFields.Fields.Item("U_SER_EST").Value.ToString() + "-" + oDocumento.UserFields.Fields.Item("U_SER_PE").Value.ToString() + "-" + oDocumento.FolioNumber.ToString()).ToString()
            oUserTable.UserFields.Fields.Item("U_CardCode").Value = oDocumento.CardCode
            oUserTable.UserFields.Fields.Item("U_CardName").Value = oDocumento.CardName
            oUserTable.UserFields.Fields.Item("U_Valor").Value = oDocumento.DocTotal.ToString
            oUserTable.UserFields.Fields.Item("U_ClaAcce").Value = oDocumento.UserFields.Fields.Item("U_SSCLAVE").Value.ToString()
            oUserTable.UserFields.Fields.Item("U_Estado").Value = Estado
            oUserTable.UserFields.Fields.Item("U_Tipo").Value = tipoDocumento
            oUserTable.UserFields.Fields.Item("U_Log").Value = log
            oUserTable.UserFields.Fields.Item("U_Fecha").Value = Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))

            Try
                oUserTable.UserFields.Fields.Item("U_ObjTypeR").Value = rsboApp.Forms.Item("frmDocumento").Items.Item("objR").Specific.value.ToString()
            Catch ex As Exception
            End Try
            Try
                oUserTable.UserFields.Fields.Item("U_DocEntryR").Value = rsboApp.Forms.Item("frmDocumento").Items.Item("docR").Specific.value()
            Catch ex As Exception
            End Try


            oUserTable.Add()
            '// Check for errors
            rCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                rsboApp.StatusBar.SetText(NombreAddon + " - Error al ingresar configuracion previa: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                rsboApp.StatusBar.SetText(NombreAddon + " - Grabando Log como: " + Estado.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If


        Catch ex As Exception

        Finally
#Disable Warning BC42104 ' La variable 'oUserTable' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            If oUserTable IsNot Nothing Then
#Enable Warning BC42104 ' La variable 'oUserTable' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
            End If

        End Try
    End Sub

    Private Function ReturnDocEntryDocBase(XMLITEMSRELACIONADOS As String, itemCode As String) As Integer()
        Dim DOCENTRY() As Integer
        Try
            ' http://www.dotnetcurry.com/linq/564/linq-to-xml-tutorials-examples

            Dim xelement As XElement = XElement.Parse(XMLITEMSRELACIONADOS)

            ' HAGO LINQ PARA DOCENTRY DEL DOCUMENTO RELACIONADO
            Dim Rows = xelement.Elements("Rows").Elements("Row").ToList() ' OBTENGO TODAS LAS LINEAS DEL GRID RELACIONADO
            Dim contador As Integer = 0
            For Each xEle As XElement In Rows
                Dim Codigo As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)
                Codigo = xEle.Descendants("Cells").Elements("Cell").Elements("Value").Skip(2).Take(1) ' OBTENGO EL CODIGO DEL ARTICULO DEL GRID RELACIONADO
                If Codigo.Value = itemCode Then                                                           ' COMPARO EL CODIGO DEL ARTICULO CON ITEM CODE DE SAP
                    Dim DocEntrys As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)
                    DocEntrys = xEle.Descendants("Cells").Elements("Cell").Elements("Value").Take(1)  ' SI SON IGUALES OBTENGO EL DOCENTRY DEL DOC RELACIONADO
                    ReDim Preserve DOCENTRY(contador)
                    DOCENTRY(contador) = DocEntrys.Value
                    contador += 1
                End If

            Next xEle

#Disable Warning BC42104 ' La variable 'DOCENTRY' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            Return DOCENTRY
#Enable Warning BC42104 ' La variable 'DOCENTRY' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
        Catch ex As Exception

            Return DOCENTRY
        End Try
    End Function
    ''' <summary>
    ''' Valido que todos los codigo de articulos mapeados existan en el grid de los documentos relacionados
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidarMapeoDeItems() As String
        Dim ItemCodeRelacionado As String = ""
        Dim ItemCodeMapeado As String = ""
        Dim XML As String = ""
        Dim Encontrado As Boolean = True
        Try
            Dim dtRECIBIDO As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs")    ' DATA TABLE DOCUMENTOS RECIBIDO
            Dim dtRELACIONADO As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocr") ' DATA TABLE DOCUMENTOS RELACIONADO

            XML = dtRELACIONADO.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly) ' SERIALIZO EL DT PARA PODER HACER UN SELECT
            Dim xelement As XElement = XElement.Parse(XML)

            For index As Integer = 0 To dtRECIBIDO.Rows.Count - 1
                ItemCodeMapeado = Convert.ToString(dtRECIBIDO.GetValue(2, index))
                Dim Rows = xelement.Elements("Rows").Elements("Row").ToList() ' OBTENGO TODAS LAS LINEAS DEL GRID RELACIONADO
                If Encontrado Then
                    Encontrado = False
                    For Each xEle As XElement In Rows
                        Dim Codigo As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)
                        Codigo = xEle.Descendants("Cells").Elements("Cell").Elements("Value").Skip(2).Take(1) ' OBTENGO EL CODIGO DEL ARTICULO DEL GRID RELACIONADO
                        If Codigo.Value = ItemCodeMapeado Then                                                           ' COMPARO EL CODIGO DEL ARTICULO CON ITEM CODE DE SAP                     
                            Encontrado = True
                            ItemCodeMapeado = ""
                            Exit For
                        End If
                    Next xEle
                Else
                    Return ItemCodeMapeado
                    Exit Function
                End If

            Next
            Return ItemCodeMapeado
        Catch ex As Exception
            ItemCodeMapeado = "Error"
            Return ItemCodeMapeado
        End Try

    End Function

    Private Function ContarItems(xml As String, itemCode As String) As Integer
        Try
            Dim xelement As XElement = XElement.Parse(xml)
            Dim result As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)
            result = xelement.Descendants("Rows").Elements("Row").Elements("Cells").Elements("Cell").Where(Function(n) n.Element("Value").Value = itemCode)
            Return result.Count()
        Catch ex As Exception

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


#Region "Utiles"
    Public Function CierraFormulario(ByVal bIsOpenFormulario As Boolean, ByVal strNombreFormulario As String) As Boolean
        If bIsOpenFormulario Then
            Try
                rsboApp.Forms.Item(strNombreFormulario).Close()
            Catch ex As Exception
                Return False
            End Try
        End If
        Return False
    End Function
    Public Function GetFechaYYYYMMDD(ByVal fecha As Date) As String
        Dim sFecha As String = ""
        Dim dia As String = ""
        Dim mes As String = ""
        Dim año As String = ""
        Try

            año = Strings.Right("0000" & CType(fecha.Year, String), 4)
            mes = Strings.Right("00" & CType(fecha.Month, String), 2)
            dia = Strings.Right("00" & CType(fecha.Day, String), 2)
            sFecha = año & mes & dia
            Return sFecha

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End Try
    End Function

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
                    ' result = Double.Parse((numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString())), CultureInfo.InvariantCulture)
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

    'Public Shared Function formatDecimal(ByVal numero As String) As Decimal

    '    Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
    '    Dim result As Double = 0
    '    Try
    '        If numero IsNot Nothing Then
    '            If Not numero.Contains(",") Then
    '                result = Double.Parse(numero, CultureInfo.InvariantCulture)
    '            Else
    '                result = Convert.ToDouble(numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
    '            End If
    '        End If
    '    Catch e As Exception
    '        Try
    '            result = Convert.ToDouble(numero)
    '        Catch
    '            Try
    '                result = Convert.ToDouble(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."))
    '            Catch
    '                Throw New Exception("Wrong string-to-double format")
    '            End Try
    '        End Try
    '    End Try
    '    Return result

    '    'Dim formato As Decimal
    '    'If Not numero.Equals(String.Empty) Then
    '    '    Dim sep As Char = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator
    '    '    Select Case sep
    '    '        Case "."
    '    '            formato = numero.Replace(",", sep)
    '    '        Case ","
    '    '            formato = numero.Replace(".", sep)
    '    '    End Select
    '    'End If
    '    'Return formato

    'End Function

    Private Function ConvertToDouble(s As String) As Double
        Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
        Dim result As Double = 0
        Try
            If s IsNot Nothing Then
                If Not s.Contains(",") Then
                    result = Double.Parse(s, CultureInfo.InvariantCulture)
                Else
                    result = Convert.ToDouble(s.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
                End If
            End If
        Catch e As Exception
            Try
                result = Convert.ToDouble(s)
            Catch
                Try
                    result = Convert.ToDouble(s.Replace(",", ";").Replace(".", ",").Replace(";", "."))
                Catch
                    Throw New Exception("Wrong string-to-double format")
                End Try
            End Try
        End Try
        Return result
    End Function


#End Region

#Region "Funciones Comentadas"

    'Private Function CrearFacturaPreliminarRelacionada() As Boolean

    '    'Dim S As String = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocr").SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly)
    '    'rsboApp.MessageBox(S.ToString())

    '    Dim RetVal As Long
    '    Dim ErrCode As Long
    '    Dim ErrMsg As String
    '    rsboApp.StatusBar.SetText("GS - Creando Factura por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    'Create the Documents object
    '    Dim GRPO As SAPbobsCOM.Documents
    '    GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

    '    Try
    '        ' If baseGRPO.GetByKey(PO_DocEntry) = True Then
    '        GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
    '        GRPO.CardCode = oCardCode
    '        GRPO.DocDate = _oDocumento.RetCabecera._fechaEmision
    '        GRPO.DocDueDate = Today.Date

    '        'iTotalPO_Line = baseGRPO.Lines.Count
    '        'iTotalFrgChg_Line = baseGRPO.Expenses.Count

    '        ' DATOS DE AUTORIZACION
    '        GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = _oDocumento.RetCabecera._NumeroAutorizacion
    '        GRPO.UserFields.Fields.Item("U_SER_EST").Value =  _oDocumento.RetCabecera._estab
    '        GRPO.UserFields.Fields.Item("U_SER_PE").Value = _oDocumento.RetCabecera._ptoEmi

    '        'U_EREC_CREADO 
    '        GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"

    '        GRPO.FolioNumber = _oDocumento.RetCabecera._secuencial

    '        Dim dtRECIBIDO As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs")    ' DATA TABLE DOCUMENTOS RECIBIDO
    '        Dim dtRELACIONADO As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocr") ' DATA TABLE DOCUMENTOS RELACIONADO

    '        'Dim nItemsRecibidos As Integer = dtRECIBIDO.Rows.Count()
    '        'Dim nItemsRelacionados As Integer = dtRELACIONADO.Rows.Count()


    '        Dim XMLITEMSRELACIONADOS As String = ""
    '        XMLITEMSRELACIONADOS = dtRELACIONADO.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly) ' SERIALIZO EL DT PARA PODER HACER UN SELECT

    '        Dim XMLITEMSRECIBIDOS As String = ""
    '        XMLITEMSRECIBIDOS = dtRECIBIDO.SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly) ' SERIALIZO EL DT PARA PODER HACER UN SELECT

    '        Dim DocEntry() As Integer
    '        Dim itemCode As String = ""
    '        Dim nItemRecibido As Integer
    '        Dim nItemRelacionado As Integer
    '        ' RECORRO EL DT DEL DOCUMENTO RECIBIDO
    '        For index As Integer = 0 To dtRECIBIDO.Rows.Count - 1
    '            itemCode = Convert.ToString(dtRECIBIDO.GetValue(2, index))

    '            ' POR CADA DOCUMENTO RECIBIDO VALIDO LAS CANTIDADES DE ITEMS EN EL GRID RECIBIDO Y CONTRA EL RELACIONADO
    '            ' SE DEBE SABER YA QUE PUEDE EXISTIR UNA FACTURA RECIBIDO CON LINEAS QUE CONTIENEN EL MISMO ITEM
    '            ' DE IGUAL FORMA PUEDE HABER UN UN ITEM REPETIDO EN EL GRID RELACIONADO
    '            nItemRecibido = ContarItems(XMLITEMSRECIBIDOS, itemCode)
    '            nItemRelacionado = ContarItems(XMLITEMSRELACIONADOS, itemCode)

    '            If nItemRecibido = 1 And nItemRelacionado = 1 Then
    '                DocEntry = ReturnDocEntryDocBase(XMLITEMSRELACIONADOS, itemCode)
    '                If ObjTypeRelacionado = 22 Then
    '                    GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
    '                ElseIf ObjTypeRelacionado = 20 Then
    '                    GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
    '                Else
    '                End If
    '                GRPO.Lines.BaseEntry = DocEntry(0)  'Convert.ToInt32(3616)
    '                'GRPO.Lines.BaseLine = Convert.ToInt32(dtT.GetValue(1, index)) '1
    '                GRPO.Lines.Quantity = formatDecimal(dtRECIBIDO.GetValue(4, index).ToString()) 'Cantidad
    '                GRPO.Lines.Price = formatDecimal(dtRECIBIDO.GetValue(5, index).ToString()) 'Precio
    '                'GRPO.Lines.Descuento = formatDecimal(dtRECIBIDO.GetValue(6, index).ToString()) 'Descuento
    '                GRPO.Lines.LineTotal = formatDecimal(dtRECIBIDO.GetValue(7, index).ToString()) 'Line Total
    '                GRPO.Lines.Add()
    '            End If
    '            If nItemRecibido = 1 And nItemRelacionado > 1 Then
    '                DocEntry = ReturnDocEntryDocBase(XMLITEMSRELACIONADOS, itemCode)
    '                ' CALCULO LA CANTIDAD DEPENDIENDO DEL TOTAL DE REGISTROS
    '                ' LA CANTIDAD PUEDE SER EN DOUBLE O ENTERO 
    '                Dim TotalCantidad As Integer = formatDecimal(dtRECIBIDO.GetValue(4, index).ToString())
    '                Dim Cantidad As Integer
    '                Dim CantidadUtima As Integer
    '                Dim residuo As Integer
    '                Dim CantdadComoInteger = True
    '                Try
    '                    TotalCantidad = formatDecimal(dtRECIBIDO.GetValue(4, index).ToString())
    '                    Cantidad = Int(TotalCantidad / nItemRelacionado)
    '                    residuo = (TotalCantidad Mod nItemRelacionado)
    '                    If residuo > 0 Then
    '                        CantidadUtima = Cantidad + 1
    '                    End If
    '                Catch ex As Exception
    '                    CantdadComoInteger = False
    '                End Try

    '                Dim TotalCantidadD As Double = formatDecimal(dtRECIBIDO.GetValue(4, index).ToString())
    '                Dim CantidadD As Double
    '                CantidadD = Math.Round((TotalCantidadD / nItemRelacionado), 2)

    '                Dim Contador As Integer = 1
    '                For Each number As Integer In DocEntry
    '                    If ObjTypeRelacionado = 22 Then
    '                        GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
    '                    ElseIf ObjTypeRelacionado = 20 Then
    '                        GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
    '                    Else
    '                    End If
    '                    GRPO.Lines.BaseEntry = number
    '                    If Contador = DocEntry.Count Then
    '                        GRPO.Lines.Quantity = IIf(CantdadComoInteger, CantidadUtima, CantidadD)
    '                        GRPO.Lines.Price = formatDecimal(dtRECIBIDO.GetValue(5, index).ToString()) 'Precio
    '                        'GRPO.Lines.Descuento = formatDecimal(dtRECIBIDO.GetValue(6, index).ToString()) 'Descuento
    '                        GRPO.Lines.LineTotal = Math.Round(IIf(CantdadComoInteger, CantidadUtima, CantidadD) * formatDecimal(dtRECIBIDO.GetValue(5, index).ToString())) 'Line Total
    '                    Else
    '                        GRPO.Lines.Quantity = IIf(CantdadComoInteger, Cantidad, CantidadD)
    '                        GRPO.Lines.Price = formatDecimal(dtRECIBIDO.GetValue(5, index).ToString()) 'Precio
    '                        'GRPO.Lines.Descuento = formatDecimal(dtRECIBIDO.GetValue(6, index).ToString()) 'Descuento
    '                        'GRPO.Lines.LineTotal = formatDecimal(dtRECIBIDO.GetValue(7, index).ToString()) 'Line Total
    '                        GRPO.Lines.LineTotal = Math.Round(IIf(CantdadComoInteger, Cantidad, CantidadD) * formatDecimal(dtRECIBIDO.GetValue(5, index).ToString())) 'Line Total
    '                    End If

    '                    GRPO.Lines.Add()
    '                    Contador += 1
    '                Next

    '            End If

    '        Next

    '        GRPO.Comments += "Creado por el addon eREP"

    '        'Dim iTotalPO_Line As Integer
    '        'Dim iTotalFrgChg_Line As Integer
    '        'If baseGRPO.GetByKey(PO_DocEntry) = True Then
    '        '    iTotalFrgChg_Line = baseGRPO.Expenses.Count
    '        '    ' Freight Charges
    '        '    If iTotalFrgChg_Line > 0 Then
    '        '        Dim fcnt As Integer
    '        '        For fcnt = 0 To iTotalFrgChg_Line - 1
    '        '            GRPO.Expenses.SetCurrentLine(fcnt)
    '        '            GRPO.Expenses.ExpenseCode = baseGRPO.Expenses.ExpenseCode
    '        '            GRPO.Expenses.BaseDocType = "22"
    '        '            GRPO.Expenses.BaseDocLine = baseGRPO.Expenses.LineNum
    '        '            GRPO.Expenses.BaseDocEntry = baseGRPO.DocEntry
    '        '            GRPO.Expenses.Add()
    '        '        Next
    '        '    End If
    '        'End If

    '        'Add the Invoice
    '        RetVal = GRPO.Add

    '        'Check the result
    '        If RetVal <> 0 Then
    '            rCompany.GetLastError(ErrCode, ErrMsg)
    '            rsboApp.MessageBox(ErrCode & " " & ErrMsg)
    '            Return False
    '        Else
    '            Return True
    '        End If

    '        'End If

    '    Catch ex As Exception
    '        rsboApp.StatusBar.SetText("GS - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    Finally
    '        GRPO = Nothing
    '        GC.Collect()
    '    End Try


    'End Function

#End Region

End Class

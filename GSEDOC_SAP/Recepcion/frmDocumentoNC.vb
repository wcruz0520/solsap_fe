Imports System.IO
Imports System.Threading
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Linq

'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security


Public Class frmDocumentoNC
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable
    Dim oCardCode As String = ""
    Dim _oDocumento As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito
    Public listaDetalleArtiulos As New List(Of Entidades.DetalleArticulo)
    Dim _fila As Integer
    Dim ObjTypeRelacionado As Integer = 0
    Dim oDocumentoSAP As SAPbobsCOM.Documents
    Dim _WS_Recepcion As String = ""
    Dim _WS_RecepcionCambiarEstado As String = ""
    Dim _WS_RecepcionClave As String = ""
    Dim _WS_RecepcionArchivo As String = ""
    Dim _ClaveAcceso As String = ""
    Dim _IdGS As Long = 0
    Dim _TipoDocumento As String = ""
    Dim oGroupFolder As SAPbouiCOM.Item

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential

    Dim _sRUC As String = ""

    Dim _EsDocumentoCargadoPorXML As Boolean = False
    Dim NumeroPedido As String = ""

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioDocumento(sRUC As String, sCardCode As String, sNombre As String, oDocumento As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito, ofila As Integer, Optional documentoCargadoPorXML As Boolean = False)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmDocumentoNC") Then
            Exit Sub
        End If
        oCardCode = sCardCode
        _fila = ofila
        _sRUC = sRUC
        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentoNC.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
                _oDocumento = oDocumento
                _ClaveAcceso = oDocumento.ClaveAcceso
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentoNC").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            'add Arturo 12042024
            'bandera para saber si se integro la info por XML

            _EsDocumentoCargadoPorXML = documentoCargadoPorXML


            oForm = rsboApp.Forms.Item("frmDocumentoNC")
            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO
            oForm.Freeze(True)
            '
            oForm.Items.Item("objR").Visible = False ' Guardo el ObjType del documento relacionado, lo lleno desde frm consultaordenes
            oForm.Items.Item("docR").Visible = False ' Guardo los docEntrys de los documentos relacionados, lo lleno desde frm consultaordenes

            _IdGS = _oDocumento.IdNotaCredito

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
            txtClaAcc.Value = _oDocumento.ClaveAcceso

            oForm.Items.Item("txtNumAut").Enabled = False
            Dim txtNumAut As SAPbouiCOM.EditText
            txtNumAut = oForm.Items.Item("txtNumAut").Specific
            txtNumAut.Value = _oDocumento.AutorizacionSRI

            oForm.Items.Item("txtFecAut").Enabled = False
            Dim txtFecAut As SAPbouiCOM.EditText
            txtFecAut = oForm.Items.Item("txtFecAut").Specific
            txtFecAut.Value = _oDocumento.FechaAutorizacion

            oForm.Items.Item("txtNumDoc").Enabled = False
            Dim txtNumDoc As SAPbouiCOM.EditText
            txtNumDoc = oForm.Items.Item("txtNumDoc").Specific
            txtNumDoc.Value = _oDocumento.Establecimiento + "-" + _oDocumento.PuntoEmision + "-" + _oDocumento.Secuencial

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


            Dim cbEnlazar As SAPbouiCOM.ButtonCombo
            cbEnlazar = oForm.Items.Item("cbEnlazar").Specific
            cbEnlazar.ValidValues.Add("Devolución de Mercadería", "Devolución de Mercadería")
            cbEnlazar.ValidValues.Add("Factura de Proveedores", "Factura de Proveedores")
            cbEnlazar.ValidValues.Add("Anticipo de Proveedores", "Anticipo de Proveedores")
            cbEnlazar.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodPrin", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodAuxi", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodSAP", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Descripc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Cantid", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Precio", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Desc", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Float, 100)

            Dim PendienteMapear As Boolean = False

            oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(oDocumento.ENTDetalleNotaCredito.Count)
            Dim i As Integer = 0

            For Each odetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito In oDocumento.ENTDetalleNotaCredito

                oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodPrin", i, IIf(IsNothing(odetalle.CodigoPrincipal), odetalle.Descripcion, odetalle.CodigoPrincipal))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodAuxi", i, IIf(IsNothing(odetalle.CodigoAuxiliar), "", odetalle.CodigoAuxiliar)) '
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", i, oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + sCardCode + "' AND ""Substitute"" = '" + Left(odetalle.CodigoPrincipal, 50) + "'", "ItemCode", ""))
                Else
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", i, oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WHERE CardCode = '" + sCardCode + "' AND Substitute = '" + Left(odetalle.CodigoPrincipal, 50) + "'", "ItemCode", ""))
                End If

                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Descripc", i, IIf(IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion).ToString().Length > 250, odetalle.Descripcion.Substring(1, 249), IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion)))
                If IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion).ToString().Length > 250 Then
                    Try
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Descripc", i, odetalle.Descripcion.Substring(1, 249))
                    Catch ex As Exception
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Descripc", i, "-")
                    End Try
                Else
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Descripc", i, IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion))
                End If

                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Cantid", i, Convert.ToDouble(odetalle.Cantidad))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Precio", i, Convert.ToDouble(odetalle.PrecioUnitario))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Desc", i, Convert.ToDouble(odetalle.Descuento))
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Total", i, Convert.ToDouble(odetalle.PrecioTotalSinImpuesto))

                Dim CodigoArticulo As String = ""
                If PendienteMapear = False Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" ='" + sCardCode + "' AND ""Substitute""  = '" + odetalle.CodigoPrincipal + "'", "ItemCode", "")
                    Else
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WITH(NOLOCK) WHERE CardCode ='" + sCardCode + "' AND Substitute  = '" + odetalle.CodigoPrincipal + "'", "ItemCode", "")
                    End If

                    If String.IsNullOrEmpty(CodigoArticulo) Then
                        PendienteMapear = True
                    End If
                End If

                i += 1
            Next

            If PendienteMapear = False Then
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = oForm.Items.Item("lbMapp").Specific
                lbMapp.Value = "SI"
                lbMapp.Item.ForeColor = RGB(7, 118, 10)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = oForm.Items.Item("btnMapear").Specific
                btnMapear.Item.Visible = False

            Else
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = oForm.Items.Item("lbMapp").Specific
                lbMapp.Value = "NO"
                'lbMap.Item.ForeColor = RGB(7, 118, 10)
                lbMapp.Item.ForeColor = ColorTranslator.ToOle(Color.Red)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = oForm.Items.Item("btnMapear").Specific
                lbMapp.Item.Visible = True
            End If


            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.Item.Enabled = False
            oGrid.Item.FromPane = 0
            oGrid.Item.ToPane = 0

            ' Codigo
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Caption = "Codigo"
            ' Codigo Auxiliar
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Codigo Auxiliar"
            ' ItemCode
            oGrid.Columns.Item(2).Description = "Cod.SAP"
            oGrid.Columns.Item(2).TitleObject.Caption = "Cod.SAP"
            oGrid.Columns.Item(2).Editable = False
            Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            oEditTextColum = oGrid.Columns.Item(2)
            oEditTextColum.LinkedObjectType = 4
            ' Descripcion
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Caption = "Descripcion"
            ' Cantidad
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Caption = "Cantidad"
            oGrid.Columns.Item(4).RightJustified = True
            ' Precio
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Precio"
            oGrid.Columns.Item(5).RightJustified = True
            ' Descuento
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Caption = "Descuento"
            oGrid.Columns.Item(6).RightJustified = True
            ' Total
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Caption = "Total"
            oGrid.Columns.Item(7).RightJustified = True

            Dim BaseImponibleIVA As Decimal = 0
            Dim BaseImponibleIVA5 As Decimal = 0
            Dim BaseImponible0 As Decimal = 0
            Dim BaseImponibleNoObjeto As Decimal = 0
            Dim BaseImponibleExento As Decimal = 0
            Dim Iva As Decimal = 0
            Dim Iva5 As Decimal = 0
            Dim ICE As Decimal = 0
            Dim BaseImponibleICE As Decimal = 0
            For Each facImpuesto As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto In _oDocumento.ENTNotaCreditoImpuesto
                If facImpuesto.Codigo = 2 Then
                    If facImpuesto.CodigoPorcentaje = 2 Or facImpuesto.CodigoPorcentaje = 3 Or facImpuesto.CodigoPorcentaje = 4 Or facImpuesto.CodigoPorcentaje = 8 Or facImpuesto.CodigoPorcentaje = 10 Then
                        BaseImponibleIVA += facImpuesto.BaseImponible
                        Iva += facImpuesto.Valor
                    ElseIf facImpuesto.CodigoPorcentaje = 5 Then
                        BaseImponibleIVA5 += facImpuesto.BaseImponible
                        Iva5 += facImpuesto.Valor
                    ElseIf facImpuesto.CodigoPorcentaje = 0 Then
                        BaseImponible0 += facImpuesto.BaseImponible
                    ElseIf facImpuesto.CodigoPorcentaje = 6 Then
                        BaseImponibleNoObjeto += facImpuesto.BaseImponible
                    ElseIf facImpuesto.CodigoPorcentaje = 7 Then
                        BaseImponibleExento += facImpuesto.BaseImponible
                    End If
                End If
            Next
            For Each facImpuesto As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto In _oDocumento.ENTNotaCreditoImpuesto
                If facImpuesto.Codigo = 3 Then
                    '  BaseImponibleICE += facImpuesto.BaseImponible
                    ICE += facImpuesto.Valor
                End If
            Next

            oForm.Items.Item("txtSub").Enabled = False
            Dim txtSub As SAPbouiCOM.EditText
            txtSub = oForm.Items.Item("txtSub").Specific
            txtSub.Item.RightJustified = True
            txtSub.Value = formatDecimal(Math.Round(BaseImponibleIVA, 2).ToString())
            txtSub.Item.FromPane = 0
            txtSub.Item.ToPane = 0

            oForm.Items.Item("txtSub5").Enabled = False
            Dim txtSub5 As SAPbouiCOM.EditText
            txtSub5 = oForm.Items.Item("txtSub5").Specific
            txtSub5.Item.RightJustified = True
            txtSub5.Value = formatDecimal(Math.Round(BaseImponibleIVA5, 2).ToString())
            txtSub5.Item.FromPane = 1
            txtSub5.Item.ToPane = 1

            oForm.Items.Item("txtSub0").Enabled = False
            Dim txtSub0 As SAPbouiCOM.EditText
            txtSub0 = oForm.Items.Item("txtSub0").Specific
            txtSub0.Item.RightJustified = True
            txtSub0.Value = formatDecimal(Math.Round(BaseImponible0, 2).ToString())
            txtSub0.Item.FromPane = 0
            txtSub0.Item.ToPane = 0

            oForm.Items.Item("txtSubN").Enabled = False
            Dim txtSubN As SAPbouiCOM.EditText
            txtSubN = oForm.Items.Item("txtSubN").Specific
            txtSubN.Item.RightJustified = True
            txtSubN.Value = formatDecimal(Math.Round(BaseImponibleNoObjeto, 2).ToString())
            txtSubN.Item.FromPane = 0
            txtSubN.Item.ToPane = 0

            oForm.Items.Item("txtSubE").Enabled = False
            Dim txtSubE As SAPbouiCOM.EditText
            txtSubE = oForm.Items.Item("txtSubE").Specific
            txtSubE.Item.RightJustified = True
            txtSubE.Value = formatDecimal(Math.Round(BaseImponibleExento, 2).ToString())
            txtSubE.Item.FromPane = 0
            txtSubE.Item.ToPane = 0

            oForm.Items.Item("txtSubS").Enabled = False
            Dim txtSubS As SAPbouiCOM.EditText
            txtSubS = oForm.Items.Item("txtSubS").Specific
            txtSubS.Item.RightJustified = True
            txtSubS.Value = formatDecimal(Math.Round(oDocumento.TotalSinImpuesto, 2).ToString())
            txtSubS.Item.FromPane = 0
            txtSubS.Item.ToPane = 0

            oForm.Items.Item("txtDes").Enabled = False
            Dim txDes As SAPbouiCOM.EditText
            txDes = oForm.Items.Item("txtDes").Specific
            txDes.Item.RightJustified = True
            txDes.Value = formatDecimal(Math.Round(oDocumento.Descuento, 2).ToString())
            txDes.Item.FromPane = 0
            txDes.Item.ToPane = 0

            oForm.Items.Item("txtTotal").Enabled = False
            Dim txtTotal As SAPbouiCOM.EditText
            txtTotal = oForm.Items.Item("txtTotal").Specific
            txtTotal.Item.RightJustified = True
            txtTotal.Value = formatDecimal(Math.Round(oDocumento.ValorModificacion, 2).ToString())
            txtTotal.Item.FromPane = 0
            txtTotal.Item.ToPane = 0

            oForm.Items.Item("txtICE").Enabled = False
            Dim txtICE As SAPbouiCOM.EditText
            txtICE = oForm.Items.Item("txtICE").Specific
            txtICE.Item.RightJustified = True
            txtICE.Value = formatDecimal(Math.Round(ICE, 2).ToString())
            txtICE.Item.FromPane = 0
            txtICE.Item.ToPane = 0

            oForm.Items.Item("txtIva").Enabled = False
            Dim txtIva As SAPbouiCOM.EditText
            txtIva = oForm.Items.Item("txtIva").Specific
            txtIva.Item.RightJustified = True
            txtIva.Value = formatDecimal(Math.Round(Iva, 2).ToString())
            txtIva.Item.FromPane = 1
            txtIva.Item.ToPane = 1

            oForm.Items.Item("txtIva5").Enabled = False
            Dim txtIva5 As SAPbouiCOM.EditText
            txtIva5 = oForm.Items.Item("txtIva5").Specific
            txtIva5.Item.RightJustified = True
            txtIva5.Value = formatDecimal(Math.Round(Iva5, 2).ToString())
            txtIva5.Item.FromPane = 1
            txtIva5.Item.ToPane = 1

            Dim cbxTipo As SAPbouiCOM.ComboBox
            cbxTipo = oForm.Items.Item("cbxTipo").Specific
            cbxTipo.Select("NC Inventariable", SAPbouiCOM.BoSearchKey.psk_ByValue)
            cbxTipo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            OcultarObjetosPorTipo("NC Inventariable")

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
            lnkP.LinkedObjectType = 112
            lnkP.Item.LinkTo = "txtFPre"

            Dim flDetalle As SAPbouiCOM.Folder
            flDetalle = oForm.Items.Item("flDetalle").Specific
            flDetalle.Select()

            'flDetalle.GroupWith("flRelacion")
            'oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2)
            'flDetalle.DataBind.SetBound(True, "", "FolderDS")

            ' GRID DE DOCUMENTOS RELACIONADO
            Try
                oForm.DataSources.DataTables.Add("dtDocr") ' DATA TABLE, PARA DETALLE DE DOCUMENTOS RELACIONADOS
                oForm.DataSources.DataTables.Add("dtDocRel") ' DATA TABLE, PARA RESUMEN DE LOS DOCUMENTOS RELACIONADOS
            Catch ex As Exception
            End Try
            Dim oGri As SAPbouiCOM.Grid = oForm.Items.Item("oGridD").Specific
            oGri.DataTable = oForm.DataSources.DataTables.Item("dtDocr")
            oGri.Item.Enabled = False
            CargaGridRelacionado()
            ' END GRID DE DOCUMENTOS RELACIONADO

            'oForm.Width = 747
            'oForm.Height = 619
            LLenarGridDatosAdicionales(oDocumento.ENTDatoAdicionalNotaCredito.ToList)

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
        If RecorreFormulario(rsboApp, "frmDocumentoNC") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentoNC.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentoNC").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmDocumentoNC")
            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO
            oForm.Freeze(True)

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
                QueryCabecera += " ,""U_Mapeado"" ,""U_ClaAcc"" ,""U_NumAut"" "
                QueryCabecera += " ,""U_FecAut"" ,""U_NumDoc"" ,""U_FPrelim"""
                QueryCabecera += " ,""U_SubTot"",""U_SubTot5"" ,""U_Sub0"" ,""U_SubNO"""
                QueryCabecera += " ,""U_SubEx"" ,""U_SubSI"" ,""U_Desc"""
                QueryCabecera += " ,""U_ICE"" ,""U_IVA"",""U_IVA5"" ,""U_vTotal"""
                QueryCabecera += " ,""U_rTades"" ,""U_rPDesc"" ,""U_rDesc"""
                QueryCabecera += " ,""U_rGast"" ,""U_rImp"" ,""U_rTotal"""
                QueryCabecera += " ,""U_Tipo"" ,""U_IdGS"" ,""U_Sincro"""
                QueryCabecera += " ,""U_SincroE"" ,""U_Estado"" ,""U_FechaS"""
                QueryCabecera += "  FROM ""@GS_NCR"" "
                QueryCabecera += "  WHERE ""DocEntry"" =  " + IdDocumentoRecibido_UDO
            Else
                QueryCabecera = " SELECT U_RUC ,U_Nombre ,U_CardCode"
                QueryCabecera += " ,U_Mapeado ,U_ClaAcc ,U_NumAut "
                QueryCabecera += " ,U_FecAut ,U_NumDoc ,U_FPrelim"
                QueryCabecera += " ,U_SubTot,U_SubTot5 ,U_Sub0 ,U_SubNO"
                QueryCabecera += " ,U_SubEx ,U_SubSI ,U_Desc"
                QueryCabecera += " ,U_ICE ,U_IVA,U_IVA5 ,U_vTotal"
                QueryCabecera += " ,U_rTades ,U_rPDesc ,U_rDesc"
                QueryCabecera += " ,U_rGast ,U_rImp ,U_rTotal"
                QueryCabecera += " ,U_Tipo ,U_IdGS ,U_Sincro"
                QueryCabecera += " ,U_SincroE ,U_Estado ,U_FechaS"
                QueryCabecera += "  FROM ""@GS_NCR"" A WITH(NOLOCK)"
                QueryCabecera += "  WHERE A.DocEntry =  " + IdDocumentoRecibido_UDO
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocCAB").ExecuteQuery(QueryCabecera)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString() + " - Query: " + QueryCabecera, "frmDocumentoNC")
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

                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = oForm.Items.Item("lbMapp").Specific
                lbMapp.Value = odt.GetValue("U_Mapeado", i).ToString()

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
                txtNumDoc.Value = odt.GetValue("U_NumDoc", i).ToString() '_oDocumento.Establecimiento + "-" + _oDocumento.PuntoEmision + "-" + _oDocumento.Secuencial

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
                txtSub.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_SubTot", i).ToString()), 2).ToString())
                txtSub.Item.FromPane = 0
                txtSub.Item.ToPane = 0

                oForm.Items.Item("txtSub5").Enabled = False
                Dim txtSub5 As SAPbouiCOM.EditText
                txtSub5 = oForm.Items.Item("txtSub5").Specific
                txtSub5.Item.RightJustified = True
                txtSub5.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_SubTot5", i).ToString()), 2).ToString())
                txtSub5.Item.FromPane = 1
                txtSub5.Item.ToPane = 1

                oForm.Items.Item("txtSub0").Enabled = False
                Dim txtSub0 As SAPbouiCOM.EditText
                txtSub0 = oForm.Items.Item("txtSub0").Specific
                txtSub0.Item.RightJustified = True
                txtSub0.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_Sub0", i).ToString()), 2).ToString())
                txtSub0.Item.FromPane = 0
                txtSub0.Item.ToPane = 0

                oForm.Items.Item("txtSubN").Enabled = False
                Dim txtSubN As SAPbouiCOM.EditText
                txtSubN = oForm.Items.Item("txtSubN").Specific
                txtSubN.Item.RightJustified = True
                txtSubN.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_SubNO", i).ToString()), 2).ToString())
                txtSubN.Item.FromPane = 0
                txtSubN.Item.ToPane = 0

                oForm.Items.Item("txtSubE").Enabled = False
                Dim txtSubE As SAPbouiCOM.EditText
                txtSubE = oForm.Items.Item("txtSubE").Specific
                txtSubE.Item.RightJustified = True
                txtSubE.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_SubEx", i).ToString()), 2).ToString())
                txtSubE.Item.FromPane = 0
                txtSubE.Item.ToPane = 0

                oForm.Items.Item("txtSubS").Enabled = False
                Dim txtSubS As SAPbouiCOM.EditText
                txtSubS = oForm.Items.Item("txtSubS").Specific
                txtSubS.Item.RightJustified = True
                txtSubS.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_SubSI", i).ToString()), 2).ToString())
                txtSubS.Item.FromPane = 0
                txtSubS.Item.ToPane = 0

                oForm.Items.Item("txtDes").Enabled = False
                Dim txDes As SAPbouiCOM.EditText
                txDes = oForm.Items.Item("txtDes").Specific
                txDes.Item.RightJustified = True
                txDes.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_Desc", i).ToString()), 2).ToString())
                txDes.Item.FromPane = 0
                txDes.Item.ToPane = 0

                oForm.Items.Item("txtTotal").Enabled = False
                Dim txtTotal As SAPbouiCOM.EditText
                txtTotal = oForm.Items.Item("txtTotal").Specific
                txtTotal.Item.RightJustified = True
                txtTotal.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_vTotal", i).ToString()), 2).ToString())
                txtTotal.Item.FromPane = 0
                txtTotal.Item.ToPane = 0

                oForm.Items.Item("txtICE").Enabled = False
                Dim txtICE As SAPbouiCOM.EditText
                txtICE = oForm.Items.Item("txtICE").Specific
                txtICE.Item.RightJustified = True
                txtICE.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_ICE", i).ToString()), 2).ToString())
                txtICE.Item.FromPane = 0
                txtICE.Item.ToPane = 0

                oForm.Items.Item("txtIva").Enabled = False
                Dim txtIva As SAPbouiCOM.EditText
                txtIva = oForm.Items.Item("txtIva").Specific
                txtIva.Item.RightJustified = True
                txtIva.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_IVA", i).ToString()), 2).ToString())
                txtIva.Item.FromPane = 1
                txtIva.Item.ToPane = 1

                oForm.Items.Item("txtIva5").Enabled = False
                Dim txtIva5 As SAPbouiCOM.EditText
                txtIva5 = oForm.Items.Item("txtIva5").Specific
                txtIva5.Item.RightJustified = True
                txtIva5.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_IVA5", i).ToString()), 2).ToString())
                txtIva5.Item.FromPane = 1
                txtIva5.Item.ToPane = 1

                Dim cbxTipo As SAPbouiCOM.ComboBox
                cbxTipo = oForm.Items.Item("cbxTipo").Specific
                cbxTipo.Select(odt.GetValue("U_Tipo", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                cbxTipo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                _TipoDocumento = odt.GetValue("U_Tipo", i).ToString()

                oForm.Items.Item("txtF").Enabled = True
                Dim Focus As SAPbouiCOM.EditText
                Focus = oForm.Items.Item("txtF").Specific
                'txtF.Item.RightJustified = True
                Focus.Value = "0"

                cbxTipo.Item.Enabled = False

                oForm.Items.Item("txtFPre").Enabled = False
                Dim txtFPre As SAPbouiCOM.EditText
                txtFPre = oForm.Items.Item("txtFPre").Specific
                txtFPre.Value = odt.GetValue("U_FPrelim", i).ToString()

                'SETEO TOTALES DE DOCUMENTOS RELACIONADOS
                Dim txtDTot As SAPbouiCOM.EditText
                txtDTot = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDTot").Specific
                txtDTot.Item.RightJustified = True
                txtDTot.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_rTades", i).ToString()), 2).ToString())
                Dim txtDP As SAPbouiCOM.EditText
                txtDP = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDP").Specific
                txtDP.Item.RightJustified = True
                txtDP.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_rPDesc", i).ToString()), 2).ToString())
                Dim txtDVP As SAPbouiCOM.EditText
                txtDVP = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDVP").Specific
                txtDVP.Item.RightJustified = True
                txtDVP.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_rDesc", i).ToString()), 2).ToString())
                Dim txtDG As SAPbouiCOM.EditText
                txtDG = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDG").Specific
                txtDG.Item.RightJustified = True
                txtDG.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_rGast", i).ToString()), 2).ToString())
                Dim txtDI As SAPbouiCOM.EditText
                txtDI = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDI").Specific
                txtDI.Item.RightJustified = True
                txtDI.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_rImp", i).ToString()), 2).ToString())
                Dim txtDT As SAPbouiCOM.EditText
                txtDT = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDT").Specific
                txtDT.Item.RightJustified = True
                txtDT.Value = formatDecimal(Math.Round(Convert.ToDouble(odt.GetValue("U_rTotal", i).ToString()), 2).ToString())
            Next

            ' DATA TABLE DETALLE
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try
            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodPrin", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodAuxi", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodSAP", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Descripc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Cantid", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Precio", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Desc", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
            Dim QueryDetalle As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryDetalle = "SELECT  A.""U_CodPrin"", A.""U_CodAuxi"", A.""U_CodSAP"", A.""U_Descripc"", A.""U_Cantid"", A.""U_Precio"", A.""U_Desc"", A.""U_Total"" "
                QueryDetalle += "  FROM ""@GS0_NCR"" A "
                QueryDetalle += "  WHERE A.""DocEntry"" =  " + IdDocumentoRecibido_UDO
            Else
                QueryDetalle = "SELECT  A.U_CodPrin, A.U_CodAuxi, A.U_CodSAP, A.U_Descripc, A.U_Cantid, A.U_Precio, A.U_Desc, A.U_Total "
                QueryDetalle += "  FROM ""@GS0_NCR"" A WITH(NOLOCK)"
                QueryDetalle += "  WHERE A.DocEntry =  " + IdDocumentoRecibido_UDO
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocs").ExecuteQuery(QueryDetalle)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString() + " - Query: " + QueryDetalle, "frmDocumentoNC")
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.Item.Enabled = False
            oGrid.Item.FromPane = 0
            oGrid.Item.ToPane = 0

            ' Codigo
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Caption = "Codigo"
            ' Codigo Auxiliar
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Codigo Auxiliar"
            ' ItemCode
            oGrid.Columns.Item(2).Description = "Cod.SAP"
            oGrid.Columns.Item(2).TitleObject.Caption = "Cod.SAP"
            oGrid.Columns.Item(2).Editable = False
            Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            oEditTextColum = oGrid.Columns.Item(2)
            oEditTextColum.LinkedObjectType = 4
            ' Descripcion
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Caption = "Descripcion"
            ' Cantidad
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Caption = "Cantidad"
            oGrid.Columns.Item(4).RightJustified = True
            ' Precio
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Precio"
            oGrid.Columns.Item(5).RightJustified = True
            ' Descuento
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Caption = "Descuento"
            oGrid.Columns.Item(6).RightJustified = True
            ' Total
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Caption = "Total"
            oGrid.Columns.Item(7).RightJustified = True

            'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            'oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            'oGrid.Item.Enabled = False
            'oGrid.Columns.Item(2).Description = "Cod.SAP"
            'oGrid.Columns.Item(2).TitleObject.Caption = "Cod.SAP"
            'oGrid.Columns.Item(2).Editable = False
            'Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            'oEditTextColum = oGrid.Columns.Item(2)
            'oEditTextColum.LinkedObjectType = 4
            ' END DATA TABLE DETALLE

            ' DATA TABLE, PARA DETALLE DE DOCUMENTOS RELACIONADOS
            Try
                oForm.DataSources.DataTables.Add("dtDocr")
            Catch ex As Exception
            End Try
            oForm.DataSources.DataTables.Item("dtDocr").Clear()
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("Dscription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("Price", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("DiscPrcnt", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("TaxCode", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("LineTotal", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10)
            oForm.DataSources.DataTables.Item("dtDocr").Rows.Clear()
            Dim QueryDetalleRelacionados As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryDetalleRelacionados = "SELECT  A.""U_DocEntr"", A.""U_LineNu"", A.""U_ItemCode"", A.""U_Descripc"", A.""U_Cantid"", A.""U_Precio"", A.""U_DiscPr"", A.""U_TaxCode"", A.""U_lTotal"", A.""U_ObjType"" "
                QueryDetalleRelacionados += "  FROM ""@GS1_NCR"" A "
                QueryDetalleRelacionados += "  WHERE A.""DocEntry"" =  " + IdDocumentoRecibido_UDO
            Else
                QueryDetalleRelacionados = "SELECT   A.U_DocEntr, A.U_LineNu, A.U_ItemCode, A.U_Descripc, A.U_Cantid, A.U_Precio, A.U_DiscPr, A.U_TaxCode, A.U_lTotal, A.U_ObjType "
                QueryDetalleRelacionados += "  FROM ""@GS1_NCR"" A WITH(NOLOCK)"
                QueryDetalleRelacionados += "  WHERE A.DocEntry =  " + IdDocumentoRecibido_UDO
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocr").ExecuteQuery(QueryDetalleRelacionados)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString() + " - Query: " + QueryDetalleRelacionados, "frmDocumentoNC")
            End Try

            Dim oGridD As SAPbouiCOM.Grid = oForm.Items.Item("oGridD").Specific
            oGridD.DataTable = oForm.DataSources.DataTables.Item("dtDocr")
            oGridD.Item.Enabled = False

            ' OBTENGO EL PRIMER OBJETYPE DEL DETALLE RELACIONADO PARA PONER EL LINK BUTTON
            Dim ObjType As String = ""
            ObjType = oFuncionesB1.getRSvalue(QueryDetalleRelacionados, "U_ObjType", "")

            If ObjType = "" Then
                ObjType = 0
            End If


            If ObjType = 21 Then
                'oGri.Columns.Item(0).Visible = False
                oGridD.Columns.Item(0).TitleObject.Caption = "Devoluciones"
                Dim oEditTextColump As SAPbouiCOM.EditTextColumn
                oEditTextColump = oGridD.Columns.Item(0)
                oEditTextColump.LinkedObjectType = 21
            ElseIf ObjType = 18 Then
                'oGri.Columns.Item(0).Visible = False
                oGridD.Columns.Item(0).TitleObject.Caption = "Facturas"
                Dim oEditTextColump As SAPbouiCOM.EditTextColumn
                oEditTextColump = oGridD.Columns.Item(0)
                oEditTextColump.LinkedObjectType = 18
            ElseIf ObjType = 204 Then
                'oGri.Columns.Item(0).Visible = False
                oGridD.Columns.Item(0).TitleObject.Caption = "Anticipos"
                Dim oEditTextColump As SAPbouiCOM.EditTextColumn
                oEditTextColump = oGridD.Columns.Item(0)
                oEditTextColump.LinkedObjectType = 204
            End If

            oGridD.Columns.Item(1).TitleObject.Caption = "Linea"
            oGridD.Columns.Item(2).Description = "Número de Artículo"
            oGridD.Columns.Item(2).TitleObject.Caption = "Número de Artículo"
            oGridD.Columns.Item(2).Editable = False
            Dim oEditTextColumD As SAPbouiCOM.EditTextColumn
            oEditTextColumD = oGridD.Columns.Item(2)
            oEditTextColumD.LinkedObjectType = 4

            oGridD.Columns.Item(3).TitleObject.Caption = "Descripción"
            oGridD.Columns.Item(4).TitleObject.Caption = "Cantidad"
            oGridD.Columns.Item(5).TitleObject.Caption = "Precio"
            oGridD.Columns.Item(6).TitleObject.Caption = "Descuento"
            oGridD.Columns.Item(7).TitleObject.Caption = "Impuesto"
            oGridD.Columns.Item(8).TitleObject.Caption = "Total"
            oGridD.Columns.Item(9).Visible = False


            ' DATA TABLE, PARA DETALLE DE DOCUMENTOS RELACIONADOS

            Dim lnkP As SAPbouiCOM.LinkedButton
            lnkP = oForm.Items.Item("lnkP").Specific
            lnkP.LinkedObjectType = 112
            lnkP.Item.LinkTo = "txtFPre"

            Dim flDetalle As SAPbouiCOM.Folder
            flDetalle = oForm.Items.Item("flDetalle").Specific
            flDetalle.Select()

            oForm.Items.Item("obtnGrabar").Visible = False
            oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
            Dim oB As SAPbouiCOM.Button
            oB = oForm.Items.Item("2").Specific
            oB.Caption = "OK"

            OcultarObjetosPorTipo(_TipoDocumento)

            Dim btnMapear As SAPbouiCOM.Button
            btnMapear = oForm.Items.Item("btnMapear").Specific
            btnMapear.Item.Enabled = False

            Dim obuttonCombo As SAPbouiCOM.ButtonCombo
            obuttonCombo = oForm.Items.Item("cbEnlazar").Specific
            obuttonCombo.Item.Enabled = False

            LLenarGridDatosAdicionalesExistente(IdDocumentoRecibido_UDO)
            'oForm.Width = 747
            'oForm.Height = 619

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - Error al cargar formulario recibido existente: " + ex.Message.ToString())
        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Public Sub CargaGridRelacionado()
        Try
            oForm.Freeze(True)
            oForm.DataSources.DataTables.Item("dtDocr").Clear()
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("Dscription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("Price", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("DiscPrcnt", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("TaxCode", SAPbouiCOM.BoFieldsType.ft_Float, 100)
            oForm.DataSources.DataTables.Item("dtDocr").Columns.Add("LineTotal", SAPbouiCOM.BoFieldsType.ft_Float, 100)

            Dim oGridD As SAPbouiCOM.Grid = oForm.Items.Item("oGridD").Specific
            'oGridD.DataTable = dt
            'oGridD.Item.Enabled = False

            oGridD.Columns.Item(2).Description = "Número de Artículo"
            oGridD.Columns.Item(2).TitleObject.Caption = "Número de Artículo"
            oGridD.Columns.Item(2).Editable = False
            Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            oEditTextColum = oGridD.Columns.Item(2)
            oEditTextColum.LinkedObjectType = 4

            oForm.Freeze(False)
        Catch ex As Exception

        End Try

    End Sub

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

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent
        Try
            Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            If typeEx = "frmDocumentoNC" Then
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
                                    ConsutarPDFRecibido(_IdGS, 3)
                                Case "cbEnlazar"
                                    rsboApp.StatusBar.SetText("  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None) ' Se lo usa para blanquear el mensaje anterior
                                    Dim cbEnlazar As SAPbouiCOM.ButtonCombo
                                    cbEnlazar = oForm.Items.Item("cbEnlazar").Specific
                                    If Not cbEnlazar.Selected Is Nothing Then
                                        'cbEnlazar.ValidValues.Add("Devolución de Mercadería", "Devolución de Mercadería")
                                        'cbEnlazar.ValidValues.Add("Factura de Proveedores", "Factura de Proveedores")
                                        'cbEnlazar.ValidValues.Add("Anticipo de Proveedores", "Anticipo de Proveedores")

                                        If cbEnlazar.Selected.Value = "Devolución de Mercadería" Then ' tabla: ORPD objectype: 21
                                            ofrmConsultaOrdenes.CargaFormularioConsulta(oCardCode, "21", "NC")
                                            ObjTypeRelacionado = 21
                                        ElseIf cbEnlazar.Selected.Value = "Factura de Proveedores" Then ' tabla: OPCH objectype: 18
                                            ofrmConsultaOrdenes.CargaFormularioConsulta(oCardCode, "18", "NC")
                                            ObjTypeRelacionado = 18
                                        ElseIf cbEnlazar.Selected.Value = "Anticipo de Proveedores" Then ' tabla: ODPO objectype: 204
                                            ofrmConsultaOrdenes.CargaFormularioConsulta(oCardCode, "204", "NC")
                                            ObjTypeRelacionado = 204
                                        End If
                                        cbEnlazar.Caption = "Relacionar a :"
                                    End If

                                Case "cbxTipo"
                                    Dim cbxTipo As SAPbouiCOM.ComboBox
                                    cbxTipo = oForm.Items.Item("cbxTipo").Specific
                                    OcultarObjetosPorTipo(cbxTipo.Value)


                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "flDetalle"
                                    oForm.Freeze(True)
                                    oForm.Items.Item("txtSub").Visible = True
                                    oForm.Items.Item("txtSub0").Visible = True
                                    oForm.Items.Item("txtSubN").Visible = True
                                    oForm.Items.Item("txtSubE").Visible = True
                                    oForm.Items.Item("txtSubS").Visible = True
                                    oForm.Items.Item("txtDes").Visible = True
                                    oForm.Items.Item("txtICE").Visible = True
                                    oForm.Items.Item("txtIva").Visible = True
                                    oForm.Items.Item("txtIva5").Visible = True
                                    oForm.Items.Item("txtTotal").Visible = True
                                    oForm.Items.Item("oGrid").Visible = True
                                    oForm.Freeze(False)
                                Case "flRelacion"
                                    oForm.Freeze(True)
                                    oForm.Items.Item("txtSub").Visible = False
                                    oForm.Items.Item("txtSub0").Visible = False
                                    oForm.Items.Item("txtSubN").Visible = False
                                    oForm.Items.Item("txtSubE").Visible = False
                                    oForm.Items.Item("txtSubS").Visible = False
                                    oForm.Items.Item("txtDes").Visible = False
                                    oForm.Items.Item("txtICE").Visible = False
                                    oForm.Items.Item("txtIva").Visible = False
                                    oForm.Items.Item("txtIva5").Visible = False
                                    oForm.Items.Item("txtTotal").Visible = False
                                    oForm.Items.Item("oGridD").Visible = True
                                    oForm.Freeze(False)

                                Case "flDatosAd"
                                    oForm.Freeze(True)
                                    oForm.Items.Item("txtSub").Visible = False
                                    oForm.Items.Item("txtSub0").Visible = False
                                    oForm.Items.Item("txtSubN").Visible = False
                                    oForm.Items.Item("txtSubE").Visible = False
                                    oForm.Items.Item("txtSubS").Visible = False
                                    oForm.Items.Item("txtDes").Visible = False
                                    oForm.Items.Item("txtICE").Visible = False
                                    oForm.Items.Item("txtIva").Visible = False
                                    oForm.Items.Item("txtIva5").Visible = False
                                    oForm.Items.Item("txtTotal").Visible = False
                                    oForm.Items.Item("oGrid").Visible = False
                                    oForm.Items.Item("oGridD").Visible = False

                                    oForm.Freeze(False)

                                Case "btnMapear"
                                    rsboApp.StatusBar.SetText("  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None) ' Se lo usa para blanquear el mensaje anterior

                                    Dim sClaveAcceso As String = _oDocumento.ClaveAcceso
                                    Dim sNombre As String = _oDocumento.RazonSocial

                                    listaDetalleArtiulos.Clear()
                                    For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito In _oDocumento.ENTDetalleNotaCredito
                                        listaDetalleArtiulos.Add(New Entidades.DetalleArticulo(sClaveAcceso, oDetalle.CodigoPrincipal, _
                                                                                               oDetalle.CodigoAuxiliar, oDetalle.Descripcion, oDetalle.Cantidad, oDetalle.PrecioUnitario, oDetalle.Descuento, oDetalle.PrecioTotalSinImpuesto))
                                    Next
                                    ofrmMapeo.CargaFormularioMapeo(_oDocumento.Ruc, oCardCode, sNombre, listaDetalleArtiulos, _fila, "NC")

                                Case "obtnGrabar"

                                    'guardar pdf en ruta compartida parametrizada
                                    Dim rl As String = ""
                                    rl = Functions.VariablesGlobales._RecepcionLite
                                    If rl = "Y" Then
                                        GuardarArchivoNCRecibido(_IdGS, 3)
                                    End If

                                    Dim obtnGrabar As SAPbouiCOM.Button
                                    obtnGrabar = oForm.Items.Item("obtnGrabar").Specific
                                    Dim cbxTipo As SAPbouiCOM.ComboBox
                                    cbxTipo = oForm.Items.Item("cbxTipo").Specific

                                    Dim Exitoso As Boolean = False
                                    Dim sDocEntryPreliminar As String = "0"
                                    If cbxTipo.Value = "NC Inventariable Relacionada" Then '"Documento Relacionado"

                                        If oForm.DataSources.DataTables.Item("dtDocr").Rows.Count > 0 Then
                                            Dim mensajeValidacionTotales As String = ""

                                            Dim PermiteDescuadre As String = ""
                                            PermiteDescuadre = Functions.VariablesGlobales._PermiteDescuadre
                                            If PermiteDescuadre = "N" Then
                                                ' VALIDAR TOTALES

                                                Dim txtTotalFE As SAPbouiCOM.EditText = oForm.Items.Item("txtTotal").Specific
                                                Dim txtTotalRE As SAPbouiCOM.EditText = oForm.Items.Item("txtDT").Specific
                                                Dim dTotalFE As Decimal = formatDecimal(txtTotalFE.Value)
                                                Dim dTotalRE As Decimal = formatDecimal(txtTotalRE.Value)
                                                If dTotalFE > dTotalRE Then
                                                    mensajeValidacionTotales = "El valor de la Nota de Crédito recibida es MAYOR que los documentos relacionados, por : " + (dTotalFE - dTotalRE).ToString()
                                                ElseIf dTotalFE < dTotalRE Then
                                                    mensajeValidacionTotales = "El valor de la Nota de Crédito recibida es MENOR que los documentos relacionados, por : " + (dTotalRE - dTotalFE).ToString()
                                                Else
                                                    ' GUARDO EL DOCUMENTO RECIBIDO EN EL UDO FACTURA RECIBIDA
                                                    Dim DocEntryNCRecibida_UDO As String = 0
                                                    If Guarda_DocumentoRecibido_NC(DocEntryNCRecibida_UDO) Then
                                                        Exitoso = CrearNCPreliminarRelacionada(sDocEntryPreliminar, DocEntryNCRecibida_UDO)
                                                    End If

                                                End If
                                            Else
                                                ' GUARDO EL DOCUMENTO RECIBIDO EN EL UDO FACTURA RECIBIDA
                                                Dim DocEntryNCRecibida_UDO As String = 0
                                                If Guarda_DocumentoRecibido_NC(DocEntryNCRecibida_UDO) Then
                                                    Exitoso = CrearNCPreliminarRelacionada(sDocEntryPreliminar, DocEntryNCRecibida_UDO)
                                                End If
                                            End If

                                            If mensajeValidacionTotales <> "" Then
                                                'Dim iReturnValue As Integer
                                                'iReturnValue = rsboApp.MessageBox("GS " + mensajeValidacionTotales + ", Desea Continuar", 1, "&SI", "&NO")
                                                'If iReturnValue = 1 Then
                                                '    Dim iReturnValue2 As Integer
                                                '    iReturnValue2 = rsboApp.MessageBox("GS - Se Creará la Factura Preliminar en base a los DOCUMENTO RELACIONADO, Desea Continuar", 1, "&SI", "&NO")
                                                '    If iReturnValue2 = 1 Then
                                                '        rsboApp.StatusBar.SetText("GS - Creando Factura por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                '        Exitoso = CrearFacturaPreliminarRelacionada()
                                                '    Else
                                                '        Exitoso = False
                                                '    End If
                                                'Else
                                                '    Exitoso = False
                                                'End If
                                                rsboApp.StatusBar.SetText(NombreAddon + " - " + mensajeValidacionTotales, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If

                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - No existe ningun documento relacionado, para poder usar esta opción!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exitoso = False
                                        End If
                                    ElseIf cbxTipo.Value = "NC Inventariable" Then '"Mapeo de Items"
                                        Dim txtMapeo As SAPbouiCOM.EditText = oForm.Items.Item("lbMapp").Specific
                                        If txtMapeo.Value = "SI" Then
                                            Dim iReturnValue As Integer
                                            iReturnValue = rsboApp.MessageBox(NombreAddon + " -  Se Creará la Nota de Crédito Preliminar en base al MAPEO DE ITEMS, Desea Continuar", 1, "&SI", "&NO")
                                            If iReturnValue = 1 Then
                                                rsboApp.StatusBar.SetText(NombreAddon + " -  Creando la Nota de Crédito por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                Dim DocEntryNCRecibida_UDO As String = 0
                                                If Guarda_DocumentoRecibido_NC(DocEntryNCRecibida_UDO) Then
                                                    Exitoso = CrearNCPreliminarMapeada(oCardCode, _oDocumento, sDocEntryPreliminar, DocEntryNCRecibida_UDO)
                                                End If
                                            Else
                                                Exitoso = False
                                            End If
                                        Else
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Los items del documento no se encuentran Mapeados, para poder usar esta opción!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Exitoso = False
                                        End If

                                    ElseIf cbxTipo.Value = "NC de Servicio" Then '"Servicio"
                                        Dim iReturnValue As Integer
                                        iReturnValue = rsboApp.MessageBox(NombreAddon + " - Se Creará la Nota de Crédito Preliminar de SERVICIO, Desea Continuar", 1, "&SI", "&NO")
                                        If iReturnValue = 1 Then
                                            rsboApp.StatusBar.SetText(NombreAddon + "- Creando la Nota de Crédito por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            Dim DocEntryNCRecibida_UDO As String = 0
                                            If Guarda_DocumentoRecibido_NC(DocEntryNCRecibida_UDO) Then
                                                Exitoso = CrearNCPremilinarServicio(oCardCode, _oDocumento, sDocEntryPreliminar, DocEntryNCRecibida_UDO)
                                            End If
                                        Else
                                            Exitoso = False
                                        End If

                                    End If

                                    If Exitoso = True Then
                                        ' BLOQUEO BOTON GRABAR
                                        oForm.Items.Item("obtnGrabar").Visible = False
                                        oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                        Dim oB As SAPbouiCOM.Button
                                        oB = oForm.Items.Item("2").Specific
                                        oB.Caption = "OK"
                                        ' MUESTRO EL LINK BUTTON DE FACTURA PRELIMINAR 
                                        ' Asigno el docentry de la factura preliminar guardada.
                                        ' se busca por el numero de autorizacion top 1 descendente
                                        oForm.Items.Item("Item_23").Visible = True
                                        oForm.Items.Item("lnkP").Visible = True
                                        oForm.Items.Item("txtFPre").Visible = True
                                        Try
                                            cbxTipo.Active = False
                                            cbxTipo.Item.Enabled = False
                                        Catch ex As Exception
                                        End Try

                                        Dim txtFPre As SAPbouiCOM.EditText
                                        txtFPre = oForm.Items.Item("txtFPre").Specific
                                        txtFPre.Value = sDocEntryPreliminar

                                        ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                        rsboApp.Forms.Item("frmDocumentosRecibidos").Freeze(True)
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidos").DataSources.DataTables.Item("dtDocs")
                                        odt.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                        rsboApp.Forms.Item("frmDocumentosRecibidos").Freeze(False)

                                        oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, " Proceso terminado Exitosamente!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                                        rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                    End If
                            End Select
                        End If

                End Select
            End If

        Catch ex As Exception
            rsboApp.MessageBox("Error en: rSboApp_ItemEvent," + ex.Message().ToString(), NombreAddon)
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

            If BusinessObjectInfo.FormTypeEx = "181" Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If Not BusinessObjectInfo.BeforeAction Then
                            Select Case BusinessObjectInfo.ActionSuccess
                                Case True
                                    oDocumentoSAP = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                                    oDocumentoSAP.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                    If Not oDocumentoSAP.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation Then
                                        If Not oDocumentoSAP.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then
                                            Dim CREADAPORGSEDOC As String = "NO"
                                            Dim idDocumentoRecibido_UDO As String = ""
                                            Try
                                                CREADAPORGSEDOC = oDocumentoSAP.UserFields.Fields.Item("U_SSCREADAR").Value.ToString()
                                                idDocumentoRecibido_UDO = oDocumentoSAP.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                            Catch ex As Exception
                                            End Try
                                            If CREADAPORGSEDOC = "SI" Then
                                                ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                                Dim idFacturaGS As String = ""
                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                    idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_NCR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                                Else
                                                    idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_NCR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                                End If
                                                '' RECUPERO EL ID DE LA FACTURA PRELIMINAR SAP, PARA CERRARLA
                                                'Dim DocEntryPreliminar As String = ""
                                                'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                '    DocEntryPreliminar = oFuncionesB1.getRSvalue("SELECT ""U_FPrelim"" FROM ""@GS_NCR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO, "U_FPrelim", "")
                                                'Else
                                                '    DocEntryPreliminar = oFuncionesB1.getRSvalue(" select U_FPrelim from ""@GS_NCR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_FPrelim", "")
                                                'End If
                                                ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                                '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                    _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_NCR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                                Else
                                                    _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_NCR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                                End If
                                                ' LE CAMBIA EL ESTADO A LA FACTURA UDO A DOCFINAL
                                                ActualizadoEstado_DocumentoRecibido_NC(idDocumentoRecibido_UDO, "docFinal")
                                                ' ACTUALIZA EL CAMPO SINCRO A 1, ESTE CAMPO IDENTIFICA QUE YA ESTA SINCRONIZADA EN SAP
                                                ActualizadoEstadoSincronizado_DocumentoRecibido_NC(idDocumentoRecibido_UDO, 1)
                                                ' MARCA EL DOCUMENTO COMO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                                MarcarVisto(Integer.Parse(idFacturaGS), 3, mensaje, idDocumentoRecibido_UDO)
                                                ' EL WINDOWS SERVICE DEBE SIEMPRE TOMAR COMO REFERENCIA EL CAMPO SINCRO, Y ENVIAR A EDOC LO QUE TENGA EL CAMPO SINCRO.
                                                ' ES DECIR SI EL CAMPO SINCRO ES IGUAL A 1, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO TRUE
                                                ' SI EL CAMPO ES IGUAL A 0, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO FALSE

                                                ' SI LA PANTALLA DE DOCUMENTOS RECIBIDOS ESTA ABIERTA ELIMINO LA LINEA DE LA FACTURA RECIBIDA
                                                ' YA QUE YA ESTA INTEGRADA
                                                Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                                                    If rsboApp.Forms.Item("frmDocumentosRecibidos").Visible = True Then
                                                        rsboApp.Forms.Item("frmDocumentosRecibidos").Freeze(True)
                                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidos").DataSources.DataTables.Item("dtDocs")
                                                        odt.Rows.Remove(_fila)
                                                        rsboApp.Forms.Item("frmDocumentosRecibidos").Freeze(False)
                                                    End If
                                                Catch ex As Exception
                                                End Try

                                                'oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Documento Preliminar # " + DocEntryPreliminar.ToString() + ", cerrado Satisfactoriamente!!", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)

                                                '' NO ES NECESARIO CERRAR EL DOCUMENTO PRELIMINAR, YA QUE AL ATARLO A LA FACTURA FIJA, YA NO QUEDA PENDIENTE
                                                '' OBTENGO EL DRAF Y LO CIERRO
                                                'oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, " Obteniendo Documento Preliminar # " + DocEntryPreliminar.ToString() + ", para cerrarlo", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                'Dim oDraft As SAPbobsCOM.Documents
                                                'Try
                                                '    oDraft = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                                '    If oDraft.GetByKey(CInt(DocEntryPreliminar)) Then
                                                '        Dim RetVal As Long
                                                '        Dim ErrCode As Long = 0
                                                '        Dim ErrMsg As String = ""
                                                '        RetVal = oDraft.Close()
                                                '        If RetVal <> 0 Then
                                                '            rCompany.GetLastError(ErrCode, ErrMsg)
                                                '            oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, " Error al Cerrar Documento Preliminar # " + DocEntryPreliminar.ToString() + ", " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                '        Else
                                                '            oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, " Documento Preliminar # " + DocEntryPreliminar.ToString() + ", cerrado Satisfactoriamente!!", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                '        End If
                                                '    End If
                                                'Catch ex As Exception
                                                '    rsboApp.StatusBar.SetText(NombreAddon + " - Error al cerrar el documento preliminar : " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                '    oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, " Error al Cerrar Documento Preliminar # " + DocEntryPreliminar.ToString() + ": " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                'Finally
                                                '    oDraft = Nothing
                                                'End Try


                                            End If
                                        End If
                                    Else ' SI ES UNA CANCELACION CAMBIO EL ESTADO EN EDOC A NO SINCRONIZADO
                                        Dim CREADAPORGSEDOC As String = "NO"
                                        Dim idDocumentoRecibido_UDO As String = ""
                                        Try
                                            CREADAPORGSEDOC = oDocumentoSAP.UserFields.Fields.Item("U_SSCREADAR").Value.ToString()
                                            idDocumentoRecibido_UDO = oDocumentoSAP.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                        Catch ex As Exception
                                        End Try
                                        If CREADAPORGSEDOC = "SI" Then
                                            ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                            Dim idFacturaGS As String = ""
                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_NCR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                            Else
                                                idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_NCR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                            End If
                                            ' RECUPERO LA CLAVE DE ACCESO - CLAVE DE ACCESO ES UN VARIABLE GLOBAL, QUE SE USA EN FUNCIONES COMO MARCARVISTO
                                            '                             - SE LA VUELVE A SETEAR YA QUE ESTE EVENTO PUEDE GENERARSE SIN EMPEZAR POR LA CREACION DEL PRELIMINAR
                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                _ClaveAcceso = oFuncionesB1.getRSvalue("SELECT ""U_ClaAcc"" FROM ""@GS_NCR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_ClaAcc", "")
                                            Else
                                                _ClaveAcceso = oFuncionesB1.getRSvalue(" select U_ClaAcc from ""@GS_NCR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_ClaAcc", "")
                                            End If
                                            ActualizadoEstado_DocumentoRecibido_NC(idDocumentoRecibido_UDO, "docCancelado")

                                            ' ACTUALIZA EL CAMPO SINCRO A 0, AL CANCELARLO EL SE DEBE MARCAR COMO NO SINCRONIZADO
                                            ActualizadoEstadoSincronizado_DocumentoRecibido_NC(idDocumentoRecibido_UDO, 0)

                                            ' MARCA EL DOCUMENTO COMO NO VISTO(SINCRONIZADO) EN EDOC A TRAVEZ DEL WS, SI DA ERROR UN WINDOWS SERVICE DEBE REPROCESARLO
                                            MarcarNOVisto(Integer.Parse(idFacturaGS), 3, mensaje, idDocumentoRecibido_UDO)

                                            ' EL WINDOWS SERVICE DEBE SIEMPRE TOMAR COMO REFERENCIA EL CAMPO SINCRO, Y ENVIAR A EDOC LO QUE TENGA EL CAMPO SINCRO.
                                            ' ES DECIR SI EL CAMPO SINCRO ES IGUAL A 1, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO TRUE
                                            ' SI EL CAMPO ES IGUAL A 0, DEBE ENVIAR A EDOC A SINCRONIZAR CON ESTADO FALSE
                                        End If
                                    End If

                            End Select
                        End If
                End Select

            End If

        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " ERROR - CATH DATA EVENT :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - rSboApp_FormDataEvent: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function MarcarVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String) As Boolean
        Try
            _WS_Recepcion = Functions.VariablesGlobales._WS_Recepcion
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = Functions.VariablesGlobales._WS_RecepcionCambiarEstado
            _WS_RecepcionClave = Functions.VariablesGlobales._WS_RecepcionClave

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
                'oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Documento Marcado como Visto(Integrado) en EDOC ", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(idDocumentoRecibido_UDO, 1)
                Return True
            Else
                oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Error al marcar documento como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS ", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error MarcarVisto : " + ex.Message.ToString, "frmDocumentoNC")
            Return False
        End Try
    End Function

    Public Function MarcarNOVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String) As Boolean
        Try
            _WS_Recepcion = Functions.VariablesGlobales._WS_Recepcion
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = Functions.VariablesGlobales._WS_RecepcionCambiarEstado
            _WS_RecepcionClave = Functions.VariablesGlobales._WS_RecepcionClave

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
            If WS.MarcarNoVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(idDocumentoRecibido_UDO, 0)
                oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Documento Marcado como NO Visto(Integrado) en EDOC Satisfactoriamente! ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return True
            Else
                oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Error al marcar documento NO como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub ConsutarPDFRecibido(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer)
        Try
            Dim rl As String = ""
            rl = Functions.VariablesGlobales._RecepcionLite
            If rl = "Y" Then
                Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO
                Dim rutarl = Functions.VariablesGlobales._Ruta_Compartida
                Dim fechacarpeta As String = Date.Today.ToString("dd-MM-yyyy")
                Dim fechacreacion As String = oFuncionesB1.getRSvalue("SELECT ""CreateDate"" FROM ""@GS_NCR"" WITH(NOLOCK) where ""U_IdGS"" = " + IdDocumento.ToString, "CreateDate", "")
                fechacreacion = CDate(fechacreacion).Date.ToString("dd-MM-yyyy")
                Dim rutaFC As String = ""
                rutaFC = rutarl & "\" & "NOTASDECREDITO" & "\" & fechacarpeta & "\"
                Dim filepath As String = rutaFC
                If fechacarpeta = fechacreacion Then
                    'obtengo la ruta completa añadiendo la clave de acceso
                    filepath += _ClaveAcceso + ".pdf"
                    'If Not File.Exists(filepath) Then
                    '    GuardarArchivoNCRecibido(IdDocumento, TipoDocumento)
                    'End If
                Else
                    rutaFC = rutarl & "\" & "NOTASDECREDITO" & "\" & fechacreacion & "\"
                    filepath = rutaFC
                    filepath += _ClaveAcceso + ".pdf"
                End If
                Dim Proc As New Process()
                Proc.StartInfo.FileName = filepath
                rsboApp.StatusBar.SetText(NombreAddon + " -  PDF abierto exitosamente..!! : " + filepath, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Proc.Start()
                Proc.Dispose()
            Else
                _WS_Recepcion = Functions.VariablesGlobales._WS_Recepcion
                If _WS_Recepcion = "" Then
                    rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                _WS_RecepcionArchivo = Functions.VariablesGlobales._WS_RecepcionConsultaArchivo
                _WS_RecepcionClave = Functions.VariablesGlobales._WS_RecepcionClave

                rsboApp.StatusBar.SetText(NombreAddon + " - Ruta Recepcion: " + _WS_RecepcionArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                rsboApp.StatusBar.SetText(NombreAddon + " - Clave Recepcion: " + _WS_RecepcionClave, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


                Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO
                WS.Url = _WS_RecepcionArchivo
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

                Dim filepath As String = Path.GetTempPath()
                filepath += _ClaveAcceso + ".pdf"


                'Logica Agregada Arturo
                'Para ver si existe el archivo
                If oManejoDocumentos.ExisterchivoLocal(Functions.VariablesGlobales._RutaIntegracionXML & "\" & _ClaveAcceso + ".pdf") Then Exit Sub



                ' SI NO EXISTE EN LA CARPETA TEMPORAL, LO CONSULTO AL WS
                If Not File.Exists(filepath) Then
                    rsboApp.SetStatusBarMessage("Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Dim FS As FileStream = Nothing
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If
                    'oManejoDocumentos.SetProtocolosdeSeguridad()
                    ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                    Dim dbbyte As Byte() = WS.ConsultaArchivoProveedor_PDF(_WS_RecepcionClave, TipoDocumento, _IdGS, mensaje)
                    If dbbyte Is Nothing Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Else
                        FS = New FileStream(filepath, System.IO.FileMode.Create)
                        FS.Write(dbbyte, 0, dbbyte.Length)
                        FS.Close()

                    End If

                End If

                Dim Proc As New Process()
                Proc.StartInfo.FileName = filepath
                Proc.Start()
                Proc.Dispose()
            End If

            

        Catch ex As Exception
            rsboApp.SetStatusBarMessage(NombreAddon + " - Ocurrio un error al generar el PDF recibido! " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
        End Try
    End Sub

    Public Sub GuardarArchivoNCRecibido(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer)
        'obtengo la ruta compartida
        Dim rutarl = Functions.VariablesGlobales._Ruta_Compartida
        'ontegno la fecha actual
        Dim fechacarpeta As String = Date.Today.ToString("dd-MM-yyyy")
        Dim rutaFC As String = ""
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO

        Try
            _WS_Recepcion = Functions.VariablesGlobales._WS_Recepcion
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionArchivo = Functions.VariablesGlobales._WS_RecepcionConsultaArchivo
            _WS_RecepcionClave = Functions.VariablesGlobales._WS_RecepcionClave

            rsboApp.StatusBar.SetText(NombreAddon + " - Ruta Recepcion: " + _WS_RecepcionArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            rsboApp.StatusBar.SetText(NombreAddon + " - Clave Recepcion: " + _WS_RecepcionClave, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Dim pdf As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO

            rsboApp.StatusBar.SetText(NombreAddon + " - Seteando Entidad ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            pdf.Url = _WS_RecepcionArchivo

            'verifico si la carpeta NOTADECREDITO existe en la ruta compartida
            If Not Directory.Exists(rutarl & "\" & "NOTASDECREDITO") Then

                Directory.CreateDirectory(rutarl & "\" & "NOTASDECREDITO")
                Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + rutarl & "\" & "NOTASDECREDITO".ToString, "frmDocumento")
            End If
            'verifico si la carpeta con la fecha actual existe dentro de la carpeta NOTADECREDITOA
            If Not Directory.Exists(rutarl & "\" & "NOTASDECREDITO" & "\" & fechacarpeta) Then
                Directory.CreateDirectory(rutarl & "\" & "NOTASDECREDITO" & "\" & fechacarpeta)
                Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + rutarl & "\" & "NOTASDECREDITO" & "\" & fechacarpeta, "frmDocumento")
               
            End If

            rutaFC = rutarl & "\" & "NOTASDECREDITO" & "\" & fechacarpeta & "\"
            WS.Url = _WS_RecepcionArchivo

            Dim filepath As String = rutaFC
            'obtengo la ruta completa añadiendo la clave de acceso
            filepath += _ClaveAcceso + ".pdf"

            'verifico si la clave de acceso existe dentro de la carpeta NOTADECREDITO
            If Not File.Exists(filepath) Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - Guardando pdf, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim FS As FileStream = Nothing
                Dim dbbyte As Byte() = WS.ConsultaArchivoProveedor_PDF(_WS_RecepcionClave, TipoDocumento, _IdGS, mensaje)
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
                rsboApp.StatusBar.SetText("(GS) Error al ingresar configuracion previa: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                rsboApp.StatusBar.SetText("(GS) Grabando Log como: " + Estado.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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

    Private Sub OcultarObjetosPorTipo(Tipo As String)

        Try
            oForm = rsboApp.Forms.Item("frmDocumentoNC")
            oForm.Freeze(True)

            If Tipo = "NC Inventariable" Then
                Dim flDetalle As SAPbouiCOM.Folder
                flDetalle = oForm.Items.Item("flRelacion").Specific
                flDetalle.Item.Visible = False

                oForm.Items.Item("Item_25").Visible = False
                oForm.Items.Item("Item_20").Visible = False
                oForm.Items.Item("Item_24").Visible = False
                oForm.Items.Item("Item_26").Visible = False
                oForm.Items.Item("Item_29").Visible = False
                oForm.Items.Item("Item_31").Visible = False
                oForm.Items.Item("Item_32").Visible = False
                oForm.Items.Item("txtDTot").Visible = False
                oForm.Items.Item("txtDP").Visible = False
                oForm.Items.Item("txtDVP").Visible = False
                oForm.Items.Item("txtDG").Visible = False
                oForm.Items.Item("txtDI").Visible = False
                oForm.Items.Item("txtDT").Visible = False
                oForm.Items.Item("cbEnlazar").Visible = False

                oForm.Items.Item("Item_8").Visible = True
                oForm.Items.Item("lbMapp").Visible = True
                oForm.Items.Item("btnMapear").Visible = True
                Dim oGrid As SAPbouiCOM.Grid
                oGrid = oForm.Items.Item("oGrid").Specific
                oGrid.Columns.Item(2).Visible = True

                '
            ElseIf Tipo = "NC Inventariable Relacionada" Then
                Dim flDetalle As SAPbouiCOM.Folder
                flDetalle = oForm.Items.Item("flRelacion").Specific
                flDetalle.Item.Visible = True
                oForm.Items.Item("Item_25").Visible = True
                oForm.Items.Item("Item_20").Visible = True
                oForm.Items.Item("Item_24").Visible = True
                oForm.Items.Item("Item_26").Visible = True
                oForm.Items.Item("Item_29").Visible = True
                oForm.Items.Item("Item_31").Visible = True
                oForm.Items.Item("Item_32").Visible = True
                oForm.Items.Item("txtDTot").Visible = True
                oForm.Items.Item("txtDP").Visible = True
                oForm.Items.Item("txtDVP").Visible = True
                oForm.Items.Item("txtDG").Visible = True
                oForm.Items.Item("txtDI").Visible = True
                oForm.Items.Item("txtDT").Visible = True
                oForm.Items.Item("cbEnlazar").Visible = True
                oForm.Items.Item("Item_8").Visible = False
                oForm.Items.Item("lbMapp").Visible = False
                oForm.Items.Item("btnMapear").Visible = False
                Dim oGrid As SAPbouiCOM.Grid
                oGrid = oForm.Items.Item("oGrid").Specific
                oGrid.Columns.Item(2).Visible = False
                '
            ElseIf Tipo = "NC de Servicio" Then
                Dim flDetalle As SAPbouiCOM.Folder
                flDetalle = oForm.Items.Item("flRelacion").Specific
                flDetalle.Item.Visible = False
                oForm.Items.Item("Item_25").Visible = False
                oForm.Items.Item("Item_20").Visible = False
                oForm.Items.Item("Item_24").Visible = False
                oForm.Items.Item("Item_26").Visible = False
                oForm.Items.Item("Item_29").Visible = False
                oForm.Items.Item("Item_31").Visible = False
                oForm.Items.Item("Item_32").Visible = False
                oForm.Items.Item("txtDTot").Visible = False
                oForm.Items.Item("txtDP").Visible = False
                oForm.Items.Item("txtDVP").Visible = False
                oForm.Items.Item("txtDG").Visible = False
                oForm.Items.Item("txtDI").Visible = False
                oForm.Items.Item("txtDT").Visible = False
                oForm.Items.Item("cbEnlazar").Visible = False

                oForm.Items.Item("Item_8").Visible = False
                oForm.Items.Item("lbMapp").Visible = False
                oForm.Items.Item("btnMapear").Visible = False
                Dim oGrid As SAPbouiCOM.Grid
                oGrid = oForm.Items.Item("oGrid").Specific
                oGrid.Columns.Item(2).Visible = False
            End If
            oForm.Freeze(False)
        Catch ex As Exception

        Finally
            If oForm IsNot Nothing Then
                'Descongelamos el formulario haya habido o no excepción
                oForm.Freeze(False)
            End If
        End Try
    End Sub

    Private Function CrearNCPreliminarRelacionada(ByRef sDocEntryPreliminar As String, ByVal DocEntryNERecibida_UDO As String) As Boolean

        'Dim S As String = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocr").SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly)
        'rsboApp.MessageBox(S.ToString())

        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Creando Nota de Crédito Preliminar de tipo: Relacionada", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Nota de Crédito por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        Dim Prefijo As String = Functions.VariablesGlobales._Prefijo_NC
        Dim DocEntryBase As Integer
        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

        Try
            ' If baseGRPO.GetByKey(PO_DocEntry) = True Then
            GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes
            GRPO.CardCode = oCardCode
            'GRPO.DocDate = _oDocumento.FechaEmision
            If Functions.VariablesGlobales._vgFechaEmisionNotaCredito = "Y" Then
                GRPO.DocDate = _oDocumento.FechaEmision
                GRPO.DocDueDate = _oDocumento.FechaEmision
                GRPO.TaxDate = _oDocumento.FechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionNotaCreditoP = "Y" Then
                GRPO.DocDate = _oDocumento.FechaEmision
            ElseIf IIf(String.IsNullOrEmpty(Functions.VariablesGlobales._FechaAutEnFechaContabNC), "N", Functions.VariablesGlobales._FechaAutEnFechaContabNC) = "Y" Then
                GRPO.DocDate = _oDocumento.FechaAutorizacion
            Else
                GRPO.DocDate = Today.Date
                'GRPO.DocDueDate = 
                GRPO.TaxDate = _oDocumento.FechaEmision
            End If
            

            'iTotalPO_Line = baseGRPO.Lines.Count
            'iTotalFrgChg_Line = baseGRPO.Expenses.Count

            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SER_EST").Value = _oDocumento.Establecimiento
                GRPO.UserFields.Fields.Item("U_SER_PE").Value = _oDocumento.PuntoEmision

                GRPO.FolioNumber = _oDocumento.Secuencial
                GRPO.FolioPrefixString = Prefijo
                Try
                    GRPO.UserFields.Fields.Item("U_fecha_emi_doc_rel").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                End Try

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                GRPO.UserFields.Fields.Item("U_NO_AUTORI").Value = _oDocumento.AutorizacionSRI
                GRPO.NumAtCard = _oDocumento.Establecimiento + _oDocumento.PuntoEmision + _oDocumento.Secuencial

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                GRPO.NumAtCard = _oDocumento.Establecimiento + "-" + _oDocumento.PuntoEmision + "-" + _oDocumento.Secuencial
                GRPO.UserFields.Fields.Item("U_SYP_SERIESUC").Value = _oDocumento.Establecimiento
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSD").Value = _oDocumento.PuntoEmision
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDSD").Value = _oDocumento.PuntoEmision
                End Try
                'Try
                '    GRPO.UserFields.Fields.Item("U_SYP_MDSD").Value = _oDocumento.PuntoEmision
                'Catch ex As Exception
                '    GRPO.UserFields.Fields.Item("U_BPP_MDSD").Value = _oDocumento.PuntoEmision
                'End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDCD").Value = _oDocumento.Secuencial
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDCD").Value = _oDocumento.Secuencial
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _oDocumento.AutorizacionSRI
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_SYP_NroAuto").Value = _oDocumento.AutorizacionSRI
                End Try

                'CAMPOS USUARIOS PARA TOTALTEK
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDTD").Value = _oDocumento.CodigoDocumento
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_MDTD: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_TIPO_EMIS").Value = "E"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_TIPO_EMIS: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_FORMAP").Value = "20"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FORMAP: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NUM_SRI").Value = "001-001"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FORMAP: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_SERIESUCO").Value = _oDocumento.Establecimiento
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_SERIESUCO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSO").Value = _oDocumento.PuntoEmision
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_MDSO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NDCO").Value = _oDocumento.Secuencial
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NDCO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NROAUTOO").Value = _oDocumento.AutorizacionSRI
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTOO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_FECHAREF").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FECHAREF: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                'FIN CAMPOS USUARIOS PARA TOTALTEK
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                GRPO.UserFields.Fields.Item("U_TM_NAUT").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_TM_DATEA").Value = _oDocumento.FechaAutorizacion
                GRPO.NumAtCard = _oDocumento.Establecimiento.ToString + "-" + _oDocumento.PuntoEmision.ToString + "-" + _oDocumento.Secuencial.ToString

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                GRPO.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_HBT_SER_EST").Value = _oDocumento.Establecimiento
                GRPO.UserFields.Fields.Item("U_HBT_PTO_EST").Value = _oDocumento.PuntoEmision
                GRPO.NumAtCard = _oDocumento.Secuencial.ToString

                GRPO.UserFields.Fields.Item("U_HBT_TIDOMO").Value = "01"
                GRPO.UserFields.Fields.Item("U_HBT_EST_MOD").Value = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
                GRPO.UserFields.Fields.Item("U_HBT_PUFAMO").Value = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
                GRPO.UserFields.Fields.Item("U_HBT_NUFAMO").Value = Right(_oDocumento.NumDocModificado.ToString, 9)
                GRPO.UserFields.Fields.Item("U_HBT_FEDOMO").Value = _oDocumento.FechaEmisionDocModificado

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                GRPO.UserFields.Fields.Item("U_SS_NumAut").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SS_Est").Value = _oDocumento.Establecimiento
                GRPO.UserFields.Fields.Item("U_SS_Pemi").Value = _oDocumento.PuntoEmision

                GRPO.FolioNumber = _oDocumento.Secuencial
                GRPO.FolioPrefixString = Prefijo

                Try
                    GRPO.UserFields.Fields.Item("U_SS_TipCom").Value = "01"
                Catch ex As Exception
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SS_FecEmiDocRel").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                End Try

            End If

            'GRPO.UserFields.Fields.Item("U_SSCLAVE").Value = _oDocumento.ClaveAcceso
            GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
            GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryNERecibida_UDO.ToString()

            Dim dtRELACIONADO As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocr") ' DATA TABLE DOCUMENTOS RELACIONADO

            ' RECORRO EL DT DEL DOCUMENTO RECIBIDO
            For index As Integer = 0 To dtRELACIONADO.Rows.Count - 1
                If ObjTypeRelacionado = 21 Then ' Devolución de Mercadería" Then ' tabla: ORPD objectype: 21
                    GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                ElseIf ObjTypeRelacionado = 18 Then ' Factura de Proveedores" Then ' tabla: OPCH objectype: 18
                    GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                ElseIf ObjTypeRelacionado = 204 Then ' Anticipo de Proveedores" Then ' tabla: ODPO objectype: 204
                    GRPO.Lines.BaseType = Convert.ToInt32(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments)
                Else
                End If
                DocEntryBase = Convert.ToInt32(dtRELACIONADO.GetValue(0, index))
                GRPO.Lines.BaseEntry = Convert.ToInt32(dtRELACIONADO.GetValue(0, index))
                GRPO.Lines.BaseLine = Convert.ToInt32(dtRELACIONADO.GetValue(1, index)) '1
                'GRPO.Lines.Quantity = formatDecimal(dtRELACIONADO.GetValue(4, index).ToString()) 'Cantidad
                'GRPO.Lines.Price = formatDecimal(dtRELACIONADO.GetValue(5, index).ToString()) 'Precio
                ''GRPO.Lines.Descuento = formatDecimal(dtRECIBIDO.GetValue(6, index).ToString()) 'Descuento
                'GRPO.Lines.LineTotal = formatDecimal(dtRELACIONADO.GetValue(8, index).ToString()) 'Line Total
                GRPO.Lines.Add()
                'GRPO.ApplyCurrentVATRatesForDownPaymentsToDraw = SAPbobsCOM.BoYesNoEnum.tYES
                'GRPO.WithholdingTaxData
            Next

            'If ObjTypeRelacionado = 18 Then
            '    Try
            '        Dim sQueryRet As String
            '        Dim oDocEntrys As String = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("docR").Specific.value
            '        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '            sQueryRet = "select ""WTCode"",""Rate"",""TaxbleAmnt"",""TxblAmntSC"",""TxblAmntFC"",""WTAmnt"",""WTAmntSC"",""WTAmntFC"","
            '            sQueryRet += """ApplAmnt"",""ApplAmntSC"",""ApplAmntFC"",""Category"",""Criteria"",""Account"",""Type"",""BaseType"",""LineNum"" "
            '            sQueryRet += "from PCH5 WHERE ""AbsEntry"" IN (" + oDocEntrys + ")"
            '        Else
            '            sQueryRet = "select P.WTCode,O.DocEntry,P.ObjType,P.LineNum,O.DocNum "
            '            sQueryRet += "from OPCH O INNER JOIN PCH5 P ON P.AbsEntry=O.DocEntry "
            '            sQueryRet += "WHERE O.DocEntry IN (" + oDocEntrys + ")"
            '        End If

            '        rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Add("dtDocRelRet")
            '        Utilitario.Util_Log.Escribir_Log("Consulta Detalle Retencion " & sQueryRet, "frmDocumentoNC")
            '        Dim dtRet As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocRelRet")
            '        dtRet.ExecuteQuery(sQueryRet)

            '        For index As Integer = 0 To dtRet.Rows.Count - 1
            '            GRPO.WithholdingTaxData.SetCurrentLine(GRPO.WithholdingTaxData.Count - 1)
            '            GRPO.WithholdingTaxData.WTCode = dtRet.GetValue(0, index)
            '            GRPO.WithholdingTaxData.BaseDocEntry = oDocEntrys
            '            GRPO.WithholdingTaxData.BaseDocType = 18
            '            Dim numerodoc As Integer = GRPO.WithholdingTaxData.BaseDocumentReference
            '            'GRPO.WithholdingTaxData.BaseDocLine = CInt(dtRet.GetValue(3, index))
            '            'GRPO.WithholdingTaxData.
            '            GRPO.WithholdingTaxData.Add()
            '            'GRPO.WithholdingTaxData.TaxableAmount = formatDecimal(dtRet.GetValue(2, index).ToString)
            '            'GRPO.WithholdingTaxData.WTAmount = formatDecimal(dtRet.GetValue(5, index).ToString)
            '            'GRPO.Lines.WithholdingTaxLines.WTCode = dtRet.GetValue(0, index)
            '            'GRPO.Lines.WithholdingTaxLines.BaseDocEntry = oDocEntrys
            '            'GRPO.Lines.WithholdingTaxLines.Add()

            '        Next
            '    Catch ex As Exception

            '    End Try
            'End If
            Dim DocBase As SAPbobsCOM.Documents

            If ObjTypeRelacionado = 21 Then ' Devolución de Mercadería" Then ' tabla: ORPD objectype: 21
                DocBase = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
            ElseIf ObjTypeRelacionado = 18 Then ' Factura de Proveedores" Then ' tabla: OPCH objectype: 18
                DocBase = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            ElseIf ObjTypeRelacionado = 204 Then ' Anticipo de Proveedores" Then ' tabla: ODPO objectype: 204
                DocBase = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments)
            End If

            If DocBase.GetByKey(DocEntryBase) Then
                Dim lineas As Integer = DocBase.Expenses.Count
                If lineas > 0 Then
                    For x As Integer = 0 To lineas - 1
                        DocBase.Expenses.SetCurrentLine(x)

                        If DocBase.Expenses.LineTotal > 0 Then

                            GRPO.Expenses.ExpenseCode = DocBase.Expenses.ExpenseCode
                            If ObjTypeRelacionado = 21 Then
                                GRPO.Expenses.BaseDocType = "21"
                            ElseIf ObjTypeRelacionado = 18 Then
                                GRPO.Expenses.BaseDocType = "18"
                            ElseIf ObjTypeRelacionado = 204 Then
                                GRPO.Expenses.BaseDocType = "204"
                            End If

                            GRPO.Expenses.BaseDocLine = DocBase.Expenses.LineNum
                            GRPO.Expenses.BaseDocEntry = DocBase.DocEntry
                            GRPO.Expenses.TaxCode = DocBase.Expenses.TaxCode
                            GRPO.Expenses.TaxSum = DocBase.Expenses.TaxSum 'vatsum
                            GRPO.Expenses.VatGroup = DocBase.Expenses.VatGroup 'VatGroup
                            GRPO.Expenses.LineTotal = DocBase.Expenses.LineTotal 'VatPrcnt
                            GRPO.Expenses.Remarks = DocBase.Expenses.Remarks 'Comments
                            GRPO.Expenses.Add()

                        End If

                    Next

                End If


            End If

            GRPO.Comments += "Creado por el addon de Recepcion (SolSap)"

            'Dim iTotalPO_Line As Integer
            'Dim iTotalFrgChg_Line As Integer
            'If baseGRPO.GetByKey(PO_DocEntry) = True Then
            '    iTotalFrgChg_Line = baseGRPO.Expenses.Count
            '    ' Freight Charges
            '    If iTotalFrgChg_Line > 0 Then
            '        Dim fcnt As Integer
            '        For fcnt = 0 To iTotalFrgChg_Line - 1
            '            GRPO.Expenses.SetCurrentLine(fcnt)
            '            GRPO.Expenses.ExpenseCode = baseGRPO.Expenses.ExpenseCode
            '            GRPO.Expenses.BaseDocType = "22"
            '            GRPO.Expenses.BaseDocLine = baseGRPO.Expenses.LineNum
            '            GRPO.Expenses.BaseDocEntry = baseGRPO.DocEntry
            '            GRPO.Expenses.Add()
            '        Next
            '    End If
            'End If

            'Add the Invoice
            RetVal = GRPO.Add

            'Check the result
            If RetVal <> 0 Then
                Elimina_DocumentoRecibido_NC(DocEntryNERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Nota de Crédito Preliminar de tipo: Relacionada:" + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "Factura Preliminar de tipo: Relacionada, Creada Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_NC(DocEntryNERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If

            'End If

        Catch ex As Exception
            Elimina_DocumentoRecibido_NC(DocEntryNERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Factura Preliminar de tipo: Relacionada:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            GRPO = Nothing
            GC.Collect()
        End Try


    End Function
    Private Function CrearNCPremilinarServicio(ByVal sCardCode As String, ByVal oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito, ByRef sDocEntryPreliminar As String, ByVal DocEntryNCRecibida_UDO As String) As Boolean

        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Try
            oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Creando Nota de Crédito de tipo: Servicio", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Nota de Crédito por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes
            GRPO.CardCode = sCardCode
            'GRPO.DocDate = oNC.FechaEmision
            '
            If Functions.VariablesGlobales._vgFechaEmisionNotaCredito = "Y" Then
                GRPO.DocDate = oNC.FechaEmision
                GRPO.DocDueDate = oNC.FechaEmision
                GRPO.TaxDate = oNC.FechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionNotaCreditoP = "Y" Then
                GRPO.DocDate = oNC.FechaEmision
            ElseIf IIf(String.IsNullOrEmpty(Functions.VariablesGlobales._FechaAutEnFechaContabNC), "N", Functions.VariablesGlobales._FechaAutEnFechaContabNC) = "Y" Then
                GRPO.DocDate = oNC.FechaAutorizacion
            Else
                GRPO.DocDate = Today.Date
                GRPO.TaxDate = oNC.FechaEmision
            End If
            

            ' DATOS DE AUTORIZACION
            Dim Prefijo As String = Functions.VariablesGlobales._Prefijo_NC
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = oNC.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SER_EST").Value = oNC.Establecimiento
                GRPO.UserFields.Fields.Item("U_SER_PE").Value = oNC.PuntoEmision

                GRPO.FolioNumber = oNC.Secuencial
                GRPO.FolioPrefixString = Prefijo
                Try
                    GRPO.UserFields.Fields.Item("U_fecha_emi_doc_rel").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                End Try


            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                GRPO.UserFields.Fields.Item("U_NO_AUTORI").Value = oNC.AutorizacionSRI
                GRPO.NumAtCard = oNC.Establecimiento + oNC.PuntoEmision + oNC.Secuencial

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                GRPO.NumAtCard = _oDocumento.Establecimiento + "-" + _oDocumento.PuntoEmision + "-" + _oDocumento.Secuencial
                GRPO.UserFields.Fields.Item("U_SYP_SERIESUC").Value = _oDocumento.Establecimiento
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSD").Value = _oDocumento.PuntoEmision
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDSD").Value = _oDocumento.PuntoEmision
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDCD").Value = _oDocumento.Secuencial
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDCD").Value = _oDocumento.Secuencial
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _oDocumento.AutorizacionSRI
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_SYP_NroAuto").Value = _oDocumento.AutorizacionSRI
                End Try

                'CAMPOS USUARIOS PARA TOTALTEK
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDTD").Value = _oDocumento.CodigoDocumento
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_MDTD: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_TIPO_EMIS").Value = "E"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_TIPO_EMIS: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_FORMAP").Value = "20"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FORMAP: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NUM_SRI").Value = "001-001"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FORMAP: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_SERIESUCO").Value = _oDocumento.Establecimiento
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_SERIESUCO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSO").Value = _oDocumento.PuntoEmision
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_MDSO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NDCO").Value = _oDocumento.Secuencial
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NDCO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NROAUTOO").Value = _oDocumento.AutorizacionSRI
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTOO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_FECHAREF").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FECHAREF: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                'FIN CAMPOS USUARIOS PARA TOTALTEK
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                GRPO.UserFields.Fields.Item("U_TM_NAUT").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_TM_DATEA").Value = _oDocumento.FechaAutorizacion
                GRPO.NumAtCard = _oDocumento.Establecimiento.ToString + "-" + _oDocumento.PuntoEmision.ToString + "-" + _oDocumento.Secuencial.ToString

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                GRPO.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_HBT_SER_EST").Value = _oDocumento.Establecimiento
                GRPO.UserFields.Fields.Item("U_HBT_PTO_EST").Value = _oDocumento.PuntoEmision
                GRPO.NumAtCard = _oDocumento.Secuencial.ToString

                GRPO.UserFields.Fields.Item("U_HBT_TIDOMO").Value = "01"
                GRPO.UserFields.Fields.Item("U_HBT_EST_MOD").Value = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
                GRPO.UserFields.Fields.Item("U_HBT_PUFAMO").Value = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
                GRPO.UserFields.Fields.Item("U_HBT_NUFAMO").Value = Right(_oDocumento.NumDocModificado.ToString, 9)
                GRPO.UserFields.Fields.Item("U_HBT_FEDOMO").Value = _oDocumento.FechaEmisionDocModificado

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                GRPO.UserFields.Fields.Item("U_SS_NumAut").Value = oNC.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SS_Est").Value = oNC.Establecimiento
                GRPO.UserFields.Fields.Item("U_SS_Pemi").Value = oNC.PuntoEmision

                GRPO.FolioNumber = oNC.Secuencial
                GRPO.FolioPrefixString = Prefijo
                Try
                    GRPO.UserFields.Fields.Item("U_SS_TipCom").Value = "01"
                Catch ex As Exception
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SS_FecEmiDocRel").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                End Try
            End If
            'Dim pe As String = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
            'Dim es As String = _oDocumento.NumDocModificado.ToString.Substring(4, 3)
            'Dim s As String = Right(_oDocumento.NumDocModificado.ToString, 9)
            'U_EREC_CREADO 
            'GRPO.UserFields.Fields.Item("U_SSCLAVE").Value = oNC.ClaveAcceso
            GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
            GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryNCRecibida_UDO.ToString()

            'Dim serviceInvoice As Documents = TryCast(B1Connections.diCompany.GetBusinessObject(BoObjectTypes.oInvoices), Documents)
            'serviceInvoice.CardCode = "C20000"
            GRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            GRPO.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO

            Dim FormatCode As String = ""
            Dim sQueryAcctCode As String = ""

            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" Where ""CardCode"" =  '" + sCardCode + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardCode =  '" + sCardCode + "'"
            End If
            FormatCodeProveedor = oFuncionesB1.getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            If FormatCodeProveedor = "" Then
                FormatCode = Functions.VariablesGlobales._Cuenta_NC
            Else
                FormatCode = FormatCodeProveedor
            End If

            ' FormatCode = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "NC", "Cuenta")
            If FormatCode = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "ERROR - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode + "'"
            End If

            Dim Cuenta As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")

            Dim line As Integer = 0
            Dim sQueryCodImp As String
            Dim CodImp As String
            Dim sQueryCodImpV As String
            Dim CodImpV As String
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito In oNC.ENTDetalleNotaCredito
                sQueryCodImp = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oNC.ENTNotaCreditoImpuesto(0).CodigoPorcentaje.ToString + "' "
                CodImp = oFuncionesB1.getRSvalue(sQueryCodImp, "U_SSID", "")
                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImp + "Resultado :" + CodImp.ToString(), "frmDocumentoNC")

                sQueryCodImpV = " SELECT ""U_SSCOD"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oNC.ENTNotaCreditoImpuesto(0).CodigoPorcentaje.ToString + "' "
                CodImpV = oFuncionesB1.getRSvalue(sQueryCodImpV, "U_SSCOD", "")
                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImpV + "Resultado :" + CodImpV.ToString(), "frmDocumentoNC")

                GRPO.Lines.AccountCode = Cuenta
                GRPO.Lines.LineTotal = formatDecimal(oDetalle.PrecioUnitario)
                If CodImpV = oNC.ENTNotaCreditoImpuesto(0).CodigoPorcentaje.ToString Then
                    GRPO.Lines.TaxCode = CodImp.ToString
                End If
                GRPO.Lines.ItemDescription = LTrim(RTrim(Left(oDetalle.Descripcion.ToString, 100)))
                GRPO.Lines.Quantity = 1
                GRPO.Lines.Add()
                line += 1
            Next

            GRPO.Comments += "Creado por el addon de Recepcion (SolSap)"

            RetVal = GRPO.Add()
            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Elimina_DocumentoRecibido_NC(DocEntryNCRecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Nota de Crédito Preliminar de tipo: Servicio:" + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Nota de Crédito Preliminar de tipo: Servicio, Creada Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_NC(DocEntryNCRecibida_UDO, sDocEntryPreliminar)
                Return True
            End If

        Catch ex As Exception
            Elimina_DocumentoRecibido_NC(DocEntryNCRecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            GRPO = Nothing
            GC.Collect()
        End Try

    End Function
    Private Function CrearNCPreliminarMapeada(ByVal sCardCode As String, ByVal oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito, ByRef sDocEntryPreliminar As String, ByVal DocEntryNCRecibida_UDO As String) As Boolean

        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Creando Nota de Crédito Preliminar de tipo: Mapeo de Items", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Nota de Crédito por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Disable Warning BC42024 ' Variable local sin usar: 'iTotalPO_Line'.
        Dim iTotalPO_Line As Integer
#Enable Warning BC42024 ' Variable local sin usar: 'iTotalPO_Line'.
#Disable Warning BC42024 ' Variable local sin usar: 'iTotalFrgChg_Line'.
        Dim iTotalFrgChg_Line As Integer
#Enable Warning BC42024 ' Variable local sin usar: 'iTotalFrgChg_Line'.

        'Dim baseGRPO As SAPbobsCOM.Documents
        'baseGRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)


        Dim Prefijo As String = Functions.VariablesGlobales._Prefijo_NC

        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

        Try

            ' If baseGRPO.GetByKey(PO_DocEntry) = True Then
            GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes
            GRPO.CardCode = sCardCode
            'GRPO.DocDate = oNC.FechaEmision
            'GRPO.DocDueDate = Today.Date
            If Functions.VariablesGlobales._vgFechaEmisionNotaCredito = "Y" Then
                GRPO.DocDate = oNC.FechaEmision
                GRPO.DocDueDate = oNC.FechaEmision
                GRPO.TaxDate = oNC.FechaEmision
            ElseIf Functions.VariablesGlobales._vgFechaEmisionNotaCreditoP = "Y" Then
                GRPO.DocDate = oNC.FechaEmision
            ElseIf IIf(String.IsNullOrEmpty(Functions.VariablesGlobales._FechaAutEnFechaContabNC), "N", Functions.VariablesGlobales._FechaAutEnFechaContabNC) = "Y" Then
                GRPO.DocDate = oNC.FechaAutorizacion
            Else
                GRPO.DocDate = Today.Date
                GRPO.TaxDate = oNC.FechaEmision
            End If
            

            GRPO.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
            GRPO.DocTotal = formatDecimal(oNC.ValorModificacion)

            'iTotalPO_Line = baseGRPO.Lines.Count
            'iTotalFrgChg_Line = baseGRPO.Expenses.Count

            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = oNC.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SER_EST").Value = oNC.Establecimiento
                GRPO.UserFields.Fields.Item("U_SER_PE").Value = oNC.PuntoEmision

                GRPO.FolioNumber = oNC.Secuencial
                GRPO.FolioPrefixString = Prefijo

                
                Try
                    GRPO.UserFields.Fields.Item("U_fecha_emi_doc_rel").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                End Try

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                GRPO.UserFields.Fields.Item("U_NO_AUTORI").Value = oNC.AutorizacionSRI
                GRPO.NumAtCard = oNC.Establecimiento + oNC.PuntoEmision + oNC.Secuencial

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                GRPO.NumAtCard = _oDocumento.Establecimiento + "-" + _oDocumento.PuntoEmision + "-" + _oDocumento.Secuencial
                GRPO.UserFields.Fields.Item("U_SYP_SERIESUC").Value = _oDocumento.Establecimiento
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSD").Value = _oDocumento.PuntoEmision
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDSD").Value = _oDocumento.PuntoEmision
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDCD").Value = _oDocumento.Secuencial
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDCD").Value = _oDocumento.Secuencial
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _oDocumento.AutorizacionSRI
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_SYP_NroAuto").Value = _oDocumento.AutorizacionSRI
                End Try

                'CAMPOS USUARIOS PARA TOTALTEK
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDTD").Value = _oDocumento.CodigoDocumento
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_MDTD: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_TIPO_EMIS").Value = "E"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_TIPO_EMIS: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_FORMAP").Value = "20"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FORMAP: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NUM_SRI").Value = "001-001"
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FORMAP: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_SERIESUCO").Value = _oDocumento.Establecimiento
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_SERIESUCO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSO").Value = _oDocumento.PuntoEmision
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_MDSO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NDCO").Value = _oDocumento.Secuencial
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NDCO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_NROAUTOO").Value = _oDocumento.AutorizacionSRI
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_NROAUTOO: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_FECHAREF").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("U_SYP_FECHAREF: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                'FIN CAMPOS USUARIOS PARA TOTALTEK
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                GRPO.UserFields.Fields.Item("U_TM_NAUT").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_TM_DATEA").Value = _oDocumento.FechaAutorizacion
                GRPO.NumAtCard = _oDocumento.Establecimiento.ToString + "-" + _oDocumento.PuntoEmision.ToString + "-" + _oDocumento.Secuencial.ToString

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                GRPO.UserFields.Fields.Item("U_HBT_AUT_FAC").Value = _oDocumento.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_HBT_SER_EST").Value = _oDocumento.Establecimiento
                GRPO.UserFields.Fields.Item("U_HBT_PTO_EST").Value = _oDocumento.PuntoEmision
                GRPO.NumAtCard = _oDocumento.Secuencial.ToString

                GRPO.UserFields.Fields.Item("U_HBT_TIDOMO").Value = "01"
                GRPO.UserFields.Fields.Item("U_HBT_EST_MOD").Value = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
                GRPO.UserFields.Fields.Item("U_HBT_PUFAMO").Value = _oDocumento.NumDocModificado.ToString.Substring(0, 3)
                GRPO.UserFields.Fields.Item("U_HBT_NUFAMO").Value = Right(_oDocumento.NumDocModificado.ToString, 9)
                GRPO.UserFields.Fields.Item("U_HBT_FEDOMO").Value = _oDocumento.FechaEmisionDocModificado

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                GRPO.UserFields.Fields.Item("U_SS_NumAut").Value = oNC.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SS_Est").Value = oNC.Establecimiento
                GRPO.UserFields.Fields.Item("U_SS_Pemi").Value = oNC.PuntoEmision

                GRPO.FolioNumber = oNC.Secuencial
                GRPO.FolioPrefixString = Prefijo

                Try
                    GRPO.UserFields.Fields.Item("U_SS_TipCom").Value = "01"
                Catch ex As Exception
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SS_FecEmiDocRel").Value = _oDocumento.FechaEmision
                Catch ex As Exception
                End Try

            End If

            'U_EREC_CREADO 
            'GRPO.UserFields.Fields.Item("U_SSCLAVE").Value = oNC.ClaveAcceso
            GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
            GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryNCRecibida_UDO.ToString()

            Dim itemCode As String = ""
            Dim line As Integer = 0

            Dim sQueryCodImp As String
            Dim CodImp As String
            Dim sQueryCodImpV As String
            Dim CodImpV As String

            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito In oNC.ENTDetalleNotaCredito
                'SELECT DfltWH FROM OITM WITH(NOLOCK) WHERE ItemCode
                sQueryCodImp = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oNC.ENTNotaCreditoImpuesto(0).CodigoPorcentaje.ToString + "' "
                CodImp = oFuncionesB1.getRSvalue(sQueryCodImp, "U_SSID", "")
                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImp + "Resultado :" + CodImp.ToString(), "frmDocumentoNC")

                sQueryCodImpV = " SELECT ""U_SSCOD"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oNC.ENTNotaCreditoImpuesto(0).CodigoPorcentaje.ToString + "' "
                CodImpV = oFuncionesB1.getRSvalue(sQueryCodImpV, "U_SSCOD", "")
                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImpV + "Resultado :" + CodImpV.ToString(), "frmDocumentoNC")


                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    itemCode = oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + sCardCode + "' AND ""Substitute"" = '" + Left(oDetalle.CodigoPrincipal, 50) + "'", "ItemCode", "")
                Else
                    itemCode = oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "' AND Substitute = '" + Left(oDetalle.CodigoPrincipal, 50) + "'", "ITemCode", "")
                End If

                GRPO.Lines.ItemCode = itemCode
                GRPO.Lines.Quantity = formatDecimal(oDetalle.Cantidad)
                'GRPO.Lines.ItemDescription = "xxxx"

                If CodImpV = oNC.ENTNotaCreditoImpuesto(0).CodigoPorcentaje.ToString Then
                    GRPO.Lines.TaxCode = CodImp.ToString
                End If

                GRPO.Lines.DiscountPercent = 0
                'GRPO.Lines.WarehouseCode = oFuncionesB1.getRSvalue("SELECT DfltWH FROM OITM WITH(NOLOCK) WHERE ItemCode = '" + itemCode + "'", "DfltWH", "")

                GRPO.Lines.UnitPrice = formatDecimal(oDetalle.PrecioUnitario)

                GRPO.Lines.Add()
                line += 1
            Next

            GRPO.Comments += "Creado por el addon de Recepcion (SolSap)"

            'Add the Invoice
            RetVal = GRPO.Add

            'Check the result
            If RetVal <> 0 Then
                Elimina_DocumentoRecibido_NC(DocEntryNCRecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Nota de Crédito Preliminar de tipo: Mapeo de Items:" + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Nota de Crédito Preliminar de tipo: Mapeo de Items, Creada Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_NC(DocEntryNCRecibida_UDO, sDocEntryPreliminar)
                Return True
            End If

            'End If

        Catch ex As Exception
            Elimina_DocumentoRecibido_NC(DocEntryNCRecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("NCR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Nota de Crédito Preliminar de tipo: Mapeo de Items:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            GRPO = Nothing
            GC.Collect()
        End Try


    End Function

    Public Sub CargaDocumentoRelacionados()
        Try
            Dim oObjType As String = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("objR").Specific.value
            Dim oDocEntrys As String = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("docR").Specific.value

            Dim sQuery As String = ""
            If oObjType = 21 Then ' "Devolución de Mercadería" Then ' tabla: ORPD objectype: 21
                ' sQuery = "EXEC RP_ConsultaDevolucionDeMercancia'" + oDocEntrys + "'"
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT ""DocEntry"",""LineNum"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""DiscPrcnt"",""TaxCode"",""LineTotal"",""ObjType"" FROM ""RPD1"" WHERE ""LineStatus""='O' AND ""DocEntry"" IN (" + oDocEntrys + ")"
                Else
                    sQuery = "SELECT DocEntry,LineNum,ItemCode,Dscription,Quantity,Price,DiscPrcnt,TaxCode,LineTotal,ObjType FROM RPD1 WHERE LineStatus='O' AND DocEntry IN (" + oDocEntrys + ")"
                End If

            ElseIf oObjType = 18 Then ' Factura de Proveedores" Then ' tabla: OPCH objectype: 18
                ' sQuery = "EXEC RP_ConsultaFacturaDeProveedor '" + oDocEntrys + "'"
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT ""DocEntry"",""LineNum"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""DiscPrcnt"",""TaxCode"",""LineTotal"",""ObjType"" FROM ""PCH1"" WHERE ""LineStatus""='O' AND ""DocEntry"" IN (" + oDocEntrys + ")"
                Else
                    sQuery = "SELECT DocEntry,LineNum,ItemCode,Dscription,Quantity,Price,DiscPrcnt,TaxCode,LineTotal,ObjType FROM PCH1 WHERE LineStatus='O' AND DocEntry IN (" + oDocEntrys + ")"
                End If

            ElseIf oObjType = 204 Then '"Anticipo de Proveedores" Then ' tabla: ODPO objectype: 204
                ' sQuery = "EXEC RP_ConsultaAnticipoDeProveedor '" + oDocEntrys + "'"
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT ""DocEntry"",""LineNum"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""DiscPrcnt"",""TaxCode"",""LineTotal"",""ObjType"" FROM ""DPO1"" WHERE ""LineStatus""='O' AND ""DocEntry"" IN (" + oDocEntrys + ")"
                Else
                    sQuery = "SELECT DocEntry,LineNum,ItemCode,Dscription,Quantity,Price,DiscPrcnt,TaxCode,LineTotal,ObjType FROM DPO1 WHERE LineStatus='O' AND DocEntry IN (" + oDocEntrys + ")"
                End If

            End If
            Utilitario.Util_Log.Escribir_Log("Consulta Cabecera " & sQuery, "frmDocumentoNC")
            Try
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Add("dtDocr") ' DATA TABLE, PARA DETALLE DE DOCUMENTOS RELACIONADOS
            Catch ex As Exception
            End Try
            Try
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Add("dtDocRel") ' DATA TABLE, PARA RESUMEN DE LOS DOCUMENTOS RELACIONADOS
            Catch ex As Exception
            End Try

            Dim oGri As SAPbouiCOM.Grid = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("oGridD").Specific
            rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocr").ExecuteQuery(sQuery)
            oGri.DataTable = rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocr")
            oGri.Item.Enabled = False
            If oObjType = 21 Then
                'oGri.Columns.Item(0).Visible = False
                oGri.Columns.Item(0).TitleObject.Caption = "Devoluciones"
                Dim oEditTextColump As SAPbouiCOM.EditTextColumn
                oEditTextColump = oGri.Columns.Item(0)
                oEditTextColump.LinkedObjectType = 21
            ElseIf oObjType = 18 Then
                'oGri.Columns.Item(0).Visible = False
                oGri.Columns.Item(0).TitleObject.Caption = "Facturas"
                Dim oEditTextColump As SAPbouiCOM.EditTextColumn
                oEditTextColump = oGri.Columns.Item(0)
                oEditTextColump.LinkedObjectType = 18
            ElseIf oObjType = 204 Then
                'oGri.Columns.Item(0).Visible = False
                oGri.Columns.Item(0).TitleObject.Caption = "Anticipos"
                Dim oEditTextColump As SAPbouiCOM.EditTextColumn
                oEditTextColump = oGri.Columns.Item(0)
                oEditTextColump.LinkedObjectType = 204
            End If

            'oGri.Columns.Item(1).Visible = False
            oGri.Columns.Item(1).TitleObject.Caption = "Linea"

            oGri.Columns.Item(2).Description = "Número de Artículo"
            oGri.Columns.Item(2).TitleObject.Caption = "Número de Artículo"
            oGri.Columns.Item(2).Editable = False
            Dim oEditTextColum As SAPbouiCOM.EditTextColumn
            oEditTextColum = oGri.Columns.Item(2)
            oEditTextColum.LinkedObjectType = 4

            rsboApp.Forms.Item("frmDocumentoNC").Freeze(True)

            Try
                rsboApp.Forms.Item("frmConsultaOrdenes").Close() ' CIERRO PANTALLA DE CONSULTA
            Catch ex As Exception
            End Try


            ' SETEO LOS TOTALES DE LA PESTAÑA DOCUMENTOS RELACIONADOS
            Dim sQueryResumen As String = ""
            If oObjType = 21 Then ' "Devolución de Mercadería" Then ' tabla: ORPD objectype: 21
                'sQueryResumen = "EXEC RP_ConsultaPedidosDeCompraResumen '" + oDocEntrys + "'"
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQueryResumen = "SELECT ((SUM(""DocTotal"") - SUM(""VatSum"") - SUM(""TotalExpns""))+SUM(""DiscSum"")) AS BASE, "
                    sQueryResumen += " SUM(""DiscPrcnt"") AS DiscPrcnt, SUM(""DiscSum"")AS DiscSum, SUM(""TotalExpns"") AS TotalExpns,SUM(""VatSum"")AS Vatsum,"
                    sQueryResumen += "SUM(""DocTotal"") AS DocTotal FROM ""ORPD"" WHERE ""DocEntry"" IN (" + oDocEntrys + ")"
                Else
                    sQueryResumen = "SELECT ((SUM(DocTotal) - SUM(Vatsum) - SUM(TotalExpns))+SUM(DiscSum)) AS BASE, "
                    sQueryResumen += " SUM(DiscPrcnt) AS DiscPrcnt, SUM(DiscSum)AS DiscSum, SUM(TotalExpns) AS TotalExpns,SUM(Vatsum)AS Vatsum,"
                    sQueryResumen += "SUM(DocTotal) AS DocTotal FROM ORPD WHERE DocEntry IN (" + oDocEntrys + ")"
                End If
                
            ElseIf oObjType = 18 Then ' Factura de Proveedores" Then ' tabla: OPCH objectype: 18
                'sQueryResumen = "EXEC RP_ConsultaEntradasDeMercanciaResumen '" + oDocEntrys + "'"
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQueryResumen = "SELECT ((SUM(""DocTotal"") - SUM(""VatSum"") - SUM(""TotalExpns""))+SUM(""DiscSum"")) AS BASE, "
                    sQueryResumen += " SUM(""DiscPrcnt"") AS DiscPrcnt, SUM(""DiscSum"")AS DiscSum, SUM(""TotalExpns"") AS TotalExpns,SUM(""VatSum"")AS Vatsum,"
                    sQueryResumen += "SUM(""DocTotal"") AS DocTotal FROM ""OPCH"" WHERE ""DocEntry"" IN (" + oDocEntrys + ")"
                Else
                    sQueryResumen = "SELECT ((SUM(DocTotal) - SUM(Vatsum) - SUM(TotalExpns))+SUM(DiscSum)) AS BASE, "
                    sQueryResumen += " SUM(DiscPrcnt) AS DiscPrcnt, SUM(DiscSum)AS DiscSum, SUM(TotalExpns) AS TotalExpns,SUM(Vatsum)AS Vatsum,"
                    sQueryResumen += "SUM(DocTotal) AS DocTotal FROM OPCH WHERE DocEntry IN (" + oDocEntrys + ")"
                End If

            ElseIf oObjType = 204 Then '"Anticipo de Proveedores" Then ' tabla: ODPO objectype: 204
                'sQueryResumen = "EXEC RP_ConsultaEntradasDeMercanciaResumen '" + oDocEntrys + "'"
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQueryResumen = "SELECT ((SUM(""DocTotal"") - SUM(""VatSum"") - SUM(""TotalExpns""))+SUM(""DiscSum"")) AS BASE, "
                    sQueryResumen += " SUM(""DiscPrcnt"") AS DiscPrcnt, SUM(""DiscSum"")AS DiscSum, SUM(""TotalExpns"") AS TotalExpns,SUM(""VatSum"")AS Vatsum,"
                    sQueryResumen += "SUM(""DocTotal"") AS DocTotal FROM ""ODPO"" WHERE ""DocEntry"" IN (" + oDocEntrys + ")"
                Else
                    sQueryResumen = "SELECT ((SUM(DocTotal) - SUM(Vatsum) - SUM(TotalExpns))+SUM(DiscSum)) AS BASE, "
                    sQueryResumen += " SUM(DiscPrcnt) AS DiscPrcnt, SUM(DiscSum)AS DiscSum, SUM(TotalExpns) AS TotalExpns,SUM(Vatsum)AS Vatsum,"
                    sQueryResumen += "SUM(DocTotal) AS DocTotal FROM ODPO WHERE DocEntry IN (" + oDocEntrys + ")"
                End If
                
            End If
            Utilitario.Util_Log.Escribir_Log("Consulta Detalle " & sQueryResumen, "frmDocumentoNC")
            Dim dt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocRel")
            dt.ExecuteQuery(sQueryResumen)
            For index As Integer = 0 To dt.Rows.Count - 1
                Dim BaseImponible As Decimal = dt.GetValue(0, index) ' BASE IMPONIBLE
                Dim PrcDescuento As Decimal = dt.GetValue(1, index) ' PORCENTAJE DE DESCUENTO
                Dim Descuento As Decimal = dt.GetValue(2, index) ' DESCUENTO
                Dim TotalGastos As Decimal = dt.GetValue(3, index) ' TOTAL GASTOS ADICIONALES
                Dim Impuesto As Decimal = dt.GetValue(4, index) ' IMPUESTO
                Dim Total As Decimal = dt.GetValue(5, index) ' TOTAL

                Dim txtDTot As SAPbouiCOM.EditText
                txtDTot = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDTot").Specific
                txtDTot.Item.RightJustified = True
                txtDTot.Value = formatDecimal(Math.Round(BaseImponible, 2).ToString())
                Dim txtDP As SAPbouiCOM.EditText
                txtDP = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDP").Specific
                txtDP.Item.RightJustified = True
                txtDP.Value = Math.Round(PrcDescuento, 0).ToString()
                Dim txtDVP As SAPbouiCOM.EditText
                txtDVP = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDVP").Specific
                txtDVP.Item.RightJustified = True
                txtDVP.Value = formatDecimal(Math.Round(Descuento, 2).ToString())
                Dim txtDG As SAPbouiCOM.EditText
                txtDG = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDG").Specific
                txtDG.Item.RightJustified = True
                txtDG.Value = formatDecimal(Math.Round(TotalGastos, 2).ToString())
                Dim txtDI As SAPbouiCOM.EditText
                txtDI = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDI").Specific
                txtDI.Item.RightJustified = True
                txtDI.Value = formatDecimal(Math.Round(Impuesto, 2).ToString())
                Dim txtDT As SAPbouiCOM.EditText
                txtDT = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("txtDT").Specific
                txtDT.Item.RightJustified = True
                txtDT.Value = formatDecimal(Math.Round(Total, 2).ToString())
            Next

            rsboApp.Forms.Item("frmDocumentoNC").Freeze(False)
        Catch ex As Exception
            rsboApp.MessageBox(ex.Message().ToString())
        End Try
    End Sub

    Public Function Guarda_DocumentoRecibido_NC(ByRef DocEntryNCRecibida_UDO As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Creando registro de Nota de Créddito Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oForm = rsboApp.Forms.Item("frmDocumentoNC")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralData.SetProperty("Code", conta)
            oGeneralData.SetProperty("U_RUC", oForm.Items.Item("txtRUC").Specific.Value.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oForm.Items.Item("txtNombre").Specific.Value.ToString(), 99))
            oGeneralData.SetProperty("U_CardCode", oForm.Items.Item("txtCodigo").Specific.Value.ToString())
            oGeneralData.SetProperty("U_Mapeado", oForm.Items.Item("lbMapp").Specific.Value.ToString())
            oGeneralData.SetProperty("U_ClaAcc", oForm.Items.Item("txtClaAcc").Specific.Value.ToString())
            oGeneralData.SetProperty("U_NumAut", oForm.Items.Item("txtNumAut").Specific.Value.ToString())
            oGeneralData.SetProperty("U_FecAut", oForm.Items.Item("txtFecAut").Specific.Value.ToString())
            oGeneralData.SetProperty("U_NumDoc", oForm.Items.Item("txtNumDoc").Specific.Value.ToString())
            oGeneralData.SetProperty("U_FPrelim", oForm.Items.Item("txtFPre").Specific.Value.ToString())
            oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSub").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_SubTot5", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSub5").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSub0").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSubN").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSubE").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtSubS").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDes").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_ICE", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtICE").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_IVA", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtIva").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_IVA5", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtIva5").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtTotal").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_rTades", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDTot").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_rPDesc", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDP").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_rDesc", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDVP").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_rGast", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDG").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_rImp", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDI").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_rTotal", Convert.ToDouble(formatDecimal(oForm.Items.Item("txtDT").Specific.Value.ToString())))
            oGeneralData.SetProperty("U_IdGS", _oDocumento.IdNotaCredito.ToString())
            oGeneralData.SetProperty("U_Sincro", 0)

            If _EsDocumentoCargadoPorXML Then
                oGeneralData.SetProperty("U_Estado", "docPrelXML")
            Else
                oGeneralData.SetProperty("U_Estado", "docPreliminar")
            End If



            Dim cbxTipo As SAPbouiCOM.ComboBox
            cbxTipo = oForm.Items.Item("cbxTipo").Specific
            oGeneralData.SetProperty("U_Tipo", cbxTipo.Value.ToString())

            oChildren = oGeneralData.Child("GS0_NCR")
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1
                oChild = oChildren.Add
                oChild.SetProperty("U_CodPrin", Left(odt.GetValue(0, i).ToString(), 99))
                oChild.SetProperty("U_CodAuxi", odt.GetValue(1, i).ToString())
                oChild.SetProperty("U_CodSAP", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_Descripc", odt.GetValue(3, i).ToString())
                oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(odt.GetValue(4, i).ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(odt.GetValue(5, i).ToString())))
                oChild.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(odt.GetValue(6, i).ToString())))
                oChild.SetProperty("U_Total", Convert.ToDouble(formatDecimal(odt.GetValue(7, i).ToString())))
            Next

            oChildren = oGeneralData.Child("GS1_NCR")
            odt = oForm.DataSources.DataTables.Item("dtDocr")
            For i = 0 To oForm.DataSources.DataTables.Item("dtDocr").Rows.Count - 1
                oChild = oChildren.Add
                oChild.SetProperty("U_DocEntr", Integer.Parse(odt.GetValue(0, i).ToString()))
                oChild.SetProperty("U_LineNu", Integer.Parse(odt.GetValue(1, i).ToString()))
                oChild.SetProperty("U_ItemCode", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_Descripc", odt.GetValue(3, i).ToString())
                oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(odt.GetValue(4, i).ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(odt.GetValue(5, i).ToString())))
                oChild.SetProperty("U_DiscPr", Convert.ToDouble(formatDecimal(odt.GetValue(6, i).ToString())))
                oChild.SetProperty("U_TaxCode", odt.GetValue(7, i).ToString())
                oChild.SetProperty("U_lTotal", Convert.ToDouble(formatDecimal(odt.GetValue(8, i).ToString())))
                oChild.SetProperty("U_ObjType", odt.GetValue(9, i).ToString())
            Next

            oChildren = oGeneralData.Child("GS2_NCR")
            odt = oForm.DataSources.DataTables.Item("dtDocsDA")
            For i = 0 To oForm.DataSources.DataTables.Item("dtDocsDA").Rows.Count - 1
                oChild = oChildren.Add
                oChild.SetProperty("U_Nombre", odt.GetValue(0, i).ToString())
                oChild.SetProperty("U_Valor", odt.GetValue(1, i).ToString())

            Next

            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryNCRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Se creo registro de Nota de Crédito Recibida UDO satisfactoriamente, # : " + DocEntryNCRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Ocurrior un error al crear registro de Nota de Crédito Recibida UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Nota de Crédito Recibida en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Public Sub Actualiza_DocumentoRecibido_NC(DocEntryNCRecibida_UDO As String, DocEntryPreliminar As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryNCRecibida_UDO)

            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

            oGeneralService.Update(oGeneralData)

            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Error: Actualizando Numero de Documento Preliminar en Documento Recibido UDO: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub
    Public Sub Elimina_DocumentoRecibido_NC(DocEntryNCRecibida_UDO As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Eliminando Documento Recibido UDO Retención # " + DocEntryNCRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryNCRecibida_UDO)

            'oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            'oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

            oGeneralService.Delete(oGeneralParams)


            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO Retención..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub

    ''' <summary>
    ''' Se usa para enviar cambiar el estado a sincronizado en SAP BO
    ''' </summary>
    ''' <param name="DocEntryFacturaRecibida_UDO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ActualizadoEstadoSincronizado_DocumentoRecibido_NC(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Actualizando a Sincronizado = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")

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
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
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
    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")

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
            Utilitario.Util_Log.Escribir_Log("Error ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC : " + ex.Message.ToString, "frmDocumento")
            'rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
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
    Public Function ActualizadoEstado_DocumentoRecibido_NC(ByRef DocEntryFacturaRecibida_UDO As String, Estado As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")

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
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
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
    Private Function ReturnDocEntryDocBase(XMLITEMSRELACIONADOS As String, itemCode As String) As Integer()
        Dim DOCENTRY() As Integer
        Try
            ' http://www.dotnetcurry.com/linq/564/linq-to-xml-tutorials-examples

            Dim xelement As XElement = xelement.Parse(XMLITEMSRELACIONADOS)

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
            Dim xelement As XElement = xelement.Parse(XML)

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
            Dim xelement As XElement = xelement.Parse(xml)
            Dim result As System.Collections.Generic.IEnumerable(Of System.Xml.Linq.XElement)
            result = xelement.Descendants("Rows").Elements("Row").Elements("Cells").Elements("Cell").Where(Function(n) n.Element("Value").Value = itemCode)
            Return result.Count()
        Catch ex As Exception

        End Try
    End Function
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
    '        GRPO.DocDate = _oDocumento.FechaEmision
    '        GRPO.DocDueDate = Today.Date

    '        'iTotalPO_Line = baseGRPO.Lines.Count
    '        'iTotalFrgChg_Line = baseGRPO.Expenses.Count

    '        ' DATOS DE AUTORIZACION
    '        GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = _oDocumento.AutorizacionSRI
    '        GRPO.UserFields.Fields.Item("U_SER_EST").Value = _oDocumento.Establecimiento
    '        GRPO.UserFields.Fields.Item("U_SER_PE").Value = _oDocumento.PuntoEmision

    '        'U_EREC_CREADO 
    '        GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"

    '        GRPO.FolioNumber = _oDocumento.Secuencial

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

    Public Function LLenarGridDatosAdicionales(listInfoAdicional As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDatoAdicionalNotaCredito)) As Boolean

        Try
            oForm = rsboApp.Forms.Item("frmDocumentoNC")
            oForm.Freeze(True)

            Try
                oForm.DataSources.DataTables.Add("dtDocsDA")
            Catch ex As Exception
            End Try

            NumeroPedido = ""
            oForm.DataSources.DataTables.Item("dtDocsDA").Clear()
            oForm.DataSources.DataTables.Item("dtDocsDA").Columns.Add("Nombre", Left(SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 253))
            oForm.DataSources.DataTables.Item("dtDocsDA").Columns.Add("Valor", Left(SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 253))

            oForm.DataSources.DataTables.Item("dtDocsDA").Rows.Add(listInfoAdicional.Count)

            Dim i As Integer = 0
            For Each Info As Entidades.wsEDoc_ConsultaRecepcion.ENTDatoAdicionalNotaCredito In listInfoAdicional

                'If Info.Nombre.ToLower = Functions.VariablesGlobales._NombreCampoPedidoInfoAd.ToLower Then
                '    NumeroPedido = Info.Descripcion.ToString
                'End If

                oForm.DataSources.DataTables.Item("dtDocsDA").SetValue("Nombre", i, IIf(IsNothing(Info.Nombre), "", Left(Info.Nombre, 253)))
                oForm.DataSources.DataTables.Item("dtDocsDA").SetValue("Valor", i, IIf(IsNothing(Info.Descripcion), "", Left(Info.Descripcion, 253)))

                i += 1
            Next
            '
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGridDA").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocsDA")
            oGrid.Item.Enabled = False
            oGrid.Item.FromPane = 3
            oGrid.Item.ToPane = 3


            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Caption = "Nombre"

            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Valor"

            oGrid.AutoResizeColumns()

            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al cargar datos adicionales: " + ex.Message.ToString(), "frmDocumento")
            Return False
        End Try

        Return True
    End Function

    Public Function LLenarGridDatosAdicionalesExistente(IdUdo As String) As Boolean

        Try
            oForm = rsboApp.Forms.Item("frmDocumentoNC")
            oForm.Freeze(True)


            NumeroPedido = ""
            Try
                oForm.DataSources.DataTables.Add("dtDocsDA")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocsDA").Clear()
            oForm.DataSources.DataTables.Item("dtDocsDA").Columns.Add("Nombre", Left(SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 253))
            oForm.DataSources.DataTables.Item("dtDocsDA").Columns.Add("Valor", Left(SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 253))
            oForm.DataSources.DataTables.Item("dtDocsDA").Rows.Clear()


            Dim QueryDetalle As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryDetalle = "SELECT  A.""U_Nombre"", A.""U_Valor"" "
                QueryDetalle += "  FROM ""@GS2_NCR"" A "
                QueryDetalle += "  WHERE A.""DocEntry"" =  " + IdUdo
            Else
                QueryDetalle = "SELECT  A.U_Nombre, A.U_Valor "
                QueryDetalle += "  FROM ""@GS2_NCR"" A WITH(NOLOCK)"
                QueryDetalle += "  WHERE A.DocEntry =  " + IdUdo
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocsDA").ExecuteQuery(QueryDetalle)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error dtDocsDA: " + ex.Message.ToString() + " - Query: " + QueryDetalle, "frmDocumentoNC")
            End Try


            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGridDA").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocsDA")
            oGrid.Item.Enabled = False
            oGrid.Item.FromPane = 3
            oGrid.Item.ToPane = 3


            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Caption = "Nombre"

            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Valor"

            oGrid.AutoResizeColumns()

            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al cargar datos adicionales: " + ex.Message.ToString(), "frmDocumentoNC")
            Return False
        End Try

        Return True
    End Function

End Class

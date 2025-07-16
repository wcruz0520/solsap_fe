Imports Entidades
'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Imports System.IO
Imports System.Threading
Imports System.Globalization
Imports System.Xml

Public Class frmDocumentosRecibidosXML
    Public oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    '
    Private listaDocumentoPorUsuario As List(Of Entidades.DocumentoTipo)
    Private listaFCs As New List(Of Factura)
    'Private listaFCsFecha As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
    Private listaREs As New List(Of Retencion)
    Private listaNCs As New List(Of NotaCredito)

    Public listaDetalleArtiulos As New List(Of Entidades.DetalleArticulo)
    Private oFactura As Factura
    Private oNotaDeCredito As NotaCredito
    Private oRetencion As Retencion

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

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormularioDocumentosRecibidos()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        rsboApp.SetStatusBarMessage(NombreAddon + " - Consultando, Espere Por favor!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        If RecorreFormulario(rsboApp, "frmDocumentosRecibidosXML") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentosRecibidosXML.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentosRecibidosXML").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidosXML")

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

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            'Dim typeEx, idForm As String
            'typeEx = oFuncionesB1.FormularioActivo(idForm)
            If pVal.FormTypeEx = "frmDocumentosRecibidosXML" Then
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
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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
                                oForm = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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

                                Case "oGrid"
                                    ofila = pVal.Row
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                                    Dim oGrid As SAPbouiCOM.Grid = oFor.Items.Item("oGrid").Specific
                                    oGrid.Rows.SelectedRows.Add(ofila)
                                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                        ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                                        ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                                    Next
                                Case "btnAnt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) - 1, TotalDocs)
                                Case "btnSig"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) + 1, TotalDocs)
                                Case "btnPri"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), 1, TotalDocs)
                                Case "btnUlt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbNP.Caption), TotalDocs)
                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pVal.ColUID = "RUC" And pVal.BeforeAction = True Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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
                                    Utilitario.Util_Log.Escribir_Log("Query Buscar SN Cliente/Proveedor: " + sQueryProveedor.ToString + " Resultado:" + _sQueryProveedor.ToString, "frmDocumentosRecibidosXML")
                                    If _sQueryProveedor = "PROVEEDOR" Then
                                        sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")

                                    Else
                                        sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                                        Utilitario.Util_Log.Escribir_Log("Query Buscar SN Cliente: " + QueryExisteCliente.ToString + " Resultado:" + sCardCode.ToString, "frmDocumentosRecibidosXML")
                                    End If

                                Else
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                                    Utilitario.Util_Log.Escribir_Log("Query Buscar SN Proveedor: " + QueryExisteProveedor.ToString + " Resultado:" + sCardCode.ToString, "frmDocumentosRecibidosXML")
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
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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
                                Utilitario.Util_Log.Escribir_Log("Query Buscar SN Cliente/Proveedor: " + sQueryProveedor.ToString + " Resultado:" + _sQueryProveedor.ToString, "frmDocumentosRecibidosXML")
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

                                rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Documento de " + sNombre + ", por favor espere..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                If oDataTable.GetValue(0, ofila).ToString() = "Factura" Then
                                    Dim results As List(Of Factura)
                                    results = listaFCs.FindAll(Function(column) column.FacturaCabecera._claveAcceso = sClaveAcceso)
                                    For Each oFac As Factura In results
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
                                        ofrmDocumentoXML.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oFactura, ofila)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumentoXML.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If

                                ElseIf oDataTable.GetValue(0, ofila).ToString() = "Nota de Crédito" Then
                                    Dim results As List(Of NotaCredito)
                                    results = listaNCs.FindAll(Function(column) column.NotaCreditoCabecera._claveAcceso = sClaveAcceso)
                                    For Each oNC As NotaCredito In results
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
                                        ofrmDocumentoNCXML.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oNotaDeCredito, ofila)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumentoNCXML.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                                    End If

                                ElseIf oDataTable.GetValue(0, ofila).ToString() = "Retención de Cliente" Then
                                    Dim results As List(Of Retencion)
                                    results = listaREs.FindAll(Function(column) column.RetCabecera._claveAcceso = sClaveAcceso)
                                    For Each oRE As Retencion In results
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
                                        ofrmDocumentoREXML.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oRetencion, ofila)
                                    Else
                                        idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                                        ofrmDocumentoREXML.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
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

                Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                oDataTable.Rows.Clear()

                'MarcarVistosDocumentosPendientes()

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
    Private Sub MarcarVistosDocumentosPendientes()

        Try
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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

            Utilitario.Util_Log.Escribir_Log("QueryV: " + QueryV.ToString(), "frmDocumentosRecibidosXML")
            Try
                oForm.DataSources.DataTables.Item("dtVIS").ExecuteQuery(QueryV)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error QueryV : " + ex.Message.ToString(), "frmDocumentosRecibidosXML")
            End Try

            Dim U_IdGS As Integer = 0
            Dim IdDocumentoRecibido As String = ""
            Dim TipoDoc As Integer = 0
            Dim ClaveAcceso As String = ""
            Dim Mensaje As String = ""

            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos pendientes de marcar integrados : " + dtVIS.Rows.Count().ToString(), "frmDocumentosRecibidosXML")

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

                Utilitario.Util_Log.Escribir_Log("U_IdGS : " + U_IdGS.ToString(), "frmDocumentosRecibidosXML")
                Utilitario.Util_Log.Escribir_Log("IdDocumentoRecibido : " + IdDocumentoRecibido.ToString(), "frmDocumentosRecibidosXML")
                Utilitario.Util_Log.Escribir_Log("TipoDoc : " + TipoDoc.ToString(), "frmDocumentosRecibidosXML")
                Utilitario.Util_Log.Escribir_Log("ClaveAcceso : " + ClaveAcceso.ToString(), "frmDocumentosRecibidosXML")

                If codDoc = 1 Then
                    Try
                        ofrmDocumento.MarcarVisto(U_IdGS, codDoc, Mensaje, IdDocumentoRecibido)
                        Utilitario.Util_Log.Escribir_Log("ReProceso Visto(Integrado FC) en EDOC: " + ClaveAcceso + Mensaje.ToString(), "frmDocumentosRecibidosXML")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error ofrmDocumento.MarcarVisto : " + ex.Message.ToString(), "frmDocumentosRecibidosXML")
                    End Try

                ElseIf codDoc = 3 Then
                    ofrmDocumentoNC.MarcarVisto(U_IdGS, codDoc, Mensaje, IdDocumentoRecibido)
                    Utilitario.Util_Log.Escribir_Log("ReProceso Visto(Integrado NC) en EDOC: " + ClaveAcceso + Mensaje.ToString(), "frmDocumentosRecibidosXML")
                ElseIf codDoc = 2 Then
                    ofrmDocumentoRE.MarcarVisto(U_IdGS, codDoc, Mensaje, IdDocumentoRecibido)
                    Utilitario.Util_Log.Escribir_Log("ReProceso Visto(Integrado RT) en EDOC: " + ClaveAcceso + Mensaje.ToString(), "frmDocumentosRecibidosXML")
                End If

                rsboApp.StatusBar.SetText(NombreAddon + " - Documento  " + ClaveAcceso + " Actualizado", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Next

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Reproceso Marcar Integrados eDoc: " + ex.Message.ToString(), "frmDocumentosRecibidosXML")
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

            If String.IsNullOrEmpty(_WS_RecepcionCargaEstados) Then _WS_RecepcionCargaEstados = "1"
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
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
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

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific

                    Dim z = CargarEntidadFCs()

                    If Not z Is Nothing Then
                        i = z.Count
                        listaFCs = DirectCast(z, List(Of Factura))
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
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
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

                    mensaje = ""
                    Dim z = CargarEntidadNCs()

                    If Not z Is Nothing Then
                        i = z.Count
                        listaNCs = DirectCast(z, List(Of NotaCredito))
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

                    Dim z = CargarEntidadRTs()

                    If Not z Is Nothing Then
                        i = z.Count
                        listaREs = DirectCast(z, List(Of Retencion))
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

                Case "05"
                    'Dim z = WS.ConsultarND(_WS_RecepcionClave, LicTradNum, "01", mensaje).ToList
                    'If Not z Is Nothing Then
                    '    i = z.Count
                    '    listaNDs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaDebito))
                    'Else
                    '    Return False
                    'End If
                Case "06"
                    'Dim z = WS.ConsultarGR(_WS_RecepcionClave, LicTradNum, "01", mensaje).ToList
                    'If Not z Is Nothing Then
                    '    i = z.Count
                    '    listaGRs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTGuiaRemision))
                    'Else
                    '    Return False
                    'End If
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

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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

                For Each oFactura As Factura In listaFCs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    Dim sRUC As String = oFactura.FacturaCabecera._ruc
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
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_NUM_AUTOR"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_NUM_AUTOR = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_NO_AUTORI"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_NO_AUTORI = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_SYP_NroAuto
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Try
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NROAUTOO"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NroAuto"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            End Try

                        Else
                            Try
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SYP_NROAUTO = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SYP_NroAuto = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            End Try

                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_TM_NAUT"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_TM_NAUT = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_HBT_AUT_FAC"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_HBT_AUT_FAC = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SS_NumAut"" = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SS_NumAut = '" + oFactura.FacturaCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    End If


                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, "DocEntry", "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Factura")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, CDate(oFactura.FacturaCabecera._fechaEmision))
                    If String.IsNullOrEmpty(oFactura.FacturaCabecera._FechaAutorizacion) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, CDate(oFactura.FacturaCabecera._fechaEmision))
                    Else
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, CDate(oFactura.FacturaCabecera._FechaAutorizacion))
                    End If


                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oFactura.FacturaCabecera._estab + "-" + oFactura.FacturaCabecera._ptoEmi + "-" + oFactura.FacturaCabecera._secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oFactura.FacturaCabecera._ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(oFactura.FacturaCabecera._RazonSocial, 250).Trim)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oFactura.FacturaCabecera._importeTotal))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oFactura.FacturaCabecera._claveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oFactura.FacturaCabecera._NumeroAutorizacion)


                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)

                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Factura de " + oFactura.FacturaCabecera._RazonSocial
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Next
                i = 0

            ElseIf tipoDoc = "04" Then
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)

                For Each oNC As NotaCredito In listaNCs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    sQuery = ""
                    Dim sQuerySucursal = ""
                    Dim sRUC As String = oNC.NotaCreditoCabecera._ruc
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
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_NUM_AUTOR"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_NUM_AUTOR = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_NO_AUTORI"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_NO_AUTORI = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_SYP_NroAuto
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Try
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SYP_NROAUTO"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SYP_NroAuto"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            End Try

                        Else
                            Try
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SYP_NROAUTO = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SYP_NroAuto = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                            End Try

                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_TM_NAUT"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_TM_NAUT = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_HBT_AUT_FAC"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_HBT_AUT_FAC = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SS_NumAut"" = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SS_NumAut = '" + oNC.NotaCreditoCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    End If

                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, "DocEntry", "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Nota de Crédito")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, CDate(oNC.NotaCreditoCabecera._fechaEmision))

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, CDate(oNC.NotaCreditoCabecera._FechaAutorizacion))

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oNC.NotaCreditoCabecera._estab + "-" + oNC.NotaCreditoCabecera._ptoEmi + "-" + oNC.NotaCreditoCabecera._secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oNC.NotaCreditoCabecera._ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Trim(Left(oNC.NotaCreditoCabecera._RazonSocial, 250).Trim))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oNC.NotaCreditoCabecera._valorModificacion))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oNC.NotaCreditoCabecera._claveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oNC.NotaCreditoCabecera._NumeroAutorizacion)
                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Nota de Crédito de " + oNC.NotaCreditoCabecera._RazonSocial

                Next
                i = 0
            ElseIf tipoDoc = "07" Then
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)


                For Each oRE As Retencion In listaREs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    Dim sucursal As String = ""
                    sQuery = ""
                    sSucursal = ""
                    Dim sQuerySucursal = ""
                    Dim sRUC As String = oRE.RetCabecera._ruc
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
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_CXS_NUM_AUTO_RETE"" = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_CXS_NUM_AUTO_RETE = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        IdCol = "DocNum"
                        'U_NUM_AUT
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_NUM_AUT"" = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_NUM_AUT = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_FX_AUTO_RETENCION
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            '
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""OPDF"" where ""U_FX_AUTO_RETENCION"" = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from OPDF WITH(NOLOCK) where U_FX_AUTO_RETENCION = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                        IdCol = "DocNum"
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""@TM_LE_RETVH"" where ""U_TM_CASRI"" = '" + oRE.RetCabecera._NumeroAutorizacion + "' and ""U_TM_STATUS""='Borrador' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from ""@TM_LE_RETVH"" WITH(NOLOCK) where U_TM_CASRI = '" + oRE.RetCabecera._NumeroAutorizacion + "' and U_TM_STATUS='Borrador' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        'U_FX_AUTO_RETENCION
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            '
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""OPDF"" where ""U_HBT_NUM_AUT"" = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from OPDF WITH(NOLOCK) where U_HBT_NUM_AUT = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        IdCol = "DocNum"
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_SS_AutRetRec"" = '" + oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_SS_AutRetRec = '" +  oRE.RetCabecera._NumeroAutorizacion + "' ORDER BY 1 DESC"
                        End If
                    End If

                    Dim numretener As String

                    Dim odetalle As New RetDetalleImpuestos

                    numretener = odetalle._numDocSustento
                    Dim valorRetenido As Decimal = 0

                    For Each detalle As RetDetalleImpuestos In oRE.RetDetalleImp

                        valorRetenido += detalle._valorRetenido
                        numretener = detalle._numDocSustento

                    Next

                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, IdCol, "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If


                    'sucursal = oFuncionesB1.getRSvalue(sSucursal, "Sucursal", "")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Retención de Cliente")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, CDate(oRE.RetCabecera._fechaEmision))

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, CDate(oRE.RetCabecera._FechaAutorizacion))

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oRE.RetCabecera._estab + "-" + oRE.RetCabecera._ptoEmi + "-" + oRE.RetCabecera._secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oRE.RetCabecera._ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Trim(Left(oRE.RetCabecera._RazonSocial, 250)))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(valorRetenido))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oRE.RetCabecera._claveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oRE.RetCabecera._NumeroAutorizacion)

                    'oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRe", i, IIf(IsNothing(odetalle.NumDocRetener), "", odetalle.NumDocRetener))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRetener", i, IIf(String.IsNullOrEmpty(numretener), "0", numretener))

                    If Not String.IsNullOrEmpty(DocPreliminar) Then

                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Retención Recibida de " + oRE.RetCabecera._RazonSocial
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
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

        If FormUID = "frmDocumentosRecibidosXML" Then
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidosXML")

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
                            QueryidDocRec = "SELECT ""DocEntry"" from ""@GS_NCR"" where ""U_IdGS"" ='" + IdDocumento + "'"
                        Else
                            QueryidDocRec = "SELECT DocEntry from ""@GS_NCR"" WITH(NOLOCK) where U_IdGS ='" + IdDocumento + "'"
                        End If
                        idDocRec = oFuncionesB1.getRSvalue(QueryidDocRec, "DocEntry", "")
                        ofrmDocumentoNC.ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(idDocRec, 1)

                    ElseIf TipoDocumento = 2 Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QueryidDocRec = "SELECT ""DocEntry"" from ""@GS_RER"" where ""U_IdGS"" ='" + IdDocumento + "'"
                        Else
                            QueryidDocRec = "SELECT DocEntry from ""@GS_RER"" WITH(NOLOCK) where U_IdGS ='" + IdDocumento + "'"
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
            If eventInfo.FormUID = "frmDocumentosRecibidosXML" And eventInfo.ItemUID = "oGrid" Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                If eventInfo.BeforeAction Then
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    ofila = eventInfo.Row
                    Dim oFor As SAPbouiCOM.Form
                    oFor = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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

                    Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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
                    If typeEx = "frmDocumentosRecibidosXML" Then
                        If ofila > 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
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
                                    'Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                    'results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    'listaDetalleArtiulos.Clear()
                                    'If results.Count > 0 Then
                                    '    For Each oFC As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                    '        For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFC.ENTDetalleFactura
                                    '            listaDetalleArtiulos.Add(New Entidades.DetalleArticulo(sClaveAcceso, oDetalle.CodigoPrincipal,
                                    '                                                              oDetalle.CodigoAuxiliar, oDetalle.Descripcion, oDetalle.Cantidad, oDetalle.PrecioUnitario, oDetalle.Descuento, oDetalle.PrecioTotalSinImpuesto))
                                    '        Next
                                    '    Next
                                    'End If
                                    'ofrmMapeo.CargaFormularioMapeo(sRUC, sCardCode, sNombre, listaDetalleArtiulos, ofila, "FV")
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
                    If typeEx = "frmDocumentosRecibidosXML" Then
                        If ofila >= 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosRecibidosXML")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                            Dim tipoDocumento As String = oDataTable.GetValue(0, ofila).ToString()
                            Dim QueryidDocUDO As String = ""
                            Dim idDocUDO As String = ""
                            If tipoDocumento = "Factura" Then
                                Dim results As List(Of Factura)
                                results = listaFCs.FindAll(Function(column) column.FacturaCabecera._claveAcceso = sClaveAcceso)
                                For Each oFac As Factura In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then

                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_FVR"" where ""U_IdGS"" ='" + oFac.FacturaCabecera._DocEntry.ToString + "'"
                                        Else
                                            QueryidDocUDO = "SELECT DocEntry from ""@GS_FVR"" WITH(NOLOCK) where U_IdGS ='" + oFac.FacturaCabecera._DocEntry.ToString + "'"
                                        End If
                                        idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")
                                        If idDocUDO <> "0" Then
                                            If ActualizadoEstado_DocumentoRecibido_Factura(idDocUDO, "docMarcado") Then
                                                If (ofrmDocumentoXML.ActualizadoEstadoUdoFacturaXML(oFac.FacturaCabecera._DocEntry.ToString, "Marcado")) Then
                                                    'ofrmDocumento.MarcarVisto(idDocRec, 1, mensaje, oFac.IdFactura.ToString)
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("FE")
                                                    'rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                                End If

                                            End If

                                        Else
                                            If Guarda_DocumentoRecibido_Factura(oFac) Then
                                                If ofrmDocumentoXML.ActualizadoEstadoUdoFacturaXML(oFac.FacturaCabecera._DocEntry.ToString, "Marcado") Then
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
                                Dim results As List(Of NotaCredito)
                                results = listaNCs.FindAll(Function(column) column.NotaCreditoCabecera._claveAcceso = sClaveAcceso)
                                For Each oNC As NotaCredito In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_NCR"" where ""U_IdGS"" ='" + oNC.NotaCreditoCabecera._DocEntry.ToString + "'"
                                        Else
                                            QueryidDocUDO = "SELECT DocEntry from ""@GS_NCR"" WITH(NOLOCK) where U_IdGS ='" + oNC.NotaCreditoCabecera._DocEntry.ToString + "'"
                                        End If
                                        idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")
                                        If idDocUDO <> "0" Then
                                            If ActualizadoEstado_DocumentoRecibido_NCredito(idDocUDO, "docMarcado") Then
                                                If ofrmDocumentoNCXML.ActualizadoEstadoUdoNotaCreditoXML(oNC.NotaCreditoCabecera._DocEntry, "Marcado") Then
                                                    'ofrmDocumentoNC.MarcarVisto(idDocUDO, 3, mensaje, oNC.IdNotaCredito)
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("NE")
                                                End If
                                            End If

                                        Else
                                            If Guarda_DocumentoRecibido_NotaCredito(oNC) Then

                                                If ofrmDocumentoNCXML.ActualizadoEstadoUdoNotaCreditoXML(oNC.NotaCreditoCabecera._DocEntry, "Marcado") Then
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
                                Dim results As List(Of Retencion)
                                results = listaREs.FindAll(Function(column) column.RetCabecera._claveAcceso = sClaveAcceso)
                                For Each oRE As Retencion In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            QueryidDocUDO = "SELECT ""DocEntry"" from ""@GS_RER"" where ""U_IdGS"" ='" + oRE.RetCabecera._DocEntry.ToString + "'"
                                        Else
                                            QueryidDocUDO = "SELECT DocEntry from ""@GS_RER"" WITH(NOLOCK) where U_IdGS ='" + oRE.RetCabecera._DocEntry.ToString + "'"
                                        End If
                                        idDocUDO = oFuncionesB1.getRSvalue(QueryidDocUDO, "DocEntry", "")

                                        If idDocUDO <> "0" Then

                                            If ActualizadoEstado_DocumentoRecibido_Retencion(idDocUDO, "docMarcado") Then
                                                If ofrmDocumentoREXML.ActualizadoEstadoUdoRetencionXML(oRE.RetCabecera._DocEntry, "Marcado") Then
                                                    'ofrmDocumentoNC.MarcarVisto(idDocUDO, 3, mensaje, oNC.IdNotaCredito)
                                                    oDataTable.Rows.Remove(ofila)
                                                    CargaDocumentosFormato("RE")
                                                End If
                                            End If

                                        Else
                                            If Guarda_DocumentoRecibido_Retencion(oRE) Then
                                                If ofrmDocumentoREXML.ActualizadoEstadoUdoRetencionXML(oRE.RetCabecera._DocEntry, "Marcado") Then
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

        End Try

    End Sub

    Public Function Guarda_DocumentoRecibido_Factura(ByVal oFactura As Factura) As Boolean

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
            cardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + oFactura.FacturaCabecera._ruc.ToString() + "'", "CardCode", "")
        Else
            cardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + oFactura.FacturaCabecera._ruc.ToString() + "'", "CardCode", "")
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
        For Each facImpuesto As FacturaCabeceraImpuestos In oFactura.FacturaCabecera._impuestos
            If facImpuesto._codigo = 2 Then
                If facImpuesto._codigoPorcentaje = 2 Or facImpuesto._codigoPorcentaje = 3 Then
                    BaseImponibleIVA += facImpuesto._baseImponible
                    Iva += facImpuesto._valor
                ElseIf facImpuesto._codigoPorcentaje = 0 Then
                    BaseImponible0 += facImpuesto._baseImponible
                ElseIf facImpuesto._codigoPorcentaje = 6 Then
                    BaseImponibleNoObjeto += facImpuesto._baseImponible
                ElseIf facImpuesto._codigoPorcentaje = 7 Then
                    BaseImponibleExento += facImpuesto._baseImponible
                End If
            End If
        Next
        For Each facImpuesto As FacturaCabeceraImpuestos In oFactura.FacturaCabecera._impuestos
            If facImpuesto._codigo = 3 Then
                '  BaseImponibleICE += facImpuesto.BaseImponible
                ICE += facImpuesto._valor
            End If
        Next

        Try
            oFuncionesAddon.GuardaLOG("REE", oFactura.FacturaCabecera._claveAcceso.ToString(), "Registrando Factura Marcada", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("REE " + oFactura.FacturaCabecera._claveAcceso.ToString() + " Registrando Factura Marcada", "frmDocumentosRecibidosXML")
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidosXML")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralData.SetProperty("U_RUC", oFactura.FacturaCabecera._ruc.ToString())
            'oGeneralData.SetProperty("U_Nombre", Left(oFactura.NombreComercial.ToString().Replace(Chr(10), ""), 99))
            oGeneralData.SetProperty("U_Nombre", Left(oFactura.FacturaCabecera._RazonSocial.ToString().Replace(Chr(10), ""), 99))
            oGeneralData.SetProperty("U_CardCode", cardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", "")
            oGeneralData.SetProperty("U_ClaAcc", oFactura.FacturaCabecera._claveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oFactura.FacturaCabecera._NumeroAutorizacion.ToString())
            oGeneralData.SetProperty("U_FecAut", oFactura.FacturaCabecera._FechaAutorizacion.ToString())
            'oGeneralData.SetProperty("U_FechaS", Date.Now.ToString())
            oGeneralData.SetProperty("U_NumDoc", oFactura.FacturaCabecera._estab.ToString() + "-" + oFactura.FacturaCabecera._ptoEmi.ToString() + "-" + oFactura.FacturaCabecera._secuencial.ToString())
            'oGeneralData.SetProperty("U_FPrelim", oForm.Items.Item("txtFPre").Specific.Value.ToString())
            oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleIVA, 2).ToString())))
            oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponible0, 2).ToString())))
            oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleNoObjeto, 2).ToString())))
            oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleExento, 2).ToString())))
            oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oFactura.FacturaCabecera._totalSinImpuestos, 2).ToString())))
            oGeneralData.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oFactura.FacturaCabecera._totalDescuento, 2).ToString())))
            oGeneralData.SetProperty("U_ICE", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleICE, 2).ToString())))
            oGeneralData.SetProperty("U_IVA", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(Iva, 2).ToString())))
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oFactura.FacturaCabecera._importeTotal, 2).ToString())))
            oGeneralData.SetProperty("U_rTades", "0")
            oGeneralData.SetProperty("U_rPDesc", "0")
            oGeneralData.SetProperty("U_rDesc", "0")
            oGeneralData.SetProperty("U_rGast", "0")
            oGeneralData.SetProperty("U_rImp", "0")
            oGeneralData.SetProperty("U_rTotal", "0")
            oGeneralData.SetProperty("U_IdGS", oFactura.FacturaCabecera._DocEntry.ToString())
            oGeneralData.SetProperty("U_Sincro", "0")
            oGeneralData.SetProperty("U_Tipo", "Factura de Servicio")
            oGeneralData.SetProperty("U_SincroE", "1")
            oGeneralData.SetProperty("U_Estado", "docMarcado")


            oChildren = oGeneralData.Child("GS0_FVR")
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            For Each facDetalle As FacturaDetalle In oFactura.facturaDetalle
                oChild = oChildren.Add
                Dim CodAux As String = ""
                Dim CodPrin As String = ""
                If facDetalle._codigoAuxiliar = Nothing Then
                    CodAux = "N/A"
                Else
                    CodAux = facDetalle._codigoAuxiliar.ToString()
                End If
                If facDetalle._codigoPrincipal = Nothing Then
                    CodPrin = "N/A"
                Else
                    CodPrin = facDetalle._codigoPrincipal.ToString()
                End If
                oChild.SetProperty("U_CodPrin", Left(CodPrin, 99))
                oChild.SetProperty("U_CodAuxi", CodAux)
                'oChild.SetProperty("U_CodSAP", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_Descripc", Left(facDetalle._descripcion.ToString(), 100))
                oChild.SetProperty("U_Cantid", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle._cantidad.ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle._precioUnitario.ToString())))
                oChild.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle._descuento.ToString())))
                oChild.SetProperty("U_Total", Convert.ToDouble(frmDocumento.formatDecimal(facDetalle._precioTotalSinImpuesto.ToString())))
            Next
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("REE", oFactura.FacturaCabecera._claveAcceso.ToString(), "Se creo registro de Factura Marcada, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("REE", oFactura.FacturaCabecera._claveAcceso.ToString(), "Ocurrior un error al crear registro de Factura marcada UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Factura marcada en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("REE " + oFactura.FacturaCabecera._claveAcceso.ToString() + " Ocurrior un error al crear registro de Factura marcada UDO: " + ex.Message.ToString(), "frmDocumentosRecibidosXML")
            mensaje = ex.Message.ToString
            Return False
        End Try
    End Function

    Public Function Guarda_DocumentoRecibido_NotaCredito(ByVal oNC As NotaCredito) As Boolean

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
            cardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + oNC.NotaCreditoCabecera._ruc.ToString() + "'", "CardCode", "")
        Else
            cardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + oNC.NotaCreditoCabecera._ruc.ToString() + "'", "CardCode", "")
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
        For Each NCImpuesto As NotaCreditoCabeceraImpuesto In oNC.NotaCreditoCabecera._impuestos
            If NCImpuesto._codigo = 2 Then
                If NCImpuesto._codigoPorcentaje = 2 Or NCImpuesto._codigoPorcentaje = 3 Then
                    BaseImponibleIVA += NCImpuesto._baseImponible
                    Iva += NCImpuesto._valor
                ElseIf NCImpuesto._codigoPorcentaje = 0 Then
                    BaseImponible0 += NCImpuesto._baseImponible
                ElseIf NCImpuesto._codigoPorcentaje = 6 Then
                    BaseImponibleNoObjeto += NCImpuesto._baseImponible
                ElseIf NCImpuesto._codigoPorcentaje = 7 Then
                    BaseImponibleExento += NCImpuesto._baseImponible
                End If
            End If
        Next
        For Each NCImpuesto As NotaCreditoCabeceraImpuesto In oNC.NotaCreditoCabecera._impuestos
            If NCImpuesto._codigo = 3 Then
                '  BaseImponibleICE += facImpuesto.BaseImponible
                ICE += NCImpuesto._valor
            End If
        Next

        Try
            oFuncionesAddon.GuardaLOG("NCR", oNC.NotaCreditoCabecera._claveAcceso.ToString(), "Registrando NotaCredito Marcada", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oForm = rsboApp.Forms.Item("frmDocumentosRecibidosXML")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralData.SetProperty("U_RUC", oNC.NotaCreditoCabecera._ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oNC.NotaCreditoCabecera._RazonSocial.ToString().Replace(Chr(10), ""), 99))
            oGeneralData.SetProperty("U_CardCode", cardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", "")
            oGeneralData.SetProperty("U_ClaAcc", oNC.NotaCreditoCabecera._claveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oNC.NotaCreditoCabecera._NumeroAutorizacion.ToString())
            oGeneralData.SetProperty("U_FecAut", oNC.NotaCreditoCabecera._FechaAutorizacion.ToString())
            'oGeneralData.SetProperty("U_FechaS", Date.Now.ToString())
            oGeneralData.SetProperty("U_NumDoc", oNC.NotaCreditoCabecera._estab.ToString() + "-" + oNC.NotaCreditoCabecera._ptoEmi.ToString() + "-" + oNC.NotaCreditoCabecera._secuencial.ToString())
            'oGeneralData.SetProperty("U_FPrelim", oForm.Items.Item("txtFPre").Specific.Value.ToString())
            oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleIVA, 2).ToString())))
            oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponible0, 2).ToString())))
            oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleNoObjeto, 2).ToString())))
            oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleExento, 2).ToString())))
            oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oNC.NotaCreditoCabecera._totalSinImpuestos, 2).ToString())))
            oGeneralData.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oNC.NotaCreditoCabecera._totalDescuento, 2).ToString())))
            oGeneralData.SetProperty("U_ICE", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(BaseImponibleICE, 2).ToString())))
            oGeneralData.SetProperty("U_IVA", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(Iva, 2).ToString())))
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(frmDocumento.formatDecimal(Math.Round(oNC.NotaCreditoCabecera._valorModificacion, 2).ToString())))
            oGeneralData.SetProperty("U_rTades", "0")
            oGeneralData.SetProperty("U_rPDesc", "0")
            oGeneralData.SetProperty("U_rDesc", "0")
            oGeneralData.SetProperty("U_rGast", "0")
            oGeneralData.SetProperty("U_rImp", "0")
            oGeneralData.SetProperty("U_rTotal", "0")
            oGeneralData.SetProperty("U_IdGS", oNC.NotaCreditoCabecera._DocEntry.ToString())
            oGeneralData.SetProperty("U_Sincro", "0")
            oGeneralData.SetProperty("U_Tipo", "NC de Servicio")
            oGeneralData.SetProperty("U_SincroE", "1")
            oGeneralData.SetProperty("U_Estado", "docMarcado")


            oChildren = oGeneralData.Child("GS0_NCR")
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            For Each NCDetalle As NotaCreditoDetalle In oNC.NotaCreditoDetalle
                oChild = oChildren.Add
                Dim CodAux As String = ""
                Dim CodPrin As String = ""
                If NCDetalle._codigoAdicional = Nothing Then
                    CodAux = "N/A"
                Else
                    CodAux = NCDetalle._codigoAdicional.ToString()
                End If
                If NCDetalle._codigoInterno = Nothing Then
                    CodPrin = "N/A"
                Else
                    CodPrin = NCDetalle._codigoInterno.ToString()
                End If
                oChild.SetProperty("U_CodPrin", Left(CodPrin, 99))
                oChild.SetProperty("U_CodAuxi", CodAux)
                'oChild.SetProperty("U_CodSAP", odt.GetValue(2, i).ToString())
                oChild.SetProperty("U_Descripc", Left(NCDetalle._descripcion.ToString(), 100))
                oChild.SetProperty("U_Cantid", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle._cantidad.ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle._precioUnitario.ToString())))
                oChild.SetProperty("U_Desc", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle._descuento.ToString())))
                oChild.SetProperty("U_Total", Convert.ToDouble(frmDocumento.formatDecimal(NCDetalle._precioTotalSinImpuesto.ToString())))
            Next
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("REE", oNC.NotaCreditoCabecera._claveAcceso.ToString(), "Se creo registro de NotaCredito Marcada, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("REE", oNC.NotaCreditoCabecera._claveAcceso.ToString(), "Ocurrior un error al crear registro de nota de credito marcada UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar nota de credito marcada en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            mensaje = ex.Message.ToString
            Return False
        End Try
    End Function

    Public Function Guarda_DocumentoRecibido_Retencion(ByVal oRE As Retencion) As Boolean
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim DocEntryFacturaRecibida_UDO As String = 0


        Try
            rsboApp.StatusBar.SetText(NombreAddon + "- Creando registro de Pago Recibido(Retencion) Recibida UDO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", oRE.RetCabecera._claveAcceso, "Creando registro de Pago Recibido(Retencion) Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Try
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sCardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM OCRD where ""LicTradNum"" = '" + oRE.RetCabecera._ruc.ToString() + "' AND ""CardType"" = 'C' ", "CardCode", "")
                Else
                    sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + oRE.RetCabecera._ruc.ToString() + "' AND CardType = 'C' ", "CardCode", "")
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("REE - error al obtener CardCode DM : " + ex.ToString, "frmDocumentosRecibidosXML")
            End Try

            Dim numDoc As String = oRE.RetCabecera._estab + "-" + oRE.RetCabecera._ptoEmi + "-" + oRE.RetCabecera._secuencial

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralData.SetProperty("Code", conta)
            oGeneralData.SetProperty("U_RUC", oRE.RetCabecera._ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oRE.RetCabecera._RazonSocial.ToString(), 99))
            oGeneralData.SetProperty("U_CardCode", sCardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", oForm.Items.Item("lbMapp").Specific.Value.ToString())
            oGeneralData.SetProperty("U_ClaAcc", oRE.RetCabecera._claveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oRE.RetCabecera._NumeroAutorizacion.ToString())
            oGeneralData.SetProperty("U_FecAut", oRE.RetCabecera._FechaAutorizacion.ToString())
            oGeneralData.SetProperty("U_NumDoc", numDoc.ToString())
            'oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString())

            oGeneralData.SetProperty("U_IdGS", oRE.RetCabecera._DocEntry.ToString())
            oGeneralData.SetProperty("U_Sincro", 0)
            oGeneralData.SetProperty("U_SincroE", 1)
            oGeneralData.SetProperty("U_Estado", "docMarcado")

            Dim totalRetencion As Decimal = 0

            oChildren = oGeneralData.Child("GS0_RER")
            For Each detalleRet As RetDetalleImpuestos In oRE.RetDetalleImp

                Dim ejeFiscal As String = oRE.RetCabecera._periodoFiscal
                Dim _numDocRet As String = ""
                oChild = oChildren.Add
                oChild.SetProperty("U_CodRet", detalleRet._codigoRetencion.ToString())
                If IsNothing(detalleRet._numDocSustento) Then
                    _numDocRet = "0"
                Else
                    _numDocRet = detalleRet._numDocSustento.ToString
                End If
                oChild.SetProperty("U_NumDocRe", _numDocRet)
                oChild.SetProperty("U_Fecha", detalleRet._fechaEmisionDocSustento.ToString())
                oChild.SetProperty("U_pFiscal", ejeFiscal.ToString())
                oChild.SetProperty("U_Base", Convert.ToDouble(frmDocumentoRE.formatDecimal(detalleRet._baseImponible.ToString())))
                If detalleRet._codigoRetencion = "1" Then
                    oChild.SetProperty("U_Impuesto", "RENTA")
                Else
                    oChild.SetProperty("U_Impuesto", "IVA")
                End If
                oChild.SetProperty("U_Porcent", Convert.ToDouble(frmDocumentoRE.formatDecimal(detalleRet._porcentajeRetener.ToString())))
                oChild.SetProperty("U_valorR", Convert.ToDouble(frmDocumentoRE.formatDecimal(detalleRet._valorRetenido.ToString())))

                totalRetencion += detalleRet._valorRetenido

            Next
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(frmDocumentoRE.formatDecimal(totalRetencion.ToString())))
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("PRR", oRE.RetCabecera._claveAcceso, "Se creo registro de Pago Recibido(Retencion) marcada UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Se creo registro de Pago Recibido(Retencion) marcada UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", oRE.RetCabecera._claveAcceso, "Ocurrior un error al crear registro de Pago Recibido(Retencion) marcada UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Pago Recibido(Retencion) marcada en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstado_DocumentoRecibido_Factura(ByRef DocEntryUDO As String, Estado As String) As Boolean
        'Dim oBusP As SAPbobsCOM.BusinessPartners = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", Estado)
            oGeneralData.SetProperty("U_FechaS", Integer.Parse(Date.Now.ToString("yyyyMMdd")))
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstado_DocumentoRecibido_NCredito(ByRef DocEntryUDO As String, Estado As String) As Boolean
        'Dim oBusP As SAPbobsCOM.BusinessPartners = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", Estado)
            oGeneralData.SetProperty("U_FechaS", Integer.Parse(Date.Now.ToString("yyyyMMdd")))
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstado_DocumentoRecibido_Retencion(ByRef DocEntryUDO As String, Estado As String) As Boolean
        'Dim oBusP As SAPbobsCOM.BusinessPartners = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", Estado)
            oGeneralData.SetProperty("U_FechaS", Integer.Parse(Date.Now.ToString("yyyyMMdd")))
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
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
                query = "select ""DocEntry"" from OPCH where ""U_NUM_AUTOR""='" + xoFactura.AutorizacionSRI.ToString + "'"
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
#Enable Warning BC42105 ' La función 'MarcarDocumentosContabilizadosManualFC' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Public Function MarcarDocumentosContabilizadosManualNC(ByVal Lista As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito))
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
        WS.Url = _WS_RecepcionCambiarEstado
        Try
            For Each xoNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In Lista
                Dim query As String = "", rsult As String = ""
                query = "select ""DocEntry"" from ORPC where ""U_NUM_AUTOR""='" + xoNC.AutorizacionSRI.ToString + "'"
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
#Enable Warning BC42105 ' La función 'MarcarDocumentosContabilizadosManualNC' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Public Function MarcarDocumentosContabilizadosManualRT(ByVal Lista As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))
        Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
        WS.Url = _WS_RecepcionCambiarEstado
        Try
            For Each xoRT As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In Lista
                Dim query As String = "", rsult As String = ""
                query = "select TOP 1 T0.""DocEntry"" from ORCT T0 INNER JOIN RCT3 T1 ON T1.""DocNum""=T0.""DocEntry"" WHERE T1.""U_CXS_NUM_AUTO_RETE""='" + xoRT.AutorizacionSRI.ToString + "'"
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

    Shared Function customCertValidation(ByVal sender As Object, _
                                             ByVal cert As X509Certificate, _
                                             ByVal chain As X509Chain, _
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


    Public Function CargarEntidadFCs() As List(Of Factura)

        Dim ListaFacturas As New List(Of Factura)

        Dim Cabfc As SAPbobsCOM.Recordset = Nothing

        Dim docentry As Integer = 0

        Dim txtRUC As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific

        Dim LicTradNum As String = ""
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
        Else
            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
        End If

        Try

            Dim spCab As String
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                spCab = "call" & rCompany.CompanyDB & ".GS_CABECERA_XML ('" + LicTradNum + "')"
            Else
                spCab = "EXEC GS_CABECERA_XML '" + LicTradNum + "'"
            End If

            Cabfc = oFuncionesB1.getRecordSet(spCab)

            If Cabfc.RecordCount > 0 Then

                While (Cabfc.EoF = False)

                    Dim Factura As New Factura
                    Factura.FacturaCabecera = New FacturaCabecera

                    Factura.FacturaCabecera._impuestos = New List(Of FacturaCabeceraImpuestos)

                    Factura.facturaDetalle = New List(Of FacturaDetalle)

                    docentry = Cabfc.Fields.Item("DocEntry").Value.ToString()
                    Factura.FacturaCabecera._DocEntry = docentry
                    Factura.FacturaCabecera._NumeroAutorizacion = Cabfc.Fields.Item("NumAutorizacion").Value.ToString()
                    Factura.FacturaCabecera._FechaAutorizacion = IIf(IsNothing(Cabfc.Fields.Item("FechaAutorizacion").Value.ToString()), Cabfc.Fields.Item("FechaEmision").Value.ToString(), Cabfc.Fields.Item("FechaAutorizacion").Value.ToString())
                    Factura.FacturaCabecera._RazonSocial = Cabfc.Fields.Item("RazonSocial").Value.ToString()
                    Factura.FacturaCabecera._ruc = Cabfc.Fields.Item("Ruc").Value.ToString()
                    Factura.FacturaCabecera._claveAcceso = Cabfc.Fields.Item("ClaveAcceso").Value.ToString()
                    Factura.FacturaCabecera._estab = Cabfc.Fields.Item("Establecimiento").Value.ToString()
                    Factura.FacturaCabecera._ptoEmi = Cabfc.Fields.Item("PuntoEmision").Value.ToString()
                    Factura.FacturaCabecera._secuencial = Cabfc.Fields.Item("Secuencial").Value.ToString()
                    Factura.FacturaCabecera._fechaEmision = Cabfc.Fields.Item("FechaEmision").Value.ToString()
                    Factura.FacturaCabecera._dirEstablecimiento = Cabfc.Fields.Item("DireccionEstablecimiento").Value.ToString()
                    Factura.FacturaCabecera._contribuyenteEspecial = Cabfc.Fields.Item("ContribuyenteEspecial").Value.ToString()
                    Factura.FacturaCabecera._razonSocialComprador = Cabfc.Fields.Item("RazonSocialComprador").Value.ToString()
                    Factura.FacturaCabecera._identificacionComprador = Cabfc.Fields.Item("IdentificacionComprador").Value.ToString()
                    Factura.FacturaCabecera._direccionComprador = Cabfc.Fields.Item("DireccionComprador").Value.ToString()
                    Factura.FacturaCabecera._totalSinImpuestos = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("TotalSinImpuesto").Value.ToString()))
                    Factura.FacturaCabecera._totalDescuento = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("TotalDescuento").Value.ToString()))
                    Factura.FacturaCabecera._importeTotal = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("ImporteTotal").Value.ToString()))

                    'BASE 0
                    If Not String.IsNullOrEmpty(Cabfc.Fields.Item("Cod0").Value.ToString()) Then

                        Dim facCabImp As New FacturaCabeceraImpuestos

                        facCabImp._codigo = Cabfc.Fields.Item("Cod0").Value.ToString()
                        facCabImp._codigoPorcentaje = Cabfc.Fields.Item("CodPorc0").Value.ToString()
                        facCabImp._baseImponible = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("BaseImp0").Value.ToString()))
                        facCabImp._tarifa = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("Tarifa0").Value.ToString()))
                        facCabImp._valor = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("Valor0").Value.ToString()))

                        Factura.FacturaCabecera._impuestos.Add(facCabImp)
                    End If

                    'BASE 12
                    If Not String.IsNullOrEmpty(Cabfc.Fields.Item("Cod12").Value.ToString()) Then

                        Dim facCabImp As New FacturaCabeceraImpuestos

                        facCabImp._codigo = Cabfc.Fields.Item("Cod12").Value.ToString()
                        facCabImp._codigoPorcentaje = Cabfc.Fields.Item("CodPorc12").Value.ToString()
                        facCabImp._baseImponible = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("BaseImp12").Value.ToString()))
                        facCabImp._tarifa = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("Tarifa12").Value.ToString()))
                        facCabImp._valor = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("Valor12").Value.ToString()))

                        Factura.FacturaCabecera._impuestos.Add(facCabImp)
                    End If

                    'BASE 8
                    If Not String.IsNullOrEmpty(Cabfc.Fields.Item("Cod8").Value.ToString()) Then

                        Dim facCabImp As New FacturaCabeceraImpuestos

                        facCabImp._codigo = Cabfc.Fields.Item("Cod12").Value.ToString()
                        facCabImp._codigoPorcentaje = Cabfc.Fields.Item("CodPorc8").Value.ToString()
                        facCabImp._baseImponible = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("BaseImp8").Value.ToString()))
                        facCabImp._tarifa = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("Tarifa8").Value.ToString()))
                        facCabImp._valor = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("Valor8").Value.ToString()))

                        Factura.FacturaCabecera._impuestos.Add(facCabImp)
                    End If

                    'BASE NOI
                    If Not String.IsNullOrEmpty(Cabfc.Fields.Item("CodNoi").Value.ToString()) Then

                        Dim facCabImp As New FacturaCabeceraImpuestos

                        facCabImp._codigo = Cabfc.Fields.Item("CodNoi").Value.ToString()
                        facCabImp._codigoPorcentaje = Cabfc.Fields.Item("CodPorcNoi").Value.ToString()
                        facCabImp._baseImponible = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("BaseImpNoi").Value.ToString()))
                        facCabImp._tarifa = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("TarifaNoi").Value.ToString()))
                        facCabImp._valor = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("ValorNoi").Value.ToString()))

                        Factura.FacturaCabecera._impuestos.Add(facCabImp)
                    End If

                    'BASE EXENTA
                    If Not String.IsNullOrEmpty(Cabfc.Fields.Item("CodExe").Value.ToString()) Then

                        Dim facCabImp As New FacturaCabeceraImpuestos

                        facCabImp._codigo = Cabfc.Fields.Item("CodExe").Value.ToString()
                        facCabImp._codigoPorcentaje = Cabfc.Fields.Item("CodPorcExe").Value.ToString()
                        facCabImp._baseImponible = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("BaseImpExe").Value.ToString()))
                        facCabImp._tarifa = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("TarifaExe").Value.ToString()))
                        facCabImp._valor = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("ValorExe").Value.ToString()))

                        Factura.FacturaCabecera._impuestos.Add(facCabImp)
                    End If

                    'BASE ICE
                    If Not String.IsNullOrEmpty(Cabfc.Fields.Item("CodIce").Value.ToString()) Then

                        Dim facCabImp As New FacturaCabeceraImpuestos

                        facCabImp._codigo = Cabfc.Fields.Item("CodIce").Value.ToString()
                        facCabImp._codigoPorcentaje = Cabfc.Fields.Item("CodPorcIce").Value.ToString()
                        facCabImp._baseImponible = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("BaseImpIce").Value.ToString()))
                        facCabImp._tarifa = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("TarifaIce").Value.ToString()))
                        facCabImp._valor = Convert.ToDouble(formatDecimal(Cabfc.Fields.Item("ValorIce").Value.ToString()))

                        Factura.FacturaCabecera._impuestos.Add(facCabImp)
                    End If

                    'llenar detalle de la entidad
                    Dim spDet As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        spDet = "call" & rCompany.CompanyDB & ".GS_DETALLE_XML (" + docentry.ToString + ")"
                    Else
                        spDet = "EXEC GS_DETALLE_XML " + docentry.ToString
                    End If

                    Dim Detfc As SAPbobsCOM.Recordset
                    Detfc = oFuncionesB1.getRecordSet(spDet)

                    If Detfc.RecordCount > 0 Then
                        While (Detfc.EoF = False)
                            Dim FacDet As New FacturaDetalle

                            FacDet._impuestos = New List(Of FacturaDetalleImpuesto)

                            FacDet._codigoPrincipal = Detfc.Fields.Item("CodigoPrincipal").Value.ToString()
                            FacDet._codigoAuxiliar = Detfc.Fields.Item("CodigoAuxiliar").Value.ToString()
                            FacDet._descripcion = Detfc.Fields.Item("Descripcion").Value.ToString()
                            FacDet._cantidad = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("Cantidad").Value.ToString()))
                            FacDet._precioUnitario = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("PrecioUnitario").Value.ToString()))
                            FacDet._descuento = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("Descuento").Value.ToString()))
                            FacDet._precioTotalSinImpuesto = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("TotalSinImpuesto").Value.ToString()))

                            'impuestos
                            If Not String.IsNullOrEmpty(Detfc.Fields.Item("Cod").Value.ToString()) Then

                                Dim FacDetImp As New FacturaDetalleImpuesto

                                FacDetImp._codigo = Detfc.Fields.Item("Cod").Value.ToString()
                                FacDetImp._codigoPorcentaje = Detfc.Fields.Item("CodPorc").Value.ToString()
                                FacDetImp._baseImponible = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("BaseImp").Value.ToString()))
                                FacDetImp._tarifa = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("Tarifa").Value.ToString()))
                                FacDetImp._valor = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("Valor").Value.ToString()))

                                FacDet._impuestos.Add(FacDetImp)

                            End If

                            'impuestoice
                            If Not String.IsNullOrEmpty(Detfc.Fields.Item("CodIce").Value.ToString()) Then

                                Dim FacDetImp As New FacturaDetalleImpuesto

                                FacDetImp._codigo = Detfc.Fields.Item("CodIce").Value.ToString()
                                FacDetImp._codigoPorcentaje = Detfc.Fields.Item("CodPorcIce").Value.ToString()
                                FacDetImp._baseImponible = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("BaseImpIce").Value.ToString()))
                                FacDetImp._tarifa = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("TarifaIce").Value.ToString()))
                                FacDetImp._valor = Convert.ToDouble(formatDecimal(Detfc.Fields.Item("ValorIce").Value.ToString()))

                                FacDet._impuestos.Add(FacDetImp)

                            End If

                            Factura.facturaDetalle.Add(FacDet)
                            Detfc.MoveNext()
                        End While
                    End If

                    ListaFacturas.Add(Factura)
                    Cabfc.MoveNext()
                End While
            End If

            Return ListaFacturas
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + "Error - al obtener datos de la fc desde el udo docentry: " + docentry + " - " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try

    End Function

    Public Function CargarEntidadNCs() As List(Of NotaCredito)

        Dim ListaNotasCreditos As New List(Of NotaCredito)

        Dim CabNc As SAPbobsCOM.Recordset = Nothing

        Dim docentryNc As Integer = 0


        Dim txtRUC As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific

        Dim LicTradNum As String = ""
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
        Else
            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
        End If

        Try

            Dim spCabNC As String
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                spCabNC = "call" & rCompany.CompanyDB & ".GS_CABECERANC_XML ('" + LicTradNum + "')"
            Else
                spCabNC = "EXEC GS_CABECERANC_XML '" + LicTradNum + "'"
            End If

            CabNc = oFuncionesB1.getRecordSet(spCabNC)

            If CabNc.RecordCount > 0 Then

                While (CabNc.EoF = False)

                    Dim NC As New NotaCredito
                    NC.NotaCreditoCabecera = New NotaCreditoCabecera

                    NC.NotaCreditoCabecera._impuestos = New List(Of NotaCreditoCabeceraImpuesto)

                    NC.NotaCreditoDetalle = New List(Of NotaCreditoDetalle)

                    docentryNc = CabNc.Fields.Item("DocEntry").Value.ToString()
                    NC.NotaCreditoCabecera._DocEntry = docentryNc
                    NC.NotaCreditoCabecera._NumeroAutorizacion = CabNc.Fields.Item("NumAutorizacion").Value.ToString()
                    NC.NotaCreditoCabecera._FechaAutorizacion = IIf(IsNothing(CabNc.Fields.Item("FechaAutorizacion").Value.ToString()), "", CabNc.Fields.Item("FechaAutorizacion").Value.ToString())
                    NC.NotaCreditoCabecera._RazonSocial = CabNc.Fields.Item("RazonSocial").Value.ToString()
                    NC.NotaCreditoCabecera._ruc = CabNc.Fields.Item("Ruc").Value.ToString()
                    NC.NotaCreditoCabecera._claveAcceso = CabNc.Fields.Item("ClaveAcceso").Value.ToString()
                    NC.NotaCreditoCabecera._estab = CabNc.Fields.Item("Establecimiento").Value.ToString()
                    NC.NotaCreditoCabecera._ptoEmi = CabNc.Fields.Item("PuntoEmision").Value.ToString()
                    NC.NotaCreditoCabecera._secuencial = CabNc.Fields.Item("Secuencial").Value.ToString()
                    NC.NotaCreditoCabecera._fechaEmision = CabNc.Fields.Item("FechaEmision").Value.ToString()
                    NC.NotaCreditoCabecera._dirEstablecimiento = CabNc.Fields.Item("DireccionEstablecimiento").Value.ToString()
                    NC.NotaCreditoCabecera._contribuyenteEspecial = CabNc.Fields.Item("ContribuyenteEspecial").Value.ToString()
                    NC.NotaCreditoCabecera._razonSocialComprador = CabNc.Fields.Item("RazonSocialComprador").Value.ToString()
                    NC.NotaCreditoCabecera._identificacionComprador = CabNc.Fields.Item("IdentificacionComprador").Value.ToString()
                    NC.NotaCreditoCabecera._direccionComprador = CabNc.Fields.Item("DireccionComprador").Value.ToString()
                    NC.NotaCreditoCabecera._totalSinImpuestos = CDec(CabNc.Fields.Item("TotalSinImpuesto").Value.ToString())
                    NC.NotaCreditoCabecera._totalDescuento = CDec(CabNc.Fields.Item("TotalDescuento").Value.ToString())
                    NC.NotaCreditoCabecera._importeTotal = CDec(CabNc.Fields.Item("ImporteTotal").Value.ToString())

                    NC.NotaCreditoCabecera._CodDocMod = CabNc.Fields.Item("CodDocMod").Value.ToString()
                    NC.NotaCreditoCabecera._numDocModificado = CabNc.Fields.Item("NumDocMod").Value.ToString()
                    NC.NotaCreditoCabecera._fechaEmisionDocSustento = CabNc.Fields.Item("FecDocMod").Value.ToString()
                    NC.NotaCreditoCabecera._motivo = CabNc.Fields.Item("Motivo").Value.ToString()
                    NC.NotaCreditoCabecera._valorModificacion = CDec(CabNc.Fields.Item("ValorModificacion").Value.ToString())

                    'BASE 0
                    If Not String.IsNullOrEmpty(CabNc.Fields.Item("Cod0").Value.ToString()) Then

                        Dim NCCabImp As New NotaCreditoCabeceraImpuesto

                        NCCabImp._codigo = CabNc.Fields.Item("Cod0").Value.ToString()
                        NCCabImp._codigoPorcentaje = CabNc.Fields.Item("CodPorc0").Value.ToString()
                        NCCabImp._baseImponible = CabNc.Fields.Item("BaseImp0").Value.ToString()
                        NCCabImp._tarifa = CabNc.Fields.Item("Tarifa0").Value.ToString()
                        NCCabImp._valor = CabNc.Fields.Item("Valor0").Value.ToString()

                        NC.NotaCreditoCabecera._impuestos.Add(NCCabImp)
                    End If

                    'BASE 12
                    If Not String.IsNullOrEmpty(CabNc.Fields.Item("Cod12").Value.ToString()) Then

                        Dim NCCabImp As New NotaCreditoCabeceraImpuesto

                        NCCabImp._codigo = CabNc.Fields.Item("Cod12").Value.ToString()
                        NCCabImp._codigoPorcentaje = CabNc.Fields.Item("CodPorc12").Value.ToString()
                        NCCabImp._baseImponible = CabNc.Fields.Item("BaseImp12").Value.ToString()
                        NCCabImp._tarifa = CabNc.Fields.Item("Tarifa12").Value.ToString()
                        NCCabImp._valor = CabNc.Fields.Item("Valor12").Value.ToString()

                        NC.NotaCreditoCabecera._impuestos.Add(NCCabImp)
                    End If

                    'BASE 8
                    If Not String.IsNullOrEmpty(CabNc.Fields.Item("Cod8").Value.ToString()) Then

                        Dim NCCabImp As New NotaCreditoCabeceraImpuesto

                        NCCabImp._codigo = CabNc.Fields.Item("Cod8").Value.ToString()
                        NCCabImp._codigoPorcentaje = CabNc.Fields.Item("CodPorc8").Value.ToString()
                        NCCabImp._baseImponible = CabNc.Fields.Item("BaseImp8").Value.ToString()
                        NCCabImp._tarifa = CabNc.Fields.Item("Tarifa8").Value.ToString()
                        NCCabImp._valor = CabNc.Fields.Item("Valor8").Value.ToString()

                        NC.NotaCreditoCabecera._impuestos.Add(NCCabImp)
                    End If

                    'BASE NOI
                    If Not String.IsNullOrEmpty(CabNc.Fields.Item("CodNoi").Value.ToString()) Then

                        Dim NCCabImp As New NotaCreditoCabeceraImpuesto

                        NCCabImp._codigo = CabNc.Fields.Item("CodNoi").Value.ToString()
                        NCCabImp._codigoPorcentaje = CabNc.Fields.Item("CodPorcNoi").Value.ToString()
                        NCCabImp._baseImponible = CabNc.Fields.Item("BaseImpNoi").Value.ToString()
                        NCCabImp._tarifa = CabNc.Fields.Item("TarifaNoi").Value.ToString()
                        NCCabImp._valor = CabNc.Fields.Item("ValorNoi").Value.ToString()

                        NC.NotaCreditoCabecera._impuestos.Add(NCCabImp)
                    End If

                    'BASE EXENTA
                    If Not String.IsNullOrEmpty(CabNc.Fields.Item("CodExe").Value.ToString()) Then

                        Dim NCCabImp As New NotaCreditoCabeceraImpuesto

                        NCCabImp._codigo = CabNc.Fields.Item("CodExe").Value.ToString()
                        NCCabImp._codigoPorcentaje = CabNc.Fields.Item("CodPorcExe").Value.ToString()
                        NCCabImp._baseImponible = CabNc.Fields.Item("BaseImpExe").Value.ToString()
                        NCCabImp._tarifa = CabNc.Fields.Item("TarifaExe").Value.ToString()
                        NCCabImp._valor = CabNc.Fields.Item("ValorExe").Value.ToString()

                        NC.NotaCreditoCabecera._impuestos.Add(NCCabImp)
                    End If

                    'BASE ICE
                    If Not String.IsNullOrEmpty(CabNc.Fields.Item("CodIce").Value.ToString()) Then

                        Dim NCCabImp As New NotaCreditoCabeceraImpuesto

                        NCCabImp._codigo = CabNc.Fields.Item("CodIce").Value.ToString()
                        NCCabImp._codigoPorcentaje = CabNc.Fields.Item("CodPorcIce").Value.ToString()
                        NCCabImp._baseImponible = CabNc.Fields.Item("BaseImpIce").Value.ToString()
                        NCCabImp._tarifa = CabNc.Fields.Item("TarifaIce").Value.ToString()
                        NCCabImp._valor = CabNc.Fields.Item("ValorIce").Value.ToString()

                        NC.NotaCreditoCabecera._impuestos.Add(NCCabImp)
                    End If

                    'llenar detalle de la entidad
                    Dim spDetNc As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        spDetNc = "call" & rCompany.CompanyDB & ".GS_DETALLENC_XML (" + docentryNc.ToString + ")"
                    Else
                        spDetNc = "EXEC GS_DETALLENC_XML " + docentryNc.ToString
                    End If

                    Dim DetNc As SAPbobsCOM.Recordset
                    DetNc = oFuncionesB1.getRecordSet(spDetNc)

                    If DetNc.RecordCount > 0 Then
                        While (DetNc.EoF = False)
                            Dim NcDet As New NotaCreditoDetalle

                            NcDet._impuestos = New List(Of NotaCreditoDetalleImpuesto)

                            NcDet._codigoInterno = DetNc.Fields.Item("CodigoPrincipal").Value.ToString()
                            NcDet._codigoAdicional = DetNc.Fields.Item("CodigoAuxiliar").Value.ToString()
                            NcDet._descripcion = DetNc.Fields.Item("Descripcion").Value.ToString()
                            NcDet._cantidad = DetNc.Fields.Item("Cantidad").Value.ToString()
                            NcDet._precioUnitario = DetNc.Fields.Item("PrecioUnitario").Value.ToString()
                            NcDet._descuento = DetNc.Fields.Item("Descuento").Value.ToString()
                            NcDet._precioTotalSinImpuesto = DetNc.Fields.Item("TotalSinImpuesto").Value.ToString()

                            'impuestos
                            If Not String.IsNullOrEmpty(DetNc.Fields.Item("Cod").Value.ToString()) Then

                                Dim NcDetImp As New NotaCreditoDetalleImpuesto

                                NcDetImp._codigo = DetNc.Fields.Item("Cod").Value.ToString()
                                NcDetImp._codigoPorcentaje = DetNc.Fields.Item("CodPorc").Value.ToString()
                                NcDetImp._baseImponible = DetNc.Fields.Item("BaseImp").Value.ToString()
                                NcDetImp._tarifa = DetNc.Fields.Item("Tarifa").Value.ToString()
                                NcDetImp._valor = DetNc.Fields.Item("Valor").Value.ToString()

                                NcDet._impuestos.Add(NcDetImp)

                            End If

                            'impuestoice
                            If Not String.IsNullOrEmpty(DetNc.Fields.Item("CodIce").Value.ToString()) Then

                                Dim NcDetImp As New FacturaDetalleImpuesto

                                NcDetImp._codigo = DetNc.Fields.Item("CodIce").Value.ToString()
                                NcDetImp._codigoPorcentaje = DetNc.Fields.Item("CodPorcIce").Value.ToString()
                                NcDetImp._baseImponible = DetNc.Fields.Item("BaseImpIce").Value.ToString()
                                NcDetImp._tarifa = DetNc.Fields.Item("TarifaIce").Value.ToString()
                                NcDetImp._valor = DetNc.Fields.Item("ValorIce").Value.ToString()

                                DetNc._impuestos.Add(NcDetImp)

                            End If

                            NC.NotaCreditoDetalle.Add(NcDet)
                            DetNc.MoveNext()
                        End While
                    End If

                    ListaNotasCreditos.Add(NC)
                    CabNc.MoveNext()
                End While
            End If

            Return ListaNotasCreditos
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + "Error - al obtener datos de la fc desde el udo docentry: " + docentryNc + " - " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try






    End Function

    Public Function CargarEntidadRTs() As List(Of Retencion)

        Dim ListaRetenciones As New List(Of Retencion)

        Dim CabRT As SAPbobsCOM.Recordset = Nothing

        Dim docentryRT As Integer = 0

        Dim txtRUC As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific

        Dim LicTradNum As String = ""
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
        Else
            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'C' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
        End If

        Try

            Dim spCabRT As String
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                spCabRT = "call" & rCompany.CompanyDB & ".GS_CABECERART_XML ('" + LicTradNum + "')"
            Else
                spCabRT = "EXEC GS_CABECERART_XML '" + LicTradNum + "'"
            End If

            CabRT = oFuncionesB1.getRecordSet(spCabRT)

            If CabRT.RecordCount > 0 Then

                While (CabRT.EoF = False)

                    Dim RT As New Retencion
                    RT.RetCabecera = New RetCabecera

                    RT.RetDetalleImp = New List(Of RetDetalleImpuestos)

                    docentryRT = CabRT.Fields.Item("DocEntry").Value.ToString()
                    RT.RetCabecera._DocEntry = docentryRT
                    RT.RetCabecera._NumeroAutorizacion = CabRT.Fields.Item("NumAutorizacion").Value.ToString()
                    RT.RetCabecera._FechaAutorizacion = IIf(IsNothing(CabRT.Fields.Item("FechaAutorizacion").Value.ToString()), "", CabRT.Fields.Item("FechaAutorizacion").Value.ToString())
                    RT.RetCabecera._RazonSocial = CabRT.Fields.Item("RazonSocial").Value.ToString()
                    RT.RetCabecera._ruc = CabRT.Fields.Item("Ruc").Value.ToString()
                    RT.RetCabecera._claveAcceso = CabRT.Fields.Item("ClaveAcceso").Value.ToString()
                    RT.RetCabecera._estab = CabRT.Fields.Item("Establecimiento").Value.ToString()
                    RT.RetCabecera._ptoEmi = CabRT.Fields.Item("PuntoEmision").Value.ToString()
                    RT.RetCabecera._secuencial = CabRT.Fields.Item("Secuencial").Value.ToString()
                    RT.RetCabecera._fechaEmision = CabRT.Fields.Item("FechaEmision").Value.ToString()
                    RT.RetCabecera._dirEstablecimiento = CabRT.Fields.Item("DireccionEstablecimiento").Value.ToString()
                    RT.RetCabecera._contribuyenteEspecial = CabRT.Fields.Item("ContribuyenteEspecial").Value.ToString()
                    RT.RetCabecera._razonSocialSujetoRetenido = CabRT.Fields.Item("RazonSocialSujetoRetenido").Value.ToString()
                    RT.RetCabecera._identificacionSujetoRetenido = CabRT.Fields.Item("IdentificacionSujetoRetenido").Value.ToString()
                    RT.RetCabecera._periodoFiscal = CabRT.Fields.Item("Periodo").Value.ToString()




                    'llenar detalle de la entidad
                    Dim spDetRT As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        spDetRT = "call" & rCompany.CompanyDB & ".GS_DETALLERT_XML (" + docentryRT.ToString + ")"
                    Else
                        spDetRT = "EXEC GS_DETALLERT_XML " + docentryRT.ToString
                    End If

                    Dim DetRT As SAPbobsCOM.Recordset
                    DetRT = oFuncionesB1.getRecordSet(spDetRT)

                    If DetRT.RecordCount > 0 Then

                        While (DetRT.EoF = False)

                            Dim RtDet As New RetDetalleImpuestos

                            RtDet._codigo = CInt(DetRT.Fields.Item("Codigo").Value.ToString())
                            RtDet._codigoRetencion = DetRT.Fields.Item("CodigoRetenido").Value.ToString()
                            RtDet._baseImponible = DetRT.Fields.Item("BaseImponible").Value.ToString()
                            RtDet._porcentajeRetener = DetRT.Fields.Item("PorcentajeRetenido").Value.ToString()
                            RtDet._valorRetenido = DetRT.Fields.Item("ValorRetenido").Value.ToString()
                            RtDet._codDocSustento = DetRT.Fields.Item("CodDocSus").Value.ToString()
                            RtDet._numDocSustento = DetRT.Fields.Item("NumDocSus").Value.ToString()
                            RtDet._fechaEmisionDocSustento = CDate(DetRT.Fields.Item("FecEmiDocSus").Value.ToString())

                            RT.RetDetalleImp.Add(RtDet)
                            DetRT.MoveNext()
                        End While
                    End If

                    ListaRetenciones.Add(RT)
                    CabRT.MoveNext()
                End While
            End If

            Return ListaRetenciones
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + "Error - al obtener datos de la fc desde el udo docentry: " + docentryRT + " - " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
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
End Class

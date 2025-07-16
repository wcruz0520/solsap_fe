Imports System.IO
Imports Negocio

Public Class frmDocumentosEnviados
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Dim LQE As String = ""
    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential
    Dim lnkRutaExp As SAPbouiCOM.LinkedButton
    Dim lblRutaExp As SAPbouiCOM.StaticText
    Private GetfileThreadFE As Threading.Thread
    Dim contExpDoc As Integer

    Dim mensajeRC As String = ""
    Dim DocSRI As String = ""
    Dim NumDocEmi As String = ""

    Private CoreRest As CoreRest

    Dim odtDE As SAPbouiCOM.DataTable
    Dim oUserDataSourceDE As SAPbouiCOM.UserDataSource
    Dim oCFLsDE As SAPbouiCOM.ChooseFromListCollection
    Dim oConsDE As SAPbouiCOM.Conditions
    Dim oConDE As SAPbouiCOM.Condition
    Dim oCFLDE As SAPbouiCOM.ChooseFromList
    'Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParamsDE As SAPbouiCOM.ChooseFromListCreationParams

    Public ListDocEntrys As New List(Of String)
    Dim btnAL2 As SAPbouiCOM.ButtonCombo


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp

        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
            CoreRest = New CoreRest()
            CoreRest.WS_EnvioDocumento = Functions.VariablesGlobales._WsEmisionEcua
            CoreRest.WS_ConsultaDocumento = Functions.VariablesGlobales._WsEmisionConsultaEcua
        End If

    End Sub

    Public Sub CreaFormularioDocumentosEnviados()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmDocumentosEnviados") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentosEnviados.srf"
        xmlDoc.Load(strPath)
        ListDocEntrys.Clear()

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentosEnviados").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmDocumentosEnviados")

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
            oCFLsDE = oForm.ChooseFromLists
            oCFLCreationParamsDE = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParamsDE.MultiSelection = False
            oCFLCreationParamsDE.ObjectType = "2"
            'oCFLCreationParams.ObjectType = "Exx_DEPOTRANS"
            oCFLCreationParamsDE.UniqueID = "CFL1"
            oCFLDE = oCFLsDE.Add(oCFLCreationParamsDE)
            ' Adding Conditions to CFL1
            oConsDE = oCFLDE.GetConditions()

            oConDE = oConsDE.Add()
            oConDE.Alias = "CardType"
            oConDE.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oConDE.CondVal = "C"
            oCFLDE.SetConditions(oConsDE)
            ' END CHOOSE FROM LIST

            Dim txtRSN As SAPbouiCOM.EditText
            txtRSN = oForm.Items.Item("txtRSN").Specific
            oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtRSN.DataBind.SetBound(True, "", "EditDS")
            txtRSN.ChooseFromListUID = "CFL1"
            txtRSN.ChooseFromListAlias = "CardType"

            Dim lnkSN As SAPbouiCOM.LinkedButton
            lnkSN = oForm.Items.Item("lnkSN").Specific
            lnkSN.LinkedObjectType = 2
            lnkSN.Item.LinkTo = "txtRSN"

            Dim cmbTipo As SAPbouiCOM.ComboBox
            cmbTipo = oForm.Items.Item("cbxTipo").Specific
            cmbTipo.ValidValues.Add("0", "Todos")
            cmbTipo.ValidValues.Add("01", "Factura")
            cmbTipo.ValidValues.Add("03", "LQ de Compra")
            cmbTipo.ValidValues.Add("04", "Nota de Crédito")
            cmbTipo.ValidValues.Add("05", "Nota de Débito")
            cmbTipo.ValidValues.Add("06", "Guía de Remisión")
            cmbTipo.ValidValues.Add("07", "Comp. de Retención")
            cmbTipo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim txtFchIni As SAPbouiCOM.EditText
            txtFchIni = oForm.Items.Item("finicial").Specific
            txtFchIni.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtFchFin As SAPbouiCOM.EditText
            txtFchFin = oForm.Items.Item("ffinal").Specific
            txtFchFin.Value = DateTime.Now.ToString("yyyyMMdd")

            'Estados Documentos.
            Dim cmbEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or
                Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                cmbEstado.ValidValues.Add("9", "Todos")
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or
                Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                cmbEstado.ValidValues.Add("", "Todos")
            End If

            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then

                cmbEstado.ValidValues.Add("0", "NO ENVIADO")
                cmbEstado.ValidValues.Add("1", "EN PROCESO")
                cmbEstado.ValidValues.Add("2", "AUTORIZADO")
                cmbEstado.ValidValues.Add("4", "NO ENCONTRADO")
                cmbEstado.ValidValues.Add("6", "CON ERROR")
                cmbEstado.ValidValues.Add("11", "ANULADO")

            Else
                cmbEstado.ValidValues.Add("0", "NO ENVIADO")
                cmbEstado.ValidValues.Add("1", "EN PROCESO")
                cmbEstado.ValidValues.Add("2", "AUTORIZADO")
                cmbEstado.ValidValues.Add("3", "NO AUTORIZADO")
                cmbEstado.ValidValues.Add("4", "VALIDAR DATOS")
                cmbEstado.ValidValues.Add("5", "EN PROCESO SRI")
                cmbEstado.ValidValues.Add("6", "DEVUELTA")
                cmbEstado.ValidValues.Add("7", "ERROR EN RECEPCION")
                cmbEstado.ValidValues.Add("11", "ANULADO")
            End If


            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or
               Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                cmbEstado.Select("9", SAPbouiCOM.BoSearchKey.psk_ByValue)
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or
                Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                cmbEstado.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

            Dim cbxTipDoc As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipDoc").Specific
            cbxTipDoc.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            cbxTipDoc.ValidValues.Add("PDF", "PDF")
            cbxTipDoc.ValidValues.Add("XML", "XML")
            cbxTipDoc.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'cbxTipDoc.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'cmbEstado.Select("PDF", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'Dim lnkRutaExp As SAPbouiCOM.LinkedButton
            lnkRutaExp = oForm.Items.Item("lnkRutaExp").Specific
            lnkRutaExp.Item.Visible = True

            'Dim lblRutaExp As SAPbouiCOM.StaticText
            'lblRutaExp = oForm.Items.Item("lblRutaExp").Specific
            'lblRutaExp.Item.Visible = False
            Dim txtRutaExp As SAPbouiCOM.EditText
            txtRutaExp = oForm.Items.Item("txtRuta").Specific
            txtRutaExp.Value = "C:\SAED\PDF"
            'txtRutaExp.Item.Enabled = False

            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then

                Dim Button As SAPbouiCOM.Button
                'Button = oForm.Items.Item("btnSelec").Specific
                'Button.Item.Visible = False
                Button = oForm.Items.Item("btnRC").Specific
                Button.Item.Visible = False

            End If

            If Functions.VariablesGlobales._ReenviarDocsPantala = "Y" Then
                Dim Button2 As SAPbouiCOM.Button
                Button2 = oForm.Items.Item("btnCEST").Specific
                Button2.Caption = "Reenviar / Consultar Doc"
            End If

            If Functions.VariablesGlobales._ReenviarListaDocEnv = "Y" Then
                Dim btnAgregaL As SAPbouiCOM.Button
                btnAgregaL = oForm.Items.Item("btnAgregaL").Specific
                btnAgregaL.Item.Visible = False

                btnAL2 = oForm.Items.Item("btnAL2").Specific
                btnAL2.Item.Visible = True
                btnAL2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAL2.Item.AffectsFormMode = False
                btnAL2.Item.Visible = True

                btnAL2.ValidValues.Add("Agregar a Lista", "Agregar a Lista")
                btnAL2.ValidValues.Add("Ver Detalle Lista", "Ver Detalle Lista")
                btnAL2.Select("Agregar a Lista", SAPbouiCOM.BoSearchKey.psk_ByValue)

            Else
                Dim btnAgregaL As SAPbouiCOM.Button
                btnAgregaL = oForm.Items.Item("btnAgregaL").Specific
                btnAgregaL.Item.Visible = False

                btnAL2 = oForm.Items.Item("btnAL2").Specific
                btnAL2.Item.Visible = False

            End If

            FormularioDocumentosEnviadosCargarGrid()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub

    Private Sub FormularioDocumentosEnviadosCargarGrid()
        oForm.Freeze(True)

        Dim cbxTipoDoc As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipo").Specific
        Dim cbxEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
        Dim txtfinicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
        Dim txtffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific
        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific
        Dim txtEst As SAPbouiCOM.EditText = oForm.Items.Item("txtEst").Specific
        Dim txtPemi As SAPbouiCOM.EditText = oForm.Items.Item("txtPemi").Specific
        Dim txtSec As SAPbouiCOM.EditText = oForm.Items.Item("txtSec").Specific
        Dim txtSN As SAPbouiCOM.EditText = oForm.Items.Item("txtRSN").Specific


        If (String.IsNullOrEmpty(txtfinicial.Value) Or String.IsNullOrEmpty(txtffinal.Value)) Then
            rsboApp.SetStatusBarMessage("Debe ingresar un rango de fechas a consultar!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        Dim dfechaDesde As Date
        Dim dfechaHasta As Date
        Dim sQuery As String = ""
        Dim sTipoDoc As String = cbxTipoDoc.Value.Trim()
        Dim sEstado As String = cbxEstado.Value.Trim()
        Dim sfolioIni As String = txtfinicial.Value.Trim()
        Dim sfoliofin As String = txtffinal.Value.Trim()
        Dim sEst As String = txtEst.Value.Trim()
        Dim sPemi As String = txtPemi.Value.Trim()
        Dim sSec As String = txtSec.Value.Trim()
        Dim sSN As String = txtSN.Value.Trim()


        If Not oFuncionesB1.BobStringToDate(txtfinicial.Value, dfechaDesde) Then
            rsboApp.SetStatusBarMessage("El formato de las fechas es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        If Not oFuncionesB1.BobStringToDate(txtffinal.Value, dfechaHasta) Then
            rsboApp.SetStatusBarMessage("El formato de las fechas es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            'CALL GS_SAP_FE_ONE_OBTENERDOCUMENTOS ('0','2',{d'2016-06-16'},{d'2017-09-28'})
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                sQuery = "CALL GS_SAP_FE_ONE_ObtenerDocumentos ("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                sQuery = "CALL " & rCompany.CompanyDB & ".GS_SAP_FE_ObtenerDocumentos ("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                sQuery = "CALL GS_SAP_FE_HEI_OBTENERDOCUMENTOS ("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                sQuery = "CALL GS_SAP_FE_SYP_ObtenerDocumentos("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                sQuery = "CALL GS_SAP_FE_TM_ObtenerDocumentos("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                sQuery = "CALL GS_SAP_FE_SS_ObtenerDocumentos("
            End If
            sQuery += "'" + sTipoDoc + "'"
            sQuery += ",'" + sEstado + "'"
            sQuery += ",'" + sEst + "'"
            sQuery += ",'" + sPemi + "'"
            sQuery += ",'" + sSec + "'"
            sQuery += ",'" + sSN + "'"
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta) + ")"
        Else
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                sQuery = "EXEC GS_SAP_FE_ONE_ObtenerDocumentos "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                sQuery = "EXEC GS_SAP_FE_ObtenerDocumentos "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                sQuery = "EXEC GS_SAP_FE_HEI_ObtenerDocumentos "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                sQuery = "EXEC GS_SAP_FE_SYP_ObtenerDocumentos "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                sQuery = "EXEC GS_SAP_FE_TM_ObtenerDocumentos "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                sQuery = "EXEC GS_SAP_FE_SS_ObtenerDocumentos "
            End If

            sQuery += "'" + sTipoDoc + "'"
            sQuery += ",'" + sEstado + "'"
            sQuery += ",'" + sEst + "'"
            sQuery += ",'" + sPemi + "'"
            sQuery += ",'" + sSec + "'"
            sQuery += ",'" + sSN + "'"
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta)
        End If


        'sQuery += "," + Util.FechaSql(fechaDesde)
        'sQuery += "," + Util.FechaSql(fechaHasta)

        Try
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + sQuery, "frmDocumentosEnviados")
                oGrid.DataTable.ExecuteQuery(sQuery)
                Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + sQuery, "frmDocumentosEnviados")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmDocumentosEnviados")
            End Try

            oGrid.Columns.Item(0).Description = "Tipo Documento"
            oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "#"
            oGrid.Columns.Item(1).TitleObject.Caption = "#"
            oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "DocEntry"
            oGrid.Columns.Item(2).TitleObject.Caption = "DocEntry"
            oGrid.Columns.Item(2).Editable = False

            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oEditTextColumn = oGrid.Columns.Item(2)
            'oEditTextColumn.LinkedObjectType = 13
            oEditTextColumn.LinkedObjectType = "SSGRNEW"

            oGrid.Columns.Item(3).Description = "Doc. Num."
            oGrid.Columns.Item(3).TitleObject.Caption = "Doc. Num."
            oGrid.Columns.Item(3).Editable = False

            oGrid.Columns.Item(4).Description = "Folio"
            oGrid.Columns.Item(4).TitleObject.Caption = "Folio"
            oGrid.Columns.Item(4).Editable = False

            oGrid.Columns.Item(5).Description = "Cliente"
            oGrid.Columns.Item(5).TitleObject.Caption = "Cliente"
            oGrid.Columns.Item(5).Editable = False

            oGrid.Columns.Item(6).Description = "Nombre"
            oGrid.Columns.Item(6).TitleObject.Caption = "Nombre"
            oGrid.Columns.Item(6).Editable = False

            oGrid.Columns.Item(7).Description = "Doc. Total"
            oGrid.Columns.Item(7).TitleObject.Caption = "Doc. Total"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).RightJustified = True

            'Dim Colcap As SAPbouiCOM.EditTextColumn
            'Colcap = oGrid.Columns.Item(7)
            'Colcap.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oGrid.Columns.Item(8).Description = "Fecha Emisión"
            oGrid.Columns.Item(8).TitleObject.Caption = "Fecha Emisión"
            oGrid.Columns.Item(8).Editable = False

            oGrid.Columns.Item(9).Description = "Estado Documento"
            oGrid.Columns.Item(9).TitleObject.Caption = "Estado Documento"
            oGrid.Columns.Item(9).Editable = False

            oGrid.Columns.Item(10).Width = 0
            oGrid.Columns.Item(10).Visible = False
            oGrid.Columns.Item(10).Editable = False

            oGrid.Columns.Item(11).Width = 0
            oGrid.Columns.Item(11).Visible = False
            oGrid.Columns.Item(11).Editable = False

            oGrid.Columns.Item(12).Width = 0
            oGrid.Columns.Item(12).Visible = False
            oGrid.Columns.Item(12).Editable = False

            oGrid.Columns.Item(13).Width = 0
            oGrid.Columns.Item(13).Visible = False
            oGrid.Columns.Item(13).Editable = False

            oGrid.Columns.Item(14).Description = "NroDocSustento"
            oGrid.Columns.Item(14).TitleObject.Caption = "N° DocSustento"
            oGrid.Columns.Item(14).Editable = False

            oGrid.Columns.Item(15).Visible = True
            oGrid.Columns.Item(15).Editable = True
            oGrid.Columns.Item(15).Description = "Seleccionar"
            oGrid.Columns.Item(15).TitleObject.Caption = "Seleccionar"
            oGrid.Columns.Item(15).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            'If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
            '    oGrid.Columns.Item(15).Visible = False

            'End If

            oGrid.Columns.Item(16).Editable = True
            oGrid.Columns.Item(16).Width = 30
            oGrid.Columns.Item(16).Description = "Ver Documento"
            oGrid.Columns.Item(16).TitleObject.Caption = "Ver Documento"
            oGrid.Columns.Item(16).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            Dim oComboBoxColumn As SAPbouiCOM.ComboBoxColumn
            oComboBoxColumn = oGrid.Columns.Item(16)
            oComboBoxColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oComboBoxColumn.ValidValues.Add("", "")
            oComboBoxColumn.ValidValues.Add("PDF", "PDF")
            oComboBoxColumn.ValidValues.Add("XML", "XML")
            oComboBoxColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oComboBoxColumn.Editable = True
            'oGrid.Columns.Item(16).Description = "PDF"
            'oGrid.Columns.Item(16).TitleObject.Caption = "PDF"
            'oGrid.Columns.Item(16).Editable = False

            'Dim oEditTextColumnPDF As SAPbouiCOM.EditTextColumn
            'oEditTextColumnPDF = oGrid.Columns.Item(16)
            'oEditTextColumnPDF.LinkedObjectType = 13

            'oGrid.Columns.Item(17).Description = "XML"
            'oGrid.Columns.Item(17).TitleObject.Caption = "XML"
            'oGrid.Columns.Item(17).Editable = False
            oGrid.Columns.Item(17).Width = 0
            oGrid.Columns.Item(17).Visible = False
            oGrid.Columns.Item(17).Editable = False

            oGrid.Columns.Item(18).Description = "Comentario"
            oGrid.Columns.Item(18).TitleObject.Caption = "Comentario"
            oGrid.Columns.Item(18).Editable = False

            'oGrid.Columns.Item(19).Width = 0
            'oGrid.Columns.Item(19).Visible = False
            'oGrid.Columns.Item(19).Editable = True

            oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()
            schk.Checked = False

            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos:" + sQuery + " - " + ex.Message.ToString, "frmDocumentosEnviados")
        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If FormUID = "frmDocumentosEnviados" Then


                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If Not pVal.Before_Action Then

                            Select Case pVal.ItemUID

                                Case "obtnBuscar"

                                    FormularioDocumentosEnviadosCargarGrid()

                                Case "btnCEST"

                                    Dim resp As Integer = 0

                                    If ListDocEntrys.Count > 0 Then

                                        resp = rsboApp.MessageBox("Desea Reenviar los documentos agregados a la lista ?", 1, "SI", "NO")

                                    Else

                                        resp = rsboApp.MessageBox("Desea actualizar los Estados ?", 1, "SI", "NO")

                                    End If

                                    Select Case resp
                                        Case 1

                                            ConsultarEstados()

                                        Case 2


                                    End Select

                                Case "btnSelec"
                                    Dim contador As Integer = 0
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("btnSelec").Specific
                                    If Seleccionar.Caption = "Seleccionar Todo" Then
                                        If SeleccionarDocumentosPendientes(contador) Then
                                            Seleccionar.Caption = "Desmarcar Todo"
                                            rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se seleccionaron " + contador.ToString + " registros para reenviar el correo con los archivos adjuntos.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    Else
                                        If DesmarcarDocumentosPendientes(contador) Then
                                            Seleccionar.Caption = "Seleccionar Todo"
                                            rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se desmarcarón " + contador.ToString + " registros para reenviar el correo con los archivos adjuntos.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            'oForm.Items.Item("btnCon").Enabled = False
                                            rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - No existen registros marcados.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        End If
                                    End If

                                Case "btnRC"
                                    Dim contador As Integer = 0
                                    mensajeRC = ""
                                    If ReenviarCorreosDocMarcados(contador) Then
                                        rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se reenviaron " + contador.ToString + " correos con los archivos adjuntos.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    Else
                                        rsboApp.SetStatusBarMessage(mensajeRC, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    End If

                                Case "obtnCerrar"

                                    oForm.Close()

                                Case "schk"

                                    ExpandirContraer()

                                Case "btExportar"

                                    contExpDoc = 0
                                    Dim cbxTipDoc As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipDoc").Specific
                                    Dim TipDoc As String = cbxTipDoc.Value.Trim()
                                    If TipDoc = "" Then

                                        Exit Sub
                                    End If
                                    If Not String.IsNullOrEmpty(TipDoc) Then

                                        If ListDocEntrys.Count > 0 Then
                                            Dim ValoresListA As String = ""
                                            Dim Combo As SAPbouiCOM.ComboBox
                                            Combo = oForm.Items.Item("cbxTipDoc").Specific
                                            Dim extension As String = ""
                                            If cbxTipDoc.Value = "PDF" Then
                                                extension = "PDF"
                                            ElseIf cbxTipDoc.Value = "XML" Then
                                                extension = "XML"
                                            End If

                                            Try
                                                For j As Integer = 0 To ListDocEntrys.Count - 1
                                                    ValoresListA = ListDocEntrys(j)

                                                    Exportar_PDF_XML(Trim(ValoresListA.Split(";")(7)), extension, Trim(ValoresListA.Split(";")(2)), "", Trim(ValoresListA.Split(";")(1)))
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Documento " + Trim(ValoresListA.Split(";")(2)) + " exportado", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                    ValoresListA = ""
                                                Next
                                                rsboApp.StatusBar.SetText(Functions.VariablesGlobales._vgNombreAddOn + " - Documentos " + extension + " exportados con éxito...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Catch ex As Exception
                                                rsboApp.StatusBar.SetText(Functions.VariablesGlobales._vgNombreAddOn + " - Error al exportar Lista " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End Try

                                            'Exportar_PDF_XML(ClaveAccesoExp, Extension, NumeracionSRI, NombreSN, oGridDetExp.GetValue("NomDocSap", i))
                                            'ListDocEntrys.Add(DocEntry + " ; " + tipotabla + " ; " + numeracionSRI + " ; " + estado + " ; " + comentario + " ; " + Seleccion + " ; " + objtype + " ; " + ClaveAcceso)
                                        Else

                                            If ExportarArchivo(TipDoc) Then
                                                If contExpDoc > 0 Then
                                                    rsboApp.StatusBar.SetText(Functions.VariablesGlobales._vgNombreAddOn + " - Proceso terminado con éxito...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                Else
                                                    rsboApp.SetStatusBarMessage(NombreAddon + " - Por favor seleccionar documentos a exportar... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If

                                            End If

                                        End If


                                    End If
                                Case "lnkRutaExp"


                                    Dim _ruta As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific

                                    Dim link = _ruta.Value.ToString
                                    Dim MiProceso As New System.Diagnostics.Process
                                    MiProceso.Start("explorer.exe", link)


                                Case "btnExplor"

                                    Dim selectFileDialog As New SelectFileDialog("C:\", "", "|*.rpt", DialogType.FOLDER)
                                    selectFileDialog.Open()



                                    If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFolder) Then
                                        Dim s As String
                                        s = selectFileDialog.SelectedFolder
                                        oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
                                        Dim _ruta2 As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific
                                        _ruta2.Value = s
                                    End If
                                    'explorador()

                                    'Case "btnAgregaL2"
                                Case "btnAL2"
                                    Dim caption As SAPbouiCOM.ButtonCombo
                                    caption = oForm.Items.Item("btnAL2").Specific

                                    If caption.Caption = "Agregar a Lista" Then
                                        AgregarDocLista()
                                    Else
                                        OfrmListaAEnviar.CreaFormularioLista(ListDocEntrys)
                                    End If

                                    'ReenviarListaDocs(ListDocEntrys)

                            End Select
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                        If pVal.BeforeAction Then

                            Event_MatrixLinkPressed(pVal)

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If Not pVal.Before_Action Then
                            Try
                                oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
                                Dim cbxTipo As SAPbouiCOM.ComboBox
                                cbxTipo = oForm.Items.Item("cbxTipo").Specific

                                Try
                                    oConsDE = oCFLDE.GetConditions()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al obtener condiciones:" + ex.Message.ToString, "frmImpresionPorBloque")
                                End Try


                                'Dim lbSocio As SAPbouiCOM.StaticText
                                'lbSocio = oForm.Items.Item("lbSocio").Specific



                                If oConsDE.Count > 0 Then 'If there are already user conditions.
                                    If cbxTipo.Value = "07" Or cbxTipo.Value = "03" Then ' SI ES 07, SIGNIFICA QUE ES RETENCION, POR ENDE PAGO RECIBIDO DE CLIENTE
                                        oConsDE.Item(oConsDE.Count - 1).CondVal = "S"
                                        'lbSocio.Caption = " Proveedor:"

                                    Else
                                        oConsDE.Item(oConsDE.Count - 1).CondVal = "C"
                                        'lbSocio.Caption = "Cliente:"

                                    End If
                                End If

                                oCFLDE.SetConditions(oConsDE)
                                Dim txtRuc As SAPbouiCOM.EditText
                                txtRuc = oForm.Items.Item("txtRSN").Specific
                                txtRuc.ChooseFromListUID = "CFL1"
                                txtRuc.ChooseFromListAlias = "CardType"

                            Catch ex As Exception

                            End Try

                            Select Case pVal.ItemUID
                                Case "cbxTipDoc"

                                    oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
                                    Dim cbxTipDoc As SAPbouiCOM.ComboBox
                                    cbxTipDoc = oForm.Items.Item("cbxTipDoc").Specific

                                    Dim txtRuta As SAPbouiCOM.EditText
                                    txtRuta = oForm.Items.Item("txtRuta").Specific

                                    If cbxTipDoc.Value = "PDF" Then
                                        txtRuta.Value = "C:\SAED\PDF"
                                    ElseIf cbxTipDoc.Value = "XML" Then
                                        txtRuta.Value = "C:\SAED\XML"
                                    End If

                                Case "oGrid"

                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    Dim TipoArchivo As String = oGrid.DataTable.GetValue("VerDoc", oGrid.GetDataTableRowIndex(pVal.Row))

                                    Dim ClaveAccesoDoc As String = oGrid.DataTable.GetValue("ClaveAcceso", oGrid.GetDataTableRowIndex(pVal.Row))
                                    Dim estadoDoc As String = oGrid.DataTable.GetValue("EstadoDoc", oGrid.GetDataTableRowIndex(pVal.Row))

                                    'Dim oGridDet2 As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                                    'Dim contLineas As Integer = oGridDet2.Rows.Count - 1

                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then

                                        If estadoDoc = "AUTORIZADA" Then
                                            Dim _objtype As String = oGrid.DataTable.GetValue("ObjType", oGrid.GetDataTableRowIndex(pVal.Row))
                                            Dim _DocSubtype As String = oGrid.DataTable.GetValue("DocSubType", oGrid.GetDataTableRowIndex(pVal.Row))
                                            Dim _ss_tipotabla As String = obtenerTipoTabla(_objtype, _DocSubtype)
                                            Dim _DocEntry As Integer = CInt(oGrid.DataTable.GetValue("DocEntry", oGrid.GetDataTableRowIndex(pVal.Row)))
                                            If TipoArchivo = "XML" Then
                                                oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAccesoDoc, _DocEntry, _ss_tipotabla, "xml")
                                            End If
                                            If TipoArchivo = "PDF" Then
                                                oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAccesoDoc, _DocEntry, _ss_tipotabla, "pdf")
                                            End If

                                        Else
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Solo se puede consultar Documentos Autorizados..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        End If
                                    Else
                                        If estadoDoc = "AUTORIZADA" Then

                                            If TipoArchivo = "XML" Then
                                                oManejoDocumentos.ConsultaXML(ClaveAccesoDoc)
                                            End If
                                            If TipoArchivo = "PDF" Then
                                                oManejoDocumentos.ConsultaPDF(ClaveAccesoDoc)
                                            End If

                                        Else
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Solo se puede consultar Documentos Autorizados..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        End If
                                    End If


                                    'oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
                                    'oForm.Freeze(True)
                                    'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    'Dim oDataTable As SAPbouiCOM.DataTable = oGrid.DataTable
                                    ''oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                                    'Dim dtrow As Integer = oGrid.GetDataTableRowIndex(pVal.Row)
                                    'Dim dsd As String = oDataTable.GetValue(16, dtrow)

                                    'Dim pruebas As String = oGrid.DataTable.GetValue(pVal.ColUID, dtrow).ToString
                                    'Dim comboDesc As String = (CType(oGrid.Columns.Item(16), SAPbouiCOM.ComboBoxColumn)).GetSelectedValue(pVal.Row).Description

                                    'oGrid.DataTable.SetValue(16, dtrow, "XML")
                                    'oForm.Freeze(False)
                                    'Dim ofilaDE As Integer = 0

                                    'oForm = rsboApp.Forms.Item("frmDocumentosEnviados")

                                    'ofilaDE = pVal.Row

                                    'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    'oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                                    'oGrid.Rows.SelectedRows.Add(ofilaDE)
                                    'For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                    '    ofilaDE = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                                    '    ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                                    'Next

                                    'Dim oCol As SAPbouiCOM.ComboBoxColumn
                                    'oCol = oGrid.Columns.Item(16)

                                    'Dim sTipoDoc As String = oCol.SetSelectedValue

                                    'Dim dvSelectedValue As String = oCol.GetSelectedValue(pVal.Row).Description
                                    'MsgBox(dvSelectedValue)


                                    'Dim oComboBox As SAPbouiCOM.ComboBox = Nothing
                                    'oComboBox = oForm.Items.Item(16).Specific
                                    'SelectedValue = oComboBox.Value.Trim()
                                    'MsgBox(SelectedValue)

                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        If oCFLEvento.BeforeAction = False Then
                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosEnviados")

                            oCFLDE = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty

                            If Not oDataTable Is Nothing Then
                                val = oDataTable.GetValue(0, 0)
                                val1 = oDataTable.GetValue(1, 0)
                                Try

                                    oUserDataSourceDE = oForm.DataSources.UserDataSources.Item("EditDS")
                                    oUserDataSourceDE.ValueEx = val

                                Catch ex As Exception
                                End Try

                                Try
                                    Dim txtRaz As SAPbouiCOM.EditText
                                    txtRaz = oForm.Items.Item("Item_12").Specific
                                    txtRaz.Value = val1
                                Catch ex As Exception
                                End Try
                            Else
                                Dim txtRaz As SAPbouiCOM.EditText
                                txtRaz = oForm.Items.Item("Item_12").Specific
                                txtRaz.Value = ""
                            End If

                        Else
                            'BubbleEvent = False
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub
    Public Sub explorador()
        Try
            rsboApp.SetStatusBarMessage(NombreAddon + " - Minimizar la pantalla para que pueda escoger la ruta..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            GetfileThreadFE = New Threading.Thread(AddressOf GetNombreArchivoPartFE, 1)
            GetfileThreadFE.SetApartmentState(Threading.ApartmentState.STA)
            GetfileThreadFE.Start()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub GetNombreArchivoPartFE()

        Try
            Dim folderBrowserDialog1 As FolderBrowserDialog = New FolderBrowserDialog

            folderBrowserDialog1.Description = "Selecccionar ruta para guardar los archivos seleccionados"
            folderBrowserDialog1.ShowNewFolderButton = True

            'Dim f2 As New System.Windows.Forms.Form() With {.TopMost = True, .Visible = False}
            'Dim sv = New FolderBrowserDialog

            'Dim resul = f2.ShowDialog(sv)

            Dim result As DialogResult = folderBrowserDialog1.ShowDialog(New Form() With {.TopMost = True, .Visible = False})
            If result = DialogResult.OK Then
                Dim s As String
                s = folderBrowserDialog1.SelectedPath

                folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyComputer

                oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
                Dim _ruta2 As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific
                _ruta2.Value = s



            End If




            'guardar
            'Dim saveFileDialog1 As New SaveFileDialog
            'Dim myStream As Stream
            'Dim path As String

            'saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            'saveFileDialog1.FilterIndex = 2
            'saveFileDialog1.RestoreDirectory = True

            'If saveFileDialog1.ShowDialog() = DialogResult.OK Then

            '    myStream = saveFileDialog1.OpenFile()
            '    myStream.Close()

            'End If
            'fin guardar
        Catch ex As Exception
            rsboApp.StatusBar.SetText(" Error al cargar archivo, " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub ConsultarEstados()

        Try
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
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
            Dim Marca As String = "N"
            For x = 0 To oDatable.Rows.Count - 1
                nombre_estado = oDatable.GetValue("EstadoDoc", x)
                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                    If nombre_estado = Estados_docenviados.EN_PROCESO Then
                        LQE = oDatable.GetValue("NomDocSap", x)
                        ss_tipotabla = obtenerTipoTabla(oDatable.GetValue("ObjType", x), oDatable.GetValue("DocSubType", x))
                        identificador = CInt(oDatable.GetValue("DocEntry", x))
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
                            oGrid.Rows.SelectedRows.Add(y)
                            oManejoDocumentosEcua.ProcesaEnvioDocumento(identificador, ss_tipotabla, True)
                            '   rsboApp.MessageBox(ss_tipotabla & " " & oDatable.GetValue("DocEntry", x))
                            gcss.SetRowBackColor(y + 1, 255000)
                        Catch ex As Exception
                            gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                        End Try


                    End If
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Else
                    If nombre_estado = Estados_docenviados.EN_PROCESO_SRI Or nombre_estado = Estados_docenviados.ERROR_EN_RECEPCION Then
                        LQE = oDatable.GetValue("NomDocSap", x)
                        ss_tipotabla = obtenerTipoTabla(oDatable.GetValue("ObjType", x), oDatable.GetValue("DocSubType", x))
                        identificador = CInt(oDatable.GetValue("DocEntry", x))
                        For y = 1 To oGrid.Rows.Count
                            indexgrid = oGrid.GetDataTableRowIndex(y)
                            If indexgrid = x Then
                                gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                                Exit For
                            End If
                        Next

                        Try
                            oGrid.Rows.SelectedRows.Add(y)
                            oManejoDocumentos.ProcesaEnvioDocumento(identificador, ss_tipotabla, True)
                            gcss.SetRowBackColor(y + 1, 255000)
                            'oGrid.Rows.SelectedRows.Add(y)
                        Catch ex As Exception
                            gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                        End Try

                        '' 26/01/2023 RO - ENVIAR LOS NO ENVIADOS - FOLEADOS PREVIAMENTE EN LA CREACION
                        'ElseIf nombre_estado = Estados_docenviados.NO_ENVIADO Then
                        '' 03/02/2023 DM - ENVIAR LOS NO ENVIADOS - FOLEADOS PREVIAMENTE EN LA CREACION Unicamente para clientes SOLSAP y que este activo parametro reenviar documentos
                    ElseIf (nombre_estado = Estados_docenviados.NO_ENVIADO Or nombre_estado = Estados_docenviados.NO_AUTORIZADA Or nombre_estado = Estados_docenviados.DEVUELTA) And Functions.VariablesGlobales._ReenviarDocsPantala = "Y" And ListDocEntrys.Count = 0 Then 'And Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP
                        LQE = oDatable.GetValue("NomDocSap", x)
                        ss_tipotabla = obtenerTipoTabla(oDatable.GetValue("ObjType", x), oDatable.GetValue("DocSubType", x))
                        identificador = CInt(oDatable.GetValue("DocEntry", x))

                        Marca = oDatable.GetValue("Seleccionar", x)
                        If Marca = "Y" Then
                            For y = 1 To oGrid.Rows.Count
                                indexgrid = oGrid.GetDataTableRowIndex(y)
                                If indexgrid = x Then
                                    gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                                    Exit For
                                End If
                            Next

                            Try
                                'oGrid.SetCellFocus(x, 10)
                                oGrid.Rows.SelectedRows.Add(y)
                                oManejoDocumentos.ProcesaEnvioDocumento(identificador, ss_tipotabla, False)
                                gcss.SetRowBackColor(y + 1, 255000)
                            Catch ex As Exception
                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                            End Try
                        End If

                    ElseIf Functions.VariablesGlobales._ReenviarListaDocEnv = "Y" And ListDocEntrys.Count > 0 Then
                        ReenviarListaDocs(ListDocEntrys)
                        ListDocEntrys.Clear()
                        rsboApp.SetStatusBarMessage("Lista reenviada con exito", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        rsboApp.SetStatusBarMessage("Lista vacia", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        Dim txtEst As SAPbouiCOM.EditText = oForm.Items.Item("txtEst").Specific
                        Dim txtPemi As SAPbouiCOM.EditText = oForm.Items.Item("txtPemi").Specific
                        Dim txtSec As SAPbouiCOM.EditText = oForm.Items.Item("txtSec").Specific
                        Dim txtSN As SAPbouiCOM.EditText = oForm.Items.Item("txtRSN").Specific

                        txtEst.Value = ""
                        txtPemi.Value = ""
                        txtSec.Value = ""
                        txtSN.Value = ""

                    End If
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                End If

                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next

            rsboApp.StatusBar.SetText("(SAED) El estado de los documentos han sido actualizados correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
            FormularioDocumentosEnviadosCargarGrid()

        Catch ex As Exception
            ListDocEntrys.Clear()
            rsboApp.SetStatusBarMessage("Ocurrio un error al llamar la funcion ConsultarEstados " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    Private Sub ExpandirContraer()

        Dim oGrid As SAPbouiCOM.Grid

        Try
            oGrid = oForm.Items.Add("oGrid", SAPbouiCOM.BoFormItemTypes.it_GRID).Specific
        Catch ex As Exception
            oGrid = oForm.Items.Item("oGrid").Specific
        End Try

        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

        If schk.Checked Then

            oGrid.Rows.CollapseAll()
        Else
            oGrid.Rows.ExpandAll()

        End If

        oGrid.AutoResizeColumns()

    End Sub

    Private Sub Event_MatrixLinkPressed(ByVal pVal As SAPbouiCOM.ItemEvent)

        If pVal.FormTypeEx = "frmDocumentosEnviados" Then

            Select Case pVal.ItemUID

                Case "oGrid"

                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                    Dim oObjType As String = oGrid.DataTable.GetValue("ObjType", oGrid.GetDataTableRowIndex(pVal.Row))
                    Dim oColumns As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("DocEntry")
                    Dim DocNum As String = oGrid.DataTable.GetValue("DocNum", oGrid.GetDataTableRowIndex(pVal.Row))
                    Select Case oObjType

                        Case 13
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oInvoices

                        Case 203
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oDownPayments

                        Case 14
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oCreditNotes

                        Case 18
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices

                        Case 204
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments

                        Case 15
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oDeliveryNotes

                        Case 60
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit

                        Case 67
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer

                        Case 1250000001
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest

                        Case "SSGRNEW"

                            'rsboApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "SSGRNEW", DOCENTRY)
                            ' SAPbouiCOM
                            oColumns.LinkedObjectType = "SSGRNEW"
                            'oForm.Freeze(True)
                            ofrmGuiasRemision.CargaFormularioGuia()
                            Dim oFormG = rsboApp.Forms.Item("frmGuiasRemision")
                            oFormG.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            Dim txtDocNum As SAPbouiCOM.EditText
                            txtDocNum = oFormG.Items.Item("txtentry").Specific
                            txtDocNum.Item.Enabled = True
                            txtDocNum.Value = DocNum
                            oFormG.Items.Item("1").Click()
                            txtDocNum.Item.Enabled = False
                            'oForm.Freeze(False)

                        Case Else
                            Exit Sub

                    End Select

            End Select

        End If
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

            'For Each oForm In oApp.Forms
            '    If oForm.UniqueID = Formulario Then
            '        oForm.Visible = True
            '        oForm.Select()
            '        ' oForm.Close()
            '        Return True
            '    End If
            'Next


            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function obtenerTipoTabla(ByVal objType As String, ByVal subtype As String) As String

        Dim oTipoTabla As String = ""

        If objType = "13" Then  ' FACTURA DE CLIENTE or  NOTA DE DEBITO
            If subtype = "--" Then
                oTipoTabla = "FCE"
            Else
                oTipoTabla = "NDE"
            End If
        ElseIf objType = "203" Then
            oTipoTabla = "FAE"
        ElseIf objType = "14" Then 'NOTA DE CREDITO DE CLIENTES
            oTipoTabla = "NCE"
        ElseIf objType = "15" Then 'GUIA DE REMISION - ENTREGA
            oTipoTabla = "GRE"
        ElseIf objType = "67" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
            oTipoTabla = "TRE"
        ElseIf objType = "1250000001" Then 'GUIA DE REMISION - SOLICITUD TRASLADOS                                            
            oTipoTabla = "TLE"
        ElseIf objType = "18" Then  'FACTURA DE PROVEEDOR/RETENCION
            If LQE = "LQ DE COMPRA" Then
                oTipoTabla = "LQE"
            Else
                oTipoTabla = "REE"
            End If

        ElseIf objType = "204" Then  'FACTURA DE ANTICIPO PROVEEDOR/RETENCION                             

            oTipoTabla = "REA"
        ElseIf objType = "SSGRNEW" Then
            oTipoTabla = "SSGR"
        End If


        Return oTipoTabla
    End Function

    Private Structure Estados_docenviados
        Const EN_PROCESO_SRI = "EN PROCESO SRI"
        Const ERROR_EN_RECEPCION = "ERROR EN RECEPCION"
        Const AUTORIZADA = "AUTORIZADA"
        Const EN_PROCESO = "EN PROCESO"
        Const NO_ENVIADO = "NO ENVIADO"
        Const NO_AUTORIZADA = "NO AUTORIZADA"
        Const VALIDAR_DATOS = "VALIDAR DATOS"
        Const DEVUELTA = "DEVUELTA"
    End Structure

    Private Function SeleccionarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                estado = oGridDet.GetValue("EstadoDoc", i)
                '  If estado = "AUTORIZADA" Then
                oGridDet.SetValue("Seleccionar", i, "Y")
                    contador += 1
                    resul = True
                ' End If
            Next
            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmDocumentosEnviados")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al SeleccionarDocumentosPendientes:" + ex.Message().ToString(), "frmDocumentosEnviados")
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function

    Private Function DesmarcarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                estado = oGridDet.GetValue("EstadoDoc", i)
                '  If estado = "AUTORIZADA" Then
                oGridDet.SetValue("Seleccionar", i, "N")
                    contador += 1
                    resul = True
                '  End If
            Next
            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmDocumentosEnviados")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al SeleccionarDocumentosPendientes:" + ex.Message().ToString(), "frmDocumentosEnviados")
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function

    Private Function ReenviarCorreosDocMarcados(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Dim marcado As String = ""
        Dim SQUERY As String = ""
        Dim Tabla As String = ""
        Dim objType As String = ""
        Dim ClaveAcceso As String = ""
        Dim docentry As String = ""
        Try
            oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                Utilitario.Util_Log.Escribir_Log("ingresando al for", "ManejoDeDocumentos")
                If oGridDet.GetValue("Seleccionar", i) = "Y" Then
                    Utilitario.Util_Log.Escribir_Log("ingresando al if", "ManejoDeDocumentos")
                    rsboApp.SetStatusBarMessage(NombreAddon + " - Re Enviando Mail... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    ClaveAcceso = oGridDet.GetValue("ClaveAcceso", i)
                    docentry = oGridDet.GetValue("DocEntry", i)
                    objType = oGridDet.GetValue("ObjType", i)

                    If objType = "13" Then  ' FACTURA DE CLIENTE or  NOTA DE DEBITO
                        Tabla = "OINV"
                    ElseIf objType = "203" Then
                        Tabla = "OINV"
                    ElseIf objType = "14" Then 'NOTA DE CREDITO DE CLIENTES
                        Tabla = "ORIN"
                    ElseIf objType = "15" Then 'GUIA DE REMISION - ENTREGA
                        Tabla = "ODLN"
                    ElseIf objType = "67" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
                        Tabla = "OWTR"
                    ElseIf objType = "18" Then  'FACTURA DE PROVEEDOR/RETENCION
                        Tabla = "OPCH"
                    ElseIf objType = "204" Then  'FACTURA DE ANTICIPO PROVEEDOR/RETENCION                             
                        Tabla = "OPCH"
                    End If
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", Tabla), "IDENTIFICADOR", docentry)
                        'SQUERY = String.Format("SELECT A.""E_Mail"" FROM ""{0}"" O INNER JOIN ""OCRD"" A ON O.""CardCode"" = A.""CardCode"" WHERE O.""DocEntry"" =  {1}", oTabla, docentry)
                    Else
                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", Tabla), "IDENTIFICADOR", docentry)
                        'SQUERY = String.Format("SELECT A.E_Mail FROM {0} O WITH(NOLOCK) INNER JOIN OCRD A WITH(NOLOCK) ON O.CardCode = A.CardCode WHERE O.DocEntry = {1} ", oTabla, docentry)
                    End If
                    Dim sCorreoNuevo As String = oFuncionesB1.getRSvalue(SQUERY, "Email", "")

                    If String.IsNullOrEmpty(sCorreoNuevo) Then
                        'rsboApp.SetStatusBarMessage(NombreAddon + " - No se encontro consulta, verificar en la parametrización..!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        mensajeRC = NombreAddon + " - No se encontro consulta, verificar en la parametrización..!! "
                    Else
                        Utilitario.Util_Log.Escribir_Log("Consulta: " + SQUERY.ToString() + " - Respuesta: " + sCorreoNuevo.ToString, "frmDocumentosEnviados")
                        Try
                            If oManejoDocumentos.ReenvioMail(sCorreoNuevo, ClaveAcceso) Then
                                rsboApp.SetStatusBarMessage(NombreAddon + " - Mail Re Enviado, Listo!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                contador += 1
                                resul = True
                            End If

                        Catch ex As Exception
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Error al al reenviar el Mail.!!: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            Utilitario.Util_Log.Escribir_Log("error al reenviar el e-mail: " + ex.Message.ToString, "ManejoDeDocumentos")
                            mensajeRC = "error al reenviar el e-mail: " + ex.Message.ToString
                            resul = False
                            Return resul
                        End Try

                    End If
                End If
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            Next
            Utilitario.Util_Log.Escribir_Log("Cantidad de correos reenviados : " + contador.ToString(), "frmDocumentosEnviados")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al reenviar correos:" + ex.Message().ToString(), "frmDocumentosEnviados")
            mensajeRC = "Error al reenviar correos:" + ex.Message().ToString()
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function

    Private Function ExportarArchivo(Extension As String) As Boolean

        'Dim resul As Boolean = False
        'Dim marcado As String = ""
        'Dim SQUERY As String = ""
        'Dim Tabla As String = ""
        'Dim objType As String = ""
        Dim ClaveAccesoExp As String = ""
        Dim NumeracionSRI As String = ""
        Dim NombreSN As String = ""
        Dim estadoExp As String = ""
        Try
            oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
            oForm.Freeze(True)
            Dim oGridDetExp As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDetExp.Rows.Count - 1
                Utilitario.Util_Log.Escribir_Log("ingresando al for exportar documentos", "frmDocumentosEnviados")
                estadoExp = oGridDetExp.GetValue("EstadoDoc", i)
                DocSRI = oGridDetExp.GetValue("SRI", i)
                NumDocEmi = oGridDetExp.GetValue("Folio", i)
                If estadoExp = Estados_docenviados.AUTORIZADA Then
                    If oGridDetExp.GetValue("Seleccionar", i) = "Y" Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - Exportando documentos por favor espere un momento... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        ClaveAccesoExp = oGridDetExp.GetValue("ClaveAcceso", i)
                        NumeracionSRI = oGridDetExp.GetValue("Folio", i)
                        NombreSN = oGridDetExp.GetValue("CardName", i)
                        LQE = oGridDetExp.GetValue("NomDocSap", i)
                        Dim TipoTabla = obtenerTipoTabla(oGridDetExp.GetValue("ObjType", i), oGridDetExp.GetValue("DocSubType", i))

                        Try
                            Exportar_PDF_XML(ClaveAccesoExp, Extension, NumeracionSRI, NombreSN, TipoTabla)
                            rsboApp.SetStatusBarMessage(NombreAddon + " - PDF con clave " + ClaveAccesoExp + " guardado exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        Catch ex As Exception
                            rsboApp.SetStatusBarMessage(NombreAddon + " - Error al guardar pdf: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            Utilitario.Util_Log.Escribir_Log("error al guardar pdf " + ex.Message.ToString, "frmDocumentosEnviados")

                        End Try
                        contExpDoc += 1
                    End If
                End If

                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            Next
            Return True
            'Utilitario.Util_Log.Escribir_Log("Cantidad de correos reenviados : " + contador.ToString(), "frmDocumentosEnviados")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al guardarpdf:" + ex.Message().ToString(), "frmDocumentosEnviados")
            Return False
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try


    End Function
    Public Sub Exportar_PDF_XML(ClaveAccesoExp As String, ext As String, numeracionsri As String, Socionegocio As String, TipoDoc As String)
        Try
            Dim TipoWebServicesExp As String = "LOCAL"
            Dim mensajeExp As String = ""
            TipoWebServicesExp = Functions.VariablesGlobales._TipoWS
            Dim url As String = ""
            url = Functions.VariablesGlobales._wsConsultaEmision
            If url = "" Then
                If Not Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                    rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

                    Exit Sub
                End If

            End If

            'Utilitario.Util_Log.Escribir_Log("VisualizaPDF_Bytes :  " + ClaveAccesoExp, "frmDocumentosEnviados")

            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmDocumentosEnviados")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            ' Dim rutaExp As String = "C:\SAED"
            Dim carpeta As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific
            Dim rutaExp As String = carpeta.Value.ToString
            Dim URLExp As String = ""
            Dim ws As Object

            Utilitario.Util_Log.Escribir_Log("VER PDF WS : " + TipoWebServicesExp, "frmDocumentosEnviados")

            If TipoWebServicesExp = "LOCAL" Then
                ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
            ElseIf TipoWebServicesExp = "NUBE" Then
                ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                'ElseIf TipoWebServices = "NUBE_4_1" Then
                '    ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA
            ElseIf TipoWebServicesExp = "NUBE_4_1" Then
                ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

            End If

            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = Functions.VariablesGlobales._vgProxy_puerto
                Proxy_IP = Functions.VariablesGlobales._vgProxy_IP
                Proxy_Usuario = Functions.VariablesGlobales._vgProxy_Usuario
                Proxy_Clave = Functions.VariablesGlobales._vgProxy_Clave

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmDocumentosEnviados")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmDocumentosEnviados")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmDocumentosEnviados")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmDocumentosEnviados")

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

            Dim VisualizaPDF_Bytes As String = "N"
            Dim FSExp As FileStream = Nothing
            VisualizaPDF_Bytes = Functions.VariablesGlobales._VisualizaPDFByte

            'If ext = "PDF" Then
            '    rutaExp += "\PDF"
            'Else
            '    rutaExp += "\XML"
            'End If

            If Not Directory.Exists(rutaExp) Then
                Directory.CreateDirectory(rutaExp)
                Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + rutaExp & "\" & "FACTURAS".ToString, "frmDocumentosEnviados")
            End If

            'If ext = "PDF" Then
            '    rutaExp += "\" + rCompany.CompanyName + " - " + Socionegocio + " - " + numeracionsri + ".pdf"
            'Else
            '    rutaExp += "\" + rCompany.CompanyName + " - " + Socionegocio + " - " + numeracionsri + ".xml"
            'End If
            If ext = "PDF" Then
                rutaExp += "\" + TipoDoc + " - " + numeracionsri + ".pdf"
            Else
                rutaExp += "\" + TipoDoc + " - " + numeracionsri + ".xml"
            End If

            If File.Exists(rutaExp) Then
                File.Delete(rutaExp)
                Utilitario.Util_Log.Escribir_Log("archivo eliminado" + ext + " " + rutaExp.ToString, "frmDocumentosEnviados")
            End If

            Utilitario.Util_Log.Escribir_Log("ruta guardar " + ext + " " + rutaExp.ToString, "frmDocumentosEnviados")

            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then

                If Not File.Exists(rutaExp) Then

                    Dim Sincro_ruc As String = "", Sincro_Tipo_doc As String = "", Sincro_sec_ERP As String = "", Sincro_Num_Doc As String

                    Sincro_ruc = oFuncionesB1.getRSvalue("select ""TaxIdNum"" from OADM", "TaxIdNum", "")

                    Sincro_Num_Doc = NumDocEmi

                    Sincro_Tipo_doc = DocSRI

                    Dim respuesta_WS As String = ""
                    Dim ObjetoRespuesta As New Entidades.ConsultaDocRespuesta

                    Dim ConsultarEstadoDoc As New Entidades.ConsultaDocumento
                    ConsultarEstadoDoc.NombreWs = Functions.VariablesGlobales._NombreWsEcua
                    ConsultarEstadoDoc.clave = Functions.VariablesGlobales._Token
                    ConsultarEstadoDoc.ruc = Sincro_ruc
                    ConsultarEstadoDoc.docType = Sincro_Tipo_doc
                    ConsultarEstadoDoc.docNumber = Sincro_Num_Doc

                    ObjetoRespuesta = CoreRest.ConsultaDocumento(ConsultarEstadoDoc, respuesta_WS)

                    If Not ObjetoRespuesta Is Nothing Then

                        Dim Archivobyte As Byte() = Nothing

                        Dim _nombreFile = ObjetoRespuesta.authorizationNumber.ToString

                        Dim _path = carpeta.Value.ToString
                        If ext = "PDF" Then
                            Archivobyte = Convert.FromBase64String(ObjetoRespuesta.pdf)
                            _path = _path & _nombreFile & ".pdf"

                        Else
                            Archivobyte = Convert.FromBase64String(ObjetoRespuesta.xml)
                            _path = _path & _nombreFile & ".xml"
                        End If

                        FSExp = New FileStream(rutaExp, System.IO.FileMode.Create)
                        FSExp.Write(Archivobyte, 0, Archivobyte.Length)
                        FSExp.Close()

                        'System.IO.File.WriteAllBytes(_path, Archivobyte)
                    End If

                End If

            Else
                If Not File.Exists(rutaExp) Then

                    SetProtocolosdeSeguridad()
                    Dim dbbyteExp As Byte() = Nothing
                    mensajeExp = ""
                    If TipoWebServicesExp = "LOCAL" Then
                        dbbyteExp = ws.ConsultarDocumento(ClaveAccesoExp, "PDF")
                    ElseIf TipoWebServicesExp = "NUBE" Then
                        dbbyteExp = ws.ConsultarDocumento(ClaveAccesoExp, "PDF")
                    ElseIf TipoWebServicesExp = "NUBE_4_1" Then
                        If ext = "PDF" Then
                            dbbyteExp = ws.ConsultarDocumento(ClaveAccesoExp, "PDF", mensajeExp)
                        Else
                            dbbyteExp = ws.ConsultarDocumento(ClaveAccesoExp, "XML", mensajeExp)
                        End If


                    End If
                    If dbbyteExp Is Nothing Then
                        rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensajeExp.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    Else
                        FSExp = New FileStream(rutaExp, System.IO.FileMode.Create)
                        FSExp.Write(dbbyteExp, 0, dbbyteExp.Length)
                        FSExp.Close()
                    End If
                Else
                    Dim dbbyteExp As Byte() = File.ReadAllBytes(rutaExp)
                    FSExp = New FileStream(rutaExp, System.IO.FileMode.Create)
                    FSExp.Write(dbbyteExp, 0, dbbyteExp.Length)
                    FSExp.Close()
                End If
            End If


            'rsboApp.SetStatusBarMessage("PDF Abierto! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
#Disable Warning BC42353 ' La función 'ConsultaPDF' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    End Sub

    Public Sub dialogo()

        'Dim FolderBrowserDialog1 As FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog

        'With FolderBrowserDialog1
        '    .RootFolder = Environment.SpecialFolder.CommonProgramFiles
        '    '.SelectedPath = "C:\Temp"
        '    .ShowNewFolderButton = True
        '    .Description = "Escoger la carpeta donde se almacenaran los documentos"
        'End With

        'If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
        '    Console.WriteLine(FolderBrowserDialog1.SelectedPath)
        'End If

    End Sub

    Public Function GetComboBoxSelectedValue(ByVal oForm As SAPbouiCOM.Form, ByVal oComUID As String) As Object
        '*****************************************
        'DECLARE LOCAL VARIABLE(S)
        '*****************************************
        Dim oComboBox As SAPbouiCOM.ComboBox = Nothing
        '*****************************************
        oComboBox = oForm.Items.Item(oComUID).Specific
        Try
            Return oComboBox.Value.Trim()
        Catch ex As Exception
            Return Nothing
        Finally

        End Try
    End Function

    Private Sub AgregarDocLista()

        Try
            oForm = rsboApp.Forms.Item("frmDocumentosEnviados")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            Dim Seleccion As String = ""
            Dim DocEntry As String = 0
            Dim tipotabla As String = ""


            For i As Integer = 0 To oGridDet.Rows.Count - 1
                Seleccion = oGridDet.GetValue("Seleccionar", i)

                If Seleccion = "Y" Then
                    LQE = oGridDet.GetValue("NomDocSap", i)
                    DocEntry = oGridDet.GetValue("DocEntry", i)
                    tipotabla = obtenerTipoTabla(oGridDet.GetValue("ObjType", i), oGridDet.GetValue("DocSubType", i))
                    Dim concatenado = DocEntry + " ; " + tipotabla
                    Dim numeracionSRI = oGridDet.GetValue("Folio", i)
                    Dim estado = oGridDet.GetValue("EstadoDoc", i)
                    Dim comentario = oGridDet.GetValue("Comentario", i)
                    Dim objtype = oGridDet.GetValue("ObjType", i)
                    Dim ClaveAcceso = oGridDet.GetValue("ClaveAcceso", i)
                    If ListDocEntrys.Count > 0 Then
                        Dim index = ListDocEntrys.FindAll(Function(p) p.Contains(concatenado)).Count
                        If index > 0 Then
                            rsboApp.SetStatusBarMessage("Error: " + "El documento que intenta agregar ya se encuentra en la lista", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        Else
                            ListDocEntrys.Add(DocEntry + " ; " + tipotabla + " ; " + numeracionSRI + " ; " + estado + " ; " + comentario + " ; " + Seleccion + " ; " + objtype + " ; " + ClaveAcceso)
                            rsboApp.SetStatusBarMessage("Se agrego el documento: " + tipotabla + " - " + numeracionSRI.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        End If
                    Else
                        ListDocEntrys.Add(DocEntry + " ; " + tipotabla + " ; " + numeracionSRI + " ; " + estado + " ; " + comentario + " ; " + Seleccion + " ; " + objtype + " ; " + ClaveAcceso)
                        rsboApp.SetStatusBarMessage("Se agrego el documento: " + tipotabla + " - " + numeracionSRI.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End If


                End If

            Next

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al AgregarDocLista:" + ex.Message().ToString(), "frmDocumentosEnviados")

        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub ReenviarListaDocs(Lista As List(Of String))

        Dim ValoresListaDocs = ""
        Dim docentry As Integer = 0
        Dim tipotabla As String = ""

        Try
            For j As Integer = 0 To Lista.Count - 1
                ValoresListaDocs = ListDocEntrys(j)
                docentry = Trim(ValoresListaDocs.Split(";")(0))
                tipotabla = Trim(ValoresListaDocs.Split(";")(1))

                oManejoDocumentos.ProcesaEnvioDocumento(docentry, tipotabla, False)

                ValoresListaDocs = ""
            Next

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ReenviarListaDocs:" + ex.Message().ToString(), "frmDocumentosEnviados")
            ListDocEntrys.Clear()
        Finally
            oForm.Freeze(False)
        End Try

    End Sub


End Class

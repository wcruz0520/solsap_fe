Public Class frmAnexoCompras

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private odt As SAPbouiCOM.DataTable

    Dim odtDE As SAPbouiCOM.DataTable
    Dim oUserDataSourceDE As SAPbouiCOM.UserDataSource
    Dim oCFLsDE As SAPbouiCOM.ChooseFromListCollection
    Dim oConsDE As SAPbouiCOM.Conditions
    Dim oConDE As SAPbouiCOM.Condition
    Dim oCFLDE As SAPbouiCOM.ChooseFromList
    'Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParamsDE As SAPbouiCOM.ChooseFromListCreationParams

    'cenbtro de costo
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oConditions As SAPbouiCOM.Conditions
    Dim oCondition As SAPbouiCOM.Condition

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub


    Public Sub CargaFormularioAnexoCompras()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmAnexoCompras") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmAnexoCompras.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmAnexoCompras").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmAnexoCompras")

            oForm.Freeze(True)

            Dim txtFchIni As SAPbouiCOM.EditText
            txtFchIni = oForm.Items.Item("finicial").Specific
            txtFchIni.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtFchFin As SAPbouiCOM.EditText
            txtFchFin = oForm.Items.Item("ffinal").Specific
            txtFchFin.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim ipLogoSS As SAPbouiCOM.PictureBox
            ipLogoSS = oForm.Items.Item("Item_6").Specific
            ipLogoSS.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"

            '' CHOOSE FROM LIST
            'oCFLsDE = oForm.ChooseFromLists
            'oCFLCreationParamsDE = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            'oCFLCreationParamsDE.MultiSelection = False
            'oCFLCreationParamsDE.ObjectType = "2"
            ''oCFLCreationParams.ObjectType = "Exx_DEPOTRANS"
            'oCFLCreationParamsDE.UniqueID = "CFLP"
            'oCFLDE = oCFLsDE.Add(oCFLCreationParamsDE)
            '' Adding Conditions to CFL1
            'oConsDE = oCFLDE.GetConditions()

            'oConDE = oConsDE.Add()
            'oConDE.Alias = "CardType"
            'oConDE.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oConDE.CondVal = "S"
            'oCFLDE.SetConditions(oConsDE)
            oCFLs = oForm.ChooseFromLists

            ' Crear parámetros para el CFL
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = 2 ' Código para OOCR (centros de coste)
            oCFLCreationParams.UniqueID = "CFLP"

            ' Crear el CFL
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Configurar condiciones del CFL para filtrar por DimCode = 3
            oConditions = oCFL.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "CardType" ' Campo a filtrar
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "S" ' Valor del filtro

            oCFL.SetConditions(oConditions)
            ' END CHOOSE FROM LIST

            Dim txtCodSN As SAPbouiCOM.EditText
            txtCodSN = oForm.Items.Item("txtCodSN").Specific
            oForm.DataSources.UserDataSources.Add("EditDSP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCodSN.DataBind.SetBound(True, "", "EditDSP")
            txtCodSN.ChooseFromListUID = "CFLP"
            txtCodSN.ChooseFromListAlias = "CardCode"

            Dim chkConExp As SAPbouiCOM.CheckBox
            chkConExp = oForm.Items.Item("chkConExp").Specific
            oForm.DataSources.UserDataSources.Add("chkConExp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkConExp.ValOn = "Y"
            chkConExp.ValOff = "N"
            chkConExp.DataBind.SetBound(True, "", "chkConExp")

            Dim lnkP As SAPbouiCOM.LinkedButton
            lnkP = oForm.Items.Item("lnkP").Specific
            lnkP.LinkedObjectType = 2
            lnkP.Item.LinkTo = "txtCodSN"

            oForm.Left = 0
            oForm.Top = 0

            CargarGrid()

            ' CargaDatos()
            Dim screenWidth As Integer = rsboApp.Desktop.Width
            Dim screenHeight As Integer = rsboApp.Desktop.Height

            ' Obtener el tamaño del formulario
            Dim formWidth As Integer = oForm.Width
            Dim formHeight As Integer = oForm.Height

            ' Calcular la posición para centrar el formulario
            Dim centeredLeft As Integer = (screenWidth - formWidth) / 2
            Dim centeredTop As Integer = (screenHeight - formHeight) / 2

            ' Ajustar la posición del formulario
            oForm.Left = centeredLeft
            oForm.Top = centeredTop - 75


            oCFLs = oForm.ChooseFromLists

            ' Crear parámetros para el CFL
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = 62 ' Código para OOCR (centros de coste)
            oCFLCreationParams.UniqueID = "CFL_Dim3"

            ' Crear el CFL
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Configurar condiciones del CFL para filtrar por DimCode = 3
            oConditions = oCFL.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "DimCode" ' Campo a filtrar
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "3" ' Valor del filtro

            oCFL.SetConditions(oConditions)

            Dim txtCC As SAPbouiCOM.EditText
            txtCC = oForm.Items.Item("txtCC").Specific
            oForm.DataSources.UserDataSources.Add("EditDSPCC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCC.DataBind.SetBound(True, "", "EditDSPCC")
            txtCC.ChooseFromListUID = "CFL_Dim3"
            txtCC.ChooseFromListAlias = "OcrCode"

            Dim lnkCC As SAPbouiCOM.LinkedButton
            lnkCC = oForm.Items.Item("lnkCC").Specific
            lnkCC.LinkedObjectType = 62
            lnkCC.Item.LinkTo = "txtCC"

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargarGrid()
        oForm.Freeze(True)


        Dim txtfinicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
        Dim txtffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific
        Dim txtCodCli As SAPbouiCOM.EditText = oForm.Items.Item("txtCodSN").Specific
        Dim txtCC As SAPbouiCOM.EditText = oForm.Items.Item("txtCC").Specific
        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

        If (String.IsNullOrEmpty(txtfinicial.Value) Or String.IsNullOrEmpty(txtffinal.Value)) Then
            rsboApp.SetStatusBarMessage("Debe ingresar un rango de fechas a consultar!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        Dim dfechaDesde As Date
        Dim dfechaHasta As Date
        Dim sQuery As String = ""

        Dim sfolioIni As String = txtfinicial.Value.Trim()
        Dim sfoliofin As String = txtffinal.Value.Trim()

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

        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    'CALL GS_SAP_FE_ONE_OBTENERDOCUMENTOS ('0','2',{d'2016-06-16'},{d'2017-09-28'})
        '    sQuery = "CALL GS_CO_SAP_FE_ObtenerDocumentos ("

        '    sQuery += "'" + sTipoDoc + "'"
        '    sQuery += ",'" + sEstado + "'"
        '    'sQuery += ",''"
        '    'sQuery += ",''"        
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta) + ")"
        'Else
        '    sQuery = "EXEC GS_CO_SAP_FE_ObtenerDocumentos "

        '    sQuery += "'" + sTipoDoc + "'"
        '    sQuery += ",'" + sEstado + "'"
        '    'sQuery += ",''"
        '    'sQuery += ",''"        
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta)
        'End If


        'ESTO ES PARA PROBAR QUERYS ENCRIPTADOS DESDE DB

        sQuery = Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._SS_ComprasQRY.Replace("{", "").Replace("}", "").ToString(), sKey)

        If String.IsNullOrWhiteSpace(sQuery) Then

            rsboApp.SetStatusBarMessage("(GS) Pantalla Inactiva , Por Favor Revisar la Parametrizacion de los Documentos Enviados", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

        End If


        'sQuery = sQuery.Replace("REPPLACE_TIPODOC", "'" + sTipoDoc + "'")
        'sQuery = sQuery.Replace("REPPLACE_ESTADO", "'" + sEstado + "'")


        sQuery = sQuery.Replace("@f1", Functions.FuncionesB1.FechaSql(dfechaDesde))
        sQuery = sQuery.Replace("@f2", Functions.FuncionesB1.FechaSql(dfechaHasta))
        sQuery = sQuery.Replace("@SN", IIf(String.IsNullOrEmpty(txtCodCli.Value), "", txtCodCli.Value.ToString))
        sQuery = sQuery.Replace("@CC", IIf(String.IsNullOrEmpty(txtCC.Value), "", txtCC.Value.ToString))

        'FIN QUERY CONSULTA ENCRYPTADA
        Try
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                Utilitario.Util_Log.Escribir_Log("QUERY: " + sQuery, "frmAnexoCompras")
                oGrid.DataTable.ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmAnexoCompras")
            End Try

            Dim oColumn As SAPbouiCOM.GridColumn
            If oGrid.DataTable.Rows.Count > 0 Then

                For y As Integer = 0 To oGrid.Columns.Count - 1


                    oGrid.Columns.Item(y).Editable = False
                    oColumn = oGrid.Columns.Item(y)
                    Dim _oEditTextColumn As SAPbouiCOM.EditTextColumn = CType(oColumn, SAPbouiCOM.EditTextColumn)
                    Dim cellValue As String = _oEditTextColumn.GetText(0)
                    If IsNumeric(cellValue) Then
                        If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                            Dim oST As SAPbouiCOM.BoColumnSumType = oColumn.ColumnSetting.SumType
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                        End If

                    End If

                Next

            End If

            'oGrid.Columns.Item(0).Description = "Tipo Documento"
            'oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            'oGrid.Columns.Item(0).Editable = False

            'oGrid.Columns.Item(1).Description = "#"
            'oGrid.Columns.Item(1).TitleObject.Caption = "#"
            'oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "DocEntry"
            oGrid.Columns.Item(2).TitleObject.Caption = "DocEntry"


            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oEditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = 18


            'oGrid.Columns.Item(3).Description = "Fecha Emisión"
            'oGrid.Columns.Item(3).TitleObject.Caption = "Fecha Emisión"
            'oGrid.Columns.Item(3).Editable = False

            'oGrid.Columns.Item(4).Description = "Doc. Num."
            'oGrid.Columns.Item(4).TitleObject.Caption = "Doc. Num."
            'oGrid.Columns.Item(4).Editable = False

            'oGrid.Columns.Item(5).Description = "Cliente"
            'oGrid.Columns.Item(5).TitleObject.Caption = "Cliente"
            'oGrid.Columns.Item(5).Editable = False


            'oGrid.Columns.Item(6).Description = "Doc. Total"
            'oGrid.Columns.Item(6).TitleObject.Caption = "Doc. Total"
            'oGrid.Columns.Item(6).Editable = False
            'oGrid.Columns.Item(6).RightJustified = True


            'oGrid.Columns.Item(7).Description = "Estado Documento"
            'oGrid.Columns.Item(7).TitleObject.Caption = "Estado Documento"
            'oGrid.Columns.Item(7).Editable = False

            'oGrid.Columns.Item(8).Description = "CUF"
            'oGrid.Columns.Item(8).TitleObject.Caption = "CUF"
            'oGrid.Columns.Item(8).Editable = False

            'oGrid.Columns.Item(9).Description = "EXT1"
            'oGrid.Columns.Item(9).TitleObject.Caption = "EXT1"
            'oGrid.Columns.Item(9).Editable = False
            'oGrid.Columns.Item(9).Visible = False

            'oGrid.Columns.Item(10).Description = "EXT2"
            'oGrid.Columns.Item(10).TitleObject.Caption = "EXT2"
            'oGrid.Columns.Item(10).Editable = False
            'oGrid.Columns.Item(10).Visible = False

            'oGrid.Columns.Item(11).Description = "EXT3"
            'oGrid.Columns.Item(11).TitleObject.Caption = "EXT3"
            'oGrid.Columns.Item(11).Editable = False
            'oGrid.Columns.Item(11).Visible = False

            'oGrid.Columns.Item(12).Description = "EXT4"
            'oGrid.Columns.Item(12).TitleObject.Caption = "EXT4"
            'oGrid.Columns.Item(12).Editable = False
            'oGrid.Columns.Item(12).Visible = False


            oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()
            schk.Checked = False

            oForm.Freeze(False)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.FormTypeEx = "frmAnexoCompras" Then

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                        Case "btnBuscar"
                            If pVal.BeforeAction = False Then
                                CargarGrid()
                            Else

                            End If

                        Case "chkConExp"
                            If pVal.BeforeAction = False Then
                                ExpandirContraer()
                            End If

                        Case "btnExcel"
                            If pVal.BeforeAction = False Then
                                Dim oGrid As SAPbouiCOM.Grid
                                oGrid = oForm.Items.Item("oGrid").Specific
                                oFuncionesAddon.ExportGridToExcel(oGrid)
                            End If

                        Case Else

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                    If pVal.BeforeAction Then

                        Event_MatrixLinkPressed(pVal)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    If oCFLEvento.BeforeAction = False Then
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmAnexoCompras")

                        If sCFL_ID = "CFLP" Then
                            oCFLDE = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty

                            If Not oDataTable Is Nothing Then
                                val = oDataTable.GetValue(0, 0)
                                val1 = oDataTable.GetValue(1, 0)
                                Try

                                    oUserDataSourceDE = oForm.DataSources.UserDataSources.Item("EditDSP")
                                    oUserDataSourceDE.ValueEx = val

                                Catch ex As Exception
                                End Try

                                Try
                                    Dim lblNomCli As SAPbouiCOM.StaticText
                                    lblNomCli = oForm.Items.Item("Item_5").Specific
                                    lblNomCli.Caption = val1
                                Catch ex As Exception
                                End Try
                            Else
                                Dim lblNomCli As SAPbouiCOM.StaticText
                                lblNomCli = oForm.Items.Item("Item_5").Specific
                                lblNomCli.Caption = ""
                            End If
                        End If

                        If sCFL_ID = "CFL_Dim3" Then
                            oCFLDE = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty

                            If Not oDataTable Is Nothing Then
                                val = oDataTable.GetValue(0, 0)
                                val1 = oDataTable.GetValue(1, 0)
                                Try

                                    oUserDataSourceDE = oForm.DataSources.UserDataSources.Item("EditDSPCC")
                                    oUserDataSourceDE.ValueEx = val

                                Catch ex As Exception
                                End Try

                                Try
                                    Dim lblNomCli As SAPbouiCOM.StaticText
                                    lblNomCli = oForm.Items.Item("lblCC").Specific
                                    lblNomCli.Caption = val1
                                Catch ex As Exception
                                End Try
                            Else
                                Dim lblNomCli As SAPbouiCOM.StaticText
                                lblNomCli = oForm.Items.Item("lblCC").Specific
                                lblNomCli.Caption = ""
                            End If
                        End If


                    Else
                        'BubbleEvent = False
                    End If

                Case Else

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


    Private Sub ExpandirContraer()

        Dim oGrid As SAPbouiCOM.Grid

        Try
            oGrid = oForm.Items.Add("oGrid", SAPbouiCOM.BoFormItemTypes.it_GRID).Specific
        Catch ex As Exception
            oGrid = oForm.Items.Item("oGrid").Specific
        End Try

        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("chkConExp").Specific

        If schk.Checked Then

            oGrid.Rows.CollapseAll()
        Else
            oGrid.Rows.ExpandAll()

        End If

        oGrid.AutoResizeColumns()

    End Sub

    Private Sub Event_MatrixLinkPressed(ByVal pVal As SAPbouiCOM.ItemEvent)

        If pVal.FormTypeEx = "frmAnexoVentas" Then

            Select Case pVal.ItemUID

                Case "oGrid"

                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                    Dim oObjType As String = oGrid.DataTable.GetValue("ObjType", oGrid.GetDataTableRowIndex(pVal.Row))
                    Dim oColumns As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("DocEntry")

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

                        Case 67
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer

                        Case 19
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes

                        Case Else
                            Exit Sub

                    End Select

            End Select

        End If
    End Sub

End Class

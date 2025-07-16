Imports System.IO

Public Class frmGuiasRemision
    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private comboSeries As SAPbouiCOM.ComboBox

    Private num As Integer

    Private oTabla As String
    Private oTipoTabla As String

    Private nombreFormulario As String = "GUIA REMISION"

    Private EsElectronico As String

    Dim lbComentario As SAPbouiCOM.StaticText
    Dim lbEstado As SAPbouiCOM.StaticText
    Dim cbEstado As SAPbouiCOM.ComboBox
    Dim btnAccion As SAPbouiCOM.ButtonCombo


    Dim btnPruebaCombo As SAPbouiCOM.ButtonCombo
    Dim btnPrueba As SAPbouiCOM.Button

    Dim txtcliente As SAPbouiCOM.EditText
    Dim txtnombreCliente As SAPbouiCOM.EditText

    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oUserDataSourceDE As SAPbouiCOM.UserDataSource
    Dim oCFLDE As SAPbouiCOM.ChooseFromList

    Dim truco As Boolean = False


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub


    Public Sub CargaFormularioGuia()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmGuiasRemision") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmGuiasRemision.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmGuiasRemision").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmGuiasRemision")

            Dim cmbTipo As SAPbouiCOM.ComboBox
            cmbTipo = oForm.Items.Item("cmbTSN").Specific
            cmbTipo.ValidValues.Add("01", "Cliente")
            cmbTipo.ValidValues.Add("02", "Proveedor")
            cmbTipo.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue)

            Inicioform(oForm)

            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            'Dim mtx As SAPbouiCOM.Matrix = oForm.Items.Item("MtxCont").Specific
            'mtx.Item.Enabled = True
            ' Configurar las propiedades del formulario
            Dim logo As SAPbouiCOM.PictureBox
            logo = oForm.Items.Item("logo").Specific
            logo.Picture = Application.StartupPath & "\LogoSS.png"

            '' CHOOSE FROM LIST
            'oCFLs = oForm.ChooseFromLists
            'oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            'oCFLCreationParams.MultiSelection = False
            'oCFLCreationParams.ObjectType = "2"
            ''oCFLCreationParams.ObjectType = "Exx_DEPOTRANS"
            'oCFLCreationParams.UniqueID = "CFL1"
            'oCFL = oCFLs.Add(oCFLCreationParams)
            '' Adding Conditions to CFL1
            'oCons = oCFL.GetConditions()

            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)
            '' END CHOOSE FROM LIST

            'Dim txtcodCli As SAPbouiCOM.EditText
            'txtcodCli = oForm.Items.Item("txtcodCli").Specific
            'oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'txtcodCli.DataBind.SetBound(True, "", "EditDS")
            'txtcodCli.ChooseFromListUID = "CFL1"
            'txtcodCli.ChooseFromListAlias = "CardType"

            oForm.Left = 0
            oForm.Top = 0

            ' Obtener el tamaño de la pantalla
            Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
            Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
            Dim screenLeft As Integer = Screen.PrimaryScreen.Bounds.Left
            Dim screenTop As Integer = Screen.PrimaryScreen.Bounds.Top
            ' Ajustar el tamaño del formulario en base al tamaño de la pantalla
            'oForm.Width = screenWidth - 205
            oForm.Height = screenHeight - 150

            'oForm.Width = screenWidth - 210
            'oForm.Height = screenHeight - 135

            oForm.Left = screenLeft + 150
            oForm.Top = screenTop

            'oForm.Resize(screenWidth, screenHeight)
            'HandleAppEvent(oForm)


            oForm.Visible = True
            oForm.Select()

            rsboApp.ActivateMenuItem("1282")

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla frmGuiasRemision: " + ex.Message.ToString())
        End Try

    End Sub
    Private Sub HandleAppEvent(pform As SAPbouiCOM.Form)
        Try
            For i As Integer = 0 To pform.Items.Count - 1
                Dim oItem As SAPbouiCOM.Item = pform.Items.Item(i)
                ' Ajustar propiedades del elemento según sea necesario
                oItem.Width += 4
                oItem.Height += 4
                ' Puedes ajustar otras propiedades o realizar otras acciones según sea necesario
            Next
        Catch ex As Exception
            rsboApp.MessageBox("Error ajustando los elementos: " & ex.Message)
        End Try
    End Sub
    Private Sub Inicioform(oForm As SAPbouiCOM.Form)

        LlenarComboSeries()

        creaBotonPrueba(oForm)

        CargaItemEnFormulario(oForm, "NUEVO", oForm.TypeEx, oTabla)

        'se agrega el choose from list

        agregarChooseFromList(oForm)




    End Sub

    Private Sub agregarChooseFromList(oForm As SAPbouiCOM.Form)
        oForm.DataSources.UserDataSources.Add("dscodcli", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("dsnomcli", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        'choose para clientes
        AddChooseFromList(oForm)

        AddChooseFromListP(oForm)

        txtcliente = oForm.Items.Item("txtcodCli").Specific
        ' txtcliente.DataBind.SetBound(True, "", "dscodcli")

        txtcliente.ChooseFromListUID = "CFL1"
        txtcliente.ChooseFromListAlias = "CardCode"

        txtnombreCliente = oForm.Items.Item("txtnomCli").Specific
        ' txtnombreCliente.DataBind.SetBound(True, "", "dsnomcli")


    End Sub



    Private Sub LlenarComboSeries()


        Try


            comboSeries = oForm.Items.Item("cbxSerie").Specific
            comboSeries.Item.AffectsFormMode = False

            Dim rs As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(Functions.VariablesGlobales._QueryGuiasDesatendidasSeries)

            If rs.RecordCount > 0 Then

                While Not rs.EoF

                    Dim v = rs.Fields.Item("Series").Value
                    Dim d = rs.Fields.Item("SeriesName").Value

                    comboSeries.ValidValues.Add(v, d)

                    rs.MoveNext()
                End While



            End If

            comboSeries.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            comboSeries.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            oFuncionesB1.Release(rs)


        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("Error en fx LlenarComboSeries " & ex.Message, "frmGuiasRemision")

        End Try




    End Sub


    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

        Try
            Dim typeEx, idForm As String
            typeEx = oFuncionesB1.FormularioActivo(idForm)
            Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(idForm)


            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            If eventInfo.BeforeAction Then
                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                oMenuItem = rsboApp.Menus.Item("1280")
                If oMenuItem.SubMenus.Exists("SS_LOG") Then
                    oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("SS_LOG"))
                End If

                oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "SS_LOG"
                oCreationPackage.String = "(SS) Log Emisión de Documentos..."
                oCreationPackage.Enabled = True
                oCreationPackage.Position = 20
                oMenuItem = rsboApp.Menus.Item("1280")
                oMenus = oMenuItem.SubMenus
                oMenus.AddEx(oCreationPackage)

            Else
                oMenuItem = rsboApp.Menus.Item("1280")
                If oMenuItem.SubMenus.Exists("SS_LOG") Then
                    oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("SS_LOG"))
                End If


            End If


        Catch ex As Exception
            rsboApp.MessageBox("Error: " & ex.Message)
        End Try

    End Sub

    Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

        If eventInfo.FormUID = "frmGuiasRemision" Then


            If eventInfo.ItemUID = "MtxCont" Then

                If eventInfo.ColUID = "COL1" Then
                    Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                    Dim oMenus As SAPbouiCOM.Menus = Nothing

                    If eventInfo.BeforeAction = True Then

                        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

                        If mForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or mForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
                            Try
                                num = eventInfo.Row

                                oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oMenuItem = rsboApp.Menus.Item("1280")
                                If oMenuItem.SubMenus.Exists("Agregar") Then
                                    rsboApp.Menus.RemoveEx("Agregar")

                                End If
                                If oMenuItem.SubMenus.Exists("Eliminar") Then
                                    rsboApp.Menus.RemoveEx("Eliminar")
                                End If
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "Agregar"
                                oCreationPackage.String = "Agregar fila"
                                oCreationPackage.Enabled = True
                                oCreationPackage.Position = 20
                                oMenuItem = rsboApp.Menus.Item("1280")
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)

                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "Eliminar"
                                oCreationPackage.String = "Eliminar fila"
                                oCreationPackage.Enabled = True
                                oCreationPackage.Position = 21
                                oMenuItem = rsboApp.Menus.Item("1280")
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)

                                If oMenuItem.SubMenus.Exists("1283") Then
                                    rsboApp.Menus.RemoveEx("1283")
                                End If

                            Catch ex As Exception
                                'MessageBox.Show(ex.Message)
                            End Try
                        End If
                    Else
                        Try
                            oMenuItem = rsboApp.Menus.Item("1280")
                            If oMenuItem.SubMenus.Exists("Agregar") Then
                                rsboApp.Menus.RemoveEx("Agregar")

                            End If
                            If oMenuItem.SubMenus.Exists("Eliminar") Then
                                rsboApp.Menus.RemoveEx("Eliminar")
                            End If
                        Catch ex As Exception
                            rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                    End If

                End If



            ElseIf eventInfo.ItemUID = "MtxInfad" Then

                If eventInfo.ColUID = "COL1" Then
                    Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                    Dim oMenus As SAPbouiCOM.Menus = Nothing

                    If eventInfo.BeforeAction = True Then

                        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

                        If mForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or mForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
                            Try
                                num = eventInfo.Row

                                oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oMenuItem = rsboApp.Menus.Item("1280")
                                If oMenuItem.SubMenus.Exists("AgregarINF") Then
                                    rsboApp.Menus.RemoveEx("AgregarINF")

                                End If
                                If oMenuItem.SubMenus.Exists("EliminarINF") Then
                                    rsboApp.Menus.RemoveEx("EliminarINF")
                                End If
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "AgregarINF"
                                oCreationPackage.String = "Agregar fila"
                                oCreationPackage.Enabled = True
                                oCreationPackage.Position = 20
                                oMenuItem = rsboApp.Menus.Item("1280")
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)

                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "EliminarINF"
                                oCreationPackage.String = "Eliminar fila"
                                oCreationPackage.Enabled = True
                                oCreationPackage.Position = 21
                                oMenuItem = rsboApp.Menus.Item("1280")
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)

                                If oMenuItem.SubMenus.Exists("1283") Then
                                    rsboApp.Menus.RemoveEx("1283")
                                End If

                            Catch ex As Exception
                                'MessageBox.Show(ex.Message)
                            End Try
                        End If
                    Else
                        Try
                            oMenuItem = rsboApp.Menus.Item("1280")
                            If oMenuItem.SubMenus.Exists("AgregarINF") Then
                                rsboApp.Menus.RemoveEx("AgregarINF")

                            End If
                            If oMenuItem.SubMenus.Exists("EliminarINF") Then
                                rsboApp.Menus.RemoveEx("EliminarINF")
                            End If
                        Catch ex As Exception
                            rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                    End If

                End If



            End If

        End If


    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            Dim typeExx, idFormm As String
            typeExx = oFuncionesB1.FormularioActivo(idFormm)

            If typeExx = "frmGuiasRemision" Then
                If pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
                    rsboApp.Forms.ActiveForm.Freeze(True)
                    Try
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MtxCont").Specific
                        mMatrix.AddRow()
                        For i As Integer = 1 To mMatrix.RowCount
                            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                        Next

                        'Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
                        'comb3.Select("FC", BoSearchKey.psk_ByValue)


                        'mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
                        'mMatrix.Columns.Item("SerDesc").Cells.Item(mMatrix.RowCount).Specific.String = ""

                        'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                        '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                        'End If
                    Catch ex As Exception
                    Finally
                        rsboApp.Forms.ActiveForm.Freeze(False)
                    End Try


                ElseIf pVal.MenuUID = "AgregarINF" And pVal.BeforeAction = False Then
                    rsboApp.Forms.ActiveForm.Freeze(True)
                    Try
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MtxInfad").Specific
                        mMatrix.AddRow()
                        For i As Integer = 1 To mMatrix.RowCount
                            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                        Next

                        'Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
                        'comb3.Select("FC", BoSearchKey.psk_ByValue)


                        'mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
                        'mMatrix.Columns.Item("SerDesc").Cells.Item(mMatrix.RowCount).Specific.String = ""

                        'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                        '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                        'End If
                    Catch ex As Exception
                    Finally
                        rsboApp.Forms.ActiveForm.Freeze(False)
                    End Try


                ElseIf pVal.MenuUID = "1282" Then

                    If pVal.BeforeAction = False Then


                        oTabla = "@SS_GRCAB"

                        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.ActiveForm()

                        Dim txtentry As SAPbouiCOM.EditText

                        txtentry = mForm.Items.Item("txtentry").Specific

                        comboSeries.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                        Dim valorserie = comboSeries.Value
                        Dim anio = Year(Date.Now).ToString
                        'Dim xquery As String = $"Select ""NextNumber"" from ""NNM1"" where ""Remark"" LIKE '%DESATENTIDA% SSGRNEW' and ""Series"" = {valorserie} and ""Indicator""= {anio} "
                        Dim xquery As String = Functions.VariablesGlobales._QueryGuiasDesatendidasProximoDocNum.Replace("@SERIE", valorserie)
                        txtentry.Item.Enabled = False
                        txtentry.Value = CDbl(oFuncionesB1.getRSvalue(xquery, "NextNumber", "0"))

                        Dim Combo As SAPbouiCOM.ComboBox
                        Combo = oForm.Items.Item("cmbTSN").Specific
                        Combo.Item.Enabled = True

                        Combo = oForm.Items.Item("cbxSerie").Specific
                        Combo.Item.Enabled = True

                        Dim txt As SAPbouiCOM.EditText
                        txt = oForm.Items.Item("Item_21").Specific
                        txt.Item.Enabled = True

                        txt = oForm.Items.Item("Item_23").Specific
                        txt.Item.Enabled = True

                        txt = oForm.Items.Item("Item_6").Specific
                        txt.Item.Enabled = True

                        Dim mtx As SAPbouiCOM.Matrix = oForm.Items.Item("MtxCont").Specific
                        mtx.Item.Enabled = True

                        CargaItemEnFormulario(mForm, "NUEVO", mForm.TypeEx, oTabla)


                    End If
                ElseIf pVal.MenuUID = "SS_LOG" And pVal.BeforeAction = False Then
                    Try
                        Dim typeEx As String = "", idForm As String = ""
                        typeEx = oFuncionesB1.FormularioActivo(idForm)
                        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(idForm)

                        'SeteaTipoTabla_FormTypeEx(typeEx)

                        Dim lDocEntry As String = mForm.DataSources.DBDataSources.Item("@SS_GRCAB").GetValue("DocEntry", 0)
                        Dim lObjType As String = "'SSGR'" 'mForm.DataSources.DBDataSources.Item("@SS_GRCAB").GetValue("Object", 0)
                        Dim lDocSubType As String = "--"

                        ofrmLogEmision.CargaFormularioLogEmision(lDocEntry, lObjType, lDocSubType)
                    Catch ex As Exception

                    End Try

                End If
            End If

            'If rsboApp.Forms.ActiveForm.UniqueID = "frmGuiasRemision" And pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
            '    rsboApp.Forms.ActiveForm.Freeze(True)
            '    Try
            '        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MtxCont").Specific
            '        mMatrix.AddRow()
            '        For i As Integer = 1 To mMatrix.RowCount
            '            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            '        Next

            '        'Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
            '        'comb3.Select("FC", BoSearchKey.psk_ByValue)


            '        'mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
            '        'mMatrix.Columns.Item("SerDesc").Cells.Item(mMatrix.RowCount).Specific.String = ""

            '        'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '        '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '        'End If
            '    Catch ex As Exception
            '    Finally
            '        rsboApp.Forms.ActiveForm.Freeze(False)
            '    End Try


            'ElseIf rsboApp.Forms.ActiveForm.UniqueID = "frmGuiasRemision" And pVal.MenuUID = "AgregarINF" And pVal.BeforeAction = False Then
            '    rsboApp.Forms.ActiveForm.Freeze(True)
            '    Try
            '        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MtxInfad").Specific
            '        mMatrix.AddRow()
            '        For i As Integer = 1 To mMatrix.RowCount
            '            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            '        Next

            '        'Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
            '        'comb3.Select("FC", BoSearchKey.psk_ByValue)


            '        'mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
            '        'mMatrix.Columns.Item("SerDesc").Cells.Item(mMatrix.RowCount).Specific.String = ""

            '        'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '        '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '        'End If
            '    Catch ex As Exception
            '    Finally
            '        rsboApp.Forms.ActiveForm.Freeze(False)
            '    End Try


            'ElseIf rsboApp.Forms.ActiveForm.UniqueID = "frmGuiasRemision" And pVal.MenuUID = "1282" Then

            '    If pVal.BeforeAction = False Then


            '        oTabla = "@SS_GRCAB"

            '        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.ActiveForm()

            '        Dim txtentry As SAPbouiCOM.EditText

            '        txtentry = mForm.Items.Item("txtentry").Specific

            '        comboSeries.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            '        Dim valorserie = comboSeries.Value
            '        Dim anio = Year(Date.Now).ToString
            '        'Dim xquery As String = $"Select ""NextNumber"" from ""NNM1"" where ""Remark"" LIKE '%DESATENTIDA% SSGRNEW' and ""Series"" = {valorserie} and ""Indicator""= {anio} "
            '        Dim xquery As String = Functions.VariablesGlobales._QueryGuiasDesatendidasProximoDocNum.Replace("@SERIE", valorserie)
            '        txtentry.Item.Enabled = False
            '        txtentry.Value = CDbl(oFuncionesB1.getRSvalue(xquery, "NextNumber", "0"))

            '        Dim Combo As SAPbouiCOM.ComboBox
            '        Combo = oForm.Items.Item("cmbTSN").Specific
            '        Combo.Item.Enabled = True

            '        Combo = oForm.Items.Item("cbxSerie").Specific
            '        Combo.Item.Enabled = True

            '        Dim txt As SAPbouiCOM.EditText
            '        txt = oForm.Items.Item("Item_21").Specific
            '        txt.Item.Enabled = True

            '        txt = oForm.Items.Item("Item_23").Specific
            '        txt.Item.Enabled = True

            '        txt = oForm.Items.Item("Item_6").Specific
            '        txt.Item.Enabled = True

            '        Dim mtx As SAPbouiCOM.Matrix = oForm.Items.Item("MtxCont").Specific
            '        mtx.Item.Enabled = True

            '        CargaItemEnFormulario(mForm, "NUEVO", mForm.TypeEx, oTabla)


            '    End If
            'ElseIf pVal.MenuUID = "SS_LOG" And pVal.BeforeAction = False Then
            '    Try
            '        Dim typeEx As String = "", idForm As String = ""
            '        typeEx = oFuncionesB1.FormularioActivo(idForm)
            '        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(idForm)

            '        'SeteaTipoTabla_FormTypeEx(typeEx)

            '        Dim lDocEntry As String = mForm.DataSources.DBDataSources.Item("@SS_GRCAB").GetValue("DocEntry", 0)
            '        Dim lObjType As String = "'SSGR'" 'mForm.DataSources.DBDataSources.Item("@SS_GRCAB").GetValue("Object", 0)
            '        Dim lDocSubType As String = "--"

            '        ofrmLogEmision.CargaFormularioLogEmision(lDocEntry, lObjType, lDocSubType)
            '    Catch ex As Exception

            '    End Try

            'End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmGuiasRemision")

        End Try
    End Sub

    Private Sub rsboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.FormDataEvent

        If BusinessObjectInfo.FormTypeEx = "frmGuiasRemision" Then


            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If Not BusinessObjectInfo.BeforeAction Then
                        Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(BusinessObjectInfo.FormUID)

                        '  nombreFormulario = mForm.Title

                        If Not nombreFormulario.Contains("ELECTRONICO") Then
                            nombreFormulario = mForm.Title
                        End If
                        Try
                            Dim oFormMode As Integer = 1
                            oFormMode = mForm.Mode



                            oTabla = "@SS_GRCAB"

                            oTipoTabla = "SSGR"


                            CargaItemEnFormulario(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla, BusinessObjectInfo.FormUID)

                            Dim Combo As SAPbouiCOM.ComboBox
                            Combo = oForm.Items.Item("cmbTSN").Specific
                            Combo.Item.Enabled = False

                            Combo = oForm.Items.Item("cbxSerie").Specific
                            Combo.Item.Enabled = False

                            Dim txt As SAPbouiCOM.EditText
                            txt = oForm.Items.Item("Item_21").Specific
                            txt.Item.Enabled = False

                            txt = oForm.Items.Item("Item_23").Specific
                            txt.Item.Enabled = False

                            txt = oForm.Items.Item("Item_6").Specific
                            txt.Item.Enabled = False

                            Dim mtx As SAPbouiCOM.Matrix = oForm.Items.Item("MtxCont").Specific
                            mtx.Item.Enabled = False

                            mForm.Mode = oFormMode


                        Catch ex As Exception
                            rsboApp.SetStatusBarMessage("et_FORM_DATA_LOAD", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try


                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If Not BusinessObjectInfo.BeforeAction Then
                        Select Case BusinessObjectInfo.ActionSuccess
                            Case True
                                ' If Functions.VariablesGlobales._gActivarEnvioDocumentosBackGround = "Y" Then
                                '   rsboApp.SetStatusBarMessage(NombreAddon + " - El documento se enviara despues de un momento a la DIAN, y posterior le llegara el correo al cliente", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                ' Else
                                If BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts.ToString Then
                                    rsboApp.SetStatusBarMessage(NombreAddon + " - El documento creado es un documento Preliminar, no se procesará..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                Else
                                    Utilitario.Util_Log.Escribir_Log("paso oDrafts", "EventosEmision")

                                    Dim form As SAPbouiCOM.Form = rsboApp.Forms.Item(BusinessObjectInfo.FormUID)

                                    If EsElectronico = "FE" Then

                                        oTipoTabla = "SSGR"

                                        '   SeteaTipoTabla_FormTypeEx(BusinessObjectInfo.FormTypeEx)

                                        Dim odbds As SAPbouiCOM.DBDataSource = CType(form.DataSources.DBDataSources.Item(0), SAPbouiCOM.DBDataSource)
                                        Dim soDocEntry As String = ""
                                        Try
                                            soDocEntry = odbds.GetValue("DocEntry", odbds.Offset).Trim
                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Error al obtener el valor de DocEntry:" + ex.Message.ToString(), "EventosEmision")
                                        End Try

                                        Dim soCANCELED As String = ""
                                        Try
                                            soCANCELED = odbds.GetValue("CANCELED", odbds.Offset).Trim
                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Error al obtener el valor de CANCELED:" + ex.Message.ToString(), "EventosEmision")
                                        End Try

                                        Dim Series As String = ""
                                        Try
                                            Series = odbds.GetValue("Series", odbds.Offset).Trim
                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Error al obtener el valor de Series:" + ex.Message.ToString(), "EventosEmision")
                                        End Try

                                        oFuncionesAddon.GuardaLOG(oTipoTabla, soDocEntry, "Obteniendo información del objeto creado", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                        If Not soCANCELED = "C" Then ' SI EL DOCUMENTO NO ES UNA CANCELACION
                                            Utilitario.Util_Log.Escribir_Log("paso Cancelacion", "EventosEmision")
                                            ' 'If Not oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then   
                                            ' Utilitario.Util_Log.Escribir_Log("Es oDrafts", "EventosEmision")
                                            'If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then

                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                            oFuncionesAddon.GuardaLOG(oTipoTabla, soDocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                            'Dim querysqlTipoECF = "SELECT B.""U_TipoD"" FROM ""@GS_SERD"" B WHERE B.""U_SerId"" =" + Series

                                            'Dim TipoECF = oFuncionesB1.getRSvalue(querysqlTipoECF, "U_TipoD", "")
                                            Dim _CardCode As String = ""
                                            Try
                                                _CardCode = odbds.GetValue("CardCode", odbds.Offset).Trim
                                            Catch ex As Exception
                                                Utilitario.Util_Log.Escribir_Log("Error al obtener el valor de CardCode:" + ex.Message.ToString(), "EventosEmision")
                                            End Try


                                            Dim NombreNCF As String = "", TipoECF As String = "", IndicadorElectronico As String = ""

                                            ' ObtnerNombreCodigoNCF(form, TipoECF, NombreNCF, IndicadorElectronico, _CardCode)

                                            oManejoDocumentos.ProcesaEnvioDocumento(CInt(soDocEntry), oTipoTabla)

                                            'End If
                                        End If
                                    End If
                                End If
                                ' End If
                        End Select

                    Else
                        'antes de la Accion

                        If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts.ToString Then

                            If EsElectronico = "FE" Then

                                Dim form As SAPbouiCOM.Form = rsboApp.Forms.Item(BusinessObjectInfo.FormUID)
                                Dim oUDFForm As SAPbouiCOM.Form
                                'oUDFForm = rsboApp.Forms.Item(form.UDFFormUID)

                                'Try
                                '    oUDFForm.Items.Item("U_SS_PeticionID").Specific.String = Guid.NewGuid.ToString.ToUpper
                                'Catch ex As Exception
                                '    Utilitario.Util_Log.Escribir_Log("Error al generar peticion ID: " + ex.Message.ToString(), "EventosEmision")
                                '    Throw
                                'End Try


                                Try
                                    form.Items.Item("txtclav").Specific.String = ""
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al limpiar clave de acceso: " + ex.Message.ToString(), "EventosEmision")
                                    Throw
                                End Try

                                Try
                                    Dim cbEstado As SAPbouiCOM.ComboBox
                                    cbEstado = form.Items.Item("cbxestaut").Specific
                                    cbEstado.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    'oUDFForm.Items.Item("U_ESTADO_AUTORIZACIO").Specific.String = "0"
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al limpiar el estado de la autorizacion: " + ex.Message.ToString(), "EventosEmision")
                                    Throw
                                End Try


                                Try
                                    form.Update()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al actualizar formulario : " + ex.Message.ToString(), "EventosEmision")
                                    Throw
                                Finally
                                    '    oFuncionesB1.Release(oUDFForm)
                                    oFuncionesB1.Release(form)
                                End Try


                            End If

                        End If

                    End If

            End Select



        End If




    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent



        If pVal.FormTypeEx = "frmGuiasRemision" Then


            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    Select Case pVal.ItemUID
                        Case "cbxSerie"

                            If Not pVal.BeforeAction AndAlso pVal.ItemChanged Then

                                oTipoTabla = "SSGR"
                                oTabla = "@SS_GRCAB"

                                Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)

                                Dim txtentry As SAPbouiCOM.EditText

                                txtentry = mForm.Items.Item("txtentry").Specific

                                Dim valorserie = comboSeries.Value

                                'Dim xquery As String = $"Select ""NextNumber"" from ""NNM1"" where ""ObjectCode""='SSGRNEW' and ""Series"" = {valorserie}"
                                Dim xquery As String = Functions.VariablesGlobales._QueryGuiasDesatendidasProximoDocNum.Replace("@SERIE", valorserie)

                                txtentry.Item.Enabled = False
                                txtentry.Value = CDbl(oFuncionesB1.getRSvalue(xquery, "NextNumber", "0"))

                                CargaItemEnFormulario(mForm, "NUEVO", mForm.TypeEx, oTabla)

                            Else

                            End If

                        Case "cmbTSN"

                            If Not pVal.Before_Action Then
                                Try

                                    oForm = rsboApp.Forms.Item("frmGuiasRemision")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox
                                    cbxTipo = oForm.Items.Item("cmbTSN").Specific

                                    If Not IsNothing(oCFL) Then
                                        oCons = oCFL.GetConditions()

                                        Dim lbSocio As SAPbouiCOM.StaticText
                                        lbSocio = oForm.Items.Item("5").Specific



                                        If oCons.Count > 0 Then 'If there are already user conditions.
                                            If cbxTipo.Value = "01" Then ' SI ES 07, SIGNIFICA QUE ES RETENCION, POR ENDE PAGO RECIBIDO DE CLIENTE
                                                oCons.Item(oCons.Count - 1).CondVal = "C"
                                                lbSocio.Caption = "Cliente :"
                                                txtcliente.ChooseFromListUID = "CFL1"
                                                txtcliente.ChooseFromListAlias = "CardCode"

                                            Else
                                                'oCons.Item(oCons.Count - 1).CondVal = "S"
                                                lbSocio.Caption = "Proveedor :"
                                                txtcliente.ChooseFromListUID = "CFL2"
                                                txtcliente.ChooseFromListAlias = "CardCode"


                                            End If
                                        End If
                                        'oCFLCreationParams.ObjectType = "2"
                                        'oCFL.SetConditions(oCons)
                                        'Dim txtRuc As SAPbouiCOM.EditText
                                        'txtRuc = oForm.Items.Item("txtcodCli").Specific
                                        'txtRuc.ChooseFromListUID = "CFL1"
                                        'txtRuc.ChooseFromListAlias = "CardType"
                                    End If



                                Catch ex As Exception

                                End Try


                            End If

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    Try

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form
                        oForm = rsboApp.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)


                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim codigo As String = ""
                            Dim nombre As String = ""
                            Try

                                codigo = oDataTable.GetValue(0, 0)
                                nombre = oDataTable.GetValue(1, 0)

                            Catch ex As Exception

                            End Try

                            'para el cliente
                            If truco = False Then
                                If (pVal.ItemUID = "txtcodCli" And oCFL.UniqueID = "CFL1") Then


                                    Try
                                        truco = True

                                        Try

                                            'oUserDataSourceDE = oForm.DataSources.UserDataSources.Item("EditDS")
                                            'oUserDataSourceDE.ValueEx = codigo
                                            Dim txtnomCli As SAPbouiCOM.EditText
                                            txtnomCli = oForm.Items.Item("txtnomCli").Specific
                                            txtnomCli.Value = nombre

                                        Catch ex As Exception
                                        End Try

                                        Try
                                            Dim txtcodCli As SAPbouiCOM.EditText
                                            txtcodCli = oForm.Items.Item("txtcodCli").Specific
                                            txtcodCli.Value = codigo
                                        Catch ex As Exception
                                        End Try

                                        'txtnombreCliente.Value = nombre
                                        'txtcliente.Value = codigo

                                    Catch ex As Exception
                                    Finally
                                        ' Rehabilitar el evento después de setear el valor
                                        truco = False

                                    End Try
                                ElseIf (pVal.ItemUID = "txtcodCli" And oCFL.UniqueID = "CFL2") Then


                                    Try
                                            truco = True

                                            Try

                                                'oUserDataSourceDE = oForm.DataSources.UserDataSources.Item("EditDS")
                                                'oUserDataSourceDE.ValueEx = codigo
                                                Dim txtnomCli As SAPbouiCOM.EditText
                                                txtnomCli = oForm.Items.Item("txtnomCli").Specific
                                                txtnomCli.Value = nombre

                                            Catch ex As Exception
                                            End Try

                                            Try
                                                Dim txtcodCli As SAPbouiCOM.EditText
                                                txtcodCli = oForm.Items.Item("txtcodCli").Specific
                                                txtcodCli.Value = codigo
                                            Catch ex As Exception
                                            End Try

                                            'txtnombreCliente.Value = nombre
                                            'txtcliente.Value = codigo

                                        Catch ex As Exception
                                        Finally
                                            ' Rehabilitar el evento después de setear el valor
                                            truco = False

                                        End Try

                                        'para la matrix
                                    ElseIf (pVal.ItemUID = "MtxCont" And pVal.ColUID = "txtitCod" And oCFL.UniqueID = "CFL_i") Then
                                    'el tipo de la columna debe ser tipo liken but
                                    Dim mtx As SAPbouiCOM.Matrix = oForm.Items.Item("MtxCont").Specific


                                    mtx.SetCellWithoutValidation(pVal.Row, "txtitCod", codigo)
                                    mtx.SetCellWithoutValidation(pVal.Row, "txtitNom", nombre)


                                End If
                            End If


                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error en choosefromlist Cliente " & ex.Message, "frmGuiasRemision")
                    End Try



                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "1"

                            If Not pVal.BeforeAction Then

                                'oTabla = "@SS_GRCAB"

                                'Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)

                                'CargaItemEnFormulario(mForm, "NUEVO", mForm.TypeEx, oTabla)


                            Else

                            End If

                        Case "btnIma"

                            If Not pVal.BeforeAction Then



                                Dim selectFileDialog As New SelectFileDialog("C:\", "", "Image Files|*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tiff;*.ico;*.svg|All Files|*.*", DialogType.OPEN)
                                selectFileDialog.Open()



                                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                                    Dim s As String
                                    s = selectFileDialog.SelectedFile
                                    'consulto ruta donde se almacena imagenea en parametrizacion general de sap
                                    Dim rutaSAP = ""
                                    rutaSAP = oFuncionesB1.getRSvalue("select ""BitmapPath"" from OADP", "BitmapPath", "")
                                    Dim nombreImagen = System.IO.Path.GetFileNameWithoutExtension(s) & (System.IO.Path.GetExtension(s))
                                    Dim rutaNombre = rutaSAP & nombreImagen
                                    If File.Exists(s) Then
                                        File.Copy(s, rutaNombre, True)
                                    End If
                                    If rutaSAP = "" Or IsNothing(rutaSAP) Then
                                        rsboApp.StatusBar.SetText(NombreAddon + " - No se encuentra parametrizada la ruta de imagenes en la Parametrizacion General de SAP", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        'Exit case
                                    Else

                                        oForm = rsboApp.Forms.Item("frmGuiasRemision")
                                        Dim Item_12 As SAPbouiCOM.PictureBox = oForm.Items.Item("Item_12").Specific
                                        Item_12.Picture = rutaNombre
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If

                                End If
                            End If
                        Case "btnAccion"

                            If Not pVal.BeforeAction Then

                                Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)
                                btnAccion = mForm.Items.Item("btnAccion").Specific

                                oTipoTabla = "SSGR"

                                If btnAccion.Caption = "(GS) Ver RIDE" Then

                                    Try
                                        Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_CLAVE_ACCESO", 0)
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Consultando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        Dim docentry As String = ""
                                        docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAcceso, docentry, oTipoTabla, "pdf")
                                        Else
                                            oManejoDocumentos.ConsultaPDF(ClaveAcceso)
                                        End If

                                    Catch x As TimeoutException
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del documento! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Catch ex As Exception
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End Try

                                ElseIf btnAccion.Caption = "(GS) Ver XML" Then
                                    Dim ClaveAccesoXML As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_CLAVE_ACCESO", 0)
                                    Try
                                        rsboApp.SetStatusBarMessage(NombreAddon + " -Consultando el XML, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                                        '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                                        'End If
                                        'oManejoDocumentos.SetProtocolosdeSeguridad()
                                        Dim docentry As String = ""
                                        docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAccesoXML, docentry, oTipoTabla, "xml")
                                        Else
                                            oManejoDocumentos.ConsultaXML(ClaveAccesoXML)
                                        End If

                                    Catch x As TimeoutException
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del XML! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Catch ex As Exception
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End Try


                                ElseIf btnAccion.Caption = "(GS) Consultar AUT" Then

                                    Try
                                        Dim docentry As String = ""
                                        docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla, True)
                                        Else
                                            oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla, True)
                                        End If

                                        Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                                            mForm.Freeze(True)
                                            rsboApp.ActivateMenuItem("1289")
                                            rsboApp.ActivateMenuItem("1288")
                                        Catch ex As Exception
                                        Finally
                                            mForm.Freeze(False)
                                        End Try

                                    Catch ex As Exception
                                        rsboApp.SetStatusBarMessage("Error al intentar Consultar la autorizacion desde eDoc " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    End Try

                                ElseIf btnAccion.Caption = "(GS) Reenviar SRI" Then
                                    oTipoTabla = "SSGR"

                                    If EsElectronico = "FE" Then

                                        Dim odbds As SAPbouiCOM.DBDataSource = CType(mForm.DataSources.DBDataSources.Item(0), SAPbouiCOM.DBDataSource)
                                        Dim soDocEntry As String = ""
                                        Try
                                            soDocEntry = odbds.GetValue("DocEntry", odbds.Offset).Trim
                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Error al obtener el valor de DocEntry:" + ex.Message.ToString(), "EventosEmision")
                                        End Try


                                        oManejoDocumentos.ProcesaEnvioDocumento(CInt(soDocEntry), oTipoTabla)

                                        Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                                            mForm.Freeze(True)
                                            rsboApp.ActivateMenuItem("1289")
                                            rsboApp.ActivateMenuItem("1288")
                                        Catch ex As Exception
                                        Finally
                                            mForm.Freeze(False)
                                        End Try

                                    End If


                                End If

                            Else

                                'Destinado para antes del evento

                            End If

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    Select Case pVal.ItemUID

                        Case "txtnomCli"

                            If pVal.BeforeAction = False Then

                                If pVal.CharPressed = 13 Then

                                    rsboApp.Forms.ActiveForm.Freeze(True)
                                    Try
                                        Dim idRow = 0
                                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MtxCont").Specific
                                        mMatrix.AddRow()
                                        For i As Integer = 1 To mMatrix.RowCount
                                            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                                            idRow = i
                                        Next
                                        mMatrix.Columns.Item("txtitCod").Cells.Item(idRow).Click()
                                    Catch ex As Exception
                                    Finally
                                        rsboApp.Forms.ActiveForm.Freeze(False)
                                    End Try
                                End If

                            End If
                        Case "MtxCont"

                            If pVal.ColUID = "Col_5" Then

                                If pVal.BeforeAction = False Then

                                    rsboApp.Forms.ActiveForm.Freeze(True)
                                    Try
                                        Dim idRow = 0
                                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MtxCont").Specific
                                        mMatrix.AddRow()
                                        For i As Integer = 1 To mMatrix.RowCount
                                            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                                            idRow = i
                                        Next
                                        mMatrix.Columns.Item("txtitCod").Cells.Item(idRow).Click()
                                    Catch ex As Exception
                                    Finally
                                        rsboApp.Forms.ActiveForm.Freeze(False)
                                    End Try

                                End If

                            End If

                    End Select



            End Select
        End If



    End Sub

    Private Sub AddChooseFromList(ByRef oForm As SAPbouiCOM.Form)
        Try



            oCFLs = oForm.ChooseFromLists


            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            ' oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL1"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)



        Catch
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub AddChooseFromListP(ByRef oForm As SAPbouiCOM.Form)
        Try



            oCFLs = oForm.ChooseFromLists


            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            ' oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL2"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            txtcliente = oForm.Items.Item("txtcodCli").Specific
            ' txtcliente.DataBind.SetBound(True, "", "dscodcli")

            txtcliente.ChooseFromListUID = "CFL2"
            txtcliente.ChooseFromListAlias = "CardCode"

            txtnombreCliente = oForm.Items.Item("txtnomCli").Specific

        Catch
            MsgBox(Err.Description)
        End Try
    End Sub


    Private Sub creaBotonPrueba(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try
            Dim _left As Integer = 0, _top As Integer = 0, _espaciado As Integer = 15

            Dim ItemParaLinkeo As String = "5"

            If oTipoTabla = "TRE" Then
                _left = oForm.Items.Item("9").Left + 135

            ElseIf oTipoTabla = "TLE" Then
                _left = oForm.Items.Item("9").Left + 135
            Else
                _left = oForm.Items.Item("txtnomCli").Left
            End If


            ' COMENTARIO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbComen", "(GS) Documento Electrónico", _left, 75 - 10, 250, 14, 0, False)
                oForm.Items.Item("lbComen").LinkTo = ItemParaLinkeo

            Catch ex As Exception
                rsboApp.SetStatusBarMessage("No se pudo crear la label de Comentario", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try


            ' ESTADO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbEstado", "(GS) Estado", _left, 90 - 10, 250, 14, 0, False)
                oForm.Items.Item("lbEstado").LinkTo = ItemParaLinkeo

            Catch ex As Exception
                rsboApp.SetStatusBarMessage("No se pudo crear la label de estado", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

            ' Boton COmbo
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO, "btnAccion", "", _left, 105 - 10, 110, 19, 0, False)
                oForm.Items.Item("btnAccion").LinkTo = ItemParaLinkeo
            Catch ex As Exception
                rsboApp.SetStatusBarMessage("No se pudo crear El boton combo Accion", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try



        Catch ex As Exception
            rsboApp.SetStatusBarMessage(NombreAddon + " - Botones Prueba ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)

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


    Private Sub CargaItemEnFormulario(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String, Optional formUID As String = "")
        Dim codDoc As String = "00"

        Try

            Dim oSerie As String = ""
            Dim sql As String = ""
            ' OBTENGO EL NUMERO DE SERIE

            oSerie = oForm.Items.Item("cbxSerie").Specific.value.ToString()


            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                If oTipoTabla = "SSGR" Then

                    sql = "Select IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + oSerie

                End If

            Else

                If oTipoTabla = "SSGR" Then

                    sql = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + oSerie

                End If

            End If

            Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + sql, "EventosEmision")


            EsElectronico = oFuncionesB1.getRSvalue(sql, "U_FE_TipoEmision", "")

            If EsElectronico = "NAN" Or EsElectronico = "" Then
                EsElectronico = "NA"
            Else
                EsElectronico = "FE"
            End If




            'FRIZEO ANTES DE CUALQUIER DIBUJADO

            oForm.Freeze(True)

            If evento = "NUEVO" Then

                lbEstado = oForm.Items.Item("lbEstado").Specific
                lbEstado.Item.ForeColor = RGB(204, 0, 0)
                lbEstado.Caption = "Estado: NO ENVIADO"
                lbEstado.Item.Visible = True

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                lbComentario = oForm.Items.Item("lbComen").Specific

                If EsElectronico = "FE" Then

                    If Not oForm.Title.Contains("ELECTRONICO") Then
                        oForm.Title += " - ELECTRONICO"
                    End If



                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(7, 118, 10)
                    ' lbComentario.Caption = "ELECTRONICO"
                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO NO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(204, 0, 0)
                End If
                lbComentario.Item.Visible = True
                If (oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM") And EsElectronico <> "FE" Then
                    lbComentario.Item.Visible = False
                    lbEstado.Item.Visible = False
                End If
                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rsboApp.SetStatusBarMessage("NUEVO-Eliminacion valores validos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                ' se trata de crear un codigo QR





            ElseIf evento = "ACTUALIZAR" Then

                lbComentario = oForm.Items.Item("lbComen").Specific
                lbEstado = oForm.Items.Item("lbEstado").Specific
                'cbEstado = oForm.Items.Item("cbEstado").Specific
                'btnAccion = oForm.Items.Item("btnAccion").Specific

                'Obtengo valor almacenado en base
                Dim UDFEA As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_ESTADO_AUTORIZACIO", 0)

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rsboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEA) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEA = "0"
                End If
                UDFEA = UDFEA.Trim

                Try
                    If EsElectronico = "FE" Then
                        If UDFEA = "2" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: AUTORIZADO"

                        ElseIf UDFEA = "5" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: RECIBIDO"

                        ElseIf UDFEA = "7" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: ERROR RECEPCION SRI"

                        ElseIf UDFEA = "4" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: VALIDAR DATOS"

                        ElseIf UDFEA = "3" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO AUTORIZADA"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEA = "11" Then
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = False
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: ANULADO"

                        Else
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO ENVIADO"

                        End If

                    End If
                Catch ex As Exception
                    rsboApp.SetStatusBarMessage("Actualizar-FE ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If EsElectronico = "FE" Then

                    If Not oForm.Title.Contains("ELECTRONICO") Then
                        oForm.Title += " - ELECTRONICO"
                    End If

                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(7, 118, 10)

                    Select Case oForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            lbComentario.Item.Visible = True
                            lbEstado.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Then
                                If UDFEA = "0" Then
                                    btnAccion.Item.Visible = False
                                End If
                            End If

                        Case Else
                            btnAccion.Item.Visible = False

                    End Select

                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO NO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(204, 0, 0)

                    lbEstado.Caption = "Estado: NO ENVIADO"
                    lbEstado.Item.ForeColor = RGB(204, 0, 0)
                    btnAccion.Item.Visible = False

                    If oTipoTabla = "REE" Then
                        lbComentario.Item.Visible = False
                        lbEstado.Item.Visible = False
                    End If


                End If


                'se intenta crear imagenes QR

                Try

                    If oForm.PaneLevel = 50 Then

                        '  ObtenerEnlacesURLyGenerarRQ(oForm, formUID)

                    End If
                Catch ex As Exception

                End Try


                If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                    'se oculta folios generados si no es modo ADD

                    Try

                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            '  oForm.Items.Item("ptexto").Visible = False
                            '  oForm.Items.Item("pfolio").Visible = False

                        End If

                    Catch ex As Exception

                    End Try

                    'Solo para mostrar el folio de la solicitud
                    Try

                        If oForm.TypeEx = "1250000940" Then

                            Dim foliosolicitud As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0).Trim

                            oForm.Items.Item("etssloc80").Enabled = False
                            oForm.Items.Item("etssloc80").Specific.value = foliosolicitud

                        End If


                    Catch ex As Exception

                    End Try

                End If



            End If

        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oFrmUser.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Private Sub SeteaTipoTabla_FormTypeEx(FormTypeEx As String)
        Try
            If FormTypeEx = "133" Or FormTypeEx = "60090" Or FormTypeEx = "65307" Then
                oTabla = "OINV"
                oTipoTabla = "FCE"
            ElseIf FormTypeEx = "60091" Then
                oTabla = "OINV"
                oTipoTabla = "FRE"
            ElseIf FormTypeEx = "65303" Then
                oTabla = "OINV"
                oTipoTabla = "NDE"
            ElseIf FormTypeEx = "65300" Then
                oTabla = "ODPI"
                oTipoTabla = "FAE"
            ElseIf FormTypeEx = "179" Then
                oTabla = "ORIN"
                oTipoTabla = "NCE"
            ElseIf FormTypeEx = "140" Then
                oTabla = "ODLN"
                oTipoTabla = "GRE"
            ElseIf FormTypeEx = "940" Then
                oTabla = "OWTR"
                oTipoTabla = "TRE"
            ElseIf FormTypeEx = "1250000940" Then
                oTabla = "OWTQ"
                oTipoTabla = "TLE"
            ElseIf FormTypeEx = "141" Then
                oTabla = "OPCH"
                oTipoTabla = "REE"

            ElseIf FormTypeEx = "65306" Then
                oTabla = "OPCH"
                oTipoTabla = "RDM"

            ElseIf FormTypeEx = "60092" Then   'FACTURA DE RESERVA PROVEEDOR/RETENCION
                oTabla = "OPCH"
                oTipoTabla = "RER"
            ElseIf FormTypeEx = "65301" Then ' 'FACTURA DE ANTICIPO DE PROVEEDORES
                oTabla = "ODPO"
                oTipoTabla = "REA"
            ElseIf FormTypeEx = "720" Then ' SOLICITUD DE MERCANCIAS
                oTabla = "OIGE"
                oTipoTabla = "GRSM"
            ElseIf FormTypeEx = "frmGuiasRemision" Then
                oTabla = "SSGR"
                oTipoTabla = "SSGR"
            End If
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(NombreAddon + " - Error al Setear Tipo Tabla: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

End Class

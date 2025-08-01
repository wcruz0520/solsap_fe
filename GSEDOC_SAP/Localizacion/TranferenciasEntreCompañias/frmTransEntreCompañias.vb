Imports System.Deployment
Imports System.Globalization
Imports System.Threading
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class frmTransEntreCompanias

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private rCompany2 As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Dim oUserDataSource As SAPbouiCOM.UserDataSource

    Public _num As Integer = 0

    Private rMatrix As SAPbouiCOM.Matrix

    Dim CodigoProd As SAPbouiCOM.EditText = Nothing
    Dim NombreProd As SAPbouiCOM.EditText = Nothing
    Dim CostoProd As SAPbouiCOM.EditText = Nothing

    Dim EditText As SAPbouiCOM.EditText
    Dim Label As SAPbouiCOM.StaticText



    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
        rCompany2 = Nothing
    End Sub

    Public Sub CreaFormulario_frmTransEntreCompanias()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmTransEntreCompanias") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmTransEntreCompanias.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmTransEntreCompanias").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmTransEntreCompanias")

            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"


            Dim btnVer As SAPbouiCOM.Button
            btnVer = oForm.Items.Item("btnVer").Specific
            btnVer.Item.Visible = False



            Dim txtBO As SAPbouiCOM.EditText
            txtBO = oForm.Items.Item("txtBO").Specific
            txtBO.Item.Enabled = False
            txtBO.Value = rCompany.CompanyName

            Dim oCFLsA As SAPbouiCOM.ChooseFromListCollection
            Dim oConsA As SAPbouiCOM.Conditions
            Dim oConA As SAPbouiCOM.Condition
            oCFLsA = oForm.ChooseFromLists
            Dim oCFLA As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParamsA As SAPbouiCOM.ChooseFromListCreationParams
            ' CHOOSE FROM LIST
            oCFLsA = oForm.ChooseFromLists
            oCFLCreationParamsA = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParamsA.MultiSelection = False
            oCFLCreationParamsA.ObjectType = "64"
            'oCFLCreationParams.ObjectType = "Exx_DEPOTRANS"
            oCFLCreationParamsA.UniqueID = "CFL1A"
            oCFLA = oCFLsA.Add(oCFLCreationParamsA)
            ' Adding Conditions to CFL1
            oConsA = oCFLA.GetConditions()

            oConA = oConsA.Add()
            oConA.Alias = "WhsCode"
            oConA.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oConA.CondVal = "C"
            oCFLA.SetConditions(oConsA)
            ' END CHOOSE FROM LIST
            oCFLCreationParamsA.UniqueID = "CFL2A"
            oCFLA = oCFLsA.Add(oCFLCreationParamsA)
            oCFLA.SetConditions(oConsA)

            Dim txtAlmO As SAPbouiCOM.EditText
            txtAlmO = oForm.Items.Item("txtAlmO").Specific
            oForm.DataSources.UserDataSources.Add("EditDSA", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtAlmO.DataBind.SetBound(True, "", "EditDSA")
            txtAlmO.ChooseFromListUID = "CFL2A"
            txtAlmO.ChooseFromListAlias = "WhsCode"


            oForm.DataSources.UserDataSources.Add("dtFC", SAPbouiCOM.BoDataType.dt_DATE, 20)

            Dim txtFecha As SAPbouiCOM.EditText
            txtFecha = oForm.Items.Item("txtFecha").Specific
            txtFecha.DataBind.SetBound(True, "", "dtFC")
            txtFecha.Value = DateTime.Now.ToString("yyyyMMdd")


            'Dim txtAlmacen As SAPbouiCOM.EditText
            'txtAlmacen = oForm.Items.Item("txtAlmacen").Specific
            'oForm.DataSources.UserDataSources.Add("EditDSA2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'txtAlmacen.DataBind.SetBound(True, "", "EditDSA2")
            'txtAlmacen.ChooseFromListUID = "CFL1AD"
            'txtAlmacen.ChooseFromListAlias = "WhsCode"

            Dim cbcBases As SAPbouiCOM.ComboBox
            cbcBases = oForm.Items.Item("cbcBases").Specific
            Dim oRecordSet As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If rCompany.DbServerType = 9 Then
                CargaComboboxSAP("select ""Code"",""Name"" from " + rCompany.CompanyDB.ToString() + ".""@SS_BASES"" ", cbcBases, "Name", "", oRecordSet)
            Else
                CargaComboboxSAP("select Code,Name from ""@SS_BASES"" ", cbcBases, "Name", "", oRecordSet)
            End If
            cbcBases.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
            cbcBases.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            'Dim Item_1 As SAPbouiCOM.StaticText
            'Item_1 = oForm.Items.Item("Item_1").Specific
            'Item_1.Item.Visible = False

            Dim txtSalMer As SAPbouiCOM.EditText
            txtSalMer = oForm.Items.Item("Item_8").Specific
            txtSalMer.Item.Enabled = False

            'Dim lnkMer As SAPbouiCOM.LinkedButton
            'lnkMer = oForm.Items.Item("lnkMer").Specific
            'lnkMer.LinkedObject = 59
            'lnkMer.Item.Visible = False

            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mtxDetalle").Specific

            mMatrix.AddRow()
            For i As Integer = 1 To mMatrix.RowCount
                mMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                mMatrix.Columns.Item("COL").DisplayDesc = True
            Next


            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Utilitario.Util_Log.Escribir_Log("ex CreaFormulario_frmChequeP  " + ex.Message.ToString(), "frmTransEntreCompanias")
        End Try

    End Sub

    Public Sub CreaFormularioExistente_frmTransEntreCompanias(Salida As SalidaCabecera)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        If RecorreFormulario(rsboApp, "frmTransEntreCompanias") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmTransEntreCompanias.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmTransEntreCompanias").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmTransEntreCompanias")

            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"

            Dim txtFocus As SAPbouiCOM.EditText
            txtFocus = oForm.Items.Item("txtFocus").Specific
            oForm.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Dim txtFecha As SAPbouiCOM.EditText
            txtFecha = oForm.Items.Item("txtFecha").Specific
            txtFecha.Item.Enabled = False
            txtFecha.Value = Salida.Fecha.ToString("dd/MM/yyyy")


            Dim txtBO As SAPbouiCOM.EditText
            txtBO = oForm.Items.Item("txtBO").Specific
            txtBO.Item.Enabled = False
            txtBO.Value = rCompany.CompanyName

            EditText = oForm.Items.Item("Item_8").Specific
            EditText.Item.Enabled = False
            EditText.Value = Salida.DocEntrySalida

            EditText = oForm.Items.Item("txtAlmO").Specific
            EditText.Item.Enabled = False
            EditText.Value = Salida.AlmacenOrigen

            Dim qryNOmbreBod As String = ""
            Dim _qryNOmbreBod As String = ""

            qryNOmbreBod = "Select ""WhsName"" from OWHS where ""WhsCode""='" + Salida.AlmacenDestino + "'"
            _qryNOmbreBod = oFuncionesB1.getRSvalue(qryNOmbreBod, "WhsName", "")

            Label = oForm.Items.Item("Item_4").Specific
            Label.Caption = _qryNOmbreBod

            EditText = oForm.Items.Item("txtAlmacen").Specific
            EditText.Item.Enabled = False
            EditText.Value = Salida.AlmacenDestino

            qryNOmbreBod = "Select ""WhsName"" from OWHS where ""WhsCode""='" + Salida.AlmacenOrigen + "'"
            _qryNOmbreBod = oFuncionesB1.getRSvalue(qryNOmbreBod, "WhsName", "")

            Label = oForm.Items.Item("Item_9").Specific
            Label.Caption = _qryNOmbreBod

            Dim cbcBases As SAPbouiCOM.ComboBox
            cbcBases = oForm.Items.Item("cbcBases").Specific
            cbcBases.ValidValues.Add(Salida.EmpresaDestino, Salida.EmpresaDestino)
            cbcBases.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'cbcBases.Item.Enabled = False


            EditText = oForm.Items.Item("txtComen").Specific
            EditText.Item.Enabled = False
            EditText.Value = Salida.comentario

            'detalle carga matrix
            Dim oItemCode As SAPbouiCOM.EditText
            Dim oItemName As SAPbouiCOM.EditText
            Dim oCantidad As SAPbouiCOM.EditText
            Dim oPrecio As SAPbouiCOM.EditText

            Dim oMatrix As SAPbouiCOM.Matrix
            oMatrix = oForm.Items.Item("mtxDetalle").Specific

            Dim source1 As SAPbouiCOM.UserDataSource = oForm.DataSources.UserDataSources.Item("UD_3")
            Dim source2 As SAPbouiCOM.UserDataSource = oForm.DataSources.UserDataSources.Item("UD_4")
            Dim source3 As SAPbouiCOM.UserDataSource = oForm.DataSources.UserDataSources.Item("UD_5")
            Dim source4 As SAPbouiCOM.UserDataSource = oForm.DataSources.UserDataSources.Item("UD_6")

            'Dim columnID As String = "Col_4" ' ID de la columna
            'Dim columnTitle As String = "Column Title" ' Título de la columna
            'Dim columnType As SAPbouiCOM.BoFormItemTypes = SAPbouiCOM.BoFormItemTypes.it_EDIT ' Tipo de datos de la columna


            'oMatrix.Columns.Add(columnID, columnType)

            Dim oUserDataSrc As SAPbouiCOM.UserDataSource = oForm.DataSources.UserDataSources.Add("UD_7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            Dim oColumn As SAPbouiCOM.Column = oMatrix.Columns.Add("Col_4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.DataBind.SetBound(True, "", "UD_7")
            oColumn.TitleObject.Caption = "Costo"
            oColumn.Width = 100
            Dim linea As Integer = 1

            If Salida.Detalles.Count > 0 Then

                For Each X As SalidaDetalle In Salida.Detalles

                    'funciona
                    oMatrix.AddRow()


                    oMatrix.Columns.Item("COL").Cells.Item(linea).Specific.String = linea
                    oMatrix.Columns.Item("COL").DisplayDesc = True

                    oMatrix.Columns.Item("Col_0").Cells.Item(linea).Specific.String = X.Codigo
                    oMatrix.Columns.Item("Col_0").DisplayDesc = True


                    oMatrix.Columns.Item("Col_1").Cells.Item(linea).Specific.String = X.Nombre
                    oMatrix.Columns.Item("Col_1").DisplayDesc = True


                    oMatrix.Columns.Item("Col_2").Cells.Item(linea).Specific.String = X.Cantidad
                    oMatrix.Columns.Item("Col_2").DisplayDesc = True


                    oMatrix.Columns.Item("Col_3").Cells.Item(linea).Specific.String = formatDecimal(X.Precio.ToString)
                    oMatrix.Columns.Item("Col_3").DisplayDesc = True
                    oForm.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oMatrix.Columns.Item("Col_3").Visible = False

                    oMatrix.Columns.Item("Col_4").Cells.Item(linea).Specific.String = X.Precio.ToString
                    oMatrix.Columns.Item("Col_4").DisplayDesc = True

                    '-----------------------------------------------------
                    oMatrix.FlushToDataSource()
                    linea += 1
                Next
                '
            End If
            'oMatrix.FlushToDataSource()
            oForm.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'For i As Integer = 1 To oMatrix.RowCount
            oMatrix.Columns.Item("Col_0").Editable = False
            oMatrix.Columns.Item("Col_1").Editable = False
            oMatrix.Columns.Item("Col_2").Editable = False
            oMatrix.Columns.Item("Col_3").Editable = False
            oMatrix.Columns.Item("Col_4").Editable = False
            'Next

            Dim lnkMer As SAPbouiCOM.LinkedButton
            lnkMer = oForm.Items.Item("Item_10").Specific
            lnkMer.LinkedObject = 60

            oForm.Items.Item("cbcBases").Enabled = False

            oForm.Items.Item("btnVer").Left = oForm.Items.Item("2").Left
            Dim btnVer As SAPbouiCOM.Button
            btnVer = oForm.Items.Item("btnVer").Specific
            btnVer.Item.Visible = True

            oForm.Items.Item("btnProcesa").Visible = False
            oForm.Items.Item("2").Left = oForm.Items.Item("btnProcesa").Left
            Dim oB As SAPbouiCOM.Button
            oB = oForm.Items.Item("2").Specific
            oB.Caption = "OK"



            oForm.Visible = True
            oForm.Select()



            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Utilitario.Util_Log.Escribir_Log("ex CreaFormulario_frmChequeP  " + ex.Message.ToString(), "frmTransEntreCompanias")
        End Try

    End Sub


    Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent
        Try


            If eventInfo.FormUID = "frmTransEntreCompanias" Then

                If eventInfo.ItemUID = "mtxDetalle" Then

                    If eventInfo.ColUID = "COL" Then
                        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                        Dim oMenus As SAPbouiCOM.Menus = Nothing

                        If eventInfo.BeforeAction = True Then

                            Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

                            'If mForm.Mode = BoFormMode.fm_ADD_MODE Or mForm.Mode = BoFormMode.fm_OK_MODE Then

                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
                            Try
                                _num = eventInfo.Row

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
                            'End If
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
                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent

        Try
            Dim typeExx, idFormm As String
            typeExx = oFuncionesB1.FormularioActivo(idFormm)

            If typeExx = "frmTransEntreCompanias" Then
                If pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
                    rsboApp.Forms.ActiveForm.Freeze(True)
                    Try
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("mtxDetalle").Specific
                        mMatrix.AddRow()
                        For i As Integer = 1 To mMatrix.RowCount
                            mMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                            mMatrix.Columns.Item("COL").DisplayDesc = True
                        Next


                        'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                        '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                        'End If
                    Catch ex As Exception
                    Finally
                        rsboApp.Forms.ActiveForm.Freeze(False)
                    End Try

                End If

                If pVal.MenuUID = "Eliminar" And pVal.BeforeAction = False Then
                    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("mtxDetalle").Specific
                    mMatrix.DeleteRow(_num)
                    For i As Integer = 1 To mMatrix.RowCount
                        mMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                    Next

                    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                    End If
                End If

                If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                    'oForm = rsboApp.Forms.Item("frmTransEntreCompanias")
                    oForm.Close()
                    CreaFormulario_frmTransEntreCompanias()


                End If

            End If

            'If rsboApp.Forms.ActiveForm.UniqueID = "frmTransEntreCompanias" And pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
            '    rsboApp.Forms.ActiveForm.Freeze(True)
            '    Try
            '        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("mtxDetalle").Specific
            '        mMatrix.AddRow()
            '        For i As Integer = 1 To mMatrix.RowCount
            '            mMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
            '            mMatrix.Columns.Item("COL").DisplayDesc = True
            '        Next


            '        'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '        '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '        'End If
            '    Catch ex As Exception
            '    Finally
            '        rsboApp.Forms.ActiveForm.Freeze(False)
            '    End Try

            'End If

            'If rsboApp.Forms.ActiveForm.UniqueID = "frmTransEntreCompanias" And pVal.MenuUID = "Eliminar" And pVal.BeforeAction = False Then
            '    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("mtxDetalle").Specific
            '    mMatrix.DeleteRow(_num)
            '    For i As Integer = 1 To mMatrix.RowCount
            '        mMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
            '    Next

            '    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '    End If
            'End If

            'If pVal.MenuUID = "1282" And rsboApp.Forms.ActiveForm.UniqueID = "frmTransEntreCompanias" And pVal.BeforeAction = False Then

            '    'oForm = rsboApp.Forms.Item("frmTransEntreCompanias")
            '    oForm.Close()
            '    CreaFormulario_frmTransEntreCompanias()


            'End If


        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmTransEntreCompañias")

        End Try
    End Sub


    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        'Try
        If pVal.FormTypeEx = "frmTransEntreCompanias" Then


            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If Not pVal.Before_Action Then

                        Select Case pVal.ItemUID

                            Case "btnProcesa"

                                Try

                                    Dim Result As clsError
                                    'If conectSAPBase2() Then
                                    '    Result = CrearEntradaMercancias()
                                    'End If
                                    oForm.Freeze(True)
                                    rsboApp.StatusBar.SetText(NombreAddon + " - Creando Salida de Mercancias", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Result = CrearSalidaMercancias()
                                    If Result.ErrorEstado = True Then
                                        oForm.Freeze(False)
                                        Throw New System.Exception("Error al crear Salida de Mercancias: " + Result.ErrorDescripcion) 'este seria el ex.message del catch
                                    Else
                                        If conectSAPBase2() Then
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Conectando a la base " + rCompany2.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Result = CrearEntradaMercancias()
                                            If Result.ErrorEstado = False Then
                                                'rsboApp.StatusBar.SetText(NombreAddon + " - Actualizando Id en Salida Mercancias " + rCompany2.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                'If ActualizarSalidaMercancias() Then
                                                '    rsboApp.StatusBar.SetText(NombreAddon + " - Salida Mercancias actualizada " + rCompany2.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)



                                                If rCompany2.InTransaction Then
                                                    Try

                                                        rCompany2.EndTransaction(BoWfTransOpt.wf_Commit)
                                                    Catch ex As Exception
                                                        If Not IsNothing(rCompany) Then
                                                            If rCompany.InTransaction Then
                                                                rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                                                            End If
                                                        End If

                                                        If Not IsNothing(rCompany2) Then
                                                            If rCompany2.InTransaction Then
                                                                rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                                                            End If
                                                        End If
                                                        oForm.Freeze(False)
                                                        Throw New ArgumentException("Error en commit entrada de mercancias: " + ex.Message.ToString())

                                                    End Try

                                                End If

                                                If rCompany.InTransaction Then
                                                    Try

                                                        rCompany.EndTransaction(BoWfTransOpt.wf_Commit)
                                                    Catch ex As Exception
                                                        oForm.Freeze(False)
                                                        Dim error_code As Integer = 0
                                                        Dim error_message As String = ex.Message.ToString()
                                                        Throw New ArgumentException("Error en commit Salida de Mercancias: " + ex.Message.ToString())

                                                    End Try

                                                End If
                                                rsboApp.StatusBar.SetText(NombreAddon + " - Actualizando Id en Salida Mercancias " + rCompany2.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                If ActualizarSalidaMercancias() Then
                                                    rsboApp.StatusBar.SetText(NombreAddon + " - Salida Mercancias actualizada " + rCompany2.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                End If


                                                Dim txtSalMer As SAPbouiCOM.EditText
                                                txtSalMer = oForm.Items.Item("Item_8").Specific
                                                txtSalMer.Value = Functions.VariablesGlobales._SS_IdSalidaMercancias

                                                Dim lnkMer As SAPbouiCOM.LinkedButton
                                                lnkMer = oForm.Items.Item("Item_10").Specific
                                                lnkMer.LinkedObject = 60

                                                oForm.Items.Item("btnProcesa").Visible = False
                                                oForm.Items.Item("2").Left = oForm.Items.Item("btnProcesa").Left
                                                Dim oB As SAPbouiCOM.Button
                                                oB = oForm.Items.Item("2").Specific
                                                oB.Caption = "OK"

                                                Try
                                                    rCompany2.Disconnect()
                                                Catch ex As Exception
                                                    rsboApp.StatusBar.SetText(NombreAddon + " - Error al desconectar conexion rCompany2: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                End Try
                                            Else
                                                oForm.Freeze(False)
                                                Throw New System.Exception("Error al crear Entrada de Mercancias: " + Result.ErrorDescripcion)
                                            End If
                                        Else
                                            Utilitario.Util_Log.Escribir_Log("Hubo un error al conectarse a la base 2, se hizo rollback: ", "frmTransEntreCompanias")

                                            If Not IsNothing(rCompany) Then
                                                If rCompany.InTransaction Then
                                                    rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                                                End If
                                            End If

                                            If Not IsNothing(rCompany2) Then
                                                If rCompany2.InTransaction Then
                                                    rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                                                End If
                                            End If

                                            oForm.Freeze(False)
                                        End If
                                    End If
                                    oForm.Freeze(False)

                                Catch comEx As System.Runtime.InteropServices.COMException

                                    Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Salida de Mercancias, se hizo Rollback comEx: " + comEx.Message.ToString, "frmTransEntreCompanias")
                                    rsboApp.SetStatusBarMessage("Error en Try General" + comEx.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)

                                    If Not IsNothing(rCompany) Then
                                        If rCompany.InTransaction Then
                                            rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                                        End If
                                    End If

                                    If Not IsNothing(rCompany2) Then
                                        If rCompany2.InTransaction Then
                                            rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                                        End If
                                    End If

                                    oForm.Freeze(False)

                                Catch ex As Exception

                                    Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Salida de Mercancias, se hizo Rollback: " + ex.Message.ToString, "frmTransEntreCompanias")
                                    rsboApp.SetStatusBarMessage("Error en Try General" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)

                                    If Not IsNothing(rCompany) Then
                                        If rCompany.InTransaction Then
                                            rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                                        End If
                                    End If

                                    If Not IsNothing(rCompany2) Then
                                        If rCompany2.InTransaction Then
                                            rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                                        End If
                                    End If

                                    oForm.Freeze(False)
                                Finally
                                    Utilitario.Util_Log.Escribir_Log("Finally", "frmTransEntreCompanias")
                                    oForm.Freeze(False)
                                End Try

                            Case "btnVer"

                                Dim rutaReporte As String = ""
                                Dim docentry As SAPbouiCOM.EditText = oForm.Items.Item("Item_8").Specific
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                                    rutaReporte = System.Windows.Forms.Application.StartupPath & "\" + rsboApp.Company.DatabaseName + "_SS_TransferenciaSEHana.rpt"
                                    ofrmChequePD.PresentarPDFNDHANA(docentry.Value, rutaReporte, "Salida", "DOCENTRY")

                                Else
                                    rutaReporte = System.Windows.Forms.Application.StartupPath & "\" + rsboApp.Company.DatabaseName + "_SS_TransferenciaSESQL.rpt"
                                    ofrmChequePD.PresentarPDFNDSQL(docentry.Value, rutaReporte, "Salida", "DOCENTRY")
                                End If



                        End Select
                    End If
                    '                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                    '                        If pVal.BeforeAction Then

                    '                            Event_MatrixLinkPressed(pVal)

                    '                        End If

                    '                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    '                        If pVal.BeforeAction = False And pVal.ItemUID = "oGrid" Then
                    '                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmTransEntreCompanias")
                    '                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                    '                            Dim ofila As Integer = 0
                    '                            ofila = pVal.Row
                    '                            'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                    '                            'oGrid.Rows.SelectedRows.Add(ofila)
                    '                            'For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                    '                            '    ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                    '                            '    ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                    '                            'Next

                    '                            If ofila = -1 Then '' clic en la cabecera
                    '                                Exit Sub
                    '                            End If
                    '                            Dim oCheque As New clsCheque(oDataTable.GetValue("NumPago", ofila).ToString(),
                    '                                                           oDataTable.GetValue("Cheque_Num", ofila).ToString(),
                    '                                                           oDataTable.GetValue("Cheque_Valor", ofila),
                    '                                                           oDataTable.GetValue("Banco", ofila).ToString(),
                    '                                                           oDataTable.GetValue("Cliente_Codigo", ofila).ToString(),
                    '                                                           oDataTable.GetValue("Cliente", ofila).ToString(),
                    '                                                           oDataTable.GetValue("Doc_Protesto", ofila).ToString(),
                    '                                                           oDataTable.GetValue("Pago_Coments", ofila).ToString(),
                    '                                                           oDataTable.GetValue("CuentaContableDeposito", ofila).ToString(),
                    '                                                           oDataTable.GetValue("NombreCuentaContableDeposito", ofila).ToString(),
                    '                                                           oDataTable.GetValue("NumeroDeposito", ofila).ToString())

                    '                            Dim ValorProtesto As Decimal = 0
                    '                            If oCheque.Doc_Protesto <> "0" Then
                    '                                Dim sQuery As String
                    '                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    '                                    sQuery = "SELECT TOP 1 B.""LineTotal"" FROM ""OINV"" A INNER JOIN ""INV1"" B ON A.""DocEntry"" = B.""DocEntry"" WHERE B.""Dscription"" = 'MONTO PROTESTO' AND A.""DocEntry"" =  '" + oCheque.Doc_Protesto + "'"
                    '                                Else
                    '                                    sQuery = "SELECT TOP 1 B.""LineTotal"" FROM ""OINV"" A INNER JOIN ""INV1"" B ON A.""DocEntry"" = B.""DocEntry"" WHERE B.""Dscription"" = 'MONTO PROTESTO' AND A.""DocEntry"" =  '" + oCheque.Doc_Protesto + "'"
                    '                                End If
                    '                                Utilitario.Util_Log.Escribir_Log("Query Obtener Valor del protesto:  " + sQuery, "frmTransEntreCompanias")

                    '                                ValorProtesto = formatDecimal(oFuncionesB1.getRSvalue(sQuery, "LineTotal", ""))

                    '                            End If
                    '                            ofrmChequePD.CreaFormulario_frmChequePD(oCheque, ValorProtesto, ofila)

                    '                        End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal

                    If oCFLEvento.BeforeAction = False Then
                        Try
                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID
                            Dim oForm As SAPbouiCOM.Form
                            oForm = rsboApp.Forms.Item(FormUID)
                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty
                            If Not oDataTable Is Nothing Then
                                Try
                                    val = oDataTable.GetValue(0, 0)
                                    val1 = oDataTable.GetValue(1, 0)
                                Catch ex As Exception
                                End Try

                                Select Case pVal.ItemUID
                                    Case "txtAlmacen"
                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSA2")
                                        oUserDataSource.ValueEx = oDataTable.GetValue("WhsCode", 0)
                                        Dim lbCliente As SAPbouiCOM.StaticText
                                        lbCliente = oForm.Items.Item("Item_9").Specific
                                        lbCliente.Caption = oDataTable.GetValue("WhsName", 0)

                                    Case "txtAlmO"
                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSA")
                                        oUserDataSource.ValueEx = oDataTable.GetValue("WhsCode", 0)
                                        Dim lbCliente As SAPbouiCOM.StaticText
                                        lbCliente = oForm.Items.Item("Item_4").Specific
                                        lbCliente.Caption = oDataTable.GetValue("WhsName", 0)

                                End Select
                            End If
                        Catch ex As Exception
                            rsboApp.MessageBox("et_CHOOSE_FROM_LIST " + ex.Message.ToString())
                            Utilitario.Util_Log.Escribir_Log("ex et_CHOOSE_FROM_LIST  " + ex.Message.ToString(), "frmTransEntreCompanias")
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If Not pVal.Before_Action Then
                        If pVal.ItemUID = "mtxDetalle" And pVal.ColUID = "Col_0" Then '
                            rMatrix = oForm.Items.Item("mtxDetalle").Specific

                            Dim RE = pVal.Row
                            If pVal.Row = 0 Then
                                Exit Sub
                            End If
                            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
                            CodigoProd = rMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific
                            NombreProd = rMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific
                            CostoProd = rMatrix.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific

                            ofrmConsultaDetalleTrans.CargaFormularioConsulta(pVal.Row, CodigoProd, NombreProd, CostoProd)

                        ElseIf pVal.ItemUID = "txtAlmacen" Then
                            conectSAPBase2()
                            EditText = oForm.Items.Item("txtAlmacen").Specific
                            Label = oForm.Items.Item("Item_9").Specific
                            ofrmConsultaBodega.CargaFormularioConsulta(EditText, Label, rCompany2)
                            Try
                                rCompany2.Disconnect()
                            Catch ex As Exception
                                rsboApp.StatusBar.SetText(NombreAddon + " - Error al desconectar conexion rCompany2: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If
                    End If
                    '                End Select
                    '            End If
                    '        Catch ex As Exception

                    '        End Try

                    '    End Sub

                    '    Private Sub ExpandirContraer()

                    '        Dim oGrid As SAPbouiCOM.Grid

                    '        Try
                    '            oGrid = oForm.Items.Add("oGrid", SAPbouiCOM.BoFormItemTypes.it_GRID).Specific
                    '        Catch ex As Exception
                    '            oGrid = oForm.Items.Item("oGrid").Specific
                    '        End Try

                    '        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

                    '        If schk.Checked Then

                    '            oGrid.Rows.CollapseAll()
                    '        Else
                    '            oGrid.Rows.ExpandAll()

                    '        End If

                    '        oGrid.AutoResizeColumns()

                    '    End Sub

                    '    Private Sub Event_MatrixLinkPressed(ByVal pVal As SAPbouiCOM.ItemEvent)

                    '        If pVal.FormTypeEx = "frmTransEntreCompanias" Then

                    '            Select Case pVal.ItemUID

                    '                Case "oGrid"

                    '                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                    '                    Dim oObjType As String = oGrid.DataTable.GetValue("ObjType", oGrid.GetDataTableRowIndex(pVal.Row))
                    '                    Dim oColumns As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("DocEntry")

                    '                    Select Case oObjType

                    '                        Case 13
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oInvoices

                    '                        Case 203
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oDownPayments

                    '                        Case 14
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oCreditNotes

                    '                        Case 18
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices

                    '                        Case 204
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments

                    '                        Case 15
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oDeliveryNotes

                    '                        Case 67
                    '                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer

                    '                        Case Else
                    '                            Exit Sub

                    '                    End Select

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


    '    Public Shared Function FechaSql(ByVal fecha As DateTime) As String

    '        Dim anio As String = fecha.Year
    '        Dim mes As String = fecha.Month
    '        Dim dia As String = fecha.Day

    '        If anio.Length = 2 Then
    '            anio = "20" & anio
    '        End If

    '        Return "{d'" & anio & "-" & mes.PadLeft(2, "0") & "-" & dia.PadLeft(2, "0") & "'}"

    '    End Function

    '    Public Shared Function formatDecimal(ByVal numero As String) As Decimal

    '        Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
    '        Dim result As Double = 0
    '        Try
    '            If numero = "" Then
    '                numero = "0"
    '            End If
    '            If numero IsNot Nothing Then
    '                If Not numero.Contains(",") Then
    '                    result = Double.Parse(numero, CultureInfo.InvariantCulture)
    '                Else
    '                    result = Convert.ToDouble(numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
    '                    'result = Double.Parse((numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString())), CultureInfo.InvariantCulture)
    '                End If
    '            End If
    '        Catch e As Exception
    '            Try
    '                'result = Convert.ToDouble(numero)
    '                result = Double.Parse(numero, CultureInfo.InvariantCulture)
    '            Catch
    '                Try
    '                    'result = Convert.ToDouble(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."))
    '                    result = Double.Parse(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."), CultureInfo.InvariantCulture)
    '                Catch
    '                    Throw New Exception("Wrong string-to-double format")
    '                End Try
    '            End Try
    '        End Try
    '        Return result

    '    End Function
    Public Sub CargaComboboxSAP(ByVal Consulta As String, ByRef Combobox As SAPbouiCOM.ComboBox, ByVal Campo As String, ByVal Descripcion As String, ByVal RecordSet As SAPbobsCOM.Recordset)

        Dim oValidValues As SAPbouiCOM.ValidValues

        Utilitario.Util_Log.Escribir_Log("Query Consulta Combo: " + Consulta, "frmParametrosAddOnLE")

        RecordSet.DoQuery(Consulta)

        oValidValues = Combobox.ValidValues

        While Combobox.ValidValues.Count > 0
            Combobox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End While

        If Descripcion.Equals(String.Empty) Then
            While Not RecordSet.EoF
                oValidValues.Add(RecordSet.Fields.Item(Campo).Value, Descripcion)
                RecordSet.MoveNext()
            End While
        Else
            While Not RecordSet.EoF
                oValidValues.Add(RecordSet.Fields.Item(Campo).Value, RecordSet.Fields.Item(Descripcion).Value)
                RecordSet.MoveNext()
            End While
        End If

    End Sub

    Private Function CrearSalidaMercancias() As clsError

        'System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator = "."
        'System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator = ","

        Dim oSalida As SAPbobsCOM.Documents = Nothing
        Dim CodBodSalida As SAPbouiCOM.EditText = oForm.Items.Item("txtAlmO").Specific
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim oClsError As New clsError

        Dim txtCodigo As SAPbouiCOM.EditText
        Dim txtNombre As SAPbouiCOM.EditText
        Dim txtCantidad As SAPbouiCOM.EditText
        Dim txtCosto As SAPbouiCOM.EditText

        Dim cbcBases As SAPbouiCOM.ComboBox
        cbcBases = oForm.Items.Item("cbcBases").Specific

        Dim bodegaEntrada As SAPbouiCOM.EditText
        bodegaEntrada = oForm.Items.Item("txtAlmacen").Specific

        Functions.VariablesGlobales._SS_IdSalidaMercancias = ""

        Try
            If Not rCompany.InTransaction Then

                rCompany.StartTransaction()
                Utilitario.Util_Log.Escribir_Log("Transaccion Iniciada crear salida de mercancias", "frmTransEntreCompanias")
            End If

            oSalida = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            oSalida.DocObjectCode = BoObjectTypes.oInventoryGenExit
            oSalida.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            Dim FECHA As String = oForm.Items.Item("txtFecha").Specific.value.ToString()

            oSalida.DocDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now
            oSalida.TaxDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now
            Utilitario.Util_Log.Escribir_Log("Fechas asignadas a la salida", "frmTransEntreCompanias")
            Dim oMatrix = oForm.Items.Item("mtxDetalle").Specific

            If oMatrix.VisualRowCount() = 0 Then
                oClsError.ErrorEstado = True
                Return oClsError
            End If

            For i As Integer = 1 To oMatrix.VisualRowCount()

                txtCodigo = oMatrix.GetCellSpecific("Col_0", i)

                If txtCodigo.Value <> "" Then

                    Utilitario.Util_Log.Escribir_Log("Salida Codigo detalle : " + txtCodigo.Value.ToString, "frmTransEntreCompanias")

                    txtNombre = oMatrix.GetCellSpecific("Col_1", i)
                    txtCantidad = oMatrix.GetCellSpecific("Col_2", i)
                    txtCosto = oMatrix.GetCellSpecific("Col_3", i)

                    Utilitario.Util_Log.Escribir_Log("Salida Codigo detalle : " + txtCodigo.Value.ToString, "frmTransEntreCompanias")
                    Utilitario.Util_Log.Escribir_Log("Salida descripcion detalle : " + txtNombre.Value.ToString, "frmTransEntreCompanias")
                    Utilitario.Util_Log.Escribir_Log("Salida cantidad detalle : " + txtCantidad.Value.ToString, "frmTransEntreCompanias")
                    Utilitario.Util_Log.Escribir_Log("Salida costo detalle : " + txtCosto.Value.ToString, "frmTransEntreCompanias")

                    Dim costo As String = CDec(txtCosto.Value).ToString("###0.0000")
                    Dim cantidad As String = CDec(txtCantidad.Value).ToString("###0.0000")

                    oSalida.Lines.ItemCode = txtCodigo.Value
                    oSalida.Lines.ItemDescription = txtNombre.Value
                    oSalida.Lines.Quantity = CDbl(txtCantidad.Value)
                    oSalida.Lines.Price = CDbl(txtCosto.Value)

                    oSalida.Lines.WarehouseCode = CodBodSalida.Value
                    Utilitario.Util_Log.Escribir_Log("Salida costo detalle : " + CodBodSalida.Value.ToString, "frmTransEntreCompanias")
                    Dim qryCuenta As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        qryCuenta = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + Functions.VariablesGlobales._SS_CuentaTransferencia + "'"
                    Else
                        qryCuenta = "Select AcctCode from OACT Where FormatCode =  '" + Functions.VariablesGlobales._SS_CuentaTransferencia + "'"
                    End If
                    Dim Cuenta As String = oFuncionesB1.getRSvalue(qryCuenta, "AcctCode", "")

                    oSalida.Lines.AccountCode = Cuenta
                    Utilitario.Util_Log.Escribir_Log("Salida Cuenta detalle : " + Cuenta.ToString, "frmTransEntreCompanias")

                    oSalida.Lines.Add()
                    Utilitario.Util_Log.Escribir_Log("Salida linea agregada", "frmTransEntreCompanias")
                End If

            Next

            Dim Commentario As SAPbouiCOM.EditText = oForm.Items.Item("txtComen").Specific
            oSalida.Comments = Commentario.Value.ToString
            Utilitario.Util_Log.Escribir_Log("Salida comentario agregado" + Commentario.Value.ToString, "frmTransEntreCompanias")
            oSalida.UserFields.Fields.Item("U_SS_BaseDestino").Value = cbcBases.Value.ToString
            oSalida.UserFields.Fields.Item("U_SS_IdBodegaEnt").Value = bodegaEntrada.Value.ToString
            Utilitario.Util_Log.Escribir_Log("Salida linea agregada", "frmTransEntreCompanias")

            RetVal = oSalida.Add()
            Utilitario.Util_Log.Escribir_Log("RetVal Salida" + RetVal.ToString, "frmTransEntreCompanias")
            If RetVal <> 0 Then
                rCompany.GetLastError(ErrCode, ErrMsg)
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oClsError.ErrorEstado = True
                oClsError.ErrorDescripcion = "Error: " + ErrCode.ToString + " - " + ErrMsg
                Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Salida de Mercancias, se hizo Rollback - Error: " + ErrCode.ToString + ": " + ErrMsg, "frmTransEntreCompanias")
                If Not IsNothing(rCompany) Then
                    If rCompany.InTransaction Then
                        rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                    End If
                End If
                Return oClsError

            Else
                rCompany.GetNewObjectCode(Functions.VariablesGlobales._SS_IdSalidaMercancias)
                oClsError.ErrorEstado = False
                Utilitario.Util_Log.Escribir_Log("Id Salida Mercancias" + Functions.VariablesGlobales._SS_IdSalidaMercancias, "frmTransEntreCompanias")
                'If rCompany.InTransaction Then
                '    Try

                '        rCompany.EndTransaction(BoWfTransOpt.wf_Commit)
                '    Catch ex As Exception
                '        Dim error_code As Integer = 0
                '        Dim error_message As String = ex.Message.ToString()
                '        Throw New ArgumentException("Error en commit Salida de Mercancias: " + ex.Message.ToString())

                '    End Try

                'End If
                Utilitario.Util_Log.Escribir_Log("No genero error en la creacion de Salida", "frmTransEntreCompanias")
                Return oClsError
            End If

        Catch comEx As System.Runtime.InteropServices.COMException
            Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Salida de Mercancias, se hizo Rollback: " + comEx.Message.ToString, "frmTransEntreCompanias")
            rsboApp.SetStatusBarMessage("Error  System.Runtime: " + comEx.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oClsError.ErrorEstado = True
            oClsError.ErrorDescripcion = comEx.Message
            If Not IsNothing(rCompany) Then
                If rCompany.InTransaction Then

                    rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                End If
            End If
            Return oClsError

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Salida de Mercancias, se hizo Rollback: " + ex.Message.ToString, "frmTransEntreCompanias")
            rsboApp.SetStatusBarMessage("Error Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oClsError.ErrorEstado = True
            oClsError.ErrorDescripcion = ex.Message
            If Not IsNothing(rCompany) Then
                If rCompany.InTransaction Then

                    rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                End If
            End If
            Return oClsError
        End Try


    End Function

    Private Function CrearEntradaMercancias() As clsError

        'System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator = "."
        'System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator = ","

        Dim oEntrada As SAPbobsCOM.Documents = Nothing
        Dim CodBodEntrada As SAPbouiCOM.EditText = oForm.Items.Item("txtAlmacen").Specific
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim oClsError As New clsError

        Functions.VariablesGlobales._SS_IdEntradaMercancias = ""
        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Entrada de Mercancias", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Dim BaseOrigen As SAPbouiCOM.EditText
        BaseOrigen = oForm.Items.Item("txtBO").Specific

        Dim bodegaSalida As SAPbouiCOM.EditText
        bodegaSalida = oForm.Items.Item("txtAlmO").Specific

        Try
            If Not rCompany2.InTransaction Then

                rCompany2.StartTransaction()
                Utilitario.Util_Log.Escribir_Log("Transaccion Iniciada entradad de mercancias", "frmTransEntreCompanias")
            End If

            oEntrada = rCompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
            oEntrada.DocObjectCode = BoObjectTypes.oInventoryGenEntry
            oEntrada.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            Dim FECHA As String = oForm.Items.Item("txtFecha").Specific.value.ToString()

            oEntrada.DocDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now
            oEntrada.TaxDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now

            Utilitario.Util_Log.Escribir_Log("Fechas agregadas a la entrada", "frmTransEntreCompanias")

            Dim oMatrix = oForm.Items.Item("mtxDetalle").Specific

            For i As Integer = 1 To oMatrix.VisualRowCount()

                Dim txtCodigo As SAPbouiCOM.EditText = oMatrix.GetCellSpecific("Col_0", i)
                Dim txtNombre As SAPbouiCOM.EditText = oMatrix.GetCellSpecific("Col_1", i)
                Dim txtCantidad As SAPbouiCOM.EditText = oMatrix.GetCellSpecific("Col_2", i)
                Dim txtCosto As SAPbouiCOM.EditText = oMatrix.GetCellSpecific("Col_3", i)

                If txtCodigo.Value <> "" Then
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle codigo: " + txtCodigo.Value.ToString, "frmTransEntreCompanias")

                    oEntrada.Lines.ItemCode = txtCodigo.Value
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle codigo: " + txtCodigo.Value.ToString, "frmTransEntreCompanias")
                    oEntrada.Lines.ItemDescription = txtNombre.Value
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle nombre: " + txtNombre.Value.ToString, "frmTransEntreCompanias")
                    oEntrada.Lines.Quantity = CDbl(txtCantidad.Value)
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle cantidad: " + txtCantidad.Value.ToString, "frmTransEntreCompanias")
                    oEntrada.Lines.UnitPrice = CDbl(txtCosto.Value)
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle costo: " + txtCosto.Value.ToString, "frmTransEntreCompanias")
                    'oEntrada.Lines.Price = CDbl(txtCosto.Value)
                    oEntrada.Lines.WarehouseCode = CodBodEntrada.Value
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle bodega: " + CodBodEntrada.Value.ToString, "frmTransEntreCompanias")

                    Dim qryCuenta As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        qryCuenta = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + Functions.VariablesGlobales._SS_CuentaTransferencia + "'"
                    Else
                        qryCuenta = "Select AcctCode from OACT Where FormatCode =  '" + Functions.VariablesGlobales._SS_CuentaTransferencia + "'"
                    End If
                    Dim Cuenta As String = getRSvalue(qryCuenta, "AcctCode", "")

                    oEntrada.Lines.AccountCode = Cuenta
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle Cuenta: " + CodBodEntrada.Value.ToString, "frmTransEntreCompanias")
                    oEntrada.Lines.Add()
                    Utilitario.Util_Log.Escribir_Log("Entrada detalle agregado", "frmTransEntreCompanias")
                End If

            Next

            oEntrada.UserFields.Fields.Item("U_SS_IdSalMer").Value = Functions.VariablesGlobales._SS_IdSalidaMercancias
            Utilitario.Util_Log.Escribir_Log("Entrada IdSalidaMercancias: " + Functions.VariablesGlobales._SS_IdSalidaMercancias, "frmTransEntreCompanias")
            oEntrada.UserFields.Fields.Item("U_SS_EntTrans").Value = "SI"
            oEntrada.UserFields.Fields.Item("U_SS_BaseOrigen").Value = BaseOrigen.Value
            Utilitario.Util_Log.Escribir_Log("Entrada BaseOrigen: " + BaseOrigen.Value.ToString, "frmTransEntreCompanias")
            oEntrada.UserFields.Fields.Item("U_SS_IdBodegaSal").Value = bodegaSalida.Value
            Utilitario.Util_Log.Escribir_Log("Entrada bodegaSalida: " + bodegaSalida.Value.ToString, "frmTransEntreCompanias")


            Dim Commentario As SAPbouiCOM.EditText = oForm.Items.Item("txtComen").Specific
            oEntrada.Comments = Commentario.Value.ToString

            Utilitario.Util_Log.Escribir_Log("Entrada comentario: " + Commentario.Value.ToString, "frmTransEntreCompanias")

            'oEntrada.Comments = "PRUEBA TRASLADOS ENTRE COMPAÑIAS"
            RetVal = oEntrada.Add()
            Utilitario.Util_Log.Escribir_Log("Entrada RetVal: " + RetVal.ToString, "frmTransEntreCompanias")
            If RetVal <> 0 Then
                rCompany2.GetLastError(ErrCode, ErrMsg)
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oClsError.ErrorEstado = True
                oClsError.ErrorDescripcion = "Error: " + ErrCode.ToString + " - " + ErrMsg
                Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Entrada de Mercancias, se hizo Rollback Error: " + ErrCode.ToString + " - " + ErrMsg, "frmTransEntreCompanias")
                If Not IsNothing(rCompany2) Then
                    If rCompany2.InTransaction Then

                        rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                    End If
                End If

                If Not IsNothing(rCompany) Then
                    If rCompany.InTransaction Then
                        Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la Entrada de Mercancias, se hizo Rollback compañia origen Error: " + ErrCode.ToString + " - " + ErrMsg, "frmTransEntreCompanias")
                        rCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                    End If
                End If
                Return oClsError



            Else
                rsboApp.StatusBar.SetText(NombreAddon + " - Entrada de Mercancias creada con éxito..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                rCompany2.GetNewObjectCode(Functions.VariablesGlobales._SS_IdEntradaMercancias)
                Utilitario.Util_Log.Escribir_Log("Id Entrada de mercancias: " + Functions.VariablesGlobales._SS_IdEntradaMercancias, "frmTransEntreCompanias")
                oClsError.ErrorEstado = False

                'If rCompany2.InTransaction Then
                '    Try

                '        rCompany2.EndTransaction(BoWfTransOpt.wf_Commit)
                '    Catch ex As Exception
                '        Dim error_code As Integer = 0
                '        Dim error_message As String = ex.Message.ToString()
                '        Throw New ArgumentException("Error en commit entrada de mercancias: " + ex.Message.ToString())

                '    End Try

                'End If

                Return oClsError
            End If

        Catch comEx As System.Runtime.InteropServices.COMException
            rsboApp.SetStatusBarMessage("Error  System.Runtime entrada de mercancias: " + comEx.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oClsError.ErrorEstado = True
            oClsError.ErrorDescripcion = comEx.Message
            If Not IsNothing(rCompany2) Then
                If rCompany2.InTransaction Then
                    Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la entrada de Mercancias, se hizo Rollback: " + comEx.Message.ToString, "frmTransEntreCompanias")
                    rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                End If
            End If
            Return oClsError

        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Error Exception crear entrada de mercancias: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oClsError.ErrorEstado = True
            oClsError.ErrorDescripcion = ex.Message
            If Not IsNothing(rCompany2) Then
                If rCompany2.InTransaction Then
                    Utilitario.Util_Log.Escribir_Log("Hubo un error al crear la entrada de Mercancias, se hizo Rollback: " + ex.Message.ToString, "frmTransEntreCompanias")
                    rCompany2.EndTransaction(BoWfTransOpt.wf_RollBack)
                End If
            End If
            Return oClsError
        End Try


    End Function

    Private Function conectSAPBase2() As Boolean
        Try
            Dim ErrCode As Long
            Dim ErrMsg As String = ""

            'rCompany2 = New SAPbobsCOM.Company()
            If IsNothing(rCompany2) Then
                rCompany2 = New SAPbobsCOM.Company()
                Utilitario.Util_Log.Escribir_Log("Inicializado rCompany2", "frmTransEntreCompanias")
            End If

            Dim cbcBases As SAPbouiCOM.ComboBox
            cbcBases = oForm.Items.Item("cbcBases").Specific

            Utilitario.Util_Log.Escribir_Log("Iniciando Conexion a la base " + cbcBases.Value.ToString, "frmTransEntreCompanias")

            'Dim TipoServer As String = ""
            'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '    TipoServer = "9"
            'Else
            '    TipoServer = "15"
            'End If

            rCompany2.DbServerType = rCompany.DbServerType
            rCompany2.UseTrusted = False
            rCompany2.CompanyDB = cbcBases.Value
            rCompany2.UserName = Functions.VariablesGlobales._SS_UserSAP
            rCompany2.Password = Functions.VariablesGlobales._SS_PassSAP

            Try
                If Functions.VariablesGlobales._ipServer = "" Then
                    rCompany2.Server = Functions.VariablesGlobales._ipServer
                    rCompany2.LicenseServer = rCompany.Server
                    rCompany2.DbUserName = Functions.VariablesGlobales._SS_UserDB
                    rCompany2.DbPassword = Functions.VariablesGlobales._SS_PassDB
                Else
                    rCompany2.Server = rCompany.Server
                End If
            Catch ex As Exception
                'GuardaLog(ex.Message)
            End Try

            Utilitario.Util_Log.Escribir_Log("DbServerType" + rCompany.DbServerType.ToString(), "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("CompanyDB" + cbcBases.Value.ToString, "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("UserName" + Functions.VariablesGlobales._SS_UserSAP, "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("Password" + Functions.VariablesGlobales._SS_PassSAP, "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("Server" + Functions.VariablesGlobales._ipServer, "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("LicenseServer" + rCompany.Server, "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("DbUserName" + Functions.VariablesGlobales._SS_UserDB, "frmTransEntreCompanias")
            Utilitario.Util_Log.Escribir_Log("DbPassword" + Functions.VariablesGlobales._SS_PassDB, "frmTransEntreCompanias")

            'GuardaLog("DevServerType2 " + ConfigurationManager.AppSettings("DevServerType2").ToString())
            'GuardaLog("UseTrusted2 " + ConfigurationManager.AppSettings("UseTrusted2").ToString())
            'GuardaLog("DevDatabase2 " + ConfigurationManager.AppSettings("DevDatabase2").ToString())
            'GuardaLog("DevSBOUser2 " + ConfigurationManager.AppSettings("DevSBOUser2").ToString())
            'GuardaLog("DevSBOPassword2 " + ConfigurationManager.AppSettings("DevSBOPassword2").ToString())
            'GuardaLog("DevServer2 " + ConfigurationManager.AppSettings("DevServer2").ToString())

            If rCompany2.Connected Then
                'GuardaLog("Company en estado Conectado a SAP BO " + oCompanyBase2.CompanyDB.ToString())
                Utilitario.Util_Log.Escribir_Log("Company en estado Conectado a SAP BO " + rCompany2.CompanyDB, "frmTransEntreCompanias")
                Return True
            End If



            ErrCode = rCompany2.Connect()

            If ErrCode <> 0 Then
                rCompany2.GetLastError(ErrCode, ErrMsg)
                'GuardaLog("Error al conectarse a SAP BASE 2 ,funcion : oCompany.Connect :" + ErrCode.ToString() + " - " + ErrMsg.ToString)
                Utilitario.Util_Log.Escribir_Log("Error al conectarse a SAP BASE 2 ,funcion : oCompany.Connect : " + ErrCode.ToString() + " - " + ErrMsg.ToString + " NombreUsuario: " + rCompany2.UserName, "frmTransEntreCompanias")
                rsboApp.SetStatusBarMessage("Error al conectarse a la Base 2, intente el proceso nuevamente" + ErrCode.ToString() + " - " + ErrMsg.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            Else
                ' GuardaLog("Conectado a SAP BO" + oCompany.CompanyDB.ToString())
                Utilitario.Util_Log.Escribir_Log("Conectado a SAP BO " + rCompany2.CompanyDB.ToString(), "frmTransEntreCompanias")
                Return True
            End If

        Catch ex As Exception
            'GuardaLog("Error al conectarse a SAP , funcion: conectSAP , EX :" + ex.Message.ToString())
            Utilitario.Util_Log.Escribir_Log("Error al conectarse a SAP , funcion: conectSAP , EX : " + ex.Message.ToString(), "frmTransEntreCompanias")
            Return False
            Exit Function
        End Try

    End Function

    Private Function ActualizarSalidaMercancias() As Boolean


        Dim oSalida As SAPbobsCOM.Documents = Nothing
        Dim CodBodSalida As SAPbouiCOM.EditText = oForm.Items.Item("txtAlmO").Specific
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String

        Try

            oSalida = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            oSalida.DocObjectCode = BoObjectTypes.oInventoryGenExit
            oSalida.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None

            oSalida.GetByKey(Functions.VariablesGlobales._SS_IdSalidaMercancias)

            oSalida.UserFields.Fields.Item("U_SS_IdEntMer").Value = Functions.VariablesGlobales._SS_IdEntradaMercancias
            oSalida.UserFields.Fields.Item("U_SS_SalTrans").Value = "SI"

            RetVal = oSalida.Update
            If RetVal <> 0 Then
                rCompany.GetLastError(ErrCode, ErrMsg)
                rsboApp.SetStatusBarMessage("Error al actualizar Salida de Mercancias: " + ErrCode.ToString + " - " + ErrMsg.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
            Return True

        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Error al actualizar Salida de Mercancias: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try


    End Function

    Public Function getRSvalue(ByVal query As String, ByVal columnaRet As String, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRSvalue Catch:" + ex.Message().ToString() + "-QUERY: " + query, "FuncionesB1")
        End Try
        Return ret
    End Function

    Public Function nzString(ByVal unString As String, Optional ByVal formatoSQL As Boolean = False, Optional ByVal valorSiNulo As String = "") As String
        Try
            If Not IsDBNull(unString) Then
                If formatoSQL Then
                    unString = unString.Replace("'", "' + CHAR(39) + '")
                End If
                valorSiNulo = unString
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("nzString Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
        Return valorSiNulo
    End Function

    Public Sub Release(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Release Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
    End Sub

    Public Function getRecordSet(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = rCompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
        End Try
        Return fRS
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

    Function RegularInput(valor As String) As String



        Dim pa = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator



        If pa = "," Then



            Return valor.Replace(".", pa)



        Else



            Return valor.Replace(",", pa)



        End If



    End Function
End Class

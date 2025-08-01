Imports SAPbobsCOM

Public Class frmParametrosAddonLE
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Private mors As SAPbobsCOM.Recordset = Nothing
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim odt As SAPbouiCOM.DataTable
    Dim Alia As String = ""

    Dim cbxCO As SAPbouiCOM.ComboBox
    Dim cbxINH As SAPbouiCOM.ComboBox
    Dim cbxSeries As SAPbouiCOM.ComboBox

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioParametrosADDON()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmParametrosAddonLE") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmParametrosAddonLE.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmParametrosAddonLE").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmParametrosAddonLE")

            Dim CHK_GLE As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            CHK_GLE = oForm.Items.Item("CHK_GLE").Specific
            oForm.DataSources.UserDataSources.Add("CHK_GLE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            CHK_GLE.ValOn = "Y"
            CHK_GLE.ValOff = "N"
            CHK_GLE.DataBind.SetBound(True, "", "CHK_GLE")

            'add 13092022
            Dim chkDBRPT As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            chkDBRPT = oForm.Items.Item("chkDBRPT").Specific
            oForm.DataSources.UserDataSources.Add("chkDBRPT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkDBRPT.ValOn = "Y"
            chkDBRPT.ValOff = "N"
            chkDBRPT.DataBind.SetBound(True, "", "chkDBRPT")


            'add 07102022
            Dim chkgenflo As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            chkgenflo = oForm.Items.Item("chkgenflo").Specific
            oForm.DataSources.UserDataSources.Add("chkgenflo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkgenflo.ValOn = "Y"
            chkgenflo.ValOff = "N"
            chkgenflo.DataBind.SetBound(True, "", "chkgenflo")


            ' add 14102022
            Dim chkVsdn As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            chkVsdn = oForm.Items.Item("chkVsdn").Specific
            oForm.DataSources.UserDataSources.Add("chkVsdn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkVsdn.ValOn = "Y"
            chkVsdn.ValOff = "N"
            chkVsdn.DataBind.SetBound(True, "", "chkVsdn")

            Dim chkVDocs As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            chkVDocs = oForm.Items.Item("chkVDocs").Specific
            oForm.DataSources.UserDataSources.Add("chkVDocs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkVDocs.ValOn = "Y"
            chkVDocs.ValOff = "N"
            chkVDocs.DataBind.SetBound(True, "", "chkVDocs")


            'add 14072023
            'ver boton impresion
            Dim chkverImp As SAPbouiCOM.CheckBox ' 
            chkverImp = oForm.Items.Item("chkverImp").Specific
            oForm.DataSources.UserDataSources.Add("chkverImp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkverImp.ValOn = "Y"
            chkverImp.ValOff = "N"
            chkverImp.DataBind.SetBound(True, "", "chkverImp")

            'Label que muestra la version tributaria , se lo coloca en Negrita
            'Dim LBVH As SAPbouiCOM.StaticText
            'LBVH = oForm.Items.Item("LBVH").Specific
            'LBVH.Item.TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            'LBVH.Item.FontSize = 13
            'LBVH.Caption = functions.VariablesGlobales._gVersiondelMinisterioDeHacienda

            '' CHOOSE FROM LIST CUENTA 
            'Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            'Dim oCons As SAPbouiCOM.Conditions
            'Dim oCon As SAPbouiCOM.Condition
            'oCFLs = oForm.ChooseFromLists
            'Dim oCFL As SAPbouiCOM.ChooseFromList
            'Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            'oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            'oCFLCreationParams.MultiSelection = False
            'oCFLCreationParams.ObjectType = "1"
            'oCFLCreationParams.UniqueID = "CFL1"
            'oCFL = oCFLs.Add(oCFLCreationParams)
            'oCons = oCFL.GetConditions()

            'oCon = oCons.Add()
            'oCon.Alias = "FormatCode"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            ''oCon.CondVal = "S"
            'oCFL.SetConditions(oCons)

            ''Choose from list, para NC
            'oCFLCreationParams.UniqueID = "CFL2"
            'oCFL = oCFLs.Add(oCFLCreationParams)
            'oCFL.SetConditions(oCons)

            'Dim txtMC As SAPbouiCOM.EditText
            'txtMC = oForm.Items.Item("txtMC").Specific
            'oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'txtMC.DataBind.SetBound(True, "", "EditDS")
            'txtMC.ChooseFromListUID = "CFL1"
            'txtMC.ChooseFromListAlias = "FormatCode"

            'Dim txtMP As SAPbouiCOM.EditText
            'txtMP = oForm.Items.Item("txtMP").Specific
            'oForm.DataSources.UserDataSources.Add("EditDST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'txtMP.DataBind.SetBound(True, "", "EditDST")
            'txtMP.ChooseFromListUID = "CFL2"
            'txtMP.ChooseFromListAlias = "FormatCode"

            ' CHOOSE FROM LIST IMPUESTO 
            Dim oCFLsI As SAPbouiCOM.ChooseFromListCollection
            Dim oConsI As SAPbouiCOM.Conditions
            Dim oConI As SAPbouiCOM.Condition
            oCFLsI = oForm.ChooseFromLists
            Dim oCFLI As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParamsI As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParamsI = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParamsI.MultiSelection = False
            oCFLCreationParamsI.ObjectType = "128"
            oCFLCreationParamsI.UniqueID = "CFL1MC"
            oCFLI = oCFLsI.Add(oCFLCreationParamsI)
            oConsI = oCFLI.GetConditions()

            oConI = oConsI.Add()
            oConI.Alias = "Code"
            oConI.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oCon.CondVal = "S"
            oCFLI.SetConditions(oConsI)

            oCFLCreationParamsI.UniqueID = "CFL2MP"
            oCFLI = oCFLsI.Add(oCFLCreationParamsI)
            oCFLI.SetConditions(oConsI)

            Dim txtImpMC As SAPbouiCOM.EditText
            txtImpMC = oForm.Items.Item("txtImpMC").Specific
            oForm.DataSources.UserDataSources.Add("EditDSMC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtImpMC.DataBind.SetBound(True, "", "EditDSMC")
            txtImpMC.ChooseFromListUID = "CFL1MC"
            txtImpMC.ChooseFromListAlias = "Code"

            Dim txtImpMP As SAPbouiCOM.EditText
            txtImpMP = oForm.Items.Item("txtImpMP").Specific
            oForm.DataSources.UserDataSources.Add("EditDSTMP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtImpMP.DataBind.SetBound(True, "", "EditDSTMP")
            txtImpMP.ChooseFromListUID = "CFL2MP"
            txtImpMP.ChooseFromListAlias = "Code"

            Dim CHK_GCHP As SAPbouiCOM.CheckBox ' ActivarOpcionesParaManejoChequesProtestos
            CHK_GCHP = oForm.Items.Item("CHK_GCHP").Specific
            oForm.DataSources.UserDataSources.Add("CHK_GCHP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            CHK_GCHP.ValOn = "Y"
            CHK_GCHP.ValOff = "N"
            CHK_GCHP.DataBind.SetBound(True, "", "CHK_GCHP")

            Dim CHK_GTEB As SAPbouiCOM.CheckBox ' ActivarOpcionesParaTrasnferStockEntreBDs
            CHK_GTEB = oForm.Items.Item("CHK_GTEB").Specific
            oForm.DataSources.UserDataSources.Add("CHK_GTEB", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            CHK_GTEB.ValOn = "Y"
            CHK_GTEB.ValOff = "N"
            CHK_GTEB.DataBind.SetBound(True, "", "CHK_GTEB")

            Dim chkDina As SAPbouiCOM.CheckBox ' Activar Manu dinardap
            chkDina = oForm.Items.Item("chkDina").Specific
            oForm.DataSources.UserDataSources.Add("chkDina", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkDina.ValOn = "Y"
            chkDina.ValOff = "N"
            chkDina.DataBind.SetBound(True, "", "chkDina")

            Dim cbxSeries As SAPbouiCOM.ComboBox
            cbxSeries = oForm.Items.Item("cbSeries").Specific
            Dim oRecordSet As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim filasRecuperadas As Integer

            If rCompany.DbServerType = 9 Then
                CargaComboboxSAP("select ""SeriesName"",""Series"" from " + rCompany.CompanyDB.ToString() + ".""NNM1"" where ""ObjectCode"" = '13' AND ""DocSubType""='DN'", cbxSeries, "SeriesName", "", oRecordSet, filasRecuperadas)
            Else
                CargaComboboxSAP("select SeriesName,Series from NNM1 where ObjectCode = '13' AND DocSubType='DN'", cbxSeries, "SeriesName", "", oRecordSet, filasRecuperadas)
            End If

            If filasRecuperadas > 0 Then

                cbxSeries.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
                cbxSeries.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            End If


            ' CHOOSE FROM LIST CUENTA TRANSFERENCIA 
            Dim oCFLsICT As SAPbouiCOM.ChooseFromListCollection
            Dim oConsICT As SAPbouiCOM.Conditions
            Dim oConICT As SAPbouiCOM.Condition
            oCFLsICT = oForm.ChooseFromLists
            Dim oCFLICT As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParamsICT As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParamsICT = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParamsICT.MultiSelection = False
            oCFLCreationParamsICT.ObjectType = "1"
            oCFLCreationParamsICT.UniqueID = "CFL1MCCT"
            oCFLICT = oCFLsICT.Add(oCFLCreationParamsICT)
            oConsICT = oCFLICT.GetConditions()

            oConICT = oConsICT.Add()
            oConICT.Alias = "FormatCode"
            oConICT.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oCon.CondVal = "S"
            oCFLICT.SetConditions(oConsICT)

            Dim txtCuentaT As SAPbouiCOM.EditText
            txtCuentaT = oForm.Items.Item("txtCuentaT").Specific
            oForm.DataSources.UserDataSources.Add("EditDSCT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCuentaT.DataBind.SetBound(True, "", "EditDSCT")
            txtCuentaT.ChooseFromListUID = "CFL1MCCT"
            txtCuentaT.ChooseFromListAlias = "FormatCode"

            Dim chkSetCU As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            chkSetCU = oForm.Items.Item("chkSetCU").Specific
            oForm.DataSources.UserDataSources.Add("chkSetCU", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkSetCU.ValOn = "Y"
            chkSetCU.ValOff = "N"
            chkSetCU.DataBind.SetBound(True, "", "chkSetCU")

            Dim chkGRDE As SAPbouiCOM.CheckBox 'guias desatendidsas
            chkGRDE = oForm.Items.Item("chkGRDE").Specific
            oForm.DataSources.UserDataSources.Add("chkGRDE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkGRDE.ValOn = "Y"
            chkGRDE.ValOff = "N"
            chkGRDE.DataBind.SetBound(True, "", "chkGRDE")

            Dim chk_PM As SAPbouiCOM.CheckBox 'pagos masivos
            chk_PM = oForm.Items.Item("chk_PM").Specific
            oForm.DataSources.UserDataSources.Add("chk_PM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chk_PM.ValOn = "Y"
            chk_PM.ValOff = "N"
            chk_PM.DataBind.SetBound(True, "", "chk_PM")

            Dim chkActSB As SAPbouiCOM.CheckBox 'Servicios Basicos
            chkActSB = oForm.Items.Item("chkActSB").Specific
            oForm.DataSources.UserDataSources.Add("chkActSB", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkActSB.ValOn = "Y"
            chkActSB.ValOff = "N"
            chkActSB.DataBind.SetBound(True, "", "chkActSB")

            CargaDatos()

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddonLOC + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargaDatos()
        oForm = rsboApp.Forms.Item("frmParametrosAddonLE")
        oForm.Freeze(True)
        Try
            Dim ACTUALIZA As Integer = 0
            ' DATA TABLE CABECERA
            Try
                oForm.DataSources.DataTables.Add("odt")
            Catch ex As Exception
            End Try
            Dim QueryFC As String = ""
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                QueryFC = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryFC += "FROM ""@SS_CONFD"" A INNER JOIN "
                QueryFC += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryFC += " WHERE  B.""U_Modulo"" = '" + NombreAddonLOC + "' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'CONFIGURACION'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = '" + NombreAddonLOC + "' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'CONFIGURACION'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")
            'cbxINH = oForm.Items.Item("cbxINH").Specific

            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                If odt.GetValue("U_Nombre", i).ToString().Equals("Param_Ambiente") Then
                    cbxCO = oForm.Items.Item("cbxCO").Specific
                    If Not odt.GetValue("U_Valor", i).ToString() = "" Then
                        cbxCO.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarQueSeGuardeLogEmision") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GLE")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_Licencia") Then
                    If Not odt.GetValue("U_Valor", i).ToString() = "" Then
                        oForm.Items.Item("ws_LIC").Specific.value = odt.GetValue("U_Valor", i).ToString()
                    Else
                        oForm.Items.Item("ws_LIC").Specific.value = "https://labcr.guru-soft.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc" 'LICENCIA
                    End If
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Param_Ruc") Then
                    oForm.Items.Item("txtNIT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("UserDB") Then
                    oForm.Items.Item("txtUserDB").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PassDB") Then
                    oForm.Items.Item("txtPassDB").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("IpServer") Then
                    oForm.Items.Item("txtIp").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaRPT") Then
                    oForm.Items.Item("txtRutaRPT").Specific.value = odt.GetValue("U_Valor", i).ToString()


                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("AbrirRPTCargadosDB") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDBRPT")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("GenerarFolio") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkgenflo")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Dinardap") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDina")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ValidarSocioNegociosUDF") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkVsdn")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ValidarDocumentosUDF") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkVDocs")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()


                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("IDSerie") Then
                    cbxSeries = oForm.Items.Item("cbSeries").Specific
                    If Not odt.GetValue("U_Valor", i).ToString() = "" Then
                        cbxSeries.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If


                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarOpcionesParaManejoChequesProtestos") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GCHP")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarOpcionesParaTrasnferStockEntreBDs") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GTEB")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ImpuestoMontoCheque") Then
                    oForm.Items.Item("txtImpMC").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("impuestoMontoProtesto") Then
                    oForm.Items.Item("txtImpMP").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreImpuestoMontoProtesto") Then
                    Dim label As SAPbouiCOM.StaticText = oForm.Items.Item("lbIMP").Specific
                    label.Caption = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreImpuestoMontoCheque") Then
                    Dim label As SAPbouiCOM.StaticText = oForm.Items.Item("lbIMC").Specific
                    label.Caption = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaRporteND") Then
                    oForm.Items.Item("txtRutaND").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CuentaTransferencia") Then
                    oForm.Items.Item("txtCuentaT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreCuentaTransferencia") Then
                    Dim label As SAPbouiCOM.StaticText = oForm.Items.Item("lblCT").Specific
                    label.Caption = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("UserBaseDatos") Then
                    oForm.Items.Item("txtUBD").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PassBaseDatos") Then
                    oForm.Items.Item("txtPBD").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("UserSAP") Then
                    oForm.Items.Item("txtUSAP").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PassSAP") Then
                    oForm.Items.Item("txtPSAP").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SetearCamposUsuario") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkSetCU")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'add 14072023
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("MostrarBotonImpresion") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkverImp")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SConexionHana") Then
                    oForm.Items.Item("txtSCone").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("DriverHana") Then
                    oForm.Items.Item("txtDriHna").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ParametroRPT") Then
                    oForm.Items.Item("txtParRPT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'items tabs Localizacion pos XY

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PosicionItemTabX") Then
                    oForm.Items.Item("txtItemx").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PosicionItemTabY") Then
                    oForm.Items.Item("txtItemy").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("GuiasDesatendidas") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkGRDE")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("QryGuiasDesSerie") Then
                    oForm.Items.Item("qry_GRD").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("QryGuiasDesNumDoc") Then
                    oForm.Items.Item("qry_GRDND").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PagosMasivos") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_PM")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'Add JP 23/08/2024
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaArchivoRPTPM") Then
                    oForm.Items.Item("txtRRPTPM").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CadenaConexionRPTPM") Then
                    oForm.Items.Item("txtCRPTPM").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ADD JP 25/10/2024
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CuentaTransitoriaPM") Then
                    oForm.Items.Item("txtCTATRA").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ADD JP 13/11/2024
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SBMenuPadreRptPreInf") Then
                    oForm.Items.Item("txtMPSBPI").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    '    'ADD JP 15/11/2024 'Se utilizara general para rpt pm
                    'ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaArchivoCEPM") Then
                    '    oForm.Items.Item("txtRRPTCE").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaArchivoCHQPM") Then
                    oForm.Items.Item("txtRPTCHQ").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ADD DM 05/12/2024
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreColumnasAnexos") Then
                    oForm.Items.Item("txtNCOL").Specific.value = odt.GetValue("U_Valor", i).ToString()


                    'ADD DM 05/12/2024
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaRepCM") Then
                    oForm.Items.Item("txtRepCM").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarServiciosBasicos") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkActSB")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                End If

                ACTUALIZA = 1
            Next

            If ACTUALIZA = 1 Then
                Dim obtnGrabar As SAPbouiCOM.Button
                obtnGrabar = oForm.Items.Item("obtnGrabar").Specific
                obtnGrabar.Caption = "Actualizar"
            End If

        Catch ex As Exception
            rsboApp.MessageBox(ex.Message.ToString())
        Finally
            oForm.Freeze(False)
            mors = Nothing
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmParametrosAddonLE" Then
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "cbxINH"
                                    oForm = rsboApp.Forms.Item("frmParametrosAddonLE")

                                Case "cbxCO"
                                    oForm = rsboApp.Forms.Item("frmParametrosAddonLE")
                                    Dim cbxCO As SAPbouiCOM.ComboBox
                                    cbxCO = oForm.Items.Item("cbxCO").Specific
                                    Dim ws_LIC As SAPbouiCOM.EditText
                                    ws_LIC = oForm.Items.Item("ws_LIC").Specific 'WS_LICENCIA

                                    If cbxCO.Value = "PRUEBAS" Then
                                        ws_LIC.Value = "https://labcr.guru-soft.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc" 'LICENCIA
                                    ElseIf cbxCO.Value = "PRODUCCION" Then
                                        ws_LIC.Value = "https://cr.edocnube.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc"
                                    End If


                            End Select
                        End If


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "obtnGrabar"

                                    Try
                                        Dim oConfiguracion As Entidades.Configuracion
                                        Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                        oForm = rsboApp.Forms.Item("frmParametrosAddonLE")

                                        Dim cbxCO As SAPbouiCOM.ComboBox
                                        cbxCO = oForm.Items.Item("cbxCO").Specific 'Compania

                                        Dim ws_LIC As SAPbouiCOM.EditText
                                        ws_LIC = oForm.Items.Item("ws_LIC").Specific 'Licencia WS

                                        Dim txtNIT As SAPbouiCOM.EditText
                                        txtNIT = oForm.Items.Item("txtNIT").Specific 'NIT,RUC,IDENTIFICACION

                                        Dim txtUserDB As SAPbouiCOM.EditText
                                        txtUserDB = oForm.Items.Item("txtUserDB").Specific

                                        Dim txtPassDB As SAPbouiCOM.EditText
                                        txtPassDB = oForm.Items.Item("txtPassDB").Specific

                                        Dim txtIp As SAPbouiCOM.EditText
                                        txtIp = oForm.Items.Item("txtIp").Specific

                                        Dim txtRutaRPT As SAPbouiCOM.EditText
                                        txtRutaRPT = oForm.Items.Item("txtRutaRPT").Specific


                                        Dim txtImpMC As SAPbouiCOM.EditText
                                        txtImpMC = oForm.Items.Item("txtImpMC").Specific
                                        Dim lbIMC As SAPbouiCOM.StaticText
                                        lbIMC = oForm.Items.Item("lbIMC").Specific

                                        Dim txtImpMP As SAPbouiCOM.EditText
                                        txtImpMP = oForm.Items.Item("txtImpMP").Specific
                                        Dim lbIMP As SAPbouiCOM.StaticText
                                        lbIMP = oForm.Items.Item("lbIMP").Specific

                                        Dim cbSeries As SAPbouiCOM.ComboBox
                                        cbSeries = oForm.Items.Item("cbSeries").Specific 'IDSerie

                                        Dim txtRutaND As SAPbouiCOM.EditText
                                        txtRutaND = oForm.Items.Item("txtRutaND").Specific

                                        Dim txtCuentaT As SAPbouiCOM.EditText
                                        txtCuentaT = oForm.Items.Item("txtCuentaT").Specific
                                        Dim lblCT As SAPbouiCOM.StaticText
                                        lblCT = oForm.Items.Item("lblCT").Specific

                                        Dim txtUBD As SAPbouiCOM.EditText
                                        txtUBD = oForm.Items.Item("txtUBD").Specific

                                        Dim txtPBD As SAPbouiCOM.EditText
                                        txtPBD = oForm.Items.Item("txtPBD").Specific

                                        Dim txtUSAP As SAPbouiCOM.EditText
                                        txtUSAP = oForm.Items.Item("txtUSAP").Specific

                                        Dim txtPSAP As SAPbouiCOM.EditText
                                        txtPSAP = oForm.Items.Item("txtPSAP").Specific

                                        'add Artur 17072023

                                        Dim txtSCone As SAPbouiCOM.EditText
                                        txtSCone = oForm.Items.Item("txtSCone").Specific


                                        Dim txtDriHna As SAPbouiCOM.EditText
                                        txtDriHna = oForm.Items.Item("txtDriHna").Specific

                                        Dim txtParRPT As SAPbouiCOM.EditText
                                        txtParRPT = oForm.Items.Item("txtParRPT").Specific

                                        'add 05204

                                        Dim txtItemx As SAPbouiCOM.EditText
                                        txtItemx = oForm.Items.Item("txtItemx").Specific

                                        Dim txtItemy As SAPbouiCOM.EditText
                                        txtItemy = oForm.Items.Item("txtItemy").Specific

                                        Dim qry_GRD As SAPbouiCOM.EditText
                                        qry_GRD = oForm.Items.Item("qry_GRD").Specific

                                        Dim qry_GRDND As SAPbouiCOM.EditText
                                        qry_GRDND = oForm.Items.Item("qry_GRDND").Specific

                                        'Add JP 23/08/2024
                                        Dim txtRRPTPM As SAPbouiCOM.EditText = oForm.Items.Item("txtRRPTPM").Specific 'Ruta archivo rpt PM
                                        Dim txtCRPTPM As SAPbouiCOM.EditText = oForm.Items.Item("txtCRPTPM").Specific 'Cadena conexion rpt PM
                                        'ADD JP 25/10/2024
                                        Dim txtCTATRA As SAPbouiCOM.EditText = oForm.Items.Item("txtCTATRA").Specific 'Cuenta trasnitoria para la creacion de PM

                                        'ADD JP 13/11/2024                                        
                                        Dim txtMPSBPI As SAPbouiCOM.EditText = oForm.Items.Item("txtMPSBPI").Specific 'Menu padre de rpt preliminar/informe
                                        'ADD JP 15/11/2024 'Se utilizara general para rpt pm
                                        'Dim txtRRPTCE As SAPbouiCOM.EditText = oForm.Items.Item("txtRRPTCE").Specific 'Ruta archivo rpt Comprobante de egreso PM

                                        Dim txtRPTCHQ As SAPbouiCOM.EditText = oForm.Items.Item("txtRPTCHQ").Specific 'Ruta archivo rpt Cheque PM

                                        'ADD DM 05/12/2024
                                        Dim txtNCOL As SAPbouiCOM.EditText = oForm.Items.Item("txtNCOL").Specific 'Nombre de columnas a formatear en excel

                                        'ADD JP 15/01/2025
                                        Dim txtRepCM As SAPbouiCOM.EditText = oForm.Items.Item("txtRepCM").Specific 'RutaParaArchivosCM

                                        'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
                                        oConfiguracion = New Entidades.Configuracion
                                        oConfiguracion.Modulo = NombreAddonLOC
                                        oConfiguracion.Tipo = "PARAMETROS"
                                        oConfiguracion.SubTipo = "CONFIGURACION"
                                        olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Param_Ambiente", cbxCO.Value))
                                        'olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Param_Inhouse", cbxINH.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_Licencia", ws_LIC.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Param_Ruc", txtNIT.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GLE")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarQueSeGuardeLogEmision", oUserDataSource.ValueEx.ToString()))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("UserDB", txtUserDB.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PassDB", txtPassDB.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("IpServer", txtIp.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaRPT", txtRutaRPT.Value))


                                        'Add 13092022
                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDBRPT")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("AbrirRPTCargadosDB", oUserDataSource.ValueEx.ToString()))

                                        'Add 07102022
                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkgenflo")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("GenerarFolio", oUserDataSource.ValueEx.ToString()))

                                        Try
                                            Functions.VariablesGlobales._SS_GenerarFolio = oUserDataSource.ValueEx.ToString()
                                        Catch ex As Exception

                                        End Try


                                        'add 14102022

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkVsdn")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ValidarSocioNegociosUDF", oUserDataSource.ValueEx.ToString()))

                                        Try
                                            Functions.VariablesGlobales._SS_ValidarSocioNegociosUDF = oUserDataSource.ValueEx.ToString()
                                        Catch ex As Exception

                                        End Try

                                        '------------------------------

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkVDocs")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ValidarDocumentosUDF", oUserDataSource.ValueEx.ToString()))

                                        Try
                                            Functions.VariablesGlobales._SS_ValidarDocumentosUDF = oUserDataSource.ValueEx.ToString()
                                        Catch ex As Exception

                                        End Try

                                        '' PARAMETROS DE MANEJO DE CHEQUES PROTESTOS
                                        'olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CuentaMontoCheque", txtMC.Value))
                                        'olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreCuentaMontoCheque", lbMC.Caption))

                                        'olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CuentaMontoProtesto", txtMP.Value))
                                        'olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreCuentaMontoProtesto", lbMP.Caption))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("IDSerie", cbSeries.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ImpuestoMontoCheque", txtImpMC.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreImpuestoMontoCheque", lbIMC.Caption))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("impuestoMontoProtesto", txtImpMP.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreImpuestoMontoProtesto", lbIMP.Caption))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaRporteND", txtRutaND.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GCHP")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarOpcionesParaManejoChequesProtestos", oUserDataSource.ValueEx.ToString()))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GTEB")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarOpcionesParaTrasnferStockEntreBDs", oUserDataSource.ValueEx.ToString()))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDina")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Dinardap", oUserDataSource.ValueEx.ToString()))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CuentaTransferencia", txtCuentaT.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreCuentaTransferencia", lblCT.Caption))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("UserBaseDatos", txtUBD.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PassBaseDatos", txtPBD.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("UserSAP", txtUSAP.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PassSAP", txtPSAP.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkSetCU")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SetearCamposUsuario", oUserDataSource.ValueEx.ToString()))

                                        'add 14072023

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkverImp")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("MostrarBotonImpresion", oUserDataSource.ValueEx.ToString()))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("DriverHana", txtDriHna.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SConexionHana", txtSCone.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ParametroRPT", txtParRPT.Value))

                                        'Posicion de los items en los tabs

                                        'add KingArtur 052024

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PosicionItemTabX", txtItemx.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PosicionItemTabY", txtItemy.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkGRDE")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("GuiasDesatendidas", oUserDataSource.ValueEx.ToString()))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QryGuiasDesSerie", qry_GRD.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QryGuiasDesNumDoc", qry_GRDND.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_PM")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PagosMasivos", oUserDataSource.ValueEx.ToString()))

                                        Try

                                            Functions.VariablesGlobales._PosicionItemTabX = txtItemx.Value
                                            Functions.VariablesGlobales._PosicionItemTabY = txtItemy.Value

                                        Catch ex As Exception

                                        End Try

                                        'Add JP 23/08/2024
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaArchivoRPTPM", txtRRPTPM.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CadenaConexionRPTPM", txtCRPTPM.Value))
                                        'ADD JP 25/10/2024
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CuentaTransitoriaPM", txtCTATRA.Value))

                                        'ADD JP 13/11/2024
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SBMenuPadreRptPreInf", txtMPSBPI.Value))

                                        'ADD JP 15/11/2024  'Se utilizara general para rpt pm
                                        'olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaArchivoCEPM", txtRRPTCE.Value))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaArchivoCHQPM", txtRPTCHQ.Value))

                                        'ADD DM 05/12/2024
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreColumnasAnexos", txtNCOL.Value))

                                        'ADD JP 15/01/2025
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaRepCM", txtRepCM.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("chkActSB")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarServiciosBasicos", oUserDataSource.ValueEx.ToString()))

                                        Try
                                            Functions.VariablesGlobales._RutaArchivoRPTPM = txtRRPTPM.Value
                                            Functions.VariablesGlobales._CadenaConexionRPTPM = txtCRPTPM.Value
                                            Functions.VariablesGlobales._RutaArchivoCHQPM = txtRPTCHQ.Value
                                        Catch ex As Exception

                                        End Try

                                        oConfiguracion.Detalle = olistaDetalleConfiguracion
                                        GuardaCONF(oConfiguracion)

                                        oForm.Items.Item("obtnGrabar").Visible = False
                                        oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                        Dim oB As SAPbouiCOM.Button
                                        oB = oForm.Items.Item("2").Specific
                                        oB.Caption = "OK"

                                    Catch ex As Exception
                                        Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent - obtnGrabar: " + ex.Message.ToString(), "frmParametrosAddonLE")
                                    Finally

                                    End Try

                                Case "btnEstSB"
                                    Try
                                        rEstructura.CreacionEstructuraSB()
                                    Catch ex As Exception
                                        Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent - btnEstSB: " + ex.Message.ToString(), "frmParametrosAddonLE")
                                    End Try
                            End Select
                        End If
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
                                        Case "txtImpMC"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSMC")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("Code", 0)
                                            Dim lbIMC As SAPbouiCOM.StaticText
                                            lbIMC = oForm.Items.Item("lbIMC").Specific
                                            lbIMC.Caption = oDataTable.GetValue("Name", 0)

                                        Case "txtImpMP"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSTMP")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("Code", 0)
                                            Dim lbIMP As SAPbouiCOM.StaticText
                                            lbIMP = oForm.Items.Item("lbIMP").Specific
                                            lbIMP.Caption = oDataTable.GetValue("Name", 0)

                                        Case "txtCuentaT"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSCT")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                            Dim lblCT As SAPbouiCOM.StaticText
                                            lblCT = oForm.Items.Item("lblCT").Specific
                                            lblCT.Caption = oDataTable.GetValue("AcctName", 0)

                                    End Select
                                End If
                            Catch ex As Exception
                                rsboApp.MessageBox("et_CHOOSE_FROM_LIST " + ex.Message.ToString())
                            End Try
                        End If

                End Select


            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmParametrosAddonLE")
            System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try
    End Sub

    Public Sub GuardaCONF(ByVal oConfiguracion As Entidades.Configuracion)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            Dim query As String
            Dim CodeExist As String = "0"
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@SS_CONF"" Where ""U_Modulo"" = '" + oConfiguracion.Modulo + "' AND ""U_Tipo"" = '" + oConfiguracion.Tipo + "' AND ""U_Subtipo"" = '" + oConfiguracion.SubTipo + "'"
            Else
                query = "Select DocEntry From [@SS_CONF] Where U_Modulo = '" + oConfiguracion.Modulo + "' AND U_Tipo = '" + oConfiguracion.Tipo + "' AND U_Subtipo = '" + oConfiguracion.SubTipo + "'"
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            'mRst = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CodeExist = "0" Then ' SI EXISTE, ELIMINO Y ACTUALIZO

                ' SI EXISTE ELIMINA PARA VOLVER A CREAR
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONFLOC")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)
                oGeneralService.Delete(oGeneralParams)

                'CREA NUEVAMENTE EL REGISTRO
                oGeneralService = oCompanyService.GetGeneralService("SS_CONFLOC")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("SS_CONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)

            Else

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONFLOC")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("SS_CONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Function ConsultaParametro(ByVal Modulo As String, ByVal Tipo As String, ByVal Subtipo As String, ByVal Nombre As String) As String
        Try
            Dim valor As String = ""
            Dim sQueryPrefijo As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryPrefijo = "SELECT A.""U_Valor"" "
                sQueryPrefijo += "FROM ""@SS_CONFD"" A INNER JOIN "
                sQueryPrefijo += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                sQueryPrefijo += " WHERE  B.""U_Modulo"" = '" + Modulo + "' AND B.""U_Tipo"" = '" + Tipo + "' "
                sQueryPrefijo += " AND B.""U_Subtipo"" = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.""U_Nombre"" = '" + Nombre + "'"
            Else
                sQueryPrefijo = "SELECT A.U_Valor "
                sQueryPrefijo += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                sQueryPrefijo += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                sQueryPrefijo += " WHERE B.U_Modulo = '" + Modulo + "' AND  B.U_Tipo = '" + Tipo + "' "
                sQueryPrefijo += " AND B.U_Subtipo = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.U_Nombre = '" + Nombre + "'"
            End If

            valor = oFuncionesB1.getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
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

    Public Sub CargaComboboxSAP(ByVal Consulta As String, ByRef Combobox As SAPbouiCOM.ComboBox, ByVal Campo As String, ByVal Descripcion As String, ByVal RecordSet As SAPbobsCOM.Recordset, ByRef nfilas As Integer)
        Try
            Dim oValidValues As SAPbouiCOM.ValidValues

            Utilitario.Util_Log.Escribir_Log("Query Consulta Combo: " + Consulta, "frmParametrosAddOnLE")

            RecordSet.DoQuery(Consulta)

            nfilas = RecordSet.RecordCount

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
        Catch ex As Exception

        End Try
    End Sub

End Class

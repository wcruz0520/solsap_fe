Imports SAPbobsCOM

Public Class frmParametrosRecepcion
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Private mors As SAPbobsCOM.Recordset = Nothing
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim odt As SAPbouiCOM.DataTable
    Dim Alia As String = ""

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioParametrosRecepcion()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmParametrosRecepcion") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmParametrosRecepcion.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmParametrosRecepcion").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmParametrosRecepcion")

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

            ' CHOOSE FROM LIST CUENTA FACTURA
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "FormatCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            'Choose from list, para NC
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCFL.SetConditions(oCons)

            Dim txtFCue As SAPbouiCOM.EditText
            txtFCue = oForm.Items.Item("txtFCue").Specific
            oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtFCue.DataBind.SetBound(True, "", "EditDS")
            txtFCue.ChooseFromListUID = "CFL1"
            txtFCue.ChooseFromListAlias = "FormatCode"

            Dim txtCCue As SAPbouiCOM.EditText
            txtCCue = oForm.Items.Item("txtCCue").Specific
            oForm.DataSources.UserDataSources.Add("EditDST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCCue.DataBind.SetBound(True, "", "EditDST")
            txtCCue.ChooseFromListUID = "CFL2"
            txtCCue.ChooseFromListAlias = "FormatCode"

            ' CFL TARJETAS DE CREDITO - RETENCION OCRC
            Dim oCFLs2 As SAPbouiCOM.ChooseFromListCollection
            Dim oCons2 As SAPbouiCOM.Conditions
            Dim oCon2 As SAPbouiCOM.Condition
            oCFLs2 = oForm.ChooseFromLists
            Dim oCFL2 As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams2 As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams2 = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams2.MultiSelection = False

            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                oCFLCreationParams2.ObjectType = "36" '65015
                Alia = "CreditCard"
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                oCFLCreationParams2.ObjectType = "1" '65015
                Alia = "FormatCode"
            End If

            oCFLCreationParams2.UniqueID = "CFL3"
            oCFL2 = oCFLs2.Add(oCFLCreationParams2)
            oCons2 = oCFL2.GetConditions()

            oCon2 = oCons2.Add()
            oCon2.Alias = Alia
            oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oCon.CondVal = "S"
            oCFL2.SetConditions(oCons2)

            'Choose from list, para retencion de la renta
            oCFLCreationParams2.UniqueID = "CFL4"
            oCFL2 = oCFLs2.Add(oCFLCreationParams2)
            oCFL2.SetConditions(oCons2)

            Dim txtCodigo As SAPbouiCOM.EditText
            txtCodigo = oForm.Items.Item("txtCodigo").Specific
            oForm.DataSources.UserDataSources.Add("EditDS3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCodigo.DataBind.SetBound(True, "", "EditDS3")
            txtCodigo.ChooseFromListUID = "CFL3"
            txtCodigo.ChooseFromListAlias = Alia

            Dim txtCodigoR As SAPbouiCOM.EditText
            txtCodigoR = oForm.Items.Item("txtCodigoR").Specific
            oForm.DataSources.UserDataSources.Add("EditDS3R", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCodigoR.DataBind.SetBound(True, "", "EditDS3R")
            txtCodigoR.ChooseFromListUID = "CFL4"
            txtCodigoR.ChooseFromListAlias = Alia

            Dim chkPedido As SAPbouiCOM.CheckBox
            chkPedido = oForm.Items.Item("chkPedido").Specific
            oForm.DataSources.UserDataSources.Add("udChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkPedido.ValOn = "Y"
            chkPedido.ValOff = "N"
            chkPedido.DataBind.SetBound(True, "", "udChk")

            Dim chkBFR As SAPbouiCOM.CheckBox
            chkBFR = oForm.Items.Item("chkPedidoR").Specific
            oForm.DataSources.UserDataSources.Add("udChkR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkBFR.ValOn = "Y"
            chkBFR.ValOff = "N"
            chkBFR.DataBind.SetBound(True, "", "udChkR")

            Dim chkDesc As SAPbouiCOM.CheckBox
            chkDesc = oForm.Items.Item("chkDesc").Specific
            oForm.DataSources.UserDataSources.Add("chkDesc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkDesc.ValOn = "Y"
            chkDesc.ValOff = "N"
            chkDesc.DataBind.SetBound(True, "", "chkDesc")

            Dim chkFRF As SAPbouiCOM.CheckBox 'fecha emision fc recibida
            chkFRF = oForm.Items.Item("chkFRF").Specific
            oForm.DataSources.UserDataSources.Add("chkFRF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkFRF.ValOn = "Y"
            chkFRF.ValOff = "N"
            chkFRF.DataBind.SetBound(True, "", "chkFRF")

            Dim chkFRN As SAPbouiCOM.CheckBox 'fecha emision nota de credito recibida
            chkFRN = oForm.Items.Item("chkFRN").Specific
            oForm.DataSources.UserDataSources.Add("chkFRN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkFRN.ValOn = "Y"
            chkFRN.ValOff = "N"
            chkFRN.DataBind.SetBound(True, "", "chkFRN")

            Dim chkFRR As SAPbouiCOM.CheckBox  'fecha emision retencion recibida
            chkFRR = oForm.Items.Item("chkFRR").Specific
            oForm.DataSources.UserDataSources.Add("chkFRR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkFRR.ValOn = "Y"
            chkFRR.ValOff = "N"
            chkFRR.DataBind.SetBound(True, "", "chkFRR")

            Dim chk_FFCF As SAPbouiCOM.CheckBox  'fecha emision factura recibida en fecha de contabilizacion del preliminar
            chk_FFCF = oForm.Items.Item("chk_FFCF").Specific
            oForm.DataSources.UserDataSources.Add("chk_FFCF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chk_FFCF.ValOn = "Y"
            chk_FFCF.ValOff = "N"
            chk_FFCF.DataBind.SetBound(True, "", "chk_FFCF")

            Dim chk_FNCF As SAPbouiCOM.CheckBox  'fecha emision nota de credito recibida en fecha de contabilizacion del preliminar
            chk_FNCF = oForm.Items.Item("chk_FNCF").Specific
            oForm.DataSources.UserDataSources.Add("chk_FNCF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chk_FNCF.ValOn = "Y"
            chk_FNCF.ValOff = "N"
            chk_FNCF.DataBind.SetBound(True, "", "chk_FNCF")

            Dim chk_FRTF As SAPbouiCOM.CheckBox  'fecha emision retencion recibida en fecha de contabilizacion del preliminar
            chk_FRTF = oForm.Items.Item("chk_FRTF").Specific
            oForm.DataSources.UserDataSources.Add("chk_FRTF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chk_FRTF.ValOn = "Y"
            chk_FRTF.ValOff = "N"
            chk_FRTF.DataBind.SetBound(True, "", "chk_FRTF")

            Dim chkFPFP As SAPbouiCOM.CheckBox  'se agrego el 8/09/2021 por solicitud de pasquel
            chkFPFP = oForm.Items.Item("chkFPFP").Specific
            oForm.DataSources.UserDataSources.Add("chkFPFP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkFPFP.ValOn = "Y"
            chkFPFP.ValOff = "N"
            chkFPFP.DataBind.SetBound(True, "", "chkFPFP")

            Dim chkMarFC As SAPbouiCOM.CheckBox  'se agrego el 11/10/2021 por solicitud de pespesca
            chkMarFC = oForm.Items.Item("chkMarFC").Specific
            oForm.DataSources.UserDataSources.Add("chkMarFC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkMarFC.ValOn = "Y"
            chkMarFC.ValOff = "N"
            chkMarFC.DataBind.SetBound(True, "", "chkMarFC")

            Dim chkMarNC As SAPbouiCOM.CheckBox  'se agrego el 11/10/2021 por solicitud de pespesca
            chkMarNC = oForm.Items.Item("chkMarNC").Specific
            oForm.DataSources.UserDataSources.Add("chkMarNC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkMarNC.ValOn = "Y"
            chkMarNC.ValOff = "N"
            chkMarNC.DataBind.SetBound(True, "", "chkMarNC")

            Dim chkMarRT As SAPbouiCOM.CheckBox  'se agrego el 11/10/2021 por solicitud de pespesca
            chkMarRT = oForm.Items.Item("chkMarRT").Specific
            oForm.DataSources.UserDataSources.Add("chkMarRT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkMarRT.ValOn = "Y"
            chkMarRT.ValOff = "N"
            chkMarRT.DataBind.SetBound(True, "", "chkMarRT")

            Dim FechasCTK As SAPbouiCOM.CheckBox
            FechasCTK = oForm.Items.Item("FechasCTK").Specific
            oForm.DataSources.UserDataSources.Add("FechasCTK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            FechasCTK.ValOn = "Y"
            FechasCTK.ValOff = "N"
            FechasCTK.DataBind.SetBound(True, "", "FechasCTK")

            Dim txtDiasV As SAPbouiCOM.EditText
            txtDiasV = oForm.Items.Item("txtDiasV").Specific
            txtDiasV.Item.Enabled = False

            Dim chkPL3F As SAPbouiCOM.CheckBox
            chkPL3F = oForm.Items.Item("chkPL3F").Specific
            oForm.DataSources.UserDataSources.Add("chkPL3F", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkPL3F.ValOn = "Y"
            chkPL3F.ValOff = "N"
            chkPL3F.DataBind.SetBound(True, "", "chkPL3F")

            Dim chkPL1F As SAPbouiCOM.CheckBox
            chkPL1F = oForm.Items.Item("chkPL1F").Specific
            oForm.DataSources.UserDataSources.Add("chkPL1F", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkPL1F.ValOn = "Y"
            chkPL1F.ValOff = "N"
            chkPL1F.DataBind.SetBound(True, "", "chkPL1F")

            If Functions.VariablesGlobales._UsuarioParamDias <> "" Then
                'Item_10.Item.Visible = True
                txtDiasV.Item.Enabled = True
            End If


            Dim chk_FAFC As SAPbouiCOM.CheckBox
            chk_FAFC = oForm.Items.Item("chk_FAFC").Specific
            oForm.DataSources.UserDataSources.Add("chk_FAFC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chk_FAFC.ValOn = "Y"
            chk_FAFC.ValOff = "N"
            chk_FAFC.DataBind.SetBound(True, "", "chk_FAFC")


            Dim chk_FANC As SAPbouiCOM.CheckBox
            chk_FANC = oForm.Items.Item("chk_FANC").Specific
            oForm.DataSources.UserDataSources.Add("chk_FANC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chk_FANC.ValOn = "Y"
            chk_FANC.ValOff = "N"
            chk_FANC.DataBind.SetBound(True, "", "chk_FANC")


            Dim CHK_FART As SAPbouiCOM.CheckBox
            CHK_FART = oForm.Items.Item("CHK_FART").Specific
            oForm.DataSources.UserDataSources.Add("CHK_FART", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            CHK_FART.ValOn = "Y"
            CHK_FART.ValOff = "N"
            CHK_FART.DataBind.SetBound(True, "", "CHK_FART")

            Dim chkffmaPL As SAPbouiCOM.CheckBox
            chkffmaPL = oForm.Items.Item("chkffmaPL").Specific
            oForm.DataSources.UserDataSources.Add("chkffmaPL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkffmaPL.ValOn = "Y"
            chkffmaPL.ValOff = "N"
            chkffmaPL.DataBind.SetBound(True, "", "chkffmaPL")

            Dim chkFFMA As SAPbouiCOM.CheckBox
            chkFFMA = oForm.Items.Item("chkFFMA").Specific
            oForm.DataSources.UserDataSources.Add("chkFFMA", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkFFMA.ValOn = "Y"
            chkFFMA.ValOff = "N"
            chkFFMA.DataBind.SetBound(True, "", "chkFFMA")

            CargaDatos()

            Dim flFactura As SAPbouiCOM.Folder
            flFactura = oForm.Items.Item("Item_4").Specific
            flFactura.Select()
            flFactura = oForm.Items.Item("Item_10").Specific
            flFactura.Item.Visible = False

            If Functions.VariablesGlobales._vgPreLot = "Y" Or Functions.VariablesGlobales._PreliminarLoteXML = "Y" Then
                flFactura = oForm.Items.Item("Item_10").Specific
                flFactura.Item.Visible = True
            End If

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargaDatos()
        oForm = rsboApp.Forms.Item("frmParametrosRecepcion")
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
                QueryFC += "FROM ""@GS_CONFD"" A INNER JOIN "
                QueryFC += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryFC += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'FC'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'FC'"
            End If

            Dim QueryNC As String = ""
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                QueryNC = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryNC += "FROM ""@GS_CONFD"" A INNER JOIN "
                QueryNC += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryNC += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryNC += " AND B.""U_Subtipo"" = 'NC'"
            Else
                QueryNC = "SELECT A.U_Nombre,A.U_Valor "
                QueryNC += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryNC += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryNC += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                QueryNC += " AND  B.U_Subtipo = 'NC'"
            End If
            Dim QueryRE As String = ""
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                QueryRE = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryRE += "FROM ""@GS_CONFD"" A INNER JOIN "
                QueryRE += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryRE += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryRE += " AND B.""U_Subtipo"" = 'RE'"
            Else
                QueryRE = "SELECT A.U_Nombre,A.U_Valor "
                QueryRE += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryRE += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryRE += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                QueryRE += " AND  B.U_Subtipo = 'RE'"
            End If
            Dim QueryPL As String = ""
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                QueryPL = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryPL += "FROM ""@GS_CONFD"" A INNER JOIN "
                QueryPL += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryPL += " WHERE  B.""U_Modulo"" = 'RECEPCION' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryPL += " AND B.""U_Subtipo"" = 'PL'"
            Else
                QueryPL = "SELECT A.U_Nombre,A.U_Valor "
                QueryPL += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryPL += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryPL += " WHERE  B.U_Modulo = 'RECEPCION' AND  B.U_Tipo = 'PARAMETROS' "
                QueryPL += " AND  B.U_Subtipo = 'PL'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")
            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1
                If odt.GetValue("U_Nombre", i).ToString().Equals("Prefijo") Then
                    oForm.Items.Item("txtFPref").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Cuenta") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreCuenta") Then
                    oForm.Items.Item("lCuentaF").Specific.Caption = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CreaPedido") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("udChk")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PermiteDescuadre") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDesc")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaEmisionFactura") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFRF")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaEmisionFacturaP") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FFCF")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FormaPagoCompras") Then 'se agrego el 8/09/2021 por solicitud de pasquel
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFPFP")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ValorFormaPagoCompras") Then
                    oForm.Items.Item("txtValorFP").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("MarcarDocFC") Then 'se agrego el 8/09/2021 por solicitud de pasquel
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMarFC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaAutEnFechaContabFC") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FAFC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                End If

                ACTUALIZA = 1
            Next

            ' CARGANDO CONFIGURACION DE NOTA DE CREDITO
            oForm.DataSources.DataTables.Item("odt").Clear()
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryNC)
            odt = oForm.DataSources.DataTables.Item("odt")
            'Dim i As Integer
            i = 0
            For i = 0 To odt.Rows.Count - 1
                If odt.GetValue("U_Nombre", i).ToString().Equals("Prefijo") Then
                    oForm.Items.Item("txtNPref").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Cuenta") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDST")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreCuenta") Then
                    oForm.Items.Item("lCuentaN").Specific.Caption = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaEmisionNotaCredito") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFRN")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaEmisionNotaCreditoP") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FNCF")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("MarcarDocNC") Then 'se agrego el 8/09/2021 por solicitud de pasquel
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMarNC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaAutEnFechaContabNC") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FANC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                End If
                ACTUALIZA = 1
            Next

            ' CARGANDO CONFIGURACION DE RETENCION
            oForm.DataSources.DataTables.Item("odt").Clear()
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryRE)
            odt = oForm.DataSources.DataTables.Item("odt")
            i = 0
            For i = 0 To odt.Rows.Count - 1
                If odt.GetValue("U_Nombre", i).ToString().Equals("CodigoRetencion") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS3")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreRetencion") Then
                    oForm.Items.Item("txtNam").Specific.Value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CodigoRetencionR") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS3R")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreRetencionR") Then
                    oForm.Items.Item("txtNamR").Specific.Value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("BFR") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("udChkR")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaEmisionRetencion") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFRR")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaEmisionRetencionP") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FRTF")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("MarcarDocRT") Then 'se agrego el 8/09/2021 por solicitud de pasquel
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMarRT")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaAutEnFechaContabRT") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_FART")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ComentarioPago") Then
                    oForm.Items.Item("txtCom").Specific.Value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaFinMesAnterior") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFFMA")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                End If
                ACTUALIZA = 1
            Next

            ' CARGANDO PROCESO EN LOTE
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryPL)
            odt = oForm.DataSources.DataTables.Item("odt")
            i = 0
            For i = 0 To odt.Rows.Count - 1
                If odt.GetValue("U_Nombre", i).ToString().Equals("ValidarFechasCTK") Then 'se agrego el 8/09/2021 por solicitud de pasquel
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("FechasCTK")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("DiasValidarProcesoLote") Then
                    oForm.Items.Item("txtDiasV").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("3Fechas") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkPL3F")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaContabilizacion") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkPL1F")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaFinMesAnteriorPL") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkffmaPL")
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
        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST _
                   And pVal.FormTypeEx = "frmParametrosRecepcion" Then
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
                            Case "txtFCue"
                                oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                                oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                Dim lCuentaF As SAPbouiCOM.StaticText
                                lCuentaF = oForm.Items.Item("lCuentaF").Specific
                                lCuentaF.Caption = oDataTable.GetValue("AcctName", 0)

                            Case "txtCCue"
                                oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDST")
                                oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                Dim lCuentaN As SAPbouiCOM.StaticText
                                lCuentaN = oForm.Items.Item("lCuentaN").Specific
                                lCuentaN.Caption = oDataTable.GetValue("AcctName", 0)

                            Case "txtCodigo"
                                oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS3")
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                        Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                    oUserDataSource.ValueEx = oDataTable.GetValue("CreditCard", 0)
                                    Dim txtNam As SAPbouiCOM.EditText
                                    txtNam = oForm.Items.Item("txtNam").Specific
                                    Try
                                        txtNam.Value = oDataTable.GetValue("CardName", 0)
                                    Catch ex As Exception
                                        txtNam.Value = val1
                                    End Try
                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                    oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                    Dim txtNam As SAPbouiCOM.EditText
                                    txtNam = oForm.Items.Item("txtNam").Specific
                                    Try
                                        txtNam.Value = oDataTable.GetValue("AcctName", 0)
                                    Catch ex As Exception
                                        txtNam.Value = val1
                                    End Try

                                End If

                            Case "txtCodigoR"
                                oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS3R")
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                       Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then

                                    oUserDataSource.ValueEx = oDataTable.GetValue("CreditCard", 0)
                                    Dim txtNam As SAPbouiCOM.EditText
                                    txtNam = oForm.Items.Item("txtNamR").Specific
                                    Try
                                        txtNam.Value = oDataTable.GetValue("CardName", 0)
                                    Catch ex As Exception
                                        txtNam.Value = val1
                                    End Try

                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

                                    oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                    Dim txtNam As SAPbouiCOM.EditText
                                    txtNam = oForm.Items.Item("txtNamR").Specific
                                    Try
                                        txtNam.Value = oDataTable.GetValue("AcctName", 0)
                                    Catch ex As Exception
                                        txtNam.Value = val1
                                    End Try

                                End If
                                
                        End Select
                    End If
                Catch ex As Exception
                    rsboApp.MessageBox("et_CHOOSE_FROM_LIST " + ex.Message.ToString())
                End Try
            End If
        End If

        If pVal.FormTypeEx = "frmParametrosRecepcion" Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If Not pVal.Before_Action Then
                        Select Case pVal.ItemUID
                            Case "obtnGrabar"

                                Try
                                    Dim oConfiguracion As Entidades.Configuracion
                                    Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                    oForm = rsboApp.Forms.Item("frmParametrosRecepcion")
                                    Dim txtFPref As SAPbouiCOM.EditText
                                    txtFPref = oForm.Items.Item("txtFPref").Specific
                                    Dim txtFCue As SAPbouiCOM.EditText
                                    txtFCue = oForm.Items.Item("txtFCue").Specific
                                    Dim lCuentaF As SAPbouiCOM.StaticText
                                    lCuentaF = oForm.Items.Item("lCuentaF").Specific
                                    Dim txtValorFP As SAPbouiCOM.EditText
                                    txtValorFP = oForm.Items.Item("txtValorFP").Specific
                                    Dim txtCom As SAPbouiCOM.EditText
                                    txtCom = oForm.Items.Item("txtCom").Specific

                                    'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
                                    oConfiguracion = New Entidades.Configuracion
                                    oConfiguracion.Modulo = "RECEPCION"
                                    oConfiguracion.Tipo = "PARAMETROS"
                                    oConfiguracion.SubTipo = "FC"
                                    olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Prefijo", txtFPref.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Cuenta", txtFCue.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreCuenta", lCuentaF.Caption))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("udChk")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CreaPedido", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDesc")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PermiteDescuadre", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFRF")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaEmisionFactura", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FFCF")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaEmisionFacturaP", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFPFP")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FormaPagoCompras", oUserDataSource.ValueEx.ToString()))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ValorFormaPagoCompras", txtValorFP.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMarFC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("MarcarDocFC", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FAFC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaAutEnFechaContabFC", oUserDataSource.ValueEx.ToString()))

                                    oConfiguracion.Detalle = olistaDetalleConfiguracion
                                    GuardaCONF(oConfiguracion)

                                    Dim txtNPref As SAPbouiCOM.EditText
                                    txtNPref = oForm.Items.Item("txtNPref").Specific
                                    Dim txtCCue As SAPbouiCOM.EditText
                                    txtCCue = oForm.Items.Item("txtCCue").Specific
                                    Dim lCuentaN As SAPbouiCOM.StaticText
                                    lCuentaN = oForm.Items.Item("lCuentaN").Specific
                                    'GrabaParametrizacion("04", "Nota de Credito de Proveedor", txtNPref.Value, txtCCue.Value, lCuentaN.Caption)
                                    oConfiguracion = New Entidades.Configuracion
                                    oConfiguracion.Modulo = "RECEPCION"
                                    oConfiguracion.Tipo = "PARAMETROS"
                                    oConfiguracion.SubTipo = "NC"
                                    olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Prefijo", txtNPref.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Cuenta", txtCCue.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreCuenta", lCuentaN.Caption))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFRN")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaEmisionNotaCredito", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FNCF")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaEmisionNotaCreditoP", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMarNC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("MarcarDocNC", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FANC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaAutEnFechaContabNC", oUserDataSource.ValueEx.ToString()))


                                    oConfiguracion.Detalle = olistaDetalleConfiguracion
                                    GuardaCONF(oConfiguracion)

                                    Dim txtCodigo As SAPbouiCOM.EditText
                                    txtCodigo = oForm.Items.Item("txtCodigo").Specific
                                    Dim txtNam As SAPbouiCOM.EditText
                                    txtNam = oForm.Items.Item("txtNam").Specific
                                    Dim txtCodigoR As SAPbouiCOM.EditText
                                    txtCodigoR = oForm.Items.Item("txtCodigoR").Specific
                                    Dim txtNamR As SAPbouiCOM.EditText
                                    txtNamR = oForm.Items.Item("txtNamR").Specific
                                    ' GrabaParametrizacion("07", "Retención de Cliente", "", txtCodigo.Value, txtNam.Value)
                                    oConfiguracion = New Entidades.Configuracion
                                    oConfiguracion.Modulo = "RECEPCION"
                                    oConfiguracion.Tipo = "PARAMETROS"
                                    oConfiguracion.SubTipo = "RE"
                                    olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CodigoRetencion", txtCodigo.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreRetencion", txtNam.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CodigoRetencionR", txtCodigoR.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreRetencionR", txtNamR.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("udChkR")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("BFR", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFRR")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaEmisionRetencion", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chk_FRTF")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaEmisionRetencionP", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMarRT")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("MarcarDocRT", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_FART")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaAutEnFechaContabRT", oUserDataSource.ValueEx.ToString()))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ComentarioPago", txtCom.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFFMA")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaFinMesAnterior", oUserDataSource.ValueEx.ToString()))


                                    oConfiguracion.Detalle = olistaDetalleConfiguracion
                                    GuardaCONF(oConfiguracion)

                                    Dim txtDiasV As SAPbouiCOM.EditText
                                    txtDiasV = oForm.Items.Item("txtDiasV").Specific
                                    oConfiguracion = New Entidades.Configuracion
                                    oConfiguracion.Modulo = "RECEPCION"
                                    oConfiguracion.Tipo = "PARAMETROS"
                                    oConfiguracion.SubTipo = "PL"
                                    olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("FechasCTK")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ValidarFechasCTK", oUserDataSource.ValueEx.ToString()))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("DiasValidarProcesoLote", txtDiasV.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkPL3F")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("3Fechas", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkPL1F")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaContabilizacion", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkffmaPL")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaFinMesAnteriorPL", oUserDataSource.ValueEx.ToString()))

                                    oConfiguracion.Detalle = olistaDetalleConfiguracion
                                    GuardaCONF(oConfiguracion)

                                    oForm.Items.Item("obtnGrabar").Visible = False
                                    oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                    Dim oB As SAPbouiCOM.Button
                                    oB = oForm.Items.Item("2").Specific
                                    oB.Caption = "OK"

                                Catch ex As Exception
                                Finally

                                End Try

                        End Select
                    End If
            End Select
        End If
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
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@GS_CONF"" Where ""U_Modulo"" = '" + oConfiguracion.Modulo + "' AND ""U_Tipo"" = '" + oConfiguracion.Tipo + "' AND ""U_Subtipo"" = '" + oConfiguracion.SubTipo + "'"
            Else
                query = "Select DocEntry From [@GS_CONF] Where U_Modulo = '" + oConfiguracion.Modulo + "' AND U_Tipo = '" + oConfiguracion.Tipo + "' AND U_Subtipo = '" + oConfiguracion.SubTipo + "'"
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            'mRst = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CodeExist = "0" Then ' SI EXISTE, ELIMINO Y ACTUALIZO

                ' SI EXISTE ELIMINA PARA VOLVER A CREAR
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)
                oGeneralService.Delete(oGeneralParams)

                'CREA NUEVAMENTE EL REGISTRO
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("GS_CONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)

            Else

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("GS_CONFD")
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
End Class

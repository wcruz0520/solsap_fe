Imports SAPbobsCOM

Public Class frmParametrosAddon
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Private mors As SAPbobsCOM.Recordset = Nothing
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim odt As SAPbouiCOM.DataTable
    Dim Alia As String = ""

    Dim cbxProveedor As SAPbouiCOM.ComboBox
    Dim cbxTipoWS As SAPbouiCOM.ComboBox


    'Public Sub AsignarURL()
    '    Dim url As String
    '    cbxTipoWS = oForm.Items.Item("CBX_TIP").Specific
    '    cbxTipoWS.Selected.Value(1, )
    '    url = cbxTipoWS.ValidValues.Item()
    '    If url = "NUBE" Then
    '        rsboApp.MessageBox("simon limon")
    '    End If

    'End Sub

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioParametrosADDON()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmParametrosAddon") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmParametrosAddon.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmParametrosAddon").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmParametrosAddon")

            Dim chkDesc As SAPbouiCOM.CheckBox
            chkDesc = oForm.Items.Item("CHK_ACC").Specific
            oForm.DataSources.UserDataSources.Add("CHK_ACC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkDesc.ValOn = "Y"
            chkDesc.ValOff = "N"
            chkDesc.DataBind.SetBound(True, "", "CHK_ACC")

            Dim chkEnviaDocEnBackGround As SAPbouiCOM.CheckBox
            chkEnviaDocEnBackGround = oForm.Items.Item("CHK_EDB").Specific
            oForm.DataSources.UserDataSources.Add("CHK_EDB", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkEnviaDocEnBackGround.ValOn = "Y"
            chkEnviaDocEnBackGround.ValOff = "N"
            chkEnviaDocEnBackGround.DataBind.SetBound(True, "", "CHK_EDB")


            Dim chkVisualizaPDF_Bytes As SAPbouiCOM.CheckBox
            chkVisualizaPDF_Bytes = oForm.Items.Item("CHK_VPB").Specific
            oForm.DataSources.UserDataSources.Add("CHK_VPB", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkVisualizaPDF_Bytes.ValOn = "Y"
            chkVisualizaPDF_Bytes.ValOff = "N"
            chkVisualizaPDF_Bytes.DataBind.SetBound(True, "", "CHK_VPB")

            Dim chkDesactivaPreguntaDeclaraable As SAPbouiCOM.CheckBox
            chkDesactivaPreguntaDeclaraable = oForm.Items.Item("CHK_NPP").Specific
            oForm.DataSources.UserDataSources.Add("CHK_NPP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkDesactivaPreguntaDeclaraable.ValOn = "Y"
            chkDesactivaPreguntaDeclaraable.ValOff = "N"
            chkDesactivaPreguntaDeclaraable.DataBind.SetBound(True, "", "CHK_NPP")

            'SALIDA HTTPS
            Dim chkSalidaporHttps As SAPbouiCOM.CheckBox
            chkSalidaporHttps = oForm.Items.Item("CHK_SH").Specific
            oForm.DataSources.UserDataSources.Add("CHK_SH", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkSalidaporHttps.ValOn = "Y"
            chkSalidaporHttps.ValOff = "N"
            chkSalidaporHttps.DataBind.SetBound(True, "", "CHK_SH")

            'FOLIACION POR POSTN

            Dim CHK_POSNT As SAPbouiCOM.CheckBox
            CHK_POSNT = oForm.Items.Item("CHK_POSNT").Specific
            oForm.DataSources.UserDataSources.Add("CHK_POSNT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            CHK_POSNT.ValOn = "Y"
            CHK_POSNT.ValOff = "N"
            CHK_POSNT.DataBind.SetBound(True, "", "CHK_POSNT")

            'recepcion lite
            Dim chkRecepcionLite As SAPbouiCOM.CheckBox
            chkRecepcionLite = oForm.Items.Item("CHKRL").Specific
            oForm.DataSources.UserDataSources.Add("CHKRL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkRecepcionLite.ValOn = "Y"
            chkRecepcionLite.ValOff = "N"
            chkRecepcionLite.DataBind.SetBound(True, "", "CHKRL")

            'validar valor recibido con el preliminar
            Dim chkValidarValRec As SAPbouiCOM.CheckBox
            chkValidarValRec = oForm.Items.Item("CHKPRE").Specific
            oForm.DataSources.UserDataSources.Add("CHKPRE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkValidarValRec.ValOn = "Y"
            chkValidarValRec.ValOff = "N"
            chkValidarValRec.DataBind.SetBound(True, "", "CHKPRE")

            'mostrar campo fecha de autorizacion
            Dim chkMostrarFA As SAPbouiCOM.CheckBox
            chkMostrarFA = oForm.Items.Item("chkMFA").Specific
            oForm.DataSources.UserDataSources.Add("chkMFA", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkMostrarFA.ValOn = "Y"
            chkMostrarFA.ValOff = "N"
            chkMostrarFA.DataBind.SetBound(True, "", "chkMFA")

            'Controlar series electronicas por UDF
            Dim chkFEUDF As SAPbouiCOM.CheckBox
            chkFEUDF = oForm.Items.Item("chkFEUDF").Specific
            oForm.DataSources.UserDataSources.Add("chkFEUDF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkFEUDF.ValOn = "Y"
            chkFEUDF.ValOff = "N"
            chkFEUDF.DataBind.SetBound(True, "", "chkFEUDF")

            'guardar log udf
            Dim chkGLOG As SAPbouiCOM.CheckBox
            chkGLOG = oForm.Items.Item("chkGLOG").Specific
            oForm.DataSources.UserDataSources.Add("chkGLOG", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkGLOG.ValOn = "Y"
            chkGLOG.ValOff = "N"
            chkGLOG.DataBind.SetBound(True, "", "chkGLOG")

            'foliacion LQ por udf (DIBEAL)
            Dim chkSLQUDF As SAPbouiCOM.CheckBox
            chkSLQUDF = oForm.Items.Item("chkSLQUDF").Specific
            oForm.DataSources.UserDataSources.Add("chkSLQUDF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkSLQUDF.ValOn = "Y"
            chkSLQUDF.ValOff = "N"
            chkSLQUDF.DataBind.SetBound(True, "", "chkSLQUDF")

            'mostrar pantalla impresion por bloque (DIBEAL)
            Dim chkImpBlo As SAPbouiCOM.CheckBox
            chkImpBlo = oForm.Items.Item("chkImpBlo").Specific
            oForm.DataSources.UserDataSources.Add("chkImpBlo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkImpBlo.ValOn = "Y"
            chkImpBlo.ValOff = "N"
            chkImpBlo.DataBind.SetBound(True, "", "chkImpBlo")

            'mostrar pantalla preliminares por lote
            Dim chkPreLot As SAPbouiCOM.CheckBox
            chkPreLot = oForm.Items.Item("chkPreLot").Specific
            oForm.DataSources.UserDataSources.Add("chkPreLot", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkPreLot.ValOn = "Y"
            chkPreLot.ValOff = "N"
            chkPreLot.DataBind.SetBound(True, "", "chkPreLot")

            'pantalla proceso lote manamer
            Dim chkProLotM As SAPbouiCOM.CheckBox
            chkProLotM = oForm.Items.Item("chkProLotM").Specific
            oForm.DataSources.UserDataSources.Add("chkProLotM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkProLotM.ValOn = "Y"
            chkProLotM.ValOff = "N"
            chkProLotM.DataBind.SetBound(True, "", "chkProLotM")

            'no enviar retencion cuando la liquidacion no este con estado 3,4,6
            Dim chkNoEnvRT As SAPbouiCOM.CheckBox
            chkNoEnvRT = oForm.Items.Item("chkNoEnvRT").Specific
            oForm.DataSources.UserDataSources.Add("chkNoEnvRT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkNoEnvRT.ValOn = "Y"
            chkNoEnvRT.ValOff = "N"
            chkNoEnvRT.DataBind.SetBound(True, "", "chkNoEnvRT")

            'bloquear boton de reenviar - para carvallo y no se dupliquen las fc
            Dim chkDesReen As SAPbouiCOM.CheckBox
            chkDesReen = oForm.Items.Item("chkDesReen").Specific
            oForm.DataSources.UserDataSources.Add("chkDesReen", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkDesReen.ValOn = "Y"
            chkDesReen.ValOff = "N"
            chkDesReen.DataBind.SetBound(True, "", "chkDesReen")

            'mostrar logo solsap
            Dim chkLogoSS As SAPbouiCOM.CheckBox
            chkLogoSS = oForm.Items.Item("chkLogoSS").Specific
            oForm.DataSources.UserDataSources.Add("chkLogoSS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkLogoSS.ValOn = "Y"
            chkLogoSS.ValOff = "N"
            chkLogoSS.DataBind.SetBound(True, "", "chkLogoSS")

            'ASIGNAR NUMERO EST+PTO+SECUENCIA EN LICTRADNUM
            Dim chkNumDoc As SAPbouiCOM.CheckBox
            chkNumDoc = oForm.Items.Item("chkNumDoc").Specific
            oForm.DataSources.UserDataSources.Add("chkNumDoc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkNumDoc.ValOn = "Y"
            chkNumDoc.ValOff = "N"
            chkNumDoc.DataBind.SetBound(True, "", "chkNumDoc")

            'CREAR FACTURA DE RESERVA DE PROVEEDORES A PARTIR DE LA FC DE CLIENTES RECIBIDAS
            Dim chkFPR As SAPbouiCOM.CheckBox
            chkFPR = oForm.Items.Item("chkFPR").Specific
            oForm.DataSources.UserDataSources.Add("chkFPR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkFPR.ValOn = "Y"
            chkFPR.ValOff = "N"
            chkFPR.DataBind.SetBound(True, "", "chkFPR")

            'PASAR EST PEMI FOLIO AL CAMPO NUMATCARD PARA EXXIS Y SOLSAP
            Dim chkAnulDoc As SAPbouiCOM.CheckBox
            chkAnulDoc = oForm.Items.Item("chkAnulDoc").Specific
            oForm.DataSources.UserDataSources.Add("chkAnulDoc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkAnulDoc.ValOn = "Y"
            chkAnulDoc.ValOff = "N"
            chkAnulDoc.DataBind.SetBound(True, "", "chkAnulDoc")

            'RECEPCION SERVICIO XML HESION
            Dim chkRecHei As SAPbouiCOM.CheckBox
            chkRecHei = oForm.Items.Item("chkRecHei").Specific
            oForm.DataSources.UserDataSources.Add("chkRecHei", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkRecHei.ValOn = "Y"
            chkRecHei.ValOff = "N"
            chkRecHei.DataBind.SetBound(True, "", "chkRecHei")

            'PRELIMINAR LOTE XML
            Dim PreLoteXml As SAPbouiCOM.CheckBox
            PreLoteXml = oForm.Items.Item("PreLoteXml").Specific
            oForm.DataSources.UserDataSources.Add("PreLoteXml", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            PreLoteXml.ValOn = "Y"
            PreLoteXml.ValOff = "N"
            PreLoteXml.DataBind.SetBound(True, "", "PreLoteXml")

            'Imprimir doc autorizado
            Dim chkDocAut As SAPbouiCOM.CheckBox
            chkDocAut = oForm.Items.Item("chkDocAut").Specific
            oForm.DataSources.UserDataSources.Add("chkDocAut", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkDocAut.ValOn = "Y"
            chkDocAut.ValOff = "N"
            chkDocAut.DataBind.SetBound(True, "", "chkDocAut")

            'Integracion Ecuanexus
            'Integracion Ecuanexus
            Dim chkIntEcu As SAPbouiCOM.CheckBox
            chkIntEcu = oForm.Items.Item("chkIntEcu").Specific
            oForm.DataSources.UserDataSources.Add("chkIntEcu", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkIntEcu.ValOn = "Y"
            chkIntEcu.ValOff = "N"
            chkIntEcu.DataBind.SetBound(True, "", "chkIntEcu")

            Dim FolioReen As SAPbouiCOM.CheckBox
            FolioReen = oForm.Items.Item("FolioReen").Specific
            oForm.DataSources.UserDataSources.Add("FolioReen", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            FolioReen.ValOn = "Y"
            FolioReen.ValOff = "N"
            FolioReen.DataBind.SetBound(True, "", "FolioReen")

            Dim chkRD As SAPbouiCOM.CheckBox
            chkRD = oForm.Items.Item("chkRD").Specific
            oForm.DataSources.UserDataSources.Add("chkRD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkRD.ValOn = "Y"
            chkRD.ValOff = "N"
            chkRD.DataBind.SetBound(True, "", "chkRD")


            Dim chkRLDE As SAPbouiCOM.CheckBox
            chkRLDE = oForm.Items.Item("chkRLDE").Specific
            oForm.DataSources.UserDataSources.Add("chkRLDE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkRLDE.ValOn = "Y"
            chkRLDE.ValOff = "N"
            chkRLDE.DataBind.SetBound(True, "", "chkRLDE")

            'validar campos nulos
            Dim chkVCN As SAPbouiCOM.CheckBox
            chkVCN = oForm.Items.Item("chkVCN").Specific
            oForm.DataSources.UserDataSources.Add("chkVCN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkVCN.ValOn = "Y"
            chkVCN.ValOff = "N"
            chkVCN.DataBind.SetBound(True, "", "chkVCN")

            'impresion doble cara
            Dim chkImpDC As SAPbouiCOM.CheckBox
            chkImpDC = oForm.Items.Item("chkImpDC").Specific
            oForm.DataSources.UserDataSources.Add("chkImpDC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkImpDC.ValOn = "Y"
            chkImpDC.ValOff = "N"
            chkImpDC.DataBind.SetBound(True, "", "chkImpDC")


            'add ArturDev 05042024
            'Funcionalidad Locaizacion
            Dim chkLocEC As SAPbouiCOM.CheckBox
            chkLocEC = oForm.Items.Item("chkLocEC").Specific
            oForm.DataSources.UserDataSources.Add("chkLocEC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkLocEC.ValOn = "Y"
            chkLocEC.ValOff = "N"
            chkLocEC.DataBind.SetBound(True, "", "chkLocEC")

            'Add 09072024
            Dim chkCM As SAPbouiCOM.CheckBox = oForm.Items.Item("chkCM").Specific
            chkCM = oForm.Items.Item("chkCM").Specific
            oForm.DataSources.UserDataSources.Add("chkCM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkCM.ValOn = "Y"
            chkCM.ValOff = "N"
            chkCM.DataBind.SetBound(True, "", "chkCM")

            'Add 03/09/2024
            Dim chkCPRPL As SAPbouiCOM.CheckBox = oForm.Items.Item("chkCPRPL").Specific
            chkCPRPL = oForm.Items.Item("chkCPRPL").Specific
            oForm.DataSources.UserDataSources.Add("chkCPRPL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkCPRPL.ValOn = "Y"
            chkCPRPL.ValOff = "N"
            chkCPRPL.DataBind.SetBound(True, "", "chkCPRPL")

            CargaDatos()

            Dim flFactura As SAPbouiCOM.Folder
            flFactura = oForm.Items.Item("Item_4").Specific
            flFactura.Select()

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargaDatos()
        oForm = rsboApp.Forms.Item("frmParametrosAddon")
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
                QueryFC += " WHERE  B.""U_Modulo"" = '" + Functions.VariablesGlobales._vgNombreAddOn + "' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'CONFIGURACION'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = '" + Functions.VariablesGlobales._vgNombreAddOn + "' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'CONFIGURACION'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")

            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1
                If odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionFC") Then
                    oForm.Items.Item("ws_FC").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionND") Then
                    oForm.Items.Item("ws_ND").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionNC") Then
                    oForm.Items.Item("ws_NC").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionRetencion") Then
                    oForm.Items.Item("ws_RE").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionGuia") Then
                    oForm.Items.Item("ws_GR").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_LiquidacionCompra") Then
                    oForm.Items.Item("ws_LQ").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionConsulta") Then
                    oForm.Items.Item("ws_COE").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_EmisionReenvioMail") Then
                    oForm.Items.Item("ws_REE").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("EmisionClave") Then
                    oForm.Items.Item("ws_CLA").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("EmisionTipo") Then
                    oForm.Items.Item("ws_TIP").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_RecepcionConsulta") Then
                    oForm.Items.Item("ws_COR").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_RecepcionEstado") Then
                    oForm.Items.Item("ws_EST").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_RecepcionConsultaArchivo") Then
                    oForm.Items.Item("ws_CORA").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RecepcionClave") Then
                    oForm.Items.Item("ws_CLAR").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ProveedorSAP") Then
                    cbxProveedor = oForm.Items.Item("CBX_PRO").Specific
                    cbxProveedor.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("TipoWebServices") Then
                    cbxTipoWS = oForm.Items.Item("CBX_TIP").Specific
                    cbxTipoWS.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("BD_User") Then
                    oForm.Items.Item("BD_USR").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("BD_Pass") Then
                    oForm.Items.Item("BD_PAS").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SRI_TA") Then
                    oForm.Items.Item("txtSRIA").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SRI_TE") Then
                    oForm.Items.Item("txtSRIT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActualizaCamposDeUsuario") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_ACC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("EnviaDocumentosEnBackGround") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_EDB")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("VisualizaPDF_Bytes") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_VPB")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("DesactivaPreguntaDeclaraable") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_NPP")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()
                    'FAMC 12022019
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Estados_docs") Then
                    oForm.Items.Item("RE_ESTDOCS").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SINCRO_RET") Then
                    oForm.Items.Item("sinc_rete").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SINCRO_LQE") Then
                    oForm.Items.Item("sinc_lqe").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SINCRO_DOC") Then
                    oForm.Items.Item("sinc_doc").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("QUERY_CORREO") Then
                    oForm.Items.Item("qry_correo").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Registros_por_paginas") Then
                    oForm.Items.Item("txt_RegPag").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Adicional_FC") Then
                    oForm.Items.Item("adc_fac").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Adicional_NC") Then
                    oForm.Items.Item("adc_nc").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Adicional_RET") Then
                    oForm.Items.Item("adc_ret").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'parametro salidahttps
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SalidaporHttps") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_SH")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'Foliacion por POSTN
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FoliacionPOSTN") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_POSNT")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'RECEPCION LITE
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RecepcionLite") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHKRL")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'ruta compartida
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Ruta_Compartida") Then
                    oForm.Items.Item("txt_RL").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'Validar valor recibido
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PagoRecibido_Seidor_exxis") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHKPRE")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'nombre campo adicional
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Nombre_CA") Then
                    oForm.Items.Item("txt_NomCA").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'fecha salida en vivo
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("FechaSalidaEnVivo") Then
                    oForm.Items.Item("txtFecSalV").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ws licencia
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WsLicencia") Then
                    oForm.Items.Item("txtWsLic").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'RUC COMPAÑIA
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RucCompañia") Then
                    oForm.Items.Item("txtRucCom").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'SERVERNODE
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ServerNode") Then
                    oForm.Items.Item("txtServer").Specific.value = odt.GetValue("U_Valor", i).ToString()


                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("TipoWsLicencia") Then
                    cbxTipoWS = oForm.Items.Item("cboWsLic").Specific
                    cbxTipoWS.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)

                    'MOSTRAR CAMPO FECHA AUTORIZACION
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("MostrarFechaAutorizacion") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMFA")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'SERIES FE POR UDF
                    'MOSTRAR CAMPO FECHA AUTORIZACION
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("SeriesFEUDF") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFEUDF")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'guardar log
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("GuardarLog") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkGLOG")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'foliacion LQ por udf (DIBEAL)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("foliacionLQ") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkSLQUDF")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'mostrar pantalla impresion por bloque (DIBEAL)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ImpresionBloque") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkImpBlo")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'mostrar pantalla preliminares por lote
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PreliminaresLote") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkPreLot")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'no enviar retencion cuando la liquidacion no este con estado 3,4,6
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NoEnviarRT") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkNoEnvRT")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'pantalla proceso lote manamer  
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ProcesoLoteManamer") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkProLotM")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'bloquear boton de reenviar - para carvallo y no se dupliquen las fc
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("BloquearReenviarSRI") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDesReen")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'mostrar logo solsap
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("MostrarLogoSolSap") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkLogoSS")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'ASIGNAR NUMERO EST+PTO+SECUENCIA EN LICTRADNUM
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("AsignarNumeroDocEnNumAtCard") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkNumDoc")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'CREAR FACTURAS DE RESERVA DE PROVEEDORES A PARTIR DE LA FACTURA DE CLIENTES RECIBIDA
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CrearFacturaDeResarvaProveedores") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFPR")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'NOMBRE DEL CAMPO DONDE SE COLOCARA EL SECUENCIAL DE LA RETENCION RECIBIDA EN EL MEDIO DE PAGO EXXIS SAP 10
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreCampoNumRet") Then
                    oForm.Items.Item("txtNumRet").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ASIGNAR NUMERO EST+PTO+SECUENCIA EN LICTRADNUM Y PONER EL FOLIO Y FOLIOPREF VACIOS AL DAR CLIC EN CANCELAR
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NumeroDocEnNumAtCardCancelar") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkAnulDoc")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'xml recepcion heison
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("XMLRecepcionHesion") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkRecHei")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'ruta fc xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaFC") Then
                    oForm.Items.Item("txtRutaFC").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ruta nc xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaNC") Then
                    oForm.Items.Item("txtRutaNC").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ruta rt xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaRT") Then
                    oForm.Items.Item("txtRutaRT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ruta pro fc xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaProFC") Then
                    oForm.Items.Item("txtProFc").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ruta pro nc xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaProNC") Then
                    oForm.Items.Item("txtProNC").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ruta pro rt xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaProRT") Then
                    oForm.Items.Item("txtProRT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'prelimnar lotes xml
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("PreliminarLotesXML") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("PreLoteXml")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ImprimirDocAut") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDocAut")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'Integracion ecuanexus
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("IntegracionEcuanexus") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkIntEcu")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'web service emision ecuanexus
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WsEmisionEcu") Then
                    oForm.Items.Item("WsEcuEmi").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'web service consulta ecuanexus
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WsConsultaEcu") Then
                    oForm.Items.Item("WsEcuCon").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'token ecuanexus
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("TokenEcu") Then
                    oForm.Items.Item("txtToken").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'nomre ws para consultar estado doc
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NombreWsEcu") Then
                    oForm.Items.Item("txtNomWs").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'query para obtener el secuencial (Localizacion solsap)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("QrySecuencial") Then
                    oForm.Items.Item("qrySecLCE").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'asignar folio al reenviar - solsap ya que el addon de localizacion generara el secuencial al crear pero no al reenviar
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("AsignarFolioReenvio") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("FolioReen")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'se agrego parametro para visualizar un nuevo boton para reenviar desde la pantalla de doc enviados al sri
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ReenviarDoc") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkRD")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'parametro para ir agregando documento a una lista por medio del check seleccionar para plasti
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ReenviarListaDocEnv") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkRLDE")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ValidarCamposNulos") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkVCN")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ImpresionDobleCara") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkImpDC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()


                    'add ArturDev 05042024 activar Localizacion EC

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarLocalizacionEC") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkLocEC")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    ' RutaIntegracionXML
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RutaIntegracionXML") Then
                    oForm.Items.Item("txtRutXML").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("TablasNativasReplace") Then
                    oForm.Items.Item("txtNatRep").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'ADD 09072024
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarCMFML") Then 'Activar Cash Management fuera de menu de LocEC
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkCM")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("QueryGRUdo") Then
                    oForm.Items.Item("qry_GRUdo").Specific.value = odt.GetValue("U_Valor", i).ToString()

                    'Add 03/09/2024 
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CantabilizarPRProcLot") Then 'Activar Cash Management fuera de menu de LocEC
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkCPRPL")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                    'Add 17/12/2024 
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("NomCampoPedInfoAdicional") Then
                    oForm.Items.Item("txtNCPIA").Specific.value = odt.GetValue("U_Valor", i).ToString()

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

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent

        If pVal.FormTypeEx = "frmParametrosAddon" Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If Not pVal.Before_Action Then
                        Select Case pVal.ItemUID
                            Case "obtnGrabar"

                                Try
                                    Dim oConfiguracion As Entidades.Configuracion
                                    Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                    oForm = rsboApp.Forms.Item("frmParametrosAddon")
                                    Dim ws_FC As SAPbouiCOM.EditText
                                    ws_FC = oForm.Items.Item("ws_FC").Specific 'WS_EmisionFC
                                    Dim ws_ND As SAPbouiCOM.EditText
                                    ws_ND = oForm.Items.Item("ws_ND").Specific 'WS_EmisionND
                                    Dim ws_NC As SAPbouiCOM.EditText
                                    ws_NC = oForm.Items.Item("ws_NC").Specific 'WS_EmisionNC
                                    Dim ws_RE As SAPbouiCOM.EditText
                                    ws_RE = oForm.Items.Item("ws_RE").Specific 'WS_EmisionRetencion
                                    Dim ws_GR As SAPbouiCOM.EditText
                                    ws_GR = oForm.Items.Item("ws_GR").Specific 'WS_EmisionGuia

                                    Dim ws_LQ As SAPbouiCOM.EditText
                                    ws_LQ = oForm.Items.Item("ws_LQ").Specific 'WS_LiquidacionCompra

                                    Dim ws_COE As SAPbouiCOM.EditText
                                    ws_COE = oForm.Items.Item("ws_COE").Specific 'WS_EmisionConsulta
                                    Dim ws_REE As SAPbouiCOM.EditText
                                    ws_REE = oForm.Items.Item("ws_REE").Specific 'WS_EmisionReenvioMail
                                    Dim ws_CLA As SAPbouiCOM.EditText
                                    ws_CLA = oForm.Items.Item("ws_CLA").Specific 'EmisionClave
                                    Dim ws_TIP As SAPbouiCOM.EditText
                                    ws_TIP = oForm.Items.Item("ws_TIP").Specific 'EmisionTipo

                                    Dim ws_COR As SAPbouiCOM.EditText
                                    ws_COR = oForm.Items.Item("ws_COR").Specific 'WS_RecepcionConsulta
                                    Dim ws_EST As SAPbouiCOM.EditText
                                    ws_EST = oForm.Items.Item("ws_EST").Specific 'WS_RecepcionEstado
                                    Dim ws_CORA As SAPbouiCOM.EditText
                                    ws_CORA = oForm.Items.Item("ws_CORA").Specific 'WS_RecepcionConsultaArchivo
                                    Dim ws_CLAR As SAPbouiCOM.EditText
                                    ws_CLAR = oForm.Items.Item("ws_CLAR").Specific 'RecepcionClave


                                    Dim CBX_PRO As SAPbouiCOM.ComboBox
                                    CBX_PRO = oForm.Items.Item("CBX_PRO").Specific 'ProveedorSAP
                                    Dim CBX_TIP As SAPbouiCOM.ComboBox
                                    CBX_TIP = oForm.Items.Item("CBX_TIP").Specific 'TipoWebServices
                                    Dim ServerNode As SAPbouiCOM.EditText
                                    ServerNode = oForm.Items.Item("txtServer").Specific 'SERVERNODE
                                    Dim BD_USR As SAPbouiCOM.EditText
                                    BD_USR = oForm.Items.Item("BD_USR").Specific 'BD_User
                                    Dim BD_PAS As SAPbouiCOM.EditText
                                    BD_PAS = oForm.Items.Item("BD_PAS").Specific 'BD_Pass

                                    Dim txtSRIA As SAPbouiCOM.EditText
                                    txtSRIA = oForm.Items.Item("txtSRIA").Specific 'SRI TIPO AMBIENTE
                                    Dim txtSRIT As SAPbouiCOM.EditText
                                    txtSRIT = oForm.Items.Item("txtSRIT").Specific 'SRI TIPO EMISION

                                    'FAMC estados de recepcion
                                    Dim txtEstados_docs As SAPbouiCOM.EditText
                                    txtEstados_docs = oForm.Items.Item("RE_ESTDOCS").Specific 'Estados documentos

                                    'FAMC parametros de sincronizacion 12/02/2019

                                    Dim txt_sincroRET As SAPbouiCOM.EditText, txt_sincroDOC As SAPbouiCOM.EditText, txt_sincroLQE As SAPbouiCOM.EditText

                                    txt_sincroRET = oForm.Items.Item("sinc_rete").Specific
                                    txt_sincroLQE = oForm.Items.Item("sinc_lqe").Specific
                                    txt_sincroDOC = oForm.Items.Item("sinc_doc").Specific

                                    Dim txt_QUERY_CORREO As SAPbouiCOM.EditText
                                    txt_QUERY_CORREO = oForm.Items.Item("qry_correo").Specific

                                    Dim txt_RegPag As SAPbouiCOM.EditText
                                    txt_RegPag = oForm.Items.Item("txt_RegPag").Specific 'Registro por pagina

                                    Dim txt_adcFc As SAPbouiCOM.EditText, txt_adcNc As SAPbouiCOM.EditText, txt_adcRet As SAPbouiCOM.EditText
                                    txt_adcFc = oForm.Items.Item("adc_fac").Specific
                                    txt_adcNc = oForm.Items.Item("adc_nc").Specific
                                    txt_adcRet = oForm.Items.Item("adc_ret").Specific

                                    'ruta compartida
                                    Dim txt_rl As SAPbouiCOM.EditText
                                    txt_rl = oForm.Items.Item("txt_RL").Specific

                                    'nombre campo adicional en documentos rebidos
                                    Dim txt_ncr As SAPbouiCOM.EditText
                                    txt_ncr = oForm.Items.Item("txt_NomCA").Specific

                                    'Fecha Salido en Vivo
                                    Dim txtFecSalV As SAPbouiCOM.EditText
                                    txtFecSalV = oForm.Items.Item("txtFecSalV").Specific

                                    Dim txtWsLic As SAPbouiCOM.EditText
                                    txtWsLic = oForm.Items.Item("txtWsLic").Specific 'WS_LICENCIA

                                    Dim txtRucCom As SAPbouiCOM.EditText
                                    txtRucCom = oForm.Items.Item("txtRucCom").Specific 'RUC_LICENCIA

                                    Dim cboWsLic As SAPbouiCOM.ComboBox
                                    cboWsLic = oForm.Items.Item("cboWsLic").Specific 'Tipo WS LICENCIA

                                    Dim txtNumRet As SAPbouiCOM.EditText
                                    txtNumRet = oForm.Items.Item("txtNumRet").Specific 'SRI TIPO EMISION

                                    Dim txtRutaFC As SAPbouiCOM.EditText
                                    txtRutaFC = oForm.Items.Item("txtRutaFC").Specific 'SRI TIPO EMISION

                                    Dim txtRutaNC As SAPbouiCOM.EditText
                                    txtRutaNC = oForm.Items.Item("txtRutaNC").Specific 'SRI TIPO EMISION

                                    Dim txtRutaRT As SAPbouiCOM.EditText
                                    txtRutaRT = oForm.Items.Item("txtRutaRT").Specific 'SRI TIPO EMISION

                                    Dim txtProFc As SAPbouiCOM.EditText
                                    txtProFc = oForm.Items.Item("txtProFc").Specific 'SRI TIPO EMISION

                                    Dim txtProNC As SAPbouiCOM.EditText
                                    txtProNC = oForm.Items.Item("txtProNC").Specific 'SRI TIPO EMISION

                                    Dim txtProRT As SAPbouiCOM.EditText
                                    txtProRT = oForm.Items.Item("txtProRT").Specific 'SRI TIPO EMISION

                                    Dim WsEcuEmi As SAPbouiCOM.EditText
                                    WsEcuEmi = oForm.Items.Item("WsEcuEmi").Specific 'ws emision ecuanexus

                                    Dim WsEcuCon As SAPbouiCOM.EditText
                                    WsEcuCon = oForm.Items.Item("WsEcuCon").Specific 'ws consulta ecuanexus

                                    Dim txtToken As SAPbouiCOM.EditText
                                    txtToken = oForm.Items.Item("txtToken").Specific 'token ecuanexus

                                    Dim txtNomWs As SAPbouiCOM.EditText
                                    txtNomWs = oForm.Items.Item("txtNomWs").Specific 'Nombre del ws para consultar estado


                                    Dim qrySecLCE As SAPbouiCOM.EditText
                                    qrySecLCE = oForm.Items.Item("qrySecLCE").Specific 'qyeru para calcular el secuencial en localizacion solsap

                                    'add Artur RutaDocumentos cargados con XML

                                    Dim txtRutXML As SAPbouiCOM.EditText
                                    txtRutXML = oForm.Items.Item("txtRutXML").Specific

                                    'tablas para replace
                                    Dim txtNatRep As SAPbouiCOM.EditText
                                    txtNatRep = oForm.Items.Item("txtNatRep").Specific

                                    'query guias remision udo
                                    Dim qry_GRUdo As SAPbouiCOM.EditText
                                    qry_GRUdo = oForm.Items.Item("qry_GRUdo").Specific

                                    Dim txtNCPIA As SAPbouiCOM.EditText
                                    txtNCPIA = oForm.Items.Item("txtNCPIA").Specific

                                    'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
                                    oConfiguracion = New Entidades.Configuracion
                                    oConfiguracion.Modulo = Functions.VariablesGlobales._vgNombreAddOn
                                    oConfiguracion.Tipo = "PARAMETROS"
                                    oConfiguracion.SubTipo = "CONFIGURACION"
                                    olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionFC", ws_FC.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionND", ws_ND.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionNC", ws_NC.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionRetencion", ws_RE.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionGuia", ws_GR.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_LiquidacionCompra", ws_LQ.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionConsulta", ws_COE.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_EmisionReenvioMail", ws_REE.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("EmisionClave", ws_CLA.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("EmisionTipo", ws_TIP.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_RecepcionConsulta", ws_COR.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_RecepcionEstado", ws_EST.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_RecepcionConsultaArchivo", ws_CORA.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RecepcionClave", ws_CLAR.Value))


                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ProveedorSAP", CBX_PRO.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("TipoWebServices", CBX_TIP.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ServerNode", ServerNode.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("BD_User", BD_USR.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("BD_Pass", BD_PAS.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SRI_TA", txtSRIA.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SRI_TE", txtSRIT.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_ACC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActualizaCamposDeUsuario", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_EDB")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("EnviaDocumentosEnBackGround", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_VPB")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("VisualizaPDF_Bytes", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_NPP")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("DesactivaPreguntaDeclaraable", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_SH")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SalidaporHttps", oUserDataSource.ValueEx.ToString()))
                                    'FOLIACION POSTN
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_POSNT")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FoliacionPOSTN", oUserDataSource.ValueEx.ToString()))

                                    'Recepcion Lite
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHKRL")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RecepcionLite", oUserDataSource.ValueEx.ToString()))

                                    'Validar valor recibido con el preliminar
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHKPRE")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PagoRecibido_Seidor_exxis", oUserDataSource.ValueEx.ToString()))

                                    'mostrar campo fecha autorizacion en doc recibidos
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkMFA")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("MostrarFechaAutorizacion", oUserDataSource.ValueEx.ToString()))

                                    'series electronicas por udf
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFEUDF")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SeriesFEUDF", oUserDataSource.ValueEx.ToString()))

                                    'GUARGAR LOG
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkGLOG")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("GuardarLog", oUserDataSource.ValueEx.ToString()))

                                    'foliacion LQ por udf (DIBEAL)
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkSLQUDF")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("foliacionLQ", oUserDataSource.ValueEx.ToString()))

                                    'mostrar pantalla impresion por bloque  (DIBEAL)
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkImpBlo")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ImpresionBloque", oUserDataSource.ValueEx.ToString()))

                                    'mostrar pantalla impresion por lote
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkPreLot")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PreliminaresLote", oUserDataSource.ValueEx.ToString()))

                                    'mostrar pantalla impresion por lote
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkNoEnvRT")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NoEnviarRT", oUserDataSource.ValueEx.ToString()))

                                    'mostrar pantalla impresion por lote
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDesReen")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("BloquearReenviarSRI", oUserDataSource.ValueEx.ToString()))

                                    'FAMC agrego el campo de estados de documentos

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Estados_docs", txtEstados_docs.Value))

                                    'FAMC sincro 12022019
                                    'param docs
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SINCRO_DOC", txt_sincroDOC.Value))
                                    'param rete
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SINCRO_RET", txt_sincroRET.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("SINCRO_LQE", txt_sincroLQE.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QUERY_CORREO", txt_QUERY_CORREO.Value))

                                    'param num para paginacion
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Registros_por_paginas", txt_RegPag.Value))
                                    'ruta compartida 
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Ruta_Compartida", txt_rl.Value))


                                    'parm adiconal en doc recibidos
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Adicional_FC", txt_adcFc.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Adicional_NC", txt_adcNc.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Adicional_RET", txt_adcRet.Value))

                                    'nombre campo adicional en documentos recibidos
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Nombre_CA", txt_ncr.Value))

                                    'fecha salida en vivo
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("FechaSalidaEnVivo", txtFecSalV.Value))

                                    'WS licencia
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WsLicencia", txtWsLic.Value))
                                    'RUC licencia
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RucCompañia", txtRucCom.Value))
                                    'Tipo WS licencia
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("TipoWsLicencia", cboWsLic.Value))
                                    'proceso preliminares lote de manamer
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkProLotM")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ProcesoLoteManamer", oUserDataSource.ValueEx.ToString()))
                                    'mostrar logo solsap
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkLogoSS")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("MostrarLogoSolSap", oUserDataSource.ValueEx.ToString()))

                                    'ASIGNAR NUMERO EST+PTO+SECUENCIA EN LICTRADNUM
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkNumDoc")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("AsignarNumeroDocEnNumAtCard", oUserDataSource.ValueEx.ToString()))

                                    'CREAR FACTURAS DE RESERVA DE PROVEEDORES A PARTIR DE LA FACTURA DE CLIENTES RECIBIDA
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkFPR")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CrearFacturaDeResarvaProveedores", oUserDataSource.ValueEx.ToString()))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreCampoNumRet", txtNumRet.Value))

                                    'ASIGNAR NUMERO EST+PTO+SECUENCIA EN LICTRADNUM
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkAnulDoc")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NumeroDocEnNumAtCardCancelar", oUserDataSource.ValueEx.ToString()))

                                    'XML RECEPCION HEISON
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkRecHei")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("XMLRecepcionHesion", oUserDataSource.ValueEx.ToString()))

                                    'RUTA XML Y RUTA PROCESADOS
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaFC", txtRutaFC.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaNC", txtRutaNC.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaRT", txtRutaRT.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaProFC", txtProFc.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaProNC", txtProNC.Value))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaProRT", txtProRT.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("PreLoteXml")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("PreliminarLotesXML", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkDocAut")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ImprimirDocAut", oUserDataSource.ValueEx.ToString()))

                                    'integracion ecuanexus
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkIntEcu")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("IntegracionEcuanexus", oUserDataSource.ValueEx.ToString()))
                                    'ws emision ecuanexus
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WsEmisionEcu", WsEcuEmi.Value))
                                    'ws consulta ecuanexus
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WsConsultaEcu", WsEcuCon.Value))
                                    'toekn ecuanexus
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("TokenEcu", txtToken.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NombreWsEcu", txtNomWs.Value))

                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QrySecuencial", qrySecLCE.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("FolioReen")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("AsignarFolioReenvio", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkRD")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ReenviarDoc", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkRLDE")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ReenviarListaDocEnv", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkVCN")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ValidarCamposNulos", oUserDataSource.ValueEx.ToString()))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkImpDC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ImpresionDobleCara", oUserDataSource.ValueEx.ToString()))

                                    'ArturDev Se añade opcion de Localizacion
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkLocEC")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarLocalizacionEC", oUserDataSource.ValueEx.ToString()))

                                    'ruta cargados XML artur
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaIntegracionXML", txtRutXML.Value))

                                    'Tablas Para Replace artur
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("TablasNativasReplace", txtNatRep.Value))

                                    'query gr udo
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryGRUdo", qry_GRUdo.Value))

                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkCM")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarCMFML", oUserDataSource.ValueEx.ToString()))

                                    'ADD 03/09/2024 
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("chkCPRPL")
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CantabilizarPRProcLot", oUserDataSource.ValueEx.ToString()))

                                    'nnombre del campo pedido que vendra en el xml
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("NomCampoPedInfoAdicional", txtNCPIA.Value))

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

                            Case "obtnVECM"
                                rsboApp.StatusBar.SetText(NombreAddon + " - Creando la estructura necesaria para Cash Management.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                rEstructura.CreacionEstructuraCM()
                                rsboApp.StatusBar.SetText(NombreAddon + " - Estructura de Cash Management creada con éxito! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        End Select
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If Not pVal.Before_Action Then
                        Select Case pVal.ItemUID
                            Case "cboWsLic"
                                oForm = rsboApp.Forms.Item("frmParametrosAddon")
                                Dim cboWsLic As SAPbouiCOM.ComboBox
                                cboWsLic = oForm.Items.Item("cboWsLic").Specific
                                Dim txtWsLic As SAPbouiCOM.EditText
                                txtWsLic = oForm.Items.Item("txtWsLic").Specific 'WS_LICENCIA

                                ''EMISION
                                'Dim ws_FC As SAPbouiCOM.EditText
                                'ws_FC = oForm.Items.Item("ws_FC").Specific

                                'Dim ws_ND As SAPbouiCOM.EditText
                                'ws_ND = oForm.Items.Item("ws_ND").Specific

                                'Dim ws_NC As SAPbouiCOM.EditText
                                'ws_NC = oForm.Items.Item("ws_NC").Specific

                                'Dim ws_RE As SAPbouiCOM.EditText
                                'ws_RE = oForm.Items.Item("ws_RE").Specific

                                'Dim ws_GR As SAPbouiCOM.EditText
                                'ws_GR = oForm.Items.Item("ws_GR").Specific

                                'Dim ws_LQ As SAPbouiCOM.EditText
                                'ws_LQ = oForm.Items.Item("ws_LQ").Specific

                                'Dim ws_COE As SAPbouiCOM.EditText
                                'ws_COE = oForm.Items.Item("ws_COE").Specific

                                'Dim ws_REE As SAPbouiCOM.EditText
                                'ws_REE = oForm.Items.Item("ws_REE").Specific

                                'Dim ws_TIP As SAPbouiCOM.EditText
                                'ws_TIP = oForm.Items.Item("ws_TIP").Specific

                                ''RECEPCION
                                'Dim ws_COR As SAPbouiCOM.EditText
                                'ws_COR = oForm.Items.Item("ws_COR").Specific

                                'Dim ws_EST As SAPbouiCOM.EditText
                                'ws_EST = oForm.Items.Item("ws_EST").Specific

                                'Dim ws_CORA As SAPbouiCOM.EditText
                                'ws_CORA = oForm.Items.Item("ws_CORA").Specific

                                'Dim txt_RegPag As SAPbouiCOM.EditText
                                'txt_RegPag = oForm.Items.Item("txt_RegPag").Specific

                                If cboWsLic.Value = "PRUEBAS" Then

                                    txtWsLic.Value = "https://labcr.guru-soft.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc" 'LICENCIA

                                    'ws_FC.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_FACTURAS.svc"
                                    'ws_ND.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_NOTAS_DEBITO.svc"
                                    'ws_NC.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_NOTAS_CREDITO.svc"
                                    'ws_RE.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_RETENCIONES.svc"
                                    'ws_GR.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_GUIAS_REMISION.svc"
                                    'ws_LQ.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_LIQUIDACIONES_COMPRA.svc"
                                    'ws_COE.Value = "https://labec.guru-soft.com/eDocEcuador/4.3/Nube/WSEDOC/WSEDOCNUBE_CONSULTA.svc"
                                    'ws_REE.Value = "https://labec.guru-soft.com/eDocEcuador/WSEDOC_REENVIO/WSEDOC_ENVIARMAIL.svc"
                                    'ws_TIP.Value = "1"

                                    'ws_COR.Value = "https://labec.guru-soft.com/eDocEcuador/WSEDOC_RECEPCION/WSRAD_KEY_CONSULTA.svc"
                                    'ws_EST.Value = "https://labec.guru-soft.com/eDocEcuador/WSEDOC_RECEPCION/WSRAD_KEY_CAMBIARESTADO.svc"
                                    'ws_CORA.Value = "https://labec.guru-soft.com/eDocEcuador/WSEDOC_RECEPCION/WSRAD_KEY_ARCHIVO.svc"
                                    'txt_RegPag.Value = "20"

                                ElseIf cboWsLic.Value = "PRODUCCION" Then

                                    txtWsLic.Value = "https://cr.edocnube.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc"

                                    'ws_FC.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_FACTURAS.svc"
                                    'ws_ND.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_NOTAS_DEBITO.svc"
                                    'ws_NC.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_NOTAS_CREDITO.svc"
                                    'ws_RE.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_RETENCIONES.svc"
                                    'ws_GR.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_GUIAS_REMISION.svc"
                                    'ws_LQ.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_LIQUIDACIONES_COMPRA.svc"
                                    'ws_COE.Value = "https://edocnube.com/4.3/Nube/WSEDOC/WSEDOCNUBE_CONSULTA.svc"
                                    'ws_REE.Value = "https://edocnube.com/EDOCWS_ENVIARMAIL/WSEDOC_ENVIARMAIL.svc"
                                    'ws_TIP.Value = "1"

                                    'ws_COR.Value = "https://edocnube.com/EDOCWSRAD_KEYRECEPCION_SSL/WSRAD_KEY_CONSULTA.svc"
                                    'ws_EST.Value = "https://edocnube.com/EDOCWSRAD_KEYRECEPCION_SSL/WSRAD_KEY_CAMBIARESTADO.svc"
                                    'ws_CORA.Value = "https://edocnube.com/EDOCWSRAD_KEYRECEPCION_SSL/WSRAD_KEY_ARCHIVO.svc"
                                    'txt_RegPag.Value = "20"

                                End If


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

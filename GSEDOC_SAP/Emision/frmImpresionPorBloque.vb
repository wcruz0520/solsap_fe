Imports System.Data.SqlClient
Imports Functions
Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports iTextSharp.text.pdf
Imports iTextSharp.text

'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Drawing.Printing
Imports Spire.Pdf
Imports Spire.Pdf.Print

Public Class frmImpresionPorBloque
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim ofila As Integer
    Dim _tipoManejo As String
    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oCFL As SAPbouiCOM.ChooseFromList
    'Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub
    Public Sub CreaFormularioImpresionPorLote()
        Dim xmlDoc As New System.Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmImpresionPorBloque") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmImpresionPorBloque.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmImpresionPorBloque").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmImpresionPorBloque")

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

            Dim lblRuta As SAPbouiCOM.StaticText
            lblRuta = oForm.Items.Item("lblRuta").Specific
            lblRuta.Caption = ""
            lblRuta.Item.Visible = False
            'lblRuta.Item.ForeColor = RGB(062,095,138)
            Dim lnkRuta As SAPbouiCOM.LinkedButton
            lnkRuta = oForm.Items.Item("lnkRuta").Specific

            Dim cmbTipo As SAPbouiCOM.ComboBox
            cmbTipo = oForm.Items.Item("cbxTipo").Specific
            'cmbTipo.ValidValues.Add("0", "Todos")
            cmbTipo.ValidValues.Add("01", "Factura")
            cmbTipo.ValidValues.Add("03", "LQ de Compra")
            cmbTipo.ValidValues.Add("04", "Nota de Crédito")
            cmbTipo.ValidValues.Add("05", "Nota de Débito")
            cmbTipo.ValidValues.Add("06", "Guía de Remisión")
            cmbTipo.ValidValues.Add("07", "Comp. de Retención")
            cmbTipo.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue)


            Dim cbImp As SAPbouiCOM.ComboBox
            cbImp = oForm.Items.Item("cbImp").Specific
            Dim i As Integer = 0
            Dim imp As String = Nothing
            For i = 0 To PrinterSettings.InstalledPrinters.Count - 1
                Dim nombreX = PrinterSettings.InstalledPrinters.Item(i)
                If i = 0 Then
                    imp = nombreX
                End If
                cbImp.ValidValues.Add(nombreX, nombreX)
            Next
            cbImp.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
            cbImp.Select(imp, SAPbouiCOM.BoSearchKey.psk_ByDescription)

            'Dim txtFchIni As SAPbouiCOM.EditText
            'txtFchIni = oForm.Items.Item("finicial").Specific
            'txtFchIni.Value = DateTime.Now.ToString("yyyyMMdd")


            'Dim txtFchFin As SAPbouiCOM.EditText
            'txtFchFin = oForm.Items.Item("ffinal").Specific
            'txtFchFin.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtNumIni As SAPbouiCOM.EditText
            txtNumIni = oForm.Items.Item("NumInicial").Specific

            Dim txtNumFin As SAPbouiCOM.EditText
            txtNumFin = oForm.Items.Item("NumFinal").Specific

            Dim lnkPr As SAPbouiCOM.LinkedButton
            lnkPr = oForm.Items.Item("lnkPr").Specific
            lnkPr.LinkedObjectType = 2
            lnkPr.Item.LinkTo = "txtRuc"

            'Dim txtNumCop As SAPbouiCOM.EditText
            'txtNumCop = oForm.Items.Item("txtNumCop").Specific

            Dim chkBN As SAPbouiCOM.CheckBox
            chkBN = oForm.Items.Item("chkBN").Specific
            oForm.DataSources.UserDataSources.Add("chkBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            chkBN.ValOn = "Y"
            chkBN.ValOff = "N"
            chkBN.DataBind.SetBound(True, "", "chkBN")

            'Dim chkCopiaBN As SAPbouiCOM.CheckBox
            'chkCopiaBN = oForm.Items.Item("chkCopiaBN").Specific
            'oForm.DataSources.UserDataSources.Add("chkCopiaBN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            ''chkPedido.DataBind.SetBound(True, "", "udChk")
            'chkCopiaBN.ValOn = "Y"
            'chkCopiaBN.ValOff = "N"
            'chkCopiaBN.DataBind.SetBound(True, "", "chkCopiaBN")

            'Estados Documentos.
            Dim cmbEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
            'If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or _
            '    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
            '    cmbEstado.ValidValues.Add("9", "Todos")
            'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            '    cmbEstado.ValidValues.Add("", "Todos")
            'End If
            cmbEstado.ValidValues.Add("", "Todos")
            cmbEstado.ValidValues.Add("0", "NO ENVIADO")
            cmbEstado.ValidValues.Add("1", "EN PROCESO")
            cmbEstado.ValidValues.Add("2", "AUTORIZADO")
            cmbEstado.ValidValues.Add("3", "NO AUTORIZADO")
            cmbEstado.ValidValues.Add("4", "VALIDAR DATOS")
            cmbEstado.ValidValues.Add("5", "EN PROCESO SRI")
            cmbEstado.ValidValues.Add("6", "DEVUELTA")
            cmbEstado.ValidValues.Add("7", "ERROR EN RECEPCION")
            cmbEstado.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue)

            'colocar la fecha actual
            'Dim txtFchIni As SAPbouiCOM.EditText
            'txtFchIni = oForm.Items.Item("FechaIni").Specific
            'txtFchIni.Value = DateTime.Now.ToString("yyyyMMdd")

            'Dim txtFchFin As SAPbouiCOM.EditText
            'txtFchFin = oForm.Items.Item("FechaFin").Specific
            'txtFchFin.Value = DateTime.Now.ToString("yyyyMMdd")

            'If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Or _
            '   Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
            '    cmbEstado.Select("9", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            '    cmbEstado.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'End If

            'FormularioImpresionPorLoteCargarGrid()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub FormularioImpresionPorLoteCargarGrid()
        oForm.Freeze(True)
        Dim FiltroFecha As String = "NO"
        Dim cbxTipoDoc As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipo").Specific
        Dim cbxEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
        'Dim txtfinicial As SAPbouiCOM.EditText = oForm.Items.Item("FechaIni").Specific
        'Dim txtffinal As SAPbouiCOM.EditText = oForm.Items.Item("FechaFin").Specific
        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

        Dim txtNuminicial As SAPbouiCOM.EditText = oForm.Items.Item("NumInicial").Specific
        Dim txtNumfinal As SAPbouiCOM.EditText = oForm.Items.Item("NumFinal").Specific

        Dim txtDespacho As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDes").Specific
        Dim txtSocioNegocio As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific
        'Dim chkImpBN As SAPbouiCOM.CheckBox = oForm.Items.Item("chkImpBN").Specific

        Dim txtfinicial As SAPbouiCOM.EditText = oForm.Items.Item("FechaIni").Specific
        Dim txtffinal As SAPbouiCOM.EditText = oForm.Items.Item("FechaFin").Specific


        'If (String.IsNullOrEmpty(txtfinicial.Value) and String.IsNullOrEmpty(txtffinal.Value)) Then
        '    rsboApp.SetStatusBarMessage("Debe ingresar un rango de fechas a consultar!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        '    oForm.Freeze(False)
        '    Exit Sub
        'End If

        Dim dfechaDesde As Date
        Dim dfechaHasta As Date
        Dim sQuery As String = ""
        Dim sTipoDoc As String = cbxTipoDoc.Value.Trim()
        Dim sEstado As String = cbxEstado.Value.Trim()
        Dim sfolioIni As String = txtfinicial.Value.Trim()
        Dim sfoliofin As String = txtffinal.Value.Trim()

        Dim sNumInicial As String = txtNuminicial.Value.Trim()
        Dim sNumfinal As String = txtNumfinal.Value.Trim()
        Dim sNumDespacho As String = txtDespacho.Value.Trim()
        Dim sSN As String = txtSocioNegocio.Value.Trim()
        'If sNumInicial = "" Then 'se comento por requerimiento de mega ya que solo sera un filtro mas que puede o no colocar esta informacion 12/04/2020
        '    rsboApp.SetStatusBarMessage("Colocar el numero de documento Inicial para la búsqueda..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        '    oForm.Freeze(False)
        '    Exit Sub
        'End If
        'If sNumfinal = "" Then
        '    rsboApp.SetStatusBarMessage("Colocar el numero de documento final para la búsqueda..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        '    oForm.Freeze(False)
        '    Exit Sub
        'End If
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
        If sNumfinal < sNumInicial Then
            rsboApp.SetStatusBarMessage("El numero final no puede ser menor al numero inicial, por favor corregir..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        If (txtfinicial.Value = "" And txtffinal.Value <> "") Or (txtfinicial.Value <> "" And txtffinal.Value = "") Then
            rsboApp.SetStatusBarMessage("Por favor colocar la fecha inicial y final..!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        If txtfinicial.Value = "" And txtffinal.Value = "" Then
            FiltroFecha = "NO"
        Else
            FiltroFecha = "SI"
        End If



        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            'CALL GS_SAP_FE_ONE_OBTENERDOCUMENTOS ('0','2',{d'2016-06-16'},{d'2017-09-28'})
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                sQuery = "CALL GS_SAP_FE_ONE_ObtenerDocumentosImpresionPorBloque ("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                sQuery = "CALL " & rCompany.CompanyDB & ".GS_SAP_FE_ObtenerDocumentosImpresionPorBloque ("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                sQuery = "CALL GS_SAP_FE_HEI_OBTENERDOCUMENTOSImpresionPorBloque ("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                sQuery = "CALL GS_SAP_FE_SYP_ObtenerDocumentosImpresionPorBloque("
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                sQuery = "CALL " & rCompany.CompanyDB & ".GS_SAP_FE_SS_ObtenerDocumentosImpresionPorBloque ("
            End If
            sQuery += "'" + sTipoDoc + "'"
            sQuery += ",'" + sEstado + "'"
            sQuery += ",'" + sNumInicial + "'"
            sQuery += ",'" + sNumfinal + "'"
            sQuery += ",'" + sNumDespacho + "'"
            sQuery += ",'" + sSN + "'"
            sQuery += ",'" + FiltroFecha + "'"
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta) + ")"

        Else
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                sQuery = "EXEC GS_SAP_FE_ONE_ObtenerDocumentosImpresionPorBloque "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                sQuery = "EXEC GS_SAP_FE_ObtenerDocumentosImpresionPorBloque "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                sQuery = "EXEC GS_SAP_FE_HEI_ObtenerDocumentosImpresionPorBloque "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                sQuery = "EXEC GS_SAP_FE_SYP_ObtenerDocumentosImpresionPorBloque "
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                sQuery = "EXEC GS_SAP_FE_SS_ObtenerDocumentosImpresionPorBloque "
            End If

            sQuery += "'" + sTipoDoc + "'"
            sQuery += ",'" + sEstado + "'"
            sQuery += ",'" + sNumInicial + "'"
            sQuery += ",'" + sNumfinal + "'"
            sQuery += ",'" + sNumDespacho + "'"
            sQuery += ",'" + sSN + "'"
            sQuery += ",'" + FiltroFecha + "'"
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
                Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + sQuery, "frmImpresionPorBloque")
                oGrid.DataTable.ExecuteQuery(sQuery)
                Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + sQuery, "frmImpresionPorBloque")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmImpresionPorBloque")
            End Try

            oGrid.Columns.Item(0).Description = "Tipo Documento"
            oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).Visible = False

            oGrid.Columns.Item(1).Description = "#"
            oGrid.Columns.Item(1).TitleObject.Caption = "#"
            oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "DocEntry"
            oGrid.Columns.Item(2).TitleObject.Caption = "DocEntry"
            oGrid.Columns.Item(2).Editable = False

            'Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            'oEditTextColumn = oGrid.Columns.Item(2)
            'oEditTextColumn.LinkedObjectType = 13

            oGrid.Columns.Item(3).Description = "Doc. Num."
            oGrid.Columns.Item(3).TitleObject.Caption = "Doc. Num."
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).Visible = False

            oGrid.Columns.Item(4).Description = "Folio"
            oGrid.Columns.Item(4).TitleObject.Caption = "Folio"
            oGrid.Columns.Item(4).Editable = False

            oGrid.Columns.Item(5).Description = "Cliente"
            oGrid.Columns.Item(5).TitleObject.Caption = "Cliente"
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).Visible = False

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

            oGrid.Columns.Item(11).Description = "Numero Autorizacion"
            oGrid.Columns.Item(11).TitleObject.Caption = "Numero Autorizacion"
            oGrid.Columns.Item(11).Visible = True
            oGrid.Columns.Item(11).Editable = False

            oGrid.Columns.Item(12).Width = 0
            oGrid.Columns.Item(12).Visible = False
            oGrid.Columns.Item(12).Editable = False

            oGrid.Columns.Item(13).Width = 0
            oGrid.Columns.Item(13).Visible = False
            oGrid.Columns.Item(13).Editable = False

            oGrid.Columns.Item(14).Visible = True
            oGrid.Columns.Item(14).Editable = True
            oGrid.Columns.Item(14).Description = "Seleccionar"
            oGrid.Columns.Item(14).TitleObject.Caption = "Seleccionar"
            oGrid.Columns.Item(14).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()
            schk.Checked = False

            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos:" + sQuery + " - " + ex.Message.ToString, "frmImpresionPorBloque")
        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If FormUID = "frmImpresionPorBloque" Then


                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If Not pVal.Before_Action Then

                            Select Case pVal.ItemUID

                                Case "BuscarPdf"

                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    'oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

                                    Dim contLineas As Integer = oGrid.Rows.Count

                                    If contLineas > 0 Then
                                        Dim Seleccionar As SAPbouiCOM.Button
                                        Seleccionar = oForm.Items.Item("btnSelPdf").Specific
                                        Dim contador As Integer = 0
                                        If Seleccionar.Caption = "Desmarcar Todo" Then
                                            If DesmarcarDocumentosPendientes(contador) Then
                                                Seleccionar.Caption = "Seleccionar Todo"
                                            End If

                                        End If
                                        Dim direccionPDF As SAPbouiCOM.StaticText = oForm.Items.Item("lblRuta").Specific
                                        direccionPDF.Caption = ""
                                    End If


                                    FormularioImpresionPorLoteCargarGrid()

                                Case "GenPdf"

                                    'Dim resp As Integer = rsboApp.MessageBox("Desea actualizar los Estados ?", 1, "SI", "NO")

                                    'Select Case resp
                                    '    Case 1

                                    Generar()
                                'generarPDF2()

                                '    Case 2


                                'End Select
                                Case "btnImpri"
                                    'Dim ImpresoraCompCORRECTA = ""
                                    ''Dim j As Integer = 0
                                    'Dim NombreImpresoraUsuario = (From k In Functions.VariablesGlobales._SS_Impresoras Where (k.Usuario = "0" Or k.Usuario = rCompany.UserSignature.ToString)).ToList
                                    'If Not IsNothing(NombreImpresoraUsuario) Then
                                    '    For Each j As Functions.DatosImpresora In NombreImpresoraUsuario
                                    '        ImpresoraCompCORRECTA = ObtenerImpresoraPorNombreParcial(j.Impresora)
                                    '    Next
                                    'End If
                                    'Imprmir(ImpresoraCompCORRECTA)
                                    Imprmir()

                                Case "btnSelPdf"

                                    If pVal.BeforeAction = False Then
                                        Dim contador As Integer = 0
                                        Dim Seleccionar As SAPbouiCOM.Button
                                        Seleccionar = oForm.Items.Item("btnSelPdf").Specific
                                        If Seleccionar.Caption = "Seleccionar Todo" Then
                                            If SeleccionarDocumentosPendientes(contador) Then
                                                Seleccionar.Caption = "Desmarcar Todo"
                                                rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se seleccionaron " + contador.ToString + " registros.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            End If
                                        Else
                                            If DesmarcarDocumentosPendientes(contador) Then
                                                Seleccionar.Caption = "Seleccionar Todo"
                                                rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - Se desmarcarón " + contador.ToString + " registros.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            Else
                                                'oForm.Items.Item("btnCon").Enabled = False
                                                rsboApp.SetStatusBarMessage(oFuncionesAddon.NombreAddon + " - No existen registros marcados.. ", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            End If
                                        End If
                                    End If


                                Case "obtnCerrar"

                                    oForm.Close()

                                Case "lnkRuta"
                                    Dim _rutapdf As SAPbouiCOM.StaticText = oForm.Items.Item("lblRuta").Specific
                                    Dim link = _rutapdf.Caption.ToString
                                    Dim Proc As New Process()
                                    Proc.StartInfo.FileName = link
                                    Proc.Start()
                                    Proc.Dispose()
                                    'ExpandirContraer()

                            End Select
                        End If
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
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmImpresionPorBloque")

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
                                oForm = rsboApp.Forms.Item("frmImpresionPorBloque")
                                Dim cbxTipo As SAPbouiCOM.ComboBox
                                cbxTipo = oForm.Items.Item("cbxTipo").Specific

                                Try
                                    oCons = oCFL.GetConditions()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al obtener condiciones:" + ex.Message.ToString, "frmImpresionPorBloque")
                                End Try


                                Dim lbSocio As SAPbouiCOM.StaticText
                                lbSocio = oForm.Items.Item("lbSocio").Specific



                                If oCons.Count > 0 Then 'If there are already user conditions.
                                    If cbxTipo.Value = "07" Or cbxTipo.Value = "03" Then ' SI ES 07, SIGNIFICA QUE ES RETENCION, POR ENDE PAGO RECIBIDO DE CLIENTE
                                        oCons.Item(oCons.Count - 1).CondVal = "S"
                                        lbSocio.Caption = " Proveedor:"

                                    Else
                                        oCons.Item(oCons.Count - 1).CondVal = "C"
                                        lbSocio.Caption = "Cliente:"

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

                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Event_MatrixLinkPressed(ByVal pVal As SAPbouiCOM.ItemEvent)

        If pVal.FormTypeEx = "frmImpresionPorBloque" Then

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

                        Case 1250000001
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest

                        Case Else
                            Exit Sub

                    End Select

            End Select

        End If
    End Sub

    Private Sub Generar()

        Try
            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmImpresionPorBloque")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim contador As Integer = 0
            Dim Seleccionar As String
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oDatable As SAPbouiCOM.DataTable
            'pintar filas
            Dim gcss As SAPbouiCOM.CommonSetting
            gcss = oGrid.CommonSetting
            ' oDatable = oForm.DataSources.DataTables.Item("dtDocs")
            oDatable = oGrid.DataTable
            'Dim oGrid As SAPbouiCOM.Grid
            'oGrid = oForm.Items.Item("oGrid").Specific
            'Dim oDatable As SAPbouiCOM.DataTable
            ''pintar filas
            'Dim gcss As SAPbouiCOM.CommonSetting
            'gcss = oGrid.CommonSetting
            '' oDatable = oForm.DataSources.DataTables.Item("dtDocs")
            'oDatable = oGrid.DataTable
#Disable Warning BC42024 ' Variable local sin usar: 'ss_clave'.
            Dim ss_clave As PdfReader
#Enable Warning BC42024 ' Variable local sin usar: 'ss_clave'.

            Dim x As Integer, y As Integer
            Dim nombre_estado As String = ""
            Dim ss_tipotabla As String = ""
            Dim identificador As Integer = 0
            Dim indexgrid As Integer = 0
            Dim lista As New List(Of PdfReader)
            Dim filepath As String = ""
            Dim temp As String = Path.GetTempPath()
            Dim nomcarp As String = "SAED_DocUni"
            'Dim s_clave As New PdfReader(filepath)
            For x = 0 To oDatable.Rows.Count - 1
                nombre_estado = oDatable.GetValue("EstadoDoc", x)
                Seleccionar = oDatable.GetValue("Seleccionar", x)
                If Seleccionar = "Y" Then
                    contador += 1

                    For y = 1 To oGrid.Rows.Count
                        indexgrid = oGrid.GetDataTableRowIndex(y)
                        If indexgrid = x Then
                            gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                            Exit For
                        End If

                    Next
                    ofila = indexgrid
                    Dim sNumDoc As String = oDataTable.GetValue(4, ofila).ToString()
                    Dim sNumAutorizacion As String = oDataTable.GetValue(11, ofila).ToString()
                    Try
                        'rsboApp.SetStatusBarMessage("Imprimiendo documento numero " + sNumDoc + ", por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                        'ConsultaPDF(sNumAutorizacion)

                        Dim TipoWebServices As String = "LOCAL"
                        'TipoWebServices = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
                        TipoWebServices = Functions.VariablesGlobales._TipoWS
                        Utilitario.Util_Log.Escribir_Log("Tipo Web Service: " + TipoWebServices.ToString, "frmImpresionPorBloque")

                        Dim url As String = ""
                        'url = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionConsulta")
                        url = Functions.VariablesGlobales._wsConsultaEmision
                        If url = "" Then
                            rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            Exit Sub
                        End If

                        Dim SALIDA_POR_PROXY As String = ""
                        'SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
                        SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
                        Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmImpresionPorBloque")
                        Dim Proxy_puerto As String = ""
                        Dim Proxy_IP As String = ""
                        Dim Proxy_Usuario As String = ""
                        Dim Proxy_Clave As String = ""

                        Dim ruta As String = ""
                        Dim ws As Object
                        If TipoWebServices = "LOCAL" Then
                            ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
                        ElseIf TipoWebServices = "NUBE" Then
                            ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                            'ElseIf TipoWebServices = "NUBE_4_1" Then
                            '    ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA
                        ElseIf TipoWebServices = "NUBE_4_1" Then
                            ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

                        End If

                        If SALIDA_POR_PROXY = "Y" Then

                            Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                            Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                            Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                            Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                            Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmImpresionPorBloque")
                            Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmImpresionPorBloque")
                            Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmImpresionPorBloque")
                            Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmImpresionPorBloque")

                            If Not Proxy_puerto = "" Then
                                proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                            Else
                                proxyobject = New System.Net.WebProxy(Proxy_IP)
                            End If
                            cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                            proxyobject.Credentials = cred

#Disable Warning BC42104 ' La variable 'ws' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                            ws.Proxy = proxyobject
#Enable Warning BC42104 ' La variable 'ws' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                            ws.Credentials = cred

                        End If

                        ws.Url = url


                        'Dim VisualizaPDF_Bytes As String = "N"
                        'VisualizaPDF_Bytes = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "VisualizaPDF_Bytes")

                        'If VisualizaPDF_Bytes = "Y" Then

                        'BYTES
                        If Not Directory.Exists(temp & "\" & nomcarp) Then
                            Directory.CreateDirectory(temp & "\" & nomcarp)
                            Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + filepath & "\" & nomcarp.ToString, "frmImpresionPorBloque")
                        End If
                        filepath = temp + nomcarp + "\" + sNumAutorizacion + ".pdf"

                        'If File.Exists(filepath) Then
                        '    File.Delete(filepath)
                        'End If

                        If Not File.Exists(filepath) Then

                            Dim FS As FileStream = Nothing
                            'If Functions.VariablesGlobales._vgHttps = "Y" Then
                            '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                            'End If
                            'Dim dbbyte As Byte() = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                            oManejoDocumentos.SetProtocolosdeSeguridad()

                            Dim dbbyte As Byte() = Nothing
                            mensaje = ""
                            If TipoWebServices = "LOCAL" Then
                                dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                            ElseIf TipoWebServices = "NUBE" Then
                                dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                            ElseIf TipoWebServices = "NUBE_4_1" Then
                                dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF", mensaje)
                            End If
                            If dbbyte Is Nothing Then
                                rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                            Else
                                FS = New FileStream(filepath, System.IO.FileMode.Create)
                                FS.Write(dbbyte, 0, dbbyte.Length)
                                FS.Close()
                                'Dim s_clave As New PdfReader(filepath)
                                ''ss_clave = New PdfReader(filepath)
                                'lista.Add(s_clave)

                                's_clave.Close()
                                'Dim mensaje As String = ""
                                'imprimir_XPDF_Bytes(dbbyte, mensaje)
                                'Utilitario.Util_Log.Escribir_Log("Error al Imprimir: " + mensaje.ToString, "frmImpresionPorBloque")
                                'File.Delete(filepath)
                            End If
                        End If
                        If File.Exists(filepath) Then
                            Dim s_clave As New PdfReader(filepath)
                            'ss_clave = New PdfReader(filepath)
                            lista.Add(s_clave)
                        End If
                        '   rsboApp.MessageBox(ss_tipotabla & " " & oDatable.GetValue("DocEntry", x))
                        gcss.SetRowBackColor(y + 1, 255000)
                    Catch ex As Exception
                        gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                    End Try

                End If

                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next

            If contador > 0 Then
                generarPDF2(lista)
                'ss_clave.Close()
                rsboApp.StatusBar.SetText("(SAED) Se genero el pdf exitosamente, revisar la ruta configurada", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                rsboApp.StatusBar.SetText("(SAED) Por favor marca los documentos para generar el PDF..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

            'FormularioDocumentosEnviadosCargarGrid()

        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Ocurrio un error en la funcion generarpdf " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    Public Function ConsultaPDF(sClaveAcceso As String) As Boolean
        Try
            Dim TipoWebServices As String = "LOCAL"
            Dim filepath As String = Path.GetTempPath()
            Dim nomcarp As String = "SAED_DocUni"
            'TipoWebServices = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
            TipoWebServices = Functions.VariablesGlobales._TipoWS
            Utilitario.Util_Log.Escribir_Log("Tipo Web Service: " + TipoWebServices.ToString, "frmImpresionPorBloque")

            Dim url As String = ""
            'url = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionConsulta")
            url = Functions.VariablesGlobales._wsConsultaEmision
            If url = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            Dim SALIDA_POR_PROXY As String = ""
            'SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmImpresionPorBloque")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            Dim ruta As String = ""
            Dim ws As Object
            If TipoWebServices = "LOCAL" Then
                ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
            Else
                ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                'ElseIf TipoWebServices = "NUBE_4_1" Then
                '    ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

            End If

            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmImpresionPorBloque")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmImpresionPorBloque")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmImpresionPorBloque")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmImpresionPorBloque")

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


            'Dim VisualizaPDF_Bytes As String = "N"
            'VisualizaPDF_Bytes = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "VisualizaPDF_Bytes")

            'If VisualizaPDF_Bytes = "Y" Then

            'BYTES
            filepath = filepath
            filepath += sClaveAcceso + ".pdf"
            Dim lista As New List(Of PdfReader)
            If Not File.Exists(filepath) Then
                Dim FS As FileStream = Nothing
                oManejoDocumentos.SetProtocolosdeSeguridad()
                Dim dbbyte As Byte() = ws.ConsultarDocumento(sClaveAcceso, "PDF")
                If dbbyte Is Nothing Then
                    rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Else
                    FS = New FileStream(filepath, System.IO.FileMode.Create)
                    FS.Write(dbbyte, 0, dbbyte.Length)
                    FS.Close()
                    Dim s_clave As New PdfReader(filepath)
                    lista.Add(s_clave)
                    Dim mensaje As String = ""
                    'imprimir_XPDF_Bytes(dbbyte, mensaje)
                    'Utilitario.Util_Log.Escribir_Log("Error al Imprimir: " + mensaje.ToString, "frmImpresionPorBloque")
                    'File.Delete(filepath)
                End If
            End If
            'Else
            'rsboApp.SetStatusBarMessage("Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
            'Dim FS As FileStream = Nothing
            'Dim dbbyte As Byte() = ws.ConsultarDocumento(sClaveAcceso, "PDF")
            'Dim mensaje As String = ""
            'imprimir_XPDF_Bytes(dbbyte, mensaje)
            'Utilitario.Util_Log.Escribir_Log("Error al Imprimir: " + mensaje.ToString, "frmImpresionPorBloque")
            'End If
            'BYTES
            'Dim Proc As New Process()
            'Proc.StartInfo.FileName = filepath
            'Proc.Start()
            'Proc.Dispose()

            'Else

            ' '' RUTA
            'rsboApp.SetStatusBarMessage("Consultando url: " + url.ToString() + " Clave Acceso: " + sClaveAcceso, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            'ruta = ws.ConsultarDocumentoRuta(sClaveAcceso, "PDF")
            'If ruta Is Nothing Then
            '    rsboApp.SetStatusBarMessage("El ws NO devolvio la ruta", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            'End If
            'If ruta.Contains("win-u8ppvmocuel") Then
            '    ruta = ruta.Replace("win-u8ppvmocuel", "gurusoft-lab.com")
            'End If
            ''ruta
            'rsboApp.SetStatusBarMessage("Abriendo la siguiente ruta: " + ruta.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
            'Dim Proc As New Process()
            'Proc.StartInfo.FileName = ruta
            'Proc.Start()
            'Proc.Dispose()
            ' '' END RUTA

            'End If

            rsboApp.SetStatusBarMessage("PDF Abierto! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
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

    Private Sub Imprmir()
        Try
            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmImpresionPorBloque")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim oGrid As SAPbouiCOM.Grid
            Dim contador As Integer = 0
            Dim Seleccionar As String
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oDatable As SAPbouiCOM.DataTable
            'pintar filas
            Dim gcss As SAPbouiCOM.CommonSetting
            gcss = oGrid.CommonSetting
            oDatable = oGrid.DataTable
            Dim x As Integer, y As Integer
            Dim nombre_estado As String = ""
            Dim ss_tipotabla As String = ""
            Dim identificador As Integer = 0
            Dim indexgrid As Integer = 0
            'Dim lista As New List(Of PdfReader)
            Dim filepath As String = ""
            Dim temp As String = Path.GetTempPath()
            Dim nomcarp As String = "SAED_DocUni"
            For x = 0 To oDatable.Rows.Count - 1
                nombre_estado = oDatable.GetValue("EstadoDoc", x)
                Seleccionar = oDatable.GetValue("Seleccionar", x)
                If Seleccionar = "Y" Then

                    For y = 1 To oGrid.Rows.Count
                        indexgrid = oGrid.GetDataTableRowIndex(y)
                        If indexgrid = x Then
                            gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                            Exit For
                        End If

                    Next
                    ofila = indexgrid
                    Dim sNumDoc As String = oDataTable.GetValue(4, ofila).ToString()
                    Dim sNumAutorizacion As String = oDataTable.GetValue(11, ofila).ToString()
                    Try
                        Dim TipoWebServices As String = "LOCAL"
                        'TipoWebServices = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "TipoWebServices")
                        TipoWebServices = Functions.VariablesGlobales._TipoWS
                        Utilitario.Util_Log.Escribir_Log("Tipo Web Service: " + TipoWebServices.ToString, "frmImpresionPorBloque")

                        Dim url As String = ""
                        'url = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_EmisionConsulta")
                        url = Functions.VariablesGlobales._wsConsultaEmision
                        If url = "" Then
                            rsboApp.SetStatusBarMessage("GS - No existe informacion del Web Service, revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            Exit Sub
                        End If

                        Dim SALIDA_POR_PROXY As String = ""
                        'SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
                        SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
                        Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmImpresionPorBloque")
                        Dim Proxy_puerto As String = ""
                        Dim Proxy_IP As String = ""
                        Dim Proxy_Usuario As String = ""
                        Dim Proxy_Clave As String = ""

                        Dim ruta As String = ""
                        Dim ws As Object
                        If TipoWebServices = "LOCAL" Then
                            ws = New Entidades.wsEDoc_ConsultaEmision_LOCAL.WSEDOC_CONSULTA
                        ElseIf TipoWebServices = "NUBE" Then
                            ws = New Entidades.wsEDoc_ConsultaEmision.WSEDOCNUBE_CONSULTA
                            'ElseIf TipoWebServices = "NUBE_4_1" Then
                            '    ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA
                        ElseIf TipoWebServices = "NUBE_4_1" Then
                            ws = New Entidades.WSEDOCNUBE_CONSULTA_v4_3.WSEDOCNUBE_CONSULTA

                        End If

                        If SALIDA_POR_PROXY = "Y" Then

                            Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                            Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                            Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                            Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                            Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmImpresionPorBloque")
                            Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmImpresionPorBloque")
                            Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmImpresionPorBloque")
                            Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmImpresionPorBloque")

                            If Not Proxy_puerto = "" Then
                                proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                            Else
                                proxyobject = New System.Net.WebProxy(Proxy_IP)
                            End If
                            cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                            proxyobject.Credentials = cred

#Disable Warning BC42104 ' La variable 'ws' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                            ws.Proxy = proxyobject
#Enable Warning BC42104 ' La variable 'ws' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                            ws.Credentials = cred

                        End If

                        ws.Url = url
                        '********************************************
                        'If Not Directory.Exists(temp & "\" & nomcarp) Then
                        '    Directory.CreateDirectory(temp & "\" & nomcarp)
                        '    Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + filepath & "\" & nomcarp.ToString, "frmImpresionPorBloque")
                        'End If

                        'filepath = temp + nomcarp + "\" + sNumAutorizacion + ".pdf"
                        'Dim dbbyte As Byte() = Nothing
                        'If File.Exists(filepath) Then
                        '    Dim mensaje As String = ""
                        '    dbbyte = File.ReadAllBytes(filepath)
                        '    imprimir_XPDF_Bytes_SpirePDF(dbbyte, mensaje)
                        'Else
                        '    If TipoWebServices = "LOCAL" Then
                        '        dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                        '    ElseIf TipoWebServices = "NUBE" Then
                        '        dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                        '    ElseIf TipoWebServices = "NUBE_4_1" Then
                        '        Utilitario.Util_Log.Escribir_Log("Antes de consultar el pdf: " + mensaje.ToString, "frmImpresionPorBloque")
                        '        dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF", mensaje)
                        '        Utilitario.Util_Log.Escribir_Log("despues de consultar el pdf: " + mensaje.ToString, "frmImpresionPorBloque")
                        '    End If
                        '    If dbbyte Is Nothing Then
                        '        rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                        '    Else
                        '        imprimir_XPDF_Bytes_SpirePDF(dbbyte, mensaje)
                        '        'Dim s_clave As New PdfReader(filepath)
                        '        'lista.Add(s_clave)
                        '        'Dim mensaje As String = ""
                        '        'imprimir_XPDF_Bytes(dbbyte, mensaje)
                        '        'Utilitario.Util_Log.Escribir_Log("Error al Imprimir: " + mensaje.ToString, "frmImpresionPorBloque")
                        '        'File.Delete(filepath)
                        '    End If
                        'End If
                        ''**************************
                        'BYTES
                        If Not Directory.Exists(temp & "\" & nomcarp) Then
                            Directory.CreateDirectory(temp & "\" & nomcarp)
                            Utilitario.Util_Log.Escribir_Log("Se creo exitosamente la carpeta " + filepath & "\" & nomcarp.ToString, "frmImpresionPorBloque")
                        End If
                        filepath = temp + nomcarp + "\" + sNumAutorizacion + ".pdf"

                        'If File.Exists(filepath) Then
                        '    File.Delete(filepath)
                        'End If

                        If Not File.Exists(filepath) Then

                            Dim FS As FileStream = Nothing
                            SetProtocolosdeSeguridad()
                            'Dim dbbyte As Byte() = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                            Dim dbbyte As Byte() = Nothing
                            mensaje = ""
                            If TipoWebServices = "LOCAL" Then
                                dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                            ElseIf TipoWebServices = "NUBE" Then
                                dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF")
                            ElseIf TipoWebServices = "NUBE_4_1" Then
                                dbbyte = ws.ConsultarDocumento(sNumAutorizacion, "PDF", mensaje)
                            End If
                            If dbbyte Is Nothing Then
                                rsboApp.SetStatusBarMessage("GS" + " - Arreglo de bytes vacío,! " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
                            Else
                                FS = New FileStream(filepath, System.IO.FileMode.Create)
                                FS.Write(dbbyte, 0, dbbyte.Length)
                                FS.Close()
                                'Dim s_clave As New PdfReader(filepath)
                                'lista.Add(s_clave)
                                'Dim mensaje As String = ""
                                'imprimir_XPDF_Bytes(dbbyte, mensaje)
                                'Utilitario.Util_Log.Escribir_Log("Error al Imprimir: " + mensaje.ToString, "frmImpresionPorBloque")
                                'File.Delete(filepath)
                            End If
                        End If
                        If File.Exists(filepath) Then
                            Dim mensaje As String = ""
                            'Dim dbbyte2 As Byte() = File.ReadAllBytes(filepath)
                            'Dim dbbyteabase64 As String = Convert.ToBase64String(dbbyte2)
                            'imprimir_XPDF_Bytes_SpirePDF(filepath, mensaje)
                            'imprimir_XPDF_Bytes_SpirePDF2(filepath, mensaje) 'ANTES DE CAMBIO SI FUNCIONA



                            imprimir_XPDF_Bytes_SpirePDF2CanDuplax(filepath, mensaje)
                            'imprimir_XPDF_Bytes_SpirePDF2CanDuplax2(filepath, mensaje, NombreImpresora)

                        End If
                        '   rsboApp.MessageBox(ss_tipotabla & " " & oDatable.GetValue("DocEntry", x))
                        gcss.SetRowBackColor(y + 1, 255000)
                    Catch ex As Exception
                        gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                    End Try
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                    rsboApp.StatusBar.SetText("(SAED) Se imprimió el pdf exitosamente, Numero: " + sNumDoc.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    contador += 1
                End If

                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next
            If contador > 0 Then
                rsboApp.StatusBar.SetText("(SAED) Se termino el proceso con éxito", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                rsboApp.StatusBar.SetText("(SAED) Por favor marca los documentos a imprimir..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            'generarPDF2(lista)

            'rsboApp.StatusBar.SetText("(SAED) Se genero el pdf exitosamente, revisar la ruta configurada", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'FormularioDocumentosEnviadosCargarGrid()

        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Ocurrio un error en la funcion generarpdf " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Public Function imprimir_XPDF_Bytes(ByVal bytesPdfFile As Byte(), ByRef mensaje As String) As Boolean
        Dim caras As Boolean = False
        Try
            'Dim numcop As Integer = CInt(ncopia)
            'Dim i As Integer
            ''Using doc As New Spire.Pdf.PdfDocument

            ''Dim doc_copia As New Spire.Pdf.PdfDocument
            ''doc.LoadFromBytes(bytesPdfFile)

            'doc.PrintSettings.PrinterName = ("PDFCreator")

            'If check = "Y" Then
            '    doc.PrintSettings.Color = False
            'Else
            '    doc.PrintSettings.Color = True
            'End If
            ''doc.PrintSettings.PrintController = New System.Drawing.Printing.StandardPrintController
            'caras = doc.PrintSettings.CanDuplex
            'If caras = True Then
            '    doc.PrintSettings.Duplex = Printing.Duplex.Default
            'End If
            '' doc.Print()

            'If numcop <> 0 Then
            '    ' For i = 0 To numcop - 1
            '    doc_copia.LoadFromBytes(bytesPdfFile)
            '    If check1 = "Y" Then

            '        doc_copia.PrintSettings.Color = False
            '    Else
            '        doc_copia.PrintSettings.Color = True
            '    End If
            '    doc_copia.PrintSettings.Copies = numcop
            '    doc_copia.Print()
            '    'Next
            'End If

            'doc.PrintSettings.Copies = (1)



            ''End Using
            mensaje = "OK"
            Return True
        Catch ex As Exception

            mensaje = "Error al intentar cargar Pdfs por bytes : " & ex.Message
            Utilitario.Util_Log.Escribir_Log("Error al Imprimir: " + mensaje.ToString, "frmImpresionPorBloque")
            Return False

        End Try


    End Function
    Public Function imprimir_XPDF_Bytes_SpirePDF(ByVal ruta As String, ByRef mensaje As String) As Boolean
        'Dim txtNumCop As SAPbouiCOM.EditText = oForm.Items.Item("txtNumCop").Specific
        'Dim copias As String = txtNumCop.Value.Trim
        'Dim numcopias As Integer
        'If copias = "" Then
        '    numcopias = 0
        'Else
        '    numcopias = CInt(copias)
        'End If


        'For i As Integer = 1 To numcopias


        Try
            Dim impreso As Boolean = False
            'MessageBox.Show("ruta: " + ruta.ToString)
            Dim psi As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo
            psi.UseShellExecute = True
            psi.CreateNoWindow = True
            psi.Verb = "print"
            psi.FileName = ruta
            psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            psi.ErrorDialog = False
            psi.Arguments = "/p"

            'psi.WindowStyle = ProcessWindowStyle.Hidden
            'Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start("Chrome.exe", ruta)
            Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
            p.CloseMainWindow()

            'impreso = p.WaitForInputIdle()
            ' If impreso = True Then
            p.Close()
            'End If
        Catch ex As Exception
            mensaje = "Error al intentar imprimir Pdfs : " & ex.Message
        End Try
        'Next

        'Dim printername As String
        'Dim oPS As New System.Drawing.Printing.PrinterSettings
        'printername = oPS.PrinterName
        '*******************
        'Using p As New Process
        '    p.StartInfo.FileName = ruta

        '    p.StartInfo.Verb = "Print"
        '    p.StartInfo.CreateNoWindow = True
        '    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

        '    p.Start()
        '    p.Kill()
        'End Using
        '***************************
        'Proc.EnableRaisingEvents = True
        'Proc.StartInfo.FileName = ruta
        'Proc.StartInfo.Arguments = Chr(34) + printername + Chr(34)
        'Proc.StartInfo.Verb = "Print"
        'Proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        'Proc.StartInfo.CreateNoWindow = False

        'Proc.Start()
        'Dim pathToExecutable As String = "Chrome.exe"
        'Dim p As New ProcessStartInfo("Chrome.exe", ruta)
        'Dim Process As New Process()
        'Process.StartInfo = p
        ''

        'Process.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
        'Process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        ''Process.WaitForExit(7000)
        ''Process.Kill()
        'Process.Start()
        'Process.StartInfo.Verb = "Print"
        'Process.Close()

    End Function

    Public Function imprimir_XPDF_Bytes_SpirePDF2(ByVal ruta As String, ByRef mensaje As String) As Boolean

        Try
            Dim dbbyte As Byte() = Nothing
            dbbyte = File.ReadAllBytes(ruta)
            Dim Stream = New MemoryStream(dbbyte)
            Using doc As New Spire.Pdf.PdfDocument
                'doc.LoadFromXPS(arrb)
                doc.LoadFromStream(Stream)
                doc.PrintSettings.PrintController = New StandardPrintController
                doc.Print()
                Utilitario.Util_Log.Escribir_Log("Impresion Realizada: " + doc.ToString(), "ImpresionAutomatica")
            End Using
        Catch ex As Exception
            mensaje = "Error al intentar imprimir Pdfs : " & ex.Message
        End Try

    End Function

    Public Function imprimir_XPDF_Bytes_SpirePDF2CanDuplax(ByVal ruta As String, ByRef mensaje As String) As Boolean

        Try
            Dim cbImpresoras As SAPbouiCOM.ComboBox = oForm.Items.Item("cbImp").Specific
            Dim NombreImpresora As String = cbImpresoras.Value.Trim()

            Dim n As Integer
            Dim NumCopias As Integer = 0

            Dim BN As SAPbouiCOM.CheckBox
            BN = oForm.Items.Item("chkBN").Specific


            Dim txtNumCopias As SAPbouiCOM.EditText
            txtNumCopias = oForm.Items.Item("txtNumCop").Specific

            If Not String.IsNullOrEmpty(txtNumCopias.Value) Then
                NumCopias = CShort(txtNumCopias.Value)
            End If

            Dim dbbyte As Byte() = Nothing
            dbbyte = File.ReadAllBytes(ruta)
            Dim Stream = New MemoryStream(dbbyte)
            Using doc As New Spire.Pdf.PdfDocument
                'doc.LoadFromXPS(arrb)
                doc.LoadFromStream(Stream)
                doc.PrintSettings.PrintController = New StandardPrintController
                Utilitario.Util_Log.Escribir_Log("NombreImresora: " + NombreImpresora.ToString(), "ImpresionAutomatica")
                doc.PrintSettings.PrinterName = NombreImpresora
                Dim pageCount As Integer = doc.Pages.Count

                If NumCopias > 0 Then
                    If pageCount > 1 Then
                        doc.PrintSettings.Copies = NumCopias + 1
                    Else
                        doc.PrintSettings.Copies = NumCopias + 1
                    End If

                End If

                If Functions.VariablesGlobales._ImpresionDobleCara = "Y" Then
                    If pageCount > 1 Then
                        If doc.PrintSettings.CanDuplex Then
                            doc.PrintSettings.PaperSize = New System.Drawing.Printing.PaperSize("A4", 827, 1169)
                            doc.PrintSettings.Duplex = Duplex.Vertical
                        Else
                            Utilitario.Util_Log.Escribir_Log("Impresion no admite doble cara", "ImpresionAutomatica")
                        End If
                    End If
                End If

                If BN.Checked = True Then
                    doc.PrintSettings.SelectPageRange(1, doc.Pages.Count)
                    doc.PrintSettings.Color = False
                End If

                doc.Print()
                doc.Close()

                Utilitario.Util_Log.Escribir_Log("Impresion Realizada: " + doc.ToString(), "ImpresionAutomatica")
            End Using
        Catch ex As Exception
            mensaje = "Error al intentar imprimir Pdfs : " & ex.Message
        End Try

    End Function

    Public Function imprimir_XPDF_Bytes_SpirePDF2CanDuplax2(ByVal ruta As String, ByRef mensaje As String, ByRef NombreImpresora As String) As Boolean

        Try
            Dim n As Integer
            Dim NumCopias As Integer = 0

            Dim BN As SAPbouiCOM.CheckBox
            BN = oForm.Items.Item("chkBN").Specific


            Dim txtNumCopias As SAPbouiCOM.EditText
            txtNumCopias = oForm.Items.Item("txtNumCop").Specific

            If Not String.IsNullOrEmpty(txtNumCopias.Value) Then
                NumCopias = CShort(txtNumCopias.Value)
            End If

            Dim dbbyte As Byte() = Nothing
            dbbyte = File.ReadAllBytes(ruta)

            Dim tempFilePath As String = Path.GetTempFileName()
            tempFilePath += "001.pdf"
            ' File.WriteAllBytes(tempFilePath + ".pdf", dbbyte)
            Dim FS As FileStream = Nothing
            FS = New FileStream(tempFilePath, System.IO.FileMode.Create)
            FS.Write(dbbyte, 0, dbbyte.Length)
            FS.Close()

            'Dim Stream = New MemoryStream(dbbyte)
            Dim doc As Spire.Pdf.PdfDocument = New Spire.Pdf.PdfDocument()

            doc.LoadFromFile(tempFilePath)
            Utilitario.Util_Log.Escribir_Log("NombreImresora: " + NombreImpresora.ToString(), "ImpresionAutomatica")
            doc.PrintSettings.PrinterName = NombreImpresora
            Dim pageCount As Integer = doc.Pages.Count


            If NumCopias > 0 Then
                If pageCount > 1 Then
                    doc.PrintSettings.Copies = NumCopias + 1
                Else
                    doc.PrintSettings.Copies = NumCopias + 1
                End If

            End If

            If Functions.VariablesGlobales._ImpresionDobleCara = "Y" Then
                If pageCount > 1 Then
                    If doc.PrintSettings.CanDuplex Then
                        doc.PrintSettings.PaperSize = New System.Drawing.Printing.PaperSize("A4", 827, 1169)
                        doc.PrintSettings.Duplex = Duplex.Vertical
                    Else
                        Utilitario.Util_Log.Escribir_Log("Impresion no admite doble cara", "ImpresionAutomatica")
                    End If
                End If
            End If

            If BN.Checked = True Then
                doc.PrintSettings.SelectPageRange(1, doc.Pages.Count)
                doc.PrintSettings.Color = False
            End If

            doc.Print()
            doc.Close()
            File.Delete(tempFilePath)
            Utilitario.Util_Log.Escribir_Log("Impresion Realizada: " + doc.ToString(), "ImpresionAutomatica")


        Catch ex As Exception
            mensaje = "Error al intentar imprimir Pdfs : " & ex.Message
        End Try

    End Function


    Private Function generarPDF()
        'sClaveAcceso As String, ruta As String, Optional ByVal imprimir As Boolean = False, Optional ByVal numVeces As Integer = 0, Optional ByVal formato As String = "PDF"
        Dim pdf1 As New PdfReader("1305202003179130416000110010020000215534358176216.pdf")
        Dim pdf2 As New PdfReader("BELTRAN TORRES JULIO CESAR_001-001-000000197.pdf")
        Dim pdfNOVO As New Document

        Dim juntar As New PdfCopy(pdfNOVO, New FileStream("unionpdf.pdf", FileMode.Create))
        pdfNOVO.Open()

        Dim Importa1 As PdfImportedPage = juntar.GetImportedPage(pdf1, 1)
        Dim Importa2 As PdfImportedPage = juntar.GetImportedPage(pdf2, 1)

        juntar.AddPage(Importa1)
        juntar.AddPage(Importa2)

        pdf1.Close()
        pdf2.Close()
        juntar.Close()
        pdfNOVO.Close()

#Disable Warning BC42105 ' La función 'generarPDF' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'generarPDF' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Private Function generarPDF2(listao As IList(Of PdfReader))

        'lista.Add(pdf1)
        'lista.Add(pdf2)
        'Dim rutapdf = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Ruta_Compartida")
        Dim rutapdf = Functions.VariablesGlobales._Ruta_Compartida
        Dim txtNuminicial As SAPbouiCOM.EditText = oForm.Items.Item("NumInicial").Specific
        Dim txtNumfinal As SAPbouiCOM.EditText = oForm.Items.Item("NumFinal").Specific
        Dim cbxTipoDoc As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipo").Specific
        Dim _rutapdf As SAPbouiCOM.StaticText = oForm.Items.Item("lblRuta").Specific
        Dim link As SAPbouiCOM.LinkedButton = oForm.Items.Item("lnkRuta").Specific
        Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific
        Dim txtNumDes As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDes").Specific
        Dim sNumInicial As String = txtNuminicial.Value.Trim()
        Dim sNumfinal As String = txtNumfinal.Value.Trim()
        Dim sTipoDoc As String = cbxTipoDoc.Value.Trim()
        Dim sCodSN As String = txtRuc.Value.Trim()
        Dim SNumeroDespacho As String = txtNumDes.Value.Trim()



        Dim doc As String = sTipoDoc
        If doc = "01" Then
            doc = "Factura"
        ElseIf doc = "03" Then
            doc = "LiquidacionCompra"
        ElseIf doc = "04" Then
            doc = "NotaCredito"
        ElseIf doc = "05" Then
            doc = "NotaDebito"
        ElseIf doc = "06" Then
            doc = "GuiaRemision"
        ElseIf doc = "07" Then
            doc = "Retencion"
        End If
        'rutapdf = rutapdf + "\" + doc + " " + sNumInicial + " a " + sNumfinal + ".pdf"
        rutapdf = rutapdf + "\" + doc + " " + sNumInicial + " a " + sNumfinal + " CodSN " + sCodSN + " NumDespacho " + SNumeroDespacho + ".pdf"

        'If System.IO.File.Exists(rutapdf) = True Then
        '    'If System.IO.File.Exists(rutapdf) = True
        '    'System.IO.File.Delete(rutapdf)
        'End If

        Dim lista As New List(Of PdfReader)
        'Dim ruta As File
        lista = listao
        'Dim pdf1 As New PdfReader("1305202003179130416000110010020000215534358176216.pdf")
        'Dim pdf2 As New PdfReader("BELTRAN TORRES JULIO CESAR_001-001-000000197.pdf")
        Dim pdfNOVO As New Document
        Dim pdfcopy As New PdfCopy(pdfNOVO, New FileStream(rutapdf, FileMode.Create))



        pdfNOVO.Open()
        Dim x As PdfReader

        For Each x In lista

            Dim pdfreader As New PdfReader(x)

            Dim conpag As Integer = pdfreader.NumberOfPages
            Dim y As Integer
            For y = 1 To conpag
                pdfcopy.AddPage(pdfcopy.GetImportedPage(pdfreader, ++y))
            Next
            pdfcopy.FreeReader(pdfreader)
            pdfreader.Close()

        Next
        pdfcopy.Flush()
        'If pdfNOVO Is Nothing Then
        pdfcopy.Close()
        pdfNOVO.Close()
        lista.Clear()

        _rutapdf.Caption = rutapdf.ToString
        _rutapdf.Item.ForeColor = RGB(0, 0, 255)
        _rutapdf.Item.Visible = True

        'link.Item.Visible = True

        Dim Proc As New Process()
        Proc.StartInfo.FileName = rutapdf
        Proc.Start()
        Proc.Dispose()
        Proc.Close()


#Disable Warning BC42105 ' La función 'generarPDF2' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'generarPDF2' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Private Function LinkPdf(ruta As String)
        Dim Proc As New Process()
        Proc.StartInfo.FileName = ruta
        Proc.Start()
        Proc.Dispose()
#Disable Warning BC42105 ' La función 'LinkPdf' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'LinkPdf' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Private Function DesmarcarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmImpresionPorBloque")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            'Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                'estado = oGridDet.GetValue("EstadoDoc", i)
                'If estado = "AUTORIZADA" Then
                oGridDet.SetValue("Seleccionar", i, "N")
                contador += 1
                resul = True
                'End If
            Next
            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmImpresionPorBloque")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al SeleccionarDocumentosPendientes:" + ex.Message().ToString(), "frmImpresionPorBloque")
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function

    Private Function SeleccionarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmImpresionPorBloque")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            'Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                'estado = oGridDet.GetValue("EstadoDoc", i)
                'If estado = "AUTORIZADA" Then
                oGridDet.SetValue("Seleccionar", i, "Y")
                contador += 1
                resul = True
                'End If
            Next
            Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmImpresionPorBloque")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al SeleccionarDocumentosPendientes:" + ex.Message().ToString(), "frmImpresionPorBloque")
            'resul = False
        Finally
            oForm.Freeze(False)
        End Try
        Return resul
    End Function
    Private Function ChooseFromList(ByRef pVal As SAPbouiCOM.ItemEvent, ByVal FormUID As String) As Boolean

        Dim bBubbleEvent As Boolean = True

        If FormUID = "frmDocumentosRecibidos" Then
            oForm = rsboApp.Forms.Item("frmDocumentosRecibidos")

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

    Private Function ObtenerImpresoraPorNombreParcial(nombreprinter As String) As String
        Dim i As Integer = 0
        Try
            For i = 0 To PrinterSettings.InstalledPrinters.Count - 1
                Dim nombreX = PrinterSettings.InstalledPrinters.Item(i)

                If nombreX.ToLower.Contains(nombreprinter.ToLower) Then

                    Return nombreX

                End If

            Next

        Catch ex As Exception

        End Try

        Return String.Empty

    End Function
End Class

Imports Entidades
Imports System.Threading
Imports System.Globalization
Imports System.IO

'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Xml.Serialization

Public Class frmProcesoLoteC
    Public oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    '
    Private listaDocumentoPorUsuario As List(Of Entidades.DocumentoTipo)
    Private listaFCs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
    Private listaREs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
    Private listaNCs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
    Private listaNDs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaDebito)
    Private listaGRs As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTGuiaRemision)
    Public listaDetalleArtiulos As New List(Of Entidades.DetalleArticulo)
    Private oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura
    Private oNotaDeCredito As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito
    Private oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion

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
    Dim GRPO As SAPbobsCOM.Documents
    Dim cbxTipo As SAPbouiCOM.ComboBox
    ''
    Dim oListaFacturaVenta As New List(Of Entidades.FacturaVenta)
    Dim oFacturaVenta As Entidades.FacturaVenta

    Dim _oDocumento As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
    Dim men As String = ""
    Dim Mensaje_Error As String = ""
    Dim gcss As SAPbouiCOM.CommonSetting
    Dim diasValidar As Integer = 0
    Dim ListaFila As New List(Of Integer)
    Dim ListaPRCont As New List(Of Integer)
    Dim ListaPRError As New List(Of Integer)




    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormularioProcesoLote()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        rsboApp.SetStatusBarMessage(NombreAddon + " - Consultando, Espere Por favor!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        If RecorreFormulario(rsboApp, "frmProcesoLoteC") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmProcesoLoteC.srf"
        xmlDoc.Load(strPath)
        Utilitario.Util_Log.Escribir_Log("strPath: " + strPath.ToString, "frmProcesoLoteC")
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmProcesoLoteC").Close()
                xmlDoc = Nothing
                Utilitario.Util_Log.Escribir_Log("ERROR AL CARGAR SRF: " + exx.Message.ToString, "frmProcesoLoteC")
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmProcesoLoteC")

            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO

            oForm.Freeze(True)
            Utilitario.Util_Log.Escribir_Log("Freeze True", "frmProcesoLoteC")
            diasValidar = IIf(Functions.VariablesGlobales._DiasValidarProcesoLote = "", 5, CInt(Functions.VariablesGlobales._DiasValidarProcesoLote))
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
            cmbTipo.ValidValues.Add("01", "Factura")
            cmbTipo.ValidValues.Add("07", "Comp. de Retención")
            cmbTipo.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim lnkPr As SAPbouiCOM.LinkedButton
            lnkPr = oForm.Items.Item("lnkPr").Specific
            lnkPr.LinkedObjectType = 2
            lnkPr.Item.LinkTo = "txtRuc"

            Dim focus As SAPbouiCOM.EditText
            focus = oForm.Items.Item("focus").Specific
            focus.Item.Visible = False

            oForm.DataSources.UserDataSources.Add("dtFI", SAPbouiCOM.BoDataType.dt_DATE, 20)
            oForm.DataSources.UserDataSources.Add("dtFF", SAPbouiCOM.BoDataType.dt_DATE, 20)

            Dim txtFechaD As SAPbouiCOM.EditText
            txtFechaD = oForm.Items.Item("txtFechaD").Specific
            txtFechaD.DataBind.SetBound(True, "", "dtFI")

            Dim txtFechaH As SAPbouiCOM.EditText
            txtFechaH = oForm.Items.Item("txtFechaH").Specific
            txtFechaH.DataBind.SetBound(True, "", "dtFF")

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try
            Utilitario.Util_Log.Escribir_Log("AÑADIENDO GRILLA", "frmProcesoLoteC")
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

            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("FechaEmisionFactura", SAPbouiCOM.BoFieldsType.ft_Date, 20)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Mapeado", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Borrador", SAPbouiCOM.BoFieldsType.ft_Integer, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Sucursal", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Observación", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("IdDoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Seleccionar", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            Utilitario.Util_Log.Escribir_Log("CARGANDO DOCUMENTOSD", "frmProcesoLoteC")
            cargarDocumentos()
            Utilitario.Util_Log.Escribir_Log("DOCUMENTOSD CARGADOS", "frmProcesoLoteC")


            ' Obtener el tamaño de la pantalla
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

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
            Utilitario.Util_Log.Escribir_Log("Freeze False", "frmProcesoLoteC")
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
    Private Sub VerificarSN()

        Try
            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oDatable As SAPbouiCOM.DataTable
            'pintar filas
            Dim gcss As SAPbouiCOM.CommonSetting
            gcss = oGrid.CommonSetting
            ' oDatable = oForm.DataSources.DataTables.Item("dtDocs")
            oDatable = oGrid.DataTable
            Dim x As Integer, y As Integer
            Dim indexgrid As Integer = 0


            Dim QueryExisteCliente As String = ""
            Dim QueryExisteProveedor As String = ""

            For x = 0 To oDatable.Rows.Count - 1
                For y = 1 To oGrid.Rows.Count
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    If indexgrid = x Then

                        Exit For
                    End If

                Next
                ofila = indexgrid
                cbxTipo = oForm.Items.Item("cbxTipo").Specific
                Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                Dim sNomProveedor As String = oDataTable.GetValue(5, ofila).ToString()
                If cbxTipo.Value = "07" Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
                    Else
                        QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
                    End If
                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                    Try
                        If String.IsNullOrEmpty(sCardCode) Then
                            Dim respuesta = rsboApp.MessageBox(NombreAddon + " - No existe Cliente " + sNomProveedor + " , Desea Crearlo ?", 1, "OK", "Cancelar")
                            If respuesta = 1 Then
                                rsboApp.ActivateMenuItem("2561")
                                oForm = rsboApp.Forms.GetFormByTypeAndCount(134, -1)
                                oForm.Select()
                                rsboApp.ActivateMenuItem("1282") 'NUEVO

                                Exit Sub
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Else
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sRUC + "'"
                    Else
                        QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sRUC + "'"
                    End If
                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                    Try
                        If String.IsNullOrEmpty(sCardCode) Then
                            Dim respuesta = rsboApp.MessageBox(NombreAddon + " - No existe Proveedor " + sNomProveedor + " , Desea Crearlo ?", 1, "OK", "Cancelar")
                            If respuesta = 1 Then
                                rsboApp.ActivateMenuItem("2561")
                                oForm = rsboApp.Forms.GetFormByTypeAndCount(134, -1)
                                oForm.Select()
                                rsboApp.ActivateMenuItem("1282") 'NUEVO
                                Exit Sub
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                End If
                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next
            rsboApp.StatusBar.SetText("(SAED) Verificacion Completa", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)



        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Ocurrio un error al llamar la funcion ConsultarEstados " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    Private Sub HabilitarBTN()

        Try
            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oDatable As SAPbouiCOM.DataTable
            'pintar filas
            Dim gcss As SAPbouiCOM.CommonSetting
            gcss = oGrid.CommonSetting
            ' oDatable = oForm.DataSources.DataTables.Item("dtDocs")
            oDatable = oGrid.DataTable
            Dim x As Integer, y As Integer
            Dim indexgrid As Integer = 0


            Dim QueryExisteCliente As String = ""
            Dim QueryExisteProveedor As String = ""
            Dim sinError As Boolean = True
            Dim cont As Integer = 0
            Dim cont2 As Integer = 0
            Dim nRegistros As String = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Registros_por_paginas")
            If nRegistros = "" Then
                RegistrosXPaginas = 15
            Else
                RegistrosXPaginas = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Registros_por_paginas")

            End If

            For x = 0 To oDatable.Rows.Count - 1
                For y = 1 To oGrid.Rows.Count
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    If indexgrid = x Then

                        Exit For
                    End If
                Next
                ofila = indexgrid
                cbxTipo = oForm.Items.Item("cbxTipo").Specific

                Dim sPreliminar As String = oDataTable.GetValue(12, ofila).ToString()

                If cbxTipo.Value = "07" Then
                    If sPreliminar <> "0" Then
                        cont = cont + 1
                    End If
                    cont2 = cont
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Else
                    Dim borrador As String = oDataTable.GetValue(12, ofila).ToString()
                    If sPreliminar <> "0" Then
                        cont = cont + 1
                    End If
                    cont2 = cont
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                End If
                ' rsboApp.MessageBox("numero de elementos en la grilla = " & CStr(oDatable.Rows.Count) & "   " & oDatable.GetValue("EstadoDoc", x) & "  " & oDatable.GetValue("CardName", x))
            Next
            If cont <> RegistrosXPaginas Then
                sinError = False
            End If
            If cont <> cont2 Then
                sinError = False
            End If
            If cont = 0 And cont2 = 0 Then
                sinError = False
            End If

            'rsboApp.StatusBar.SetText("(SAED) Verificacion Completa", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If sinError Then
                oForm.Items.Item("btnProc").Enabled = True
            Else
                oForm.Items.Item("btnProc").Enabled = False
            End If
        Catch ex As Exception
            'rsboApp.SetStatusBarMessage("Ocurrio un error al llamar la funcion ConsultarEstados " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try


    End Sub

    Private Sub ProcesoLote()

        Try
            oForm = rsboApp.Forms.Item("frmProcesoLoteC")
            'oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            Dim gcss As SAPbouiCOM.CommonSetting
            gcss = oGrid.CommonSetting
            Dim DocEntryFacturaRecibida_UDO As String = 0
            Dim Exitoso As Boolean = False
            Dim QryRuc As String = ""
            Dim indexgrid As Integer = 0
            Dim x As Integer, y As Integer


            'For k As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            '    gcss.SetRowBackColor(k + 2, RGB(245, 238, 81))
            '    'Dim sNomProveedor As String = oGrid.GetValue(5, k + 1).ToString()
            '    Dim sNomProveedor As String = oGridDet.GetValue(5, k).ToString()
            '    'Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
            '    'odt.SetValue(14, k - 1, "Error en funcion Guarda_DocumentoRecibido_Factura: " + Mensaje_Error.ToString)
            '    'oForm.DataSources.DataTables.Item("dtDocs").SetValue(14, k - 1, "No existe Cliente " + sNomProveedor)
            '    oGrid.DataTable.SetValue(14, k, sNomProveedor)
            'Next
            Dim k As Integer = 0
            If ListaFila.Count > 0 Then
                ListaFila.Sort()
            End If

            If ListaFila.Count > 0 Then

                For Each k In ListaFila
                    'For k As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    cbxTipo = oForm.Items.Item("cbxTipo").Specific

                    If cbxTipo.Value = "07" Then
                        Dim Borrador As String = oGridDet.GetValue(12, k).ToString()

                        If Borrador = "0" Then
                            If oGridDet.GetValue("Seleccionar", k) = "Y" Then

                                'End If
                                Dim sRUC As String = oGridDet.GetValue("RUC", k)
                                Dim sNomProveedor As String = oGridDet.GetValue(5, k).ToString()

                                Dim QueryExisteCliente As String
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
                                Else
                                    QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
                                End If
                                Dim CodigoCliente = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                                If CodigoCliente <> "" Then

                                    oRetencion = Nothing
                                    Dim iIdDocEdoc As Long = Long.Parse(oGridDet.GetValue(15, k).ToString())
                                    Dim iIdPreliminar As Long = Long.Parse(oGridDet.GetValue(12, k).ToString())
                                    Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                    WS.Url = Functions.VariablesGlobales._WS_Recepcion

                                    If iIdPreliminar.ToString <> "" Or (iIdPreliminar.ToString = "" And Functions.VariablesGlobales._ContabilizarPRPL = "Y") Then

                                        oGrid.Rows.SelectedRows.Add(k)
                                        gcss.SetRowBackColor(k + 1, RGB(245, 238, 81))

                                        ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()

                                        oRetencion = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)
                                        Mensaje_Error = ""
                                        If Guarda_DocumentoRecibido_RE(DocEntryFacturaRecibida_UDO, oRetencion, CodigoCliente) Then

                                            Dim sDocEntryPreliminar As String = "0"
                                            'Exitoso = CrearPagoRecibido(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                                            Exitoso = CrearPagoRecibido_E_ONormal(CodigoCliente, oRetencion, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)

                                            If Exitoso Then
                                                If Functions.VariablesGlobales._ContabilizarPRPL = "Y" Then
                                                    Try
                                                        Dim sClaveAcceso As String = oGridDet.GetValue(7, k).ToString()

                                                        Dim idFacturaGS As String = ""
                                                        Try

                                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + DocEntryFacturaRecibida_UDO + "'", "U_IdGS", "")
                                                            Else
                                                                idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + DocEntryFacturaRecibida_UDO, "U_IdGS", "")
                                                            End If
                                                        Catch ex As Exception
                                                            Utilitario.Util_Log.Escribir_Log("Error al obtener idDocumentoRecibido_UDO: " + ex.Message.ToString(), "ProcesoLote")
                                                        End Try

                                                        If ActualizadoEstado_DocumentoRecibido_RE(DocEntryFacturaRecibida_UDO, "docFinal", sClaveAcceso) Then
                                                            ActualizadoEstadoSincronizado_DocumentoRecibido_RE(DocEntryFacturaRecibida_UDO, 1, sClaveAcceso)
                                                            MarcarVisto(Integer.Parse(idFacturaGS), 2, mensaje, DocEntryFacturaRecibida_UDO, sClaveAcceso)
                                                            ListaPRCont.Add(k)
                                                            oForm.Freeze(True)
                                                            oGridDet.SetValue("Seleccionar", k, "N")
                                                            gcss.SetRowBackColor(k + 1, 255000)
                                                            oForm.Freeze(False)
                                                        End If
                                                    Catch ex As Exception
                                                        oForm.Freeze(True)
                                                        oGridDet.SetValue("Seleccionar", k, "N")
                                                        gcss.SetRowBackColor(k + 1, RGB(255, 0, 0))
                                                        oForm.Freeze(False)
                                                        rsboApp.SetStatusBarMessage("Error al contabilizar pago " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                    End Try
                                                Else
                                                    Try
                                                        If oRetencion.FechaAutorizacion <= Date.Now Then
                                                            oForm.Freeze(True)
                                                            Dim FecEmiFac As Date = oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener
                                                            Dim FecAutRet As Date = oRetencion.FechaAutorizacion
                                                            Dim FecEmiRet As Date = oRetencion.FechaEmision

                                                            If Date.Now.Day > 5 And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                                                oGrid.DataTable.SetValue(12, k, Integer.Parse(sDocEntryPreliminar))
                                                                oGrid.DataTable.SetValue(14, k, "Revisar: la fecha de autorizacion es mayor al 5 del mes actual")
                                                                gcss.SetRowBackColor(k + 1, RGB(251, 163, 26))
                                                            ElseIf FecEmiFac.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiRet.Month < Date.Now.Month Then

                                                                oGrid.DataTable.SetValue(12, k, Integer.Parse(sDocEntryPreliminar))
                                                                oGrid.DataTable.SetValue(14, k, "Revisar: este documento se autorizó el mes pasado")
                                                                gcss.SetRowBackColor(k + 1, RGB(251, 163, 26))
                                                            Else

                                                                oGrid.DataTable.SetValue(12, k, Integer.Parse(sDocEntryPreliminar))
                                                                oGrid.DataTable.SetValue(14, k, "Creado Preliminar con Éxito")
                                                                gcss.SetRowBackColor(k + 1, 255000)
                                                            End If
                                                            oForm.Freeze(False)
                                                        End If
                                                    Catch ex As Exception
                                                        oForm.Freeze(True)
                                                        oGrid.DataTable.SetValue(14, k, "Error al validar fechas" + ex.ToString)
                                                        gcss.SetRowBackColor(k + 1, RGB(245, 0, 0))
                                                        oForm.Freeze(False)
                                                        Mensaje_Error = "Error al validar fechas" + ex.ToString
                                                    End Try
                                                End If
                                            Else
                                                If Mensaje_Error <> "" Then
                                                    oForm.Freeze(True)
                                                    oGrid.DataTable.SetValue(14, k, "Error en funcion CrearPagoRecibido: " + Mensaje_Error.ToString)
                                                    gcss.SetRowBackColor(k + 1, RGB(245, 0, 0))
                                                    oGridDet.SetValue("Seleccionar", k, "N")
                                                    ListaPRError.Add(k)
                                                    oForm.Freeze(False)
                                                End If
                                            End If

                                        Else

                                            If Mensaje_Error <> "" Then
                                                oForm.Freeze(True)
                                                oGrid.DataTable.SetValue(14, k, "Error en funcion Guarda_DocumentoRecibido_RE: " + Mensaje_Error.ToString)
                                                gcss.SetRowBackColor(k + 1, RGB(245, 0, 0))
                                                oGridDet.SetValue("Seleccionar", k, "N")
                                                ListaPRError.Add(k)
                                                oForm.Freeze(False)
                                            End If

                                        End If
                                    End If

                                Else
                                    oForm.Freeze(True)
                                    oGrid.DataTable.SetValue(14, k, "No existe Cliente " + sNomProveedor + " con Ruc " + sRUC)
                                    gcss.SetRowBackColor(k + 1, RGB(245, 0, 0))
                                    oGridDet.SetValue("Seleccionar", k, "N")
                                    ListaPRError.Add(k)
                                    Mensaje_Error = "No existe Cliente " + sNomProveedor + " con Ruc " + sRUC
                                    oForm.Freeze(False)
                                End If
                            End If
                        End If

                    Else
                        'facturas
                        Mensaje_Error = ""
                        Dim Borrador As String = oGridDet.GetValue(12, k).ToString()

                        If Borrador = "0" Then

                            If oGridDet.GetValue("Seleccionar", k) = "Y" Then

                                'End If
                                Dim sRUC As String = oGridDet.GetValue("RUC", k)
                                Dim sNomProveedor As String = oGridDet.GetValue(5, k).ToString()
                                Dim QueryExisteProveedor As String = ""

                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sRUC + "'"
                                Else
                                    QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sRUC + "'"
                                End If
                                sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                                If sCardCode <> "" Then
                                    Try

                                        oFactura = Nothing
                                        Dim iIdDocEdoc As Long = Long.Parse(oGridDet.GetValue(15, k).ToString())
                                        Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                        WS.Url = Functions.VariablesGlobales._WS_Recepcion
                                        ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                                        oFactura = WS.ConsultarFactura_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)


                                        If Guarda_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO, oFactura) Then
                                            Dim sDocEntryPreliminar As String = "0"

                                            Exitoso = CrearFacturaPremilinarServicio(sCardCode, oFactura, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                                            Try
                                                If Exitoso = True Then
                                                    ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                                    oForm.Freeze(True)
                                                    Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                    odt.SetValue(12, k, Integer.Parse(sDocEntryPreliminar))
                                                    odt.SetValue(14, k, "Creado Preliminar con Éxito")
                                                    gcss.SetRowBackColor(k + 1, 255000)
                                                    oGridDet.SetValue("Seleccionar", k, "N")
                                                    ListaPRCont.Add(k)
                                                    oForm.Freeze(False)
                                                    ' END ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                                    'rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                Else

                                                    If Mensaje_Error <> "" Then
                                                        oForm.Freeze(True)
                                                        ListaPRError.Add(k)
                                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt.SetValue(14, k, "Error al crear el documento Preliminar " + Mensaje_Error.ToString)
                                                        gcss.SetRowBackColor(k + 1, RGB(255, 0, 0))
                                                        oGridDet.SetValue("Seleccionar", k, "N")
                                                        oForm.Freeze(False)
                                                    End If

                                                End If

                                            Catch ex As Exception
                                                oForm.Freeze(True)
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(14, k, "Error: " + ex.Message.ToString)
                                                oGridDet.SetValue("Seleccionar", k, "N")
                                                gcss.SetRowBackColor(k + 1, RGB(255, 0, 0))
                                                oForm.Freeze(False)
                                            End Try
                                        Else

                                            If Mensaje_Error <> "" Then
                                                oForm.Freeze(True)
                                                ListaPRError.Add(k)
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(14, k, "Error en funcion Guarda_DocumentoRecibido_Factura: " + Mensaje_Error.ToString)
                                                gcss.SetRowBackColor(k + 1, RGB(255, 0, 0))
                                                oGridDet.SetValue("Seleccionar", k, "N")
                                                oForm.Freeze(False)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        Utilitario.Util_Log.Escribir_Log("Error al generar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                                        'oForm.Freeze(False)
                                    End Try
                                Else
                                    oForm.Freeze(True)
                                    Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                    odt.SetValue(14, k, "No existe Proveedor " + sNomProveedor + " con Ruc " + sRUC)
                                    gcss.SetRowBackColor(k + 1, RGB(255, 0, 0))
                                    oGridDet.SetValue("Seleccionar", k, "N")
                                    Mensaje_Error = "No existe Proveedor " + sNomProveedor + " con Ruc " + sRUC
                                    oForm.Freeze(False)
                                End If

                            End If
                        End If
                    End If

                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Next

            Else
                rsboApp.StatusBar.SetText(NombreAddon + " - Debe de seleccionar documentos para inciiar con el proceso..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

            'Dim j As Integer = 0
            'For Each j In ListaPRError
            '    gcss.SetRowBackColor(j + 1, -1)
            'Next
            ListaPRError.Clear()


            Dim rowsToDelete As List(Of Integer) = ListaPRCont
            rowsToDelete.Sort()
            rowsToDelete.Reverse()

            For Each rowIndex As Integer In rowsToDelete
                If rowIndex >= 0 And rowIndex < oGridDet.Rows.Count Then
                    oGridDet.Rows.Remove(rowIndex)
                End If
            Next

            'If ListaPRCont.Count > 0 Then
            '    oForm.Freeze(True)
            '    For Each k In ListaPRCont

            '        oGridDet.Rows.Remove(k)
            If cbxTipo.Value = "01" Then
                CargaDocumentosFormato("FE") 'se reordena las filas por eso tengo que validar otros datos como sn y folio ya que al eliminar las filas se cambian de orden
            ElseIf cbxTipo.Value = "04" Then
                CargaDocumentosFormato("NE")
            ElseIf cbxTipo.Value = "07" Then
                CargaDocumentosFormato("RE")
            End If
            '        ' oForm.Freeze(False)

            '    Next
            'End If


            If ListaPRCont.Count > 0 Then
                ListaPRCont.Clear()
                'ListaPRCont = Nothing
            End If

            If ListaFila.Count > 0 Then
                ListaFila.Clear()
                'ListaFila = Nothing
            End If


            For L As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                gcss.SetRowBackColor(L + 1, -1)
            Next

            If Mensaje_Error = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado con errores..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If



            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Ocurrio un error en el PL " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub CrearDocPorLote()

        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oDatable As SAPbouiCOM.DataTable
            'pintar filas

            gcss = oGrid.CommonSetting
            ' oDatable = oForm.DataSources.DataTables.Item("dtDocs")
            oDatable = oGrid.DataTable
            Dim x As Integer, y As Integer
            Dim nombre_estado As String = ""
            Dim ss_tipotabla As String = ""
            Dim identificador As Integer = 0
            Dim indexgrid As Integer = 0

            Dim DocEntryFacturaRecibida_UDO As String = 0
            Dim Exitoso As Boolean = False

            Dim _fila As Integer

            cbxTipo = oForm.Items.Item("cbxTipo").Specific

            Dim QueryExisteCliente As String = ""
            Dim QueryExisteProveedor As String = ""
            Dim sinError As Boolean = True
            For x = 0 To oDatable.Rows.Count - 1

                For y = 1 To oGrid.Rows.Count
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    If indexgrid = x Then
                        gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                        Exit For
                    End If

                Next
                Try
                    ofila = indexgrid
                    Dim sPreliminar As String = oDataTable.GetValue(12, ofila).ToString()
                    Dim sNomProveedor As String = oDataTable.GetValue(5, ofila).ToString()
                    If sPreliminar = "0" Then
                        _fila = indexgrid
                        Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                        If cbxTipo.Value = "07" Then

                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
                            Else
                                QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
                            End If
                            sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                            If sCardCode <> "" Then

                                oRetencion = Nothing
                                Dim iIdDocEdoc As Long = Long.Parse(oDataTable.GetValue(15, ofila).ToString())
                                Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                WS.Url = Functions.VariablesGlobales._WS_Recepcion
                                ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                                oRetencion = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)



                                If Guarda_DocumentoRecibido_RE(DocEntryFacturaRecibida_UDO, oRetencion, sCardCode) Then
                                    Dim sDocEntryPreliminar As String = "0"
                                    Exitoso = CrearPagoRecibido(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                                    Try
                                        If Exitoso = True Then
                                            ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                            Try
                                                If oRetencion.FechaAutorizacion <= Date.Now Then

                                                    Dim FecEmiFac As Date = oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener
                                                    Dim FecAutRet As Date = oRetencion.FechaAutorizacion
                                                    Dim FecEmiRet As Date = oRetencion.FechaEmision

                                                    If Date.Now.Day > 5 And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                                        Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt2.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                        odt2.SetValue(14, _fila, "Revisar: la fecha de autorizacion es mayor al 5 del mes actual")
                                                        gcss.SetRowBackColor(y + 1, RGB(251, 163, 26))
                                                    ElseIf FecEmiFac.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiRet.Month < Date.Now.Month Then
                                                        Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt2.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                        odt2.SetValue(14, _fila, "Revisar: este documento se autorizó el mes pasado")
                                                        gcss.SetRowBackColor(y + 1, RGB(251, 163, 26))
                                                    Else
                                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                        odt.SetValue(14, _fila, "Creado Preliminar con Éxito")
                                                        gcss.SetRowBackColor(y + 1, 255000)
                                                    End If
                                                End If
                                            Catch ex As Exception
                                                Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt2.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                odt2.SetValue(14, _fila, "Error al validar fechas" + ex.ToString)
                                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                            End Try



                                            ' END ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                            oFuncionesAddon.GuardaLOG("REE", oRetencion.ClaveAcceso, " Proceso terminado Exitosamente!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                                            Utilitario.Util_Log.Escribir_Log("PL - REE: " + oRetencion.ClaveAcceso.ToString(), "ProcesoLote")
                                        Else
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, _fila, "Error: " + men)
                                            gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                        End If

                                    Catch ex As Exception
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                        odt.SetValue(14, _fila, "Error: " + ex.Message.ToString)
                                        gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                        sinError = False
                                        Utilitario.Util_Log.Escribir_Log("Error al ingresar Docentry en el grid: " + ex.Message.ToString(), "ProcesoLote")
                                    End Try
                                End If
                            Else

                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                odt.SetValue(14, _fila, "No existe Cliente " + sNomProveedor + " con Ruc " + sRUC)
                                sinError = False
                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                            End If

                        Else
                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sRUC + "'"
                            Else
                                QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sRUC + "'"
                            End If
                            sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                            If sCardCode <> "" Then
                                Try

                                    oFactura = Nothing
                                    Dim iIdDocEdoc As Long = Long.Parse(oDataTable.GetValue(15, ofila).ToString())
                                    Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                    WS.Url = Functions.VariablesGlobales._WS_Recepcion
                                    ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                                    oFactura = WS.ConsultarFactura_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)


                                    If Guarda_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO, oFactura) Then
                                        Dim sDocEntryPreliminar As String = "0"

                                        Exitoso = CrearFacturaPremilinarServicio(sCardCode, oFactura, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                                        Try
                                            If Exitoso = True Then
                                                ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL

                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                odt.SetValue(14, _fila, "Creado Preliminar con Éxito")
                                                gcss.SetRowBackColor(y + 1, 255000)

                                                ' END ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                                oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, " Proceso terminado Exitosamente!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                Utilitario.Util_Log.Escribir_Log("PL - REE: " + oFactura.ClaveAcceso.ToString(), "ProcesoLote")
                                            Else
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(14, _fila, "Error al crear el documento Preliminar ")
                                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))

                                            End If

                                        Catch ex As Exception
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, _fila, "Error: " + ex.Message.ToString)
                                            gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                            sinError = False
                                            Utilitario.Util_Log.Escribir_Log("Error al ingresar Docentry en el grid: " + ex.Message.ToString(), "ProcesoLote")
                                        End Try
                                    End If
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al generar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                                End Try
                            Else
                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                odt.SetValue(14, _fila, "No existe Proveedor " + sNomProveedor + " con Ruc " + sRUC)
                                sinError = False
                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                            End If
                        End If

                    End If
                    'gcss.SetRowBackColor(y + 1, 255000)
                Catch ex As Exception
                    'sinError = False
                    gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                    Exit For
                End Try
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            Next
            'HABILITAR BOTON PROCESAR PRELIMINARES
            'If sinError Then

            '    oForm.Items.Item("btnProc").Enabled = True
            'End If
            rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Ocurrio un error en el PL " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub CrearDocPorLoteCorregido()

        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
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

            Dim DocEntryFacturaRecibida_UDO As String = 0
            Dim Exitoso As Boolean = False

            Dim _fila As Integer

            cbxTipo = oForm.Items.Item("cbxTipo").Specific

            Dim QueryExisteCliente As String = ""
            Dim QueryExisteProveedor As String = ""
            Dim sinError As Boolean = True
            For x = 0 To oDatable.Rows.Count - 1

                For y = 1 To oGrid.Rows.Count
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    If indexgrid = x Then
                        gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                        Exit For
                    End If

                Next
                Try
                    ofila = indexgrid
                    Dim sPreliminar As String = oDataTable.GetValue(12, ofila).ToString()
                    Dim sNomProveedor As String = oDataTable.GetValue(5, ofila).ToString()
                    If sPreliminar = "0" Then
                        _fila = indexgrid
                        Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                        If cbxTipo.Value = "07" Then

                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
                            Else
                                QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
                            End If
                            sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                            If sCardCode <> "" Then

                                oRetencion = Nothing
                                Dim iIdDocEdoc As Long = Long.Parse(oDataTable.GetValue(15, ofila).ToString())
                                Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                WS.Url = Functions.VariablesGlobales._WS_Recepcion
                                ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                                oRetencion = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)



                                If Guarda_DocumentoRecibido_RE(DocEntryFacturaRecibida_UDO, oRetencion, sCardCode) Then
                                    Dim sDocEntryPreliminar As String = "0"
                                    Exitoso = CrearPagoRecibido(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                                    Try
                                        If Exitoso = True Then
                                            ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                            Try
                                                If oRetencion.FechaAutorizacion <= Date.Now Then

                                                    Dim FecEmiFac As Date = oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener
                                                    Dim FecAutRet As Date = oRetencion.FechaAutorizacion
                                                    Dim FecEmiRet As Date = oRetencion.FechaEmision

                                                    If Date.Now.Day > 5 And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                                        Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt2.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                        odt2.SetValue(14, _fila, "Revisar: la fecha de autorizacion es mayor al 5 del mes actual")
                                                        gcss.SetRowBackColor(y + 1, RGB(251, 163, 26))
                                                    ElseIf FecEmiFac.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiRet.Month < Date.Now.Month Then
                                                        Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt2.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                        odt2.SetValue(14, _fila, "Revisar: este documento se autorizó el mes pasado")
                                                        gcss.SetRowBackColor(y + 1, RGB(251, 163, 26))
                                                    Else
                                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                        odt.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                        odt.SetValue(14, _fila, "Creado Preliminar con Éxito")
                                                        gcss.SetRowBackColor(y + 1, 255000)
                                                    End If
                                                End If
                                            Catch ex As Exception
                                                Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt2.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                odt2.SetValue(14, _fila, "Error al validar fechas" + ex.ToString)
                                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                            End Try



                                            ' END ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                            oFuncionesAddon.GuardaLOG("REE", oRetencion.ClaveAcceso, " Proceso terminado Exitosamente!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                                            Utilitario.Util_Log.Escribir_Log("PL - REE: " + oRetencion.ClaveAcceso.ToString(), "ProcesoLote")
                                        Else
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, _fila, "Error: " + men)
                                            gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                        End If

                                    Catch ex As Exception
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                        odt.SetValue(14, _fila, "Error: " + ex.Message.ToString)
                                        gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                        sinError = False
                                        Utilitario.Util_Log.Escribir_Log("Error al ingresar Docentry en el grid: " + ex.Message.ToString(), "ProcesoLote")
                                    End Try
                                End If
                            Else

                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                odt.SetValue(14, _fila, "No existe Cliente " + sNomProveedor + " con Ruc " + sRUC)
                                sinError = False
                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                            End If

                        Else
                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sRUC + "'"
                            Else
                                QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sRUC + "'"
                            End If
                            sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                            If sCardCode <> "" Then
                                Try

                                    oFactura = Nothing
                                    Dim iIdDocEdoc As Long = Long.Parse(oDataTable.GetValue(15, ofila).ToString())
                                    Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                                    WS.Url = Functions.VariablesGlobales._WS_Recepcion
                                    ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                                    oFactura = WS.ConsultarFactura_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)


                                    If Guarda_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO, oFactura) Then
                                        Dim sDocEntryPreliminar As String = "0"

                                        Exitoso = CrearFacturaPremilinarServicio(sCardCode, oFactura, sDocEntryPreliminar, DocEntryFacturaRecibida_UDO)
                                        Try
                                            If Exitoso = True Then
                                                ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL

                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(12, _fila, Integer.Parse(sDocEntryPreliminar))
                                                odt.SetValue(14, _fila, "Creado Preliminar con Éxito")
                                                gcss.SetRowBackColor(y + 1, 255000)

                                                ' END ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                                                oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, " Proceso terminado Exitosamente!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                                                Utilitario.Util_Log.Escribir_Log("PL - REE: " + oFactura.ClaveAcceso.ToString(), "ProcesoLote")
                                            Else
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(14, _fila, "Error al crear el documento Preliminar ")
                                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))

                                            End If

                                        Catch ex As Exception
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, _fila, "Error: " + ex.Message.ToString)
                                            gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                                            sinError = False
                                            Utilitario.Util_Log.Escribir_Log("Error al ingresar Docentry en el grid: " + ex.Message.ToString(), "ProcesoLote")
                                        End Try
                                    End If
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al generar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                                End Try
                            Else
                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                odt.SetValue(14, _fila, "No existe Proveedor " + sNomProveedor + " con Ruc " + sRUC)
                                sinError = False
                                gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                            End If
                        End If

                    End If
                    'gcss.SetRowBackColor(y + 1, 255000)
                Catch ex As Exception
                    'sinError = False
                    gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                    Exit For
                End Try
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            Next
            'HABILITAR BOTON PROCESAR PRELIMINARES
            'If sinError Then

            '    oForm.Items.Item("btnProc").Enabled = True
            'End If
            rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Ocurrio un error en el PL " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            'Dim typeEx, idForm As String
            'typeEx = oFuncionesB1.FormularioActivo(idForm)
            If pVal.FormTypeEx = "frmProcesoLoteC" Then
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
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
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
                                oForm = rsboApp.Forms.Item("frmProcesoLoteC")
                                Dim cbxTipo As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipo").Specific

                                oCons = oCFL.GetConditions()

                                Dim lbSocio As SAPbouiCOM.StaticText = oForm.Items.Item("lbSocio").Specific

                                Dim btnCrear As SAPbouiCOM.Button = oForm.Items.Item("btnCrear").Specific
                                Dim btnProc As SAPbouiCOM.Button = oForm.Items.Item("btnProc").Specific

                                If oCons.Count > 0 Then 'If there are already user conditions.
                                    If cbxTipo.Value = "07" Then ' SI ES 07, SIGNIFICA QUE ES RETENCION, POR ENDE PAGO RECIBIDO DE CLIENTE
                                        oCons.Item(oCons.Count - 1).CondVal = "C"
                                        lbSocio.Caption = "Cliente :"

                                        If Functions.VariablesGlobales._ContabilizarPRPL = "Y" Then
                                            btnCrear.Caption = "Contab. Retencion"
                                            btnProc.Item.Visible = False
                                        End If
                                    Else
                                        oCons.Item(oCons.Count - 1).CondVal = "S"
                                        lbSocio.Caption = "Proveedor :"

                                        If Functions.VariablesGlobales._ContabilizarPRPL = "Y" Then
                                            btnCrear.Caption = "Crear Preliminares"
                                            btnProc.Item.Visible = True
                                        End If
                                    End If
                                End If

                                oCFL.SetConditions(oCons)
                                Dim txtRuc As SAPbouiCOM.EditText = oForm.Items.Item("txtRuc").Specific
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

                                Case "btnVSN"
                                    VerificarSN()

                                Case "btnCrear"
                                    ' CrearDocPorLote()
                                    ProcesoLote()
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("btnSele").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                                Case "btnProc"
                                    'ProcesarPreliminar()
                                    ContabilizarPreliminar()

                                Case "btnSele"
                                    Dim contador As Integer = 0
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oForm.Items.Item("btnSele").Specific
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

                                Case "oGrid"
                                    ofila = pVal.Row
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLoteC")
                                    Dim oGrid As SAPbouiCOM.Grid = oFor.Items.Item("oGrid").Specific
                                    oGrid.Rows.SelectedRows.Add(ofila)
                                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                        ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                                        ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                                        Dim isChecked As String = oGrid.DataTable.GetValue("Seleccionar", ofila)
                                        If isChecked = "Y" Then
                                            If Not ListaFila.Contains(ofila) Then
                                                ListaFila.Add(ofila)
                                            End If
                                        Else
                                            If ListaFila.Contains(ofila) Then
                                                ListaFila.Remove(ofila)
                                            End If
                                        End If
                                    Next
                                Case "btnAnt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLoteC")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) - 1, TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oFor.Items.Item("btnSele").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                                Case "btnSig"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLoteC")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific
                                    Dim Seleccionar As SAPbouiCOM.Button = oFor.Items.Item("btnSele").Specific
                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbPA.Caption) + 1, TotalDocs)
                                    Seleccionar.Caption = "Seleccionar Todo"

                                    'PintarTransparente()

                                Case "btnPri"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLoteC")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), 1, TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oFor.Items.Item("btnSele").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                                Case "btnUlt"
                                    Dim oFor As SAPbouiCOM.Form
                                    oFor = rsboApp.Forms.Item("frmProcesoLoteC")
                                    Dim cbxTipo As SAPbouiCOM.ComboBox = oFor.Items.Item("cbxTipo").Specific
                                    Dim lbNP As SAPbouiCOM.StaticText = oForm.Items.Item("lbNP").Specific
                                    Dim lbPA As SAPbouiCOM.StaticText = oForm.Items.Item("lbPA").Specific

                                    llenarGrid(cbxTipo.Selected.Value.ToString(), RegistrosXPaginas, Integer.Parse(lbNP.Caption), Integer.Parse(lbNP.Caption), TotalDocs)
                                    Dim Seleccionar As SAPbouiCOM.Button
                                    Seleccionar = oFor.Items.Item("btnSele").Specific
                                    Seleccionar.Caption = "Seleccionar Todo"

                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pVal.ColUID = "RUC" And pVal.BeforeAction = True Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
                            Dim sCardCode As String = ""
                            Dim DocEntrySN As Integer
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
                                Dim QueryDocEntrySN As String = ""
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                    Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                                    Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sLicTradNum + "'"
                                        QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sLicTradNum + "'"
                                        QueryDocEntrySN = "SELECT ""DocEntry"" FROM ""OCRD"" where  ""LicTradNum"" = '" + sLicTradNum + "'"
                                    Else
                                        QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sLicTradNum + "'"
                                        QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sLicTradNum + "'"
                                        QueryDocEntrySN = "SELECT DocEntry FROM OCRD WITH(NOLOCK) where  LicTradNum = '" + sLicTradNum + "'"
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
                                DocEntrySN = 0
                                If tipoDocumento = "Retención de Cliente" Then
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                                Else
                                    sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                                End If

                                DocEntrySN = oFuncionesB1.getRSvalue(QueryDocEntrySN, "DocEntry", "")
                            Catch ex As Exception
                            End Try
                            If sCardCode <> "" Then
                                'rsboApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_BusinessPartner, "", DocEntrySN)
                                oForm.Items.Item("txtOpch").Specific.Value = sCardCode
                                oForm.Items.Item("lnkOpch").Click()
                                ''rsboApp.SendKeys("^+U")
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
                        'Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        '    If pVal.BeforeAction = False Then
                        '        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
                        '        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

                        '        'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                        '        'oGrid.Rows.SelectedRows.Add(ofila)

                        '        Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                        '        Dim QueryExisteProveedor As String = ""
                        '        Dim QueryExisteCliente As String = ""
                        '        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                        '            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '                QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""LicTradNum"" = '" + sRUC + "'"
                        '                QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
                        '            Else
                        '                QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND LicTradNum = '" + sRUC + "'"
                        '                QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
                        '            End If
                        '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '                QueryExisteProveedor = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""U_DOCUMENTO"" = '" + sRUC + "'"
                        '                QueryExisteCliente = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""U_DOCUMENTO"" = '" + sRUC + "'"
                        '            Else
                        '                QueryExisteProveedor = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'S' AND U_DOCUMENTO = '" + sRUC + "'"
                        '                QueryExisteCliente = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND U_DOCUMENTO = '" + sRUC + "'"
                        '            End If
                        '        End If

                        '        sCardCode = ""
                        '        Dim tipoDocumento As String = oDataTable.GetValue(0, ofila).ToString()
                        '        If tipoDocumento = "Retención de Cliente" Then
                        '            sCardCode = oFuncionesB1.getRSvalue(QueryExisteCliente, "CardCode", "")
                        '        Else
                        '            sCardCode = oFuncionesB1.getRSvalue(QueryExisteProveedor, "CardCode", "")
                        '        End If

                        '        If sCardCode = "" Then
                        '            Dim respuesta = rsboApp.MessageBox(NombreAddon + " - No existe el Socio de Negocio con el RUC/Cedula seleccionado: " + sRUC + ", Desea Crearlo ?", 1, "OK", "Cancelar")
                        '            If respuesta = 1 Then
                        '                rsboApp.ActivateMenuItem("2561")
                        '                oForm = rsboApp.Forms.GetFormByTypeAndCount(134, -1)
                        '                oForm.Select()
                        '                rsboApp.ActivateMenuItem("1282") 'NUEVO
                        '            End If
                        '        Else
                        '            Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                        '            Dim sNombre As String = oDataTable.GetValue(5, ofila).ToString()
                        '            Dim sMapeado As String = oDataTable.GetValue(11, ofila).ToString()
                        '            Dim iBorrador As Integer = Integer.Parse(oDataTable.GetValue(12, ofila).ToString())

                        '            rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Documento de " + sNombre + ", por favor espere..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        '            If oDataTable.GetValue(0, ofila).ToString() = "Factura" Then
                        '                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                        '                results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                        '                For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                        '                    oFactura = oFac
                        '                Next
                        '                Dim sQueryIdDocumento As String = ""
                        '                Dim idDocumentoRecibido_UDO As String = ""
                        '                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '                    sQueryIdDocumento = "SELECT ""U_SSIDDOCUMENTO"" FROM ""ODRF"" WHERE ""ObjType"" = 18 and ""DocEntry"" = " + iBorrador.ToString()
                        '                Else
                        '                    sQueryIdDocumento = "select U_SSIDDOCUMENTO from ODRF Where ObjType = 18 and DocEntry = " + iBorrador.ToString()
                        '                End If
                        '                If iBorrador = 0 Then
                        '                    ofrmDocumento.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oFactura, ofila)
                        '                Else
                        '                    idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                        '                    ofrmDocumento.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                        '                End If

                        '            ElseIf oDataTable.GetValue(0, ofila).ToString() = "Nota de Crédito" Then
                        '                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito)
                        '                results = listaNCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                        '                For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In results
                        '                    oNotaDeCredito = oNC
                        '                Next
                        '                Dim sQueryIdDocumento As String = ""
                        '                Dim idDocumentoRecibido_UDO As String = ""
                        '                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '                    sQueryIdDocumento = "SELECT ""U_SSIDDOCUMENTO"" FROM ""ODRF"" WHERE ""ObjType"" = 19 and ""DocEntry"" = " + iBorrador.ToString()
                        '                Else
                        '                    sQueryIdDocumento = "select U_SSIDDOCUMENTO from ODRF Where ObjType = 19 and DocEntry = " + iBorrador.ToString()
                        '                End If
                        '                If iBorrador = 0 Then
                        '                    ofrmDocumentoNC.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oNotaDeCredito, ofila)
                        '                Else
                        '                    idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                        '                    ofrmDocumentoNC.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                        '                End If

                        '            ElseIf oDataTable.GetValue(0, ofila).ToString() = "Retención de Cliente" Then
                        '                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                        '                results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                        '                For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                        '                    oRetencion = oRE
                        '                Next
                        '                Dim sQueryIdDocumento As String = ""
                        '                Dim idDocumentoRecibido_UDO As String = ""
                        '                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        '                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '                        sQueryIdDocumento += " SELECT B.""U_SSIDDOCUMENTO"""
                        '                        sQueryIdDocumento += " FROM ""OPDF"" B "
                        '                        sQueryIdDocumento += " WHERE B.""DocEntry"" = " + iBorrador.ToString()
                        '                        sQueryIdDocumento += " AND B.""U_SSCREADAR"" = 'SI'"
                        '                        sQueryIdDocumento += " AND B.""ObjType"" = 24"
                        '                    Else
                        '                        sQueryIdDocumento += " SELECT B.U_SSIDDOCUMENTO"
                        '                        sQueryIdDocumento += " FROM OPDF B "
                        '                        sQueryIdDocumento += " WHERE B.DocEntry = " + iBorrador.ToString()
                        '                        sQueryIdDocumento += " AND B.U_SSCREADAR = 'SI'"
                        '                        sQueryIdDocumento += " AND B.ObjType = 24"
                        '                    End If
                        '                Else
                        '                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '                        sQueryIdDocumento += " SELECT A.""U_SSIDDOCUMENTO"""
                        '                        sQueryIdDocumento += " FROM ""PDF3"" A INNER JOIN"
                        '                        sQueryIdDocumento += " ""OPDF"" B ON A.""DocNum"" = B.""DocEntry"" AND A.""U_SSCREADAR"" = 'SI'"
                        '                        sQueryIdDocumento += " WHERE B.""DocEntry"" = " + iBorrador.ToString()
                        '                        sQueryIdDocumento += " AND B.""ObjType"" = 24"
                        '                    Else
                        '                        sQueryIdDocumento += " SELECT A.U_SSIDDOCUMENTO"
                        '                        sQueryIdDocumento += " FROM PDF3 A INNER JOIN"
                        '                        sQueryIdDocumento += " OPDF B ON A.DocNum = B.DocEntry AND A.U_SSCREADAR = 'SI'"
                        '                        sQueryIdDocumento += " WHERE B.DocEntry = " + iBorrador.ToString()
                        '                        sQueryIdDocumento += " AND B.ObjType = 24"
                        '                    End If
                        '                End If

                        '                If iBorrador = 0 Then
                        '                    ofrmDocumentoRE.CargaFormularioDocumento(sRUC, sCardCode, sNombre, oRetencion, ofila)
                        '                Else
                        '                    idDocumentoRecibido_UDO = oFuncionesB1.getRSvalue(sQueryIdDocumento, "U_SSIDDOCUMENTO", "")
                        '                    ofrmDocumentoRE.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docPreliminar")
                        '                End If
                        '            End If
                        '            rsboApp.StatusBar.SetText(NombreAddon + " - Documento de " + sNombre + ", Cargado!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        '        End If
                        '    End If

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

                Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                oDataTable.Rows.Clear()

                If CargarDocumento() Then

                    'If Functions.VariablesGlobales._ContabilizarPRPL = "Y" Then
                    '    cbxTipo = oForm.Items.Item("cbxTipo").Specific
                    '    If cbxTipo.Value = "07" Then
                    '        Dim btnCrear As SAPbouiCOM.Button
                    '        btnCrear = oForm.Items.Item("btnCrear").Specific
                    '        btnCrear.Caption = "Contab. Ret"
                    '        btnCrear = oForm.Items.Item("btnProc").Specific
                    '        btnCrear.Item.Visible = False

                    '    Else
                    '        Dim btnCrear As SAPbouiCOM.Button
                    '        btnCrear = oForm.Items.Item("btnCrear").Specific
                    '        btnCrear.Caption = "Crear Preliminares"
                    '        btnCrear = oForm.Items.Item("btnProc").Specific
                    '        btnCrear.Item.Visible = True
                    '    End If
                    'End If
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

    Public Function CargarDocumento() As Boolean
        Try
            'CONSULTO EL RUC DE LA BASE ACTUAL            
            ' _RUC = oFuncionesB1.getRSvalue("SELECT TAXIDNUM FROM OADM", "TAXIDNUM", "")

            ' OBTENGO URL DEL SERVICIO DE RECEPCION
            'RegistrosXPaginas = oFuncionesAddon.getRSvalue("SELECT TOP 1 ""U_Valor"" from ""@GS_CONFD"" where U_Nombre='Registros_por_paginas'", "U_Valor", "")
            Dim nRegistros As String = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Registros_por_paginas")
            If nRegistros = "" Then
                RegistrosXPaginas = 10
            Else
                RegistrosXPaginas = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Registros_por_paginas")
            End If


            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")
            'FAMC cargo los estados parametrizados
            _WS_RecepcionCargaEstados = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Estados_docs")

            If String.IsNullOrEmpty(_WS_RecepcionCargaEstados) Then _WS_RecepcionCargaEstados = "01"
            'RegistrosXPaginas =  oUserTable.UserFields.Fields.Item("U_ws_RP").Value

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
            WS.Url = _WS_Recepcion

            'MANEJO DE PROXY
            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmProcesoLoteC")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""
            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmProcesoLoteC")

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
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
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

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                    'End If



                    Dim oFiltrosRecepcionEC As New wsEDoc_ConsultaRecepcion.ClsBusqueda
                    oFiltrosRecepcionEC.CiaTipoAlojamientoKey = Functions.VariablesGlobales._WS_RecepcionClave
                    oFiltrosRecepcionEC.RucProveedor = LicTradNum
                    oFiltrosRecepcionEC.Estado = _WS_RecepcionCargaEstados

                    Dim txtNumDoc As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDoc").Specific
                    If Not txtNumDoc.Value = "" Then
                        oFiltrosRecepcionEC.NumDocumento = txtNumDoc.Value
                    End If

                    ' FILTRO: Rango de Fechas
                    Dim txtFechaD As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaD").Specific
                    Dim txtFechaH As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaH").Specific
                    Dim dfechaDesde As Date
                    Dim dfechaHasta As Date
                    If Not oFuncionesB1.BobStringToDate(txtFechaD.Value, dfechaDesde) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Desde es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not oFuncionesB1.BobStringToDate(txtFechaH.Value, dfechaHasta) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Hasta es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not txtFechaH.Value = "" Then
                        oFiltrosRecepcionEC.FechaEmisionDesde = dfechaDesde
                        oFiltrosRecepcionEC.FechaEmisionHasta = dfechaHasta
                    Else
                        oFiltrosRecepcionEC.FechaEmisionDesde = Nothing
                        oFiltrosRecepcionEC.FechaEmisionHasta = Nothing
                    End If

                    Try
                        Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                        Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Recepcion_FC" + ".xml"
                        If System.IO.Directory.Exists(sRutaCarpeta) Then
                            Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de Busqueda", "frmDocumentosRecibidos")

                            Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_ConsultaRecepcion.ClsBusqueda))
                            Dim writer As TextWriter = New StreamWriter(sRuta)
                            x.Serialize(writer, oFiltrosRecepcionEC)
                            writer.Close()
                            Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda" + sRuta, "frmDocumentosRecibidos")
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                    'Dim z = WS.ConsultarFactura(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    Dim z = WS.ConsultarFactura_CabeceraBuscar(oFiltrosRecepcionEC, mensaje).ToList

                    If Not z Is Nothing Then
                        i = z.Count
                        listaFCs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura))
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

                    'Case "04"
                    '    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    '        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    '            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                    '        Else
                    '            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                    '        End If
                    '    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    '        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    '            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""U_DOCUMENTO"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                    '        Else
                    '            LicTradNum = oFuncionesB1.getRSvalue("SELECT U_DOCUMENTO FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "U_DOCUMENTO", "")
                    '        End If
                    '    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    '        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    '            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'S' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                    '        Else
                    '            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'S' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                    '        End If
                    '    End If

                    '    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific
                    '    Dim z = WS.ConsultarNC(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    '    If Not z Is Nothing Then
                    '        i = z.Count
                    '        listaNCs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito))
                    '        TotalDocs = listaNCs.Count
                    '        'CALCULO EL NUMERO DE PAGINAS EN BASE A LA CANTIDAD DE REGISTROS
                    '        If TotalDocs <= RegistrosXPaginas Then
                    '            NumeroPaginas = 1

                    '        Else
                    '            NumeroPaginas = Int(TotalDocs / RegistrosXPaginas)
                    '            residuo = (TotalDocs Mod RegistrosXPaginas)
                    '            If residuo > 0 Then
                    '                NumeroPaginas += 1
                    '            End If
                    '        End If
                    '        llenarGrid("04", RegistrosXPaginas, NumeroPaginas, 1, TotalDocs)
                    '        lbInfo.Caption = "Nº Total Nota de Crédito Recibidas :" + TotalDocs.ToString()

                    '    Else
                    '        lbInfo.Caption = "Nº Total Documentos Recibidos :" + TotalDocs.ToString()
                    '        Return False
                    '    End If

                Case "07"
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '    sQuery = "SELECT TOP 1 ""County"" from ""CRD1"" where ""LicTradNum"" = '" + sRUC + "' ORDER BY 1 DESC"
                        'Else
                        '    sQuery = "select top 1 County from CRD1 where LicTradNum ='" + sRUC + "' ORDER BY 1 DESC"
                        'End If

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

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT ""LicTradNum"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""CardCode"" = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        Else
                            LicTradNum = oFuncionesB1.getRSvalue("SELECT LicTradNum FROM OCRD WITH(NOLOCK) where CardType = 'C' AND CardCode = '" + txtRUC.Value.ToString() + "'", "LicTradNum", "")
                        End If
                    End If

                    Dim lbInfo As SAPbouiCOM.StaticText = oForm.Items.Item("lbInfo").Specific

                    Dim oFiltrosRecepcionEC As New wsEDoc_ConsultaRecepcion.ClsBusqueda
                    oFiltrosRecepcionEC.CiaTipoAlojamientoKey = Functions.VariablesGlobales._WS_RecepcionClave
                    oFiltrosRecepcionEC.RucProveedor = LicTradNum
                    oFiltrosRecepcionEC.Estado = _WS_RecepcionCargaEstados

                    Dim txtNumDoc As SAPbouiCOM.EditText = oForm.Items.Item("txtNumDoc").Specific
                    If Not txtNumDoc.Value = "" Then
                        oFiltrosRecepcionEC.NumDocumento = txtNumDoc.Value
                    End If

                    ' FILTRO: Rango de Fechas
                    Dim txtFechaD As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaD").Specific
                    Dim txtFechaH As SAPbouiCOM.EditText = oForm.Items.Item("txtFechaH").Specific
                    Dim dfechaDesde As Date
                    Dim dfechaHasta As Date
                    If Not oFuncionesB1.BobStringToDate(txtFechaD.Value, dfechaDesde) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Desde es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not oFuncionesB1.BobStringToDate(txtFechaH.Value, dfechaHasta) Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - El formato de la Fecha Hasta es incorrecto..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Return False
                        Exit Function
                    End If
                    If Not txtFechaH.Value = "" Then
                        oFiltrosRecepcionEC.FechaEmisionDesde = dfechaDesde
                        oFiltrosRecepcionEC.FechaEmisionHasta = dfechaHasta
                    Else
                        oFiltrosRecepcionEC.FechaEmisionDesde = Nothing
                        oFiltrosRecepcionEC.FechaEmisionHasta = Nothing
                    End If

                    Try
                        Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                        Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Recepcion_RT" + ".xml"
                        If System.IO.Directory.Exists(sRutaCarpeta) Then
                            Utilitario.Util_Log.Escribir_Log("Serializando, Parametros de Busqueda", "frmDocumentosRecibidos")

                            Dim x As XmlSerializer = New XmlSerializer(GetType(Entidades.wsEDoc_ConsultaRecepcion.ClsBusqueda))
                            Dim writer As TextWriter = New StreamWriter(sRuta)
                            x.Serialize(writer, oFiltrosRecepcionEC)
                            writer.Close()
                            Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda" + sRuta, "frmDocumentosRecibidos")
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "ManejoDeDocumentos")
                    End Try

                    ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()

                    'Dim z = WS.ConsultarRetencion(_WS_RecepcionClave, LicTradNum, _WS_RecepcionCargaEstados, mensaje).ToList
                    Dim z = WS.ConsultarRetencion_CabeceraBuscar(oFiltrosRecepcionEC, mensaje).ToList
                    If Not z Is Nothing Then
                        i = z.Count
                        listaREs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion))
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

                    'Case "05"
                    '    Dim z = WS.ConsultarND(_WS_RecepcionClave, LicTradNum, "01", mensaje).ToList
                    '    If Not z Is Nothing Then
                    '        i = z.Count
                    '        listaNDs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaDebito))
                    '    Else
                    '        Return False
                    '    End If
                    'Case "06"
                    '    Dim z = WS.ConsultarGR(_WS_RecepcionClave, LicTradNum, "01", mensaje).ToList
                    '    If Not z Is Nothing Then
                    '        i = z.Count
                    '        listaGRs = DirectCast(z, List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTGuiaRemision))
                    '    Else
                    '        Return False
                    '    End If
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

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
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
                'Dim listaOrdenada As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                'listaOrdenada = (From M In listaFCs Order By M.RazonSocial Ascending Select M).ToList
                'Dim sb As New System.Text.StringBuilder()
                'For Each str As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In listaOrdenada
                '    sb.AppendLine(str)
                'Next
                'listaOrdenada = listaFCs

                For Each oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In listaFCs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    Dim sRUC As String = oFactura.Ruc
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
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_FC")
                            sSucursal = sSucursal.Replace("RUC", sRUC.ToString)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        Else
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_FC")
                            sSucursal = sSucursal.Replace("RUC", sRUC.ToString)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        End If

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_NUM_AUTOR"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_NUM_AUTOR = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_NO_AUTORI"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_NO_AUTORI = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_SYP_NroAuto
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            Try
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NROAUTO"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SYP_NroAuto"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            End Try

                        Else
                            Try
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SYP_NROAUTO = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            Catch ex As Exception
                                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SYP_NroAuto = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                            End Try

                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '18' AND ""U_SS_NumAut"" = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '18' AND U_SS_NumAut = '" + oFactura.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    End If


                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, "DocEntry", "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Factura")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, oFactura.FechaEmision)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, oFactura.FechaAutorizacion)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oFactura.Establecimiento + "-" + oFactura.PuntoEmision + "-" + oFactura.Secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oFactura.Ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Trim(Left(oFactura.RazonSocial, 250)))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oFactura.ImporteTotal))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oFactura.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oFactura.AutorizacionSRI)


                    If Not String.IsNullOrEmpty(DocPreliminar) Then
                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("IdDoc", i, oFactura.IdFactura.ToString)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Seleccionar", i, "N")

                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Factura de " + oFactura.RazonSocial
                    rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Next
                i = 0

                'ElseIf tipoDoc = "04" Then
                '    oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                '    oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)

                '    For Each oNC As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito In listaNCs.GetRange(NumIni, RegistrosXPaginas)
                '        PendienteMapear = False
                '        Dim DocPreliminar As String = ""
                '        sQuery = ""
                '        Dim sQuerySucursal = ""
                '        Dim sRUC As String = oNC.Ruc
                '        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                '                sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_NC")
                '                sSucursal = sSucursal.Replace("RUC", sRUC)
                '                sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "County", "")
                '            Else
                '                sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_NC")
                '                sSucursal = sSucursal.Replace("RUC", sRUC)
                '                sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                '            End If

                '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                '                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_NUM_AUTOR"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '            Else
                '                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_NUM_AUTOR = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '            End If
                '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                '                sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_NO_AUTORI"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '            Else
                '                sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_NO_AUTORI = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '            End If
                '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                '            'U_SYP_NroAuto
                '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                '                Try
                '                    sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SYP_NROAUTO"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '                Catch ex As Exception
                '                    sQuery = "SELECT TOP 1 ""DocEntry"" from ""ODRF"" where ""ObjType"" = '19' AND ""U_SYP_NroAuto"" = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '                End Try

                '            Else
                '                Try
                '                    sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SYP_NROAUTO = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '                Catch ex As Exception
                '                    sQuery = "SELECT TOP 1 DocEntry from ODRF WITH(NOLOCK) where ObjType = '19' AND U_SYP_NroAuto = '" + oNC.AutorizacionSRI + "' ORDER BY 1 DESC"
                '                End Try

                '            End If
                '        End If

                '        DocPreliminar = oFuncionesB1.getRSvalue(sQuery, "DocEntry", "")
                '        If DocPreliminar = "0" Then
                '            DocPreliminar = ""
                '        End If

                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Nota de Crédito")
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, oNC.FechaEmision)

                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, oNC.FechaAutorizacion)

                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oNC.Establecimiento + "-" + oNC.PuntoEmision + "-" + oNC.Secuencial)
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oNC.Ruc)
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(oNC.RazonSocial, 250))
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oNC.ValorModificacion))
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oNC.ClaveAcceso)
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oNC.AutorizacionSRI)
                '        If Not String.IsNullOrEmpty(DocPreliminar) Then
                '            oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                '        End If
                '        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                '        i += 1
                '        m_oProgBar.Value = i + 1
                '        m_oProgBar.Text = NombreAddon + " - Cargando Nota de Crédito de " + oNC.RazonSocial

                '    Next
                '    i = 0
            ElseIf tipoDoc = "07" Then
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(RegistrosXPaginas)


                For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In listaREs.GetRange(NumIni, RegistrosXPaginas)
                    PendienteMapear = False
                    Dim DocPreliminar As String = ""
                    Dim sucursal As String = ""
                    sQuery = ""
                    sSucursal = ""
                    Dim sQuerySucursal = ""
                    Dim sRUC As String = oRE.Ruc
                    Dim IdCol As String = "DocEntry"
                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        IdCol = "DocNum"
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_RET")
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "County", "")
                        Else
                            sSucursal = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Adicional_RET")
                            sSucursal = sSucursal.Replace("RUC", sRUC)
                            sQuerySucursal = oFuncionesB1.getRSvalue(sSucursal, "City", "")
                        End If

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_CXS_NUM_AUTO_RETE"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_CXS_NUM_AUTO_RETE = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        IdCol = "DocNum"
                        'U_NUM_AUT
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_NUM_AUT"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_NUM_AUT = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        'U_FX_AUTO_RETENCION
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            '
                            sQuery = "SELECT TOP 1 ""DocEntry"" from ""OPDF"" where ""U_FX_AUTO_RETENCION"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocEntry from OPDF WITH(NOLOCK) where U_FX_AUTO_RETENCION = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        IdCol = "DocNum"

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            sQuery = "SELECT TOP 1 ""DocNum"" from ""PDF3"" where ""U_SS_AutRetRec"" = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        Else
                            sQuery = "SELECT TOP 1 DocNum from PDF3 WITH(NOLOCK) where U_SS_AutRetRec = '" + oRE.AutorizacionSRI + "' ORDER BY 1 DESC"
                        End If
                    End If

                    Dim numretener As String
                    Dim WS As New Entidades.wsEDoc_ConsultaRecepcion.WSRAD_KEY_CONSULTA
                    WS.Url = Functions.VariablesGlobales._WS_Recepcion
                    Dim iIdDocEdoc As Long = Long.Parse(oRE.IdRetencion.ToString.ToString())

                    Dim odetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                    odetalle = WS.ConsultarRetencion_Detalle(Functions.VariablesGlobales._WS_RecepcionClave, iIdDocEdoc, mensaje)
                    numretener = odetalle.ENTDetalleRetencion(0).NumDocRetener

                    DocPreliminar = oFuncionesB1.getRSvalue(sQuery, IdCol, "")
                    If DocPreliminar = "0" Then
                        DocPreliminar = ""
                    End If
                    'Dim _FechaEmisionDocRetener As String = Convert.ToDateTime(oRE.ENTDetalleRetencion(0).FechaEmisionDocRetener)
                    Dim _FechaEmisionDocRetener As String = odetalle.ENTDetalleRetencion(0).FechaEmisionDocRetener.ToString("dd/MM/yyyy")
                    'fecha.ToString("dd/MM/yyyy")
                    'sucursal = oFuncionesB1.getRSvalue(sSucursal, "Sucursal", "")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tipo", i, "Retención de Cliente")
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Fecha", i, oRE.FechaEmision)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaAutorizacion", i, oRE.FechaAutorizacion)

                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Folio", i, oRE.Establecimiento + "-" + oRE.PuntoEmision + "-" + oRE.Secuencial)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RUC", i, oRE.Ruc)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("RazonSocial", i, Left(oRE.RazonSocial, 250).Trim)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Valor", i, Convert.ToDouble(oRE.TotalRetencion))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("ClaveAcceso", i, oRE.ClaveAcceso)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumAutorizacion", i, oRE.AutorizacionSRI)

                    'oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRe", i, IIf(IsNothing(odetalle.NumDocRetener), "", odetalle.NumDocRetener))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("NumDocRetener", i, IIf(String.IsNullOrEmpty(numretener), "0", numretener))
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("FechaEmisionFactura", i, odetalle.ENTDetalleRetencion(0).FechaEmisionDocRetener)
                    If Not String.IsNullOrEmpty(DocPreliminar) Then

                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("Borrador", i, DocPreliminar)
                    End If
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Sucursal", i, sQuerySucursal)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("IdDoc", i, oRE.IdRetencion.ToString)
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("Seleccionar", i, "N")
                    i += 1
                    m_oProgBar.Value = i + 1
                    m_oProgBar.Text = NombreAddon + " - Cargando Retención Recibida de " + oRE.RazonSocial
                    'rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                Next
                i = 0
            End If

            CargaDocumentosFormato(IIf(tipoDoc = "07", "RE", "TD"))
            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + "Error - llenarGrid: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            '// Stop the progress bar

            m_oProgBar.Stop()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
            m_oProgBar = Nothing
        End Try
    End Sub

    Private Sub PintarTransparente()
        Dim oRow As SAPbouiCOM.GridRows
        Dim x As Integer, y As Integer
        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = oForm.Items.Item("oGrid").Specific
        Dim oDatable As SAPbouiCOM.DataTable
        gcss = oGrid.CommonSetting
        oDatable = oGrid.DataTable
        Dim indexgrid As Integer = 0

        For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

        Next

    End Sub
    Private Sub CargaDocumentosFormato(TipoDocumento As String)
        Try
            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            'oGrid.DataTable.ExecuteQuery(sQuery)

            oGrid.Columns.Item(0).Description = "TipoDoc"
            oGrid.Columns.Item(0).TitleObject.Caption = "TipoDoc"
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).Visible = False

            oGrid.Columns.Item(1).Description = "Fecha"
            oGrid.Columns.Item(1).TitleObject.Caption = "Fecha-Emisión"
            oGrid.Columns.Item(1).Editable = False
            'oGrid.Columns.Item(2).Visible = False

            oGrid.Columns.Item(2).Description = "Fecha"
            oGrid.Columns.Item(2).TitleObject.Caption = "Fecha-Autorización"
            oGrid.Columns.Item(2).Editable = False
            'oGrid.Columns.Item(2).Visible = False
            'Dim FchAut As String = ""
            'FchAut = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "MostrarFechaAutorizacion")
            'If FchAut = "Y" Then
            'oGrid.Columns.Item(2).Visible = True
            'Else

            'End If

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
            'oGrid.Columns.Item(6).Visible = False

            oGrid.Columns.Item(7).Description = "ClaveAcceso"
            oGrid.Columns.Item(7).TitleObject.Caption = "ClavedeAcceso"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).Visible = False

            oGrid.Columns.Item(8).Description = "NumAutorizacion"
            oGrid.Columns.Item(8).TitleObject.Caption = "Numero de Autorizacion"
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(8).Visible = False

            oGrid.Columns.Item(9).Description = "NumDocRetener"
            oGrid.Columns.Item(9).TitleObject.Caption = "Folio Factura"
            oGrid.Columns.Item(9).Editable = False
            If TipoDocumento = "RE" Then '
                oGrid.Columns.Item(9).Visible = True
            Else
                oGrid.Columns.Item(9).Visible = False
            End If

            oGrid.Columns.Item(10).Description = "FechaEmisionFactura"
            oGrid.Columns.Item(10).TitleObject.Caption = "FechaEmiFac"
            oGrid.Columns.Item(10).Editable = False
            If TipoDocumento = "RE" Then '
                oGrid.Columns.Item(10).Visible = True
            Else
                oGrid.Columns.Item(10).Visible = False
            End If

            'Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            'oEditTextColumn = oGrid.Columns.Item(10)
            'oEditTextColumn.LinkedObjectType = 22

            oGrid.Columns.Item(11).Description = "Mapeado"
            oGrid.Columns.Item(11).TitleObject.Caption = "Mapeado"
            oGrid.Columns.Item(11).Editable = False
            oGrid.Columns.Item(11).ForeColor = ColorTranslator.ToOle(Color.White)
            oGrid.Columns.Item(11).Visible = False

            oGrid.Columns.Item(12).Description = "Borrador"
            oGrid.Columns.Item(12).TitleObject.Caption = "Documento Preliminar"
            oGrid.Columns.Item(12).Editable = False



            Dim NomCamAdi = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "Nombre_CA")

            oGrid.Columns.Item(13).Description = "sucursal"
            oGrid.Columns.Item(13).TitleObject.Caption = NomCamAdi '"Sucursal"
            oGrid.Columns.Item(13).Editable = False
            'oGrid.Columns.Item(10).ForeColor = ColorTranslator.ToOle(Color.White)
            oGrid.Columns.Item(13).Visible = False

            oGrid.Columns.Item(14).Description = "Observacion"
            oGrid.Columns.Item(14).TitleObject.Caption = "Observacion"
            oGrid.Columns.Item(14).Editable = False

            oGrid.Columns.Item(15).Description = "IdDoc"
            oGrid.Columns.Item(15).TitleObject.Caption = "IdDoc"
            oGrid.Columns.Item(15).Editable = False
            oGrid.Columns.Item(15).Visible = False

            Dim oEditTextColumn2 As SAPbouiCOM.EditTextColumn
            oEditTextColumn2 = oGrid.Columns.Item(12)
            If TipoDocumento = "RE" Then ' SI ES RETENCION - EL PAGO RECIBIDO BORRADOR ES OTRA TABLA Y OTRO OBJTYPE
                oEditTextColumn2.LinkedObjectType = 140
            Else
                oEditTextColumn2.LinkedObjectType = 112
            End If


            oGrid.Columns.Item(16).Description = "Seleccionar"
            oGrid.Columns.Item(16).TitleObject.Caption = "Seleccionar"
            oGrid.Columns.Item(16).Editable = True
            oGrid.Columns.Item(16).Visible = True
            oGrid.Columns.Item(16).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            'oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()

            'Try
            '    For numfila As Integer = 0 To oGrid.Rows.Count - 1
            '        Dim valorFila As Integer = oGrid.GetDataTableRowIndex(numfila)
            '        If (valorFila <> -1) Then
            '            If (oGrid.DataTable.GetValue("Mapeado", valorFila) = "SI") Then
            '                oGrid.CommonSetting.SetCellBackColor(numfila + 1, 11, ColorTranslator.ToOle(Color.LightGreen))
            '            Else
            '                oGrid.CommonSetting.SetCellBackColor(numfila + 1, 11, ColorTranslator.ToOle(Color.Red))
            '            End If
            '        End If
            '    Next
            'Catch ex As Exception
            'Finally
            'End Try
            'campoRET()

            ''MANAMER
            'HabilitarBTN()

            oForm.Items.Item("btnProc").Enabled = True
            oForm.Freeze(False)
        Catch ex As Exception

        End Try
    End Sub

    Private Function ChooseFromList(ByRef pVal As SAPbouiCOM.ItemEvent, ByVal FormUID As String) As Boolean

        Dim bBubbleEvent As Boolean = True

        If FormUID = "frmProcesoLoteC" Then
            oForm = rsboApp.Forms.Item("frmProcesoLoteC")

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
            SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmProcesoLoteC")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            If SALIDA_POR_PROXY = "Y" Then
                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmProcesoLoteC")

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
            oManejoDocumentos.SetProtocolosdeSeguridad()
            If WS.MarcarVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                rsboApp.SetStatusBarMessage("Documento Marcado como Integrado : " + mensaje, SAPbouiCOM.BoMessageTime.bmt_Short, False)
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
            If eventInfo.FormUID = "frmProcesoLoteC" And eventInfo.ItemUID = "oGrid" Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                If eventInfo.BeforeAction Then
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    ofila = eventInfo.Row
                    Dim oFor As SAPbouiCOM.Form
                    oFor = rsboApp.Forms.Item("frmProcesoLoteC")
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

                    Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
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

                    typeEx = oFuncionesB1.FormularioActivo(idForm)

                    If typeEx = "frmProcesoLoteC" Then
                        If ofila > 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
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
                                    Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                    results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                    listaDetalleArtiulos.Clear()
                                    If results.Count > 0 Then
                                        For Each oFC As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFC.ENTDetalleFactura
                                                listaDetalleArtiulos.Add(New Entidades.DetalleArticulo(sClaveAcceso, oDetalle.CodigoPrincipal,
                                                                                                  oDetalle.CodigoAuxiliar, oDetalle.Descripcion, oDetalle.Cantidad, oDetalle.PrecioUnitario, oDetalle.Descuento, oDetalle.PrecioTotalSinImpuesto))
                                            Next
                                        Next
                                    End If
                                    ofrmMapeo.CargaFormularioMapeo(sRUC, sCardCode, sNombre, listaDetalleArtiulos, ofila, "FV")
                                End If

                            End If
                        Else
                            rsboApp.MessageBox(NombreAddon + " - Por favor dar click en filas que tengan información..")
                        End If
                    End If
                ElseIf pVal.MenuUID = "Marcar" Then
                    Dim typeEx, idForm As String

                    typeEx = oFuncionesB1.FormularioActivo(idForm)

                    If typeEx = "frmProcesoLoteC" Then
                        If ofila >= 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                            Dim tipoDocumento As String = oDataTable.GetValue(0, ofila).ToString()
                            If tipoDocumento = "Factura" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
                                results = listaFCs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If MarcarVisto(oFac.IdFactura, 1, mensaje) Then
                                            oDataTable.Rows.Remove(ofila)
                                            CargaDocumentosFormato("FE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                Next
                            ElseIf tipoDocumento = "Retención de Cliente" Then
                                Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
                                results = listaREs.FindAll(Function(column) column.ClaveAcceso = sClaveAcceso)
                                For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
                                    Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer marcar el documento como Integrado ?", 1, "OK", "Cancelar")
                                    If respuesta = 1 Then
                                        If MarcarVisto(oRE.IdRetencion, 2, mensaje) Then
                                            oDataTable.Rows.Remove(ofila)
                                            CargaDocumentosFormato("RE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Marcado como Integrado Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

#Region "Funciones Proceso Lote"

    Public Function Guarda_DocumentoRecibido_Factura(ByRef DocEntryFacturaRecibida_UDO As String, ByRef oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura) As Boolean

        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        Dim _ClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim numDoc As String = ""
        'Dim resultsFC As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFactura)
        'resultsFC = listaFCs.FindAll(Function(column) column.ClaveAcceso = _ClaveAcceso)
        'Dim ofactura As Object
        'For Each oFac As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura In resultsFC
        Try
            oFuncionesAddon.GuardaLOG("FACTURA", _ClaveAcceso, "Creando registro de Factura Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            'Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
            'Try
            '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '        sCardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM OCRD where ""LicTradNum"" = '" + sRUC + "' AND ""CardType"" = 'S' ", "CardCode", "")
            '    Else
            '        sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + sRUC + "' AND CardType = 'S' ", "CardCode", "")
            '    End If
            'Catch ex As Exception
            '    Utilitario.Util_Log.Escribir_Log("REE - error al obtener CardCode PL : " + ex.ToString, "frmProcesoLoteC")
            'End Try


            numDoc = oFac.Establecimiento + "-" + oFac.PuntoEmision + "-" + oFac.Secuencial

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralData.SetProperty("Code", conta)

            'oGeneralData.SetProperty("U_Ruta_pdf", rutaFC.ToString())
            oGeneralData.SetProperty("U_RUC", oFac.Ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oFac.RazonSocial.ToString(), 99))
            oGeneralData.SetProperty("U_CardCode", sCardCode.ToString())
            oGeneralData.SetProperty("U_Mapeado", "PL")
            oGeneralData.SetProperty("U_ClaAcc", oFac.ClaveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oFac.AutorizacionSRI.ToString())
            oGeneralData.SetProperty("U_FecAut", oFac.FechaAutorizacion.ToString())
            oGeneralData.SetProperty("U_NumDoc", numDoc.ToString())
            oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString()) 'sDocEntryPreliminar
            oGeneralData.SetProperty("U_SubSI", Convert.ToDouble(formatDecimal(oFac.TotalSinImpuesto.ToString())))
            oGeneralData.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(oFac.TotalDescuento.ToString())))
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(formatDecimal(oFac.ImporteTotal)))
            oGeneralData.SetProperty("U_rGast", "0")
            oGeneralData.SetProperty("U_rImp", "0")
            oGeneralData.SetProperty("U_rTotal", "0")
            oGeneralData.SetProperty("U_IdGS", oFac.IdFactura.ToString())
            oGeneralData.SetProperty("U_Sincro", 0)
            oGeneralData.SetProperty("U_Estado", "docPreliminar")
            oGeneralData.SetProperty("U_Tipo", "Factura de Servicio")

            For Each impFAC As Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto In oFac.ENTFacturaImpuesto
                Dim BaseImponibleIVA As Decimal = 0
                Dim BaseImponible0 As Decimal = 0
                Dim BaseImponibleNoObjeto As Decimal = 0
                Dim BaseImponibleExento As Decimal = 0
                Dim Iva As Decimal = 0
                Dim ICE As Decimal = 0
                Dim BaseImponibleICE As Decimal = 0
                If impFAC.Codigo = 2 Then
                    If impFAC.CodigoPorcentaje = 2 Or impFAC.CodigoPorcentaje = 3 Then
                        BaseImponibleIVA += impFAC.BaseImponible
                        Iva += impFAC.Valor
                    ElseIf impFAC.CodigoPorcentaje = 0 Then
                        BaseImponible0 += impFAC.BaseImponible
                    ElseIf impFAC.CodigoPorcentaje = 6 Then
                        BaseImponibleNoObjeto += impFAC.BaseImponible
                    ElseIf impFAC.CodigoPorcentaje = 7 Then
                        BaseImponibleExento += impFAC.BaseImponible
                    End If
                ElseIf impFAC.Codigo = 3 Then
                    ICE += impFAC.Valor
                End If

                oGeneralData.SetProperty("U_SubTot", Convert.ToDouble(formatDecimal(BaseImponibleIVA.ToString())))
                oGeneralData.SetProperty("U_Sub0", Convert.ToDouble(formatDecimal(BaseImponible0.ToString())))
                oGeneralData.SetProperty("U_SubNO", Convert.ToDouble(formatDecimal(BaseImponibleNoObjeto.ToString())))
                oGeneralData.SetProperty("U_SubEx", Convert.ToDouble(formatDecimal(BaseImponibleExento.ToString())))
                oGeneralData.SetProperty("U_ICE", Convert.ToDouble(formatDecimal(ICE.ToString())))
                oGeneralData.SetProperty("U_IVA", Convert.ToDouble(formatDecimal(Iva.ToString())))
                oGeneralData.SetProperty("U_rTades", Convert.ToDouble(formatDecimal(impFAC.BaseImponible.ToString())))
                oGeneralData.SetProperty("U_rPDesc", Convert.ToDouble(formatDecimal(impFAC.Tarifa.ToString())))
                oGeneralData.SetProperty("U_rDesc", Convert.ToDouble(formatDecimal(impFAC.Valor.ToString())))
            Next

            oGeneralData.SetProperty("U_Ruta_pdf", "")
            oGeneralData.SetProperty("U_Ruta_xml", "")

            oChildren = oGeneralData.Child("GS0_FVR")
            For Each detalleFac As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFac.ENTDetalleFactura
                Dim CodPrin As String = ""
                oChild = oChildren.Add

                If IsNothing(detalleFac.CodigoPrincipal) Then
                    CodPrin = "PL"
                Else
                    CodPrin = detalleFac.CodigoPrincipal.ToString
                End If

                oChild.SetProperty("U_CodPrin", Left(CodPrin, 99))
                If String.IsNullOrEmpty(detalleFac.CodigoAuxiliar) Then
                    oChild.SetProperty("U_CodAuxi", "SCA")
                Else
                    oChild.SetProperty("U_CodAuxi", Left(detalleFac.CodigoAuxiliar.ToString(), 99))
                End If
                oChild.SetProperty("U_CodSAP", "Servicio PL")
                oChild.SetProperty("U_Descripc", Left(detalleFac.Descripcion.ToString(), 99))
                oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(detalleFac.Cantidad.ToString())))
                oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(detalleFac.PrecioUnitario.ToString())))
                oChild.SetProperty("U_Desc", Convert.ToDouble(formatDecimal(detalleFac.Descuento.ToString())))
                oChild.SetProperty("U_Total", Convert.ToDouble(formatDecimal(detalleFac.PrecioTotalSinImpuesto.ToString())))
            Next

            'oChildren = oGeneralData.Child("GS1_FVR")
            'oChild = oChildren.Add
            'oChild.SetProperty("U_DocEntr", Integer.Parse(odt.GetValue(0, i).ToString()))
            'oChild.SetProperty("U_LineNu", Integer.Parse(odt.GetValue(1, i).ToString()))
            'oChild.SetProperty("U_ItemCode", odt.GetValue(2, i).ToString())
            'oChild.SetProperty("U_Descripc", odt.GetValue(3, i).ToString())
            'oChild.SetProperty("U_Cantid", Convert.ToDouble(formatDecimal(odt.GetValue(4, i).ToString())))
            'oChild.SetProperty("U_Precio", Convert.ToDouble(formatDecimal(odt.GetValue(5, i).ToString())))
            'oChild.SetProperty("U_DiscPr", Convert.ToDouble(formatDecimal(odt.GetValue(6, i).ToString())))
            'oChild.SetProperty("U_TaxCode", odt.GetValue(7, i).ToString())
            'oChild.SetProperty("U_lTotal", Convert.ToDouble(formatDecimal(odt.GetValue(8, i).ToString())))
            'oChild.SetProperty("U_ObjType", odt.GetValue(9, i).ToString())

            '_oFactura = oFac
            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            Utilitario.Util_Log.Escribir_Log("REE - se creo el preliminar : " + DocEntryFacturaRecibida_UDO, "frmProcesoLoteC")
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Se creo registro de Factura Recibida UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " PL Ocurrior un error al crear registro de Factura Recibida UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - PL Ocurrio un error al guardar Factura Recibida en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try

        'Next


    End Function
    Private Function CrearFacturaPremilinarServicio(ByRef sCardCode As String, ByVal oFactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura, ByRef sDocEntryPreliminar As String, ByVal DocEntryFVRecibida_UDO As String) As Boolean
        'Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        'Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        'Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()

        'Dim _sCardCode As String = sCardCode

        oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, "Creando Factura Preliminar de tipo: Servicio", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Factura por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim CodImp As String = ""
        Dim sQueryCodImp As String = ""
        'Dim CodImpV As String = ""
        Dim sQueryCodImpV As String = ""
        Dim CodImpV As String
        Try

            GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
            GRPO.CardCode = sCardCode
            ' GRPO.DocDate = ofactura.FechaEmision
            'Dim QueryFechaEmision = "Select ""U_SSFECH_EMI"" from OCRD Where ""CardCode"" =  '" + sCardCode + "'"
            'Dim QFE = oFuncionesB1.getRSvalue(QueryFechaEmision, "U_SSFECH_EMI", "")
            'If QFE = "SI" Then
            '    GRPO.DocDate = oFactura.FechaEmision
            '    GRPO.TaxDate = oFactura.FechaEmision
            '    GRPO.DocDueDate = oFactura.FechaEmision
            'Else

            '    GRPO.DocDate = Date.Now
            '    GRPO.DocDate = Date.Now
            '    'GRPO.TaxDate = oFactura.FechaEmision
            '    GRPO.DocDueDate = Date.Now
            ''End If
            If Functions.VariablesGlobales._vgFechaEmisionFactura = "Y" Then
                GRPO.DocDate = oFactura.FechaEmision
                GRPO.DocDueDate = oFactura.FechaEmision
                GRPO.TaxDate = oFactura.FechaEmision

            ElseIf Functions.VariablesGlobales._vgFechaEmisionFacturaP = "Y" Then
                GRPO.DocDate = oFactura.FechaEmision

            Else
                GRPO.DocDate = Date.Now
                GRPO.TaxDate = oFactura.FechaEmision
            End If

            'GRPO.DocDueDate = oFactura.FechaEmision



            Dim Prefijo As String = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "FC", "Prefijo")

            ' DATOS DE AUTORIZACION
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = oFactura.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SER_EST").Value = oFactura.Establecimiento
                GRPO.UserFields.Fields.Item("U_SER_PE").Value = oFactura.PuntoEmision
                Try
                    GRPO.UserFields.Fields.Item("U_COD_ST").Value = "01"
                Catch ex As Exception
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_tipo_comprob").Value = "01"
                Catch ex As Exception

                End Try
                GRPO.FolioNumber = oFactura.Secuencial
                GRPO.FolioPrefixString = Prefijo

                ' COMENTAR ESTA LINEA
                'GRPO.NumAtCard = ofactura.Establecimiento + ofactura.PuntoEmision + ofactura.Secuencial

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                GRPO.UserFields.Fields.Item("U_NO_AUTORI").Value = oFactura.AutorizacionSRI
                GRPO.NumAtCard = oFactura.Establecimiento + oFactura.PuntoEmision + oFactura.Secuencial

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                GRPO.NumAtCard = oFactura.Establecimiento + "-" + oFactura.PuntoEmision + "-" + oFactura.Secuencial
                GRPO.UserFields.Fields.Item("U_SYP_SERIESUC").Value = oFactura.Establecimiento
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDSD").Value = oFactura.PuntoEmision
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDSD").Value = oFactura.PuntoEmision
                    Utilitario.Util_Log.Escribir_Log("Error2213frmDocumento: " + ex.Message.ToString(), "recepcionSeidor")
                End Try
                Try
                    GRPO.UserFields.Fields.Item("U_SYP_MDCD").Value = oFactura.Secuencial
                Catch ex As Exception
                    GRPO.UserFields.Fields.Item("U_BPP_MDCD").Value = oFactura.Secuencial
                    Utilitario.Util_Log.Escribir_Log("Error2213frmDocumento: " + ex.Message.ToString(), "recepcionSeidor")
                End Try

                GRPO.UserFields.Fields.Item("U_SYP_NROAUTO").Value = oFactura.AutorizacionSRI

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                GRPO.UserFields.Fields.Item("U_SS_NumAut").Value = oFactura.AutorizacionSRI
                GRPO.UserFields.Fields.Item("U_SS_Est").Value = oFactura.Establecimiento
                GRPO.UserFields.Fields.Item("U_SS_Pemi").Value = oFactura.PuntoEmision

                GRPO.FolioNumber = oFactura.Secuencial
                GRPO.FolioPrefixString = Prefijo

                If oFactura.ENTPagos.Count > 0 Then
                    GRPO.UserFields.Fields.Item("U_SS_FormaPagos").Value = oFactura.ENTPagos(0).FormaPago
                End If


            End If

            GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
            GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryFVRecibida_UDO.ToString()

            'Dim serviceInvoice As Documents = TryCast(B1Connections.diCompany.GetBusinessObject(BoObjectTypes.oInvoices), Documents)
            'serviceInvoice.CardCode = "C20000"
            GRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            GRPO.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO

            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" Where ""CardCode"" =  '" + sCardCode + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardCode =  '" + sCardCode + "'"
            End If
            FormatCodeProveedor = oFuncionesB1.getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            Dim FormatCode As String = ""
            Dim sQueryAcctCode As String = ""
            If FormatCodeProveedor = "" Then
                FormatCode = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "FC", "Cuenta")
            Else
                FormatCode = FormatCodeProveedor
            End If

            If FormatCode = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, "ERROR - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode + "'"
            End If

            Dim Cuenta As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")


            Dim line As Integer = 0
            'Dim CodImpV As SAPbobsCOM.Recordset = rsboApp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim result As String

            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In oFactura.ENTDetalleFactura


                sQueryCodImp = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje.ToString + "' "
                CodImp = oFuncionesB1.getRSvalue(sQueryCodImp, "U_SSID", "")

                sQueryCodImpV = " SELECT TOP 1 ""U_SSCOD"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje.ToString + "' "
                CodImpV = oFuncionesB1.getRSvalue(sQueryCodImpV, "U_SSCOD", "")
                'CodImpV.DoQuery(sQueryCodImpV)
                'result = CodImpV.Fields.Item("U_SSCOD").Value.ToString()

                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImp + "Resultado :" + CodImp.ToString(), "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImpV + "Resultado :" + CodImpV.ToString(), "frmProcesoLoteC")
                GRPO.Lines.AccountCode = Cuenta
                GRPO.Lines.LineTotal = formatDecimal(oDetalle.PrecioTotalSinImpuesto)
                'GRPO.Lines.TaxCode = oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje

                Dim sQueryCentroCosto As String = " SELECT TOP 1 ""U_SSCEN_COS"" FROM ""OCRD"" WHERE ""CardCode"" = '" + sCardCode.ToString + "' "
                Dim CodCentroCosto As String = oFuncionesB1.getRSvalue(sQueryCentroCosto, "U_SSCEN_COS", "")
                Utilitario.Util_Log.Escribir_Log("Obteniendo centro de costo - QUERY: " + sQueryCentroCosto + "Resultado :" + CodCentroCosto.ToString(), "frmProcesoLoteC")

                Dim sQueryMarca As String = " SELECT TOP 1 ""U_SSMARCA"" FROM ""OCRD"" WHERE ""CardCode"" = '" + sCardCode.ToString + "' "
                Dim CodMarca As String = oFuncionesB1.getRSvalue(sQueryMarca, "U_SSMARCA", "")
                Utilitario.Util_Log.Escribir_Log("Obteniendo marca - QUERY: " + sQueryMarca + "Resultado :" + CodMarca.ToString(), "frmProcesoLoteC")

                Try
                    If CodImpV = oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje.ToString Then
                        GRPO.Lines.TaxCode = CodImp.ToString
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("ERROR: " + ex.ToString(), "frmDocumento")
                End Try

                GRPO.Lines.ItemDescription = "SERVICIO"
                GRPO.Lines.Quantity = 1
                Try
                    'GRPO.Lines.ProfitCenter = CodCentroCosto
                    GRPO.Lines.CostingCode = CodCentroCosto
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("ERROR CodCentroCosto: " + ex.ToString(), "frmDocumento")
                End Try
                Try
                    GRPO.Lines.CostingCode2 = CodMarca
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("ERROR CodMarca: " + ex.ToString(), "frmDocumento")
                End Try

                GRPO.Lines.Add()
                line += 1
            Next

            If oFuncionesB1.checkCampoBD("OPCH", "SS_PROLOTE") Then
                GRPO.UserFields.Fields.Item("U_SS_PROLOTE").Value = "SI"
            End If
            'GRPO.Comments += "Creado por el addon SAED"

            RetVal = GRPO.Add()
            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Elimina_DocumentoRecibido_Factura(DocEntryFVRecibida_UDO)

                rCompany.GetLastError(ErrCode, ErrMsg)

                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, "PL Ocurrio Error al grabar Factura Preliminar de tipo: Servicio:" + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, "PL Factura Preliminar de tipo: Servicio, Creada Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_Factura(DocEntryFVRecibida_UDO, sDocEntryPreliminar)
                Return True
            End If

        Catch ex As Exception
            Elimina_DocumentoRecibido_Factura(DocEntryFVRecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - PL Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("REE", oFactura.ClaveAcceso, "PL Ocurrio Error al grabar Factura Preliminar de tipo: Servicio:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            GRPO = Nothing
            GC.Collect()
        End Try

    End Function
    Public Sub Actualiza_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO As String, DocEntryPreliminar As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)

            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

            oGeneralService.Update(oGeneralData)

            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Error: Actualizando Numero de Documento Preliminar en Documento Recibido UDO: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub
    Public Sub Elimina_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Eliminando Documento Recibido UDO Retención # " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)

            'oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            'oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

            oGeneralService.Delete(oGeneralParams)


            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO Retención..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub
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

    Public Function Guarda_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, ByRef _oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, _scardcode As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        Dim _ClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()

        'Dim results As List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion)
        'results = listaREs.FindAll(Function(column) column.ClaveAcceso = _ClaveAcceso)
        'For Each oRE As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion In results
        '    oRetencion = oRE
        'Next

        Try
            rsboApp.StatusBar.SetText(NombreAddon + "- Creando registro de Pago Recibido(Retencion) Recibida UDO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Creando registro de Pago Recibido(Retencion) Recibida UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            'Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
            'Try
            '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '        sCardCode = oFuncionesB1.getRSvalue("SELECT ""CardCode"" FROM OCRD where ""LicTradNum"" = '" + sRUC + "' AND ""CardType"" = 'C' ", "CardCode", "")
            '    Else
            '        sCardCode = oFuncionesB1.getRSvalue("SELECT CardCode FROM OCRD WITH(NOLOCK) where LicTradNum = '" + sRUC + "' AND CardType = 'C' ", "CardCode", "")
            '    End If
            'Catch ex As Exception
            '    Utilitario.Util_Log.Escribir_Log("REE - error al obtener CardCode PL : " + ex.ToString, "frmProcesoLoteC")
            'End Try

            Dim numDoc As String = oRetencion.Establecimiento + "-" + oRetencion.PuntoEmision + "-" + oRetencion.Secuencial

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralData.SetProperty("Code", conta)
            oGeneralData.SetProperty("U_RUC", oRetencion.Ruc.ToString())
            oGeneralData.SetProperty("U_Nombre", Left(oRetencion.RazonSocial.ToString(), 99))
            oGeneralData.SetProperty("U_CardCode", sCardCode.ToString())
            'oGeneralData.SetProperty("U_Mapeado", oForm.Items.Item("lbMapp").Specific.Value.ToString())
            oGeneralData.SetProperty("U_ClaAcc", oRetencion.ClaveAcceso.ToString())
            oGeneralData.SetProperty("U_NumAut", oRetencion.AutorizacionSRI.ToString())
            oGeneralData.SetProperty("U_FecAut", oRetencion.FechaAutorizacion.ToString())
            oGeneralData.SetProperty("U_NumDoc", numDoc.ToString())
            If Functions.VariablesGlobales._ContabilizarPRPL = "Y" Then
                oGeneralData.SetProperty("U_FPrelim", "0") 'sDocEntryPreliminar
            Else
                oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString()) 'sDocEntryPreliminar
            End If
            'oGeneralData.SetProperty("U_FPrelim", DocEntryFacturaRecibida_UDO.ToString())
            oGeneralData.SetProperty("U_vTotal", Convert.ToDouble(formatDecimal(oRetencion.TotalRetencion.ToString())))
            oGeneralData.SetProperty("U_IdGS", oRetencion.IdRetencion.ToString())
            oGeneralData.SetProperty("U_Sincro", 0)
            oGeneralData.SetProperty("U_Estado", "docPreliminar")


            oChildren = oGeneralData.Child("GS0_RER")
            For Each detalleRet As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRetencion.ENTDetalleRetencion

                Dim ejeFiscal As String = oRetencion.PeriodoFiscal
                Dim _numDocRet As String = ""
                oChild = oChildren.Add
                'oChild.SetProperty("U_CodRet", odt.GetValue(0, i).ToString() + " - " + odt.GetValue(6, i).ToString())
                oChild.SetProperty("U_CodRet", "FACTURA " + detalleRet.CodigoRetencion.ToString())
                If IsNothing(detalleRet.NumDocRetener) Then
                    _numDocRet = "0"
                Else
                    _numDocRet = detalleRet.NumDocRetener.ToString
                End If
                oChild.SetProperty("U_NumDocRe", _numDocRet)
                oChild.SetProperty("U_Fecha", detalleRet.FechaEmisionDocRetener.ToString())
                oChild.SetProperty("U_pFiscal", ejeFiscal.ToString())
                oChild.SetProperty("U_Base", Convert.ToDouble(formatDecimal(detalleRet.BaseImponible.ToString())))
                If detalleRet.CodigoRetencion = "1" Then
                    oChild.SetProperty("U_Impuesto", "RENTA")
                Else
                    oChild.SetProperty("U_Impuesto", "IVA")
                End If
                oChild.SetProperty("U_Porcent", Convert.ToDouble(formatDecimal(detalleRet.PorcentajeRetener.ToString())))
                oChild.SetProperty("U_valorR", Convert.ToDouble(formatDecimal(detalleRet.ValorRetenido.ToString())))

            Next

            oChildren = oGeneralData.Child("GS1_RER")
            For Each Info As Entidades.wsEDoc_ConsultaRecepcion.ENTDatoAdicionalRetencion In _oRetencion.ENTDatoAdicionalRetencion
                oChild = oChildren.Add
                oChild.SetProperty("U_Nombre", IIf(IsNothing(Info.Nombre), "", Left(Info.Nombre, 253)))
                oChild.SetProperty("U_Valor", IIf(IsNothing(Info.Descripcion), "", Left(Info.Descripcion, 253)))
            Next

            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Se creo registro de Pago Recibido(Retencion) Recibida UDO satisfactoriamente, # : " + DocEntryFacturaRecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Ocurrior un error al crear registro de Pago Recibido(Retencion) Recibida UDO: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al guardar Pago Recibido(Retencion) Recibida en el UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Mensaje_Error = ex.Message.ToString
            Return False
        End Try
    End Function
    Private Function CrearPagoRecibido(ByRef sCardCode As String, ByRef oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean


        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
            Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

            Return CrearPagoRecibido_E_O(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryRERecibida_UDO)

        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

            Return CrearPagoRecibido_S(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryRERecibida_UDO)
        Else
            men = "No esta definida la localizacion configurada acrtual"
        End If

    End Function
    Private Function CrearPagoRecibido_E_O(ByRef sCardCode As String, ByRef oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        'Dim RetVal As Long
        'Dim ErrCode As Long
        'Dim ErrMsg As String
        'Dim vPay As SAPbobsCOM.Payments
        'CLIENTE BANCARIO 
        Dim sQueryCB As String = ""
        sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardCode""= '" + sCardCode.ToString + "' "
        Dim clienteBancario As String = oFuncionesB1.getRSvalue(sQueryCB, "U_SSCLIENTEBANCO", "")
        If clienteBancario = "SI" Then
            Dim estado As Boolean
            estado = CrearPagoRecibido_E_OCB(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado
        Else
            Dim estado As Boolean
            estado = CrearPagoRecibido_E_ONormal(sCardCode, oRetencion, sDocEntryPreliminar, DocEntryRERecibida_UDO)
            Return estado
        End If
    End Function
    Private Function CrearPagoRecibido_E_OCB(ByRef sCardCode As String, ByRef oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCBMP As Integer = 1
        Try

            'CLIENTE BANCARIO 
            oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Creando Pago Recibido Tipo Cuenta(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido Tipo Cuenta(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rAccount
            'vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            vPay.DocCurrency = "USD"
            'vPay.DocDate = Date.Now
            'vPay.TaxDate = _oDocumento.FechaEmision
            vPay.DocDate = oRetencion.FechaEmision
            vPay.TaxDate = oRetencion.FechaEmision
            vPay.DueDate = oRetencion.FechaEmision
            vPay.DocRate = 0

            Dim RucEmpresa As String = "Select ""TaxIdNum"" from ""OADM""  "
            Dim IdRucEmp = oFuncionesB1.getRSvalue(RucEmpresa, "TaxIdNum", "")

            If IdRucEmp = "1790027791001" Then
                If sCardCode = "MZCL-045850" Then
                    vPay.JournalRemarks = "AUSTRO RET TC"
                ElseIf sCardCode = "MZCL-045955" Then
                    vPay.JournalRemarks = "PACIFICO RET TC"
                ElseIf sCardCode = "MZCL-045059" Then
                    vPay.JournalRemarks = "DINERS RET TC"
                ElseIf sCardCode = "MZCL-045236" Then
                    vPay.JournalRemarks = "AMEX RET TC"
                ElseIf sCardCode = "MZCL-045239" Then
                    vPay.JournalRemarks = "INTER RET TC"
                ElseIf sCardCode = "N1MM-000004" Then
                    vPay.JournalRemarks = "MASTERCARD RET TC"
                ElseIf sCardCode = "MZCL-045851" Then
                    vPay.JournalRemarks = "SOLIDARIO RET TC"

                End If

            ElseIf IdRucEmp = "1792890438001" Then
                If sCardCode = "MZCL-001123" Then
                    vPay.JournalRemarks = "AUSTRO RET TC"
                ElseIf sCardCode = "MZCL-001124" Then
                    vPay.JournalRemarks = "PACIFICO RET TC"
                ElseIf sCardCode = "MZCL-001125" Then
                    vPay.JournalRemarks = "DINERS RET TC"
                ElseIf sCardCode = "MZCL-000963" Then
                    vPay.JournalRemarks = "AMEX RET TC"
                ElseIf sCardCode = "MZCL-000086" Then
                    vPay.JournalRemarks = "INTER RET TC"
                ElseIf sCardCode = "MZCL-001118" Then
                    vPay.JournalRemarks = "MASTERCARD RET TC"
                ElseIf sCardCode = "MZCL-001126" Then
                    vPay.JournalRemarks = "SOLIDARIO RET TC"

                End If

            ElseIf IdRucEmp = "" Or (IdRucEmp <> "1792890438001" And IdRucEmp <> "1790027791001") Then
                vPay.JournalRemarks = ""
            End If

            'AGREGAR DETALLE DEL PAGO
            ' OBTENCION CUENTA CONTABLE
            Dim FormatCodeProveedor As String = ""
            Dim QueryCuentaProveedor As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" Where ""CardCode"" =  '" + sCardCode + "'"
            Else
                QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardCode =  '" + sCardCode + "'"
            End If
            FormatCodeProveedor = oFuncionesB1.getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

            Dim FormatCode As String = ""
            Dim sQueryAcctCode As String = ""
            If FormatCodeProveedor = "" Then
                FormatCode = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "FC", "Cuenta")
            Else
                FormatCode = FormatCodeProveedor
            End If

            If FormatCode = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "ERROR - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode + "'"
            End If
            Dim Cuenta As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")
            ' END OBTENCION CUENTA CONTABLE
            Try
                vPay.AccountPayments.AccountCode = Cuenta
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Agregando cuenta" + ex.Message.ToString + Cuenta.ToString(), "ProcesoLote")
            End Try

            'vPay.AccountPayments.AccountName = NombreCuentaRetencionCB
            Try
                vPay.AccountPayments.SumPaid = oRetencion.TotalRetencion
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("agregando total retencion " + ex.Message.ToString + oRetencion.TotalRetencion.ToString(), "ProcesoLote")
            End Try

            Try
                vPay.AccountPayments.Add()
                vPay.AccountPayments.SetCurrentLine(1)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("error al agregar lineas pago recibido " + ex.Message.ToString, "ProcesoLote")
            End Try

            'END AGREGAR DETALLE DEL PAGO


            '1 RENTA 2 IVA
            ' DETALLES
            Dim sQueryCodRetencion As String = ""
            Dim sQueryCuentaRetencion As String = ""
            Dim sQueryCrTypeCode As String = ""

            Dim CodRetencion As String = ""
            Dim CrTypeCode As String = ""
            Dim CuentaRetencion As String = ""

            Dim secuencial As Integer = 1
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRetencion.ENTDetalleRetencion

                If oDetalle.ValorRetenido > 0 Then

                    vPay.CreditCards.AdditionalPaymentSum = 0
                    vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                    If oDetalle.Codigo = 1 Then ' RENTA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmProcesoLoteC")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
                            Exit Function
                        End If
                    ElseIf oDetalle.Codigo = 2 Then ' IVA

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmProcesoLoteC")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    ElseIf oDetalle.Codigo = 6 Then ' ISD

                        sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                        CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                        Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmProcesoLoteC")
                        If CodRetencion = "" Then
                            rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                            Exit Function
                        End If
                    End If

                    sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmProcesoLoteC")

                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                    CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmProcesoLoteC")

                    'vPay.CreditCards.CreditAcct = IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    'vPay.CreditCards.CreditCard = IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Try
                        vPay.CreditCards.CreditAcct = CuentaRetencion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando cuenta retencion: " + CuentaRetencion.ToString(), "frmProcesoLoteC")
                    End Try
                    'IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                    Try
                        vPay.CreditCards.CreditCard = CodRetencion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando codigo retencion: " + CodRetencion.ToString(), "frmProcesoLoteC")
                    End Try
                    ' IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                    Try
                        vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando codigo tipo codigo: " + CrTypeCode.ToString(), "frmProcesoLoteC")
                    End Try

                    Try
                        vPay.CreditCards.CreditCardNumber = oRetencion.Secuencial
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando numero de tarjeta de credito : " + oRetencion.Secuencial.ToString(), "frmProcesoLoteC")
                    End Try

                    Try
                        vPay.CreditCards.CreditSum = oDetalle.ValorRetenido
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando valor retenido : " + oDetalle.ValorRetenido.ToString(), "frmProcesoLoteC")
                    End Try
                    ' _oDocumento.TotalRetencion ' formatDecimal(_oDocumento.TotalRetencion.ToString())
                    ' vPay.CreditCards.CreditType = 1
                    Try
                        vPay.CreditCards.FirstPaymentSum = oRetencion.TotalRetencion
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando total retencion : " + oRetencion.TotalRetencion.ToString(), "frmProcesoLoteC")
                    End Try


                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        Try
                            vPay.CreditCards.FirstPaymentDue = oRetencion.FechaAutorizacion
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + oRetencion.FechaAutorizacion.ToString(), "frmProcesoLoteC")
                        End Try


                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = oRetencion.Establecimiento + oRetencion.PuntoEmision
                            Else
                                vPay.CreditCards.VoucherNum = "123456789012345"
                                'vPay.CreditCards.OwnerPhone = "7777777"
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + oDetalle.NumDocRetener.ToString(), "frmProcesoLoteC")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando monto base : " + oDetalle.BaseImponible.ToString(), "frmProcesoLoteC")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando monto base : " + oDetalle.BaseImponible.ToString(), "frmProcesoLoteC")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = oRetencion.Secuencial
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando numero de retencion : " + oRetencion.Secuencial.ToString(), "frmProcesoLoteC")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = oRetencion.AutorizacionSRI
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando numero de autorizacion de retencion : " + oRetencion.AutorizacionSRI.ToString(), "frmProcesoLoteC")
                        End Try
                        '
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = oRetencion.Establecimiento + oRetencion.PuntoEmision
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando est y punto de emision : " + oRetencion.Establecimiento.ToString() + oRetencion.PuntoEmision, "frmProcesoLoteC")
                        End Try
                        Try
                            If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                                vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = sCardCode
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando est y punto de emision : " + sCardCode.ToString, "frmProcesoLoteC")
                        End Try

                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        vPay.CreditCards.VoucherNum = oRetencion.Establecimiento + oRetencion.PuntoEmision + oRetencion.Secuencial
                        If oFuncionesB1.checkCampoBD("RCT3", "NUM_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = oRetencion.AutorizacionSRI
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = oRetencion.FechaAutorizacion
                        End If
                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        Try
                            vPay.CreditCards.FirstPaymentDue = _oDocumento.FechaAutorizacion
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + _oDocumento.FechaAutorizacion.ToString(), "frmDocumentoRE")
                        End Try


                        Try
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
                                vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                vPay.CreditCards.OwnerPhone = _oDocumento.Establecimiento + _oDocumento.PuntoEmision
                            Else
                                vPay.CreditCards.VoucherNum = "123456789012345"
                                'vPay.CreditCards.OwnerPhone = "7777777"
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("agregando fecha autroizacion : " + oDetalle.NumDocRetener.ToString(), "frmDocumentoRE")
                        End Try
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBase") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_SecRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = oRetencion.Secuencial
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_AutRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = oRetencion.AutorizacionSRI
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = oRetencion.Establecimiento + oRetencion.PuntoEmision
                        End If
                        If oFuncionesB1.checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = sCardCode
                        End If

                        If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                            vPay.CreditCards.CreditCardNumber = oRetencion.Secuencial
                        Else
                            vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = oRetencion.Secuencial
                        End If
                    End If



                    If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
                        vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                    End If
                    Try
                        If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando docentry udo : " + DocEntryRERecibida_UDO.ToString, "frmProcesoLoteC")
                    End Try

                    Try
                        vPay.CreditCards.Add()
                        vPay.CreditCards.SetCurrentLine(secuencial)
                        secuencial += 1
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("agregando lineas medio de pago : " + ex.Message.ToString, "frmProcesoLoteC")
                    End Try


                End If
            Next
            Try
                RetVal = vPay.Add()
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("erro al agregar pago recibido tipo cuenta : " + ex.Message.ToString, "frmProcesoLoteC")
            End Try

            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode.ToString() + oRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "frmProcesoLoteC")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "frmProcesoLoteC")
                End Try

                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                Try
                    Dim xml As String = vPay.GetAsXML()
                    Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                    Dim sRuta As String = sRutaCarpeta & "\" & sCardCode.ToString() + oRetencion.IdRetencion.ToString() + ".xml"
                    'Dim xml As String = vPay.GetAsXML()
                    If System.IO.Directory.Exists(sRutaCarpeta) Then
                        Utilitario.Util_Log.Escribir_Log("Serializando...", "frmProcesoLoteC")
                        Dim writer As TextWriter = New StreamWriter(sRuta)
                        writer.Write(xml)
                        writer.Close()
                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "frmProcesoLoteC")
                End Try
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If
        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try

    End Function
    Private Function CrearPagoRecibido_E_ONormal(ByRef sCardCode As String, ByRef oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Utilitario.Util_Log.Escribir_Log("Inicio funcion creae preliminar sCardCode: " + sCardCode.ToString, "frmProcesoLoteC")
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments

        Dim sQueryCodRetencionCB As String = ""
        Dim sQueryCuentaRetencionCB As String = ""
        Dim sQueryCrTypeCodeCB As String = ""
        Dim sQueryNombreCuentaRetencionCB As String = ""

        Dim CodRetencionCB As String = ""
        Dim CrTypeCodeCB As String = ""
        Dim CuentaRetencionCB As String = ""
        Dim NombreCuentaRetencionCB As String = ""

        Dim secuencialCB As Integer = 1
        Dim secuencialCBMP As Integer = 1


        Dim valorRetencion As Double = oRetencion.TotalRetencion
        'Dim queryValorRT As String = oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString
        Dim _numDocRetener As String = oRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(0, 3).ToString + "-" + oRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(3, 3).ToString + "-" + CLng(Right(oRetencion.ENTDetalleRetencion(0).NumDocRetener, 9)).ToString
        Dim est As String = oRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(0, 3).ToString
        Dim punEmi As String = oRetencion.ENTDetalleRetencion(0).NumDocRetener.Substring(3, 3).ToString
        Dim folio As String = CLng(Right(oRetencion.ENTDetalleRetencion(0).NumDocRetener, 9)).ToString
        Dim queryValorRT As String = ""
        Dim docentry As String = ""
        Dim queryCardCode As String = ""
        Dim queryDocDate As String = ""
        Dim diaFechaAut As Integer = 0
        Dim diaFechaAct As Integer = 0

        Dim QryCardCode As String = ""

        Dim PORCENTAJES As String = ""
        Dim comentarioPago As String = ""
        Dim NUMFACT As String = ""

        'Dim _UltDiaFecha As String = UltDiaFecha.Day.ToString


        'Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        'Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        'Dim oGrid As SAPbouiCOM.Grid
        'oGrid = oForm.Items.Item("oGrid").Specific
        'Dim oDatable As SAPbouiCOM.DataTable
        'oDatable = oGrid.DataTable
        'Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    QryCardCode = "SELECT ""CardCode"" FROM ""OCRD"" where ""CardType"" = 'C' AND ""LicTradNum"" = '" + sRUC + "'"
        'Else
        '    QryCardCode = "SELECT CardCode FROM OCRD WITH(NOLOCK) where CardType = 'C' AND LicTradNum = '" + sRUC + "'"
        'End If
        'sCardCode = oFuncionesB1.getRSvalue(QryCardCode, "CardCode", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo CARDCODE - QUERY: " + QryCardCode + "Resultado :" + sCardCode.ToString(), "frmProcesoLoteC")

        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryValorRT = " SELECT (""DocTotal""-""PaidToDate"") as SaldoPendiente from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SER_EST""='" + est + "' and ""U_SER_PE""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryValorRT = " SELECT (DocTotal-PaidToDate) as SaldoPendiente from oinv where CardCode='" + sCardCode.ToString + "' and (U_SER_EST +'-'+ U_SER_PE +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryValorRT = " SELECT (""DocTotal""-""PaidToDate"") as SaldoPendiente from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SS_Est""='" + est + "' and ""U_SS_Pemi""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryValorRT = " SELECT (DocTotal-PaidToDate) as SaldoPendiente from oinv where CardCode='" + sCardCode.ToString + "' and (U_SS_Est +'-'+ U_SS_Pemi +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        End If

        Dim valorRT As String = oFuncionesAddon.getRSvalue(queryValorRT, "SaldoPendiente", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo saldo pendiendte - QUERY: " + queryValorRT + "Resultado :" + valorRT.ToString(), "frmProcesoLoteC")

        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                docentry = " SELECT ""DocEntry"" from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SER_EST""='" + est + "' and ""U_SER_PE""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                docentry = " SELECT DocEntry  from oinv where  CardCode='" + sCardCode + "' and (U_SER_EST +'-'+ U_SER_PE +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                docentry = " SELECT ""DocEntry"" from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SS_Est""='" + est + "' and ""U_SS_Pemi""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                docentry = " SELECT DocEntry  from oinv where  CardCode='" + sCardCode + "' and (U_SS_Est +'-'+ U_SS_Pemi +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        End If

        Dim _docentry As String = oFuncionesAddon.getRSvalue(docentry, "DocEntry", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo docentry de la factura - QUERY: " + docentry + "Resultado :" + _docentry.ToString(), "frmProcesoLoteC")

        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryCardCode = " SELECT ""CardCode"" from  OINV where ""U_SER_EST""='" + est + "' and ""U_SER_PE""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryCardCode = " SELECT CardCode from oinv where (U_SER_EST +'-'+ U_SER_PE +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryCardCode = " SELECT ""CardCode"" from  OINV where ""U_SS_Est""='" + est + "' and ""U_SS_Pemi""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryCardCode = " SELECT CardCode from oinv where (U_SS_Est +'-'+ U_SS_Pemi +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        End If

        Dim _CardCode As String = oFuncionesAddon.getRSvalue(queryCardCode, "CardCode", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo cardcode - QUERY: " + queryCardCode + "Resultado :" + _CardCode.ToString(), "frmProcesoLoteC")

        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryDocDate = " SELECT ""DocDate"" from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SER_EST""='" + est + "' and ""U_SER_PE""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryDocDate = " SELECT DocDate from oinv where CardCode='" + sCardCode + "' and (U_SER_EST +'-'+ U_SER_PE +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                queryDocDate = " SELECT ""DocDate"" from  OINV where ""CardCode""='" + sCardCode.ToString + "' and ""U_SS_Est""='" + est + "' and ""U_SS_Pemi""='" + punEmi + "' and cast(""FolioNum"" as varchar)='" + folio + "'"
            Else
                queryDocDate = " SELECT DocDate from oinv where CardCode='" + sCardCode + "' and (U_SS_Est +'-'+ U_SS_Pemi +'-'+ cast(FolioNum as varchar))='" + _numDocRetener + "' "
            End If
        End If

        Dim DocDate As String = oFuncionesAddon.getRSvalue(queryDocDate, "DocDate", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo fecha de la factura: " + queryDocDate + "Resultado :" + DocDate.ToString(), "frmProcesoLoteC")

        Dim MedioPago As String
        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                MedioPago = " SELECT top 1 ""DocNum"" from  RCT3 where ""VoucherNum""='" + Left(oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString(), 15) + "' and ""U_CXS_NUM_AUTO_RETE""='" + oRetencion.AutorizacionSRI + "'"
            Else
                MedioPago = " SELECT top 1 DocNum from  RCT3 where VoucherNum='" + Left(oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString(), 15) + "' and U_CXS_NUM_AUTO_RETE='" + oRetencion.AutorizacionSRI + "'"
            End If
        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                MedioPago = " SELECT top 1 ""DocNum"" from  RCT3 where ""VoucherNum""='" + Left(oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString(), 15) + "' and ""U_SS_AutRetRec""='" + oRetencion.AutorizacionSRI + "'"
            Else
                MedioPago = " SELECT top 1 DocNum from  RCT3 where VoucherNum='" + Left(oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString(), 15) + "' and U_SS_AutRetRec='" + oRetencion.AutorizacionSRI + "'"
            End If
        End If

        Dim _MedioPago As String = oFuncionesAddon.getRSvalue(MedioPago, "DocNum", "")
        Utilitario.Util_Log.Escribir_Log("Obteniendo docnum del medio de pago : " + MedioPago + "Resultado :" + _MedioPago.ToString(), "frmProcesoLoteC")

        Dim FecDocDate As Date = Format(CDate(DocDate), "dd/MM/yyyy")
        Dim UltDiaFecha As Date = DateSerial(Year(FecDocDate), Month(FecDocDate) + 1, 0)

        Dim FecEmiFac As Date = oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener
        Dim FecAutRet As Date = oRetencion.FechaAutorizacion
        Dim FecEmiRet As Date = oRetencion.FechaEmision
        Dim fechaVencRtMP As Date
        valorRT = Convert.ToDouble(valorRT)
        Dim _valorRT As Double = ConvertToDouble(valorRT)
        Utilitario.Util_Log.Escribir_Log("_valorRT Resultado :" + _valorRT.ToString(), "frmProcesoLoteC")
        'Dim jaja As Double = Convert.ToDouble(valorRT)
        'Dim ejej As Double = ConvertToDouble(valorRT)
        Utilitario.Util_Log.Escribir_Log("sCardCode Resultado :" + sCardCode.ToString(), "frmProcesoLoteC")
        _CardCode = sCardCode
        Utilitario.Util_Log.Escribir_Log("_CardCode Resultado :" + _CardCode.ToString(), "frmProcesoLoteC")
        If _docentry <> "0" Then
            Utilitario.Util_Log.Escribir_Log("_docentry Resultado :" + _docentry.ToString(), "frmProcesoLoteC")
            If valorRetencion <= _valorRT Then
                Utilitario.Util_Log.Escribir_Log("El valor de la retencion : " + valorRetencion.ToString + "es menor o igual  a :" + _valorRT.ToString(), "frmProcesoLoteC")
                'rsboApp.StatusBar.SetText(NombreAddon + " - posiposir", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If _MedioPago = "" Then
                    _MedioPago = "0"
                End If
                If Not _MedioPago <> "0" Then


                    Dim dia As String = Date.Now.Day
                    Dim mes As Integer = Date.Now.Month
                    Dim _dia As Integer = Integer.Parse(dia)
                    Dim diaRT As String = oRetencion.FechaEmision.Day
                    Dim mesRT As Integer = oRetencion.FechaEmision.Month
                    Dim _diaRT As Integer = Integer.Parse(diaRT)


                    Try

                        'Dim vPay As SAPbobsCOM.Documents
                        oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        If Functions.VariablesGlobales._ContabilizarPRPL = "Y" Then
                            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                        Else
                            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
                        End If

                        vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                        'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                        vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                        vPay.CardCode = _CardCode
                        Utilitario.Util_Log.Escribir_Log(" cardcode :" + sCardCode.ToString(), "frmProcesoLoteC")
                        vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES

                        If Functions.VariablesGlobales._ValidarFechasCTK = "Y" Then


                            If FecAutRet.Month <= Date.Now.Month Then

                                If FecEmiRet.Month = Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month = Date.Now.Month Then
                                    vPay.DocDate = Date.Now
                                    fechaVencRtMP = Date.Now
                                ElseIf FecEmiRet.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                    vPay.DocDate = UltDiaFecha 'Date.Now
                                    fechaVencRtMP = UltDiaFecha
                                    'se actualiza de 8 a 5 solicitado por Jennifer Citikold 29/09/2021
                                ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Month < Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                    vPay.DocDate = UltDiaFecha
                                    fechaVencRtMP = UltDiaFecha
                                ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Month = Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                    vPay.DocDate = UltDiaFecha
                                    fechaVencRtMP = UltDiaFecha
                                ElseIf Date.Now.Day > diasValidar And FecEmiRet.Month = Date.Now.Month And FecAutRet.Month = Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                    vPay.DocDate = Date.Now
                                    fechaVencRtMP = Date.Now
                                ElseIf Date.Now.Day > diasValidar And FecEmiRet.Month < Date.Now.Month And FecAutRet.Month < Date.Now.Month And FecEmiFac.Month < Date.Now.Month Then
                                    vPay.DocDate = Date.Now
                                    fechaVencRtMP = Date.Now
                                ElseIf FecEmiRet.Year < Date.Now.Year Then
                                    If Date.Now.Day <= diasValidar And FecEmiRet.Year < Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                        vPay.DocDate = UltDiaFecha
                                        fechaVencRtMP = UltDiaFecha
                                    ElseIf FecEmiRet.Month < Date.Now.Year And FecAutRet.Year < Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                        vPay.DocDate = UltDiaFecha 'Date.Now
                                        fechaVencRtMP = UltDiaFecha
                                    ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Year = Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                        vPay.DocDate = UltDiaFecha
                                        fechaVencRtMP = UltDiaFecha
                                    Else
                                        vPay.DocDate = Date.Now
                                        fechaVencRtMP = Date.Now
                                    End If
                                Else
                                    vPay.DocDate = Date.Now
                                    fechaVencRtMP = Date.Now
                                End If
                            ElseIf FecEmiRet.Year < Date.Now.Year Then
                                If Date.Now.Day <= diasValidar And FecEmiRet.Year < Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                    vPay.DocDate = UltDiaFecha
                                    fechaVencRtMP = UltDiaFecha
                                ElseIf FecEmiRet.Month < Date.Now.Year And FecAutRet.Year < Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                    vPay.DocDate = UltDiaFecha 'Date.Now
                                    fechaVencRtMP = UltDiaFecha
                                ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Year = Date.Now.Year And FecAutRet.Year = Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                    vPay.DocDate = UltDiaFecha
                                    fechaVencRtMP = UltDiaFecha
                                ElseIf Date.Now.Day <= diasValidar And FecEmiRet.Year < Date.Now.Year And FecAutRet.Year < Date.Now.Year And FecEmiFac.Year < Date.Now.Year Then
                                    vPay.DocDate = UltDiaFecha
                                    fechaVencRtMP = UltDiaFecha
                                Else
                                    vPay.DocDate = Date.Now
                                    fechaVencRtMP = Date.Now
                                End If
                            Else
                                vPay.DocDate = Date.Now
                                fechaVencRtMP = Date.Now
                            End If

                        ElseIf Functions.VariablesGlobales._PL3FECHAS = "Y" Then 'las 3 fechas son la fecha de emsiion del document
                            vPay.DocDate = oRetencion.FechaEmision
                            vPay.DueDate = oRetencion.FechaEmision
                            vPay.TaxDate = oRetencion.FechaEmision
                        ElseIf Functions.VariablesGlobales._PL1FECHAS = "Y" Then 'la fecha de emision se coloca como fecha de contaibilizacion
                            vPay.DocDate = oRetencion.FechaEmision
                            vPay.TaxDate = oRetencion.FechaEmision

                        ElseIf Functions.VariablesGlobales.FechaFinMesAnteriorPL = "Y" Then
                            Dim fechaActual As Date = Date.Today
                            vPay.DocDate = DateSerial(fechaActual.Year, fechaActual.Month, 1).AddDays(-1)
                            vPay.DueDate = DateSerial(fechaActual.Year, fechaActual.Month, 1).AddDays(-1)
                            vPay.TaxDate = DateSerial(fechaActual.Year, fechaActual.Month, 1).AddDays(-1)
                        Else
                            vPay.DocDate = Date.Now
                            vPay.DueDate = Date.Now
                            vPay.TaxDate = Date.Now
                        End If
                        '**********************************************************************
                        'MANAMER
                        'vPay.DocDate = Date.Now
                        'vPay.TaxDate = Date.Now
                        'vPay.DueDate = Date.Now
                        'MANAMER

                        'vPay.DocDate = Date.Now
                        'vPay.TaxDate = _oDocumento.FechaEmision
                        vPay.DocRate = 0
                        vPay.HandWritten = 0
                        'vPay.Remarks = "SIMON"
                        vPay.JournalRemarks = ""
                        'vPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES
                        vPay.Reference1 = ""

                        '' vPay.Series = 0
                        'vPay.TaxDate = DateSerial(Convert.ToInt32(sFecDep.Substring(0, 4)), Convert.ToInt32(sFecDep.Substring(4, 2)), Convert.ToInt32(sFecDep.Substring(6, 2))) 'Now            
                        ' vPay.TaxDate = Date.Now

                        'vPay.Remarks = "RET" + " " + CLng(oRetencion.Secuencial).ToString + " " + "FAC" + " " + CLng(Right(oRetencion.ENTDetalleRetencion(0).NumDocRetener, 9)).ToString

                        'vPay.Remarks = "RET" + " " + CLng(oRetencion.Secuencial).ToString + " " + "FAC" + " " + CLng(Right(oDetalle.NumDocRetener, 9)).ToString
                        'vPay.JournalRemarks = "RET" + " " + CLng(oRetencion.Secuencial).ToString + " " + "FAC" + " " + CLng(Right(oRetencion.ENTDetalleRetencion(0).NumDocRetener, 9)).ToString


                        '1 RENTA 2 IVA
                        ' DETALLES
                        Dim sQueryCodRetencion As String = ""
                        Dim sQueryCuentaRetencion As String = ""
                        Dim sQueryCrTypeCode As String = ""

                        Dim CodRetencion As String = ""
                        Dim CrTypeCode As String = ""
                        Dim CuentaRetencion As String = ""

                        Dim secuencial As Integer = 1
                        For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRetencion.ENTDetalleRetencion

                            If oDetalle.ValorRetenido > 0 Then

                                If String.IsNullOrEmpty(NUMFACT) Then
                                    If Not NUMFACT.Contains(oDetalle.NumDocRetener.Substring(0, 3).ToString + "-" + oDetalle.NumDocRetener.Substring(3, 3).ToString + "-" + Right(oDetalle.NumDocRetener, 9).ToString) Then
                                        NUMFACT = oDetalle.NumDocRetener.Substring(0, 3).ToString + "-" + oDetalle.NumDocRetener.Substring(3, 3).ToString + "-" + Right(oDetalle.NumDocRetener, 9).ToString
                                    End If

                                Else
                                    If Not NUMFACT.Contains(oDetalle.NumDocRetener.Substring(0, 3).ToString + "-" + oDetalle.NumDocRetener.Substring(3, 3).ToString + "-" + Right(oDetalle.NumDocRetener, 9).ToString) Then
                                        NUMFACT = NUMFACT + " - " + oDetalle.NumDocRetener.Substring(0, 3).ToString + "-" + oDetalle.NumDocRetener.Substring(3, 3).ToString + "-" + Right(oDetalle.NumDocRetener, 9).ToString
                                    End If
                                End If

                                vPay.CreditCards.AdditionalPaymentSum = 0
                                vPay.CreditCards.CardValidUntil = Now 'CDate("10/31/2004")

                                If oDetalle.Codigo = 1 Then ' RENTA

                                    sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_RENTA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                                    CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                                    Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmProcesoLoteC")

                                    If String.IsNullOrEmpty(PORCENTAJES) Then
                                        PORCENTAJES = oDetalle.PorcentajeRetener.ToString + "%"
                                    Else
                                        PORCENTAJES = PORCENTAJES + " - " + oDetalle.PorcentajeRetener.ToString + "%"
                                    End If

                                    If CodRetencion = "" Then
                                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString())
                                        Mensaje_Error = "No esta relacionado el codigo de Renta: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString()
                                        Exit Function
                                    End If
                                ElseIf oDetalle.Codigo = 2 Then ' IVA

                                    sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_IVA"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                                    CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                                    Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO RENTA - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmProcesoLoteC")

                                    If String.IsNullOrEmpty(PORCENTAJES) Then
                                        PORCENTAJES = oDetalle.PorcentajeRetener.ToString + "%"
                                    Else
                                        PORCENTAJES = PORCENTAJES + " - " + oDetalle.PorcentajeRetener.ToString + "%"
                                    End If

                                    If CodRetencion = "" Then
                                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                                        Mensaje_Error = "No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString()
                                        Exit Function
                                    End If
                                ElseIf oDetalle.Codigo = 6 Then ' ISD

                                    sQueryCodRetencion = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_ISD"" WHERE ""U_SSCOD"" = '" + oDetalle.CodigoRetencion.ToString + "' "
                                    CodRetencion = oFuncionesB1.getRSvalue(sQueryCodRetencion, "U_SSID", "")
                                    Utilitario.Util_Log.Escribir_Log("Obteniendo CODIGO ISD - QUERY: " + sQueryCodRetencion + "Resultado :" + CodRetencion.ToString(), "frmProcesoLoteC")
                                    If CodRetencion = "" Then
                                        Mensaje_Error = "No esta relacionado el codigo de IVA: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString()
                                        rsboApp.StatusBar.SetSystemMessage(NombreAddon + " - No esta relacionado el codigo de ISD: " + oDetalle.CodigoRetencion.ToString() + " - " + oDetalle.PorcentajeRetener.ToString() + "%")
                                        Exit Function
                                    End If
                                End If

                                sQueryCuentaRetencion = "select ""AcctCode"" from ""OCRC"" where ""CreditCard"" = '" & CodRetencion & "'"
                                CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")
                                Utilitario.Util_Log.Escribir_Log("Obteniendo CUENTA RENTA - QUERY: " + sQueryCuentaRetencion + "Resultado :" + CuentaRetencion.ToString(), "frmProcesoLoteC")

                                sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & CodRetencion & "'"
                                CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                                Dim TypeCode As Integer = CrTypeCode
                                Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmProcesoLoteC")
                                If CrTypeCode = 0 Then
                                    sQueryCrTypeCode = "select ""CrTypeCode"" from ""OCRP"" where ""CreditCard"" = '" & TypeCode & "'"
                                    CrTypeCode = oFuncionesB1.getRSvalue(sQueryCrTypeCode, "CrTypeCode", "")
                                    Utilitario.Util_Log.Escribir_Log("Obteniendo CrTypeCode (0) RENTA - QUERY: " + sQueryCrTypeCode + "Resultado :" + CrTypeCode.ToString(), "frmProcesoLoteC")
                                End If
                                'vPay.CreditCards.CreditAcct = IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                                'vPay.CreditCards.CreditCard = IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                                'vPay.CreditCards.PaymentMethodCode = IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)

                                vPay.CreditCards.CreditAcct = CuentaRetencion 'IIf(oDetalle.Codigo = 1, IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA), CuentaRetencion)
                                vPay.CreditCards.CreditCard = CodRetencion ' IIf(oDetalle.Codigo = 1, IIf(CodRetencionRENTA = "", CodRetencion, CodRetencionRENTA), CodRetencion)
                                Try
                                    vPay.CreditCards.PaymentMethodCode = CrTypeCode 'IIf(oDetalle.Codigo = 1, IIf(CrTypeCodeRENTA = "", CrTypeCode, CrTypeCodeRENTA), CrTypeCode)
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("CrTypeCode Asignado RENTA - QUERY: " + CrTypeCode.ToString(), "frmProcesoLoteC")
                                End Try



                                vPay.CreditCards.CreditCardNumber = oRetencion.Secuencial
                                vPay.CreditCards.CreditSum = oDetalle.ValorRetenido ' _oDocumento.TotalRetencion ' formatDecimal(_oDocumento.TotalRetencion.ToString())
                                ' vPay.CreditCards.CreditType = 1
                                vPay.CreditCards.FirstPaymentSum = oRetencion.TotalRetencion
                                'vPay.CreditCards.NumOfCreditPayments = 1
                                'vPay.CreditCards.NumOfPayments = 1

                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                    vPay.CreditCards.FirstPaymentDue = oRetencion.FechaAutorizacion
                                    'vPay.CreditCards.CardValidUntil = fechaVencRtMP

                                    Try
                                        If Not IsNothing(oDetalle.NumDocRetener) Then
                                            'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9)).ToString()
                                            'Left(odt.GetValue(0, i).ToString(), 99))

                                            'vPay.CreditCards.VoucherNum = Integer.Parse(oDetalle.NumDocRetener.Length)
                                            vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                            vPay.CreditCards.OwnerPhone = oRetencion.Establecimiento + oRetencion.PuntoEmision
                                        End If

                                    Catch ex As Exception
                                    End Try


                                    If oFuncionesB1.checkCampoBD("RCT3", "MONTO_BASE") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_MONTO_BASE") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_MONTO_BASE").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_RETE") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_RETE").Value = oRetencion.Secuencial
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_NUM_AUTO_RETE") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_NUM_AUTO_RETE").Value = oRetencion.AutorizacionSRI
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "CXS_SER_PTO_RET") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_CXS_SER_PTO_RET").Value = oRetencion.Establecimiento + oRetencion.PuntoEmision
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "Exx_SN_Tip_Finan") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_Exx_SN_Tip_Finan").Value = _CardCode
                                    End If
                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                    vPay.CreditCards.VoucherNum = oRetencion.Establecimiento + oRetencion.PuntoEmision + oRetencion.Secuencial
                                    If oFuncionesB1.checkCampoBD("RCT3", "NUM_AUT") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_NUM_AUT").Value = oRetencion.AutorizacionSRI
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "FEC_AUT") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_FEC_AUT").Value = oRetencion.FechaAutorizacion
                                    End If
                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                    vPay.CreditCards.FirstPaymentDue = oRetencion.FechaAutorizacion

                                    Try
                                        If Not IsNothing(oDetalle.NumDocRetener) Then
                                            vPay.CreditCards.VoucherNum = Left(oDetalle.NumDocRetener.ToString(), 15)
                                            vPay.CreditCards.OwnerPhone = oRetencion.Establecimiento + oRetencion.PuntoEmision
                                        End If

                                    Catch ex As Exception
                                    End Try


                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBaseImp") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBaseImp").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString())) 'ConvertToDouble(formatDecimal(_oDocumento.BaseImponible.ToString()))
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_MontoBase") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_MontoBase").Value = ConvertToDouble(formatDecimal(oDetalle.BaseImponible.ToString()))
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_SecRetRec") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_SecRetRec").Value = oRetencion.Secuencial
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_AutRetRec") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_AutRetRec").Value = oRetencion.AutorizacionSRI
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_EstPtoRetRec") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_EstPtoRetRec").Value = oRetencion.Establecimiento + oRetencion.PuntoEmision
                                    End If
                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_TipoFinanSN") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_TipoFinanSN").Value = _CardCode
                                    End If

                                    If oFuncionesB1.checkCampoBD("RCT3", "SS_NombreSN") Then
                                        vPay.CreditCards.UserFields.Fields.Item("U_SS_NombreSN").Value = Left(oRetencion.RazonSocial.ToString, 100)
                                    End If

                                    If String.IsNullOrEmpty(Functions.VariablesGlobales._CampoNumRetencion) Then
                                        vPay.CreditCards.CreditCardNumber = oRetencion.Secuencial
                                    Else
                                        vPay.CreditCards.UserFields.Fields.Item(Functions.VariablesGlobales._CampoNumRetencion).Value = oRetencion.Secuencial
                                    End If
                                End If


                                If oFuncionesB1.checkCampoBD("RCT3", "SSCREADAR") Then
                                    vPay.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
                                End If
                                If oFuncionesB1.checkCampoBD("RCT3", "SSIDDOCUMENTO") Then
                                    vPay.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
                                End If

                                vPay.CreditCards.Add()
                                vPay.CreditCards.SetCurrentLine(secuencial)
                                secuencial += 1



                            End If


                        Next

                        If oFuncionesB1.checkCampoBD("ORCT", "SS_PROLOTE") Then
                            vPay.CreditCards.UserFields.Fields.Item("U_SS_PROLOTE").Value = "SI"
                        End If


                        oListaFacturaVenta.Clear()
                        Dim Factura As String = ""
                        'sCardCode = _CardCode
                        Utilitario.Util_Log.Escribir_Log("CardCode:" + sCardCode.ToString(), "frmProcesoLoteC")
                        For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRetencion.ENTDetalleRetencion
                            If Not IsNothing(oDetalle.NumDocRetener) Then
                                Dim sQueryFactura As String = ""
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sQueryFactura = " SELECT ""DocEntry"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                                        sQueryFactura += " AND ""U_SER_EST"" = '" & oDetalle.NumDocRetener.ToString().Substring(0, 3) & "'"
                                        sQueryFactura += " AND ""U_SER_PE"" = '" & oDetalle.NumDocRetener.ToString().Substring(3, 3) & "'"
                                        sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9))
                                        sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "
                                    Else
                                        sQueryFactura = " SELECT DocEntry FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                                        sQueryFactura += " AND U_SER_EST = '" & oDetalle.NumDocRetener.ToString().Substring(0, 3) & "'"
                                        sQueryFactura += " AND U_SER_PE = '" & oDetalle.NumDocRetener.ToString().Substring(3, 3) & "'"
                                        sQueryFactura += " AND FolioNum = " & Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9))
                                        sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "
                                    End If
                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sQueryFactura = " SELECT ""DocEntry"" FROM ""OINV"" WHERE ""CardCode"" = '" + sCardCode + "'"
                                        sQueryFactura += " AND ""U_SS_Est"" = '" & oDetalle.NumDocRetener.ToString().Substring(0, 3) & "'"
                                        sQueryFactura += " AND ""U_SS_Pemi"" = '" & oDetalle.NumDocRetener.ToString().Substring(3, 3) & "'"
                                        sQueryFactura += " AND ""FolioNum"" = " & Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9))
                                        sQueryFactura += " AND ""DocStatus""='O' AND ""CANCELED""='N' "
                                    Else
                                        sQueryFactura = " SELECT DocEntry FROM OINV WITH(NOLOCK) WHERE CardCode = '" + sCardCode + "'"
                                        sQueryFactura += " AND U_SS_Est = '" & oDetalle.NumDocRetener.ToString().Substring(0, 3) & "'"
                                        sQueryFactura += " AND U_SS_Pemi = '" & oDetalle.NumDocRetener.ToString().Substring(3, 3) & "'"
                                        sQueryFactura += " AND FolioNum = " & Integer.Parse(oDetalle.NumDocRetener.Substring(6, 9))
                                        sQueryFactura += " AND DocStatus='O' AND CANCELED='N' "
                                    End If
                                End If


                                Try
                                    Factura = oFuncionesB1.getRSvalue(sQueryFactura, "DocEntry", "")
                                    Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + sQueryFactura.ToString() + "-Resultado: " + Factura, "frmProcesoLoteC")
                                    If Factura = "" Then
                                        Factura = "0"
                                    End If
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + ex.Message().ToString() + "-QUERY: " + sQueryFactura, "frmProcesoLoteC")
                                    men = ex.Message.ToString()
                                End Try
                                ' GUARDO LA INFO DE LA FACTURA Y EL VALOR A RETENER, PARA AL CREAR EL PAGO DESCONTARLE EL VALOR DE LA RETENCIÓN A CADA FACTURA
                                Utilitario.Util_Log.Escribir_Log(String.Format("Query:{0}..DocEntry-Respuesta:{1}", sQueryFactura, Factura), "frmProcesoLoteC")
                                If Not Factura = "0" Then
                                    oFacturaVenta = New Entidades.FacturaVenta
                                    oFacturaVenta.DocEntry = Factura
                                    oFacturaVenta.ValorARetener = Convert.ToDouble(oDetalle.ValorRetenido)
                                    Utilitario.Util_Log.Escribir_Log("Query Factura Relacionada:" + oDetalle.ValorRetenido.ToString() + "-QUERY: " + sQueryFactura, "frmProcesoLoteC")
                                    Dim query As System.Collections.Generic.IEnumerable(Of Entidades.FacturaVenta)
                                    query = oListaFacturaVenta.Where(Function(q As Entidades.FacturaVenta) q.DocEntry = Factura)
                                    If query.Count() > 0 Then
                                        query.Single().ValorARetener += Convert.ToDouble(oDetalle.ValorRetenido)
                                    Else
                                        oListaFacturaVenta.Add(oFacturaVenta)
                                    End If

                                    'END GUARDO LA INFO DE LA FACTURA Y EL VALOR A RETENER, PARA AL CREAR EL PAGO DESCONTARLE EL VALOR DE LA RETENCIÓN A CADA FACTURA
                                End If
                                'End If 'CargaFacturaRelacionadas 

                            End If
                            'Exit For
                            i += 1
                        Next
                        Try
                            Utilitario.Util_Log.Escribir_Log("Try FC Relacionada: " + Factura.ToString, "frmProcesoLoteC")
                            If Not Factura = "0" Then
                                For Each o As Entidades.FacturaVenta In oListaFacturaVenta
                                    Utilitario.Util_Log.Escribir_Log("Datos Docentry PL: " & o.DocEntry.ToString & " valor a retener: " & o.ValorARetener.ToString, "datosfacturasrelacionadas")
                                    vPay.Invoices.DocEntry = o.DocEntry
                                    vPay.Invoices.SumApplied = o.ValorARetener
                                    vPay.Invoices.Add()
                                Next
                            End If
                            Utilitario.Util_Log.Escribir_Log("Try Salida: " + Factura.ToString, "frmProcesoLoteC")
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Error Try fc relacionada: " + ex.ToString, "frmProcesoLoteC")
                            men = ex.Message.ToString()
                        End Try

                        If Not String.IsNullOrEmpty(Functions.VariablesGlobales._ComentarioPago) Then

                            comentarioPago = Functions.VariablesGlobales._ComentarioPago
                            comentarioPago = comentarioPago.Replace("@NUMRET", oRetencion.Establecimiento + "-" + oRetencion.PuntoEmision + "-" + oRetencion.Secuencial)
                            comentarioPago = comentarioPago.Replace("@FECHAEMISIONRET", CDate(oRetencion.FechaEmision).ToString("yyyy/MM/dd"))
                            comentarioPago = comentarioPago.Replace("@FECHAEMISIONFACT", CDate(oRetencion.ENTDetalleRetencion(0).FechaEmisionDocRetener).ToString("yyyy/MM/dd"))
                            comentarioPago = comentarioPago.Replace("@PORCRET", PORCENTAJES)
                            comentarioPago = comentarioPago.Replace("@NUMFACT", NUMFACT)


                            vPay.JournalRemarks = comentarioPago.ToString.Substring(0.252)
                            vPay.Remarks = comentarioPago.ToString.Substring(0.252)

                        End If

                        RetVal = vPay.Add()
                        'Dim xml As String = vPay.GetAsXML()
                        'Utilitario.Util_Log.Escribir_Log("Serializando...", "frmProcesoLoteC")



                        If RetVal <> 0 Then
                            'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                            Try
                                Dim xml As String = vPay.GetAsXML()
                                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                                Dim sRuta As String = sRutaCarpeta & "\" & _CardCode.ToString() + oRetencion.IdRetencion.ToString() + ".xml"
                                'Dim xml As String = vPay.GetAsXML()
                                If System.IO.Directory.Exists(sRutaCarpeta) Then
                                    Utilitario.Util_Log.Escribir_Log("Serializando...", "frmProcesoLoteC")
                                    Dim writer As TextWriter = New StreamWriter(sRuta)
                                    writer.Write(xml)
                                    writer.Close()
                                End If
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "frmProcesoLoteC")
                            End Try

                            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)

                            rCompany.GetLastError(ErrCode, ErrMsg)

                            'rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                            oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                            men = "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString()

                            Mensaje_Error = ErrCode.ToString + " - " + ErrMsg.ToString()

                            Return False
                            Exit Function
                        Else
                            Try
                                Dim xml As String = vPay.GetAsXML()
                                Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG"
                                Dim sRuta As String = sRutaCarpeta & "\" & _CardCode.ToString() + oRetencion.IdRetencion.ToString() + ".xml"
                                'Dim xml As String = vPay.GetAsXML()
                                If System.IO.Directory.Exists(sRutaCarpeta) Then
                                    Utilitario.Util_Log.Escribir_Log("Serializando...", "frmProcesoLoteC")
                                    Dim writer As TextWriter = New StreamWriter(sRuta)
                                    writer.Write(xml)
                                    writer.Close()
                                End If
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("EEROR " + ex.Message, "frmProcesoLoteC")
                                men = ex.Message.ToString()
                                Mensaje_Error = "Error al crear serializar " + ex.Message.ToString
                            End Try
                            rCompany.GetNewObjectCode(sDocEntryPreliminar)
                            oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                            If Functions.VariablesGlobales._ContabilizarPRPL = "N" Then
                                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                            End If
                            Return True
                        End If
                    Catch ex As Exception
                        Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
                        rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                        men = ex.Message.ToString()
                        Mensaje_Error = ex.Message.ToString()
                        Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
                        Return False
                    Finally
                        vPay = Nothing
                        GC.Collect()
                    End Try
                Else
                    rsboApp.StatusBar.SetText(NombreAddon + " - Ya existe un Medio de Pago, numero: " + _MedioPago.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    men = "Ya existe un Medio de Pago registrado para la factura" + oRetencion.ENTDetalleRetencion(0).NumDocRetener.ToString() + " , con numero de documento: " + _MedioPago.ToString
                    Mensaje_Error = men
                    Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
                    Return False
                End If

            Else
                rsboApp.StatusBar.SetText(NombreAddon + " - El valor de la retencion recibida: " + valorRetencion.ToString() + " es mayor al valor pendiente: " + _valorRT.ToString + " de la factura " + _numDocRetener.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                men = "El valor de la retencion recibida: " + valorRetencion.ToString() + " es mayor al valor pendiente: " + _valorRT.ToString + " de la factura " + _numDocRetener.ToString + " del SN " + sCardCode.ToString
                Mensaje_Error = men
                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
                Return False

            End If
        Else
            rsboApp.StatusBar.SetText(NombreAddon + " - No se encontró la Factura número: " + _numDocRetener.ToString + " para el SN: " + sCardCode, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            men = "No se encontró la Factura número: " + _numDocRetener.ToString + " para el SN: " + sCardCode
            Mensaje_Error = men
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            Return False
        End If



    End Function
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
    Private Function CrearPagoRecibido_S(ByRef sCardCode As String, ByRef oRetencion As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion, ByRef sDocEntryPreliminar As String, DocEntryRERecibida_UDO As String) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim vPay As SAPbobsCOM.Payments
        Try

            oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Creando Pago Recibido(Retencion) Preliminar", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            rsboApp.StatusBar.SetText(NombreAddon + " - Creando Pago Recibido(Retencion) Preliminar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim _ClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()

            vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            'vPay = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            vPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            vPay.CardCode = sCardCode

            'vPay.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES

            vPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
            'vPay.DocCurrency = "USD"

            vPay.DocDate = oRetencion.FechaEmision
            vPay.TaxDate = oRetencion.FechaEmision
            vPay.DueDate = oRetencion.FechaEmision

            vPay.DocRate = 0
            vPay.HandWritten = 0
            vPay.JournalRemarks = ""
            vPay.Reference1 = ""

            Try
                vPay.UserFields.Fields.Item("U_SYP_PTSC").Value = oRetencion.PuntoEmision
            Catch ex As Exception
                vPay.UserFields.Fields.Item("U_BPP_PTSC").Value = oRetencion.PuntoEmision
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: 1037 frmProcesoLoteC " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            End Try
            Try
                vPay.UserFields.Fields.Item("U_SYP_SUCRET").Value = oRetencion.Establecimiento
            Catch ex As Exception
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: 1042 frmProcesoLoteC " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            End Try
            Try
                vPay.UserFields.Fields.Item("U_SYP_PTCC").Value = oRetencion.Secuencial
            Catch ex As Exception
                vPay.UserFields.Fields.Item("U_BPP_PTCC").Value = oRetencion.Secuencial
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: 1048 frmProcesoLoteC " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            End Try


            If oFuncionesB1.checkCampoBD("ORCT", "SYP_FECHARET") Then
                Try
                    vPay.UserFields.Fields.Item("U_SYP_FECHARET").Value = oRetencion.FechaEmision
                Catch ex As Exception
                End Try
            End If
            If oFuncionesB1.checkCampoBD("ORCT", "SYP_TipoOperacion") Then
                Try
                    vPay.UserFields.Fields.Item("U_SYP_TipoOperacion").Value = "A-015"
                Catch ex As Exception
                End Try
            End If
            If oFuncionesB1.checkCampoBD("ORCT", "SYP_DETIPO") Then
                Try
                    vPay.UserFields.Fields.Item("U_SYP_DETIPO").Value = "RETENCION DE CLIENTES"
                Catch ex As Exception
                End Try
            End If

            If oFuncionesB1.checkCampoBD("ORCT", "FX_AUTO_RETENCION") Then
                vPay.UserFields.Fields.Item("U_FX_AUTO_RETENCION").Value = oRetencion.AutorizacionSRI
            End If

            If oFuncionesB1.checkCampoBD("ORCT", "SSCREADAR") Then
                vPay.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
            End If
            If oFuncionesB1.checkCampoBD("ORCT", "SSIDDOCUMENTO") Then
                vPay.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryRERecibida_UDO.ToString()
            End If

            Dim CodRetencion As String = ""
            CodRetencion = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "CodigoRetencion")
            If CodRetencion = "" Then
                oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "ERROR - Revisar la configuracion en la opcion de parametrizaciones, debe tener registrado una Cuenta Contable Retención!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
                rsboApp.StatusBar.SetText(NombreAddon + " - Revisar la configuracion en la opcion de parametrizaciones, debe tener registrado una Cuenta Contable Retención", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            'oFuncionesB1.getRSvalue
            Dim sQueryCuentaRetencion As String = ""
            Dim CuentaRetencion As String = ""
            If rCompany.DbServerType = 9 Then
                sQueryCuentaRetencion = "select ""AcctCode"" from ""OACT"" where ""FormatCode"" = '" & CodRetencion & "'"
            Else
                sQueryCuentaRetencion = "select AcctCode from OACT WITH(NOLOCK) where FormatCode  = '" & CodRetencion & "'"
            End If
            CuentaRetencion = oFuncionesB1.getRSvalue(sQueryCuentaRetencion, "AcctCode", "")


            Dim CodRetencionRENTA As String = ""
            CodRetencionRENTA = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "RE", "CodigoRetencionR")
            Dim sQueryCuentaRetencionRENTA As String = ""
            Dim CuentaRetencionRENTA As String = ""
            If rCompany.DbServerType = 9 Then
                sQueryCuentaRetencionRENTA = "select ""AcctCode"" from ""OACT"" where ""FormatCode"" = '" & CodRetencionRENTA & "'"
            Else
                sQueryCuentaRetencionRENTA = "select AcctCode from OACT WITH(NOLOCK) where FormatCode  = '" & CodRetencionRENTA & "'"
            End If
            CuentaRetencionRENTA = oFuncionesB1.getRSvalue(sQueryCuentaRetencionRENTA, "AcctCode", "")

            '1 RENTA 2 IVA
            ' DETALLES
            Dim ValorRenta As Decimal = 0
            Dim ValorIva As Decimal = 0
            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion In oRetencion.ENTDetalleRetencion
                If oDetalle.Codigo = 1 Then
                    ValorRenta += oDetalle.ValorRetenido
                ElseIf oDetalle.Codigo = 2 Then
                    ValorIva += oDetalle.ValorRetenido
                End If
            Next

            ' EN LA PESTAÑA TRANSFERENCIA VA EL VALOR RETENIDO DE IVA
            If ValorIva > 0 Then
                vPay.TransferAccount = CuentaRetencion
                vPay.TransferDate = Date.Now
                vPay.TransferSum = ValorIva
                vPay.TransferReference = "RTE " + oRetencion.Secuencial
            End If

            ' EN LA PESTAÑA EFECTIVO VA EL VALOR DE RETENIDO DE LA FUENTE
            If ValorRenta > 0 Then
                vPay.CashAccount = IIf(CuentaRetencionRENTA = "", CuentaRetencion, CuentaRetencionRENTA)
                vPay.CashSum = ValorRenta
            End If

            RetVal = vPay.Add()
            If RetVal <> 0 Then
                'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
                Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Pago Recibido(Retencion) Preliminar: " + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntryPreliminar)
                oFuncionesAddon.GuardaLOG("PRR", _oDocumento.ClaveAcceso, "Pago Recibido(Retencion) Preliminar, Creado Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
                Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO, sDocEntryPreliminar)
                Return True
            End If

        Catch ex As Exception
            Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO)
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", oRetencion.ClaveAcceso, "Error:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        Finally
            vPay = Nothing
            GC.Collect()
        End Try
    End Function
    Public Sub Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String)

        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        Dim _ClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()

        Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try

            rsboApp.StatusBar.SetText(NombreAddon + " - Eliminando Documento Recibido UDO Retención # " + DocEntryRERecibida_UDO.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Eliminando Documento Recibido UDO Retención # " + DocEntryRERecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryRERecibida_UDO)

            'oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            'oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

            oGeneralService.Delete(oGeneralParams)


            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO Retención..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub
    Public Sub Actualiza_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String, DocEntryPreliminar As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")
        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
        Dim _ClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()

        Try
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Actualizando Numero de Documento Preliminar en Documento Recibido UDO", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)

            Dim query As String
            Dim CodeExist As String = "0"
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@GS_RER"" Where ""DocEntry"" = '" + DocEntryRERecibida_UDO + "' "
            Else
                query = "Select DocEntry From [@GS_RER] Where DocEntry = '" + DocEntryRERecibida_UDO + "' "
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            If CodeExist = "0" Or CodeExist = "" Then ' 
                oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "No existen registros coincidentes, la consulta no trajo registro: " + query.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
            Else
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("GS_RER")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_FPrelim", DocEntryPreliminar)

                oGeneralService.Update(oGeneralData)

            End If

            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Error: Actualizando Numero de Documento Preliminar en Documento Recibido UDO: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
        End Try
    End Sub

    'Private Function CrearFacturaPremilinarServicio(ByVal sCardCode As String, ByVal ofactura As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura, ByRef sDocEntryPreliminar As String, ByVal DocEntryFVRecibida_UDO As String) As Boolean


    '    oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "Creando Factura Preliminar de tipo: Servicio", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '    rsboApp.StatusBar.SetText(NombreAddon + " - Creando Factura por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    'Create the Documents object
    '    Dim GRPO As SAPbobsCOM.Documents
    '    Dim RetVal As Long
    '    Dim ErrCode As Long
    '    Dim ErrMsg As String
    '    Dim CodImp As String = ""
    '    Dim sQueryCodImp As String = ""
    '    'Dim CodImpV As String = ""
    '    Dim sQueryCodImpV As String = ""

    '    Try

    '        GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
    '        GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
    '        GRPO.CardCode = sCardCode
    '        ' GRPO.DocDate = ofactura.FechaEmision
    '        GRPO.DocDate = Today.Date
    '        ' GRPO.DocDueDate = Today.Date
    '        GRPO.TaxDate = ofactura.FechaEmision

    '        Dim Prefijo As String = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "FC", "Prefijo")

    '        ' DATOS DE AUTORIZACION
    '        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
    '            GRPO.UserFields.Fields.Item("U_NUM_AUTOR").Value = ofactura.AutorizacionSRI
    '            GRPO.UserFields.Fields.Item("U_SER_EST").Value = ofactura.Establecimiento
    '            GRPO.UserFields.Fields.Item("U_SER_PE").Value = ofactura.PuntoEmision

    '            GRPO.FolioNumber = ofactura.Secuencial
    '            GRPO.FolioPrefixString = Prefijo

    '            ' COMENTAR ESTA LINEA
    '            'GRPO.NumAtCard = ofactura.Establecimiento + ofactura.PuntoEmision + ofactura.Secuencial

    '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
    '            GRPO.UserFields.Fields.Item("U_NO_AUTORI").Value = ofactura.AutorizacionSRI
    '            GRPO.NumAtCard = ofactura.Establecimiento + ofactura.PuntoEmision + ofactura.Secuencial

    '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
    '            GRPO.NumAtCard = _oDocumento.Establecimiento + "-" + _oDocumento.PuntoEmision + "-" + _oDocumento.Secuencial
    '            GRPO.UserFields.Fields.Item("U_SYP_SERIESUC").Value = _oDocumento.Establecimiento
    '            Try
    '                GRPO.UserFields.Fields.Item("U_SYP_MDSD").Value = _oDocumento.PuntoEmision
    '            Catch ex As Exception
    '                GRPO.UserFields.Fields.Item("U_BPP_MDSD").Value = _oDocumento.PuntoEmision
    '                Utilitario.Util_Log.Escribir_Log("Error2213frmDocumento: " + ex.Message.ToString(), "recepcionSeidor")
    '            End Try
    '            Try
    '                GRPO.UserFields.Fields.Item("U_SYP_MDCD").Value = _oDocumento.Secuencial
    '            Catch ex As Exception
    '                GRPO.UserFields.Fields.Item("U_BPP_MDCD").Value = _oDocumento.Secuencial
    '                Utilitario.Util_Log.Escribir_Log("Error2213frmDocumento: " + ex.Message.ToString(), "recepcionSeidor")
    '            End Try

    '            GRPO.UserFields.Fields.Item("U_SYP_NROAUTO").Value = _oDocumento.AutorizacionSRI

    '        End If

    '        GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"
    '        GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value = DocEntryFVRecibida_UDO.ToString()

    '        'Dim serviceInvoice As Documents = TryCast(B1Connections.diCompany.GetBusinessObject(BoObjectTypes.oInvoices), Documents)
    '        'serviceInvoice.CardCode = "C20000"
    '        GRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
    '        GRPO.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO

    '        Dim FormatCodeProveedor As String = ""
    '        Dim QueryCuentaProveedor As String = ""
    '        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '            QueryCuentaProveedor = "Select ""U_SSCUENTA"" from ""OCRD"" Where ""CardCode"" =  '" + sCardCode + "'"
    '        Else
    '            QueryCuentaProveedor = "Select U_SSCUENTA from OCRD Where CardCode =  '" + sCardCode + "'"
    '        End If
    '        FormatCodeProveedor = oFuncionesB1.getRSvalue(QueryCuentaProveedor, "U_SSCUENTA", "")

    '        Dim FormatCode As String = ""
    '        Dim sQueryAcctCode As String = ""
    '        If FormatCodeProveedor = "" Then
    '            FormatCode = ofrmParametrosRecepcion.ConsultaParametro("RECEPCION", "PARAMETROS", "FC", "Cuenta")
    '        Else
    '            FormatCode = FormatCodeProveedor
    '        End If

    '        If FormatCode = "" Then
    '            rsboApp.StatusBar.SetText(NombreAddon + " - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "ERROR - No existe parametrización de cuenta contable para factura de proveedor de servicio, vaya a la opcion de configurar por favor!", Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '            Return False
    '        End If
    '        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '            sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode + "'"
    '        Else
    '            sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode + "'"
    '        End If

    '        Dim Cuenta As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")

    '        Dim oCheckbox As SAPbouiCOM.CheckBox = oForm.Items.Item("chkResum").Specific
    '        If oCheckbox.Checked = True Then
    '            GRPO.Lines.AccountCode = Cuenta
    '            GRPO.Lines.LineTotal = formatDecimal(ofactura.TotalSinImpuesto)

    '            GRPO.Lines.ItemDescription = "SERVICIO"
    '            GRPO.Lines.Quantity = 1
    '            GRPO.Lines.Add()

    '        Else
    '            Dim line As Integer = 0
    '            'Dim CodImpV As SAPbobsCOM.Recordset = rsboApp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            Dim result As String
    '            For Each oDetalle As Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura In ofactura.ENTDetalleFactura


    '                sQueryCodImp = " SELECT ""U_SSID"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje.ToString + "' "
    '                CodImp = oFuncionesB1.getRSvalue(sQueryCodImp, "U_SSID", "")

    '                sQueryCodImpV = " SELECT TOP 1 ""U_SSCOD"" FROM ""@GS_MAPEO_TC"" WHERE ""U_SSCOD"" = '" + oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje.ToString + "' "
    '                CodImpV = oFuncionesB1.getRSvalue(sQueryCodImpV, "U_SSCOD", "")
    '                'CodImpV.DoQuery(sQueryCodImpV)
    '                'result = CodImpV.Fields.Item("U_SSCOD").Value.ToString()

    '                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImp + "Resultado :" + CodImp.ToString(), "frmDocumento")
    '                Utilitario.Util_Log.Escribir_Log("Obteniendo TAXCODE - QUERY: " + sQueryCodImpV + "Resultado :" + CodImpV.ToString(), "frmDocumento")
    '                GRPO.Lines.AccountCode = Cuenta
    '                GRPO.Lines.LineTotal = formatDecimal(oDetalle.PrecioTotalSinImpuesto)
    '                'GRPO.Lines.TaxCode = oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje
    '                Try
    '                    If CodImpV = oDetalle.ENTDetalleFacturaImpuesto(0).CodigoPorcentaje.ToString Then
    '                        GRPO.Lines.TaxCode = CodImp.ToString
    '                    End If

    '                Catch ex As Exception
    '                    Utilitario.Util_Log.Escribir_Log("ERROR: " + ex.ToString(), "frmDocumento")
    '                End Try

    '                GRPO.Lines.ItemDescription = "SERVICIO"
    '                GRPO.Lines.Quantity = 1
    '                GRPO.Lines.Add()
    '                line += 1
    '            Next
    '        End If

    '        'GRPO.Comments += "Creado por el addon SAED"

    '        RetVal = GRPO.Add()
    '        If RetVal <> 0 Then
    '            'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
    '            Elimina_DocumentoRecibido_Factura(DocEntryFVRecibida_UDO)
    '            rCompany.GetLastError(ErrCode, ErrMsg)
    '            rsboApp.MessageBox(ErrCode & " " & ErrMsg)
    '            oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Factura Preliminar de tipo: Servicio:" + ErrCode.ToString + " - " + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '            Return False
    '        Else
    '            rCompany.GetNewObjectCode(sDocEntryPreliminar)
    '            oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "Factura Preliminar de tipo: Servicio, Creada Exitosamente: # " + sDocEntryPreliminar.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '            Actualiza_DocumentoRecibido_Factura(DocEntryFVRecibida_UDO, sDocEntryPreliminar)
    '            Return True
    '        End If

    '    Catch ex As Exception
    '        Elimina_DocumentoRecibido_Factura(DocEntryFVRecibida_UDO)
    '        rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        oFuncionesAddon.GuardaLOG("REE", _oDocumento.ClaveAcceso, "Ocurrio Error al grabar Factura Preliminar de tipo: Servicio:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionPreliminar, Functions.FuncionesAddon.TipoLog.Recepcion)
    '        Return False
    '    Finally
    '        GRPO = Nothing
    '        GC.Collect()
    '    End Try

    'End Function

#End Region
#Region "Funciones Procesar Preliminar"
    Private Function ProcesarPreliminar() As Boolean

        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmProcesoLoteC")

            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
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

            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            Dim DocEntryFacturaRecibida_UDO As String = 0
            Dim Exitoso As Boolean = False


            Dim _fila As Integer


            cbxTipo = oForm.Items.Item("cbxTipo").Specific

            For x = 0 To oDatable.Rows.Count - 1
                For y = 1 To oGrid.Rows.Count - 1
                    ' rsboApp.MessageBox("Y: " + oGrid.GetDataTableRowIndex(y).ToString)
                    indexgrid = oGrid.GetDataTableRowIndex(y)
                    ' rsboApp.MessageBox("Y: " + oGrid.GetDataTableRowIndex(y).ToString)
                    If indexgrid = x Then
                        'gcss.SetRowBackColor(y + 1, RGB(245, 238, 81))
                        Exit For
                    End If

                Next
                Try
                    ofila = indexgrid
                    Dim sPreliminar As String = oDataTable.GetValue(12, ofila).ToString()
                    If sPreliminar <> "0" Then
                        Dim sRUC As String = oDataTable.GetValue(4, ofila).ToString()
                        '_fila = indexgrid
                        If cbxTipo.Value = "07" Then

                            Dim DocEntryPreliminar As String = oDataTable.GetValue(12, ofila).ToString()
                            Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                            ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                            Dim _ruc As String = oDataTable.GetValue(4, ofila).ToString()
                            Dim sQueryCardCode As String = "select ""CardCode"" from ""OCRD"" where ""LicTradNum"" = '" + _ruc.ToString + "' "
                            Dim sCardCode As String = oFuncionesB1.getRSvalue(sQueryCardCode, "CardCode", "")
                            Dim idDocumentoRecibido_UDO As String = ""
                            Dim voucher As String = ""


                            Dim GRPO As SAPbobsCOM.Payments
                            GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
                            GRPO.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                            'GRPO.DocType = SAPbobsCOM.BoRcptTypes.rAccount


                            Dim sQueryCB As String = ""
                            sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardCode""= '" + sCardCode.ToString + "' "
                            Dim clienteBancario As String = oFuncionesB1.getRSvalue(sQueryCB, "U_SSCLIENTEBANCO", "")
                            If clienteBancario = "SI" Then
                                GRPO.DocType = SAPbobsCOM.BoRcptTypes.rAccount
                            End If
                            GRPO.GetByKey(DocEntryPreliminar)


                            If GRPO.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                idDocumentoRecibido_UDO = GRPO.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                            End If

                            Dim idFacturaGS As String = ""
                            Try

                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                Else
                                    idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                End If
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al obtener idDocumentoRecibido_UDO: " + ex.Message.ToString(), "ProcesoLote")
                            End Try
                            Try
                                RetVal = GRPO.SaveDraftToDocument()
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al usar la funcion SaveDraftToDocument : " + ex.Message.ToString(), "ProcesoLote")
                            End Try
                            If RetVal <> 0 Then

                                rCompany.GetLastError(ErrCode, ErrMsg)

                                rsboApp.MessageBox(ErrCode & " " & ErrMsg)

                            Else
                                Try
                                    If ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docFinal", sClaveAcceso) Then
                                        ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1, sClaveAcceso)
                                        MarcarVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO, sClaveAcceso)
                                        Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                                            If rsboApp.Forms.Item("frmProcesoLoteC").Visible = True Then
                                                rsboApp.Forms.Item("frmProcesoLoteC").Freeze(True)
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.Rows.Remove(ofila)
                                                rsboApp.Forms.Item("frmProcesoLoteC").Freeze(False)
                                            End If
                                        Catch ex As Exception
                                        End Try
                                    End If
                                Catch ex As Exception

                                End Try
                            End If

                        Else

                            Try

                                Dim DocEntryPreliminar As String = oDataTable.GetValue(12, ofila).ToString()
                                Dim sClaveAcceso As String = oDataTable.GetValue(7, ofila).ToString()
                                ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                                Dim idDocumentoRecibido_UDO As String = ""
                                Dim GRPO As SAPbobsCOM.Documents
                                GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                GRPO.GetByKey(DocEntryPreliminar)
                                If GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                    idDocumentoRecibido_UDO = GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                End If

                                Dim fechaContabilizacion As String = GRPO.DocDate
                                Dim fechaDocumento As String = GRPO.TaxDate
                                'Dim fecha As String = " SELECT ""TaxDate"" FROM ""OPCH"" WHERE ""U_SSIDDOCUMENTO""= '" + idDocumentoRecibido_UDO.ToString + "' "
                                'Dim _fecha As String = oFuncionesB1.getRSvalue(fecha.ToString, "TaxDate", "")
                                'GRPO.DocDate = Date.Now
                                GRPO.DocDate = Convert.ToDateTime(fechaContabilizacion)
                                'GRPO.DocDueDate = Convert.ToDateTime(fecha)
                                GRPO.TaxDate = Convert.ToDateTime(fechaDocumento)


                                Dim idFacturaGS As String = ""
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_FVR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                                Else
                                    idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_FVR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                                End If
                                Try
                                    RetVal = GRPO.SaveDraftToDocument()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al usar la funcion SaveDraftToDocument : " + ex.Message.ToString(), "ProcesoLote")
                                End Try
                                If RetVal <> 0 Then
                                    rCompany.GetLastError(ErrCode, ErrMsg)
                                    rsboApp.MessageBox(ErrCode & " " & ErrMsg)

                                Else
                                    Try
                                        If ActualizadoEstado_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, "docFinal", sClaveAcceso) Then

                                            ActualizadoEstadoSincronizado_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, 1, sClaveAcceso)
                                            MarcarVisto(Integer.Parse(idFacturaGS), 1, mensaje, idDocumentoRecibido_UDO, sClaveAcceso)
                                            ' SI LA PANTALLA DE DOCUMENTOS RECIBIDOS ESTA ABIERTA ELIMINO LA LINEA DE LA FACTURA RECIBIDA
                                            ' YA QUE YA ESTA INTEGRADA
                                            Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                                                If rsboApp.Forms.Item("frmProcesoLoteC").Visible = True Then
                                                    rsboApp.Forms.Item("frmProcesoLoteC").Freeze(True)
                                                    Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                    odt.Rows.Remove(ofila)
                                                    rsboApp.Forms.Item("frmProcesoLoteC").Freeze(False)
                                                End If
                                            Catch ex As Exception
                                            End Try
                                        End If
                                    Catch ex As Exception
                                        Utilitario.Util_Log.Escribir_Log("Error al Procesar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                                    End Try
                                End If

                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al Procesar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                            End Try
                        End If
                    Else
                        Dim odt2 As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                        odt2.SetValue(14, ofila, "Este preliminar no contiene un preliminar")
                    End If

                    'gcss.SetRowBackColor(y + 1, 255000)
                Catch ex As Exception
                    'gcss.SetRowBackColor(y + 1, RGB(255, 0, 0))
                    Exit For
                End Try
                rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            Next
            rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Exitosamente!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Ocurrio un error en el PL " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

        Return True
    End Function

    Private Function ContabilizarPreliminar() As Boolean

        Try
            oForm = rsboApp.Forms.Item("frmProcesoLoteC")
            'oForm.Freeze(True)
            Dim ErrCode As Long
            Dim ErrMsg As String
            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            Dim RetVal As Long
            cbxTipo = oForm.Items.Item("cbxTipo").Specific
            Dim indexgrid As Integer = 0

            Dim _ActualizadoEstado_DocumentoRecibido_RE As Boolean = False
            Dim _ActualizadoEstadoSincronizado_DocumentoRecibido_RE As Boolean = False
            Dim _MarcarVisto As Boolean = False

            Dim SN_Folio As New List(Of String)

            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                'Next
                If cbxTipo.Value = "07" Then

                    'For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    Dim sPreliminar As String = oGridDet.GetValue("Borrador", i).ToString()

                    If sPreliminar <> "0" Then

                        Dim sClaveAcceso As String = oGridDet.GetValue(7, i).ToString()
                        ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                        Dim _ruc As String = oGridDet.GetValue(4, i).ToString()
                        Dim sQueryCardCode As String = "select ""CardCode"" from ""OCRD"" where ""LicTradNum"" = '" + _ruc.ToString + "' "
                        Dim sCardCode As String = oFuncionesB1.getRSvalue(sQueryCardCode, "CardCode", "")
                        Dim idDocumentoRecibido_UDO As String = ""
                        Dim voucher As String = ""

                        Dim GRPO As SAPbobsCOM.Payments
                        GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
                        GRPO.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                        Dim sQueryCB As String = ""
                        sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardCode""= '" + sCardCode.ToString + "' "
                        Dim clienteBancario As String = oFuncionesB1.getRSvalue(sQueryCB, "U_SSCLIENTEBANCO", "")
                        If clienteBancario = "SI" Then
                            GRPO.DocType = SAPbobsCOM.BoRcptTypes.rAccount
                        End If
                        GRPO.GetByKey(sPreliminar)

                        If GRPO.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                            idDocumentoRecibido_UDO = GRPO.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                        End If

                        Dim idFacturaGS As String = ""
                        Try

                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_RER"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                            Else
                                idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_RER"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Error al obtener idDocumentoRecibido_UDO: " + ex.Message.ToString(), "ProcesoLote")
                        End Try

                        Try
                            RetVal = GRPO.SaveDraftToDocument()
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Error al usar la funcion SaveDraftToDocument : " + ex.Message.ToString(), "ProcesoLote")
                        End Try

                        If RetVal <> 0 Then

                            rCompany.GetLastError(ErrCode, ErrMsg)

                            rsboApp.MessageBox(ErrCode & " " & ErrMsg)

                        Else

                            Try
                                Mensaje_Error = ""
                                If ActualizadoEstado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, "docFinal", sClaveAcceso) Then
                                    Mensaje_Error = ""
                                    If ActualizadoEstadoSincronizado_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1, sClaveAcceso) Then
                                        Mensaje_Error = ""
                                        If MarcarVisto(Integer.Parse(idFacturaGS), 2, mensaje, idDocumentoRecibido_UDO, sClaveAcceso) Then
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, i, "Contabilizado")
                                            SN_Folio.Add(oGridDet.GetValue("RazonSocial", i) + ";" + oGridDet.GetValue("Folio", i))
                                        Else
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, i, "Error: " + Mensaje_Error)
                                        End If
                                    Else
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                        odt.SetValue(14, i, "Error: " + Mensaje_Error)
                                    End If

                                    'Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                                    '    If rsboApp.Forms.Item("frmProcesoLoteC").Visible = True Then
                                    '        rsboApp.Forms.Item("frmProcesoLoteC").Freeze(True)
                                    '        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                    '        odt.Rows.Remove(i)
                                    '        rsboApp.Forms.Item("frmProcesoLoteC").Freeze(False)
                                    '    End If
                                    'Catch ex As Exception
                                    'End Try
                                Else
                                    Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                    odt.SetValue(14, i, "Error: " + Mensaje_Error)
                                End If
                            Catch ex As Exception

                            End Try

                        End If
                    End If


                    'Next

                Else

                    'For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    Try

                        Dim DocEntryPreliminar As String = oGridDet.GetValue(12, i).ToString()
                        'rsboApp.MessageBox(DocEntryPreliminar + " i:" + i.ToString)
                        If DocEntryPreliminar <> "0" Then

                            Dim sClaveAcceso As String = oGridDet.GetValue(7, i).ToString()
                            ' RECUPERO EL ID DE LA FACTURA GS, PARA MARCAR COMO INTEGRADA
                            Dim idDocumentoRecibido_UDO As String = ""
                            Dim GRPO As SAPbobsCOM.Documents
                            GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                            GRPO.GetByKey(DocEntryPreliminar)
                            If GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                idDocumentoRecibido_UDO = GRPO.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                            End If

                            Dim fechaContabilizacion As String = GRPO.DocDate
                            Dim fechaDocumento As String = GRPO.TaxDate

                            GRPO.DocDate = Convert.ToDateTime(fechaContabilizacion)

                            GRPO.TaxDate = Convert.ToDateTime(fechaDocumento)


                            Dim idFacturaGS As String = ""
                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                idFacturaGS = oFuncionesB1.getRSvalue("SELECT ""U_IdGS"" FROM ""@GS_FVR"" WHERE ""DocEntry"" = '" + idDocumentoRecibido_UDO + "'", "U_IdGS", "")
                            Else
                                idFacturaGS = oFuncionesB1.getRSvalue(" select U_IdGS from ""@GS_FVR"" Where DocEntry = " + idDocumentoRecibido_UDO, "U_IdGS", "")
                            End If
                            Try
                                RetVal = GRPO.SaveDraftToDocument()
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al usar la funcion SaveDraftToDocument : " + ex.Message.ToString(), "ProcesoLote")
                            End Try
                            If RetVal <> 0 Then
                                rCompany.GetLastError(ErrCode, ErrMsg)
                                rsboApp.MessageBox(ErrCode & " " & ErrMsg)

                            Else
                                Try
                                    Mensaje_Error = ""
                                    If ActualizadoEstado_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, "docFinal", sClaveAcceso) Then
                                        Mensaje_Error = ""
                                        If ActualizadoEstadoSincronizado_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, 1, sClaveAcceso) Then
                                            Mensaje_Error = ""
                                            If MarcarVisto(Integer.Parse(idFacturaGS), 1, mensaje, idDocumentoRecibido_UDO, sClaveAcceso) Then
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(14, i, "Contabilizado")
                                                SN_Folio.Add(oGridDet.GetValue("RazonSocial", i) + ";" + oGridDet.GetValue("Folio", i))
                                            Else
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                                odt.SetValue(14, i, "Error: " + Mensaje_Error)
                                            End If
                                            '**********para preubas***************
                                            'MarcarVisto(Integer.Parse(idFacturaGS), 1, mensaje, idDocumentoRecibido_UDO, sClaveAcceso)
                                            'Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            'odt.SetValue(14, i, "Contabilizado")
                                            '*************************************
                                        Else
                                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                            odt.SetValue(14, i, "Error: " + Mensaje_Error)
                                        End If

                                    Else
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                                        odt.SetValue(14, i, "Error: " + Mensaje_Error)
                                    End If
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al Procesar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                                End Try
                            End If

                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error al Procesar preliminar: " + ex.Message.ToString(), "ProcesoLote")
                    End Try

                    'Next


                End If
            Next
            'oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            ''Dim contador As Integer = 0 'si funcionaba pero se actualizo por ser mas seguro de que no genere error con el id de l fil
            ''Dim x As Integer, y As Integer
            ''For x = 0 To oGridDet.Rows.Count - 1

            ''    Dim DocEntryPreliminar As String = oGridDet.GetValue(12, x).ToString()
            ''    Dim comentario As String = oGridDet.GetValue(14, x).ToString()

            ''    rsboApp.MessageBox("indexgrid:" + x.ToString)
            ''    If DocEntryPreliminar <> "0" And comentario = "Contabilizado" Then
            ''        contador += 1
            ''        'oGrid.Rows.SelectedRows.Add(x)
            ''    End If

            ''Next

            ''For y = 1 To contador
            ''    BorrarFila()
            ''Next

            ' Eliminar las filas de los documentos contabilizados
            Dim SN As String = ""
            Dim Folio As String = ""
            For j As Integer = 0 To SN_Folio.Count - 1
                Dim ValoresLista = SN_Folio(j)
                For i As Integer = 0 To oGridDet.Rows.Count - 1
                    SN = oGridDet.GetValue("RazonSocial", i)
                    Folio = oGridDet.GetValue("Folio", i)
                    If SN = ValoresLista.Split(";")(0) And Folio = ValoresLista.Split(";")(1) Then
                        oGridDet.Rows.Remove(i)
                        If cbxTipo.Value = "01" Then
                            CargaDocumentosFormato("FE") 'se reordena las filas por eso tengo que validar otros datos como sn y folio ya que al eliminar las filas se cambian de orden
                        ElseIf cbxTipo.Value = "04" Then
                            CargaDocumentosFormato("NE")
                        ElseIf cbxTipo.Value = "07" Then
                            CargaDocumentosFormato("RE")
                        End If
                        oForm.Freeze(False)
                        Exit For
                    End If

                Next
            Next

            Return True
            'oForm.Freeze(True)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Ocurrio un error en el PL " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try


    End Function
    Public Function ActualizadoEstado_DocumentoRecibido_Factura(ByRef DocEntryFacturaRecibida_UDO As String, Estado As String, ByVal sClaveAcceso As String) As Boolean
        'Dim oBusP As SAPbobsCOM.BusinessPartners = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService


        Try
            oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")

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
            Mensaje_Error = "Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString()
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function
    Public Function ActualizadoEstadoSincronizado_DocumentoRecibido_Factura(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer, ByVal sClaveAcceso As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, " Actualizando a Sincronizado = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Sincro", Sincronizado)
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            Mensaje_Error = "Ocurrio al actualizar el estado de la sincronizacion: " + ex.Message.ToString()
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, " Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function
    Public Function MarcarVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String, ByVal sClaveAcceso As String) As Boolean
        Try

            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
            WS.Url = _WS_RecepcionCambiarEstado

            ' MANEJO PROXY
            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmProcesoLoteC")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            If SALIDA_POR_PROXY = "Y" Then
                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmProcesoLoteC")

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

            If WS.MarcarVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, " Documento Marcado como Visto(Integrado) en EDOC ", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                If TipoDocumento = 1 Then
                    If ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, 1, sClaveAcceso) Then
                        Return True
                    Else
                        Return False
                    End If
                    'ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, 1, sClaveAcceso)
                Else
                    If ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1, sClaveAcceso) Then
                        Return True
                    Else
                        Return False
                    End If
                    'ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1, sClaveAcceso)
                End If

                Return True
            Else
                Mensaje_Error = "Error al marcar documento como Visto(Integrado) en EDOC: " + mensaje
                oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, " Error al marcar documento como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS ", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Mensaje_Error = "Error al marcar documento como Visto(Integrado) en EDOC: " + ex.Message.ToString
            Return False
        End Try
    End Function
    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer, ByVal sClaveAcceso As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")

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
            Mensaje_Error = "Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString()
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("REE", sClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstado_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Estado As String, ByVal sClaveAcceso As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, " Actualizando el estado a : " + Estado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

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
            Mensaje_Error = " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString()
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, " Ocurrio al Actualizar Estado del Documento UDO:" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function
    Public Function ActualizadoEstadoSincronizado_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer, ByVal sClaveAcceso As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, " Actualizando a Sincronizado = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Sincro", Sincronizado)
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            Mensaje_Error = " Ocurrio al Actualizar el estado de la sincronizacion: " + ex.Message.ToString()
            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, " Ocurrio error al actualizar el estado de la sincronizacion :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function
    Public Function MarcarVistoRE(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String, ByVal sClaveAcceso As String) As Boolean
        Try
            _WS_Recepcion = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsulta")
            If _WS_Recepcion = "" Then
                rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            _WS_RecepcionCambiarEstado = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionEstado")
            _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

            Dim WS As New Entidades.wsEDoc_ConsultaRecepcionCambiaEstado.WSRAD_KEY_CAMBIARESTADO
            WS.Url = _WS_RecepcionCambiarEstado
            ' MANEJO PROXY
            Dim SALIDA_POR_PROXY As String = ""
            SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "frmProcesoLoteC")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            If SALIDA_POR_PROXY = "Y" Then
                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "frmProcesoLoteC")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "frmProcesoLoteC")

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
            If WS.MarcarVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, " Documento Marcado como Visto(Integrado) en EDOC ", Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 1, sClaveAcceso)
                Return True
            Else
                oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, " Error al marcar documento como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS : " + mensaje, Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer, ByVal sClaveAcceso As String) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

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
            Mensaje_Error = "Error el estado de la sincronizacion EDOC : " + ex.Message.ToString()
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", sClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Sub BorrarFila()
        Try
            Dim oDatable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim x As Integer


            For x = 0 To oDatable.Rows.Count - 1

                Dim DocEntryPreliminar As String = oDatable.GetValue(12, x).ToString()
                Dim comentario As String = oDatable.GetValue(14, x).ToString()

                If DocEntryPreliminar <> "0" And comentario = "Contabilizado" Then


                    Try ' SI ESTA OCULTO E FORMULARIO SE CAE
                        If rsboApp.Forms.Item("frmProcesoLoteC").Visible = True Then
                            rsboApp.Forms.Item("frmProcesoLoteC").Freeze(True)
                            Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmProcesoLoteC").DataSources.DataTables.Item("dtDocs")
                            odt.Rows.Remove(x)
                            rsboApp.Forms.Item("frmProcesoLoteC").Freeze(False)
                        End If
                    Catch ex As Exception
                    End Try

                    Exit For
                End If
            Next

            'Return True
        Catch ex As Exception
            Mensaje_Error = "Error en funcion Borrar Linea:" + ex.Message.ToString
            'Return False
        End Try


    End Sub

    Private Function SeleccionarDocumentosPendientes(ByRef contador As Integer) As Boolean
        Dim resul As Boolean = False
        Try
            oForm = rsboApp.Forms.Item("frmProcesoLoteC")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1

                oGridDet.SetValue("Seleccionar", i, "Y")

                If Not ListaFila.Contains(i) Then
                    ListaFila.Add(i)
                End If

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
            oForm = rsboApp.Forms.Item("frmProcesoLoteC")
            oForm.Freeze(True)
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1

                oGridDet.SetValue("Seleccionar", i, "N")
                If ListaFila.Contains(i) Then
                    ListaFila.Remove(i)
                End If
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
#End Region

End Class

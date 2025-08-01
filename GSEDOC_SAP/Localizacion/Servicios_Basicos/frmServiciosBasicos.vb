Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Drawing.Printing

Imports System.Windows.Forms
Imports Spire.Pdf

Public Class frmServiciosBasicos
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oConditions As SAPbouiCOM.Conditions
    Dim oCondition As SAPbouiCOM.Condition
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim ValorTotal As Double = 0
    Dim Query As String = ""

    Dim txtPI As SAPbouiCOM.EditText
    Dim txtPF As SAPbouiCOM.EditText
    Dim cbxTipSer As SAPbouiCOM.ComboBox
    Dim cbxUM As SAPbouiCOM.ComboBox
    Dim lblFact As SAPbouiCOM.StaticText
    Dim txtFact As SAPbouiCOM.EditText

    Dim lblEst As SAPbouiCOM.StaticText
    Dim cbxEstado As SAPbouiCOM.ComboBox
    Dim btnRpt As SAPbouiCOM.Button
    Dim btnInf As SAPbouiCOM.Button
    Dim btnAnexo As SAPbouiCOM.Button
    Dim txtRuta As SAPbouiCOM.EditText

    Dim btnGuardar As SAPbouiCOM.Button

    Dim Concepto As String = "P/c Pago consumo de {0} Periodo {1} al {2} "

    Dim ConsumoTotal As Integer = 0
    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioServiciosBasicos()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmServiciosBasicos") Then Exit Sub

        strPath = System.Windows.Forms.Application.StartupPath & "\frmServiciosBasicos.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmServiciosBasicos").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmServiciosBasicos")
            oForm.Freeze(True)

            'oForm.EnableMenu("1281", False)
            'oForm.EnableMenu("1282", False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

            ValorTotal = 0

            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            ipLogoSS.Item.Visible = True

            txtPF = oForm.Items.Item("txtPF").Specific
            txtPF.Value = DateTime.Now.ToString("yyyyMMdd")

            txtPI = oForm.Items.Item("txtPI").Specific
            txtPI.Value = DateTime.Now.AddMonths(-1).ToString("yyyyMMdd")

            cbxTipSer = oForm.Items.Item("cbxTipSer").Specific
            cbxTipSer.Select("Agua", SAPbouiCOM.BoSearchKey.psk_ByValue)

            cbxUM = oForm.Items.Item("cbxUM").Specific
            cbxUM.Select("m3", SAPbouiCOM.BoSearchKey.psk_ByValue)

            lblFact = oForm.Items.Item("lblFact").Specific
            lblFact.Item.Visible = False
            txtFact = oForm.Items.Item("txtFact").Specific
            txtFact.Item.Visible = False

            oForm.DataSources.DataTables.Add("dtDocs")

            Query = ""
            Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific
            Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing
            Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Try
                'Query = "SELECT ""PrcCode"", ""PrcName"" FROM ""OPRC"" WHERE ""DimCode"" = 3 AND ""Locked"" = 'N'"
                Query = "SELECT ""U_Sucursal"", ""U_DesSuc"" FROM ""@SS_SB_USR_SUC"" WHERE ""U_Usuario"" = '" & rCompany.UserName & "' AND ""U_TipSer"" = '" & cbxTipSer.Value & "'"
                rst.DoQuery(Query)
                ValoresValidos = cbxCC3.ValidValues

                If rst.RecordCount >= 1 Then
                    While (rst.EoF = False)
                        ValoresValidos.Add(rst.Fields.Item("U_sucursal").Value, rst.Fields.Item("U_DesSuc").Value.ToString)
                        rst.MoveNext()
                    End While
                    cbxCC3.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            Catch ex As Exception

            End Try

            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
            Dim columnaFac As SAPbouiCOM.Column = oMatrix.Columns.Item("U_Factor")
            Dim columnaKil As SAPbouiCOM.Column = oMatrix.Columns.Item("U_KilCon")

            columnaFac.Visible = False
            columnaKil.Visible = False

            CargaDatos()

            ConsumoTotal = 0

            BloqueaControles(False, False)

            Dim NumLecSS As Integer = 0

            Try
                Query = ""
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    ' Query = "SELECT ""AutoKey"" FROM ""ONNM"" WHERE ""ObjectCode"" = 'SSSB'"
                    Query = "SELECT IFNULL(MAX(""DocEntry""),0) + 1 AS ""DocEntry"" FROM ""@SS_SB_CAB"""
                Else
                    'Query = "SELECT AutoKey FROM ONNM WITH(NOLOCK) WHERE ObjectCode = 'SSSB'"
                    Query = "SELECT ISNULL(MAX(DocEntry),0) + 1 AS ""DocEntry"" FROM ""@SS_SB_CAB"""
                End If
                Utilitario.Util_Log.Escribir_Log("Query para obtener siguiente DocEntry: " & Query.ToString, "frmServiciosBasicos")
                NumLecSS = CInt(oFuncionesAddon.getRSvalue(Query, "DocEntry", "0"))
                Utilitario.Util_Log.Escribir_Log("Siguiente DocEntry: " & NumLecSS, "frmServiciosBasicos")
            Catch ex As Exception

            End Try

            Dim DE As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
            DE.Value = NumLecSS.ToString

            Dim txtCon As SAPbouiCOM.EditText = oForm.Items.Item("txtCon").Specific
            txtCon.Value = String.Format(Concepto, cbxTipSer.Value, txtPI.Value.Substring(0, 4) & "-" & txtPI.Value.Substring(4, 2) & "-" & txtPI.Value.Substring(6, 2), txtPF.Value.Substring(0, 4) & "-" & txtPF.Value.Substring(4, 2) & "-" & txtPF.Value.Substring(6, 2))

            oForm.Visible = True
            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error CargaFormularioServiciosBasicos " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Error al cargar formulario CargaFormularioServiciosBasicos: " & ex.Message, "frmServiciosBasicos")
        End Try
    End Sub

    Private Sub CargaDatos()
        Try
            rsboApp.StatusBar.SetText("Consultando facturas, espere por favor! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific

            Query = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Query = "CALL " & rCompany.CompanyDB & ".SS_SS_CONSULTAFACTURAS ('" & txtPI.Value & "','" & txtPF.Value & "','" & cbxTipSer.Value & "', '" & cbxCC3.Value & "')"
            Else
                Query = "EXEC SS_SS_CONSULTAFACTURAS '" & txtPI.Value & "','" & txtPF.Value & "', '" & cbxTipSer.Value & "', '" & cbxCC3.Value & "'"
            End If

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + Query, "frmServiciosBasicos")
            oGrid.DataTable.ExecuteQuery(Query)
            Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + Query, "frmServiciosBasicos")

            FormatoFacturas()

            Dim ac1 As Decimal = 0, ac2 As Decimal = 0, ac3 As Decimal = 0, ac4 As Decimal = 0, ns As Decimal = 0
            SeleccionFacturas(ac1, ac2, ac3, ac4, ns)

            oGrid.DataTable.Rows.Add(1) ' Añade una fila al final
            SeteaTotales(ac1, ac2, ac3, ac4, ns)

            oGrid.CommonSetting.SetRowBackColor(oGrid.DataTable.Rows.Count, RGB(255, 128, 0))

            For i As Integer = 1 To 10 '8 '10
                oGrid.CommonSetting.SetCellEditable(oGrid.DataTable.Rows.Count, i, False)  ' Establece la celda específica como no editable
            Next

            rsboApp.StatusBar.SetText("Facturas cargadas con éxito!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error CargaDatos " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub FormatoFacturas()
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            oGrid.Columns.Item(0).TitleObject.Caption = ""
            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            oGrid.Columns.Item(1).TitleObject.Caption = ""
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(1).Width = 16

            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(1)
            oEditTextColumn.LinkedObjectType = 18

            oGrid.Columns.Item(2).Description = "# Factura"
            oGrid.Columns.Item(2).TitleObject.Caption = "# Factura"
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Sortable = True

            oGrid.Columns.Item(3).Description = "Nombre Proveedor"
            oGrid.Columns.Item(3).TitleObject.Caption = "Nombre Proveedor"
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Sortable = True

            oGrid.Columns.Item(4).Description = "Fecha Factura"
            oGrid.Columns.Item(4).TitleObject.Caption = "Fecha Factura"
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Sortable = True

            oGrid.Columns.Item(5).Description = "Total Factura"
            oGrid.Columns.Item(5).TitleObject.Caption = "Total Factura"
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Sortable = True

            oGrid.Columns.Item(6).Description = "Ajuste"
            oGrid.Columns.Item(6).TitleObject.Caption = "Ajuste"
            oGrid.Columns.Item(6).Editable = True

            oGrid.Columns.Item(7).Description = "Total con ajuste"
            oGrid.Columns.Item(7).TitleObject.Caption = "Total con ajuste"
            oGrid.Columns.Item(7).Editable = False

            oGrid.Columns.Item(8).Description = "Consumo"
            oGrid.Columns.Item(8).TitleObject.Caption = "Consumo"
            oGrid.Columns.Item(8).Editable = True

            oGrid.Columns.Item(9).Description = "Costo"
            oGrid.Columns.Item(9).TitleObject.Caption = "Costo"
            oGrid.Columns.Item(9).Editable = False

            oGrid.Columns.Item(10).Description = "Comentario"
            oGrid.Columns.Item(10).TitleObject.Caption = "Comentario"
            oGrid.Columns.Item(10).Editable = False

            'oGrid.AutoResizeColumns()

            For fila As Integer = 1 To oGrid.Rows.Count
                For columna As Integer = 1 To oGrid.Columns.Count
                    oGrid.CommonSetting.SetCellBackColor(fila, columna, RGB(255, 255, 255))
                    Select Case columna
                        Case 1, 7
                            oGrid.CommonSetting.SetCellEditable(fila, columna, True)
                    End Select
                Next
            Next



        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error FormatoFacturas: {ex.Message}", "frmPagosMasivos")
            rsboApp.StatusBar.SetText($"Error FormatoFacturas: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmServiciosBasicos" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If Not pVal.BeforeAction Then
                            Try
                                Dim txtPI As SAPbouiCOM.EditText = oForm.Items.Item("txtPI").Specific
                                Dim txtPF As SAPbouiCOM.EditText = oForm.Items.Item("txtPF").Specific
                                Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific
                                Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific
                                Dim txtCon As SAPbouiCOM.EditText = oForm.Items.Item("txtCon").Specific
                                If txtPI.Value <> "" Then
                                    Dim fechaIni As DateTime = DateTime.ParseExact(txtPI.Value, "yyyyMMdd", Globalization.CultureInfo.InvariantCulture)
                                End If

                                If txtPF.Value <> "" Then
                                    Dim fechaFin As DateTime = DateTime.ParseExact(txtPF.Value, "yyyyMMdd", Globalization.CultureInfo.InvariantCulture)
                                End If

                                Select Case pVal.ItemUID
                                    Case "cbxTipSer"
                                        oForm.Freeze(True)
                                        CargaDatos()

                                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
                                        Dim columnaFac As SAPbouiCOM.Column = oMatrix.Columns.Item("U_Factor")
                                        Dim columnaKil As SAPbouiCOM.Column = oMatrix.Columns.Item("U_KilCon")
                                        columnaFac.Visible = False
                                        columnaKil.Visible = False
                                        lblFact = oForm.Items.Item("lblFact").Specific
                                        txtFact = oForm.Items.Item("txtFact").Specific
                                        txtFact.Item.Visible = False
                                        lblFact.Item.Visible = False
                                        If cbxTipSer.Value = "Agua" Or cbxTipSer.Value = "Gas" Then
                                            cbxUM.Select("m3", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            If cbxTipSer.Value = "Gas" Then
                                                columnaFac.Visible = True
                                                columnaKil.Visible = True
                                                lblFact.Item.Visible = True
                                                txtFact.Item.Visible = True
                                            End If
                                        Else
                                            cbxUM.Select("Kwh", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If

                                        txtCon.Value = String.Format(Concepto, cbxTipSer.Value, txtPI.Value.Substring(0, 4) & "-" & txtPI.Value.Substring(4, 2) & "-" & txtPI.Value.Substring(6, 2), txtPF.Value.Substring(0, 4) & "-" & txtPF.Value.Substring(4, 2) & "-" & txtPF.Value.Substring(6, 2))

                                        oForm.Freeze(False)
                                    Case "btnCA"
                                        Try
                                            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                            Dim ValidarCosto As String = CStr(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1))
                                            If ValidarCosto = "" Or ValidarCosto = "0" Then

                                                rsboApp.StatusBar.SetText("Por favor seleccione un documento e ingrese un valor de consumo en la sección de facturas!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                            Else

                                                Dim selectFileDialog As New SelectFileDialog("C:\", "", "Archivos de Excel (*.xls)|*.xls|Todos los archivos (*.*)|*.*", DialogType.OPEN)
                                                selectFileDialog.Open()
                                                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                                                    Dim ruta As String = ""
                                                    ruta = selectFileDialog.SelectedFile
                                                    leerXLS(ruta, cbxCC3.Value, cbxTipSer.Value)
                                                End If

                                            End If

                                        Catch ex As Exception
                                            rsboApp.StatusBar.SetText($"Error boton cargar xls: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End Try
                                    Case "btnGP"

                                        GeneraPlantilla()

                                    Case "oGrid"
                                        Try
                                            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                                If pVal.Row >= 0 Then
                                                    If pVal.ColUID = "Tipo" Then
                                                        oForm.Freeze(True)
                                                        Dim ac1 As Decimal = 0, ac2 As Decimal = 0, ac3 As Decimal = 0, ac4 As Decimal = 0, ns As Decimal = 0
                                                        SeleccionFacturas(ac1, ac2, ac3, ac4, ns)
                                                        SeteaTotales(ac1, ac2, ac3, ac4, ns)
                                                        'CalculoDeTotales()
                                                        IngresaResumenFactura()
                                                        oForm.Freeze(False)
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Exception
                                            oForm.Freeze(False)
                                            rsboApp.StatusBar.SetText("Error seccion grid:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End Try

                                    Case "btnAnexo"

                                        Dim selectFileDialog As New SelectFileDialog("C:\", "", "Archivos de Excel (*.xls)|*.xls|Todos los archivos (*.*)|*.*", DialogType.OPEN)
                                        selectFileDialog.Open() '                               

                                        If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                                            Dim ruta As String = ""
                                            ruta = selectFileDialog.SelectedFile

                                            Dim txtRuta As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific
                                            txtRuta.Value = ruta.ToString
                                        End If

                                End Select
                            Catch ex As Exception
                                rsboApp.StatusBar.SetText("Error evento  et_CLICK " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE 'et_LOST_FOCUS
                        Try
                            If Not pVal.Before_Action Then
                                Select Case pVal.ItemUID
                                    Case "oGrid"
                                        Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                        Dim txtFact As SAPbouiCOM.EditText = oForm.Items.Item("txtFact").Specific
                                        Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific
                                        Dim TotalFactura As Decimal = 0, TotalAjuste As Decimal = 0, TotalConAjuste As Decimal = 0, ac1 As Decimal = 0, ac2 As Decimal = 0, ac3 As Decimal = 0, ac4 As Decimal = 0, ns As Decimal = 0
                                        Dim Consumo As Decimal = 0, Costo As Decimal = 0
                                        If pVal.ColUID = "Ajuste" Or pVal.ColUID = "Consumo" Then
                                            oForm.Freeze(True)

                                            If oGrid.DataTable.GetValue("Tipo", pVal.Row).ToString = "Y" Then
                                                TotalFactura = CDec(oGrid.DataTable.GetValue("DocTotal", pVal.Row))
                                                TotalAjuste = CDec(oGrid.DataTable.GetValue("Ajuste", pVal.Row))
                                                TotalConAjuste = TotalFactura - TotalAjuste
                                                oGrid.DataTable.SetValue("TotalConAjuste", pVal.Row, CDbl(TotalConAjuste))

                                                Consumo = CDec(oGrid.DataTable.GetValue("Consumo", pVal.Row))
                                                If Consumo > 0 Then
                                                    'If cbxTipSer.Value <> "Gas" Then
                                                    Costo = Math.Round((TotalConAjuste / Consumo), 6)
                                                    'Else
                                                    'Costo = Math.Round((TotalConAjuste / (Consumo * CDbl(txtFact.Value))), 6)
                                                    'End If
                                                    oGrid.DataTable.SetValue("Costo", pVal.Row, CDbl(Costo))
                                                End If
                                            End If

                                            SeleccionFacturas(ac1, ac2, ac3, ac4, ns) ', ac4, ac5)
                                            SeteaTotales(ac1, ac2, ac3, ac4, ns) ', ac4, ac5)
                                            'CalculoDeTotales()

                                            IngresaResumenFactura()
                                            oForm.Freeze(False)
                                        End If

                                    Case "MTX_SS"
                                        If pVal.ColUID = "U_LecIni" Or pVal.ColUID = "U_LecFin" Then
                                            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
                                            If pVal.Row > 0 Then
                                                oForm.Freeze(True)

                                                Dim activeRow As SAPbouiCOM.CellPosition = oMatrix.GetCellFocus

                                                Dim lblCons As SAPbouiCOM.StaticText = oForm.Items.Item("lblCons").Specific
                                                Dim lblTotal As SAPbouiCOM.StaticText = oForm.Items.Item("lblTotal").Specific

                                                Dim LecIni As Integer = CInt(oMatrix.Columns.Item("U_LecIni").Cells.Item(pVal.Row).Specific.Value)
                                                Dim LecFin As Integer = CInt(oMatrix.Columns.Item("U_LecFin").Cells.Item(pVal.Row).Specific.Value)
                                                Dim ConsumoAnterior As Integer = CInt(oMatrix.Columns.Item("U_Con").Cells.Item(pVal.Row).Specific.Value)
                                                Dim TotalAnterior As Decimal = CDec(oMatrix.Columns.Item("U_Tot").Cells.Item(pVal.Row).Specific.Value)
                                                Dim Costo As Decimal = CDec(oMatrix.Columns.Item("U_Cos").Cells.Item(pVal.Row).Specific.Value)

                                                Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET1")
                                                oDBDataSource.SetValue("U_LecIni", pVal.Row - 1, LecIni)
                                                oDBDataSource.SetValue("U_LecFin", pVal.Row - 1, LecFin)

                                                Dim dif As Integer = LecFin - LecIni
                                                oDBDataSource.SetValue("U_Consumo", pVal.Row - 1, dif)

                                                Dim TotNuevo As Decimal = CDec(dif) * Costo
                                                oDBDataSource.SetValue("U_Total", pVal.Row - 1, TotNuevo)

                                                If dif < 0 Then
                                                    oMatrix.CommonSetting.SetCellBackColor(pVal.Row, 8, ColorTranslator.ToOle(Color.Red))
                                                ElseIf LecIni > LecFin Then
                                                    oMatrix.CommonSetting.SetCellBackColor(pVal.Row, 8, ColorTranslator.ToOle(Color.Red))
                                                ElseIf dif > 0 Then
                                                    oMatrix.CommonSetting.SetCellBackColor(pVal.Row, 8, ColorTranslator.ToOle(Color.LightGreen))
                                                End If

                                                Dim ConsumoActual As Integer = CInt(lblCons.Caption)
                                                lblCons.Caption = CStr(ConsumoActual - ConsumoAnterior + dif)

                                                Dim TotalActual As Decimal = CDec(lblTotal.Caption)
                                                lblTotal.Caption = CStr(TotalActual - TotalAnterior + TotNuevo)

                                                oMatrix.LoadFromDataSource()
                                                oMatrix.SetCellFocus(activeRow.rowIndex, activeRow.ColumnIndex)

                                                IngresaResumenXNivel()

                                                oForm.Freeze(False)
                                            End If
                                        End If
                                        If pVal.ColUID = "U_Obs" Then
                                            oForm.Freeze(True)
                                            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
                                            Dim activeRow As SAPbouiCOM.CellPosition = oMatrix.GetCellFocus
                                            Dim OBS As String = oMatrix.Columns.Item("U_Obs").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET1")
                                            oDBDataSource.SetValue("U_Observacion", pVal.Row - 1, OBS)
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.SetCellFocus(activeRow.rowIndex, activeRow.ColumnIndex)
                                            oForm.Freeze(False)
                                        End If

                                End Select
                            End If
                        Catch ex As Exception
                            rsboApp.StatusBar.SetText("Error evento  et_VALIDATE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Try
                            If Not pVal.BeforeAction Then
                                Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)
                                Dim mEdit As SAPbouiCOM.EditText = Nothing
                                Dim mItem As SAPbouiCOM.IItem = Nothing
                                Try
                                    mItem = oForm.Items.Add("DocEntry1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    mItem.Left = 114
                                    mItem.Top = 9
                                    mItem.Height = 14
                                    mItem.Width = 37
                                    mItem.Enabled = False
                                    mItem.DisplayDesc = False
                                    mItem.Visible = True
                                    mEdit = mItem.Specific
                                    mEdit.DataBind.SetBound(True, "@SS_SB_CAB", "DocEntry")
                                    oFuncionesB1.Release(mItem)
                                    oForm.DataBrowser.BrowseBy = "DocEntry1" 'Next
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                                Catch ex As Exception
                                    rsboApp.MessageBox("Error evento Form_Load: " + ex.Message.ToString())
                                Finally
                                    oFuncionesB1.Release(oForm)
                                    oFuncionesB1.Release(mItem)
                                    oFuncionesB1.Release(mItem)
                                End Try
                            End If
                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "btnGuardar" And pVal.BeforeAction = True Then
                            If btnGuardar.Caption = "Guardar" Then
                                If IngresaRegistroSS() Then
                                    If ActualizaFacturaProveedores() Then
                                        oForm.Freeze(True)
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                                        oForm.Items.Item("btnGP").Visible = False
                                        oForm.Items.Item("btnCA").Visible = False
                                        oForm.Items.Item("oGrid").Enabled = False
                                        oForm.Items.Item("MTX_SS").Enabled = False
                                        oForm.Items.Item("MTX_RES1").Enabled = False
                                        oForm.Items.Item("MTX_RES2").Enabled = False

                                        btnGuardar = oForm.Items.Item("btnGuardar").Specific
                                        btnGuardar.Caption = "Ok"

                                        btnRpt = oForm.Items.Item("btnRpt").Specific
                                        btnRpt.Item.Visible = True


                                        btnInf = oForm.Items.Item("btnInf").Specific
                                        btnInf.Item.Visible = True

                                        oForm.Freeze(False)
                                    End If
                                End If
                            ElseIf btnGuardar.Caption = "Actualizar" Then
                                If ActualizaRegistro() Then
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    btnGuardar.Caption = "Ok"
                                    btnRpt = oForm.Items.Item("btnRpt").Specific
                                    btnRpt.Item.Visible = True

                                    btnInf = oForm.Items.Item("btnInf").Specific
                                    btnInf.Item.Visible = True
                                End If

                            ElseIf btnGuardar.Caption = "Ok" Then
                                oForm.Close()

                            ElseIf btnGuardar.Caption = "Buscar" Then

                                Dim txtDocEntry As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
                                BuscarRegistro(txtDocEntry.Value)

                            End If
                        End If

                        If pVal.ItemUID = "btnRpt" And pVal.BeforeAction = True Then
                            Dim DocEntry1 As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
                            CargaFormatos("Preliminar", CInt(DocEntry1.Value))
                        End If

                        If pVal.ItemUID = "btnInf" And pVal.BeforeAction = True Then
                            Dim DocEntry1 As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
                            CargaFormatos("Informe", CInt(DocEntry1.Value))
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        Try
                            If Not pVal.BeforeAction Then
                                AdaptarTamano()
                            End If
                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Try
                            If Not pVal.Before_Action Then
                                Select Case pVal.ItemUID
                                    Case "txtPI"

                                        Dim txtPI As SAPbouiCOM.EditText = oForm.Items.Item("txtPI").Specific
                                        Dim txtPF As SAPbouiCOM.EditText = oForm.Items.Item("txtPF").Specific
                                        Dim txtCon As SAPbouiCOM.EditText = oForm.Items.Item("txtCon").Specific
                                        Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific

                                        txtCon.Value = String.Format(Concepto, cbxTipSer.Value, txtPI.Value.Substring(0, 4) & "-" & txtPI.Value.Substring(4, 2) & "-" & txtPI.Value.Substring(6, 2), txtPF.Value.Substring(0, 4) & "-" & txtPF.Value.Substring(4, 2) & "-" & txtPF.Value.Substring(6, 2))

                                    Case "txtPF"

                                        Dim txtPI As SAPbouiCOM.EditText = oForm.Items.Item("txtPI").Specific
                                        Dim txtPF As SAPbouiCOM.EditText = oForm.Items.Item("txtPF").Specific
                                        Dim txtCon As SAPbouiCOM.EditText = oForm.Items.Item("txtCon").Specific
                                        Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific

                                        txtCon.Value = String.Format(Concepto, cbxTipSer.Value, txtPI.Value.Substring(0, 4) & "-" & txtPI.Value.Substring(4, 2) & "-" & txtPI.Value.Substring(6, 2), txtPF.Value.Substring(0, 4) & "-" & txtPF.Value.Substring(4, 2) & "-" & txtPF.Value.Substring(6, 2))
                                End Select
                            End If
                        Catch ex As Exception

                        End Try
                End Select
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("catch rsboApp_ItemEvent: " & ex.Message.ToString(), "frmServiciosBasicos")
        End Try

    End Sub

    Private Sub AdaptarTamano()

        Try
            Dim MTX_RES1 As SAPbouiCOM.Item = oForm.Items.Item("MTX_RES1")
            Dim MTX_RES2 As SAPbouiCOM.Item = oForm.Items.Item("MTX_RES2")

            MTX_RES1.Width = 397
            MTX_RES1.Top = 124
            MTX_RES1.Height = 102


            MTX_RES2.Width = 397

            TryCast(MTX_RES1.Specific, SAPbouiCOM.Matrix).AutoResizeColumns()
            TryCast(MTX_RES2.Specific, SAPbouiCOM.Matrix).AutoResizeColumns()

            Dim MTX_SS As SAPbouiCOM.Item = oForm.Items.Item("MTX_SS")
            MTX_SS.Width = 812
            MTX_SS.Top = 239

            TryCast(MTX_SS.Specific, SAPbouiCOM.Matrix).AutoResizeColumns()

            Dim oGrid As SAPbouiCOM.Item = oForm.Items.Item("oGrid")
            oGrid.Width = 812
            oGrid.Top = 124
            oGrid.Height = 102

        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error al Ajustar Dimension Matrix" & ex.Message)
        End Try
    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            Dim typeExx, idFormm As String
            typeExx = oFuncionesB1.FormularioActivo(idFormm)

            If typeExx = "frmServiciosBasicos" Then
                If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And Not pVal.BeforeAction Then
                    oForm.Freeze(True)

                    oForm.Items.Item("btnGP").Visible = False
                    oForm.Items.Item("btnCA").Visible = False
                    oForm.Items.Item("MTX_SS").Enabled = False
                    oForm.Items.Item("MTX_RES1").Enabled = False
                    oForm.Items.Item("MTX_RES2").Enabled = False
                    oForm.Items.Item("oGrid").Enabled = False

                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                    oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                    oGrid.DataTable.Clear()

                    Dim DocEntryUDOSB As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
                    Dim query As String = ""
                    query = $"Select 'Y' AS ""Tipo"", A.""DocEntry"", A.""DocNum"", A.""CardName"", A.""DocDate"", B.""U_Valor"", 0.0 AS ""Ajuste"", B.""U_Valor"", B.""U_Consumo"", B.""U_Costo"", A.""Comments"" AS ""Comentario"" FROM ""OPCH"" A INNER JOIN ""@SS_SB_DET2"" B ON A.""U_SS_ServicioBasico"" = B.""DocEntry"" WHERE A.""U_SS_ServicioBasico"" = '{DocEntryUDOSB.Value}' AND A.""DocEntry"" = B.""U_DocEntryFac"""

                    Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + query, "frmServiciosBasicos")
                    oGrid.DataTable.ExecuteQuery(query)
                    Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + query, "frmServiciosBasicos")
                    FormatoFacturas()

                    BloqueaControles(True, True)

                    Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        query = $"SELECT ""U_Costo"", ""U_Total"" FROM ""@SS_SB_DET3"" WHERE ""DocEntry"" =  ' {DocEntryUDOSB.Value} ' AND ""U_Nivel"" = 'Total Facturar locales'"
                        rst.DoQuery(query)

                        Dim lblCons As SAPbouiCOM.StaticText = oForm.Items.Item("lblCons").Specific
                        Dim lblTotal As SAPbouiCOM.StaticText = oForm.Items.Item("lblTotal").Specific
                        If rst.RecordCount >= 1 Then
                            While (rst.EoF = False)
                                lblCons.Caption = CStr(rst.Fields.Item("U_Costo").Value)
                                lblTotal.Caption = CStr(rst.Fields.Item("U_Total").Value)
                                rst.MoveNext()
                            End While

                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error al consultar resumen NAVEGACION: " & ex.Message & " - query:" & query, "frmServiciosBasicos")
                    End Try

                    oForm.Freeze(False)
                ElseIf pVal.MenuUID = "1282" And Not pVal.BeforeAction Then 'NUEVO

                    Try
                        oForm.Freeze(True)

                        oForm.Items.Item("MTX_SS").Enabled = True
                        oForm.Items.Item("oGrid").Enabled = True

                        txtPF = oForm.Items.Item("txtPF").Specific
                        txtPF.Value = DateTime.Now.ToString("yyyyMMdd")

                        txtPI = oForm.Items.Item("txtPI").Specific
                        txtPI.Value = DateTime.Now.AddMonths(-1).ToString("yyyyMMdd")

                        cbxTipSer = oForm.Items.Item("cbxTipSer").Specific
                        cbxTipSer.Select("Agua", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        cbxUM = oForm.Items.Item("cbxUM").Specific
                        cbxUM.Select("m3", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        Query = ""
                        Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific
                        Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing
                        Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Try
                            'Query = "SELECT ""PrcCode"", ""PrcName"" FROM ""OPRC"" WHERE ""DimCode"" = 3 AND ""Locked"" = 'N'"
                            Query = "SELECT ""U_Sucursal"", ""U_DesSuc"" FROM ""@SS_SB_USR_SUC"" WHERE ""U_Usuario"" = '" & rCompany.UserName & "' AND ""U_TipSer"" = '" & cbxTipSer.Value & "'"
                            rst.DoQuery(Query)
                            ValoresValidos = cbxCC3.ValidValues

                            If rst.RecordCount >= 1 Then
                                While (rst.EoF = False)
                                    ValoresValidos.Add(rst.Fields.Item("U_sucursal").Value, rst.Fields.Item("U_DesSuc").Value.ToString)
                                    rst.MoveNext()
                                End While
                                cbxCC3.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            End If
                        Catch ex As Exception

                        End Try

                        CargaDatos()

                        Dim NumLecSS As Integer = 0

                        Try
                            Query = ""
                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                ' Query = "SELECT ""AutoKey"" FROM ""ONNM"" WHERE ""ObjectCode"" = 'SSSB'"
                                Query = "SELECT IFNULL(MAX(""DocEntry""),0) + 1 AS ""DocEntry"" FROM ""@SS_SB_CAB"""
                            Else
                                'Query = "SELECT AutoKey FROM ONNM WITH(NOLOCK) WHERE ObjectCode = 'SSSB'"
                                Query = "SELECT ISNULL(MAX(DocEntry),0) + 1 AS ""DocEntry"" FROM ""@SS_SB_CAB"""
                            End If
                            Utilitario.Util_Log.Escribir_Log("Query para obtener siguiente DocEntry: " & Query.ToString, "frmServiciosBasicos")
                            NumLecSS = CInt(oFuncionesAddon.getRSvalue(Query, "DocEntry", "0"))
                            Utilitario.Util_Log.Escribir_Log("Siguiente DocEntry: " & NumLecSS, "frmServiciosBasicos")
                        Catch ex As Exception

                        End Try

                        Dim DE As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
                        DE.Value = NumLecSS.ToString

                        BloqueaControles(False, False)

                        btnGuardar = oForm.Items.Item("btnGuardar").Specific
                        btnGuardar.Caption = "Guardar"

                        oForm.Freeze(False)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent NUEVO: {ex.Message}", "frmServiciosBasicos")
                        oForm.Freeze(False)
                    End Try

                ElseIf pVal.MenuUID = "1281" Then

                    oForm.Items.Item("DocEntry1").Enabled = True

                    btnGuardar = oForm.Items.Item("btnGuardar").Specific
                    btnGuardar.Caption = "Buscar"


                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmServiciosBasicos")
        End Try
    End Sub

    Private Sub BloqueaControles(ByVal V1 As Boolean, ByVal V3 As Boolean)
        Try
            btnGuardar = oForm.Items.Item("btnGuardar").Specific
            cbxEstado = oForm.Items.Item("cbxEstado").Specific
            btnRpt = oForm.Items.Item("btnRpt").Specific
            btnInf = oForm.Items.Item("btnInf").Specific

            oForm.Items.Item("lblEst").Visible = V1
            cbxEstado.Item.Visible = V1
            oForm.Items.Item("btnAnexo").Visible = V1
            oForm.Items.Item("txtRuta").Visible = V1
            btnRpt.Item.Visible = V3
            btnInf.Item.Visible = V3

            If cbxEstado.Value <> "" Then btnGuardar.Caption = "Actualizar"

        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error BloqueaControles " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function SeleccionFacturas(ByRef ac1 As Double, ByRef ac2 As Double, ByRef ac3 As Double, ByRef ac4 As Double, ByRef ns As Double)
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.DataTable.GetValue("Tipo", i).ToString = "Y" Then
                    ac1 += CDec(oGrid.DataTable.GetValue("DocTotal", i))
                    ac2 += CDec(oGrid.DataTable.GetValue("TotalConAjuste", i))
                    ac3 += CDec(oGrid.DataTable.GetValue("Consumo", i))
                    ac4 += CDec(oGrid.DataTable.GetValue("Costo", i))
                    ns += 1
                End If
            Next
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error SeleccionFacturas " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function SeteaTotales(ByVal ac1 As Double, ByVal ac2 As Double, ByVal ac3 As Double, ByVal ac4 As Double, ByVal ns As Double)
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable.SetValue("DocTotal", oGrid.DataTable.Rows.Count - 1, CDbl(ac1))
            oGrid.DataTable.SetValue("TotalConAjuste", oGrid.DataTable.Rows.Count - 1, CDbl(ac2))
            oGrid.DataTable.SetValue("Consumo", oGrid.DataTable.Rows.Count - 1, CDbl(ac3))

            If CDbl(ac3) > 0 Then
                If ns > 0 Then oGrid.DataTable.SetValue("Costo", oGrid.DataTable.Rows.Count - 1, Math.Round(CDbl(ac2) / CDbl(ac3), 6)) 'Math.Round((CDbl(ac4) / CDbl(ns)), 6)) 'se cambio calculo a peticion de seruvi
            Else
                oGrid.DataTable.SetValue("Costo", oGrid.DataTable.Rows.Count - 1, 0)
            End If

        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error SeteaTotales " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function GeneraPlantilla()
        Try
            Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific
            Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific

            If String.IsNullOrEmpty(cbxCC3.Value) Then
                rsboApp.StatusBar.SetText("Seleccione una sucursal por favor!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                rsboApp.StatusBar.SetText("Generando plantilla, espere por favor! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                cbxTipSer = oForm.Items.Item("cbxTipSer").Specific

                Dim excelApp As New Excel.Application()
                Dim workbook As Excel.Workbook = excelApp.Workbooks.Add
                Dim worksheet As Excel.Worksheet = CType(workbook.Sheets(1), Excel.Worksheet)

                worksheet.Range("A1", "H1").Font.Bold = True
                worksheet.Range("A1", "H1").Font.Size = 12

                worksheet.Cells(1, "A").Value = "Contrato"
                worksheet.Cells(1, "B").Value = "Contrato Anterior"
                worksheet.Cells(1, "C").Value = "Denominacion"
                worksheet.Cells(1, "D").Value = "Local"
                worksheet.Cells(1, "E").Value = "Nivel"
                worksheet.Cells(1, "F").Value = "Medidor"
                worksheet.Cells(1, "G").Value = "Lectura Inicial"
                worksheet.Cells(1, "H").Value = "Lectura Final"

                Query = ""
                Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Query = "CALL " & rCompany.CompanyDB & ".SS_GENERAPLANTILLA ('" & cbxCC3.Value & "', '" & cbxTipSer.Value & "')"
                Else
                    Query = "EXEC SS_GENERAPLANTILLA '" & cbxCC3.Value & "', '" & cbxTipSer.Value & "'"
                End If

                rst.DoQuery(Query)

                If rst.RecordCount >= 1 Then
                    Dim i As Integer = 2
                    While (rst.EoF = False)
                        worksheet.Cells(i, "A") = rst.Fields.Item("Contrato").Value
                        worksheet.Cells(i, "B") = rst.Fields.Item("ContratoAnterior").Value
                        worksheet.Cells(i, "C") = rst.Fields.Item("Denominacion").Value
                        worksheet.Cells(i, "D") = rst.Fields.Item("Locales").Value
                        worksheet.Cells(i, "E") = rst.Fields.Item("Nivel").Value
                        worksheet.Cells(i, "F") = rst.Fields.Item("Medidor").Value
                        worksheet.Cells(i, "G") = rst.Fields.Item("LecturaInicial").Value
                        worksheet.Cells(i, "H") = 0
                        i += 1
                        rst.MoveNext()
                    End While
                End If

                worksheet.Columns("A:H").AutoFit()

                Dim selectFileDialog As New SelectFileDialog("C:\", "", "Todos los archivos (*.*)|*.*", DialogType.FOLDER)
                selectFileDialog.Open()
                Dim ruta As String = ""

                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFolder) Then
                    ruta = selectFileDialog.SelectedFolder
                End If

                'Dim downloadsPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
                Dim fileName As String = cbxCC3.Value.ToString & "_ArchivoPlantilla_" & cbxTipSer.Value.ToString & "" & Date.Now.ToLongDateString.Replace(":", "") & ".xlsx"
                Dim filePath As String = Path.Combine(ruta, fileName) ' Path.Combine(downloadsPath, fileName)

                workbook.SaveAs(filePath)
                workbook.Close()
                excelApp.Quit()

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

                rsboApp.StatusBar.SetText("Plantilla generada con éxito! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error generando plantilla: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function IngresaRegistroSS() As Boolean
        Try

            Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific
            Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific
            If String.IsNullOrEmpty(cbxCC3.Value) Then
                rsboApp.StatusBar.SetText("No puede generar registros si no selecciona una sucursal!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else

                Dim txtPI As SAPbouiCOM.EditText = oForm.Items.Item("txtPI").Specific
                Dim txtPF As SAPbouiCOM.EditText = oForm.Items.Item("txtPF").Specific

                Query = ""
                Query = "SELECT COUNT(*) As ""Cantidad"", MAX(""U_FinPer"") As ""UltimaFecha"" FROM ""@SS_SB_CAB"" "
                Query += "WHERE (""U_IniPer"" BETWEEN '" & txtPI.Value & "' AND '" & txtPF.Value & "' OR ""U_FinPer"" BETWEEN '" & txtPI.Value & "' AND '" & txtPF.Value & "') AND ""U_NivCC3"" = '" & cbxCC3.Value & "' AND ""U_TipSer"" = '" & cbxTipSer.Value & "'"

                Dim Cantidad As Integer = CInt(oFuncionesAddon.getRSvalue(Query, "Cantidad", "0"))

                If Cantidad > 0 Then

                    Dim FechaUltima As String = oFuncionesAddon.getRSvalue(Query, "UltimaFecha", "")
                    rsboApp.StatusBar.SetText("Ya existe un registro dentro de este rango de fecha y sucursal, fecha máxima: " & FechaUltima & " Revise por favor!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Else

                    Dim oGeneralService As SAPbobsCOM.GeneralService
                    Dim oGeneralData As SAPbobsCOM.GeneralData
                    Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                    Dim oCompanyService As SAPbobsCOM.CompanyService
                    Dim oChildren As SAPbobsCOM.GeneralDataCollection
                    Dim oChild As SAPbobsCOM.GeneralData

                    rsboApp.StatusBar.SetText("Registrando información de Servicios Básicos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oForm = rsboApp.Forms.Item("frmServiciosBasicos")

                    oCompanyService = rCompany.GetCompanyService
                    oGeneralService = oCompanyService.GetGeneralService("SSSB")
                    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                    Dim fi As String = oForm.Items.Item("txtPI").Specific.Value '
                    Dim FechaInicio As Date = Date.ParseExact(fi, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    oGeneralData.SetProperty("U_IniPer", FechaInicio.ToString("yyyy/MM/dd")) '("MM/dd/yyyy"))

                    Dim ff As String = oForm.Items.Item("txtPF").Specific.Value '
                    Dim FechaFin As Date = Date.ParseExact(ff, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    oGeneralData.SetProperty("U_FinPer", FechaFin.ToString("yyyy/MM/dd")) '("MM/dd/yyyy"))

                    oGeneralData.SetProperty("U_NivCC3", oForm.Items.Item("cbxCC3").Specific.Value.ToString())
                    oGeneralData.SetProperty("U_TipSer", oForm.Items.Item("cbxTipSer").Specific.Value.ToString())
                    oGeneralData.SetProperty("U_UM", oForm.Items.Item("cbxUM").Specific.Value.ToString())
                    oGeneralData.SetProperty("U_Factor", oForm.Items.Item("txtFact").Specific.Value.ToString())
                    oGeneralData.SetProperty("U_Concepto", oForm.Items.Item("txtCon").Specific.Value.ToString())
                    oGeneralData.SetProperty("U_Estado", "Borrador")

                    oChildren = oGeneralData.Child("SS_SB_DET1")
                    Dim matrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
                    For i As Integer = 1 To matrix.RowCount
                        oChild = oChildren.Add
                        Dim oCheckBox As SAPbouiCOM.CheckBox = CType(matrix.Columns.Item("U_SujFac").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                        oChild.SetProperty("U_SujetoFact", If(oCheckBox.Checked, "Y", "N"))

                        oChild.SetProperty("U_Contrato", matrix.Columns.Item("U_Contrato").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_ContratoAnt", matrix.Columns.Item("U_ConAnt").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_Denominacion", matrix.Columns.Item("U_Den").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_Locales", matrix.Columns.Item("U_Locales").Cells.Item(i).Specific.Value)
                        oChild.SetProperty("U_LecIni", CInt(matrix.Columns.Item("U_LecIni").Cells.Item(i).Specific.Value))
                        oChild.SetProperty("U_LecFin", CInt(matrix.Columns.Item("U_LecFin").Cells.Item(i).Specific.Value))
                        oChild.SetProperty("U_Consumo", CInt(matrix.Columns.Item("U_Con").Cells.Item(i).Specific.Value))
                        oChild.SetProperty("U_Factor", CDbl(matrix.Columns.Item("U_Factor").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Costo", CDbl(matrix.Columns.Item("U_Cos").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Total", CDbl(matrix.Columns.Item("U_Tot").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Observacion", matrix.Columns.Item("U_Obs").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_KilCon", CDbl(matrix.Columns.Item("U_KilCon").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Nivel", matrix.Columns.Item("U_Nivel").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_Medidor", matrix.Columns.Item("U_Medidor").Cells.Item(i).Specific.Value.ToString)
                    Next

                    oChildren = oGeneralData.Child("SS_SB_DET3")
                    Dim MTX_RES2 As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_RES2").Specific
                    For i As Integer = 1 To MTX_RES2.RowCount
                        oChild = oChildren.Add
                        oChild.SetProperty("U_Nivel", MTX_RES2.Columns.Item("U_Niv").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_Consumo", CInt(MTX_RES2.Columns.Item("U_Con").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Costo", CDbl(MTX_RES2.Columns.Item("U_Cos").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Total", CDbl(MTX_RES2.Columns.Item("U_Tot").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Porcentaje", CDbl(MTX_RES2.Columns.Item("U_Por").Cells.Item(i).Specific.Value.ToString))
                    Next

                    oChildren = oGeneralData.Child("SS_SB_DET2")
                    Dim MTX_RES1 As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_RES1").Specific
                    For i As Integer = 1 To MTX_RES1.RowCount
                        oChild = oChildren.Add
                        oChild.SetProperty("U_DocEntryFac", MTX_RES1.Columns.Item("DEFP").Cells.Item(i).Specific.Value.ToString)
                        oChild.SetProperty("U_Valor", CDbl(MTX_RES1.Columns.Item("U_Val").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Consumo", CInt(MTX_RES1.Columns.Item("U_Con").Cells.Item(i).Specific.Value.ToString))
                        oChild.SetProperty("U_Costo", CDbl(MTX_RES1.Columns.Item("U_Cos").Cells.Item(i).Specific.Value.ToString))
                    Next

                    oGeneralParams = oGeneralService.Add(oGeneralData)
                    rsboApp.StatusBar.SetText("Información ingresada con éxito", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                End If
            End If
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error al ingresar información: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al ingresar información: " & ex.Message, "frmServiciosBasicos")
            Return False
        End Try
    End Function

    Public Function ActualizaFacturaProveedores() As Boolean
        Try
            If Date.Now.Month = 1 And Date.Now.Year = 2025 Then
                Return True
            Else
                Dim resultado As Integer = -1
                Dim ErrCode As Long
                Dim ErrMsg As String
                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                Dim DocEntryUDOSB As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific
                For o As Integer = 0 To oGrid.Rows.Count - 1
                    If Convert.ToString(oGrid.DataTable.GetValue("Tipo", o)) = "Y" Then
                        Dim DocEntry_fac As String = oGrid.DataTable.GetValue("DocEntry", o)

                        Dim oDocumento As SAPbobsCOM.Documents = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        'oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices 'Probar
                        'oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None 'Probar

                        If oDocumento.GetByKey(DocEntry_fac) Then

                            rsboApp.StatusBar.SetText("Actualizando factura!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            oDocumento.UserFields.Fields.Item("U_SS_ServicioBasico").Value = DocEntryUDOSB.Value
                            resultado = oDocumento.Update()

                            If resultado = 0 Then
                                rsboApp.StatusBar.SetText("Factura actualizada!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Else
                                rCompany.GetLastError(ErrCode, ErrMsg)
                                rsboApp.SetStatusBarMessage($"Error al actualizar factura: {DocEntry_fac} #Error: {ErrCode} Mensaje: {ErrMsg} ", SAPbouiCOM.BoMessageTime.bmt_Long, True)
                                Utilitario.Util_Log.Escribir_Log($"Error al actualizar factura: {DocEntry_fac} #Error: {ErrCode} Mensaje: {ErrMsg} ", "frmServiciosBasicos")
                            End If
                        End If
                    End If
                Next
                Return True
            End If
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error al actualizar factura de proveedores: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al actualizar factura de proveedores: " & ex.Message, "frmServiciosBasicos")
            Return False
        End Try
    End Function

    Public Function ActualizaRegistro() As Boolean
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService

            Dim DocEntry1 As SAPbouiCOM.EditText = oForm.Items.Item("DocEntry1").Specific

            rsboApp.StatusBar.SetText($"Actualizando registro de servicio básico: {DocEntry1.Value}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim cbxEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
            Dim txtRuta As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSSB")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", CInt(DocEntry1.Value))
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_Estado", cbxEstado.Value)
            oGeneralData.SetProperty("U_Ruta", txtRuta.Value)

            oGeneralService.Update(oGeneralData)
            rsboApp.StatusBar.SetText("Solicitud de servicio básico actualizada con éxito ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error al actualizar información: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al actualizar información: " & ex.Message, "frmServiciosBasicos")
            Return False
        End Try
    End Function

    Function leerXLS(ByVal ruta As String, ByVal Sucursal As String, ByVal TipoServicio As String) As Boolean
        Try
            oForm.Freeze(True)

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific

            rsboApp.StatusBar.SetText("Cargando archivo, espere por favor!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific

            Dim excelApp As New Excel.Application()
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(ruta)
            Dim worksheet As Excel.Worksheet = workbook.Sheets(1)

            'Dim DatosEnExcel As New List(Of Tuple(Of String, String, String, String, String, String, String))
            Dim DatosEnExcel As List(Of DatosExcel)
            DatosEnExcel = New List(Of DatosExcel)

            Dim lastRow As Integer = worksheet.Cells(worksheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim i As Integer = 2
            For rowIndex As Integer = 2 To lastRow
                If Not worksheet.Cells(rowIndex, 1).Value Is Nothing Then
                    rsboApp.StatusBar.SetText("linea " & i.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)

                    Dim datoexcel As DatosExcel = New DatosExcel

                    If worksheet.Cells(rowIndex, 1).Value IsNot Nothing Then datoexcel.Contrato = CInt(worksheet.Cells(rowIndex, 1).Value.ToString)
                    If worksheet.Cells(rowIndex, 2).Value IsNot Nothing Then datoexcel.ContratoAnterior = CInt(worksheet.Cells(rowIndex, 2).Value.ToString)
                    If worksheet.Cells(rowIndex, 3).Value IsNot Nothing Then datoexcel.Denominacion = worksheet.Cells(rowIndex, 3).Value.ToString
                    If worksheet.Cells(rowIndex, 4).Value IsNot Nothing Then datoexcel.locales = worksheet.Cells(rowIndex, 4).Value.ToString
                    If worksheet.Cells(rowIndex, 5).Value IsNot Nothing Then datoexcel.Nivel = worksheet.Cells(rowIndex, 5).Value.ToString
                    If worksheet.Cells(rowIndex, 6).Value IsNot Nothing Then datoexcel.Medidor = worksheet.Cells(rowIndex, 6).Value.ToString
                    If worksheet.Cells(rowIndex, 7).Value IsNot Nothing Then datoexcel.LecturaInicial = CInt(worksheet.Cells(rowIndex, 7).Value.ToString)
                    If worksheet.Cells(rowIndex, 8).Value IsNot Nothing Then datoexcel.LecturaFinal = CInt(worksheet.Cells(rowIndex, 8).Value.ToString)
                    DatosEnExcel.Add(datoexcel)

                    'DatosEnExcel.Add(Tuple.Create(worksheet.Cells(rowIndex, 1).Value.ToString, worksheet.Cells(rowIndex, 2).Value.ToString,
                    '                              IIf(worksheet.Cells(rowIndex, 3).Value IsNot Nothing, worksheet.Cells(rowIndex, 3).Value.ToString, "").ToString,
                    '                              IIf(worksheet.Cells(rowIndex, 4).Value IsNot Nothing, worksheet.Cells(rowIndex, 4).Value.ToString, "").ToString,
                    '                              IIf(worksheet.Cells(rowIndex, 5).Value IsNot Nothing, worksheet.Cells(rowIndex, 5).Value.ToString, "").ToString,
                    '                              IIf(worksheet.Cells(rowIndex, 6).Value IsNot Nothing, worksheet.Cells(rowIndex, 6).Value.ToString, "").ToString,
                    '                              IIf(worksheet.Cells(rowIndex, 7).Value IsNot Nothing, worksheet.Cells(rowIndex, 7).Value.ToString, "").ToString))
                    i += 1
                    End If
            Next

            workbook.Close(False)
            excelApp.Quit()

            'DatosEnExcel = DatosEnExcel.OrderBy(Function(x) x.Item1).ToList()
            'ConsumoTotal = DatosEnExcel.Sum(Function(x) CInt(x.Item7) - CInt(x.Item6))

            DatosEnExcel = DatosEnExcel.OrderBy(Function(x) x.Contrato).ToList()
            ConsumoTotal = DatosEnExcel.Sum(Function(x) x.LecturaFinal - x.LecturaInicial)

            Dim txtFact As SAPbouiCOM.EditText = oForm.Items.Item("txtFact").Specific
            Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET1")
            oDataSource.Clear()
            oMatrix.Clear()

            For j As Integer = 0 To DatosEnExcel.Count - 1
                oDataSource.InsertRecord(j)
                'oDataSource.SetValue("U_Contrato", j, DatosEnExcel(j).Item1)
                'oDataSource.SetValue("U_ContratoAnt", j, DatosEnExcel(j).Item2)
                'oDataSource.SetValue("U_Denominacion", j, DatosEnExcel(j).Item3)
                'oDataSource.SetValue("U_Locales", j, DatosEnExcel(j).Item4)
                'oDataSource.SetValue("U_Nivel", j, DatosEnExcel(j).Item5)
                'oDataSource.SetValue("U_LecIni", j, DatosEnExcel(j).Item6)
                'oDataSource.SetValue("U_LecFin", j, DatosEnExcel(j).Item7)
                'oDataSource.SetValue("U_Consumo", j, CInt(DatosEnExcel(j).Item7) - CInt(DatosEnExcel(j).Item6))
                'oDataSource.SetValue("U_Factor", j, CDec(txtFact.Value))
                'oDataSource.SetValue("U_KilCon", j, CDec(txtFact.Value) * CDec(CInt(DatosEnExcel(j).Item7) - CInt(DatosEnExcel(j).Item6)))
                'oDataSource.SetValue("U_Costo", j, CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1))) 'CDec(0))
                'Dim ConsumoMatrix As Integer = CInt(DatosEnExcel(j).Item7) - CInt(DatosEnExcel(j).Item6)
                'Dim tot As Decimal = 0
                'If cbxTipSer.Value <> "Gas" Then
                '    tot = CDec(ConsumoMatrix) * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) 'Costo
                'Else
                '    tot = CDec(CDec(txtFact.Value) * CDec(CInt(DatosEnExcel(j).Item7) - CInt(DatosEnExcel(j).Item6))) * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) ' Costo
                'End If

                'oDataSource.SetValue("U_Total", j, tot)

                oDataSource.SetValue("U_Contrato", j, DatosEnExcel(j).Contrato)
                oDataSource.SetValue("U_ContratoAnt", j, DatosEnExcel(j).ContratoAnterior)
                oDataSource.SetValue("U_Denominacion", j, DatosEnExcel(j).Denominacion)
                oDataSource.SetValue("U_Locales", j, DatosEnExcel(j).locales)
                oDataSource.SetValue("U_Nivel", j, DatosEnExcel(j).Nivel)
                oDataSource.SetValue("U_Medidor", j, DatosEnExcel(j).Medidor)
                oDataSource.SetValue("U_LecIni", j, DatosEnExcel(j).LecturaInicial)
                oDataSource.SetValue("U_LecFin", j, DatosEnExcel(j).LecturaFinal)
                oDataSource.SetValue("U_Consumo", j, CInt(DatosEnExcel(j).LecturaFinal) - CInt(DatosEnExcel(j).LecturaInicial))
                oDataSource.SetValue("U_Factor", j, CDec(txtFact.Value))
                oDataSource.SetValue("U_KilCon", j, CDec(txtFact.Value) * CDec(CInt(DatosEnExcel(j).LecturaFinal) - CInt(DatosEnExcel(j).LecturaInicial)))
                oDataSource.SetValue("U_Costo", j, CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1))) 'CDec(0))
                Dim ConsumoMatrix As Integer = CInt(DatosEnExcel(j).LecturaFinal) - CInt(DatosEnExcel(j).LecturaInicial)
                Dim tot As Decimal = 0
                If cbxTipSer.Value <> "Gas" Then
                    tot = CDec(ConsumoMatrix) * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) 'Costo
                Else
                    tot = CDec(CDec(txtFact.Value) * CDec(CInt(DatosEnExcel(j).LecturaFinal) - CInt(DatosEnExcel(j).LecturaInicial))) * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) ' Costo
                End If

                oDataSource.SetValue("U_Total", j, tot)
            Next

            oMatrix.LoadFromDataSource()

            ValidaConsumo()

            IngresaResumenXNivel()

            rsboApp.StatusBar.SetText("Datos cargados desde el archivo Excel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oForm.Freeze(False)

        Catch ex As Exception
            oForm.Freeze(False)
            rsboApp.StatusBar.SetText("Error al Leer archivo .cvs, " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function ValidaConsumo()
        Try
            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
            Dim lblCons As SAPbouiCOM.StaticText = oForm.Items.Item("lblCons").Specific
            Dim lblTotal As SAPbouiCOM.StaticText = oForm.Items.Item("lblTotal").Specific
            Dim acuCon As Integer = 0
            Dim acuTot As Decimal = 0

            For i As Integer = 1 To mMatrix.RowCount
                Dim LecIni As Integer = CInt(mMatrix.Columns.Item("U_LecIni").Cells.Item(i).Specific.Value)
                Dim LecFin As Integer = CInt(mMatrix.Columns.Item("U_LecFin").Cells.Item(i).Specific.Value)
                Dim dif As Integer = LecFin - LecIni
                If dif < 0 Then
                    mMatrix.CommonSetting.SetCellBackColor(i, 8, ColorTranslator.ToOle(Color.Red))
                ElseIf LecIni > LecFin Then
                    mMatrix.CommonSetting.SetCellBackColor(i, 8, ColorTranslator.ToOle(Color.Red))
                ElseIf dif > 0 Then
                    mMatrix.CommonSetting.SetCellBackColor(i, 8, ColorTranslator.ToOle(Color.LightGreen))
                End If

                acuCon += LecFin - LecIni
                acuTot += CDec(mMatrix.Columns.Item("U_Tot").Cells.Item(i).Specific.Value)
            Next

            lblCons.Caption = acuCon.ToString
            lblTotal.Caption = acuTot.ToString

        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error validando consumos del excel: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    'Public Function CalculoDeTotales()
    '    Try
    '        Dim Costo As Decimal = 0
    '        Dim TotalConAjuste As Decimal = 0
    '        Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
    '        Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific
    '        Dim txtFact As SAPbouiCOM.EditText = oForm.Items.Item("txtFact").Specific
    '        Dim lblCons As SAPbouiCOM.StaticText = oForm.Items.Item("lblCons").Specific
    '        Dim lblTotal As SAPbouiCOM.StaticText = oForm.Items.Item("lblTotal").Specific

    '        Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
    '        Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET1")

    '        If mMatrix.RowCount > 0 Then
    '            TotalConAjuste = oGrid.DataTable.GetValue("TotalConAjuste", oGrid.DataTable.Rows.Count - 1)

    '            If cbxTipSer.Value <> "Gas" Then
    '                Costo = Math.Round((TotalConAjuste / CDec(lblCons.Caption)), 6)
    '            Else
    '                Costo = Math.Round((TotalConAjuste / (CDec(lblCons.Caption) * CDbl(txtFact.Value))), 6)
    '            End If

    '            For i As Integer = 0 To oDBDataSource.Size - 1
    '                oDBDataSource.SetValue("U_Costo", i, Costo)

    '                Dim Consumo As Integer = CInt(oDBDataSource.GetValue("U_Consumo", i))
    '                Dim tot As Decimal = 0

    '                If cbxTipSer.Value <> "Gas" Then
    '                    tot = CDec(Consumo) * Costo
    '                Else
    '                    tot = CDec(CDec(txtFact.Value) * CDec(Consumo)) * Costo
    '                End If
    '                oDBDataSource.SetValue("U_Total", i, tot)
    '            Next

    '            lblTotal.Caption = CStr(Costo * CDec(lblCons.Caption))

    '            mMatrix.LoadFromDataSource()

    '            IngresaResumenXNivel(Costo)
    '            'IngresaResumenFactura(Costo)
    '        End If
    '    Catch ex As Exception
    '        rsboApp.StatusBar.SetText("Error calculo de costos y totales: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Function

    Public Function IngresaResumenXNivel()
        Try
            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET1")
            Dim Acumulador As Integer = 0

            Dim infoPorNivel As New List(Of Tuple(Of String, Integer))

            For i As Integer = 0 To oDBDataSource.Size - 1
                infoPorNivel.Add(Tuple.Create(oDBDataSource.GetValue("U_Nivel", i), CInt(oDBDataSource.GetValue("U_Consumo", i))))
            Next

            Dim Agrupamiento = infoPorNivel.GroupBy(Function(x) x.Item1).Select(Function(g) New With {
            .Item = g.Key, .Suma = g.Sum(Function(x) x.Item2)}).ToList()

            Dim oMatrix2 As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_RES2").Specific
            oMatrix2.LoadFromDataSource()
            Dim k As Integer = 1
            For Each item As Object In Agrupamiento
                oMatrix2.AddRow()
                oMatrix2.Columns.Item("U_Niv").Cells.Item(k).Specific.String = item.Item.ToString
                oMatrix2.Columns.Item("U_Con").Cells.Item(k).Specific.String = CDec(item.Suma)
                oMatrix2.Columns.Item("U_Cos").Cells.Item(k).Specific.String = CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) 'Costo
                oMatrix2.Columns.Item("U_Tot").Cells.Item(k).Specific.String = CDec(item.Suma) * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) 'Costo
                Acumulador += CInt(item.Suma)
                k += 1
            Next

            oMatrix2.AddRow()

            oMatrix2.Columns.Item("U_Niv").Cells.Item(k).Specific.String = "Total Facturar locales"
            oMatrix2.Columns.Item("U_Con").Cells.Item(k).Specific.String = Acumulador
            oMatrix2.Columns.Item("U_Cos").Cells.Item(k).Specific.String = CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) 'Costo
            oMatrix2.Columns.Item("U_Tot").Cells.Item(k).Specific.String = Acumulador * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1)) 'Costo
            oMatrix2.CommonSetting.SetRowBackColor(k, RGB(255, 128, 0))

            For p As Integer = 1 To oMatrix2.RowCount - 1
                Dim porcentaje As Decimal = (CDec(oMatrix2.Columns.Item("U_Tot").Cells.Item(p).Specific.String) * 100) / (Acumulador * CDec(oGrid.DataTable.GetValue("Costo", oGrid.DataTable.Rows.Count - 1))) 'Costo)
                oMatrix2.Columns.Item("U_Por").Cells.Item(p).Specific.String = Math.Round(porcentaje, 2)
            Next
            Dim sumatoria As Decimal = 0
            For p As Integer = 1 To oMatrix2.RowCount - 1
                sumatoria += CDec(oMatrix2.Columns.Item("U_Por").Cells.Item(p).Specific.String)
            Next
            oMatrix2.Columns.Item("U_Por").Cells.Item(k).Specific.String = CDec(sumatoria)

        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error IngresaResumenXNivel: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function IngresaResumenFactura()
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            Dim oMatrix3 As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_RES1").Specific

            Dim docEntryGrid As String
            Dim totalConAjuste As Double
            Dim costoValue As Double = 0
            Dim Consumo As Decimal = 0

            Dim oDataSource2 As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET2")
            oDataSource2.Clear()
            oMatrix3.Clear()

            For o As Integer = 0 To oGrid.Rows.Count - 1

                'docEntryGrid = Convert.ToString(oGrid.DataTable.GetValue("DocEntry", o))
                'totalConAjuste = CDbl(oGrid.DataTable.GetValue("TotalConAjuste", o))
                'costoValue = CDbl(oGrid.DataTable.GetValue("Costo", o)) 'oGrid.DataTable.Rows.Count - 1))
                'Consumo = CDbl(oGrid.DataTable.GetValue("Consumo", o))

                Dim i As Integer = 0
                If Convert.ToString(oGrid.DataTable.GetValue("Tipo", o)) = "Y" Then

                    docEntryGrid = Convert.ToString(oGrid.DataTable.GetValue("DocEntry", o))
                    totalConAjuste = CDbl(oGrid.DataTable.GetValue("TotalConAjuste", o))
                    costoValue = CDbl(oGrid.DataTable.GetValue("Costo", o)) 'oGrid.DataTable.Rows.Count - 1))
                    Consumo = CDbl(oGrid.DataTable.GetValue("Consumo", o))

                    oDataSource2.InsertRecord(i)
                    oDataSource2.SetValue("U_DocEntryFac", i, docEntryGrid)
                    oDataSource2.SetValue("U_Valor", i, totalConAjuste)
                    oDataSource2.SetValue("U_Consumo", i, Consumo)
                    oDataSource2.SetValue("U_Costo", i, costoValue)
                    i += 1
                End If
            Next

            oMatrix3.LoadFromDataSource()

        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error IngresaResumenFactura: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub BuscarRegistro(ByVal DocEntryUdo As String)
        Try
            oForm.Freeze(True)

            Dim oCompanyService As SAPbobsCOM.CompanyService = rCompany.GetCompanyService()
            Dim oGeneralService As SAPbobsCOM.GeneralService = oCompanyService.GetGeneralService("SSSB")
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUdo)

            Dim oGeneralData As SAPbobsCOM.GeneralData = oGeneralService.GetByParams(oGeneralParams)

            Dim FechaInicio As String = oGeneralData.GetProperty("U_IniPer")
            Dim fechaIConvertida As Date
            If Date.TryParse(FechaInicio, fechaIConvertida) Then
                Dim txtPI As SAPbouiCOM.EditText = oForm.Items.Item("txtPI").Specific
                txtPI.Value = fechaIConvertida.ToString("yyyyMMdd")
            End If

            Dim FechaFin As String = oGeneralData.GetProperty("U_FinPer")
            Dim fechaFConvertida As Date
            If Date.TryParse(FechaFin, fechaFConvertida) Then
                Dim txtPF As SAPbouiCOM.EditText = oForm.Items.Item("txtPF").Specific
                txtPF.Value = fechaFConvertida.ToString("yyyyMMdd")
            End If

            Dim cbxCC3 As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCC3").Specific
            cbxCC3.Select(oGeneralData.GetProperty("U_NivCC3"), SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim cbxEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
            cbxEstado.Select(oGeneralData.GetProperty("U_Estado"), SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim cbxTipSer As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxTipSer").Specific
            cbxTipSer.Select(oGeneralData.GetProperty("U_TipSer"), SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim cbxUM As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxUM").Specific
            cbxUM.Select(oGeneralData.GetProperty("U_UM"), SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim txtFact As SAPbouiCOM.EditText = oForm.Items.Item("txtFact").Specific
            txtFact.Value = oGeneralData.GetProperty("U_Factor")

            Dim txtRuta As SAPbouiCOM.EditText = oForm.Items.Item("txtRuta").Specific
            txtRuta.Value = oGeneralData.GetProperty("U_Ruta")

            Dim txtCon As SAPbouiCOM.EditText = oForm.Items.Item("txtCon").Specific
            txtCon.Value = oGeneralData.GetProperty("U_Concepto")

            Dim MTX_RES1 As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_RES1").Specific
            Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET2") ' "UDO_D1" es un ejemplo del nombre de la tabla hija

            Dim oConditions As SAPbouiCOM.Conditions = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            Dim oCondition As SAPbouiCOM.Condition = oConditions.Add()
            oCondition.Alias = "DocEntry"  ' Campo en la tabla que quieres filtrar
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL  ' Operación de la condición (igual, mayor, menor, etc.)
            oCondition.CondVal = DocEntryUdo.ToString
            oDBDataSource.Query(oConditions)
            MTX_RES1.LoadFromDataSource()

            Dim MTX_RES2 As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_RES2").Specific
            Dim oDBDataSource2 As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET3") ' "UDO_D1" es un ejemplo del nombre de la tabla hija
            oDBDataSource2.Query(oConditions)
            MTX_RES2.LoadFromDataSource()

            Dim MTX_SS As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SS").Specific
            Dim oDBDataSource3 As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_SB_DET1") ' "UDO_D1" es un ejemplo del nombre de la tabla hija
            oDBDataSource3.Query(oConditions)
            MTX_SS.LoadFromDataSource()


            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.DataTable.Clear()

            Dim query As String = ""
            query = $"Select 'Y' AS ""Tipo"", A.""DocEntry"", A.""DocNum"", A.""CardName"", A.""DocDate"", B.""U_Valor"", 0.0 AS ""Ajuste"", B.""U_Valor"", B.""U_Consumo"", B.""U_Costo"", A.""Comments"" AS ""Comentario"" FROM ""OPCH"" A INNER JOIN ""@SS_SB_DET2"" B ON A.""U_SS_ServicioBasico"" = B.""DocEntry"" WHERE A.""U_SS_ServicioBasico"" = '{DocEntryUdo.ToString}' AND A.""DocEntry"" = B.""U_DocEntryFac"" "

            Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + query, "frmServiciosBasicos")
            oGrid.DataTable.ExecuteQuery(query)
            Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + query, "frmServiciosBasicos")
            FormatoFacturas()

            btnGuardar = oForm.Items.Item("btnGuardar").Specific
            btnGuardar.Caption = "Ok"

            Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                query = $"SELECT ""U_Costo"", ""U_Total"" FROM ""@SS_SB_DET3"" WHERE ""DocEntry"" =  ' {DocEntryUdo.ToString} ' AND ""U_Nivel"" = 'Total Facturar locales'"
                rst.DoQuery(query)

                Dim lblCons As SAPbouiCOM.StaticText = oForm.Items.Item("lblCons").Specific
                Dim lblTotal As SAPbouiCOM.StaticText = oForm.Items.Item("lblTotal").Specific
                If rst.RecordCount >= 1 Then
                    While (rst.EoF = False)
                        lblCons.Caption = CStr(rst.Fields.Item("U_Costo").Value)
                        lblTotal.Caption = CStr(rst.Fields.Item("U_Total").Value)
                        rst.MoveNext()
                    End While

                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al consultar resumen: " & ex.Message & " - query:" & query, "frmServiciosBasicos")
            End Try

            oForm.Freeze(False)

        Catch ex As Exception
            oForm.Freeze(False)

        End Try
    End Sub

    Public Sub CargaFormatos(Reporte As String, DocEntry As Integer)
        Try
            Dim menu As Object

            If Reporte = "Preliminar" Then
                menu = oFuncionesB1.ObtenerUIDMenu("RptPre2", Functions.VariablesGlobales._SBMenuPadreRptPreInf) '"13056")
            Else
                menu = oFuncionesB1.ObtenerUIDMenu("RptFir", Functions.VariablesGlobales._SBMenuPadreRptPreInf) '"13056")
            End If

            Utilitario.Util_Log.Escribir_Log("Menu: " & menu.ToString(), "frmServiciosBasicos")

            If menu <> "" Then

                For Each f As SAPbouiCOM.Form In rsboApp.Forms
                    If f.TypeEx = "410000100" Then f.Close()
                Next

                rsboApp.ActivateMenuItem(menu)

                Dim forpara As SAPbouiCOM.Form = rsboApp.Forms.GetForm("410000100", 0)
                forpara.Select()

                TryCast(forpara.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = DocEntry.ToString
                TryCast(forpara.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                forpara.Visible = False

            End If
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error Cargando formatos " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class

Public Class DatosExcel
    Public Property Contrato As Integer
    Public Property ContratoAnterior As Integer
    Public Property Denominacion As String
    Public Property locales As String
    Public Property Nivel As String
    Public Property Medidor As String
    Public Property LecturaInicial As Integer
    Public Property LecturaFinal As Integer
End Class
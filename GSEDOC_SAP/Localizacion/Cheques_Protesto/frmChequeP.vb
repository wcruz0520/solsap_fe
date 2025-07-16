Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms

Public Class frmChequeP
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oUserDataSource As SAPbouiCOM.UserDataSource

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormulario_frmChequeP()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmChequeP") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmChequeP.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmChequeP").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmChequeP")

            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = Application.StartupPath & "\imagen_UPD.jpg"

            Dim txtFchIni As SAPbouiCOM.EditText
            txtFchIni = oForm.Items.Item("finicial").Specific
            txtFchIni.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtFchFin As SAPbouiCOM.EditText
            txtFchFin = oForm.Items.Item("ffinal").Specific
            txtFchFin.Value = DateTime.Now.ToString("yyyyMMdd")

            'Estados Documentos.
            Dim cmbEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
            cmbEstado.ValidValues.Add("TODOS", "TODOS")
            cmbEstado.ValidValues.Add("NO PROTESTADOS", "NO PROTESTADOS")
            cmbEstado.ValidValues.Add("PROTESTADOS", "PROTESTADOS")

            cmbEstado.Select("TODOS", SAPbouiCOM.BoSearchKey.psk_ByValue)

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
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            ' END CHOOSE FROM LIST

            Dim txtCliente As SAPbouiCOM.EditText
            txtCliente = oForm.Items.Item("txtCliente").Specific
            oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCliente.DataBind.SetBound(True, "", "EditDS")
            txtCliente.ChooseFromListUID = "CFL1"
            txtCliente.ChooseFromListAlias = "CardCode"


            Formulario_frmChequeP_CargarGrid()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Utilitario.Util_Log.Escribir_Log("ex CreaFormulario_frmChequeP  " + ex.Message.ToString(), "frmChequeP")
        End Try

    End Sub

    Public Sub Formulario_frmChequeP_CargarGrid()
        oForm.Freeze(True)

        Dim cbxEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
        Dim txtfinicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
        Dim txtffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific
        Dim txtCliente As SAPbouiCOM.EditText = oForm.Items.Item("txtCliente").Specific
        Dim txtCheque As SAPbouiCOM.EditText = oForm.Items.Item("txtCheque").Specific
        Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

        If (String.IsNullOrEmpty(txtfinicial.Value) Or String.IsNullOrEmpty(txtffinal.Value)) Then
            rsboApp.SetStatusBarMessage("Debe ingresar un rango de fechas a consultar!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        Dim sCheque As Integer = 0
        Try
            If txtCheque.Value.Trim() = "" Then
                sCheque = 0
            Else
                sCheque = CInt(txtCheque.Value.Trim())
            End If

        Catch ex As Exception
            rsboApp.SetStatusBarMessage("El valor ingresado en Num Cheque no es correcto!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End Try

        Dim dfechaDesde As Date
        Dim dfechaHasta As Date
        Dim sQuery As String = ""
        Dim sEstado As String = cbxEstado.Value.Trim()
        Dim sfolioIni As String = txtfinicial.Value.Trim()
        Dim sfoliofin As String = txtffinal.Value.Trim()
        Dim sCliente As String = txtCliente.Value.Trim()


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
            sQuery = "CALL SS_LOC_GET_CHEQUES ("

            sQuery += "'" + sCliente + "'"
            sQuery += "," + sCheque.ToString()
            sQuery += ",'" + sEstado + "'"
            'sQuery += ",''"
            'sQuery += ",''"        
            sQuery += "," + FechaSql(dfechaDesde)
            sQuery += "," + FechaSql(dfechaHasta) + ")"
        Else
            sQuery = "EXEC SS_LOC_GET_CHEQUES "

            sQuery += "'" + sCliente + "'"
            sQuery += "," + sCheque.ToString()
            sQuery += ",'" + sEstado + "'"
            'sQuery += ",''"
            'sQuery += ",''"        
            sQuery += "," + FechaSql(dfechaDesde)
            sQuery += "," + FechaSql(dfechaHasta)
        End If


        'ESTO ES PARA PROBAR QUERYS ENCRIPTADOS DESDE DB

        'sQuery = Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._DocsEnviadosQRY.Replace("{", "").Replace("}", "").ToString(), sKey)

        If String.IsNullOrWhiteSpace(sQuery) Then

            rsboApp.SetStatusBarMessage("(GS) Pantalla Inactiva , Por Favor Revisar la Parametrizacion de los Documentos Enviados", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

        End If

        Utilitario.Util_Log.Escribir_Log("Query Cheques:  " + sQuery.ToString(), "frmChequeP")
        'FIN QUERY CONSULTA ENCRYPTADA
        Try
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                oGrid.DataTable.ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Cheques Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmChequeP")
            End Try

            oGrid.Columns.Item(0).Description = "# de Pago"
            oGrid.Columns.Item(0).TitleObject.Caption = "# de Pago"
            oGrid.Columns.Item(0).Editable = False
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = 24

            oGrid.Columns.Item(1).Description = "Fecha Cheque"
            oGrid.Columns.Item(1).TitleObject.Caption = "Fecha Cheque"
            oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "# de Cheque"
            oGrid.Columns.Item(2).TitleObject.Caption = "# de Cheque"
            oGrid.Columns.Item(2).Editable = False

            oGrid.Columns.Item(3).Description = "Valor"
            oGrid.Columns.Item(3).TitleObject.Caption = "Valor"
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).RightJustified = True

            oGrid.Columns.Item(4).Description = "Banco"
            oGrid.Columns.Item(4).TitleObject.Caption = "Banco"
            oGrid.Columns.Item(4).Editable = False

            oGrid.Columns.Item(5).Description = "Cod Cliente"
            oGrid.Columns.Item(5).TitleObject.Caption = "Cod Cliente"
            oGrid.Columns.Item(5).Editable = False
            Dim oEditTextColumnC As SAPbouiCOM.EditTextColumn
            oEditTextColumnC = oGrid.Columns.Item(5)
            oEditTextColumnC.LinkedObjectType = 2

            oGrid.Columns.Item(6).Description = "Cliente"
            oGrid.Columns.Item(6).TitleObject.Caption = "Cliente"
            oGrid.Columns.Item(6).Editable = False

            oGrid.Columns.Item(7).Description = "Protesto"
            oGrid.Columns.Item(7).TitleObject.Caption = "Protesto"
            oGrid.Columns.Item(7).Editable = False
            Dim oEditTextColumnP As SAPbouiCOM.EditTextColumn
            oEditTextColumnP = oGrid.Columns.Item(7)
            oEditTextColumnP.LinkedObjectType = 13

            oGrid.Columns.Item(8).Width = 0
            oGrid.Columns.Item(8).Visible = False
            oGrid.Columns.Item(8).Editable = False

            oGrid.Columns.Item(9).Width = 0
            oGrid.Columns.Item(9).Visible = False
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

            'oGrid.CollapseLevel = 1
            'oGrid.AutoResizeColumns()
            'schk.Checked = False

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Utilitario.Util_Log.Escribir_Log("ex Formulario_frmChequeP_CargarGrid  " + ex.Message.ToString(), "frmChequeP")
        End Try

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmChequeP" Then


                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If Not pVal.Before_Action Then

                            Select Case pVal.ItemUID

                                Case "obtnBuscar"

                                    Formulario_frmChequeP_CargarGrid()

                                Case "obtnCerrar"

                                    oForm.Close()

                                Case "schk"

                                    ExpandirContraer()



                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                        If pVal.BeforeAction Then

                            Event_MatrixLinkPressed(pVal)

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.BeforeAction = False And pVal.ItemUID = "oGrid" Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmChequeP")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            Dim ofila As Integer = 0
                            ofila = pVal.Row
                            'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                            'oGrid.Rows.SelectedRows.Add(ofila)
                            'For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                            '    ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                            '    ' Dim sDocNum As String = odt.GetValue("Document Number", oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, BoOrderType.ot_RowOrder)))
                            'Next

                            If ofila = -1 Then '' clic en la cabecera
                                Exit Sub
                            End If
                            Dim oCheque As New clsCheque(oDataTable.GetValue("NumPago", ofila).ToString(),
                                                           oDataTable.GetValue("Cheque_Num", ofila).ToString(),
                                                           oDataTable.GetValue("Cheque_Valor", ofila),
                                                           oDataTable.GetValue("Banco", ofila).ToString(),
                                                           oDataTable.GetValue("Cliente_Codigo", ofila).ToString(),
                                                           oDataTable.GetValue("Cliente", ofila).ToString(),
                                                           oDataTable.GetValue("Doc_Protesto", ofila).ToString(),
                                                           oDataTable.GetValue("Pago_Coments", ofila).ToString(),
                                                           oDataTable.GetValue("CuentaContableDeposito", ofila).ToString(),
                                                           oDataTable.GetValue("NombreCuentaContableDeposito", ofila).ToString(),
                                                           oDataTable.GetValue("NumeroDeposito", ofila).ToString())

                            Dim ValorProtesto As Decimal = 0
                            If oCheque.Doc_Protesto <> "0" Then
                                Dim sQuery As String
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    sQuery = "SELECT TOP 1 B.""LineTotal"" FROM ""OINV"" A INNER JOIN ""INV1"" B ON A.""DocEntry"" = B.""DocEntry"" WHERE B.""Dscription"" = 'MONTO PROTESTO' AND A.""DocEntry"" =  '" + oCheque.Doc_Protesto + "'"
                                Else
                                    sQuery = "SELECT TOP 1 B.""LineTotal"" FROM ""OINV"" A INNER JOIN ""INV1"" B ON A.""DocEntry"" = B.""DocEntry"" WHERE B.""Dscription"" = 'MONTO PROTESTO' AND A.""DocEntry"" =  '" + oCheque.Doc_Protesto + "'"
                                End If
                                Utilitario.Util_Log.Escribir_Log("Query Obtener Valor del protesto:  " + sQuery, "frmChequeP")

                                ValorProtesto = formatDecimal(oFuncionesB1.getRSvalue(sQuery, "LineTotal", ""))

                            End If
                            ofrmChequePD.CreaFormulario_frmChequePD(oCheque, ValorProtesto, ofila)

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
                                        Case "txtCliente"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("CardCode", 0)
                                            Dim lbCliente As SAPbouiCOM.StaticText
                                            lbCliente = oForm.Items.Item("lbCliente").Specific
                                            lbCliente.Caption = oDataTable.GetValue("CardName", 0)



                                    End Select
                                End If
                            Catch ex As Exception
                                rsboApp.MessageBox("et_CHOOSE_FROM_LIST " + ex.Message.ToString())
                                Utilitario.Util_Log.Escribir_Log("ex et_CHOOSE_FROM_LIST  " + ex.Message.ToString(), "frmChequeP")
                            End Try
                        End If

                End Select
            End If
        Catch ex As Exception

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

        If pVal.FormTypeEx = "frmChequeP" Then

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


    Public Shared Function FechaSql(ByVal fecha As DateTime) As String

        Dim anio As String = fecha.Year
        Dim mes As String = fecha.Month
        Dim dia As String = fecha.Day

        If anio.Length = 2 Then
            anio = "20" & anio
        End If

        Return "{d'" & anio & "-" & mes.PadLeft(2, "0") & "-" & dia.PadLeft(2, "0") & "'}"

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

End Class

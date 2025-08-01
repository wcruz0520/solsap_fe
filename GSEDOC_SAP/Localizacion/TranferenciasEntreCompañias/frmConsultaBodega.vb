Imports System.Globalization
Imports System.Threading
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class frmConsultaBodega

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private rCompany2 As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable


    Dim _IdBodega As SAPbouiCOM.EditText = Nothing
    Dim _NombreBod As SAPbouiCOM.StaticText = Nothing


    Public WithEvents _oGrid As SAPbouiCOM.Grid


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    'Public Sub CargaFormularioConsulta(ofila As Integer, TipoSerie As SAPbouiCOM.ComboBox, IdSerie As SAPbouiCOM.EditText, NombreSerie As SAPbouiCOM.EditText)
    Public Sub CargaFormularioConsulta(IdBodega As SAPbouiCOM.EditText, NombreBod As SAPbouiCOM.StaticText, company2 As SAPbobsCOM.Company)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        Dim i As Integer = 0
        If RecorreFormulario(rsboApp, "frmConsultaBodega") Then
            Exit Sub
        End If
        rCompany2 = company2
        _IdBodega = IdBodega
        _NombreBod = NombreBod


        strPath = System.Windows.Forms.Application.StartupPath & "\frmConsultaBodega.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConsultaBodega").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Lista de Bodegas, por favor espere!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm = rsboApp.Forms.Item("frmConsultaBodega")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"

            Try
                oForm.DataSources.DataTables.Add("dtDocsDT")
            Catch ex As Exception
            End Try

            Dim sQuery As String = ""

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "select ""WhsCode"",""WhsName"" from OWHS"

            Else
                sQuery = "select WhsCode,WhsName from OWHS"

            End If

            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oForm.DataSources.DataTables.Item("dtDocsDT").Clear()
            oForm.DataSources.DataTables.Item("dtDocsDT").Columns.Add("IdBodega", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocsDT").Columns.Add("NombreBodega", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)


            Dim rs As SAPbobsCOM.Recordset = rCompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(sQuery)
            oForm.DataSources.DataTables.Item("dtDocsDT").Rows.Clear()
            Dim Cant = rs.RecordCount
            While (rs.EoF = False)

                oForm.DataSources.DataTables.Item("dtDocsDT").Rows.Add(1)
                oForm.DataSources.DataTables.Item("dtDocsDT").SetValue("IdBodega", i, rs.Fields.Item("WhsCode").Value)
                oForm.DataSources.DataTables.Item("dtDocsDT").SetValue("NombreBodega", i, rs.Fields.Item("WhsName").Value)
                i += 1
                rs.MoveNext()
            End While

            CargaDocumentosFormato()

            oForm.Visible = True
            oForm.Select()

            'AddHandler oGrid.ClickAfter, AddressOf Grid_ClickAfter

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + "Ocurrio un error al cargar el formulario :" + ex.Message().ToString(), 1, "")
        End Try

    End Sub
    Private Sub CargaDocumentosFormato()
        Try
            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocsDT")
            'oGrid.DataTable.ExecuteQuery(sQuery)

            oGrid.Columns.Item(0).Description = "Código"
            oGrid.Columns.Item(0).TitleObject.Caption = "Código"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "Nombre"
            oGrid.Columns.Item(1).TitleObject.Caption = "Nombre"
            oGrid.Columns.Item(1).Editable = False


            oForm.Freeze(False)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FormularioTCargarGrid()
        oForm.Freeze(True)

        Dim sQuery As String = ""

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            sQuery = "select ""ItemCode"",""ItemName"",""AvgPrice"" as ""Costo del Articulo"" from OITM"
            'sQuery += " UNION ALL SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion FROM NNM1 WHERE ObjectCode = '13' AND DocSubType = 'IX' AND Locked = 'N' AND IsForCncl = 'N' "
        Else
            sQuery = " select ItemCode,ItemName,AvgPrice as ""Costo del Articulo"" from OITM"
            'sQuery += " UNION ALL SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion FROM ""NNM1"" WHERE ""ObjectCode"" = '13' AND ""DocSubType"" = 'IX' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
        End If
        Try
            Try
                oForm.DataSources.DataTables.Add("dtDocsDT")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocsDT")
            Try
                Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + sQuery, "frmConsultaBodega")
                oGrid.DataTable.ExecuteQuery(sQuery)
                Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + sQuery, "frmConsultaBodega")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmDocumentosEnviados")
            End Try

            oGrid.Columns.Item(0).Description = "Tipo Documento"
            oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "#"
            oGrid.Columns.Item(1).TitleObject.Caption = "#"
            oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "DocEntry"
            oGrid.Columns.Item(2).TitleObject.Caption = "DocEntry"
            oGrid.Columns.Item(2).Editable = False

            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos:" + sQuery + " - " + ex.Message.ToString, "frmDocumentosEnviados")
        Finally
            oForm.Freeze(False)
        End Try

    End Sub


    Private Sub ConsultaProductosCodigo(ByVal Cadena As String)
        oForm.Freeze(True)

        Try
            Dim sQuery As String = ""

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "select ""ItemCode"",""ItemName"",""AvgPrice"" AS ""Costo del Articulo"",""OnHand"" as ""Stock"" from OITM where ""ItemCode"" Like '%" + Cadena.ToString + "%'"
                'sQuery += " UNION ALL SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion FROM NNM1 WHERE ObjectCode = '13' AND DocSubType = 'IX' AND Locked = 'N' AND IsForCncl = 'N' "
            Else
                sQuery = " select ItemCode,ItemName,AvgPrice AS ""Costo del Articulo"",""OnHand"" as ""Stock"" from OITM where ItemCode Like '%" + Cadena.ToString + "%'"
                'sQuery += " UNION ALL SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion FROM ""NNM1"" WHERE ""ObjectCode"" = '13' AND ""DocSubType"" = 'IX' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocsDT").ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Consulta productos" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmConsultaBodega")
            End Try

            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocsDT")
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(2).Visible = True
            oGrid.Columns.Item(2).Editable = False
            'oGrid.Columns.Item(3).Visible = True
            'oGrid.Columns.Item(3).Editable = False

            oGrid.AutoResizeColumns()


            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            'Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos:" + sQuery + " - " + ex.Message.ToString, "frmDocumentosEnviados")
        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub ConsultaProductosNombre(ByVal Cadena As String)
        oForm.Freeze(True)

        Try
            Dim sQuery As String = ""

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "select ""ItemCode"",""ItemName"",""AvgPrice"" AS ""Costo del Articulo"",""OnHand"" as ""Stock"" from OITM where ""ItemName"" Like '%" + Cadena.ToString + "%'"
                'sQuery += " UNION ALL SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion FROM NNM1 WHERE ObjectCode = '13' AND DocSubType = 'IX' AND Locked = 'N' AND IsForCncl = 'N' "
            Else
                sQuery = " select ItemCode,ItemName,AvgPrice AS ""Costo del Articulo"",""OnHand"" as ""Stock"" from OITM where ItemName Like '%" + Cadena.ToString + "%'"
                'sQuery += " UNION ALL SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion FROM ""NNM1"" WHERE ""ObjectCode"" = '13' AND ""DocSubType"" = 'IX' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocsDT").ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Consulta productos" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmConsultaBodega")
            End Try

            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocsDT")
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(2).Visible = True
            oGrid.Columns.Item(2).Editable = False
            'oGrid.Columns.Item(3).Visible = True
            'oGrid.Columns.Item(3).Editable = False

            oGrid.AutoResizeColumns()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            'Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos:" + sQuery + " - " + ex.Message.ToString, "frmDocumentosEnviados")
        Finally
            oForm.Freeze(False)
        End Try

    End Sub


    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try

            If FormUID = "frmConsultaBodega" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK

                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID

                                Case "btnAsig"
                                    oForm = rsboApp.Forms.Item("frmConsultaBodega")
                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    Dim oDT As SAPbouiCOM.DataTable = oGrid.DataTable
                                    Dim sIdBodega As String = ""
                                    Dim sNombreBodega As String = ""


                                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                        If oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString() <> "" Then
                                            sIdBodega = oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                            sNombreBodega = oDT.GetValue(1, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                            'sCosto = oDT.GetValue(2, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                        End If
                                    Next

                                    _IdBodega.Value = sIdBodega
                                    _NombreBod.Caption = sNombreBodega


                                    oForm = rsboApp.Forms.Item("frmConsultaBodega")
                                    oForm.Close()

                                Case "btnBuscar"
                                    'Dim columnName As String = ""
                                    'Dim cadena As SAPbouiCOM.EditText = oForm.Items.Item("txtCadena").Specific

                                    'Dim titulo = pVal.ColUID
                                    'Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    'Select Case colBuscar
                                    '    Case "ItemCode"
                                    '        ConsultaProductosCodigo(cadena.Value.ToString)
                                    '    Case "ItemName"
                                    '        ConsultaProductosNombre(cadena.Value.ToString)
                                    '    Case Else
                                    '        ConsultaProductosCodigo(cadena.Value.ToString)
                                    'End Select

                            End Select

                        End If

                        'Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        '    If Not pVal.Before_Action Then
                        '        If pVal.CharPressed = 9 Then

                        '            oForm = rsboApp.Forms.Item("frmConsultaBodega")
                        '            Dim columnName As String = ""
                        '            Dim cadena As SAPbouiCOM.EditText = oForm.Items.Item("txtCadena").Specific

                        '            Dim titulo = pVal.ColUID
                        '            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                        '            Select Case colBuscar
                        '                Case "ItemCode"
                        '                    ConsultaProductosCodigo(cadena.Value.ToString)
                        '                Case "ItemName"
                        '                    ConsultaProductosNombre(cadena.Value.ToString)
                        '                Case Else
                        '                    ConsultaProductosCodigo(cadena.Value.ToString)
                        '            End Select


                        '        End If

                        '        'Dim txtFocus As SAPbouiCOM.EditText
                        '        'txtFocus = oForm.Items.Item("txtFocus").Specific
                        '        'txtFocus.Item.Click()


                        '        'txtFocus = oForm.Items.Item("txtCadena").Specific
                        '        'txtFocus.Item.Click()
                        '        'txtFocus.Item.Click()
                        '        'oMatrix.Columns.Item("0").Cells.Item(1).Click()
                        '    End If



                End Select
            End If

        Catch ex As Exception

        Finally

        End Try
    End Sub

    Public Sub Grid_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles _oGrid.ClickAfter
        Dim columnIndex As String = pVal.ColUID ' Índice de la columna

        ' Realizar las acciones necesarias según el título de la columna seleccionada
        'colBuscar = columnIndex
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
                    oForm.Select()
                    oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
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


End Class

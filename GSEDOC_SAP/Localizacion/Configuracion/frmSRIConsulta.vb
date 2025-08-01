Imports System.Globalization
Imports System.Threading

Public Class frmSRIConsulta

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable

    Dim _fila As Integer
    Dim _TipoSerie As SAPbouiCOM.ComboBox = Nothing
    Dim _IdSerie As SAPbouiCOM.EditText = Nothing
    Dim _NombreSerie As SAPbouiCOM.EditText = Nothing
    Dim _DocNumInicial As SAPbouiCOM.EditText = Nothing

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioConsulta(ofila As Integer, TipoSerie As SAPbouiCOM.ComboBox, IdSerie As SAPbouiCOM.EditText, NombreSerie As SAPbouiCOM.EditText)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmSRIConsulta") Then
            Exit Sub
        End If

        _fila = ofila
        _TipoSerie = TipoSerie
        _IdSerie = IdSerie
        _NombreSerie = NombreSerie
        '_DocNumInicial = DocNumInicial

        strPath = System.Windows.Forms.Application.StartupPath & "\frmSRIConsulta.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmSRIConsulta").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            rsboApp.StatusBar.SetText(NombreAddon + " - Cargando Formularios de Consultas de Series, por favor espere!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm = rsboApp.Forms.Item("frmSRIConsulta")
            oForm.Freeze(True)


            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim sQuery As String = ""
            If _TipoSerie.Value = "FV" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion FROM NNM1 WHERE ObjectCode = '13' AND DocSubType = '--' AND Locked = 'N' AND IsForCncl = 'N' "
                    sQuery += " UNION ALL SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion FROM NNM1 WHERE ObjectCode = '13' AND DocSubType = 'IX' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion FROM ""NNM1"" WHERE ""ObjectCode"" = '13' AND ""DocSubType"" = '--' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                    sQuery += " UNION ALL SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion FROM ""NNM1"" WHERE ""ObjectCode"" = '13' AND ""DocSubType"" = 'IX' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If
            ElseIf _TipoSerie.Value = "NC" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion  FROM NNM1 WHERE ObjectCode = '14' AND DocSubType = '--' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion  FROM ""NNM1"" WHERE ""ObjectCode"" = '14' AND ""DocSubType"" = '--' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If
            ElseIf _TipoSerie.Value = "ND" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion  FROM NNM1 WHERE ObjectCode = '13' AND DocSubType = 'DN' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion  FROM ""NNM1"" WHERE ""ObjectCode"" = '13' AND ""DocSubType"" = 'DN' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If

            ElseIf _TipoSerie.Value = "GR" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion  FROM NNM1 WHERE ObjectCode = '15' AND DocSubType = '--' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion  FROM ""NNM1"" WHERE ""ObjectCode"" = '15' AND ""DocSubType"" = '--' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If



            ElseIf _TipoSerie.Value = "GRT" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion  FROM NNM1 WHERE ObjectCode = '67' AND DocSubType = '--' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion  FROM ""NNM1"" WHERE ""ObjectCode"" = '67' AND ""DocSubType"" = '--' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If

            ElseIf _TipoSerie.Value = "GRST" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion  FROM NNM1 WHERE ObjectCode = '1250000001' AND DocSubType = '--' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion  FROM ""NNM1"" WHERE ""ObjectCode"" = '1250000001' AND ""DocSubType"" = '--' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If

                'add Documentos Soportes Arturo - 20-04-2022

            ElseIf _TipoSerie.Value = "RT" Or _TipoSerie.Value = "LQ" Or _TipoSerie.Value = "LQRT" Then
                If Not rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = " SELECT Series AS Codigo,SeriesName AS Nombre ,Remark AS Observacion  FROM NNM1 WHERE ObjectCode = '18' AND DocSubType = '--' AND Locked = 'N' AND IsForCncl = 'N' "
                Else
                    sQuery = " SELECT ""Series"" AS Codigo,""SeriesName"" AS Nombre ,""Remark"" AS Observacion  FROM ""NNM1"" WHERE ""ObjectCode"" = '18' AND ""DocSubType"" = '--' AND ""Locked"" = 'N' AND ""IsForCncl"" = 'N' "
                End If



            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocs").ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query SRI Consulta Series" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmSRICONSULTA")
            End Try

            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(2).Visible = True
            oGrid.Columns.Item(2).Editable = False
            'oGrid.Columns.Item(3).Visible = True
            'oGrid.Columns.Item(3).Editable = False

            oGrid.AutoResizeColumns()

            oForm.Width = 497
            oForm.Height = 290

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + "Ocurrio un error al cargar la consulta :" + ex.Message().ToString(), 1, "")
        End Try

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try

            If FormUID = "frmSRIConsulta" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK

                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID

                                Case "btnAsig"
                                    oForm = rsboApp.Forms.Item("frmSRIConsulta")
                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                    Dim oDT As SAPbouiCOM.DataTable = oGrid.DataTable
                                    Dim sIdSerie As String = ""
                                    Dim sSerieName As String = ""
                                    'Dim sNextNumber As String = ""

                                    For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                        If oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString() <> "" Then
                                            sIdSerie = oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                            sSerieName = oDT.GetValue(1, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                            'sNextNumber = oDT.GetValue(3, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                        End If
                                    Next

                                    _IdSerie.Value = sIdSerie
                                    _NombreSerie.Value = sSerieName
                                    '_DocNumInicial.Value = sNextNumber

                                    'rsboApp.Forms.Item("frmHojaCosto").Freeze(True)
                                    'Dim dtT As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmHojaCosto").DataSources.DataTables.Item("dtIngresos")
                                    'Try

                                    '    If _sObjType = "1250000025" Then
                                    '        dtT.SetValue(2, _fila, oCode.ToString())
                                    '    ElseIf _sObjType = "4" Then
                                    '        dtT.SetValue(3, _fila, oCode.ToString())
                                    '        dtT.SetValue(4, _fila, oCantidadContenedores.ToString())
                                    '        dtT.SetValue(5, _fila, oCantidad.ToString())
                                    '        dtT.SetValue(7, _fila, oPrecio.ToString())
                                    '        dtT.SetValue(8, _fila, Math.Round((ConvertToDouble(oCantidad.ToString()) * ConvertToDouble(oPrecio.ToString())), 2))
                                    '    End If
                                    'Catch ex As Exception
                                    'Finally

                                    'End Try
                                    ' rsboApp.Forms.Item("frmSRI").Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oForm.Close()

                            End Select

                        End If

                End Select
            End If

        Catch ex As Exception

        Finally

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

Imports System.IO
Imports System.Xml.Serialization

Public Class frmConsultaSalidaEntrada

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private odt As SAPbouiCOM.DataTable


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub


    Public Sub Carga_Salidas_Entrada()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConsultaSalidaEntrada") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmConsultaSalidaEntrada.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConsultaSalidaEntrada").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmConsultaSalidaEntrada")

            oForm.Freeze(True)

            Dim txtFchIni As SAPbouiCOM.EditText
            txtFchIni = oForm.Items.Item("finicial").Specific
            txtFchIni.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtFchFin As SAPbouiCOM.EditText
            txtFchFin = oForm.Items.Item("ffinal").Specific
            txtFchFin.Value = DateTime.Now.ToString("yyyyMMdd")

            CargarGrid()

            ' CargaDatos()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargarGrid()
        oForm.Freeze(True)


        Dim txtfinicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
        Dim txtffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific
        'Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

        If (String.IsNullOrEmpty(txtfinicial.Value) Or String.IsNullOrEmpty(txtffinal.Value)) Then
            rsboApp.SetStatusBarMessage("Debe ingresar un rango de fechas a consultar!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        Dim dfechaDesde As Date
        Dim dfechaHasta As Date
        Dim sQuery As String = ""

        Dim sfolioIni As String = txtfinicial.Value.Trim()
        Dim sfoliofin As String = txtffinal.Value.Trim()

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

        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    'CALL GS_SAP_FE_ONE_OBTENERDOCUMENTOS ('0','2',{d'2016-06-16'},{d'2017-09-28'})
        '    sQuery = "CALL GS_CO_SAP_FE_ObtenerDocumentos ("

        '    sQuery += "'" + sTipoDoc + "'"
        '    sQuery += ",'" + sEstado + "'"
        '    'sQuery += ",''"
        '    'sQuery += ",''"        
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta) + ")"
        'Else
        '    sQuery = "EXEC GS_CO_SAP_FE_ObtenerDocumentos "

        '    sQuery += "'" + sTipoDoc + "'"
        '    sQuery += ",'" + sEstado + "'"
        '    'sQuery += ",''"
        '    'sQuery += ",''"        
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaDesde)
        '    sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta)
        'End If


        'ESTO ES PARA PROBAR QUERYS ENCRIPTADOS DESDE DB

        'sQuery = Utilitario.Util_Encriptador.Desencriptar(functions.VariablesGlobales._SS_ComprasQRY.Replace("{", "").Replace("}", "").ToString(), sKey)

        'If String.IsNullOrWhiteSpace(sQuery) Then

        '    rsboApp.SetStatusBarMessage("(GS) Pantalla Inactiva , Por Favor Revisar la Parametrizacion de los Documentos Enviados", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

        'End If


        'sQuery = sQuery.Replace("REPPLACE_TIPODOC", "'" + sTipoDoc + "'")
        'sQuery = sQuery.Replace("REPPLACE_ESTADO", "'" + sEstado + "'")


        'sQuery = sQuery.Replace("@f1", functions.FuncionesB1.FechaSql(dfechaDesde))
        'sQuery = sQuery.Replace("@f2", functions.FuncionesB1.FechaSql(dfechaHasta))
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            sQuery = "CALL SS_CONSULTA_SALIDA_ENTRADA ("
            sQuery += Functions.FuncionesB1.FechaSql(dfechaDesde)
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta) + ")"
        Else
            sQuery = "EXEC SS_CONSULTA_SALIDA_ENTRADA "
            sQuery += Functions.FuncionesB1.FechaSql(dfechaDesde)
            sQuery += "," + Functions.FuncionesB1.FechaSql(dfechaHasta)
        End If

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
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmConsultaSalidaEntrada")
            End Try


            If oGrid.DataTable.Rows.Count > 0 Then

                For y As Integer = 0 To oGrid.Columns.Count - 1


                    oGrid.Columns.Item(y).Editable = False


                Next

            End If

            'oGrid.Columns.Item(0).Description = "Tipo Documento"
            'oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            'oGrid.Columns.Item(0).Editable = False

            'oGrid.Columns.Item(1).Description = "#"
            'oGrid.Columns.Item(1).TitleObject.Caption = "#"
            'oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(0).Description = "Id Salida Mercancias"
            oGrid.Columns.Item(0).TitleObject.Caption = "Id Salida Mercancias"


            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = 60

            oGrid.Columns.Item(1).Description = "Id Entrada Mercancias"
            oGrid.Columns.Item(1).TitleObject.Caption = "Id Entrada Mercancias"

            'se comento debido a que pertenece a otra base
            'Dim oEditTextColumn1 As SAPbouiCOM.EditTextColumn
            'oEditTextColumn1 = oGrid.Columns.Item(1)
            'oEditTextColumn1.LinkedObjectType = 59

            'oGrid.Columns.Item(3).Description = "Fecha Emisión"
            'oGrid.Columns.Item(3).TitleObject.Caption = "Fecha Emisión"
            'oGrid.Columns.Item(3).Editable = False

            'oGrid.Columns.Item(4).Description = "Doc. Num."
            'oGrid.Columns.Item(4).TitleObject.Caption = "Doc. Num."
            'oGrid.Columns.Item(4).Editable = False

            'oGrid.Columns.Item(5).Description = "Cliente"
            'oGrid.Columns.Item(5).TitleObject.Caption = "Cliente"
            'oGrid.Columns.Item(5).Editable = False


            'oGrid.Columns.Item(6).Description = "Doc. Total"
            'oGrid.Columns.Item(6).TitleObject.Caption = "Doc. Total"
            'oGrid.Columns.Item(6).Editable = False
            'oGrid.Columns.Item(6).RightJustified = True


            'oGrid.Columns.Item(7).Description = "Estado Documento"
            'oGrid.Columns.Item(7).TitleObject.Caption = "Estado Documento"
            'oGrid.Columns.Item(7).Editable = False

            'oGrid.Columns.Item(8).Description = "CUF"
            'oGrid.Columns.Item(8).TitleObject.Caption = "CUF"
            'oGrid.Columns.Item(8).Editable = False

            'oGrid.Columns.Item(9).Description = "EXT1"
            'oGrid.Columns.Item(9).TitleObject.Caption = "EXT1"
            'oGrid.Columns.Item(9).Editable = False
            'oGrid.Columns.Item(9).Visible = False

            'oGrid.Columns.Item(10).Description = "EXT2"
            'oGrid.Columns.Item(10).TitleObject.Caption = "EXT2"
            'oGrid.Columns.Item(10).Editable = False
            'oGrid.Columns.Item(10).Visible = False

            'oGrid.Columns.Item(11).Description = "EXT3"
            'oGrid.Columns.Item(11).TitleObject.Caption = "EXT3"
            'oGrid.Columns.Item(11).Editable = False
            'oGrid.Columns.Item(11).Visible = False

            'oGrid.Columns.Item(12).Description = "EXT4"
            'oGrid.Columns.Item(12).TitleObject.Caption = "EXT4"
            'oGrid.Columns.Item(12).Editable = False
            'oGrid.Columns.Item(12).Visible = False


            'oGrid.CollapseLevel = 1
            'oGrid.AutoResizeColumns()
            'schk.Checked = False

            oForm.Freeze(False)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.FormTypeEx = "frmConsultaSalidaEntrada" Then

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                        Case "btnBuscar"
                            If pVal.BeforeAction = False Then
                                CargarGrid()
                            Else

                            End If

                        Case ""

                        Case Else

                    End Select


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.BeforeAction = False And pVal.ItemUID = "oGrid" Then
                        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmConsultaSalidaEntrada")
                        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                        Dim ofila As Integer = 0


                        ofila = pVal.Row
                        If ofila = -1 Then '' clic en la cabecera
                            Exit Sub
                        End If

                        Dim DocEntrySalida = oDataTable.GetValue("Id Salida Mercancias", ofila).ToString()

                        Dim salida As New SalidaCabecera
                        salida.DocEntrySalida = DocEntrySalida
                        salida.Fecha = oDataTable.GetValue("Fecha Salida", ofila).ToString()
                        salida.EmpresaOrigen = rCompany.CompanyName.ToString
                        salida.AlmacenOrigen = oDataTable.GetValue("Bodega Salida", ofila).ToString()
                        salida.EmpresaDestino = oDataTable.GetValue("Empresa Destino", ofila).ToString()
                        salida.AlmacenDestino = oDataTable.GetValue("Bodega Entrada", ofila).ToString()
                        salida.comentario = oDataTable.GetValue("Comentario", ofila).ToString()

                        Dim qryDetalleSalida As String = ""
                        'OINM "BASE_REF" se guarda el docnum
                        'OINM CreatedBy guarda el docentry de la salida
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                            qryDetalleSalida = "select T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"",TO_DECIMAL(REPLACE(T2.""CalcPrice"", ',', '.'), 18, 6) AS ""CalcPrice"" from OIGE T0"
                            qryDetalleSalida += " INNER JOIN IGE1 T1 ON T1.""DocEntry""=T0.""DocEntry"""
                            qryDetalleSalida += " INNER Join OINM T2 ON T2.""ItemCode""=T1.""ItemCode"" And T2.""CreatedBy""=T1.""DocEntry"" And T2.""DocLineNum""=T1.""LineNum"" And t2.""TransType""=T0.""ObjType"" "
                            qryDetalleSalida += " WHERE T0.""DocEntry""= " + DocEntrySalida

                        Else

                            qryDetalleSalida = "select T1.ItemCode,T1.Dscription,T1.Quantity,T2.CalcPrice from OIGE T0"
                            qryDetalleSalida += " INNER JOIN IGE1 T1 ON T1.DocEntry=T0.DocEntry"
                            qryDetalleSalida += " INNER Join OINM T2 ON T2.ItemCode=T1.ItemCode And T2.CreatedBy=T1.DocEntry And T2.DocLineNum=T1.LineNum And t2.TransType=T0.ObjType"
                            qryDetalleSalida += " WHERE T0.DocEntry= " + DocEntrySalida

                        End If

                        Utilitario.Util_Log.Escribir_Log("Query detalle salida de mercancias:" + qryDetalleSalida.ToString(), "frmConsultaSalidaEntrada")

                        Dim rs As SAPbobsCOM.Recordset
                        rs = oFuncionesB1.getRecordSet(qryDetalleSalida)
                        salida.Detalles = New List(Of SalidaDetalle)
                        If rs.RecordCount > 0 Then
                            While (rs.EoF = False)
                                Dim Detalle As New SalidaDetalle
                                Detalle.Codigo = rs.Fields.Item("ItemCode").Value.ToString()
                                Detalle.Nombre = rs.Fields.Item("Dscription").Value.ToString()
                                Detalle.Cantidad = CDbl(rs.Fields.Item("Quantity").Value.ToString())
                                Detalle.Precio = rs.Fields.Item("CalcPrice").Value
                                salida.Detalles.Add(Detalle)
                                rs.MoveNext()
                            End While
                        End If

                        'salida.Detalles.Add(ListaDet)
                        Try
                            Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
                            Dim sRuta As String = sRutaCarpeta & "Filtros_Consulta_Recepcion_FC" + ".xml"
                            If System.IO.Directory.Exists(sRutaCarpeta) Then
                                Utilitario.Util_Log.Escribir_Log("Serializando, clase Salida Cabecera", "frmConsultaSalidaEntrada")

                                Dim x As XmlSerializer = New XmlSerializer(GetType(SalidaCabecera))
                                Dim writer As TextWriter = New StreamWriter(sRuta)
                                x.Serialize(writer, salida)
                                writer.Close()
                                Utilitario.Util_Log.Escribir_Log("Serializado, Parametros de Busqueda" + sRuta, "frmConsultaSalidaEntrada")
                            End If

                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Serializado. Error: " + ex.Message.ToString(), "frmConsultaSalidaEntrada")
                        End Try
                        ofrmTransEntreCompanias.CreaFormularioExistente_frmTransEntreCompanias(salida)

                    End If

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



End Class

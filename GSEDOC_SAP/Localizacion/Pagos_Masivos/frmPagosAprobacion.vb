Public Class frmPagosAprobacion
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private oForm As SAPbouiCOM.Form
    Dim oFila As Integer
    Dim NivelUsr As String
    Dim contador As Integer = 0
    Dim listaDePMAut As List(Of Entidades.SS_PM_AUT)

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioPagosAprobacion()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmPagosAprobacion") Then Exit Sub

        strPath = System.Windows.Forms.Application.StartupPath & "\frmPagosAprobacion.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmPagosAprobacion").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmPagosAprobacion")

            Dim PMAut As Entidades.SS_PM_AUT
            listaDePMAut = New List(Of Entidades.SS_PM_AUT)
            Dim qry As String = ""

            Try
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    qry = "SELECT IFNULL(""U_Usuario"",'') AS ""U_Usuario"", IFNULL(""U_Nivel"",'0') AS ""U_Nivel"" FROM ""@SS_PAG_PERMISOS"""
                Else
                    qry = "SELECT ISNULL(U_Usuario,'') AS U_Usuario, ISNULL(U_Nivel,'0') AS U_Nivel FROM ""@SS_PAG_PERMISOS"""
                End If

                Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rst.DoQuery(qry)

                If rst.RecordCount >= 1 Then
                    While (rst.EoF = False)
                        PMAut = New Entidades.SS_PM_AUT
                        PMAut.Usuario = rst.Fields.Item("U_Usuario").Value
                        PMAut.Nivel = rst.Fields.Item("U_Nivel").Value
                        listaDePMAut.Add(PMAut)
                        rst.MoveNext()
                    End While
                End If

                NivelUsr = (From a In listaDePMAut Where a.Usuario = rCompany.UserName Order By a.Nivel Descending Select a.Nivel).FirstOrDefault

                Utilitario.Util_Log.Escribir_Log($"Query: {qry}, Nivel Usuario: {NivelUsr}", "frmPagosMasivos")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log($"Error obteniendo nivel de usuario: {ex.Message}", "frmPagosMasivos")
            End Try

            If NivelUsr Is Nothing Then
                rsboApp.StatusBar.SetText("Usuario no tiene permiso para utilizar este módulo! Consultelo con el administrador del sistema!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.Close()
                Exit Sub
            End If

            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            ipLogoSS.Item.Visible = True

            Dim finicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
            Dim ffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific

            finicial.Value = DateTime.Now.ToString("yyyyMMdd")
            ffinal.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim btnPro As SAPbouiCOM.Button = oForm.Items.Item("btnPro").Specific
            btnPro.Item.Visible = False

            InicializarValores()

            oForm.Visible = True
            oForm.Select()
        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla frmPagosMasivos: " + ex.Message.ToString())
        End Try
    End Sub

    Private Sub InicializarValores()
        Try
            Dim cbxEst As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEst").Specific
            cbxEst.ValidValues.Add("Revision", "Revision")
            cbxEst.ValidValues.Add("Aprobado", "Aprobado")
            cbxEst.ValidValues.Add("Rechazado", "Rechazado")
            cbxEst.ValidValues.Add("Modificar", "Modificar")
            cbxEst.ValidValues.Add("Archivo Generado", "Archivo Generado")
            cbxEst.ValidValues.Add("Archivo Procesado Banco", "Archivo Procesado Banco")

            If CInt(NivelUsr) > 0 Then
                cbxEst.Select("Revision")
            Else
                cbxEst.Select("Aprobado")
            End If

            CargarDatos()
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un Error al InicializarValores: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un Error al InicializarValores: " & ex.Message, "frmAprobacionPagosMasivos")
        End Try
    End Sub

    Private Sub CargarDatos()
        oForm.Freeze(True)
        Dim filtro As String = ""
        Dim finicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
        Dim ffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific
        Dim cbxEst As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEst").Specific

        rsboApp.StatusBar.SetText("Realizando consulta, por favor espere!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        filtro = " WHERE T0.""CreateDate"" BETWEEN '" & finicial.Value & "' AND '" & ffinal.Value & "' AND T0.""U_Estado"" = '" & cbxEst.Value & "'"

        If CInt(NivelUsr) >= 1 Then filtro += " AND T0.""U_NivelAprob"" = " & NivelUsr

        Dim oRecordSet As SAPbobsCOM.Recordset = CType(rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim query As String = "SELECT B.""Name"" AS ""DepartmentName"" FROM ""OUSR"" A LEFT JOIN ""OUDP"" B ON A.""Department"" = B.""Code"" WHERE A.""USER_CODE"" = '" & rCompany.UserName & "'"
        oRecordSet.DoQuery(query)

        If oRecordSet.RecordCount > 0 Then
            Dim departmentName As String = oRecordSet.Fields.Item("DepartmentName").Value.ToString

            If departmentName = "TALENTO HUMANO *" Then
                filtro += " AND T0.""U_Tipo"" = 'Nomina'"
            Else
                filtro += " AND (T0.""U_Tipo"" = 'Standard' OR (T0.""U_Tipo"" = 'Nomina' AND T0.""U_MedioPago"" = 'Cheque'))"
            End If
        Else
            filtro += " AND (T0.""U_Tipo"" = 'Standard' OR (T0.""U_Tipo"" = 'Nomina' AND T0.""U_MedioPago"" = 'Cheque'))"
        End If
        Utilitario.Util_Log.Escribir_Log("Paso filtro de tipo de solicitud" + filtro, "frmAprobacionPagosMasivos")

        Dim sQuery As String = ""
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            sQuery = My.Resources.SS_PM_OBTENERPAGOS_UDO_HANA & filtro & " ORDER BY T0.""DocEntry"" ASC"
        Else
            sQuery = My.Resources.SS_PM_OBTENERPAGOS_UDO_SQL & filtro & " ORDER BY T0.""DocEntry"" ASC"
        End If

        Try
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + sQuery, "frmAprobacionPagosMasivos")
                oGrid.DataTable.ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmAprobacionPagosMasivos")
            End Try

            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(0).Description = "Seleccion"
            oGrid.Columns.Item(0).TitleObject.Caption = "Seleccion"
            oGrid.Columns.Item(0).Visible = False

            oGrid.Columns.Item(1).Description = "DocEntry"
            oGrid.Columns.Item(1).TitleObject.Caption = "DocEntry"
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).Visible = False

            oGrid.Columns.Item(2).Description = "# Solicitud de pago"
            oGrid.Columns.Item(2).TitleObject.Caption = "# Solicitud de pago"
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Sortable = True

            oGrid.Columns.Item(3).Description = "Tipo de solicitud"
            oGrid.Columns.Item(3).TitleObject.Caption = "Tipo de solicitud"
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Sortable = True

            oGrid.Columns.Item(4).Description = "Medio Pago"
            oGrid.Columns.Item(4).TitleObject.Caption = "Medio Pago"
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Sortable = True

            oGrid.Columns.Item(5).Description = "Usuario solicitante"
            oGrid.Columns.Item(5).TitleObject.Caption = "Usuario solicitante"
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Sortable = True

            oGrid.Columns.Item(6).Description = "Fecha Creacion"
            oGrid.Columns.Item(6).TitleObject.Caption = "Fecha Creacion"
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Sortable = True

            oGrid.Columns.Item(7).Description = "Total a pagar"
            oGrid.Columns.Item(7).TitleObject.Caption = "Total a pagar"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).RightJustified = True

            oGrid.AutoResizeColumns()

            rsboApp.StatusBar.SetText("Consulta terminada con éxito!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error al ejecutar cargar datos: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos:" + sQuery + " - " + ex.Message.ToString, "frmAprobacionPagosMasivos")
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmPagosAprobacion" Then
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If Not pVal.BeforeAction Then
                            Select Case pVal.ItemUID
                                Case "btnBus"
                                    CargarDatos()
                                Case "oGrid"
                                    oFila = pVal.Row
                                Case "btnSolPM"
                                    ofrmPagosMasivos.CargaFormularioPagosMasivos()
                                Case "btnPro"
                                    oForm.Freeze(True)
                                    If CInt(NivelUsr) > 0 Then
                                        rsboApp.StatusBar.SetText("No puede procesar documentos con el nivel actual! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        Dim cbxEst As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEst").Specific

                                        If cbxEst.Value <> "Aprobado" Then
                                            rsboApp.StatusBar.SetText("No pueden procesar pagos con el estado actual! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            contador = 0
                                            Dim Solicitudes As List(Of Integer) = Nothing
                                            CantidadDocumentosSeleccionados(contador, Solicitudes)

                                            If contador = 0 Then
                                                rsboApp.StatusBar.SetText("No se ha seleccionado solicitudes por procesar! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            Else
                                                If ofrmPagosMasivos.CrearPagoMasivoDesdeAprobacion(Solicitudes) Then

                                                    For Each DE As Integer In Solicitudes
                                                        ' ofrmPagosMasivos.ProcesaSolicitudDePagos(DE.ToString) 'revisar
                                                    Next

                                                    Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                                                    For i As Integer = 0 To oGridDet.Rows.Count - 1
                                                        If oGridDet.GetValue("chek", i) = "Y" Then oGridDet.Rows.Remove(i)
                                                    Next
                                                End If
                                            End If
                                        End If
                                    End If
                                    oForm.Freeze(False)
                            End Select
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.BeforeAction = False And pVal.ItemUID = "oGrid" Then
                            If oFila >= 0 Then
                                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                Dim DocEntryUDO As String = ""
                                DocEntryUDO = oGrid.DataTable.GetValue("DocNum", oGrid.GetDataTableRowIndex(pVal.Row))

                                ofrmPagosMasivos.CargaFormularioPagosMasivosExistente(DocEntryUDO, NivelUsr, oFila)
                            End If
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function CantidadDocumentosSeleccionados(ByRef contador As Integer, ByRef solicitudes As List(Of Integer))
        Try
            Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            solicitudes = New List(Of Integer)
            Dim estado As String = ""
            For i As Integer = 0 To oGridDet.Rows.Count - 1
                If oGridDet.GetValue("chek", i) = "Y" Then
                    Dim solicitud As Integer = CInt(oGridDet.GetValue("DocEntry", i))
                    solicitudes.Add(solicitud)
                    contador += 1
                End If
            Next
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error verificando cantidad de documentos a procesar! " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

End Class
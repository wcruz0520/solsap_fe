Public Class frmLogEmision

    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim oGrid As SAPbouiCOM.Grid

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioLogEmision(DocEntry As String, ObjType As String, DocSubType As String)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmLogEmision") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmLogEmision.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmLogEmision").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmLogEmision")

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_Integer, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Tipo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Fecha", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Descripcion", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100000)

            Dim sQuery As String = ""
            If ObjType = "LQE" Then
                ObjType = 18
            End If
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "SELECT B.""LineId"",B.""U_Transacc"",B.""U_Fecha"",B.""U_Detalle"" "
                sQuery += " FROM ""@GS_LOG"" A INNER JOIN"
                sQuery += " ""@GS_LOGD"" B ON A.""DocEntry"" = B.""DocEntry"" "
                sQuery += " WHERE A.""U_Clave"" = '" + LTrim(RTrim(DocEntry)) + "'"
                sQuery += " AND REPLACE(A.""U_ObjType"",'LQE'," + ObjType + ") = '" + LTrim(RTrim(ObjType)) + "'"
                sQuery += " AND A.""U_Tipo"" = 'Emision'"
                sQuery += " AND A.""U_SubType""='" + LTrim(RTrim(DocSubType)) + "'"
            Else
                sQuery = "SELECT B.LineId,B.U_Transacc,B.U_Fecha,B.U_Detalle"
                sQuery += " FROM ""@GS_LOG"" A INNER JOIN"
                sQuery += " ""@GS_LOGD"" B ON A.DocEntry = B.DocEntry"
                sQuery += " WHERE A.U_Clave = " + LTrim(RTrim(DocEntry))
                sQuery += " AND replace(A.U_ObjType,'LQE'," + ObjType + ") = " + LTrim(RTrim(ObjType))
                sQuery += " AND A.U_Tipo = 'Emision'"
                sQuery += " AND A.U_SubType='" + LTrim(RTrim(DocSubType)) + "'"
            End If

            Try
                oForm.DataSources.DataTables.Item("dtDocs").ExecuteQuery(sQuery)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Detalle Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmLogEmision")
            End Try

            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            oGrid.Columns.Item(0).Description = "#"
            oGrid.Columns.Item(0).TitleObject.Caption = """"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "Tipo"
            oGrid.Columns.Item(1).TitleObject.Caption = "Tipo"
            oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "Fecha"
            oGrid.Columns.Item(2).TitleObject.Caption = "Fecha"
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            'oGrid.Columns.Item(5).RightJustified = True

            oGrid.Columns.Item(3).Description = "Detalle"
            oGrid.Columns.Item(3).TitleObject.Caption = "Detalle"
            oGrid.Columns.Item(3).Editable = False

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            'oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()

            oForm.Visible = True
            oForm.Select()

            Try ' BORRA EL MENU
                Dim oMenuItem As SAPbouiCOM.MenuItem
                oMenuItem = rsboApp.Menus.Item("1280")
                If oMenuItem.SubMenus.Exists("SS_LOG") Then
                    oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("SS_LOG"))
                End If
            Catch ex As Exception

            End Try

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
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

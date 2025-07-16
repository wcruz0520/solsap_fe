Imports System.Threading
Imports System.Globalization

Public Class frmConsultaOrdenes
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable
    Dim oCardCode As String = ""
    Dim oObjType As String = ""
    Dim oTipoDocumento As String = ""

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioConsulta(sCardCode As String, ObjType As String, TipoDocumento As String)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConsultaOrdenes") Then
            Exit Sub
        End If
        oCardCode = sCardCode
        oObjType = ObjType
        oTipoDocumento = TipoDocumento
        strPath = System.Windows.Forms.Application.StartupPath & "\frmConsultaOrdenes.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConsultaOrdenes").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmConsultaOrdenes")
            oForm.Freeze(True)

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim sQuery As String = ""
            ' PEDIDO tabla: OPOR objectype: 22
            If ObjType = "22" Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT A.""DocEntry"",A.""DocNum"" AS ""#"",A.""CardCode"",A.""CardName"" AS Nombre,A.""DocDate"" As Fecha,A.""DocTotal"" as Valor,A.""DocType"" as Tipo,A.""DocNum"",IFNULL(A.""Comments"",'') as Comentario FROM ""OPOR"" A  WHERE A.""DocStatus"" = 'O' AND A.""CardCode"" = '" + oCardCode + "'"
                Else
                    sQuery = "SELECT A.DocEntry,A.DocNum AS #,A.CardCode,A.CardName AS Nombre,A.DocDate As Fecha,A.DocTotal as Valor,A.DocType as Tipo,A.DocNum,ISNULL(A.Comments,'') FROM OPOR A WITH(NOLOCK) WHERE A.DocStatus = 'O' AND A.CardCode = '" + oCardCode + "'"
                End If
                oForm.Title = "Pedidos / Ordenes de Compra"
                ' ENTRADA DE MERCANCIA tabla: OPDN objectype: 20
            ElseIf ObjType = "20" Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT A.""DocEntry"",A.""DocNum"" AS ""#"",A.""CardCode"",A.""CardName"" AS Nombre,A.""DocDate"" As Fecha,A.""DocTotal"" as Valor,A.""DocType"" as Tipo,A.""DocNum"",IFNULL(A.""Comments"",'') as Comentario FROM ""OPDN"" A  WHERE A.""DocStatus"" = 'O' AND A.""CardCode"" = '" + oCardCode + "'"
                Else
                    sQuery = "SELECT A.DocEntry,A.DocNum AS #,A.CardCode,A.CardName AS Nombre,A.DocDate As Fecha,A.DocTotal as Valor,A.DocType as Tipo,A.DocNum,ISNULL(A.Comments,'') as Comentario FROM OPDN A WITH(NOLOCK) WHERE A.DocStatus = 'O' AND A.CardCode = '" + oCardCode + "'"
                End If
                oForm.Title = "Entrada de Mercancías"
                ' Devolución de Mercadería tabla: ORPD objectype: 21
            ElseIf ObjType = "21" Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT A.""DocEntry"",A.""DocNum"" AS ""#"",A.""CardCode"",A.""CardName"" AS Nombre,A.""DocDate"" As Fecha,A.""DocTotal"" as Valor,A.""DocNum"",IFNULL(A.""Comments"",'') as Comentario FROM ""ORPD"" A  WHERE A.""DocStatus"" = 'O' AND A.""CardCode"" = '" + oCardCode + "'"
                Else
                    sQuery = "SELECT A.DocEntry,A.DocNum AS #,A.CardCode,A.CardName AS Nombre,A.DocDate As Fecha,A.DocTotal as Valor,A.DocNum,ISNULL(A.Comments,'') as Comentario FROM ORPD A WITH(NOLOCK) WHERE A.DocStatus = 'O' AND A.CardCode = '" + oCardCode + "'"
                End If
                oForm.Title = "Devolución de Mercadería"
                ' Factura de Proveedores tabla: OPCH objectype: 18
            ElseIf ObjType = "18" Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT A.""DocEntry"",A.""DocNum"" AS ""#"",A.""CardCode"",A.""CardName"" AS Nombre,A.""DocDate"" As Fecha,A.""DocTotal"" as Valor,A.""DocNum"",IFNULL(A.""Comments"",'') as Comentario FROM ""OPCH"" A  WHERE A.""DocStatus"" = 'O' AND A.""CardCode"" = '" + oCardCode + "'"
                Else
                    sQuery = "SELECT A.DocEntry,A.DocNum AS #,A.CardCode,A.CardName AS Nombre,A.DocDate As Fecha,A.DocTotal as Valor,A.DocNum,ISNULL(A.Comments,'') as Comentario  FROM OPCH A WITH(NOLOCK) WHERE A.DocStatus = 'O' AND A.CardCode = '" + oCardCode + "'"
                End If
                oForm.Title = "Factura de Proveedores"
                ' Anticipo de Proveedores tabla: ODPO objectype: 204
            ElseIf ObjType = "204" Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    sQuery = "SELECT A.""DocEntry"",A.""DocNum"" AS ""#"",A.""CardCode"",A.""CardName"" AS Nombre,A.""DocDate"" As Fecha,A.""DocTotal"" as Valor,A.""DocNum"",IFNULL(A.""Comments"",'') as Comentario FROM ""ODPO"" A  WHERE A.""DocStatus"" = 'O' AND A.""CardCode"" = '" + oCardCode + "'"
                Else
                    sQuery = "SELECT A.DocEntry,A.DocNum AS #,A.CardCode,A.CardName AS Nombre,A.DocDate As Fecha,A.DocTotal as Valor,A.DocNum,ISNULL(A.Comments,'') as Comentario FROM ODPO A WITH(NOLOCK) WHERE A.DocStatus = 'O' AND A.CardCode = '" + oCardCode + "'"
                End If
                oForm.Title = "Anticipo de Proveedores"
            End If
            Utilitario.Util_Log.Escribir_Log("Consultando Documentos... " & sQuery, "frmConsultaOrdenes")
            oForm.DataSources.DataTables.Item("dtDocs").ExecuteQuery(sQuery)

            Dim oGrid As SAPbouiCOM.Grid
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(0).Visible = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(2).Visible = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Sortable = True
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Sortable = True
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            'If ObjType = "20" Then 'Solo para Entrada de Mercancia
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Sortable = True
            'oGrid.Columns.Item(8).Editable = False
            'oGrid.Columns.Item(8).TitleObject.Sortable = True
            'End If

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Catch CargaFormularioConsulta " & ex.Message.ToString, "frmConsultaOrdenes")
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
                   And pVal.FormTypeEx = "frmConsultaOrdenes" Then
            If Not pVal.Before_Action Then
                Select Case pVal.ItemUID
                    Case "btnSelecc"
                        Try
                            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                            Dim oDT As SAPbouiCOM.DataTable = oGrid.DataTable
                            Dim oDocEntrys As String = ""
                            If oGrid.Rows.SelectedRows.Count > 0 Then
                                For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                    'oDocEntrys = "("
                                    If oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))) <> 0 Then
                                        'A.DocEntry,A.DocNum AS #,A.CardCode,A.CardName AS Nombre,A.DocDate As Fecha,A.DocTotal
                                        If Not i = oGrid.Rows.SelectedRows.Count - 1 Then
                                            oDocEntrys += oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString() + ","
                                        Else
                                            oDocEntrys += oDT.GetValue(0, oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))).ToString()
                                        End If
                                    End If
                                Next
                                If oTipoDocumento = "FC" Then
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        rsboApp.Forms.Item("frmDocumentoXML").Items.Item("objR").Specific.value = oObjType
                                        rsboApp.Forms.Item("frmDocumentoXML").Items.Item("docR").Specific.value = oDocEntrys
                                        ofrmDocumentoXML.CargaDocumentoRelacionados()
                                    Else
                                        rsboApp.Forms.Item("frmDocumento").Items.Item("objR").Specific.value = oObjType
                                        rsboApp.Forms.Item("frmDocumento").Items.Item("docR").Specific.value = oDocEntrys
                                        ofrmDocumento.CargaDocumentoRelacionados()
                                    End If


                                ElseIf oTipoDocumento = "NC" Then
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        rsboApp.Forms.Item("frmDocumentoNCXML").Items.Item("objR").Specific.value = oObjType
                                        rsboApp.Forms.Item("frmDocumentoNCXML").Items.Item("docR").Specific.value = oDocEntrys
                                        ofrmDocumentoNCXML.CargaDocumentoRelacionados()
                                    Else
                                        rsboApp.Forms.Item("frmDocumentoNC").Items.Item("objR").Specific.value = oObjType
                                        rsboApp.Forms.Item("frmDocumentoNC").Items.Item("docR").Specific.value = oDocEntrys
                                        ofrmDocumentoNC.CargaDocumentoRelacionados()
                                    End If

                                End If
                            Else
                                rsboApp.MessageBox(NombreAddon + " Por Favor primero seleccionar un registro..")
                            End If
                            

                            ' END SETEO LOS TOTALES DE LA PESTAÑA DOCUMENTOS RELACIONADOS
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Catch rsboApp_ItemEvent Click " & ex.Message.ToString, "frmConsultaOrdenes")
                        End Try
                       
                End Select
            End If

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

    Public Shared Function formatDecimal(ByVal numero As String) As Decimal

        Dim systemSeparator As Char = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator(0)
        Dim result As Double = 0
        Try
            If numero IsNot Nothing Then
                If Not numero.Contains(",") Then
                    result = Double.Parse(numero, CultureInfo.InvariantCulture)
                Else
                    result = Convert.ToDouble(numero.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()))
                End If
            End If
        Catch e As Exception
            Try
                result = Convert.ToDouble(numero)
            Catch
                Try
                    result = Convert.ToDouble(numero.Replace(",", ";").Replace(".", ",").Replace(";", "."))
                Catch
                    Throw New Exception("Wrong string-to-double format")
                End Try
            End Try
        End Try
        Return result

        'Dim formato As Decimal
        'If Not numero.Equals(String.Empty) Then
        '    Dim sep As Char = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator
        '    Select Case sep
        '        Case "."
        '            formato = numero.Replace(",", sep)
        '        Case ","
        '            formato = numero.Replace(".", sep)
        '    End Select
        'End If
        'Return formato

    End Function
End Class

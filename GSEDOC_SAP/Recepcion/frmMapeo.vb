Public Class frmMapeo
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""  
    Dim odt As SAPbouiCOM.DataTable
    Dim _sCardCode As String = ""
    Dim _fila As String
    Dim _listaDetalleArtiulos As List(Of Entidades.DetalleArticulo)
    Dim _TipoDocumentoAMapear As String

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioMapeo(sRUC As String, sCardCode As String, sNombre As String, listaDetalleArtiulos As List(Of Entidades.DetalleArticulo), ofila As String, TipoDocumentoAMapear As String)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmMapeo") Then
            Exit Sub
        End If
        _sCardCode = sCardCode
        _listaDetalleArtiulos = listaDetalleArtiulos
        _TipoDocumentoAMapear = TipoDocumentoAMapear
        strPath = System.Windows.Forms.Application.StartupPath & "\frmMapeo.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
                _fila = ofila
            Catch exx As Exception
                rsboApp.Forms.Item("frmMapeo").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmMapeo")
            oForm.Freeze(True)

            'oForm.Items.Item("txtRUC").Enabled = False
            Dim txtRUC As SAPbouiCOM.EditText
            txtRUC = oForm.Items.Item("txtRUC").Specific
            txtRUC.Value = sRUC

            'oForm.Items.Item("txtNombre").Enabled = False
            Dim txtNommbre As SAPbouiCOM.EditText
            txtNommbre = oForm.Items.Item("txtNombre").Specific
            txtNommbre.Value = sNombre

            Dim txtCodigo As SAPbouiCOM.EditText
            txtCodigo = oForm.Items.Item("txtCodigo").Specific
            txtCodigo.Value = sCardCode
            'txtCodigo.Item.Enabled = False
            Dim lnkCuentCN As SAPbouiCOM.LinkedButton
            lnkCuentCN = oForm.Items.Item("lnkCuentC").Specific
            lnkCuentCN.LinkedObjectType = 2
            lnkCuentCN.Item.LinkTo = "txtCodigo"

            '' LOST FOCUS
            'Try
            '    oForm.Items.Item("txtCodigo").Enabled = False
            'Catch ex As Exception
            '    oForm.Items.Item("txtFC").Visible = False
            '    Dim txtFC As SAPbouiCOM.EditText
            '    txtFC = oForm.Items.Item("txtFC").Specific
            '    txtFC.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    oForm.Items.Item("txtCodigo").Enabled = False
            'End Try

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodProv", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("DesProv", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CodSAP", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("DesSAP", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 250)

            oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(listaDetalleArtiulos.Count)
            Dim i As Integer = 0
            For Each odetalle As Entidades.DetalleArticulo In _listaDetalleArtiulos
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodProv", i, odetalle.CodigoPrincipal)
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("DesProv", i, odetalle.Descripcion)
                'CodSAP
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", i, oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + sCardCode + "' AND ""Substitute"" = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                Else
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", i, oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WHERE CardCode = '" + sCardCode + "' AND Substitute = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                End If
                'DesSAP
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("DesSAP", i, oFuncionesB1.getRSvalue("SELECT ""U_ItemName"" FROM ""OSCN"" WHERE ""CardCode"" = '" + sCardCode + "' AND ""Substitute"" = '" + odetalle.CodigoPrincipal + "'", "U_ItemName", ""))
                Else
                    oForm.DataSources.DataTables.Item("dtDocs").SetValue("DesSAP", i, oFuncionesB1.getRSvalue("SELECT U_ItemName FROM OSCN WHERE CardCode = '" + sCardCode + "' AND Substitute = '" + odetalle.CodigoPrincipal + "'", "U_ItemName", ""))
                End If
                i += 1
            Next

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            oGrid.Columns.Item(0).Description = "Codigo Proveedor"
            oGrid.Columns.Item(0).TitleObject.Caption = "Codigo Item Proveedor"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "Descripcion Proveedor"
            oGrid.Columns.Item(1).TitleObject.Caption = "Descripcion Item Proveedor"
            oGrid.Columns.Item(1).Editable = False

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "ItemCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            '  oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oGrid.Columns.Item(2).Description = "Codigo"
            oGrid.Columns.Item(2).TitleObject.Caption = "Codigo"
            oGrid.Columns.Item(2).Editable = True

            Dim oItemCodeCol As SAPbouiCOM.EditTextColumn
            oItemCodeCol = CType(oGrid.Columns.Item(2), SAPbouiCOM.EditTextColumn)
            oItemCodeCol.ChooseFromListUID = "CFL1"
            oItemCodeCol.ChooseFromListAlias = "ItemCode"

            oGrid.Columns.Item(3).Description = "Descripcion"
            oGrid.Columns.Item(3).TitleObject.Caption = "Descripcion"
            oGrid.Columns.Item(3).Editable = True

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST _
                   And pVal.FormTypeEx = "frmMapeo" Then
            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
            oCFLEvento = pVal

            If oCFLEvento.BeforeAction = False Then
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
                Try
                    val = oDataTable.GetValue(0, 0)
                    val1 = oDataTable.GetValue(1, 0)
                Catch ex As Exception

                End Try
                If (pVal.ItemUID = "oGrid") And pVal.ColUID = "CodSAP" Then
                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item(pVal.ItemUID).Specific
                    oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                    oGrid.DataTable.SetValue("DesSAP", pVal.Row, val1)
                End If

            End If
        End If
        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
                   And pVal.FormTypeEx = "frmMapeo" Then
            If pVal.BeforeAction = False And pVal.ItemUID = "obtnBuscar" Then
                Dim retval As Integer
                Dim errCode As Integer
                Dim errMsg As String
                Dim oACN As SAPbobsCOM.AlternateCatNum
                Try
                    Try
                        oForm = rsboApp.Forms.Item("frmMapeo")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("oForm - BD:" + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                    End Try

                    Try
                        Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                    
                        For i As Integer = 0 To oDataTable.Rows.Count - 1
                            oACN = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAlternateCatNum)

                            Dim CodSapExistente As String = "SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + _sCardCode + "' AND ""Substitute"" = '" + Left(oDataTable.GetValue(0, i).ToString(), 20) + "'"
                            Dim _CodSapExistente As String = oFuncionesB1.getRSvalue(CodSapExistente, "ItemCode")

                            Utilitario.Util_Log.Escribir_Log("CodSapExistente: " + CodSapExistente.ToString() + " _CodSapExistente: " + _CodSapExistente.ToString + " _sCardCode: " + _sCardCode.ToString + " oDataTable.GetValue(2, i).ToString(): " + oDataTable.GetValue(2, i).ToString(), "frmMapeo")
                            'If Not oACN.GetByKey(oDataTable.GetValue(2, i).ToString(), _sCardCode, oDataTable.GetValue(0, i).ToString()) Then 'se cambio debido a que 100pre ingreaba como nuevo porque tomaba el codigo a actualizar y no actualizaba si no que lo volvia a gregar
                            If Not oACN.GetByKey(_CodSapExistente, _sCardCode, oDataTable.GetValue(0, i).ToString()) Then

                                oACN.ItemCode = oDataTable.GetValue(2, i).ToString()
                                oACN.CardCode = _sCardCode
                                oACN.Substitute = Left(oDataTable.GetValue(0, i).ToString(), 20)
                                'If oDataTable.GetValue(3, i).ToString().Length > 100 Then
                                '    Try
                                '        oACN.UserFields.Fields.Item("U_ItemName").Value = oDataTable.GetValue(3, i).ToString().Substring(0, 99)
                                '    Catch ex As Exception
                                '        Utilitario.Util_Log.Escribir_Log("agregar mayor a 100 - BD:" + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                                '    End Try

                                'Else
                                '    Try
                                '        oACN.UserFields.Fields.Item("U_ItemName").Value = oDataTable.GetValue(3, i).ToString()

                                '    Catch ex As Exception
                                '        Utilitario.Util_Log.Escribir_Log("agregar menor a 100 - BD:" + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                                '    End Try

                                'End If


                                'retval = oACN.Add() se comento para verificar si el error se continua presentando al momento de agregar
                                Try
                                    oACN.Add()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Ocurrio un error al mapear los items: BD: " + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                                End Try

                                If retval <> 0 Then
#Disable Warning BC42030 ' La variable 'errMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                                    rCompany.GetLastError(errCode, errMsg)
#Enable Warning BC42030 ' La variable 'errMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                                    rsboApp.MessageBox(NombreAddon + " - Error al Mapear los articulos! " + errMsg.ToString())
                                    GC.Collect()
                                    'Exit Sub
                                Else
                                    oACN = Nothing
                                    GC.Collect()
                                End If

                            Else
                                oACN.ItemCode = oDataTable.GetValue(2, i).ToString() ' se agrego porque anteriormente solo estaba la linea substitute y lo que se debia actualizar el itemcode
                                'oACN.CardCode = _sCardCode
                                oACN.Substitute = Left(oDataTable.GetValue(0, i).ToString(), 20)
                                'oACN.Substitute = oDataTable.GetValue(0, i).ToString()
                                'If oDataTable.GetValue(3, i).ToString().Length > 100 Then
                                '    Try
                                '        oACN.UserFields.Fields.Item("U_ItemName").Value = oDataTable.GetValue(3, i).ToString().Substring(0, 99)
                                '    Catch ex As Exception
                                '        Utilitario.Util_Log.Escribir_Log("actualizar mayor a 100 - BD:" + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                                '    End Try

                                'Else
                                '    Try
                                '        oACN.UserFields.Fields.Item("U_ItemName").Value = oDataTable.GetValue(3, i).ToString()
                                '    Catch ex As Exception
                                '        Utilitario.Util_Log.Escribir_Log("actualizar menor a 100 - BD:" + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                                '    End Try

                                'End If
                                'retval = oACN.Update() se comento para validar que no el error sea esta linea
                                Try
                                    retval = oACN.Update()
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Ocurrio un error al actualizar el mapeo de los items: BD: " + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                                End Try

                                If retval <> 0 Then
                                    rCompany.GetLastError(errCode, errMsg)
                                    rsboApp.MessageBox(NombreAddon + " - Error al Mapear los articulos! " + errMsg.ToString())
                                    ' Exit Sub
                                    GC.Collect()
                                Else
                                    oACN = Nothing
                                    GC.Collect()
                                End If
                            End If

                            rsboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                        Next
                    Catch ex As Exception
                        rsboApp.StatusBar.SetText(NombreAddon + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Utilitario.Util_Log.Escribir_Log("oDataTable - BD:" + rCompany.DbUserName.ToString() + " USER: " + rCompany.UserName.ToString + " conexion: " + rCompany.Connected.ToString + " ERROR: " + ex.Message.ToString, "frmMapeo")
                    End Try
                    oForm.Items.Item("obtnBuscar").Visible = False
                    oForm.Items.Item("2").Left = oForm.Items.Item("obtnBuscar").Left
                    Dim oB As SAPbouiCOM.Button
                    oB = oForm.Items.Item("2").Specific
                    oB.Caption = "OK"

                    If _TipoDocumentoAMapear = "FV" Then

                        If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                            ActualizaPantallaFacturaRecibidaXML()
                        Else
                            ActualizaPantallaFacturaRecibida()
                        End If

                    ElseIf _TipoDocumentoAMapear = "NC" Then
                        If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                            ActualizaPantallaNotaDeCreditoRecibidaXML()
                        Else
                            ActualizaPantallaNotaDeCreditoRecibida()
                        End If

                    End If


                Catch ex As Exception
                    rsboApp.MessageBox(NombreAddon + " Ocurrio un Error al Grabar el Mapeo de Items: " + ex.Message.ToString())
                Finally
                    oACN = Nothing
                    GC.Collect()
                End Try

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

    'Dim oGrid As SAPbouiCOM.Grid = rsboApp.Forms.Item("frmDocumento").Items.Item("oGrid").Specific
    'oGrid.DataTable = rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs")
    'oGrid.Item.Enabled = False
    'oGrid.Item.FromPane = 0
    'oGrid.Item.ToPane = 0

    'oGrid.Columns.Item(2).Description = "CodSAP"
    'oGrid.Columns.Item(2).TitleObject.Caption = "CodSAP"
    'oGrid.Columns.Item(2).Editable = False
    'Dim oEditTextColum As SAPbouiCOM.EditTextColumn
    'oEditTextColum = oGrid.Columns.Item(2)
    'oEditTextColum.LinkedObjectType = 4

    '' ACTUALIZA EL GRID DE frmDocumentosRecibidos, SETEA A MAPEADO SI
    'Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmDocumentosRecibidos").DataSources.DataTables.Item("dtDocs")
    'odt.SetValue(9, _fila, "SI")
    'Dim oGri As SAPbouiCOM.Grid = rsboApp.Forms.Item("frmDocumentosRecibidos").Items.Item("oGrid").Specific
    'oGri.DataTable = odt
    'oGri.Columns.Item(0).Description = "Tipo Documento"
    'oGri.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
    'oGri.Columns.Item(0).Editable = False

    'oGri.Columns.Item(1).Description = "Fecha"
    'oGri.Columns.Item(1).TitleObject.Caption = "Fecha"
    'oGri.Columns.Item(1).Editable = False

    'oGri.Columns.Item(2).Description = "Folio"
    'oGri.Columns.Item(2).TitleObject.Caption = "Folio"
    'oGri.Columns.Item(2).Editable = False

    'oGri.Columns.Item(3).Description = "RUC"
    'oGri.Columns.Item(3).TitleObject.Caption = "RUC"
    'oGri.Columns.Item(3).Editable = False

    'oGri.Columns.Item(4).Description = "RazonSocial"
    'oGri.Columns.Item(4).TitleObject.Caption = "RazonSocial"
    'oGri.Columns.Item(4).Editable = False

    'oGri.Columns.Item(5).Description = "Valor"
    'oGri.Columns.Item(5).TitleObject.Caption = "Valor"
    'oGri.Columns.Item(5).Editable = False

    'oGri.Columns.Item(6).Description = "ClaveAcceso"
    'oGri.Columns.Item(6).TitleObject.Caption = "Clave de Acceso"
    'oGri.Columns.Item(6).Editable = False
    ''
    'oGri.Columns.Item(7).Description = "NumAutorizacion"
    'oGri.Columns.Item(7).TitleObject.Caption = "Numero de Autorizacion"
    'oGri.Columns.Item(7).Editable = False

    'oGri.Columns.Item(8).Description = "OC"
    'oGri.Columns.Item(8).TitleObject.Caption = "# Orden Compra"
    'oGri.Columns.Item(8).Editable = False
    'oGri.Columns.Item(8).Visible = False

    'Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
    'oEditTextColumn = oGri.Columns.Item(8)
    'oEditTextColumn.LinkedObjectType = 22

    'oGri.Columns.Item(9).Description = "Mapeado"
    'oGri.Columns.Item(9).TitleObject.Caption = "Mapeado"
    'oGri.Columns.Item(9).Editable = False
    'oGri.Columns.Item(9).Visible = False

    'oGri.Columns.Item(10).Description = "Borrador"
    'oGri.Columns.Item(10).TitleObject.Caption = "Documento Preliminar"
    'oGri.Columns.Item(10).Editable = False

    'Dim oEditTextColumn2 As SAPbouiCOM.EditTextColumn
    'oEditTextColumn2 = oGri.Columns.Item(10)
    'oEditTextColumn2.LinkedObjectType = 112

    'oGri.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    'oGri.CollapseLevel = 1
    'oGri.AutoResizeColumns()

    Private Sub ActualizaPantallaFacturaRecibida()
        Try
            ' ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
            Dim PendienteMapear As Boolean = False
            rsboApp.Forms.Item("frmDocumento").Freeze(True)
            rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").Rows.Clear()
            rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").Rows.Add(_listaDetalleArtiulos.Count)
            Dim x As Integer = 0

            For Each odetalle As Entidades.DetalleArticulo In _listaDetalleArtiulos

                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("CodPrin", x, odetalle.CodigoPrincipal.ToString())
                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("CodAuxi", x, IIf(IsNothing(odetalle.CodigoAuxiliar), "", odetalle.CodigoAuxiliar)) '
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + _sCardCode + "' AND ""Substitute"" = '" + Left(odetalle.CodigoPrincipal.Trim, 20) + "'", "ItemCode", ""))
                Else
                    rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WHERE CardCode = '" + _sCardCode + "' AND Substitute = '" + Left(odetalle.CodigoPrincipal.Trim, 20) + "'", "ItemCode", ""))
                End If

                'SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'"
                '   sCardCode = oFuncionesB1.getRSvalue(SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'", "ItemCode", "")
                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("Descripc", x, IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion))
                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("Cantid", x, Convert.ToDouble(odetalle.Cantidad))

                '  Decimal.Parse(oFactura.InfoFactura.totalSinImpuestos, System.Globalization.CultureInfo.InvariantCulture)
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Precio", i, Decimal.Parse(odetalle.precioUnitario, System.Globalization.CultureInfo.InvariantCulture))
                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("Precio", x, Convert.ToDouble(odetalle.PrecioUnitario))
                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("Desc", x, Convert.ToDouble(odetalle.Descuento))
                rsboApp.Forms.Item("frmDocumento").DataSources.DataTables.Item("dtDocs").SetValue("Total", x, Convert.ToDouble(odetalle.PrecioTotalSinImpuesto))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tarifa", i, Convert.ToDouble(odetalle.TarifaIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("ValorImp", i, Convert.ToDouble(odetalle.ValorIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("TotalIm", i, (Convert.ToDouble(odetalle.precioTotalSinImpuesto) + Convert.ToDouble(odetalle.ValorIva)))

                Dim CodigoArticulo As String = ""
                If PendienteMapear = False Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" ='" + _sCardCode + "' AND ""Substitute""  = '" + Left(odetalle.CodigoPrincipal.Trim, 20) + "'", "ItemCode", "")
                    Else
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WITH(NOLOCK) WHERE CardCode ='" + _sCardCode + "' AND Substitute  = '" + Left(odetalle.CodigoPrincipal.Trim, 20) + "'", "ItemCode", "")
                    End If

                    If String.IsNullOrEmpty(CodigoArticulo) Then
                        PendienteMapear = True
                    End If
                End If

                x += 1
            Next

            If PendienteMapear = False Then
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumento").Items.Item("lbMapp").Specific
                lbMapp.Value = "SI"
                lbMapp.Item.ForeColor = RGB(7, 118, 10)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumento").Items.Item("btnMapear").Specific
                btnMapear.Item.Visible = False

            Else
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumento").Items.Item("lbMapp").Specific
                lbMapp.Value = "NO"
                'lbMap.Item.ForeColor = RGB(7, 118, 10)
                lbMapp.Item.ForeColor = ColorTranslator.ToOle(Color.Red)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumento").Items.Item("btnMapear").Specific
                lbMapp.Item.Visible = True
            End If

            rsboApp.Forms.Item("frmDocumento").Freeze(False)
            ' END ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ActualizaPantallaFacturaRecibidaXML()
        Try
            ' ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
            Dim PendienteMapear As Boolean = False
            rsboApp.Forms.Item("frmDocumentoXML").Freeze(True)
            rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").Rows.Clear()
            rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").Rows.Add(_listaDetalleArtiulos.Count)
            Dim x As Integer = 0

            For Each odetalle As Entidades.DetalleArticulo In _listaDetalleArtiulos

                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("CodPrin", x, odetalle.CodigoPrincipal.ToString())
                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("CodAuxi", x, IIf(IsNothing(odetalle.CodigoAuxiliar), "", odetalle.CodigoAuxiliar)) '
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + _sCardCode + "' AND ""Substitute"" = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                Else
                    rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WHERE CardCode = '" + _sCardCode + "' AND Substitute = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                End If

                'SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'"
                '   sCardCode = oFuncionesB1.getRSvalue(SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'", "ItemCode", "")
                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("Descripc", x, IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion))
                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("Cantid", x, Convert.ToDouble(odetalle.Cantidad))

                '  Decimal.Parse(oFactura.InfoFactura.totalSinImpuestos, System.Globalization.CultureInfo.InvariantCulture)
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Precio", i, Decimal.Parse(odetalle.precioUnitario, System.Globalization.CultureInfo.InvariantCulture))
                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("Precio", x, Convert.ToDouble(odetalle.PrecioUnitario))
                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("Desc", x, Convert.ToDouble(odetalle.Descuento))
                rsboApp.Forms.Item("frmDocumentoXML").DataSources.DataTables.Item("dtDocs").SetValue("Total", x, Convert.ToDouble(odetalle.PrecioTotalSinImpuesto))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tarifa", i, Convert.ToDouble(odetalle.TarifaIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("ValorImp", i, Convert.ToDouble(odetalle.ValorIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("TotalIm", i, (Convert.ToDouble(odetalle.precioTotalSinImpuesto) + Convert.ToDouble(odetalle.ValorIva)))

                Dim CodigoArticulo As String = ""
                If PendienteMapear = False Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" ='" + _sCardCode + "' AND ""Substitute""  = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", "")
                    Else
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WITH(NOLOCK) WHERE CardCode ='" + _sCardCode + "' AND Substitute  = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", "")
                    End If

                    If String.IsNullOrEmpty(CodigoArticulo) Then
                        PendienteMapear = True
                    End If
                End If

                x += 1
            Next

            If PendienteMapear = False Then
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumentoXML").Items.Item("lbMapp").Specific
                lbMapp.Value = "SI"
                lbMapp.Item.ForeColor = RGB(7, 118, 10)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumentoXML").Items.Item("btnMapear").Specific
                btnMapear.Item.Visible = False

            Else
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumentoXML").Items.Item("lbMapp").Specific
                lbMapp.Value = "NO"
                'lbMap.Item.ForeColor = RGB(7, 118, 10)
                lbMapp.Item.ForeColor = ColorTranslator.ToOle(Color.Red)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumentoXML").Items.Item("btnMapear").Specific
                lbMapp.Item.Visible = True
            End If

            rsboApp.Forms.Item("frmDocumentoXML").Freeze(False)
            ' END ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub ActualizaPantallaNotaDeCreditoRecibida()
        Try
            ' ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
            Dim PendienteMapear As Boolean = False
            rsboApp.Forms.Item("frmDocumentoNC").Freeze(True)
            rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").Rows.Clear()
            rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").Rows.Add(_listaDetalleArtiulos.Count)
            Dim x As Integer = 0

            For Each odetalle As Entidades.DetalleArticulo In _listaDetalleArtiulos

                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("CodPrin", x, odetalle.CodigoPrincipal.ToString())
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("CodAuxi", x, IIf(IsNothing(odetalle.CodigoAuxiliar), "", odetalle.CodigoAuxiliar)) '
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + _sCardCode + "' AND ""Substitute"" = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                Else
                    rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WHERE CardCode = '" + _sCardCode + "' AND Substitute = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                End If

                'SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'"
                '   sCardCode = oFuncionesB1.getRSvalue(SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'", "ItemCode", "")
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("Descripc", x, IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion))
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("Cantid", x, Convert.ToDouble(odetalle.Cantidad))

                '  Decimal.Parse(oFactura.InfoFactura.totalSinImpuestos, System.Globalization.CultureInfo.InvariantCulture)
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Precio", i, Decimal.Parse(odetalle.precioUnitario, System.Globalization.CultureInfo.InvariantCulture))
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("Precio", x, Convert.ToDouble(odetalle.PrecioUnitario))
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("Desc", x, Convert.ToDouble(odetalle.Descuento))
                rsboApp.Forms.Item("frmDocumentoNC").DataSources.DataTables.Item("dtDocs").SetValue("Total", x, Convert.ToDouble(odetalle.PrecioTotalSinImpuesto))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tarifa", i, Convert.ToDouble(odetalle.TarifaIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("ValorImp", i, Convert.ToDouble(odetalle.ValorIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("TotalIm", i, (Convert.ToDouble(odetalle.precioTotalSinImpuesto) + Convert.ToDouble(odetalle.ValorIva)))

                Dim CodigoArticulo As String = ""
                If PendienteMapear = False Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" ='" + _sCardCode + "' AND ""Substitute""  = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", "")
                    Else
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WITH(NOLOCK) WHERE CardCode ='" + _sCardCode + "' AND Substitute  = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", "")
                    End If

                    If String.IsNullOrEmpty(CodigoArticulo) Then
                        PendienteMapear = True
                    End If
                End If

                x += 1
            Next

            If PendienteMapear = False Then
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("lbMapp").Specific
                lbMapp.Value = "SI"
                lbMapp.Item.ForeColor = RGB(7, 118, 10)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("btnMapear").Specific
                btnMapear.Item.Visible = False

            Else
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("lbMapp").Specific
                lbMapp.Value = "NO"
                'lbMap.Item.ForeColor = RGB(7, 118, 10)
                lbMapp.Item.ForeColor = ColorTranslator.ToOle(Color.Red)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumentoNC").Items.Item("btnMapear").Specific
                lbMapp.Item.Visible = True
            End If

            rsboApp.Forms.Item("frmDocumentoNC").Freeze(False)
            ' END ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub ActualizaPantallaNotaDeCreditoRecibidaXML()
        Try
            ' ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
            Dim PendienteMapear As Boolean = False
            rsboApp.Forms.Item("frmDocumentoNCXML").Freeze(True)
            rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").Rows.Clear()
            rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").Rows.Add(_listaDetalleArtiulos.Count)
            Dim x As Integer = 0

            For Each odetalle As Entidades.DetalleArticulo In _listaDetalleArtiulos

                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("CodPrin", x, odetalle.CodigoPrincipal.ToString())
                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("CodAuxi", x, IIf(IsNothing(odetalle.CodigoAuxiliar), "", odetalle.CodigoAuxiliar)) '
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" = '" + _sCardCode + "' AND ""Substitute"" = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                Else
                    rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("CodSAP", x, oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WHERE CardCode = '" + _sCardCode + "' AND Substitute = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", ""))
                End If

                'SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'"
                '   sCardCode = oFuncionesB1.getRSvalue(SELECT ItemCode FROM OSCN WHERE CardCode = '" + sLicTradNum + "'" AND Substitute = '" + sLicTradNum + "'", "ItemCode", "")
                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("Descripc", x, IIf(IsNothing(odetalle.Descripcion), "", odetalle.Descripcion))
                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("Cantid", x, Convert.ToDouble(odetalle.Cantidad))

                '  Decimal.Parse(oFactura.InfoFactura.totalSinImpuestos, System.Globalization.CultureInfo.InvariantCulture)
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Precio", i, Decimal.Parse(odetalle.precioUnitario, System.Globalization.CultureInfo.InvariantCulture))
                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("Precio", x, Convert.ToDouble(odetalle.PrecioUnitario))
                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("Desc", x, Convert.ToDouble(odetalle.Descuento))
                rsboApp.Forms.Item("frmDocumentoNCXML").DataSources.DataTables.Item("dtDocs").SetValue("Total", x, Convert.ToDouble(odetalle.PrecioTotalSinImpuesto))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("Tarifa", i, Convert.ToDouble(odetalle.TarifaIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("ValorImp", i, Convert.ToDouble(odetalle.ValorIva))
                'oForm.DataSources.DataTables.Item("dtDocs").SetValue("TotalIm", i, (Convert.ToDouble(odetalle.precioTotalSinImpuesto) + Convert.ToDouble(odetalle.ValorIva)))

                Dim CodigoArticulo As String = ""
                If PendienteMapear = False Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ""ItemCode"" FROM ""OSCN"" WHERE ""CardCode"" ='" + _sCardCode + "' AND ""Substitute""  = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", "")
                    Else
                        CodigoArticulo = oFuncionesB1.getRSvalue("SELECT ItemCode FROM OSCN WITH(NOLOCK) WHERE CardCode ='" + _sCardCode + "' AND Substitute  = '" + Left(odetalle.CodigoPrincipal.Trim, 50) + "'", "ItemCode", "")
                    End If

                    If String.IsNullOrEmpty(CodigoArticulo) Then
                        PendienteMapear = True
                    End If
                End If

                x += 1
            Next

            If PendienteMapear = False Then
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumentoNCXML").Items.Item("lbMapp").Specific
                lbMapp.Value = "SI"
                lbMapp.Item.ForeColor = RGB(7, 118, 10)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumentoNCXML").Items.Item("btnMapear").Specific
                btnMapear.Item.Visible = False

            Else
                Dim lbMapp As SAPbouiCOM.EditText
                lbMapp = rsboApp.Forms.Item("frmDocumentoNCXML").Items.Item("lbMapp").Specific
                lbMapp.Value = "NO"
                'lbMap.Item.ForeColor = RGB(7, 118, 10)
                lbMapp.Item.ForeColor = ColorTranslator.ToOle(Color.Red)

                Dim btnMapear As SAPbouiCOM.Button
                btnMapear = rsboApp.Forms.Item("frmDocumentoNCXML").Items.Item("btnMapear").Specific
                lbMapp.Item.Visible = True
            End If

            rsboApp.Forms.Item("frmDocumentoNCXML").Freeze(False)
            ' END ACTUALIZO EL FORMULARIO frmDocumento - FACTURA DE VENTA RECIBIDA
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class

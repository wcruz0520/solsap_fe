'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Public Class frmDocumentosIntegrados
    Public oForm As SAPbouiCOM.Form
    Dim odt As SAPbouiCOM.DataTable
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim oGrid As SAPbouiCOM.Grid

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition

    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

    Dim _ClaveAcceso As String = ""

    Dim i As Integer
    Dim ofila As Integer
    Dim _WS_Recepcion As String = ""
    Dim _WS_RecepcionCambiarEstado As String = ""
    Dim _WS_RecepcionClave As String = ""

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential
    Dim chkmarcado As Boolean





    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormularioDocumentosIntegrados()
        'Dim cardcode As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion

        'Dim sQueryCB As String = ""
        'sQueryCB = " SELECT ""U_SSCLIENTEBANCO"" FROM ""OCRD"" WHERE ""CardCode""= '" + cardcode.Ruc.ToString + "' "
        'Dim clienteBancario As String = oFuncionesB1.getRSvalue(sQueryCB, "U_SSCLIENTEBANCO", "")
        'If clienteBancario = "SI" Then
        '    rsboApp.SetStatusBarMessage("cardcode!" + cardcode.Ruc.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        'End If


        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        rsboApp.SetStatusBarMessage(NombreAddon + " - Consultando, Espere Por favor!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        If RecorreFormulario(rsboApp, "frmDocumentosIntegrados") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDocumentosIntegrados.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDocumentosIntegrados").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmDocumentosIntegrados")
            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO

            oForm.Freeze(True)

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

            Dim cmbTipo As SAPbouiCOM.ComboBox
            cmbTipo = oForm.Items.Item("cbxTipo").Specific
            'cmbTipo.ValidValues.Add("0", "Todos")
            cmbTipo.ValidValues.Add("18", "Factura")
            cmbTipo.ValidValues.Add("19", "Nota de Crédito")
            'cmbTipo.ValidValues.Add("05", "Nota de Débito")
            'cmbTipo.ValidValues.Add("06", "Guía de Remisión")
            cmbTipo.ValidValues.Add("24", "Comp. de Retención")
            cmbTipo.Select("18", SAPbouiCOM.BoSearchKey.psk_ByValue)

            ' CHOOSE FROM LIST
            'Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            'Dim oCons As SAPbouiCOM.Conditions
            'Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            'Dim oCFL As SAPbouiCOM.ChooseFromList
            'Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
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
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            Dim txtRuc As SAPbouiCOM.EditText
            txtRuc = oForm.Items.Item("txtRuc").Specific
            oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtRuc.DataBind.SetBound(True, "", "EditDS")
            txtRuc.ChooseFromListUID = "CFL1"
            'txtRuc.ChooseFromListAlias = "CardCode"
            'txtRuc.SetChooseFromList("MyCFL", "AcctCode")


            Dim lnkPr As SAPbouiCOM.LinkedButton
            lnkPr = oForm.Items.Item("lnkPr").Specific
            'If then
            lnkPr.LinkedObjectType = 2
            'Else
            'lnkPr.LinkedObjectType = 1
            'End If
            lnkPr.Item.LinkTo = "txtRuc"

            'txtFIni
            Dim txtFIni As SAPbouiCOM.EditText
            txtFIni = oForm.Items.Item("txtFIni").Specific
            txtFIni.Value = DateTime.Now.ToString("yyyyMMdd")

            'txtFFin
            Dim txtFFin As SAPbouiCOM.EditText
            txtFFin = oForm.Items.Item("txtFFin").Specific
            txtFFin.Value = DateTime.Now.ToString("yyyyMMdd")

            'documentos marcados
            Dim chkDocInt As SAPbouiCOM.CheckBox
            chkDocInt = oForm.Items.Item("chkDocInt").Specific
            oForm.DataSources.UserDataSources.Add("chkDocInt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            'chkPedido.DataBind.SetBound(True, "", "udChk")
            chkDocInt.ValOn = "Y"
            chkDocInt.ValOff = "N"
            chkDocInt.DataBind.SetBound(True, "", "chkDocInt")

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Tipo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Folio", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Valor", SAPbouiCOM.BoFieldsType.ft_Price, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("ClaveAcceso", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)

            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("U_SSIDDOCUMENTO", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("U_IdGS", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)


            Dim Fecha As String
            Fecha = DateTime.Now.ToString("yyyyMMdd").Substring(0, 4) + "" + DateTime.Now.ToString("yyyyMMdd").Substring(4, 2) + "" + DateTime.Now.ToString("yyyyMMdd").Substring(6, 2)

            cargarDocumentos(Fecha, Fecha, cmbTipo.Value, "")

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST _
               And pVal.FormTypeEx = "frmDocumentosIntegrados" Then

                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                oCFLEvento = pVal

                If oCFLEvento.BeforeAction = False Then
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosIntegrados")
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oDataTable = oCFLEvento.SelectedObjects
                    Dim val As String = String.Empty
                    Dim val1 As String = String.Empty
                    Try
                        val = oDataTable.GetValue(0, 0)
                        val1 = oDataTable.GetValue(1, 0)
                        Try
                            'Dim txtRUC As SAPbouiCOM.EditText
                            'txtRUC = oForm.Items.Item("txtRuc").Specific
                            ''txtRUC.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'txtRUC.Value = val

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


                    Catch ex As Exception
                        Dim txtRaz As SAPbouiCOM.EditText
                        txtRaz = oForm.Items.Item("txtRaz").Specific
                        txtRaz.Value = ""
                    End Try

                End If
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT _
               And pVal.FormTypeEx = "frmDocumentosIntegrados" Then
                If Not pVal.Before_Action Then
                    Try
                        oForm = rsboApp.Forms.Item("frmDocumentosIntegrados")
                        Dim cbxTipo As SAPbouiCOM.ComboBox
                        cbxTipo = oForm.Items.Item("cbxTipo").Specific

                        oCons = oCFL.GetConditions()

                        Dim lbSocio As SAPbouiCOM.StaticText
                        lbSocio = oForm.Items.Item("Item_14").Specific
                        If oCons.Count > 0 Then 'If there are already user conditions.
                            If cbxTipo.Value = "07" Then ' SI ES 07, SIGNIFICA QUE ES RETENCION, POR ENDE PAGO RECIBIDO DE CLIENTE
                                oCons.Item(oCons.Count - 1).CondVal = "C"
                                lbSocio.Caption = "Cliente :"
                            Else
                                oCons.Item(oCons.Count - 1).CondVal = "S"
                                lbSocio.Caption = "Proveedor :"
                            End If
                        End If

                        oCFL.SetConditions(oCons)
                        Dim txtRuc As SAPbouiCOM.EditText
                        txtRuc = oForm.Items.Item("txtRuc").Specific
                        txtRuc.ChooseFromListUID = "CFL1"
                        txtRuc.ChooseFromListAlias = "CardType"

                    Catch ex As Exception

                    End Try


                End If

            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
             And pVal.FormTypeEx = "frmDocumentosIntegrados" Then
                If Not pVal.Before_Action Then
                    Select Case pVal.ItemUID
                        Case "obtnBuscar"
                            Dim Desde As String = ""
                            Desde = oForm.Items.Item("txtFIni").Specific.value.ToString()
                            'Fecha = DateTime.Now.ToString("yyyyMMdd").Substring(0, 4) + "" + DateTime.Now.ToString("yyyyMMdd").Substring(4, 2) + "" + DateTime.Now.ToString("yyyyMMdd").Substring(6, 2)
                            Dim Hasta As String = ""
                            Hasta = oForm.Items.Item("txtFFin").Specific.value.ToString()

                            Dim cmbTipo As SAPbouiCOM.ComboBox
                            cmbTipo = oForm.Items.Item("cbxTipo").Specific

                            cargarDocumentos(Desde, Hasta, cmbTipo.Value, oForm.Items.Item("txtRuc").Specific.value.ToString())
                    End Select
                End If

                'Buscar

            End If

            'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
            'And pVal.FormTypeEx = "frmDocumentosIntegrados" Then
            '    If pVal.Before_Action Then
            '        Event_MatrixLinkPressed(pVal)
            '    End If
            'End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                If pVal.Before_Action Then
                    Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    If typeEx = "frmDocumentosIntegrados" Then
                        Dim _fila As String = pVal.Row
                        If ofila >= 0 Then

                            If chkmarcado = True Then
                                Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosIntegrados")
                                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                oGrid.Rows.SelectedRows.Add(_fila)
                                For i As Integer = 0 To oGrid.Rows.SelectedRows.Count - 1
                                    ofila = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder))
                                Next
                                Dim tipoDocumento As String = oDataTable.GetValue(7, ofila).ToString()
                                Dim idDocumentoRecibido_UDO As String = oDataTable.GetValue(1, ofila).ToString()
                                rsboApp.SetStatusBarMessage(NombreAddon + " - Cargando Documentos Marcado, Por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                                If tipoDocumento = "18" Then
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoXML.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docMarcado")
                                    Else
                                        ofrmDocumento.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docMarcado")
                                    End If

                                ElseIf tipoDocumento = "19" Then

                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoNCXML.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docMarcado")
                                    Else
                                        ofrmDocumentoNC.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docMarcado")
                                    End If

                                ElseIf tipoDocumento = "24" Then

                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoREXML.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docMarcado")
                                    Else
                                        ofrmDocumentoRE.CargaFormularioDocumentoExistente(idDocumentoRecibido_UDO, "docMarcado")
                                    End If

                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Event_MatrixLinkPressed(ByVal pVal As SAPbouiCOM.ItemEvent)

        If pVal.FormTypeEx = "frmDocumentosIntegrados" Then



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

                        Case 19
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes

                        Case 24
                            'If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                            '    oColumns.LinkedObjectType = "TM_RETV"
                            'Else
                            oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oIncomingPayments
                            'End If


                            'Case 1
                            '    oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.

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

    Private Sub cargarDocumentos(Desde As String, Hasta As String, ObjType As String, CardCode As String)
        Try
            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosIntegrados")
            rsboApp.SetStatusBarMessage(NombreAddon + " - Cargando Documentos Integrados, Por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
            oForm.Freeze(True)

            Dim DesdeF As Date
            Dim HastaF As Date
            If Desde <> "" Then
                DesdeF = DateSerial(Convert.ToInt32(Desde.Substring(0, 4)), Convert.ToInt32(Desde.Substring(4, 2)), Convert.ToInt32(Desde.Substring(6, 2)))
                HastaF = DateSerial(Convert.ToInt32(Hasta.Substring(0, 4)), Convert.ToInt32(Hasta.Substring(4, 2)), Convert.ToInt32(Hasta.Substring(6, 2)))

                If DesdeF > HastaF Then
                    rsboApp.SetStatusBarMessage("Para poder filtrar la fecha Fin no puede ser menor a la fecha Inicio..", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Exit Sub
                End If
            End If

            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            oDataTable.Rows.Clear()



            'DesdeF.ToString("yyyyMMdd")
            Dim Query As String = ""
            Dim DocMarInt As SAPbouiCOM.CheckBox
            DocMarInt = oForm.Items.Item("chkDocInt").Specific

            If DocMarInt.Checked = True Then
                chkmarcado = True
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Query = "CALL GS_SAP_FE_ObtenerDocumentosMarcados ("
                    Query += "'" + ObjType + "'"
                    Query += ",'" + CardCode + "'"
                    Query += "," + Functions.FuncionesB1.FechaSql(DesdeF)
                    Query += "," + Functions.FuncionesB1.FechaSql(HastaF) + ")"
                Else
                    Query = "EXEC GS_SAP_FE_ObtenerDocumentosMarcados "
                    Query += "'" + ObjType + "'"
                    Query += ",'" + CardCode + "'"
                    Query += "," + Functions.FuncionesB1.FechaSql(DesdeF)
                    Query += "," + Functions.FuncionesB1.FechaSql(HastaF)
                End If
            Else
                chkmarcado = False
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Query = "CALL GS_SAP_FE_ObtenerDocumentosIntegrados ("
                    Query += "'" + ObjType + "'"
                    Query += ",'" + CardCode + "'"
                    Query += "," + Functions.FuncionesB1.FechaSql(DesdeF)
                    Query += "," + Functions.FuncionesB1.FechaSql(HastaF) + ")"
                Else
                    Query = "EXEC GS_SAP_FE_ObtenerDocumentosIntegrados "
                    Query += "'" + ObjType + "'"
                    Query += ",'" + CardCode + "'"
                    Query += "," + Functions.FuncionesB1.FechaSql(DesdeF)
                    Query += "," + Functions.FuncionesB1.FechaSql(HastaF)
                End If
            End If
            'If ObjType = "18" Then ' SI ES FACTURA DE PROVEEDORES
            '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '        Query = "SELECT 'Factura de Proveedores',A.""DocEntry"", A.""FolioNum"",A.""CardCode"",A.""CardName"",A.""DocTotal"",B.""U_ClaAcc"""
            '        Query += " FROM OPCH A "
            '        Query += " INNER JOIN ""@GS_FVR"" B ON A.""U_SSIDDOCUMENTO"" = B.""DocEntry"" "
            '        Query += " WHERE A.""U_SSCREADAR"" = 'SI'"
            '        Query += " AND A.""CANCELED"" = 'N' "
            '        Query += " AND B.""U_FechaS"" >= " + DesdeF.ToString("yyyyMMdd")
            '        Query += " AND B.""U_FechaS"" <= " + HastaF.ToString("yyyyMMdd")

            '        If Not CardCode.Equals("") Then
            '            Query += " AND ""U_CardCode"" = '" + CardCode + "'"
            '        End If
            '    Else
            '        Query = "SELECT 'Factura de Proveedores',A.DocEntry, A.FolioNum,A.CardCode,A.CardName,A.DocTotal,B.U_ClaAcc"
            '        Query += " FROM OPCH A WITH(NOLOCK) "
            '        Query += " INNER JOIN ""@GS_FVR"" B ON A.U_SSIDDOCUMENTO = B.DocEntry "
            '        Query += " WHERE A.U_SSCREADAR = 'SI'"
            '        Query += " AND A.U_SSIDDOCUMENTO != ''"
            '        Query += " AND A.CANCELED = 'N' "
            '        Query += " AND B.U_FechaS >= " + DesdeF.ToString("yyyyMMdd")
            '        Query += " AND B.U_FechaS <= " + HastaF.ToString("yyyyMMdd")

            '        If Not CardCode.Equals("") Then
            '            Query += " AND A.CardCode = '" + CardCode + "'"
            '        End If
            '    End If
            'ElseIf ObjType = "19" Then
            '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '        Query = "SELECT 'Factura de Proveedores',A.""DocEntry"", A.""FolioNum"",A.""CardCode"",A.""CardName"",A.""DocTotal"",B.""U_ClaAcc"""
            '        Query += " FROM ORPC A "
            '        Query += " INNER JOIN ""@GS_NCR"" B ON A.""U_SSIDDOCUMENTO"" = B.""DocEntry"" "
            '        Query += " WHERE A.""U_SSCREADAR"" = 'SI'"
            '        Query += " AND A.""CANCELED"" = 'N' "
            '        Query += " AND B.""U_FechaS"" >= " + DesdeF.ToString("yyyyMMdd")
            '        Query += " AND B.""U_FechaS"" <= " + HastaF.ToString("yyyyMMdd")

            '        If Not CardCode.Equals("") Then
            '            Query += " AND ""U_CardCode"" = '" + CardCode + "'"
            '        End If
            '    Else
            '        Query = "SELECT 'Nota de Crédito Proveedores',A.DocEntry, A.FolioNum,A.CardCode,A.CardName,A.DocTotal,B.U_ClaAcc"
            '        Query += " FROM ORPC A WITH(NOLOCK) "
            '        Query += " INNER JOIN ""@GS_NCR"" B ON A.U_SSIDDOCUMENTO = B.DocEntry "
            '        Query += " WHERE A.U_SSCREADAR = 'SI'"
            '        Query += " AND A.U_SSIDDOCUMENTO != ''"
            '        Query += " AND A.CANCELED = 'N' "
            '        Query += " AND B.U_FechaS >= " + DesdeF.ToString("yyyyMMdd")
            '        Query += " AND B.U_FechaS <= " + HastaF.ToString("yyyyMMdd")

            '        If Not CardCode.Equals("") Then
            '            Query += " AND A.CardCode = '" + CardCode + "'"
            '        End If
            '    End If
            'End If
            oForm.DataSources.DataTables.Item("dtDocs").ExecuteQuery(Query)
            oGrid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Utilitario.Util_Log.Escribir_Log("consulta: " + Query.ToString, "frmDocumentosIntegrados")
            oGrid.Columns.Item(0).Description = "Tipo Documento"
            oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(1).TitleObject.Caption = "DocEntry"
            If chkmarcado = False Then
                Dim oColA As SAPbouiCOM.GridColumn
                oColA = oGrid.Columns.Item(1)
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE And ObjType = "24" Then
                    oColA.LinkedObjectType = "TM_RETV"
                Else
                    oColA.LinkedObjectType = ObjType
                End If

            End If


            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).TitleObject.Caption = "Folio"

            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Caption = "CardCode"
            Dim oCol2 As SAPbouiCOM.GridColumn
            oCol2 = oGrid.Columns.Item(3)
            oCol2.LinkedObjectType = 2

            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(4).TitleObject.Sortable = True
            oGrid.Columns.Item(4).TitleObject.Caption = "CardName"

            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Sortable = True
            oGrid.Columns.Item(5).TitleObject.Caption = "Valor"
            oGrid.Columns.Item(5).RightJustified = True
            Dim col1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(5)
            Dim oST As SAPbouiCOM.BoColumnSumType = col1.ColumnSetting.SumType
            col1.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            oGrid.Columns.Item(6).TitleObject.Caption = "Clave Acceso"

            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).Visible = False
            oGrid.Columns.Item(7).TitleObject.Caption = "ObjType"

            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(8).Visible = False
            oGrid.Columns.Item(8).TitleObject.Caption = "U_SSIDDOCUMENTO"

            oGrid.Columns.Item(9).Editable = False
            oGrid.Columns.Item(9).Visible = False
            oGrid.Columns.Item(9).TitleObject.Caption = "U_IdGS"

            oGrid.Columns.Item(10).Editable = False
            oGrid.Columns.Item(10).Visible = True
            oGrid.Columns.Item(10).TitleObject.Caption = "Usuario Creador"

            oGrid.Columns.Item(11).Editable = False
            oGrid.Columns.Item(11).Visible = True
            oGrid.Columns.Item(11).TitleObject.Caption = "Fecha Creacion"

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.CollapseLevel = 1
            oGrid.AutoResizeColumns()


        Catch ex As Exception
            'oForm.Freeze(False)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

        Try
            If eventInfo.FormUID = "frmDocumentosIntegrados" And eventInfo.ItemUID = "oGrid" Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                If eventInfo.BeforeAction Then
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    ofila = eventInfo.Row
                    Dim oFor As SAPbouiCOM.Form
                    oFor = rsboApp.Forms.Item("frmDocumentosIntegrados")
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
                    If oMenuItem.SubMenus.Exists("Marcar") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("Marcar"))
                    End If


                    ' If OC > 0 Then
                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "Marcar"
                    oCreationPackage.String = "ReAbrir Documento Integrado..."
                    oCreationPackage.Enabled = True
                    oCreationPackage.Position = 21
                    oMenuItem = rsboApp.Menus.Item("1280")
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                    'End If
                Else
                    oMenuItem = rsboApp.Menus.Item("1280")

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
                If pVal.MenuUID = "Marcar" Then
                    Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    If typeEx = "frmDocumentosIntegrados" Then
                        If ofila >= 0 Then
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmDocumentosIntegrados")
                            Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
                            _ClaveAcceso = oDataTable.GetValue(6, ofila).ToString()
                            'Dim DocEntryUdoFVR As String = oDataTable.GetValue(1, ofila).ToString()
                            Dim tipoDocumento As String = oDataTable.GetValue(7, ofila).ToString()
                            Dim IDDocumentoUDO As String = oDataTable.GetValue(8, ofila).ToString()
                            Dim IDDocumentoeDoc As String = oDataTable.GetValue(9, ofila).ToString()
                            Dim mensaje As String = ""

                            If tipoDocumento = "18" Then
                                Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer ReAbrir el documento ?", 1, "OK", "Cancelar")
                                If respuesta = 1 Then
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(IDDocumentoUDO, 0)
                                        ofrmDocumentoXML.ActualizadoEstadoUdoFacturaXML(IDDocumentoeDoc, "Desmarcado")
                                        oDataTable.Rows.Remove(ofila)
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Re Abierto Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    Else
                                        If MarcarNOVisto(Integer.Parse(IDDocumentoeDoc), 1, mensaje, IDDocumentoUDO) Then
                                            oDataTable.Rows.Remove(ofila)
                                            ' CargaDocumentosFormato("FE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Re Abierto Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If

                                End If
                            ElseIf tipoDocumento = "19" Then
                                Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer ReAbrir el documento ?", 1, "OK", "Cancelar")
                                If respuesta = 1 Then
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then

                                        ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(IDDocumentoUDO, 0)
                                        ofrmDocumentoNCXML.ActualizadoEstadoUdoNotaCreditoXML(IDDocumentoeDoc, "Desmarcado")
                                        oDataTable.Rows.Remove(ofila)
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Re Abierto Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    Else
                                        If MarcarNOVisto(Integer.Parse(IDDocumentoeDoc), 3, mensaje, IDDocumentoUDO) Then
                                            oDataTable.Rows.Remove(ofila)
                                            '  CargaDocumentosFormato("NE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Re Abierto Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If


                                End If
                            ElseIf tipoDocumento = "24" Then
                                Dim respuesta = rsboApp.MessageBox(NombreAddon + " - Esta seguro de querer ReAbrir el documento ?", 1, "OK", "Cancelar")
                                If respuesta = 1 Then

                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then

                                        ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(IDDocumentoUDO, 0)
                                        ofrmDocumentoREXML.ActualizadoEstadoUdoRetencionXML(IDDocumentoeDoc, "Desmarcado")
                                        oDataTable.Rows.Remove(ofila)
                                        rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Re Abierto Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                    Else
                                        If MarcarNOVisto(Integer.Parse(IDDocumentoeDoc), 2, mensaje, IDDocumentoUDO) Then
                                            oDataTable.Rows.Remove(ofila)
                                            '  CargaDocumentosFormato("RE")
                                            rsboApp.SetStatusBarMessage(NombreAddon + " - Documento Re Abierto Exitosamente..!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        Else
                                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio error al marcar como Integrado :" + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If

                                End If
                            End If

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmDocumentosIntegrados")

        End Try

    End Sub

    Public Function MarcarNOVisto(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer, ByRef mensaje As String, idDocumentoRecibido_UDO As String) As Boolean
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
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""

            If SALIDA_POR_PROXY = "Y" Then
                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

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
            'oManejoDocumentos.SetProtocolosdeSeguridad()
            ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
            If WS.MarcarNoVisto(_WS_RecepcionClave, IdDocumento, TipoDocumento, mensaje) Then
                'If chkmarcado = True Then
                '    If TipoDocumento = "1" Then
                '        Elimina_DocumentoRecibido_Factura(idDocumentoRecibido_UDO)
                '    ElseIf TipoDocumento = "3" Then
                '        Elimina_DocumentoRecibido_NC(idDocumentoRecibido_UDO)
                '    ElseIf TipoDocumento = "2" Then
                '        Elimina_DocumentoRecibido_RE(idDocumentoRecibido_UDO)
                '    End If
                '    oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Documento Marcado como NO Visto(Marcado) en EDOC Satisfactoriamente! ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                '    Return True
                'Else
                If TipoDocumento = "1" Then
                    ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(idDocumentoRecibido_UDO, 0)
                ElseIf TipoDocumento = "3" Then
                    ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(idDocumentoRecibido_UDO, 0)
                ElseIf TipoDocumento = "2" Then
                    ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(idDocumentoRecibido_UDO, 0)
                End If
                oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Documento Marcado como NO Visto(Integrado) en EDOC Satisfactoriamente! ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return True
                'End If
            Else
                oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, " Error al marcar documento NO como Visto(Integrado) en EDOC, no se tuvo respuesta con los WS ", Functions.FuncionesAddon.Transacciones.Cancelacion, Functions.FuncionesAddon.TipoLog.Recepcion)
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_Factura(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            'oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Try
                Utilitario.Util_Log.Escribir_Log("Clave FC: " + _ClaveAcceso + " Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), "frmDocumentosIntegrados")
            Catch ex As Exception
            End Try

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_SincroE", Sincronizado)
            If chkmarcado = True Then
                oGeneralData.SetProperty("U_Estado", "docReAbierto")
            Else
                oGeneralData.SetProperty("U_Estado", "docReAbierto")
            End If

            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_NC(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            'oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Try
                Utilitario.Util_Log.Escribir_Log("Clave NC: " + _ClaveAcceso + " Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), "frmDocumentosIntegrados")
            Catch ex As Exception
            End Try
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_SincroE", Sincronizado)
            If chkmarcado = True Then
                oGeneralData.SetProperty("U_Estado", "docReAbierto")
            Else
                oGeneralData.SetProperty("U_Estado", "docReAbierto")
            End If

            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
        End Try
    End Function

    Public Function ActualizadoEstadoSincronizadoEDOC_DocumentoRecibido_RE(ByRef DocEntryFacturaRecibida_UDO As String, Sincronizado As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            'oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Try
                Utilitario.Util_Log.Escribir_Log("Clave RT: " + _ClaveAcceso + " Actualizando a Sincronizado EDOC = " + Sincronizado.ToString(), "frmDocumentosIntegrados")
            Catch ex As Exception
            End Try
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oGeneralData.SetProperty("U_SincroE", Sincronizado)
            If chkmarcado = True Then
                oGeneralData.SetProperty("U_Estado", "docReAbierto")
            Else
                oGeneralData.SetProperty("U_Estado", "docReAbierto")
            End If
            oGeneralService.Update(oGeneralData)

            Return True
            '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
            'sDocEntry = oGeneralParams.GetProperty("Code")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " -  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oFuncionesAddon.GuardaLOG("PRR", _ClaveAcceso, "  Ocurrio error al actualizar el estado de la sincronizacion EDOC :" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.CreacionFinal, Functions.FuncionesAddon.TipoLog.Recepcion)
            Return False
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

    Public Sub Elimina_DocumentoRecibido_Factura(DocEntryFacturaRecibida_UDO As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Try
            oFuncionesAddon.GuardaLOG("FVR", _ClaveAcceso, "Eliminando Registro del Documento Marcado UDO Factura # " + DocEntryFacturaRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.EliminarDocMarcado, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("Eliminando Registro del Documento Marcado UDO Factura # " + DocEntryFacturaRecibida_UDO.ToString(), "frmDocumentosIntegrados")
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_FVR")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryFacturaRecibida_UDO)
            oGeneralService.Delete(oGeneralParams)
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("FVR", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO factura..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.EliminarDocMarcado, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("Eliminando Registro del Documento Marcado UDO Factura # " + DocEntryFacturaRecibida_UDO.ToString() + " - error: " + ex.Message.ToString, "frmDocumentosIntegrados")
        End Try

    End Sub
    Public Sub Elimina_DocumentoRecibido_NC(DocEntryNCRecibida_UDO As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Try
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Eliminando registro del documento Marcado UDO Nota de credito # " + DocEntryNCRecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.EliminarDocMarcado, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("Eliminando Registro del Documento Marcado UDO nota de credito # " + DocEntryNCRecibida_UDO.ToString(), "frmDocumentosIntegrados")
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_NCR")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryNCRecibida_UDO)
            oGeneralService.Delete(oGeneralParams)

        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("NCR", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO nota de credito..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.EliminarDocMarcado, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("Eliminando Registro del Documento Marcado UDO nota de credito # " + DocEntryNCRecibida_UDO.ToString() + " - error: " + ex.Message.ToString, "frmDocumentosIntegrados")
        End Try
    End Sub
    Public Sub Elimina_DocumentoRecibido_RE(DocEntryRERecibida_UDO As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Try
            oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Eliminando Registro del Documento Marcado UDO RETENCION # " + DocEntryRERecibida_UDO.ToString(), Functions.FuncionesAddon.Transacciones.EliminarDocMarcado, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("Eliminando Registro del Documento Marcado UDO RETENCION # " + DocEntryRERecibida_UDO.ToString(), "frmDocumentosIntegrados")
            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("GS_RER")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryRERecibida_UDO)
            oGeneralService.Delete(oGeneralParams)
        Catch ex As Exception
            oFuncionesAddon.GuardaLOG("REE", _ClaveAcceso, "Error: Eliminando Documento Recibido UDO RETENCION..: " + ex.Message().ToString(), Functions.FuncionesAddon.Transacciones.EliminarDocMarcado, Functions.FuncionesAddon.TipoLog.Recepcion)
            Utilitario.Util_Log.Escribir_Log("Eliminando Registro del Documento Marcado UDO RETENCION # " + DocEntryRERecibida_UDO.ToString() + " - error: " + ex.Message.ToString, "frmDocumentosIntegrados")
        End Try
    End Sub
End Class

Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports SAPbobsCOM
Imports SAPbouiCOM


Public Class frmChequePD
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Private _oCheque As clsCheque
    Dim oDocumentoSAP As SAPbobsCOM.Payments
    Dim _fila As Integer

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CreaFormulario_frmChequePD(oCheque As clsCheque, ValorProtesto As Decimal, ofila As Integer)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmChequePD") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmChequePD.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmChequePD").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmChequePD")

            oForm.Freeze(True)

            _oCheque = oCheque
            _fila = ofila

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"

            'txtCliente
            'txtClienteN
            'txtND

            Dim labelFecha As SAPbouiCOM.StaticText
            labelFecha = oForm.Items.Item("Item_10").Specific

            oForm.DataSources.UserDataSources.Add("dtFI", SAPbouiCOM.BoDataType.dt_DATE, 20)
            Dim txtFCont As SAPbouiCOM.EditText
            txtFCont = oForm.Items.Item("txtFCont").Specific
            txtFCont.DataBind.SetBound(True, "", "dtFI")

            Dim btnImp As SAPbouiCOM.Button
            btnImp = oForm.Items.Item("btnImp").Specific ''MONTO PROTESTO
            btnImp.Item.Visible = False

            Dim txtProtest As SAPbouiCOM.EditText
            txtProtest = oForm.Items.Item("txtProtest").Specific ''MONTO PROTESTO
            txtProtest.Value = ValorProtesto

            Dim txtCheque As SAPbouiCOM.EditText
            txtCheque = oForm.Items.Item("txtCheque").Specific
            txtCheque.Value = _oCheque.Cheque_Num

            Dim txtChequeV As SAPbouiCOM.EditText
            txtChequeV = oForm.Items.Item("txtChequeV").Specific
            txtChequeV.Value = _oCheque.Cheque_Valor

            Dim lbBanco As SAPbouiCOM.StaticText
            lbBanco = oForm.Items.Item("lbBanco").Specific
            lbBanco.Caption = _oCheque.Banco

            Dim txtPago As SAPbouiCOM.EditText
            txtPago = oForm.Items.Item("txtPago").Specific
            txtPago.Value = _oCheque.NumPago

            Dim txtPagoD As SAPbouiCOM.EditText
            txtPagoD = oForm.Items.Item("txtPagoD").Specific
            txtPagoD.Value = _oCheque.Pago_Coments

            Dim txtDep As SAPbouiCOM.EditText
            txtDep = oForm.Items.Item("txtDep").Specific
            txtDep.Value = _oCheque.NumeroDeposito


            Dim txtCliente As SAPbouiCOM.EditText
            txtCliente = oForm.Items.Item("txtCliente").Specific
            txtCliente.Value = _oCheque.Cliente_Codigo

            Dim txtClienteN As SAPbouiCOM.EditText
            txtClienteN = oForm.Items.Item("txClienteN").Specific
            txtClienteN.Value = _oCheque.Cliente

            Dim txtND As SAPbouiCOM.EditText
            txtND = oForm.Items.Item("txtND").Specific
            txtND.Value = _oCheque.Doc_Protesto
            If _oCheque.Doc_Protesto > 0 Then
                Dim obtnProc As SAPbouiCOM.Button
                obtnProc = oForm.Items.Item("btnProc").Specific
                obtnProc.Item.Enabled = False
            End If

            ' CHOOSE FROM LIST CUENTA 
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "FormatCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            'Choose from list, para NC
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCFL.SetConditions(oCons)

            Dim txtMC As SAPbouiCOM.EditText
            txtMC = oForm.Items.Item("txtMC").Specific
            oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtMC.DataBind.SetBound(True, "", "EditDS")
            txtMC.ChooseFromListUID = "CFL1"
            txtMC.ChooseFromListAlias = "FormatCode"

            Dim txtMP As SAPbouiCOM.EditText
            txtMP = oForm.Items.Item("txtMP").Specific
            oForm.DataSources.UserDataSources.Add("EditDST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtMP.DataBind.SetBound(True, "", "EditDST")
            txtMP.ChooseFromListUID = "CFL2"
            txtMP.ChooseFromListAlias = "FormatCode"


            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
            'oUserDataSource.ValueEx = functions.VariablesGlobales._SS_CuentaMontoCheque
            oUserDataSource.ValueEx = _oCheque.CuentaContableDeposito
            Dim lCuentaF As SAPbouiCOM.StaticText
            lCuentaF = oForm.Items.Item("lbMC").Specific
            'lCuentaF.Caption = functions.VariablesGlobales._SS_NombreCuentaMontoCheque
            lCuentaF.Caption = _oCheque.NombreCuentaContableDeposito

            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDST")
            'oUserDataSource.ValueEx = functions.VariablesGlobales._SS_CuentaMontoProtesto
            oUserDataSource.ValueEx = _oCheque.CuentaContableDeposito
            Dim lCuentaN As SAPbouiCOM.StaticText
            lCuentaN = oForm.Items.Item("lbMP").Specific
            'lCuentaN.Caption = functions.VariablesGlobales._SS_NombreCuentaMontoProtesto
            lCuentaN.Caption = _oCheque.NombreCuentaContableDeposito


            ' CHOOSE FROM LIST IMPUESTO 
            Dim oCFLsI As SAPbouiCOM.ChooseFromListCollection
            Dim oConsI As SAPbouiCOM.Conditions
            Dim oConI As SAPbouiCOM.Condition
            oCFLsI = oForm.ChooseFromLists
            Dim oCFLI As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParamsI As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParamsI = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParamsI.MultiSelection = False
            oCFLCreationParamsI.ObjectType = "128"
            oCFLCreationParamsI.UniqueID = "CFL1MC"
            oCFLI = oCFLsI.Add(oCFLCreationParamsI)
            oConsI = oCFLI.GetConditions()

            oConI = oConsI.Add()
            oConI.Alias = "Code"
            oConI.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
            'oCon.CondVal = "S"
            oCFLI.SetConditions(oConsI)

            oCFLCreationParamsI.UniqueID = "CFL2MP"
            oCFLI = oCFLsI.Add(oCFLCreationParamsI)
            oCFLI.SetConditions(oConsI)

            Dim txtImpMC As SAPbouiCOM.EditText
            txtImpMC = oForm.Items.Item("txtImpMC").Specific
            oForm.DataSources.UserDataSources.Add("EditDSMC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtImpMC.DataBind.SetBound(True, "", "EditDSMC")
            txtImpMC.ChooseFromListUID = "CFL1MC"
            txtImpMC.ChooseFromListAlias = "Code"

            Dim txtImpMP As SAPbouiCOM.EditText
            txtImpMP = oForm.Items.Item("txtImpMP").Specific
            oForm.DataSources.UserDataSources.Add("EditDSTMP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtImpMP.DataBind.SetBound(True, "", "EditDSTMP")
            txtImpMP.ChooseFromListUID = "CFL2MP"
            txtImpMP.ChooseFromListAlias = "Code"


            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSMC")
            oUserDataSource.ValueEx = Functions.VariablesGlobales._SS_ImpuestoMontoCheque
            Dim lbIMC As SAPbouiCOM.StaticText
            lbIMC = oForm.Items.Item("lbIMC").Specific
            lbIMC.Caption = Functions.VariablesGlobales._SS_NombreImpuestoMontoCheque

            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSTMP")
            oUserDataSource.ValueEx = Functions.VariablesGlobales._SS_impuestoMontoProtesto
            Dim lbIMP As SAPbouiCOM.StaticText
            lbIMP = oForm.Items.Item("lbIMP").Specific
            lbIMP.Caption = Functions.VariablesGlobales._SS_NombreImpuestoMontoProtesto

            If oForm.Items.Item("btnProc").Enabled = False Then
                btnImp.Item.Visible = True
                txtFCont.Item.Visible = False
                labelFecha.Item.Visible = False
            End If

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.SetStatusBarMessage(ex.Message(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Utilitario.Util_Log.Escribir_Log("ex CreaFormulario_frmChequePD:  " + ex.Message.ToString(), "frmChequePD")
        End Try

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmChequePD" Then

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If Not pVal.Before_Action Then

                            Select Case pVal.ItemUID


                                Case "obtnCerrar"

                                    oForm.Close()


                                Case "btnProc"

                                    oForm = rsboApp.Forms.Item("frmChequePD")
                                    Dim DocEntry_ND As Integer = 0
                                    Dim result As Boolean = False
                                    result = CrearNotaDeDebito(oForm, _oCheque, DocEntry_ND)

                                    If result = True Then
                                        Dim txtND As SAPbouiCOM.EditText
                                        txtND = oForm.Items.Item("txtND").Specific ''DOCENTRY NOTA DE DEBITO
                                        txtND.Value = DocEntry_ND.ToString()

                                        Dim btnProc As SAPbouiCOM.Button
                                        btnProc = oForm.Items.Item("btnProc").Specific
                                        btnProc.Item.Enabled = False

                                        Dim btnImp As SAPbouiCOM.Button
                                        btnImp = oForm.Items.Item("btnImp").Specific
                                        btnImp.Item.Visible = True

                                        Dim label As SAPbouiCOM.StaticText
                                        label = oForm.Items.Item("Item_10").Specific
                                        label.Item.Visible = False

                                        oForm.Items.Item("txtComen").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                        Dim txtFCont As SAPbouiCOM.EditText
                                        txtFCont = oForm.Items.Item("txtFCont").Specific
                                        txtFCont.Item.Visible = False

                                        ActualizarPagoRecibido_Cheque(oForm, DocEntry_ND, _oCheque)
                                    End If

                                Case "btnImp"
                                    oForm = rsboApp.Forms.Item("frmChequePD")

                                    Dim docentry As SAPbouiCOM.EditText = oForm.Items.Item("txtND").Specific
                                    Dim rutaReporte As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        rutaReporte = System.Windows.Forms.Application.StartupPath & "\" + rsboApp.Company.DatabaseName + "_SS_NotaDebitoProtestoHana.rpt"
                                        PresentarPDFNDHANA(docentry.Value, rutaReporte, "Nota Debito", "Dockey@")
                                    Else
                                        rutaReporte = System.Windows.Forms.Application.StartupPath & "\" + rsboApp.Company.DatabaseName + "_SS_NotaDebitoProtestoSql.rpt"
                                        PresentarPDFNDSQL(docentry.Value, rutaReporte, "Nota Debito", "Dockey@")
                                    End If


                            End Select
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
                                        Case "txtMC"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                            Dim lCuentaF As SAPbouiCOM.StaticText
                                            lCuentaF = oForm.Items.Item("lbMC").Specific
                                            lCuentaF.Caption = oDataTable.GetValue("AcctName", 0)

                                        Case "txtMP"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDST")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("FormatCode", 0)
                                            Dim lCuentaN As SAPbouiCOM.StaticText
                                            lCuentaN = oForm.Items.Item("lbMP").Specific
                                            lCuentaN.Caption = oDataTable.GetValue("AcctName", 0)

                                        Case "txtImpMC"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSMC")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("Code", 0)
                                            Dim lbIMC As SAPbouiCOM.StaticText
                                            lbIMC = oForm.Items.Item("lbIMC").Specific
                                            lbIMC.Caption = oDataTable.GetValue("Name", 0)

                                        Case "txtImpMP"
                                            oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDSTMP")
                                            oUserDataSource.ValueEx = oDataTable.GetValue("Code", 0)
                                            Dim lbIMP As SAPbouiCOM.StaticText
                                            lbIMP = oForm.Items.Item("lbIMP").Specific
                                            lbIMP.Caption = oDataTable.GetValue("Name", 0)

                                    End Select
                                End If
                            Catch ex As Exception
                                rsboApp.MessageBox("et_CHOOSE_FROM_LIST " + ex.Message.ToString())
                                Utilitario.Util_Log.Escribir_Log("ex et_CHOOSE_FROM_LIST:  " + ex.Message.ToString(), "frmChequePD")
                            End Try
                        End If
                End Select
            End If
        Catch ex As Exception
            ' rsboApp.MessageBox("ex ActualizarPagoRecibido_Cheque: " + ex.Message.ToString())
            Utilitario.Util_Log.Escribir_Log("ex ActualizarPagoRecibido_Cheque::  " + ex.Message.ToString(), "frmChequePD")

        End Try

    End Sub

    Private Sub ActualizarPagoRecibido_Cheque(oForm As SAPbouiCOM.Form, docEntry_ND As Integer, oCheque As clsCheque)
        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            oDocumentoSAP = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            oDocumentoSAP.GetByKey(oCheque.NumPago)
            For count As Integer = 0 To oDocumentoSAP.Checks.Count - 1
                oDocumentoSAP.Checks.SetCurrentLine(count)
                oDocumentoSAP.Checks.UserFields.Fields.Item("U_SS_IDND").Value = docEntry_ND
                'oDocumentoSAP.Checks.UserFields.Fields.Item("U_SS_NDPROTE").Value = "SI"
            Next

            oDocumentoSAP.Update()
            RetVal = oDocumentoSAP.Update()
            If RetVal <> 0 Then
                rCompany.GetLastError(ErrCode, ErrMsg)
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)

            Else
                rsboApp.SetStatusBarMessage(NombreAddon + " - Generacion de Nota de Debito Protesto correctamente creada..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                'ofrmChequeP.Formulario_frmChequeP_CargarGrid()

                ' ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL
                rsboApp.Forms.Item("frmChequeP").Freeze(True)
                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmChequeP").DataSources.DataTables.Item("dtDocs")
                odt.SetValue("Doc_Protesto", _fila, Integer.Parse(docEntry_ND))
                rsboApp.Forms.Item("frmChequeP").Freeze(False)
                ' END ACTUALIZO EL GRID DE FORMULARIO PRINCIPAL

            End If


        Catch ex As Exception
            rsboApp.MessageBox("ex ActualizarPagoRecibido_Cheque: " + ex.Message.ToString())
            Utilitario.Util_Log.Escribir_Log("ex ActualizarPagoRecibido_Cheque  " + ex.Message.ToString(), "frmChequePD")
        End Try

    End Sub

    Private Function CrearNotaDeDebito(oForm As SAPbouiCOM.Form, oCheque As clsCheque, ByRef sDocEntry As String) As Boolean

        rsboApp.StatusBar.SetText(NombreAddon + " - Creando Nota de Debito por favor espere..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'Create the Documents object
        Dim GRPO As SAPbobsCOM.Documents
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim CodImp As String = ""
        Dim sQueryCodImp As String = ""
        'Dim CodImpV As String = ""
        Dim sQueryCodImpV As String = ""

        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim oSeriesService As SAPbobsCOM.SeriesService
        Dim oSeries As SAPbobsCOM.Series
        Dim oDocumentTypeParams As SAPbobsCOM.DocumentTypeParams

        'Servicio para obtener serie por default


        Try
            Dim txtProtest As SAPbouiCOM.EditText
            txtProtest = oForm.Items.Item("txtProtest").Specific ''MONTO PROTESTO

            Dim txtFCont As SAPbouiCOM.EditText
            txtFCont = oForm.Items.Item("txtFCont").Specific ''FECHA PROTESTO

            If formatDecimal(txtProtest.Value) <= 0 Then
                rsboApp.StatusBar.SetText(NombreAddon + " - Favor ingrese el valor del protesto.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If

            If txtFCont.Value = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - Favor ingrese la fecha del protesto.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If

            oCmpSrv = rCompany.GetCompanyService
            oSeriesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)
            oSeries = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries)
            oDocumentTypeParams = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams)
            oDocumentTypeParams.Document = 13 ' CAMBIAR OBJTYPE Y DOCSUBTYPE DEPENDIENDO DEL OBJETO
            oDocumentTypeParams.DocumentSubType = "DN"
            oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams)

            GRPO = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            GRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
            GRPO.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo

            GRPO.CardCode = oCheque.Cliente_Codigo
            'GRPO.DocDate = ofactura.FechaEmision
            'GRPO.DocDueDate = ofactura.FechaEmision
            'GRPO.TaxDate = ofactura.FechaEmision
            Dim FECHA As String = oForm.Items.Item("txtFCont").Specific.value.ToString()

            GRPO.DocDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now
            GRPO.TaxDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now
            GRPO.DocDueDate = DateSerial(Convert.ToInt32(FECHA.Substring(0, 4)), Convert.ToInt32(FECHA.Substring(4, 2)), Convert.ToInt32(FECHA.Substring(6, 2))) 'Now

            Dim qrySerieND As String = ""
            If rCompany.DbServerType = 9 Then
                qrySerieND = "select ""Series"" from " + rCompany.CompanyDB.ToString() + ".""NNM1"" where ""ObjectCode"" = 13 AND ""DocSubType""='DN' and ""SeriesName""='" + Functions.VariablesGlobales._SS_IDSerie + "'"
            Else
                qrySerieND = "select Series from NNM1 where ObjectCode = 13 AND DocSubType='DN' and SeriesName='" + Functions.VariablesGlobales._SS_IDSerie + "'"
            End If
            Dim SerieND As String = oFuncionesB1.getRSvalue(qrySerieND, "Series", "0")

            GRPO.Series = CInt(SerieND)

            'GRPO.UserFields.Fields.Item("U_SSCREADAR").Value = "SI"

            GRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            GRPO.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO

            Dim txtMC As SAPbouiCOM.EditText
            txtMC = oForm.Items.Item("txtMC").Specific ''CUENTA MONTO CHEQUE
            Dim txtMP As SAPbouiCOM.EditText
            txtMP = oForm.Items.Item("txtMP").Specific  ''CUENTA MONTO PROTESTO

            If txtMC.Value = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - Favor selecciones una cuenta contable para monto del cheque.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If
            If txtMP.Value = "" Then
                rsboApp.StatusBar.SetText(NombreAddon + " - Favor selecciones una cuenta contable para monto del protesto.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If

            Dim FormatCode_MontoCheque As String = txtMC.Value
            Dim FormatCode_MontoProtesto As String = txtMP.Value
            Dim sQueryAcctCode As String = ""

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode_MontoCheque + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode_MontoCheque + "'"
            End If
            Dim Cuenta_MontoCheque As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryAcctCode = "Select ""AcctCode"" from ""OACT"" Where ""FormatCode"" =  '" + FormatCode_MontoProtesto + "'"
            Else
                sQueryAcctCode = "Select AcctCode from OACT Where FormatCode =  '" + FormatCode_MontoProtesto + "'"
            End If
            Dim Cuenta_MontoProtesto As String = oFuncionesB1.getRSvalue(sQueryAcctCode, "AcctCode", "")

            '' LINEA DE MONTO CHEQUE
            Dim txtChequeV As SAPbouiCOM.EditText
            txtChequeV = oForm.Items.Item("txtChequeV").Specific ''MONTO CHEQUE
            Dim txtImpMC As SAPbouiCOM.EditText
            txtImpMC = oForm.Items.Item("txtImpMC").Specific ''IMPUESTO CHEQUE

            GRPO.Lines.AccountCode = Cuenta_MontoCheque
            GRPO.Lines.LineTotal = formatDecimal(txtChequeV.Value)
            GRPO.Lines.ItemDescription = "MONTO CHEQUE PROTESTADO"
            GRPO.Lines.Quantity = 1
            GRPO.Lines.TaxCode = txtImpMC.Value
            GRPO.Lines.Add()

            '' LINEA DE MONTO PROTESTO

            Dim txtImpMP As SAPbouiCOM.EditText
            txtImpMP = oForm.Items.Item("txtImpMP").Specific ''IMPUESTO MONTO PROTESTO

            GRPO.Lines.AccountCode = Cuenta_MontoProtesto
            GRPO.Lines.LineTotal = formatDecimal(txtProtest.Value)
            GRPO.Lines.ItemDescription = "MONTO PROTESTO"
            GRPO.Lines.Quantity = 1
            GRPO.Lines.TaxCode = txtImpMP.Value
            GRPO.Lines.Add()


            Dim txtComen As SAPbouiCOM.EditText
            txtComen = oForm.Items.Item("txtComen").Specific ''COMENTARIO
            GRPO.Comments += txtComen.Value
            GRPO.JournalMemo += txtComen.Value
            GRPO.UserFields.Fields.Item("U_SS_PROTESTO").Value = "SI"

            RetVal = GRPO.Add()
            If RetVal <> 0 Then
                rCompany.GetLastError(ErrCode, ErrMsg)
                rsboApp.MessageBox(ErrCode & " " & ErrMsg)
                Return False
            Else
                rCompany.GetNewObjectCode(sDocEntry)
                Return True
            End If

        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error:" + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("ex CrearNotaDeDebito:  " + ex.Message.ToString(), "frmChequePD")
            Return False
        Finally
            GRPO = Nothing
            GC.Collect()
        End Try

    End Function

    Public Sub PresentarPDFNDSQL(_DocEntry As String, RutaReporte As String, NombreReporte As String, Parametro1 As String)
        Try
            rsboApp.SetStatusBarMessage("Generando el Reporte...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            Dim _ipServidor = ""
            If String.IsNullOrEmpty(Functions.VariablesGlobales._ipServer) Then
                _ipServidor = rCompany.Server
            Else
                _ipServidor = Functions.VariablesGlobales._ipServer
            End If

            Dim reportDoc As New ReportDocument()
            'reportDoc.Load("C:\Users\David Macias\Documents\ECUADOR\Clientes\Nota Debito.rpt")
            reportDoc.Load(RutaReporte)

            reportDoc.SetParameterValue(Parametro1, _DocEntry)

            Dim filepath As String = Path.GetTempPath()
            'filepath += "Nota Debito.pdf"
            NombreReporte += "_" + _DocEntry
            filepath += NombreReporte + ".pdf"

            Dim connectionInfo As New ConnectionInfo()
            connectionInfo.ServerName = _ipServidor
            connectionInfo.DatabaseName = rCompany.CompanyDB
            connectionInfo.UserID = Functions.VariablesGlobales._gUsuarioDB
            connectionInfo.Password = Functions.VariablesGlobales._gPasswordDB

            For Each table As Table In reportDoc.Database.Tables
                Dim tableLogOnInfo As TableLogOnInfo = table.LogOnInfo
                tableLogOnInfo.ConnectionInfo = connectionInfo
                table.ApplyLogOnInfo(tableLogOnInfo)
            Next

            reportDoc.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, filepath)

            rsboApp.SetStatusBarMessage("Abriendo Documento " & NombreReporte & " Espere unos segundos..", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            Dim Proc As New Process()
            Proc.StartInfo.FileName = filepath
            Proc.Start()
            Proc.Dispose()

        Catch ex As Exception
            rsboApp.MessageBox("Error al presentar PDF: " + ex.Message.ToString())
            Utilitario.Util_Log.Escribir_Log("Error al presentar PDF: " + ex.Message.ToString(), "frmChequePD")
        End Try

    End Sub

    Public Sub PresentarPDFNDHANA(_DocEntry As String, RutaReporte As String, NombreReporte As String, Parametro1 As String)

        rsboApp.SetStatusBarMessage("Generando el Reporte...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        Try
            Dim crReport As ReportDocument = Nothing
            'Dim crReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim filepath As String = Path.GetTempPath()
            filepath += "Nota Debito.pdf"

            Try

                crReport = New ReportDocument
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al instanciar clase ReportDocument : " + ex.Message.ToString(), "frmGenerarRPT")
            End Try
            If (File.Exists(RutaReporte)) Then
                Utilitario.Util_Log.Escribir_Log("Ruta si Existe", "frmGenerarRPT")
                crReport.Load(RutaReporte)
                Utilitario.Util_Log.Escribir_Log("Cargando variables RPT", "frmGenerarRPT")

                Dim _ipServidor = ""
                If Functions.VariablesGlobales._ipServer = "" Then
                    _ipServidor = rCompany.Server
                Else
                    _ipServidor = Functions.VariablesGlobales._ipServer
                End If

                Dim ConexionHana As String = String.Empty
                If (IntPtr.Size = 8) Then
                    ConexionHana = String.Concat(ConexionHana, "Driver={B1CRHPROXY};")
                Else
                    ConexionHana = String.Concat(ConexionHana, "Driver={B1CRHPROXY32};")
                End If

                If String.IsNullOrEmpty(_ipServidor) Then
                    ConexionHana = String.Concat(ConexionHana, "ServerNode=", rCompany.Server & ";")

                Else
                    ConexionHana = String.Concat(ConexionHana, "SERVERNODE=", _ipServidor & ";")

                End If


                ConexionHana = String.Concat(ConexionHana, "DATABASE=", rCompany.CompanyDB, ";")

                ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._gUsuarioDB, ";")
                ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._gPasswordDB, ";")


                Utilitario.Util_Log.Escribir_Log("Conexion: " + ConexionHana, "frmGenerarRPT")

                crReport.SetParameterValue(Parametro1, _DocEntry)

                Dim logonProps2 As NameValuePairs2 = crReport.DataSourceConnections(0).LogonProperties
                Try
                    If (IntPtr.Size = 8) Then
                        logonProps2.Set("Provider", "B1CRHPROXY")
                        logonProps2.Set("Server Type", "B1CRHPROXY")
                    Else
                        logonProps2.Set("Provider", "B1CRHPROXY32")
                        logonProps2.Set("Server Type", "B1CRHPROXY32")
                    End If

                    logonProps2.Set("Connection String", ConexionHana)
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("Excepcion logonProps2 :" + ex.Message, "frmGenerarRPT")

                End Try

                Try
                    Utilitario.Util_Log.Escribir_Log("Conexion Iniciando DataSourceConnections... ", "frmGenerarRPT")
                    crReport.DataSourceConnections(0).SetLogonProperties(logonProps2)
                    If String.IsNullOrEmpty(_ipServidor) Then
                        crReport.DataSourceConnections(0).SetConnection(rCompany.Server, rCompany.CompanyDB, False)
                    Else
                        Utilitario.Util_Log.Escribir_Log("Conexion Iniciando DataSourceConnections: " + _ipServidor + " Company: " + rCompany.CompanyDB + " DB: " + Functions.VariablesGlobales._gPasswordDB, "frmGenerarRPT")
                        crReport.DataSourceConnections(0).SetConnection(_ipServidor, rCompany.CompanyDB, False)
                        'crReport.DataSourceConnections(0).SetConnection(Functions.VariablesGlobales._gIpServidorDB, rCompany.CompanyDB, Functions.VariablesGlobales._gPasswordDB)
                        Utilitario.Util_Log.Escribir_Log("Conexion Iniciando DataSourceConnections: " + crReport.DataSourceConnections(0).ToString(), "frmGenerarRPT")
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("Excepcion crReport.DataSourceConnections :" + ex.Message, "frmGenerarRPT")

                End Try

                crReport.PrintOptions.PrinterName = ""

                Dim IOST As IO.Stream = Nothing
                Try
                    IOST = crReport.ExportToStream(ExportFormatType.PortableDocFormat)
                    crReport.Close() ' libero la memoria del documento crystal
                    If IOST.Length > 0 Then

                        Dim b(IOST.Length) As Byte

                        IOST.Read(b, 0, CInt(IOST.Length))


                        NombreReporte += "_" + _DocEntry
                        Dim temporal As String = Path.GetTempPath() & "\" & NombreReporte + ".pdf"
                        Utilitario.Util_Log.Escribir_Log("ruta temporal: " + temporal.ToString, "frmGenerarRPT")
                        ''If File.Exists(temporal) Then
                        ''    File.Delete(temporal)
                        ''End If

                        'Dim ms As MemoryStream = New MemoryStream(b)
                        File.WriteAllBytes(temporal, b)


                        rsboApp.SetStatusBarMessage("Abriendo Documento " & NombreReporte & " Espere unos segundos..", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        Process.Start(temporal)



                    End If
                Catch ex As Exception
                    rsboApp.SetStatusBarMessage("Excepcion al exportar el PDF a Stream " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                End Try


            Else

                rsboApp.SetStatusBarMessage("No se encontro el Archivo RPT de Crystal", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            End If



        Catch ex As Exception

            rsboApp.SetStatusBarMessage("Solsap ,Error al generar RPT: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)

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

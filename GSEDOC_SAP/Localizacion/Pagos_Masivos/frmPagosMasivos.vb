Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports Spire.Pdf
Imports System.Drawing.Printing
Imports System.IO
Imports System.Text

Public Class frmPagosMasivos
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private btnAprob As SAPbouiCOM.ButtonCombo
    Private cbxfpago As SAPbouiCOM.ComboBox
    Private cbxCuenta As SAPbouiCOM.ComboBox
    Private cbxSuc As SAPbouiCOM.ComboBox

    Private txtPro As SAPbouiCOM.EditText
    Private txtFecCor As SAPbouiCOM.EditText

    Private Item_4 As SAPbouiCOM.StaticText

    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCon As SAPbouiCOM.Condition
    Dim oUserDataSource As SAPbouiCOM.UserDataSource

    Dim TotalFacts As Decimal = 0, totalMontoSP = 0, totalSaldoSP = 0
    Dim Query As String = ""
    Dim NivelUsuario As String = "", NivelMaximo As String = ""

    Dim DocEntry As String = "", NivelAprobacion As String = "", TipoSolicitudPMTransferencia As String = ""
    Dim referenciaBotones As Integer
    Dim Linea As Integer

    Dim ListadeLineas As New List(Of Integer)
    'Dim ofila As Integer

    Dim listaDePMAut As List(Of Entidades.SS_PM_AUT)

    Public num As Integer = 0
    Dim RutaDeArchivoGenerado As String = ""

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioPagosMasivos()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmPagosMasivos") Then Exit Sub

        strPath = System.Windows.Forms.Application.StartupPath & "\frmPagosMasivos.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmPagosMasivos").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmPagosMasivos")
            oForm.Freeze(True)

            Dim PMAut As Entidades.SS_PM_AUT
            listaDePMAut = New List(Of Entidades.SS_PM_AUT)
            Query = ""
            NivelUsuario = ""
            NivelMaximo = ""

            Try
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Query = "SELECT IFNULL(""U_Usuario"",'') AS ""U_Usuario"", IFNULL(""U_Nivel"",'0') AS ""U_Nivel"" FROM ""@SS_PAG_PERMISOS"""
                Else
                    Query = "SELECT ISNULL(U_Usuario,'') AS U_Usuario, ISNULL(U_Nivel,'0') AS U_Nivel FROM ""@SS_PAG_PERMISOS"""
                End If

                Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rst.DoQuery(Query)

                If rst.RecordCount >= 1 Then
                    While (rst.EoF = False)
                        PMAut = New Entidades.SS_PM_AUT
                        PMAut.Usuario = rst.Fields.Item("U_Usuario").Value
                        PMAut.Nivel = rst.Fields.Item("U_Nivel").Value
                        listaDePMAut.Add(PMAut)
                        rst.MoveNext()
                    End While
                End If

                NivelUsuario = (From a In listaDePMAut Where a.Usuario = rCompany.UserName Order By a.Nivel Descending Select a.Nivel).FirstOrDefault
                NivelMaximo = (From a In listaDePMAut Order By a.Nivel Descending Select a.Nivel).FirstOrDefault

                Utilitario.Util_Log.Escribir_Log($"Query: {Query}, Nivel Usuario: {NivelUsuario}, Nivel Maximo: {NivelMaximo}", "frmPagosMasivos")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log($"Error obteniendo nivel de usuario y nivel maximo: {ex.Message}", "frmPagosMasivos")
            End Try

            If NivelUsuario Is Nothing Then
                rsboApp.StatusBar.SetText("Usuario no tiene permiso para utilizar este módulo! Consultelo con el administrador del sistema!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.Close()
                Exit Sub
            End If

            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            ipLogoSS.Item.Visible = True

            Dim lnkPro As SAPbouiCOM.LinkedButton = oForm.Items.Item("lnkPro").Specific
            lnkPro.LinkedObjectType = 2
            lnkPro.Item.LinkTo = "txtPro"

            Dim txtFecCor As SAPbouiCOM.EditText = oForm.Items.Item("txtFecCor").Specific
            txtFecCor.Value = DateTime.Now.ToString("yyyyMMdd")

            ListadeLineas = New List(Of Integer)
            DocEntry = ""

            Item_4 = oForm.Items.Item("Item_4").Specific
            referenciaBotones = Item_4.Item.Top 'Guardamos la posicion del lbl de referencia para replace en modo busqueda y niveles de aprobador            

            InicializaControlesXNivel(NivelUsuario, "Modificar")

            BloqueaHabilitaBotones(NivelUsuario, "Modificar")

            Dim focus As SAPbouiCOM.Item = oForm.Items.Item("focus")
            focus.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            oForm.EnableMenu("1281", False)
            oForm.EnableMenu("1282", False)

            oForm.EnableMenu("1288", False)
            oForm.EnableMenu("1291", False)
            oForm.EnableMenu("1289", False)
            oForm.EnableMenu("1290", False)

            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim ColConsolidado As SAPbouiCOM.Column = oMatrix.Columns.Item("U_Consoli")
            Dim ColChe As SAPbouiCOM.Column = oMatrix.Columns.Item("U_NumChe") 'Ocultamos la columna de cheque
            ColConsolidado.Visible = False
            ColChe.Visible = False

            oMatrix.AutoResizeColumns()

            ''
            'Dim PruebaBanco As SAPbobsCOM.HouseBankAccounts = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oHouseBankAccounts)
            'If PruebaBanco.GetByKey("2") Then
            '    PruebaBanco.NextCheckNo = 5000
            '    PruebaBanco.Update()
            '    ''Dim CoodBanco = PruebaBanco.BankCode
            '    ''Dim NomBanco = PruebaBanco.BankName
            'End If

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            oForm.Visible = True
            oForm.Freeze(False)
            oForm.Select()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al cargar la pantalla frmPagosMasivos: " & ex.Message, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Ocurrio un error al cargar la pantalla frmPagosMasivos: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub CargaFormularioPagosMasivosExistente(ByVal DocEntryUDO As Integer, ByVal NivelUsr As String, ByVal LineaPA As Integer)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmPagosMasivos") Then Exit Sub

        strPath = System.Windows.Forms.Application.StartupPath & "\frmPagosMasivos.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmPagosMasivos").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmPagosMasivos")
            oForm.EnableMenu("1281", False) ' BUSCAR
            oForm.EnableMenu("1282", False) ' NUEVO

            oForm.EnableMenu("1288", False)
            oForm.EnableMenu("1291", False)
            oForm.EnableMenu("1289", False)
            oForm.EnableMenu("1290", False)

            oForm.Freeze(True)

            Linea = 0
            Linea = LineaPA
            RutaDeArchivoGenerado = ""
            TipoSolicitudPMTransferencia = ""

            rsboApp.StatusBar.SetText("Consultando documento, espere por favor!... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            ipLogoSS.Item.Visible = True

            Item_4 = oForm.Items.Item("Item_4").Specific
            referenciaBotones = Item_4.Item.Top 'Guardamos la posicion del lbl de referencia para replace en modo busqueda y niveles de aprobador

            Query = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Query = "SELECT TOP 1 IFNULL(""U_Nivel"",'0') AS ""Nivel"" FROM ""@SS_PAG_PERMISOS"" ORDER BY ""U_Nivel"" DESC"
            Else
                Query = "SELECT TOP 1 ISNULL(U_Nivel,'0') AS Nivel FROM ""@SS_PAG_PERMISOS"" WITH(NOLOCK) ORDER BY U_Nivel DESC"
            End If
            Utilitario.Util_Log.Escribir_Log("Query para obtener el nivel máximo de permisos: " & Query.ToString, "frmPagosMasivos")
            NivelMaximo = oFuncionesAddon.getRSvalue(Query, "Nivel", "0")
            Utilitario.Util_Log.Escribir_Log("Nivel máximo: " & NivelMaximo, "frmPagosMasivos")

            NivelUsuario = NivelUsr

            Dim oCompanyService As SAPbobsCOM.CompanyService = rCompany.GetCompanyService()
            Dim oGeneralService As SAPbobsCOM.GeneralService = oCompanyService.GetGeneralService("SSMTPAGOS")
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUDO)

            Dim oGeneralData As SAPbobsCOM.GeneralData = oGeneralService.GetByParams(oGeneralParams)

            Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific
            txtDocEnt.Value = DocEntryUDO.ToString ' oGeneralData.GetProperty("DocEntry")

            Dim txtNvAp As SAPbouiCOM.EditText = oForm.Items.Item("txtNvAp").Specific
            txtNvAp.Value = oGeneralData.GetProperty("U_NivelAprob")

            Dim txtPagCon As SAPbouiCOM.EditText = oForm.Items.Item("txtPagCon").Specific
            txtPagCon.Value = oGeneralData.GetProperty("U_IdPagCon")

            'Dim txtArcBco As SAPbouiCOM.EditText = oForm.Items.Item("txtArcBco").Specific
            'txtArcBco.Value = oGeneralData.GetProperty("U_RutArcBan")

            Dim lblABCO As SAPbouiCOM.StaticText = oForm.Items.Item("lblABCO").Specific
            lblABCO.Caption = oGeneralData.GetProperty("U_RutArcBan")

            RutaDeArchivoGenerado = oGeneralData.GetProperty("U_RutArcGen") 'Guardo ruta de archivo generado para actualizacion de nombre en el caso que modifiquen una solicitud

            Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific
            txtEstado.Value = oGeneralData.GetProperty("U_Estado")

            Dim txtCBan As SAPbouiCOM.EditText = oForm.Items.Item("txtCBan").Specific
            txtCBan.Value = oGeneralData.GetProperty("U_IdCashBan")

            Dim lblFac As SAPbouiCOM.StaticText = oForm.Items.Item("lblFac").Specific
            lblFac.Caption = oGeneralData.GetProperty("U_FacProcesadas")

            Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific
            txtMP.Value = oGeneralData.GetProperty("U_TotalPagado")

            Try
                Dim hora As DateTime
                Dim fecha As DateTime

                Dim FechaArchivo As String = oGeneralData.GetProperty("U_FechaArcRec")

                Dim esHora As Boolean = DateTime.TryParseExact(oGeneralData.GetProperty("U_FechaArcRec"), "HH:mm:ss", Nothing, Globalization.DateTimeStyles.None, hora)
                Dim esFecha As Boolean = DateTime.TryParseExact(FechaArchivo, "d/MM/yyyy", Nothing, Globalization.DateTimeStyles.None, fecha)

                If esHora = False Then
                    If esFecha Then
                        'Dim FechaArchivo As String = oGeneralData.GetProperty("U_FechaArcRec")
                        If String.IsNullOrEmpty(FechaArchivo) Or FechaArchivo <> "0:00:00" Then
                            Dim fechaConvertida As Date
                            If Date.TryParse(FechaArchivo, fechaConvertida) Then
                                Dim txtFArc As SAPbouiCOM.EditText = oForm.Items.Item("txtFArc").Specific '
                                txtFArc.Value = fechaConvertida.ToString("yyyyMMdd")
                            End If
                        End If
                    End If
                End If

                Dim FechaDevolucion As String = oGeneralData.GetProperty("U_FechaDev")

                esHora = DateTime.TryParseExact(oGeneralData.GetProperty("U_FechaDev"), "HH:mm:ss", Nothing, Globalization.DateTimeStyles.None, hora)
                esFecha = DateTime.TryParseExact(FechaDevolucion, "d/MM/yyyy", Nothing, Globalization.DateTimeStyles.None, fecha)

                If esHora = False Then
                    If esFecha Then
                        'Dim FechaDevolucion As String = oGeneralData.GetProperty("U_FechaDev")
                        If String.IsNullOrEmpty(FechaDevolucion) Or FechaDevolucion <> "0:00:00" Then
                            Dim fechaConvertida As Date
                            If Date.TryParse(FechaDevolucion, fechaConvertida) Then
                                Dim txtFDev As SAPbouiCOM.EditText = oForm.Items.Item("txtFDev").Specific
                                txtFDev.Value = fechaConvertida.ToString("yyyyMMdd")
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error seccion de carga de fechas: " + ex.Message, "frmPagosMasivos")
            End Try

            DocEntry = oGeneralData.GetProperty("DocEntry").ToString
            NivelAprobacion = oGeneralData.GetProperty("U_NivelAprob").ToString
            TipoSolicitudPMTransferencia = oGeneralData.GetProperty("U_Tipo").ToString

            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_PM_DET1") ' "UDO_D1" es un ejemplo del nombre de la tabla hija

            Dim oConditions As SAPbouiCOM.Conditions = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            Dim oCondition As SAPbouiCOM.Condition = oConditions.Add()
            oCondition.Alias = "DocEntry"  ' Campo en la tabla que quieres filtrar
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL  ' Operación de la condición (igual, mayor, menor, etc.)
            oCondition.CondVal = DocEntryUDO.ToString
            oDBDataSource.Query(oConditions)
            oMatrix.LoadFromDataSource()

            Dim focus As SAPbouiCOM.Item = oForm.Items.Item("focus")
            focus.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Dim totMon As Decimal = 0, totSal = 0

            For i As Integer = 0 To oDBDataSource.Size - 1
                Dim mon As String = oDBDataSource.GetValue("U_Monto", i).Trim()
                totMon += Math.Round(CDec(mon), 2)
                Dim sal As String = oDBDataSource.GetValue("U_Saldo", i).Trim()
                totSal += Math.Round(CDec(sal), 2)
            Next

            Dim lblSMon As SAPbouiCOM.StaticText = oForm.Items.Item("lblSMon").Specific
            lblSMon.Caption = totMon.ToString
            Dim lblSSal As SAPbouiCOM.StaticText = oForm.Items.Item("lblSSal").Specific
            lblSSal.Caption = totSal.ToString

            If txtEstado.Value = "Modificar" Then
                BloqueaHabilitaBotones(NivelUsr, "Modificar")
                InicializaControlesXNivel(NivelUsr, "Modificar")
            Else
                BloqueaHabilitaBotones(NivelUsr, "Buscar")
                InicializaControlesXNivel(NivelUsr, "Buscar", oGeneralData.GetProperty("U_MedioPago").ToString)
                LLenaComboCuenta()
            End If

            cbxCuenta = oForm.Items.Item("cbxCuenta").Specific
            cbxCuenta.Select(oGeneralData.GetProperty("U_Cuenta"), SAPbouiCOM.BoSearchKey.psk_ByValue)

            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
            cbxfpago.Select(oGeneralData.GetProperty("U_MedioPago"))

            Dim ColTra As SAPbouiCOM.Column = oMatrix.Columns.Item("U_CtaBco") 'Ocultamos la columna de cheque
            Dim ColChe As SAPbouiCOM.Column = oMatrix.Columns.Item("U_NumChe") 'Ocultamos la columna de cheque
            Dim ColCon As SAPbouiCOM.Column = oMatrix.Columns.Item("U_Consoli") 'Ocultamos la columna de cheque
            Dim btncashm As SAPbouiCOM.Button = oForm.Items.Item("btncashm").Specific
            Dim btnND As SAPbouiCOM.Button = oForm.Items.Item("btnND").Specific

            Dim result As String = ""
            Dim lblNumChe As SAPbouiCOM.StaticText = oForm.Items.Item("lblNumChe").Specific

            If cbxfpago.Value = "Transferencia" Then
                ColCon.Visible = False
                ColChe.Visible = False
                ColTra.Visible = True
            ElseIf cbxfpago.Value = "Cheque" Then
                ColCon.Visible = True
                ColChe.Visible = False
                ColTra.Visible = False
                btncashm.Item.Enabled = False
                btnND.Item.Enabled = False

                Dim separadorCuentaBanco As String() = cbxCuenta.Selected.Description.Split(":")

                'Query = ""
                'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                '    Query = "SELECT IFNULL(""U_NumChe"",0) AS ""NumCheque"" FROM ""@SS_PM_CONTROLCHEQUE"" WHERE ""Code"" = '" & separadorCuentaBanco(0) & "' AND ""Name"" = '" & cbxCuenta.Value & "'"
                'Else
                '    Query = "SELECT ISNULL(U_NumChe,0) AS NumCheque FROM ""@SS_PM_CONTROLCHEQUE"" WITH(NOLOCK) WHERE Code = '" & separadorCuentaBanco(0) & "' AND Name = '" & cbxCuenta.Value & "'"
                'End If
                'result = oFuncionesAddon.getRSvalue(Query, "NumCheque", "0")
                'lblNumChe.Caption = result.ToString

                Dim oRecordSet As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("SELECT * FROM ""DSC1"" WHERE ""BankCode"" = '" & separadorCuentaBanco(0) & "' AND ""GLAccount"" = '" & cbxCuenta.Value & "'")
                Dim CtaBcoPropio As SAPbobsCOM.HouseBankAccounts = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oHouseBankAccounts)
                CtaBcoPropio.Browser.Recordset = oRecordSet
                lblNumChe.Caption = CtaBcoPropio.NextCheckNo

                oRecordSet = Nothing
                GC.Collect()
            End If

            ColChe.Visible = False

            oMatrix.AutoResizeColumns()

            Dim lblRAG As SAPbouiCOM.StaticText = oForm.Items.Item("lblRAG").Specific
            lblRAG.Caption = oGeneralData.GetProperty("U_RutArcGen").ToString
            lblRAG.Item.ForeColor = RGB(6, 69, 173)
            lblRAG.Item.TextStyle = 4

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            oForm.Freeze(False)
            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Ocurrio un Error al CargaFormularioPagosMasivosExistente: " + ex.Message, "frmPagosMasivos")
            rsboApp.MessageBox("Ocurrio un Error al CargaFormularioPagosMasivosExistente: " + ex.Message)
        End Try
    End Sub

    Private Sub BloqueaHabilitaBotones(ByVal Nivel As String, Optional ByVal Modo As String = "")
        Try
            Dim btnND As SAPbouiCOM.Button = oForm.Items.Item("btnND").Specific
            Dim btnimpgs As SAPbouiCOM.Button = oForm.Items.Item("btnimpgs").Specific
            Dim btncashm As SAPbouiCOM.Button = oForm.Items.Item("btncashm").Specific
            Dim btnSolPag As SAPbouiCOM.Button = oForm.Items.Item("btnSolPag").Specific
            Dim btnAprob As SAPbouiCOM.ButtonCombo = oForm.Items.Item("btnAprob").Specific

            Dim Item_8 As SAPbouiCOM.StaticText = oForm.Items.Item("Item_8").Specific 'lbl proveedor
            txtPro = oForm.Items.Item("txtPro").Specific 'txt Proveedor
            Dim lblRaz As SAPbouiCOM.StaticText = oForm.Items.Item("lblRaz").Specific 'lbl Proveedor nombre
            Dim Item_12 As SAPbouiCOM.StaticText = oForm.Items.Item("Item_12").Specific 'lbl Fecha corte
            txtFecCor = oForm.Items.Item("txtFecCor").Specific 'txt Fecha corte
            Dim lblSuc As SAPbouiCOM.StaticText = oForm.Items.Item("lblSuc").Specific
            Dim cbxSuc As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxSuc").Specific 'cbx Sucursal

            Dim chkPLSB As SAPbouiCOM.CheckBox = oForm.Items.Item("chkPLSB").Specific 'chk Servicios basicos 13/12/2024

            Dim btnfilt As SAPbouiCOM.Button = oForm.Items.Item("btnfilt").Specific 'btn Filtrar
            Dim btnMDT As SAPbouiCOM.Button = oForm.Items.Item("btnMDT").Specific 'btn Marcar/desmarcar todo
            Dim btnAddMar As SAPbouiCOM.Button = oForm.Items.Item("btnAddMar").Specific 'btn Agregar marcados
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific 'Grid de facturas

            Dim Item_0 As SAPbouiCOM.StaticText = oForm.Items.Item("Item_0").Specific 'lvl id solicitud
            Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific 'txt id Solicitud (DocEntry)
            Dim Item_2 As SAPbouiCOM.StaticText = oForm.Items.Item("Item_2").Specific 'lvl nivel de aprobacion
            Dim txtNvAp As SAPbouiCOM.EditText = oForm.Items.Item("txtNvAp").Specific 'txt nivel de aprobacion

            Dim lnkPagCon As SAPbouiCOM.LinkedButton = oForm.Items.Item("lnkPagCon").Specific 'link pago consolidado
            Dim txtPagCon As SAPbouiCOM.EditText = oForm.Items.Item("txtPagCon").Specific 'txt id pago consolidado

            Dim lblArcBco As SAPbouiCOM.StaticText = oForm.Items.Item("lblArcBco").Specific 'lvl ruta archivo de banco
            Dim btnABCO As SAPbouiCOM.Button = oForm.Items.Item("btnABCO").Specific 'buttom ruta archivo banco
            Dim lblABCO As SAPbouiCOM.StaticText = oForm.Items.Item("lblABCO").Specific 'lvl ruta archivo de banco
            'Dim txtArcBco As SAPbouiCOM.EditText = oForm.Items.Item("txtArcBco").Specific 'txt ruta de archivo de banco

            Dim lblFArc As SAPbouiCOM.StaticText = oForm.Items.Item("lblFArc").Specific 'ADD 07/11/2024
            Dim txtFArc As SAPbouiCOM.EditText = oForm.Items.Item("txtFArc").Specific 'ADD 07/11/2024

            Dim lblFDev As SAPbouiCOM.StaticText = oForm.Items.Item("lblFDev").Specific 'ADD 07/11/2024
            Dim txtFDev As SAPbouiCOM.EditText = oForm.Items.Item("txtFDev").Specific 'ADD 07/11/2024

            Item_4 = oForm.Items.Item("Item_4").Specific 'lcl Cuenta
            Dim lnkCta As SAPbouiCOM.LinkedButton = oForm.Items.Item("lnkCta").Specific
            Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific 'cbx Cuenta

            Dim lblSCta As SAPbouiCOM.StaticText = oForm.Items.Item("lblSCta").Specific

            Dim Item_6 As SAPbouiCOM.StaticText = oForm.Items.Item("Item_6").Specific 'lbl Tipo forma pago
            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific 'cbx Forma Pago
            Dim lblEst As SAPbouiCOM.StaticText = oForm.Items.Item("lblEst").Specific 'lvl estado solicitud

            Dim lblNumChe As SAPbouiCOM.StaticText = oForm.Items.Item("lblNumChe").Specific

            Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific 'txt estado solicitud

            Dim lblCBan As SAPbouiCOM.StaticText = oForm.Items.Item("lblCBan").Specific 'lbl ID CASH BANCO
            Dim txtCBan As SAPbouiCOM.EditText = oForm.Items.Item("txtCBan").Specific 'txt estado solicitud

            Dim MTX_SER As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific 'Matrix
            Dim oColCombo As SAPbouiCOM.Column = MTX_SER.Columns.Item("cbxProc")
            Dim oColumn2 As SAPbouiCOM.Column = MTX_SER.Columns.Item("U_CtaBco") 'Columna de matrix
            Dim oColumn As SAPbouiCOM.Column = MTX_SER.Columns.Item("U_Pag") 'Columna de matrix
            Dim oColumObser As SAPbouiCOM.Column = MTX_SER.Columns.Item("U_Coment") 'Columna de matrix
            Dim oColumPT As SAPbouiCOM.Column = MTX_SER.Columns.Item("U_IdPT") 'Columna de matrix
            Dim oColumND As SAPbouiCOM.Column = MTX_SER.Columns.Item("U_IdND") 'Columna de matrix

            Dim it2 As SAPbouiCOM.Button = oForm.Items.Item("2").Specific 'btn Cancelar
            Dim lbl As SAPbouiCOM.StaticText = oForm.Items.Item("lbl").Specific 'lbl Facturas
            Dim lblFac As SAPbouiCOM.StaticText = oForm.Items.Item("lblFac").Specific 'lbl Num Facturas

            Dim lblSM As SAPbouiCOM.StaticText = oForm.Items.Item("lblSM").Specific 'lbl total monto sp
            Dim lblSMon As SAPbouiCOM.StaticText = oForm.Items.Item("lblSMon").Specific 'lbl total monto sp
            Dim lblSS As SAPbouiCOM.StaticText = oForm.Items.Item("lblSS").Specific 'lbl total saldo sp
            Dim lblSSal As SAPbouiCOM.StaticText = oForm.Items.Item("lblSSal").Specific 'lbl total saldo

            Dim Item_23 As SAPbouiCOM.StaticText = oForm.Items.Item("Item_23").Specific 'lbl Monto a pagar
            Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific 'txt Monto a pagar
            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            Dim focus As SAPbouiCOM.Item = oForm.Items.Item("focus")

            Dim lblRAG As SAPbouiCOM.StaticText = oForm.Items.Item("lblRAG").Specific

            'Ubicacion de controles por nivel
            If Modo = "Buscar" Or ((CInt(Nivel) >= 1 And CInt(Nivel) < CInt(NivelMaximo)) And Modo = "Modificar") Then
                'Ocultamos la seccion alta del formulario
                Item_8.Item.Visible = False
                txtPro.Item.Visible = False
                lblRaz.Item.Visible = False
                Item_12.Item.Visible = False
                focus.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                txtFecCor.Item.Visible = False
                lblSuc.Item.Visible = False
                cbxSuc.Item.Visible = False
                chkPLSB.Item.Visible = False
                btnfilt.Item.Visible = False
                btnMDT.Item.Visible = False
                btnAddMar.Item.Visible = False
                oGrid.Item.Visible = False

                'Subimos la seccion Baja a la parte alta del formulario
                Item_0.Item.Top = Item_8.Item.Top 'lcl Id solicitud
                txtDocEnt.Item.Top = Item_8.Item.Top
                Item_2.Item.Top = Item_8.Item.Top
                txtNvAp.Item.Top = Item_8.Item.Top
                lnkPagCon.Item.Top = Item_8.Item.Top
                txtPagCon.Item.Top = Item_8.Item.Top
                lblArcBco.Item.Top = Item_8.Item.Top
                btnABCO.Item.Top = Item_8.Item.Top
                lblABCO.Item.Top = Item_8.Item.Top
                'txtArcBco.Item.Top = Item_8.Item.Top

                lblFArc.Item.Top = Item_8.Item.Top
                txtFArc.Item.Top = Item_8.Item.Top
                lblFDev.Item.Top = Item_8.Item.Top
                txtFDev.Item.Top = Item_8.Item.Top

                Item_4.Item.Top = Item_12.Item.Top
                lnkCta.Item.Top = Item_12.Item.Top
                cbxCuenta.Item.Top = Item_12.Item.Top
                cbxCuenta.Item.Enabled = False
                lblSCta.Item.Top = Item_12.Item.Top
                Item_6.Item.Top = Item_12.Item.Top
                cbxfpago.Item.Top = Item_12.Item.Top
                cbxfpago.Item.Enabled = False
                lblNumChe.Item.Top = Item_12.Item.Top
                lblCBan.Item.Top = Item_12.Item.Top
                txtCBan.Item.Top = Item_12.Item.Top
                lblEst.Item.Top = Item_12.Item.Top
                txtEstado.Item.Top = Item_12.Item.Top
                btnAprob.Item.Top = Item_12.Item.Top
                MTX_SER.Item.Top = oGrid.Item.Top

                'Posicion de botones
                btnSolPag.Item.Top = referenciaBotones
                it2.Item.Top = referenciaBotones
                btnND.Item.Top = referenciaBotones
                btnimpgs.Item.Top = referenciaBotones
                btncashm.Item.Top = referenciaBotones
                lbl.Item.Top = referenciaBotones
                lblFac.Item.Top = referenciaBotones

                lblSM.Item.Top = referenciaBotones
                lblSMon.Item.Top = referenciaBotones
                lblSS.Item.Top = referenciaBotones
                lblSSal.Item.Top = referenciaBotones

                Item_23.Item.Top = referenciaBotones
                txtMP.Item.Top = referenciaBotones
                ipLogoSS.Item.Top = referenciaBotones + 18

                lblRAG.Item.Top = referenciaBotones + 25

                oForm.Height = btnSolPag.Item.Top + 90 'Recortamos el tamaño del formulario
            End If

            txtCBan.Item.Enabled = False
            txtFArc.Item.Enabled = False

            btnAprob.Item.Visible = False
            oColCombo.Visible = False
            oColumObser.Visible = False
            oColumPT.Visible = False
            oColumND.Visible = False
            lnkPagCon.Item.Visible = False
            txtPagCon.Item.Visible = False

            oColumn.Editable = False
            oColumn2.Editable = False
            btnSolPag.Item.Enabled = False
            btnND.Item.Enabled = False
            btnimpgs.Item.Enabled = False
            btncashm.Item.Enabled = False

            'Habilito botones por nivel
            If (CInt(Nivel) = 0 Or CInt(Nivel) = CInt(NivelMaximo)) And Modo = "Modificar" Then
                oColumn.Editable = True
                oColumn2.Editable = True
                btnSolPag.Item.Enabled = True
            ElseIf CInt(Nivel) >= 1 And CInt(Nivel) <= CInt(NivelMaximo) And Modo = "Buscar" Then
                btnAprob.Item.Visible = True
                btnAprob.Item.Enabled = True
            End If

            If txtEstado.Value = "Aprobado" And (CInt(Nivel) = 0 Or CInt(Nivel) = CInt(NivelMaximo)) Then ' And (CInt(Nivel) = 0 Or CInt(Nivel) = CInt(NivelMaximo)) Then
                btnAprob.Item.Visible = True
                btncashm.Item.Enabled = True
                btnimpgs.Item.Enabled = True 'Add 16/01/2025
                txtFArc.Item.Enabled = True

                'ElseIf txtEstado.Value = "Aprobado" And CInt(Nivel) = CInt(NivelMaximo) Then 'NEW
                '    btnAprob.Item.Visible = True 'NEW
                '    btncashm.Item.Enabled = True
            ElseIf txtEstado.Value = "Rechazado" Or txtEstado.Value = "Procesado" Then
                btnAprob.Item.Visible = False
            ElseIf txtEstado.Value = "Revision" And CInt(Nivel) = 0 Then 'add 23/09/2024
                btnAprob.Item.Visible = True '23/09/2024
                btnAprob.Item.Enabled = True
                btnimpgs.Item.Enabled = True 'Add 21/01/2025
            ElseIf txtEstado.Value = "Archivo Generado" Then
                txtCBan.Item.Enabled = True
                btncashm.Item.Enabled = False
                btnAprob.Item.Visible = True 'ojo
                txtFArc.Item.Enabled = True
                btnimpgs.Item.Enabled = True 'Add 15/01/2025
            ElseIf txtEstado.Value = "Archivo Procesado Banco" Then
                btnimpgs.Item.Enabled = True
                oColCombo.Visible = True
                btnND.Item.Enabled = True 'mod 25/10/2024
                btnAprob.Item.Visible = False
                oColumObser.Visible = True
                lnkPagCon.Item.Visible = True
                txtPagCon.Item.Visible = True
                btnABCO.Item.Enabled = True
                'txtArcBco.Item.Enabled = True
                oColumPT.Visible = True
                oColumND.Visible = True
                oColumPT.Editable = False
                oColumND.Editable = False
                txtFDev.Item.Enabled = True

                txtCBan.Item.Enabled = False
                txtFArc.Item.Enabled = False

            End If

            focus.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error BloqueaHabilitaBotones: " & ex.Message.ToString, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error BloqueaHabilitaBotones: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub InicializaControlesXNivel(ByVal NivelUsr As String, Optional ByVal Modo As String = "", Optional tipopago As String = "")
        Dim ErrorSeccion As String = ""
        Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific 'txt estado solicitud
        Try
            If (CInt(NivelUsr) = 0 Or CInt(NivelUsr) = CInt(NivelMaximo)) And Modo = "Modificar" Then

                oCFLs = oForm.ChooseFromLists
                oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "2"
                oCFLCreationParams.UniqueID = "CFL1"
                oCFL = oCFLs.Add(oCFLCreationParams)
                oCons = oCFL.GetConditions()

                oCon = oCons.Add()
                oCon.Alias = "CardType"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "S"
                oCFL.SetConditions(oCons)

                txtPro = oForm.Items.Item("txtPro").Specific
                oForm.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                txtPro.DataBind.SetBound(True, "", "EditDS")
                txtPro.ChooseFromListUID = "CFL1"
                txtPro.ChooseFromListAlias = "CardCode"

                Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing

                'Sucursal
                ErrorSeccion = "Obteniendo Sucursal"
                cbxSuc = oForm.Items.Item("cbxSuc").Specific
                Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rst.DoQuery("SELECT ""PrcCode"", ""PrcName"" FROM ""OPRC"" WHERE ""DimCode"" = 3 AND ""Locked"" <> 'Y'")
                ValoresValidos = cbxSuc.ValidValues

                ValoresValidos.Add("Todos", "Todos")
                If rst.RecordCount >= 1 Then
                    While (rst.EoF = False)
                        ValoresValidos.Add(rst.Fields.Item("PrcCode").Value, rst.Fields.Item("PrcName").Value.ToString)
                        rst.MoveNext()
                    End While
                End If

                cbxSuc.Select("Todos", SAPbouiCOM.BoSearchKey.psk_ByValue)

                'Cuenta
                ErrorSeccion = "Obteniendo Cuentas"
                LLenaComboCuenta()

                ErrorSeccion = "Creando tabla dtDocs"
                oForm.DataSources.DataTables.Add("dtDocs")

                cbxfpago = oForm.Items.Item("cbxfpago").Specific
                If cbxfpago.Value = "" Then
                    cbxfpago.ValidValues.Add("Transferencia", "Transferencia")
                    cbxfpago.ValidValues.Add("Cheque", "Cheque")
                    cbxfpago.ValidValues.Add("Servicios Basicos", "Servicios Basicos")
                    cbxfpago.Select("Transferencia")
                End If

                CargarDatos()

            ElseIf (CInt(NivelUsr) >= 1 And CInt(NivelUsr) <= CInt(NivelMaximo)) And txtEstado.Value = "Revision" Then
                btnAprob = oForm.Items.Item("btnAprob").Specific
                btnAprob.Item.AffectsFormMode = False
                btnAprob.ValidValues.Add("Aprobado", "Aprobado")
                btnAprob.ValidValues.Add("Rechazado", "Rechazado")
                'btnAprob.ValidValues.Add("Edicion", "Edicion")
                btnAprob.Select("Aprobado")

            ElseIf ((CInt(NivelUsr) = CInt(NivelMaximo)) And (txtEstado.Value = "Aprobado" Or txtEstado.Value = "Archivo Generado")) Or ((CInt(NivelUsr) = 0) And (txtEstado.Value = "Aprobado" Or txtEstado.Value = "Archivo Generado")) Then
                btnAprob = oForm.Items.Item("btnAprob").Specific
                btnAprob.Item.AffectsFormMode = False
                cbxfpago = oForm.Items.Item("cbxfpago").Specific

                If tipopago = "Transferencia" Then
                    btnAprob.ValidValues.Add("Modificar", "Modificar")
                    btnAprob.ValidValues.Add("Archivo Procesado Banco", "Archivo Procesado Banco")
                ElseIf tipopago = "Cheque" Or tipopago = "Servicios Basicos" Then
                    btnAprob.ValidValues.Add("Modificar", "Modificar")
                    btnAprob.ValidValues.Add("Procesar", "Procesar")
                End If
                btnAprob.Select("Modificar")


            ElseIf CInt(NivelUsr) = 0 And txtEstado.Value = "Revision" And Modo = "Buscar" Then
                btnAprob = oForm.Items.Item("btnAprob").Specific
                btnAprob.Item.AffectsFormMode = False
                btnAprob.ValidValues.Add("Modificar", "Modificar")
                btnAprob.Select("Modificar")

            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error InicializaControlesXNivel seccion: {ErrorSeccion} : {ex.Message}", "frmPagosMasivos")
            rsboApp.StatusBar.SetText($"Error InicializaControlesXNivel seccion: {ErrorSeccion} : {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LLenaComboCuenta()
        Try
            'cbxCuenta = oForm.Items.Item("cbxCuenta").Specific
            'Query = ""
            'Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing

            'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '    Query = "SELECT B.""GLAccount"", A.""AcctName"", B.""BankCode"" FROM ""OACT"" A INNER JOIN ""DSC1"" B ON A.""AcctCode"" = B.""GLAccount"""
            'Else
            '    Query = "SELECT B.GLAccount, A.AcctName, B.BankCode FROM OACT A WITH(NOLOCK) INNER JOIN DSC1 B WITH(NOLOCK) ON A.AcctCode = B.GLAccount"
            'End If

            'rst.DoQuery(Query)
            'ValoresValidos = cbxCuenta.ValidValues

            'If rst.RecordCount >= 1 Then
            '    While Not rst.EoF
            '        If cbxCuenta.Value = "" Then
            '            ValoresValidos.Add(rst.Fields.Item("GLAccount").Value, rst.Fields.Item("BankCode").Value.ToString & ":" & rst.Fields.Item("AcctName").Value)
            '            rst.MoveNext()
            '        Else
            '            If cbxCuenta.Value <> rst.Fields.Item("GLAccount").Value Then ValoresValidos.Add(rst.Fields.Item("GLAccount").Value, rst.Fields.Item("BankCode").Value.ToString & ":" & rst.Fields.Item("AcctName").Value)
            '            rst.MoveNext()
            '        End If
            '    End While
            'End If

            cbxCuenta = oForm.Items.Item("cbxCuenta").Specific
            Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing
            Dim oRecordSet As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT * FROM ""DSC1""")

            Dim CtaBcoPropio As SAPbobsCOM.HouseBankAccounts = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oHouseBankAccounts)
            CtaBcoPropio.Browser.Recordset = oRecordSet
            CtaBcoPropio.Browser.MoveFirst()
            ValoresValidos = cbxCuenta.ValidValues

            If CtaBcoPropio.Browser.RecordCount >= 1 Then
                Do While Not CtaBcoPropio.Browser.EoF
                    If cbxCuenta.Value = "" Then
                        ValoresValidos.Add(CtaBcoPropio.GLAccount, CtaBcoPropio.BankCode & ":" & CtaBcoPropio.AccountName)
                        CtaBcoPropio.Browser.MoveNext()
                    Else
                        If cbxCuenta.Value <> CtaBcoPropio.GLAccount Then ValoresValidos.Add(CtaBcoPropio.GLAccount, CtaBcoPropio.BankCode & ":" & CtaBcoPropio.AccountName)
                        CtaBcoPropio.Browser.MoveNext()
                    End If
                Loop
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error llenando combo cuenta: {ex.Message}", "frmPagosMasivos")
            rsboApp.StatusBar.SetText($"Error llenando combo cuenta: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub CargarDatos()

        txtPro = oForm.Items.Item("txtPro").Specific
        txtFecCor = oForm.Items.Item("txtFecCor").Specific
        cbxSuc = oForm.Items.Item("cbxSuc").Specific
        TotalFacts = 0
        Dim sucursal As String = ""
        If cbxSuc.Value <> "Todos" Then sucursal = cbxSuc.Value

        Query = ""
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            Query = "CALL " & rCompany.CompanyDB & ".SS_PAGOS_MASIVOS ('" & txtPro.Value & "','" & txtFecCor.Value & "','" & sucursal & "')"
        Else
            Query = "EXEC SS_PAGOS_MASIVOS '" & txtPro.Value & "','" & txtFecCor.Value & "', '" & sucursal & "'"
        End If

        Try
            rsboApp.StatusBar.SetText("Consultando Facturas, espere por favor! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + Query, "frmPagosMasivos")
            oGrid.DataTable.ExecuteQuery(Query)
            Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + Query, "frmPagosMasivos")

            FormatoFacturas()

            Dim AcumTotal As Decimal = 0, AcumSal As Decimal = 0
            Dim j As Integer = 0
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsLeaf(i) = True Then
                    AcumTotal += CDec(oGrid.DataTable.GetValue("DocTotal", j))
                    AcumSal += CDec(oGrid.DataTable.GetValue("saldo", j))

                    'Cambio de color la linea si no encuentra banco ni cuenta
                    'If oGrid.DataTable.GetValue("Banco", j) = "" Or oGrid.DataTable.GetValue("Cuenta", j) = "" Then oGrid.CommonSetting.SetRowBackColor(i, RGB(255, 255, 0))

                    j += 1
                End If
            Next

            oGrid.DataTable.Rows.Add(1) ' Añade una fila al final
            oGrid.DataTable.SetValue("DocTotal", oGrid.DataTable.Rows.Count - 1, CDbl(AcumTotal)) ' Coloca el total en la columna "Monto" de la última fila
            oGrid.DataTable.SetValue("saldo", oGrid.DataTable.Rows.Count - 1, CDbl(AcumSal)) ' Coloca el total en la columna "Monto" de la última fila
            oGrid.CommonSetting.SetRowBackColor(2, RGB(255, 128, 0))
            oGrid.CommonSetting.SetCellEditable(2, 2, False)

            rsboApp.StatusBar.SetText("Consulta terminada con éxito...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Dim focus As SAPbouiCOM.Item = oForm.Items.Item("focus")
            focus.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar cargar datos: " & ex.Message.ToString, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error al ejecutar cargar datos:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmPagosMasivos" Then

                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal

                        If oCFLEvento.BeforeAction = False Then
                            Dim sCFL_ID As String = oCFLEvento.ChooseFromListUID
                            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmPagosMasivos")
                            Dim oCFL As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty

                            If Not oDataTable Is Nothing Then
                                val = oDataTable.GetValue(0, 0)
                                val1 = oDataTable.GetValue(1, 0)
                                Try
                                    oUserDataSource = oForm.DataSources.UserDataSources.Item("EditDS")
                                    oUserDataSource.ValueEx = val
                                Catch ex As Exception
                                End Try
                                Try
                                    Dim lblRaz As SAPbouiCOM.StaticText = oForm.Items.Item("lblRaz").Specific
                                    lblRaz.Caption = val1
                                Catch ex As Exception
                                End Try
                            Else
                                Dim lblRaz As SAPbouiCOM.StaticText = oForm.Items.Item("lblRaz").Specific
                                lblRaz.Caption = ""
                            End If
                        Else

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK

                        If Not pVal.BeforeAction Then

                            Select Case pVal.ItemUID

                                Case "oGrid"
                                    Try
                                        If pVal.ColUID = "Chek" Then
                                            If pVal.Row > 0 Then
                                                Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                                Dim IdRowDataTable As Integer = oGrid.GetDataTableRowIndex(pVal.Row)
                                                Dim isChecked As String = oGrid.DataTable.GetValue("Chek", IdRowDataTable)
                                                If isChecked = "Y" Then
                                                    If Not ListadeLineas.Contains(IdRowDataTable) Then ListadeLineas.Add(IdRowDataTable)
                                                Else
                                                    If ListadeLineas.Contains(IdRowDataTable) Then ListadeLineas.Remove(IdRowDataTable)
                                                End If
                                            End If
                                        End If
                                    Catch ex As Exception
                                    End Try

                                Case "btnfilt"
                                    oForm.Freeze(True)
                                    CargarDatos()
                                    oForm.Freeze(False)

                                Case "btnMDT"

                                    MarcaDesmarcaTodo()

                                Case "btnAddMar"

                                    AgregarMarcado()

                                Case "cbxCuenta"

                                    Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific
                                    Dim lblSCta As SAPbouiCOM.StaticText = oForm.Items.Item("lblSCta").Specific
                                    Dim ObjCta As SAPbobsCOM.ChartOfAccounts = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                                    Dim cta As String = cbxCuenta.Value
                                    If ObjCta.GetByKey(cta) Then lblSCta.Caption = CStr(ObjCta.Balance)

                                Case "cbxfpago"
                                    oForm.Freeze(True)
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
                                    Dim colCtaBco As SAPbouiCOM.Column = oMatrix.Columns.Item("U_CtaBco")
                                    Dim ColConsolidado As SAPbouiCOM.Column = oMatrix.Columns.Item("U_Consoli")
                                    Dim colChe As SAPbouiCOM.Column = oMatrix.Columns.Item("U_NumChe")
                                    Dim lblNumChe As SAPbouiCOM.StaticText = oForm.Items.Item("lblNumChe").Specific
                                    Query = ""
                                    Dim result As String = ""
                                    Select Case cbxfpago.Value
                                        Case "Transferencia"
                                            colCtaBco.Visible = True
                                            colChe.Visible = False
                                            ColConsolidado.Visible = False
                                            lblNumChe.Caption = ""
                                        Case "Cheque"
                                            ColConsolidado.Visible = True
                                            colCtaBco.Visible = False

                                            If cbxCuenta.Value <> "" Then
                                                Dim separadorCuentaBanco As String() = cbxCuenta.Selected.Description.Split(":")

                                                'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                '    Query = "SELECT IFNULL(""U_NumChe"",0) AS ""NumCheque"" FROM ""@SS_PM_CONTROLCHEQUE"" WHERE ""Code"" = '" & separadorCuentaBanco(0) & "' AND ""Name"" = '" & cbxCuenta.Value & "'"
                                                'Else
                                                '    Query = "SELECT ISNULL(U_NumChe,0) AS NumCheque FROM ""@SS_PM_CONTROLCHEQUE"" WITH(NOLOCK) WHERE Code = '" & separadorCuentaBanco(0) & "' AND Name = '" & cbxCuenta.Value & "'"
                                                'End If
                                                'result = oFuncionesAddon.getRSvalue(Query, "NumCheque", "0")
                                                'lblNumChe.Caption = result.ToString

                                                Dim oRecordSet As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oRecordSet.DoQuery("SELECT * FROM ""DSC1"" WHERE ""BankCode"" = '" & separadorCuentaBanco(0) & "' AND ""GLAccount"" = '" & cbxCuenta.Value & "'")
                                                Dim CtaBcoPropio As SAPbobsCOM.HouseBankAccounts = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oHouseBankAccounts)
                                                CtaBcoPropio.Browser.Recordset = oRecordSet
                                                lblNumChe.Caption = CtaBcoPropio.NextCheckNo
                                            End If
                                    End Select
                                    oForm.Freeze(False)

                                Case "lblRAG"
                                    Try
                                        Dim lblRAG As SAPbouiCOM.StaticText = oForm.Items.Item("lblRAG").Specific
                                        If lblRAG.Caption <> "" Then
                                            If Directory.Exists(Path.GetDirectoryName(lblRAG.Caption)) Then
                                                Process.Start("explorer.exe", Path.GetDirectoryName(lblRAG.Caption))
                                            Else
                                                rsboApp.StatusBar.SetText("No se encontró el archivo ni el directorio: " & lblRAG.Caption, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        rsboApp.StatusBar.SetText("Error abriendo directorio de archivo banco: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try

                                Case "btnABCO"
                                    oForm.Freeze(True)
                                    Dim selectFileDialog As New SelectFileDialog("C:\", "", "Archivos de Excel (*.xls)|*.xls|Todos los archivos (*.*)|*.*", DialogType.OPEN)
                                    selectFileDialog.Open() '                               
                                    If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                                        Dim ruta As String = ""
                                        ruta = selectFileDialog.SelectedFile
                                        Dim lblABCO As SAPbouiCOM.StaticText = oForm.Items.Item("lblABCO").Specific
                                        lblABCO.Caption = ruta.ToString
                                    End If
                                    oForm.Freeze(False)

                                Case "lblABCO"
                                    Try
                                        Dim lblABCO As SAPbouiCOM.StaticText = oForm.Items.Item("lblABCO").Specific
                                        If lblABCO.Caption <> "" Then
                                            If Directory.Exists(Path.GetDirectoryName(lblABCO.Caption)) Then
                                                Process.Start("explorer.exe", Path.GetDirectoryName(lblABCO.Caption))
                                            Else
                                                rsboApp.StatusBar.SetText("No se encontró el archivo ni el directorio: " & lblABCO.Caption, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        rsboApp.StatusBar.SetText("Error abriendo directorio de archivo banco: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try
                            End Select
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "btnAprob" And pVal.BeforeAction = True Then
                            Dim btnAprob As SAPbouiCOM.ButtonCombo = oForm.Items.Item("btnAprob").Specific
                            Dim txtCBan As SAPbouiCOM.EditText = oForm.Items.Item("txtCBan").Specific
                            Dim txtFArc As SAPbouiCOM.EditText = oForm.Items.Item("txtFArc").Specific
                            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
                            Dim iReturnValue As Integer = 0

                            If btnAprob.Caption = "Modificar" And txtCBan.Value <> "" And txtFArc.Value <> "" Then
                                rsboApp.MessageBox("No puede modificar la solicitud teniendo un ID de banco y fecha de archivo!")
                            ElseIf btnAprob.Caption = "Archivo Procesado Banco" And txtCBan.Value = "" And txtFArc.Value = "" Then
                                rsboApp.MessageBox("No puede procesar los pagos sin tener un ID de banco y la fecha del archivo recibido!")
                            ElseIf (btnAprob.Caption = "Modificar" And txtCBan.Value = "" And txtFArc.Value = "") Or btnAprob.Caption = "Aprobado" Or btnAprob.Caption = "Rechazado" Then

                                iReturnValue = rsboApp.MessageBox($"La acción '{btnAprob.Caption.ToString.ToUpper}' se llevará a cabo. ¿Desea continuar?", 1, "&Sí", "&No")

                                If iReturnValue = 1 Then
                                    If ActualizadoSolicitudDePago(DocEntry, NivelAprobacion, btnAprob.Caption, txtCBan.Value) Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        oForm.Close()
                                        rsboApp.Forms.Item("frmPagosAprobacion").Freeze(True)
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmPagosAprobacion").DataSources.DataTables.Item("dtDocs")
                                        odt.Rows.Remove(Linea)
                                        rsboApp.Forms.Item("frmPagosAprobacion").Freeze(False)
                                    End If
                                Else
                                    rsboApp.StatusBar.SetText($"No se realizo la acción {btnAprob.Caption.ToString.ToUpper}!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                            ElseIf (btnAprob.Caption = "Archivo Procesado Banco" And txtCBan.Value <> "" And txtFArc.Value <> "") Or btnAprob.Caption = "Procesar" Then

                                iReturnValue = rsboApp.MessageBox($"La acción '{btnAprob.Caption.ToString.ToUpper}' se llevará a cabo. ¿Desea continuar?", 1, "&Sí", "&No")
                                If iReturnValue = 1 Then
                                    If CrearPagoMasivo() Then
                                        Dim btnND As SAPbouiCOM.Button = oForm.Items.Item("btnND").Specific
                                        btnND.Item.Enabled = True
                                        Dim btnimpgs As SAPbouiCOM.Button = oForm.Items.Item("btnimpgs").Specific
                                        btnimpgs.Item.Enabled = True
                                        Dim btncashm As SAPbouiCOM.Button = oForm.Items.Item("btncashm").Specific
                                        btncashm.Item.Enabled = False
                                        Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific
                                        txtEstado.Value = "Archivo Procesado Banco" 'revisar

                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        oForm.Close()
                                        rsboApp.Forms.Item("frmPagosAprobacion").Freeze(True)
                                        Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmPagosAprobacion").DataSources.DataTables.Item("dtDocs")
                                        odt.Rows.Remove(Linea)
                                        rsboApp.Forms.Item("frmPagosAprobacion").Freeze(False)
                                    End If
                                Else
                                    rsboApp.StatusBar.SetText($"No se realizo la acción {btnAprob.Caption.ToString.ToUpper}!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If
                            End If
                        End If

                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "btnSolPag"

                                    Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific

                                    If cbxCuenta.Value <> "" Then
                                        Dim DocEntryFacturaRecibida_UDO As String = ""

                                        If DocEntry <> "" Then
                                            If EdicionSolicitudPago(DocEntry) Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                                oForm.Close()

                                                rsboApp.Forms.Item("frmPagosAprobacion").Freeze(True)
                                                Dim odt As SAPbouiCOM.DataTable = rsboApp.Forms.Item("frmPagosAprobacion").DataSources.DataTables.Item("dtDocs")
                                                odt.Rows.Remove(Linea)
                                                rsboApp.Forms.Item("frmPagosAprobacion").Freeze(False)
                                            End If
                                        Else
                                            If CrearSolicitudPago(DocEntryFacturaRecibida_UDO) Then
                                                oForm.Freeze(True)
                                                CargarDatos()

                                                Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
                                                mMatrix.Clear()

                                                Dim lblFac As SAPbouiCOM.StaticText = oForm.Items.Item("lblFac").Specific
                                                lblFac.Caption = ""

                                                Dim lblSMon As SAPbouiCOM.StaticText = oForm.Items.Item("lblSMon").Specific
                                                lblSMon.Caption = ""
                                                Dim lblSSal As SAPbouiCOM.StaticText = oForm.Items.Item("lblSSal").Specific
                                                lblSSal.Caption = ""

                                                Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific
                                                txtMP.Value = ""

                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                                                oForm.Freeze(False)
                                            End If
                                        End If
                                    Else
                                        rsboApp.StatusBar.SetText(NombreAddon + " - Seleccione una cuenta...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                Case "btnND"
                                    Try
                                        Dim txtFDev As SAPbouiCOM.EditText = oForm.Items.Item("txtFDev").Specific
                                        Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific

                                        If txtEstado.Value = "Archivo Procesado Banco" Then
                                            If txtFDev.Value <> "" Then
                                                If CrearND() Then
                                                    Dim btnND As SAPbouiCOM.Button = oForm.Items.Item("btnND").Specific
                                                    btnND.Item.Enabled = False
                                                End If
                                            Else
                                                rsboApp.StatusBar.SetText("Ingrese una fecha de devolución por favor!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        rsboApp.StatusBar.SetText("Error boton de devolucion: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try

                                Case "btncashm"
                                    Try
                                        Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific

                                        If Not String.IsNullOrEmpty(txtDocEnt.Value.ToString) Then
                                            Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific 'cbx Cuenta
                                            Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific

                                            If txtEstado.Value.ToString = "Aprobado" Then
                                                Dim btncashm As SAPbouiCOM.Button = oForm.Items.Item("btncashm").Specific
                                                Dim txtCBan As SAPbouiCOM.EditText = oForm.Items.Item("txtCBan").Specific
                                                Dim txtFArc As SAPbouiCOM.EditText = oForm.Items.Item("txtFArc").Specific
                                                If GeneraArchivo(txtDocEnt.Value, cbxCuenta) Then
                                                    txtEstado.Value = "Archivo Generado"
                                                    btncashm.Item.Enabled = False
                                                    txtCBan.Item.Enabled = True
                                                    txtFArc.Item.Enabled = True
                                                    btnAprob = oForm.Items.Item("btnAprob").Specific
                                                    btnAprob.Item.AffectsFormMode = False

                                                    oForm.Items.Item("btnimpgs").Enabled = True

                                                    If RutaDeArchivoGenerado <> "" Then
                                                        Dim nombreArchivo As String = Path.GetFileName(RutaDeArchivoGenerado)
                                                        Dim nuevoNombre As String = "ANULADO_" & nombreArchivo
                                                        Try
                                                            File.Move(RutaDeArchivoGenerado, nuevoNombre)
                                                        Catch ex As Exception
                                                            rsboApp.StatusBar.SetText("Error actualizando nombre del archivo anterior: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        End Try
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Catch ex As Exception
                                        rsboApp.MessageBox("Error boton btncashm: " + ex.Message.ToString())
                                    End Try

                                Case "btnimpgs"
                                    Try
                                        Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific

                                        If Not String.IsNullOrEmpty(txtDocEnt.Value.ToString) Then
                                            Imprimir(txtDocEnt.Value)
                                        End If

                                    Catch ex As Exception
                                        rsboApp.MessageBox("Error boton btnimpgs (Imprimir pagos): " + ex.Message.ToString())
                                    End Try
                            End Select
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If Not pVal.Before_Action Then
                            If pVal.ItemUID = "MTX_SER" AndAlso pVal.ColUID = "U_Pag" Then
                                oForm.Freeze(True)
                                Try
                                    Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific
                                    Dim btnSolPag As SAPbouiCOM.Button = oForm.Items.Item("btnSolPag").Specific
                                    Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
                                    Dim CantidadDeErrores As Decimal = 0
                                    Dim TotalPagar As Decimal = 0
                                    For i As Integer = 1 To mMatrix.RowCount
                                        Dim Saldo As Decimal = Convert.ToDecimal(mMatrix.Columns.Item("U_Sal").Cells.Item(i).Specific.Value.ToString()) '.Replace(".", ","))
                                        Dim Pagar As Decimal = Convert.ToDecimal(IIf(String.IsNullOrEmpty(mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value.ToString()), "0", mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value.ToString())) '.Replace(".", ",")))

                                        If Pagar > Saldo Then
                                            mMatrix.CommonSetting.SetCellBackColor(i, 14, ColorTranslator.ToOle(Color.Red))
                                            CantidadDeErrores += 1
                                        ElseIf Pagar > 0 And Pagar <= Saldo Then
                                            mMatrix.CommonSetting.SetCellBackColor(i, 14, ColorTranslator.ToOle(Color.LightGreen))
                                        ElseIf Pagar = 0 Then
                                            mMatrix.CommonSetting.SetCellBackColor(i, 14, ColorTranslator.ToOle(Color.LightGreen))
                                            CantidadDeErrores += 1
                                        End If
                                        TotalPagar += Convert.ToDecimal(IIf(String.IsNullOrEmpty(mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value.ToString()), "0", mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value.ToString())) '.Replace(".", ","))) 
                                    Next

                                    txtMP.Value = Math.Round(TotalPagar, 2).ToString

                                    If CantidadDeErrores >= 1 Then
                                        btnSolPag.Item.Enabled = False
                                    Else
                                        btnSolPag.Item.Enabled = True
                                    End If

                                    oForm.Freeze(False)
                                Catch ex As Exception
                                    rsboApp.StatusBar.SetText("Error evento Key_Down " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pVal.BeforeAction Then Event_MatrixLinkPressed(pVal)


                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        Try
                            If Not pVal.BeforeAction Then
                                AdaptarTamano()
                            End If
                        Catch ex As Exception
                        End Try
                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub AdaptarTamano()
        Try
            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Item = oForm.Items.Item("oGrid")
            Dim MTX_SER As SAPbouiCOM.Item = oForm.Items.Item("MTX_SER")
            Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific
            If DocEntry = "" Then
                'MTX_SER.Height = 198
                oGrid.Height = 193
                TryCast(oGrid.Specific, SAPbouiCOM.Grid).AutoResizeColumns()
            Else
                If txtEstado.Value = "Modificar" Then
                    'MTX_SER.Height = 198
                    oGrid.Height = 193
                    TryCast(oGrid.Specific, SAPbouiCOM.Grid).AutoResizeColumns()
                End If
            End If
            TryCast(MTX_SER.Specific, SAPbouiCOM.Matrix).AutoResizeColumns()
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            rsboApp.SetStatusBarMessage("Error al Ajustar Dimension Matrix" & ex.Message)
        End Try
    End Sub

    Private Function ObtenerUltimoCheck(ByVal fila As Integer) As Integer
        Try
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim valor As Integer = 0
            Dim Check As Boolean
            For i As Integer = 1 To fila - 1
                Check = CType(oMatrix.Columns.Item("U_Consoli").Cells.Item(i).Specific, SAPbouiCOM.CheckBox).Checked
                If Check Then valor = CInt(oMatrix.Columns.Item("U_NumChe").Cells.Item(i).Specific.String)
            Next

            Return valor
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent
        Try
            If eventInfo.FormUID = "frmPagosMasivos" Then
                If eventInfo.ItemUID = "MTX_SER" Then
                    If eventInfo.ColUID = "U_DocEntry" Then
                        If Not rsboApp.Menus.Exists("FrmPagMas") Then
                            Dim oMenus As SAPbouiCOM.Menus = rsboApp.Menus
                            Dim oMenuItem As SAPbouiCOM.MenuItem = oMenus.Item("1280") ' ID del menú de clic derecho
                            Dim oMenuParams As SAPbouiCOM.MenuCreationParams = CType(rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams), SAPbouiCOM.MenuCreationParams)
                            num = eventInfo.Row
                            oMenuParams.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oMenuParams.UniqueID = "FrmPagMas"
                            oMenuParams.String = "Eliminar Fila"
                            oMenuParams.Enabled = True
                            oMenuItem.SubMenus.AddEx(oMenuParams)
                        Else
                            If rsboApp.Menus.Exists("FrmPagMas") Then rsboApp.Menus.RemoveEx("FrmPagMas")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            If pVal.MenuUID = "FrmPagMas" And Not pVal.BeforeAction Then
                Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
                mMatrix.DeleteRow(num)

                Dim lblFac As SAPbouiCOM.StaticText = oForm.Items.Item("lblFac").Specific
                Dim lblSMon As SAPbouiCOM.StaticText = oForm.Items.Item("lblSMon").Specific
                Dim lblSSal As SAPbouiCOM.StaticText = oForm.Items.Item("lblSSal").Specific
                Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific

                Dim Dmon As Double = 0, Dsal = 0, Dpag = 0
                Dim DCan As Integer = 0

                For i As Integer = 1 To mMatrix.RowCount
                    Dim U_Mon As String = mMatrix.Columns.Item("U_Mon").Cells.Item(i).Specific.Value
                    Dmon += CDbl(U_Mon)
                    Dim U_Sal As String = mMatrix.Columns.Item("U_Sal").Cells.Item(i).Specific.Value
                    Dsal += CDbl(U_Sal)
                    Dim U_Pag As String = mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value
                    Dpag += CDbl(U_Pag)
                    DCan += 1
                Next

                lblFac.Caption = DCan.ToString
                lblSMon.Caption = Dmon.ToString
                lblSSal.Caption = Dsal.ToString
                txtMP.Value = Dpag

                ' If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmPagosMasivos")
        End Try
    End Sub

    Private Sub MarcaDesmarcaTodo()
        Try
            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            Dim j As Integer = 0
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.Rows.IsLeaf(i) = True Then
                    Dim isChecked As String = oGrid.DataTable.GetValue("Chek", j)
                    If isChecked = "N" Then
                        If Not ListadeLineas.Contains(j) Then
                            oGrid.DataTable.SetValue("Chek", j, "Y")
                            ListadeLineas.Add(j)
                        End If
                    Else
                        If ListadeLineas.Contains(j) Then
                            oGrid.DataTable.SetValue("Chek", j, "N")
                            ListadeLineas.Remove(j)
                        End If
                    End If
                    j += 1
                End If
            Next
            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error MarcaDesmarcaTodo " & ex.Message.ToString, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error MarcaDesmarcaTodo " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AgregarMarcado()
        Try
            If ListadeLineas.Count = 0 Then
                rsboApp.StatusBar.SetText("No hay facturas seleccionadas... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            oForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_PM_DET1") 'Add Jp 17/12/2024

            TotalFacts = 0
            totalMontoSP = 0
            totalSaldoSP = 0
            ListadeLineas.Sort()
            Dim LineasSinCuenta As List(Of Integer) = Nothing
            LineasSinCuenta = New List(Of Integer)
            Dim i As Integer
            Dim j As Integer = 0
            For Each i In ListadeLineas
                If (oGrid.DataTable.GetValue("Banco", i)) = "" And (oGrid.DataTable.GetValue("Cuenta", i)) = "" And cbxfpago.Value.ToString = "Transferencia" Then
                    rsboApp.MessageBox($"El proveedor {(oGrid.DataTable.GetValue("CardCode", i))} - {(oGrid.DataTable.GetValue("CardName", i))} no tiene una cuenta bancaria registrada!")
                    LineasSinCuenta.Add(i)
                Else
                    oDataSource.InsertRecord(j)
                    oDataSource.SetValue("U_DocEntry", j, (oGrid.DataTable.GetValue("DocEntry", i)))
                    oDataSource.SetValue("U_NumDoc", j, (oGrid.DataTable.GetValue("DocNum", i)))
                    oDataSource.SetValue("U_CodProv", j, (oGrid.DataTable.GetValue("CardCode", i)))
                    oDataSource.SetValue("U_CtaBcoPr", j, (oGrid.DataTable.GetValue("Cuenta", i)))
                    oDataSource.SetValue("U_Proveedor", j, (oGrid.DataTable.GetValue("CardName", i)))
                    oDataSource.SetValue("U_Proyecto", j, (oGrid.DataTable.GetValue("Proyecto", i)))
                    oDataSource.SetValue("U_Sucursal", j, (oGrid.DataTable.GetValue("Sucursal", i)))
                    oDataSource.SetValue("U_Cuota", j, (oGrid.DataTable.GetValue("Cuota", i)))
                    oDataSource.SetValue("U_Vencimiento", j, (oGrid.DataTable.GetValue("Vencimiento", i)))
                    Dim fechaString As String = oGrid.DataTable.GetValue("DocDueDate", i)
                    oDataSource.SetValue("U_FechaVen", j, Convert.ToDateTime(fechaString).ToString("yyyyMMdd"))
                    oDataSource.SetValue("U_Monto", j, (oGrid.DataTable.GetValue("DocTotal", i)))
                    oDataSource.SetValue("U_Saldo", j, (oGrid.DataTable.GetValue("saldo", i)))
                    oDataSource.SetValue("U_Pago", j, (oGrid.DataTable.GetValue("saldo", i)))
                    oDataSource.SetValue("U_ObjType", j, (oGrid.DataTable.GetValue("ObjType", i)))
                    oDataSource.SetValue("U_ComentarioFac", j, (oGrid.DataTable.GetValue("Comments", i)))
                    oDataSource.SetValue("U_BcoPr", j, (oGrid.DataTable.GetValue("CodBanco", i)))
                    oDataSource.SetValue("U_TipCtaPr", j, (oGrid.DataTable.GetValue("TipoCuenta", i)))

                    TotalFacts += (oGrid.DataTable.GetValue("DocTotal", i))
                    totalMontoSP += (oGrid.DataTable.GetValue("DocTotal", i))
                    totalSaldoSP += (oGrid.DataTable.GetValue("saldo", i))
                    j += 1
                End If
            Next

            mMatrix.LoadFromDataSource()

            If mMatrix.RowCount > 0 Then
                If mMatrix.Columns.Item("U_DocEntry").Cells.Item(mMatrix.RowCount).Specific.Value = "" Then
                    mMatrix.DeleteRow(mMatrix.RowCount)
                End If
            End If

            If LineasSinCuenta.Count > 0 Then
                For Each a As Integer In LineasSinCuenta
                    oGrid.DataTable.SetValue("Chek", a, "N")
                    ListadeLineas.Remove(a)
                Next
            End If

            If mMatrix.RowCount = 0 Then rsboApp.StatusBar.SetText("No hay facturas seleccionadas... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            For Each k As Integer In ListadeLineas.AsEnumerable.Reverse 'j
                If Convert.ToString(oGrid.DataTable.GetValue("Chek", k)) = "Y" Then oGrid.DataTable.Rows.Remove(k)
            Next

            Dim lblFac As SAPbouiCOM.StaticText = oForm.Items.Item("lblFac").Specific
            lblFac.Caption = mMatrix.RowCount.ToString

            'Sumo total de facturas por pagar
            Dim lblSMon As SAPbouiCOM.StaticText = oForm.Items.Item("lblSMon").Specific
            lblSMon.Caption = CStr(CDec(IIf(lblSMon.Caption = "", "0", lblSMon.Caption)) + totalMontoSP)

            Dim lblSSal As SAPbouiCOM.StaticText = oForm.Items.Item("lblSSal").Specific
            lblSSal.Caption = CStr(CDec(IIf(lblSSal.Caption = "", "0", lblSSal.Caption)) + totalSaldoSP)
            '
            'NEW Resto a la linea de totales del grid
            oGrid.DataTable.SetValue("DocTotal", oGrid.DataTable.Rows.Count - 1, CDbl((oGrid.DataTable.GetValue("DocTotal", oGrid.DataTable.Rows.Count - 1))) - CDbl(totalMontoSP))
            oGrid.DataTable.SetValue("saldo", oGrid.DataTable.Rows.Count - 1, CDbl((oGrid.DataTable.GetValue("saldo", oGrid.DataTable.Rows.Count - 1))) - CDbl(totalSaldoSP))
            oGrid.CommonSetting.SetRowBackColor(2, RGB(255, 128, 0))
            oGrid.CommonSetting.SetCellEditable(2, 2, False)
            '
            Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific
            txtMP.Value = CStr(CDec(IIf(txtMP.Value = "", "0", txtMP.Value)) + totalSaldoSP) 'TotalFacts)

            Dim focus As SAPbouiCOM.Item = oForm.Items.Item("focus")
            focus.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            If ListadeLineas.Count > 0 Then ListadeLineas = New List(Of Integer)

            oForm.Freeze(False)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error AgregarMarcado: " & ex.Message.ToString, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error AgregarMarcado: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Public Function CrearSolicitudPago(ByRef DocEntryFacturaRecibida_UDO As String) As Boolean
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService

            rsboApp.StatusBar.SetText("Creando solicitud de pago...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oForm = rsboApp.Forms.Item("frmPagosMasivos")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSMTPAGOS")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            If CInt(NivelUsuario) = CInt(NivelMaximo) Then
                oGeneralData.SetProperty("U_Estado", "Aprobado") 'default
                oGeneralData.SetProperty("U_NivelAprob", NivelMaximo) 'default
            Else
                oGeneralData.SetProperty("U_Estado", "Revision") 'default
                oGeneralData.SetProperty("U_NivelAprob", "1") 'default
            End If
            oGeneralData.SetProperty("U_Cuenta", cbxCuenta.Value.ToString) 'oForm.Items.Item("txtCue").Specific.Value.ToString())
            Dim Banco As String() = cbxCuenta.Selected.Description.Split(":")
            oGeneralData.SetProperty("U_Banco", Banco(0).ToString)

            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
            oGeneralData.SetProperty("U_MedioPago", cbxfpago.Value.ToString())

            Try
                If cbxfpago.Value = "Servicios Basicos" Then oGeneralData.SetProperty("U_Estado", "Aprobado")
            Catch ex As Exception
            End Try

            oGeneralData.SetProperty("U_Tipo", "Standard") 'default
            oGeneralData.SetProperty("U_TotalPagado", oForm.Items.Item("txtMP").Specific.Value.ToString())
            oGeneralData.SetProperty("U_FacProcesadas", oForm.Items.Item("lblFac").Specific.Caption.ToString())

            oChildren = oGeneralData.Child("SS_PM_DET1")

            Dim matrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim i As Integer

            For i = 1 To matrix.RowCount
                oChild = oChildren.Add
                oChild.SetProperty("U_CodProv", matrix.Columns.Item("U_CodProv").Cells.Item(i).Specific.Value) '
                oChild.SetProperty("U_Proveedor", matrix.Columns.Item("U_Prov").Cells.Item(i).Specific.Value) '
                oChild.SetProperty("U_Vencimiento", CInt(matrix.Columns.Item("U_Venc").Cells.Item(i).Specific.Value)) '
                Dim fecha As String = matrix.Columns.Item("U_FecVec").Cells.Item(i).Specific.Value '
                Dim fechaRegistro As Date = Date.ParseExact(fecha, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                oChild.SetProperty("U_FechaVen", fechaRegistro.ToString("MM/dd/yyyy"))
                oChild.SetProperty("U_Monto", matrix.Columns.Item("U_Mon").Cells.Item(i).Specific.Value) '
                oChild.SetProperty("U_Saldo", matrix.Columns.Item("U_Sal").Cells.Item(i).Specific.Value) '
                oChild.SetProperty("U_Pago", matrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value) '
                oChild.SetProperty("U_DocEntry", matrix.Columns.Item("U_DocEntry").Cells.Item(i).Specific.Value.ToString) '
                oChild.SetProperty("U_Cuota", matrix.Columns.Item("U_Cuo").Cells.Item(i).Specific.Value.ToString) '
                oChild.SetProperty("U_ObjType", matrix.Columns.Item("U_ObjType").Cells.Item(i).Specific.Value.ToString)
                oChild.SetProperty("U_NumDoc", matrix.Columns.Item("U_NumDoc").Cells.Item(i).Specific.Value.ToString)

                'oChild.SetProperty("U_NumLinea", matrix.Columns.Item("U_NL").Cells.Item(i).Specific.Value.ToString)
                oChild.SetProperty("U_Sucursal", matrix.Columns.Item("U_Suc").Cells.Item(i).Specific.Value.ToString)
                oChild.SetProperty("U_Proyecto", matrix.Columns.Item("U_Proy").Cells.Item(i).Specific.Value.ToString)
                oChild.SetProperty("U_CtaBcoPr", matrix.Columns.Item("U_CtaBco").Cells.Item(i).Specific.Value.ToString)

                oChild.SetProperty("U_ComentarioFac", matrix.Columns.Item("U_ComFac").Cells.Item(i).Specific.Value.ToString)

                Dim oCheckBox As SAPbouiCOM.CheckBox = CType(matrix.Columns.Item("U_Consoli").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                oChild.SetProperty("U_Consolidado", If(oCheckBox.Checked, "Y", "N"))
                oChild.SetProperty("U_NumChe", matrix.Columns.Item("U_NumChe").Cells.Item(i).Specific.Value.ToString)

                oChild.SetProperty("U_BcoPr", matrix.Columns.Item("U_BcoPr").Cells.Item(i).Specific.Value.ToString)
                oChild.SetProperty("U_TipCtaPr", matrix.Columns.Item("TipCtaPr").Cells.Item(i).Specific.Value.ToString)
            Next

            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryFacturaRecibida_UDO = oGeneralParams.GetProperty("DocEntry")
            rsboApp.StatusBar.SetText("Solicitud de pago creada con éxito! " & DocEntryFacturaRecibida_UDO.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            If cbxfpago.Value = "Cheque" Then CrearOrdenDePago(DocEntryFacturaRecibida_UDO)

            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error al CrearSolicitudPago UDO: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al CrearSolicitudPago UDO: " & ex.Message, "frmPagosMasivos")
            Return False
        End Try
    End Function

    'Add JP 06/02/2025
    Public Function CrearOrdenDePago(ByVal DocEntrySolPag As String) As Boolean
        Try
            rsboApp.StatusBar.SetText("Creando Orden de pago...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm = rsboApp.Forms.Item("frmPagosMasivos")

            Dim ListaDePagos As New List(Of Entidades.PagoMasivo)
            Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific

            Try
                Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific

                For j As Integer = 1 To mMatrix.RowCount
                    Dim item As New Entidades.PagoMasivo With {
                    .DocEntryFP = IIf(String.IsNullOrEmpty(mMatrix.Columns.Item("U_DocEntry").Cells.Item(j).Specific.Value.ToString()), 0, mMatrix.Columns.Item("U_DocEntry").Cells.Item(j).Specific.Value.ToString()),
                    .CodPro = mMatrix.Columns.Item("U_CodProv").Cells.Item(j).Specific.Value.ToString(),
                    .NomPro = mMatrix.Columns.Item("U_Prov").Cells.Item(j).Specific.Value.ToString(),
                    .Vencimiento = mMatrix.Columns.Item("U_Venc").Cells.Item(j).Specific.Value.ToString(),
                    .FechaVencimiento = mMatrix.Columns.Item("U_FecVec").Cells.Item(j).Specific.Value.ToString(),
                    .Monto = mMatrix.Columns.Item("U_Mon").Cells.Item(j).Specific.Value.ToString(),
                    .Saldo = mMatrix.Columns.Item("U_Sal").Cells.Item(j).Specific.Value.ToString(),
                    .Pagar = mMatrix.Columns.Item("U_Pag").Cells.Item(j).Specific.Value.ToString(),
                    .Cuota = mMatrix.Columns.Item("U_Cuo").Cells.Item(j).Specific.Value.ToString(),
                    .Sucursal = mMatrix.Columns.Item("U_Suc").Cells.Item(j).Specific.Value.ToString(),
                    .Proyecto = mMatrix.Columns.Item("U_Proy").Cells.Item(j).Specific.Value.ToString(),
                    .ObjType = mMatrix.Columns.Item("U_ObjType").Cells.Item(j).Specific.Value.ToString(),
                    .LineId = CStr(j), 'mMatrix.Columns.Item("LineId").Cells.Item(j).Specific.Value.ToString(),
                    .Consolidado = CType(mMatrix.Columns.Item("U_Consoli").Cells.Item(j).Specific, SAPbouiCOM.CheckBox).Checked
                    }
                    ListaDePagos.Add(item)
                Next
            Catch ex As Exception
                rsboApp.StatusBar.SetText("Error agrupando pagos por socio de negocio: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Utilitario.Util_Log.Escribir_Log("Error agrupando pagos por socio de negocio: " & ex.Message, "frmPagosMasivos")
                Return False
            End Try

            Dim ListaAgrupada As Object = New Object
            ListaAgrupada = (From a In ListaDePagos Group By a.CodPro, a.Consolidado Into Group).ToList

            If Not ListaAgrupada Is Nothing Then

                For Each b As Object In ListaAgrupada

                    Dim listalineasConsolidadas As New List(Of String)
                    Dim oGeneralService As SAPbobsCOM.GeneralService
                    Dim oGeneralData As SAPbobsCOM.GeneralData
                    Dim oChild As SAPbobsCOM.GeneralData
                    Dim oChildren As SAPbobsCOM.GeneralDataCollection
                    Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                    Dim oCompanyService As SAPbobsCOM.CompanyService
                    oCompanyService = rCompany.GetCompanyService
                    oGeneralService = oCompanyService.GetGeneralService("SSOPPMPAGOS")
                    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                    oGeneralData.SetProperty("U_Estado", "Aprobado") 'default
                    oGeneralData.SetProperty("U_NivelAprob", NivelMaximo) 'default
                    Utilitario.Util_Log.Escribir_Log("U_NivelAprob:  " & NivelMaximo.ToString(), "frmPagosMasivos")

                    oGeneralData.SetProperty("U_Cuenta", cbxCuenta.Value.ToString) 'oForm.Items.Item("txtCue").Specific.Value.ToString())
                    Utilitario.Util_Log.Escribir_Log("cbxCuenta:  " & cbxCuenta.Value.ToString, "frmPagosMasivos")

                    Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
                    oGeneralData.SetProperty("U_MedioPago", cbxfpago.Value.ToString())
                    Utilitario.Util_Log.Escribir_Log("U_MedioPago:  " & cbxfpago.Value.ToString, "frmPagosMasivos")

                    oGeneralData.SetProperty("U_Tipo", "Standard") 'default
                    oGeneralData.SetProperty("U_TotalPagado", oForm.Items.Item("txtMP").Specific.Value.ToString())
                    Utilitario.Util_Log.Escribir_Log("U_TotalPagado:  " & oForm.Items.Item("txtMP").Specific.Value.ToString(), "frmPagosMasivos")

                    oGeneralData.SetProperty("U_FacProcesadas", oForm.Items.Item("lblFac").Specific.Caption.ToString())
                    Utilitario.Util_Log.Escribir_Log("U_FacProcesadas:  " & oForm.Items.Item("lblFac").Specific.Caption.ToString(), "frmPagosMasivos")

                    oGeneralData.SetProperty("U_SolicitudPago", DocEntrySolPag) 'New
                    Utilitario.Util_Log.Escribir_Log("U_MedioPago:  " & DocEntrySolPag.ToString, "frmPagosMasivos")

                    oChildren = oGeneralData.Child("SS_PM_OP_DET1")

                    For Each a As Object In b.group
                        oChild = oChildren.Add

                        Utilitario.Util_Log.Escribir_Log("U_CodProv:  " & a.CodPro.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_CodProv", a.CodPro) '

                        Utilitario.Util_Log.Escribir_Log("U_Proveedor:  " & a.NomPro.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Proveedor", a.NomPro) '

                        Utilitario.Util_Log.Escribir_Log("U_Vencimiento:  " & CInt(a.Vencimiento).ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Vencimiento", CInt(a.Vencimiento)) '

                        Utilitario.Util_Log.Escribir_Log("FechaVencimiento:  " & a.FechaVencimiento.ToString, "frmPagosMasivos")
                        Dim fecha As String = a.FechaVencimiento '
                        Dim fechaRegistro As Date = Date.ParseExact(fecha, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        Utilitario.Util_Log.Escribir_Log("FechaVencimiento2:  " & fechaRegistro.ToString("MM/dd/yyyy"), "frmPagosMasivos")
                        oChild.SetProperty("U_FechaVen", fechaRegistro.ToString("MM/dd/yyyy"))

                        Utilitario.Util_Log.Escribir_Log("Monto:  " & a.Monto.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Monto", a.Monto) '

                        Utilitario.Util_Log.Escribir_Log("Saldo:  " & a.Saldo.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Saldo", a.Saldo) '

                        Utilitario.Util_Log.Escribir_Log("Pagar:  " & a.Pagar.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Pago", a.Pagar) '

                        Utilitario.Util_Log.Escribir_Log("DocEntryFP:  " & CStr(a.DocEntryFP).ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_DocEntry", CStr(a.DocEntryFP)) '

                        Utilitario.Util_Log.Escribir_Log("Cuota:  " & a.Cuota.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Cuota", a.Cuota) '

                        Utilitario.Util_Log.Escribir_Log("ObjType:  " & a.ObjType.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_ObjType", a.ObjType)

                        Utilitario.Util_Log.Escribir_Log("Sucursal:  " & a.Sucursal.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Sucursal", a.Sucursal)

                        Utilitario.Util_Log.Escribir_Log("Proyecto:  " & a.Proyecto.ToString, "frmPagosMasivos")
                        oChild.SetProperty("U_Proyecto", a.Proyecto)

                        listalineasConsolidadas.Add(a.LineId)
                    Next
                    Dim DocEntryOrdenPago As String = ""
                    oGeneralParams = oGeneralService.Add(oGeneralData)
                    DocEntryOrdenPago = oGeneralParams.GetProperty("DocEntry")
                    rsboApp.StatusBar.SetText("Orden de pago creado con éxito! " & DocEntryOrdenPago, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    For Each IdLinea As String In listalineasConsolidadas
                        ActualizaLineaSolicitudPagoConOrdenPago(DocEntrySolPag, IdLinea, DocEntryOrdenPago)
                    Next
                Next
            End If

            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error al CrearOrdenDePago UDO: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al CrearOrdenDePago UDO: " & ex.Message, "frmPagosMasivos")
            Return False
        End Try
    End Function

    Public Function EdicionSolicitudPago(ByVal DocEntry_UDO As String) As Boolean
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific

            rsboApp.StatusBar.SetText("Actualizando solicitud de pago...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oForm = rsboApp.Forms.Item("frmPagosMasivos")

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSMTPAGOS")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntry_UDO)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            Dim txtCBan As SAPbouiCOM.EditText = oForm.Items.Item("txtCBan").Specific

            If Not String.IsNullOrEmpty(txtCBan.Value.ToString) Then
                oGeneralData.SetProperty("U_Estado", "Archivo Procesado Banco")
            Else
                If CInt(NivelUsuario) = CInt(NivelMaximo) Then
                    oGeneralData.SetProperty("U_Estado", "Aprobado") 'default
                Else
                    oGeneralData.SetProperty("U_Estado", "Revision") 'default
                End If
            End If

            'Add JP 26/02/2025
            Try
                oGeneralData.SetProperty("U_MedioPago", cbxfpago.Value.ToString)
                Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific
                oGeneralData.SetProperty("U_Cuenta", cbxCuenta.Value.ToString)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al actualizar campos medio pago y cuenta: " & ex.Message.ToString, "frmPagosMasivos")
            End Try
            '

            oGeneralData.SetProperty("U_TotalPagado", oForm.Items.Item("txtMP").Specific.Value.ToString())
            oGeneralData.SetProperty("U_FacProcesadas", CInt(oForm.Items.Item("lblFac").Specific.Caption.ToString()))

            oChildren = oGeneralData.Child("SS_PM_DET1")
            Dim matrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific

            For i As Integer = oChildren.Count - 1 To 0 Step -1
                oChildren.Remove(i)
            Next

            For J As Integer = 1 To matrix.RowCount
                oChild = oChildren.Add
                oChild.SetProperty("U_CodProv", matrix.Columns.Item("U_CodProv").Cells.Item(J).Specific.Value)
                oChild.SetProperty("U_Proveedor", matrix.Columns.Item("U_Prov").Cells.Item(J).Specific.Value)
                oChild.SetProperty("U_Vencimiento", CInt(matrix.Columns.Item("U_Venc").Cells.Item(J).Specific.Value))
                Dim fecha As String = matrix.Columns.Item("U_FecVec").Cells.Item(J).Specific.Value
                Dim fechaRegistro As Date = Date.ParseExact(fecha, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                oChild.SetProperty("U_FechaVen", fechaRegistro.ToString("MM/dd/yyyy"))
                oChild.SetProperty("U_Monto", matrix.Columns.Item("U_Mon").Cells.Item(J).Specific.Value)
                oChild.SetProperty("U_Saldo", matrix.Columns.Item("U_Sal").Cells.Item(J).Specific.Value)
                oChild.SetProperty("U_Pago", matrix.Columns.Item("U_Pag").Cells.Item(J).Specific.Value)
                oChild.SetProperty("U_DocEntry", matrix.Columns.Item("U_DocEntry").Cells.Item(J).Specific.Value.ToString)
                oChild.SetProperty("U_Cuota", matrix.Columns.Item("U_Cuo").Cells.Item(J).Specific.Value.ToString)
                oChild.SetProperty("U_ObjType", matrix.Columns.Item("U_ObjType").Cells.Item(J).Specific.Value.ToString)
                oChild.SetProperty("U_NumDoc", matrix.Columns.Item("U_NumDoc").Cells.Item(J).Specific.Value.ToString)

                'oChild.SetProperty("U_NumLinea", matrix.Columns.Item("U_NL").Cells.Item(J).Specific.Value.ToString)
                oChild.SetProperty("U_Sucursal", matrix.Columns.Item("U_Suc").Cells.Item(J).Specific.Value.ToString)
                oChild.SetProperty("U_Proyecto", matrix.Columns.Item("U_Proy").Cells.Item(J).Specific.Value.ToString)
                oChild.SetProperty("U_CtaBcoPr", matrix.Columns.Item("U_CtaBco").Cells.Item(J).Specific.Value.ToString)

                oChild.SetProperty("U_ComentarioFac", matrix.Columns.Item("U_ComFac").Cells.Item(J).Specific.Value.ToString)

                Dim oCheckBox As SAPbouiCOM.CheckBox = CType(matrix.Columns.Item("U_Consoli").Cells.Item(J).Specific, SAPbouiCOM.CheckBox)
                oChild.SetProperty("U_Consolidado", If(oCheckBox.Checked, "Y", "N"))
                oChild.SetProperty("U_NumChe", matrix.Columns.Item("U_NumChe").Cells.Item(J).Specific.Value.ToString)
            Next

            oGeneralService.Update(oGeneralData)
            rsboApp.StatusBar.SetText("Solicitud de pago actualizada con exito! " & DocEntry.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            If cbxfpago.Value = "Cheque" Then CrearOrdenDePago(DocEntry_UDO)

            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error al actualizar solicitud de pago: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error al CrearSolicitudPago: " & ex.Message.ToString, "frmPagosMasivos")
            Return False
        End Try
    End Function

    Public Function CrearPagoMasivo() As Boolean
        Try
            rsboApp.StatusBar.SetText("Agrupando pagos por socio de negocio, sucursal y proyecto...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Utilitario.Util_Log.Escribir_Log("Agrupando pagos por socio de negocio, sucursal y proyecto...", "frmPagosMasivos")

            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
            Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific
            Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific
            Dim txtCBan As SAPbouiCOM.EditText = oForm.Items.Item("txtCBan").Specific

            Dim txtFArc As SAPbouiCOM.EditText = oForm.Items.Item("txtFArc").Specific
            Dim FechaArchivoBanco As String = txtFArc.Value.ToString

            Dim lblNumChe As SAPbouiCOM.StaticText = oForm.Items.Item("lblNumChe").Specific

            Dim ListaDePagos As New List(Of Entidades.PagoMasivo)
            Dim AcumuladorDeToales As Decimal = 0
            Dim formato As String = "yyyyMMdd"
            Dim NumCheque As Integer = CInt(IIf(String.IsNullOrEmpty(lblNumChe.Caption), "0", lblNumChe.Caption))

            Try
                For i As Integer = 1 To mMatrix.RowCount
                    Dim item As New Entidades.PagoMasivo With {
                    .DocEntryFP = IIf(String.IsNullOrEmpty(mMatrix.Columns.Item("U_DocEntry").Cells.Item(i).Specific.Value.ToString()), 0, mMatrix.Columns.Item("U_DocEntry").Cells.Item(i).Specific.Value.ToString()),
                    .CodPro = mMatrix.Columns.Item("U_CodProv").Cells.Item(i).Specific.Value.ToString(),
                    .NomPro = mMatrix.Columns.Item("U_Prov").Cells.Item(i).Specific.Value.ToString(),
                    .Vencimiento = mMatrix.Columns.Item("U_Venc").Cells.Item(i).Specific.Value.ToString(),
                    .FechaVencimiento = mMatrix.Columns.Item("U_FecVec").Cells.Item(i).Specific.Value.ToString(),
                    .Monto = mMatrix.Columns.Item("U_Mon").Cells.Item(i).Specific.Value.ToString(),
                    .Saldo = mMatrix.Columns.Item("U_Sal").Cells.Item(i).Specific.Value.ToString(),
                    .Pagar = mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value.ToString(),
                    .Cuota = mMatrix.Columns.Item("U_Cuo").Cells.Item(i).Specific.Value.ToString(),
                    .Sucursal = mMatrix.Columns.Item("U_Suc").Cells.Item(i).Specific.Value.ToString(),
                    .Proyecto = mMatrix.Columns.Item("U_Proy").Cells.Item(i).Specific.Value.ToString(),
                    .ObjType = mMatrix.Columns.Item("U_ObjType").Cells.Item(i).Specific.Value.ToString(),
                    .LineId = mMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString(),
                    .Consolidado = CType(mMatrix.Columns.Item("U_Consoli").Cells.Item(i).Specific, SAPbouiCOM.CheckBox).Checked
                    }
                    ListaDePagos.Add(item)
                Next
            Catch ex As Exception
                rsboApp.StatusBar.SetText("Error agrupando pagos por socio de negocio: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Utilitario.Util_Log.Escribir_Log("Error agrupando pagos por socio de negocio: " & ex.Message, "frmPagosMasivos")
                Return False
            End Try

            Dim ListaAgrupada As Object = New Object

            rsboApp.StatusBar.SetText("Creando pagos masivos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Utilitario.Util_Log.Escribir_Log("Creando pagos masivos... ", "frmPagosMasivos")

            If cbxfpago.Value = "Cheque" Then

                Utilitario.Util_Log.Escribir_Log("Metodo de pago: " & cbxfpago.Value, "frmPagosMasivos")

                ListaAgrupada = (From a In ListaDePagos Group By a.CodPro, a.Consolidado Into Group).ToList

                If Not ListaAgrupada Is Nothing Then

                    rCompany.StartTransaction()
                    Utilitario.Util_Log.Escribir_Log("Tipo: " & TipoSolicitudPMTransferencia, "frmPagosMasivos")

                    If TipoSolicitudPMTransferencia = "Nomina" Then

                        Dim nErr As Long
                        Dim errMsg As String, Qr As String = ""
                        Dim acumulador As Double = 0
                        Dim LineId As String = ""
                        Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

                        pag.DocType = SAPbobsCOM.BoRcptTypes.rAccount
                        pag.DocDate = Date.Now
                        pag.DueDate = Date.Now
                        pag.TaxDate = Date.Now
                        pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                        pag.DocCurrency = "USD"
                        pag.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value

                        pag.Remarks = $"PAGO DE NOMINA SEGUN SOLICITUD DE PAGO MASIVO #{txtDocEnt.Value} CON CHEQUE #{NumCheque}"
                        pag.JournalRemarks = $"PAGO DE NOMINA SEGUN SOLICITUD DE PAGO MASIVO #{txtDocEnt.Value} CON CHEQUE #{NumCheque}"

                        For Each a As Object In ListaAgrupada
                            For Each b As Object In a.Group
                                LineId = b.LineId
                                acumulador += Convert.ToDecimal(b.Pagar)
                            Next
                        Next

                        pag.AccountPayments.AccountCode = "20107040101" 'Seteado Cuenta Sueldos por pagar
                        pag.AccountPayments.SumPaid = acumulador
                        pag.AccountPayments.Add()

                        Dim QryValCuentaBanco As String = "", cuentaBco As String = "", BankCode As String = "", Country As String = ""
                        QryValCuentaBanco = "SELECT TOP 1 ""Account"", ""BankCode"", ""Country"" FROM ""DSC1"" WHERE ""GLAccount"" = '" & cbxCuenta.Value & "'"

                        Utilitario.Util_Log.Escribir_Log("Query validar Cuenta Banco Cheque: " + QryValCuentaBanco.ToString, "frmPagosMasivos")
                        cuentaBco = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "Account", "")
                        BankCode = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "BankCode", "0")
                        Country = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "Country", "")

                        Utilitario.Util_Log.Escribir_Log($"Country:{Country} Bank:{BankCode} Accountt:{cuentaBco} CheckSum:{acumulador} CheckNumber:{NumCheque}", "frmPagosMasivos")
                        pag.Checks.CountryCode = Country
                        pag.Checks.BankCode = BankCode
                        pag.Checks.AccounttNum = cuentaBco
                        pag.Checks.CheckSum = acumulador
                        pag.Checks.DueDate = Date.Now
                        pag.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES
                        pag.Checks.CheckNumber = NumCheque
                        pag.Checks.Add()

                        If pag.Add = 0 Then 'Si creo los pagos
                            Dim DocEntry As String = ""
                            rCompany.GetNewObjectCode(DocEntry)
                            rsboApp.StatusBar.SetText($"P/E creado correctamente: {DocEntry}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Utilitario.Util_Log.Escribir_Log($"P/E creado correctamente: {DocEntry}", "frmPagosMasivos")

                            ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", "", "", "", LineId, DocEntry, "", "", "", "", NumCheque)
                            NumCheque += 1
                        Else
                            rCompany.GetLastError(nErr, errMsg)
                            rsboApp.StatusBar.SetText($"Error al integrar pago: {nErr} - {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log($"Error al integrar pago: {nErr} - {errMsg}", "frmPagosMasivos")
                            rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            Return False
                        End If

                    Else

                        Dim nErr As Long
                        Dim errMsg As String, Qr As String = ""
                        Dim LineId As String = ""

                        For Each b As Object In ListaAgrupada

                            If b.Consolidado = False Then
                                For Each a As Object In b.group

                                    Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                                    pag.DocType = SAPbobsCOM.BoRcptTypes.rSupplier

                                    pag.DocDate = Date.Now
                                    pag.DueDate = Date.Now
                                    pag.CardCode = b.CodPro
                                    pag.TaxDate = Date.Now
                                    pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                                    pag.DocCurrency = "USD"
                                    pag.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value

                                    pag.Remarks = $"PAGO A PROVEEDOR SEGUN SOLICITUD DE PAGO MASIVO #{txtDocEnt.Value} CON CHEQUE #{NumCheque}"
                                    pag.JournalRemarks = $"PAGO A PROVEEDOR SEGUN SOLICITUD DE PAGO MASIVO #{txtDocEnt.Value} CON CHEQUE #{NumCheque}"

                                    Utilitario.Util_Log.Escribir_Log($"DocEntry:{a.DocEntryFP} SumApplied:{a.Pagar} InstallmentId:{a.Cuota} InvoiceType:{a.ObjType}", "frmPagosMasivos")
                                    pag.Invoices.DocEntry = a.DocEntryFP
                                    pag.Invoices.SumApplied = Convert.ToDecimal(a.Pagar)
                                    pag.Invoices.InstallmentId = CInt(a.Cuota)
                                    If CStr(a.ObjType) = "18" Then
                                        pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                                    ElseIf CStr(a.ObjType) = "204" Then
                                        pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment
                                    End If
                                    pag.Invoices.Add()

                                    LineId = a.LineId

                                    Dim QryValCuentaBanco As String = "", cuentaBco As String = "", BankCode As String = "", Country As String = ""
                                    QryValCuentaBanco = "SELECT TOP 1 ""Account"", ""BankCode"", ""Country"" FROM ""DSC1"" WHERE ""GLAccount"" = '" & cbxCuenta.Value & "'"
                                    Utilitario.Util_Log.Escribir_Log($"Query validar Cuenta Banco Cheque: {QryValCuentaBanco}", "frmPagosMasivos")
                                    cuentaBco = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "Account", "")
                                    BankCode = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "BankCode", "0")
                                    Country = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "Country", "")

                                    Utilitario.Util_Log.Escribir_Log($"Country:{Country} Bank:{BankCode} Accountt:{cuentaBco} CheckSum:{a.Pagar} CheckNumber:{NumCheque}", "frmPagosMasivos")

                                    pag.Checks.CountryCode = Country
                                    pag.Checks.BankCode = BankCode.ToString
                                    pag.Checks.AccounttNum = cuentaBco
                                    pag.Checks.CheckSum = Convert.ToDecimal(a.Pagar)
                                    pag.Checks.DueDate = Date.Now
                                    pag.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES
                                    pag.Checks.CheckNumber = NumCheque
                                    pag.Checks.Add()

                                    AcumuladorDeToales += Convert.ToDecimal(a.Pagar)

                                    If pag.Add = 0 Then 'Si creo los pagos
                                        Dim DocEntry As String = ""
                                        rCompany.GetNewObjectCode(DocEntry)
                                        rsboApp.StatusBar.SetText($"P/E creado correctamente: {DocEntry}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Utilitario.Util_Log.Escribir_Log($"P/E creado correctamente: {DocEntry}", "frmPagosMasivos")

                                        ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", "", "", "", LineId, DocEntry, "", "", "", "", NumCheque)
                                        NumCheque += 1
                                    Else
                                        rCompany.GetLastError(nErr, errMsg)
                                        rsboApp.StatusBar.SetText($"Error al integrar P/E: {nErr} - {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Utilitario.Util_Log.Escribir_Log($"Error al integrar P/E: {nErr} - {errMsg}", "frmPagosMasivos")
                                        rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        Return False
                                    End If
                                Next
                            Else

                                Dim acumulador As Decimal = 0
                                Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                                Dim listalineasConsolidadas As New List(Of String)

                                pag.DocType = SAPbobsCOM.BoRcptTypes.rSupplier
                                pag.DocDate = Date.Now
                                pag.DueDate = Date.Now
                                pag.CardCode = b.CodPro
                                pag.TaxDate = Date.Now
                                pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                                pag.DocCurrency = "USD"
                                pag.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value

                                pag.Remarks = $"PAGO A PROVEEDOR SEGUN SOLICITUD DE PAGO MASIVO #{txtDocEnt.Value} CON CHEQUE #{NumCheque}"
                                pag.JournalRemarks = $"PAGO A PROVEEDOR SEGUN SOLICITUD DE PAGO MASIVO #{txtDocEnt.Value} CON CHEQUE #{NumCheque}"

                                Dim i As Integer = 0
                                For Each a As Object In b.Group
                                    Utilitario.Util_Log.Escribir_Log($"DocEntry:{a.DocEntryFP} SumApplied:{a.Pagar} InstallmentId:{a.Cuota} InvoiceType:{a.ObjType}", "frmPagosMasivos")
                                    pag.Invoices.DocEntry = a.DocEntryFP
                                    pag.Invoices.SumApplied = Convert.ToDecimal(a.Pagar)
                                    pag.Invoices.InstallmentId = CInt(a.Cuota)
                                    pag.Invoices.Add()

                                    acumulador += Convert.ToDecimal(a.Pagar)
                                    listalineasConsolidadas.Add(a.LineId)
                                    LineId = a.LineId
                                Next

                                Dim QryValCuentaBanco As String = "", cuentaBco As String = "", BankCode As String = "", Country As String = ""
                                QryValCuentaBanco = "SELECT TOP 1 ""Account"", ""BankCode"", ""Country"" FROM ""DSC1"" WHERE ""GLAccount"" = '" & cbxCuenta.Value & "'"

                                Utilitario.Util_Log.Escribir_Log($"Query validar Cuenta Banco Cheque: {QryValCuentaBanco}", "frmPagosMasivos")
                                cuentaBco = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "Account", "")
                                BankCode = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "BankCode", "0")
                                Country = oFuncionesAddon.getRSvalue(QryValCuentaBanco, "Country", "")

                                Utilitario.Util_Log.Escribir_Log($"Country:{Country} Bank:{BankCode} Accountt:{cuentaBco} CheckSum:{acumulador} CheckNumber:{NumCheque}", "frmPagosMasivos")
                                pag.Checks.CountryCode = Country
                                pag.Checks.BankCode = BankCode.ToString
                                pag.Checks.AccounttNum = cuentaBco
                                pag.Checks.CheckSum = acumulador
                                pag.Checks.DueDate = Date.Now
                                pag.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES
                                pag.Checks.CheckNumber = NumCheque
                                pag.Checks.Add()

                                AcumuladorDeToales += acumulador

                                If pag.Add = 0 Then 'Si creo los pagos
                                    Dim DocEntry As String = ""
                                    rCompany.GetNewObjectCode(DocEntry)
                                    rsboApp.StatusBar.SetText($"P/E creado correctamente: {DocEntry}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Utilitario.Util_Log.Escribir_Log($"P/E creado correctamente: {DocEntry}", "frmPagosMasivos")

                                    For Each IdLinea As String In listalineasConsolidadas
                                        ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", "", "", "", IdLinea, DocEntry, "", "", "", "", NumCheque)
                                    Next

                                    NumCheque += 1
                                Else
                                    rCompany.GetLastError(nErr, errMsg)
                                    rsboApp.StatusBar.SetText($"Error al integrar P/E: {nErr} - {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Utilitario.Util_Log.Escribir_Log($"Error al integrar P/E: {nErr} - {errMsg}", "frmPagosMasivos")
                                    rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    Return False
                                End If
                            End If
                        Next
                    End If
                End If
            End If

            If cbxfpago.Value = "Servicios Basicos" Then
                Utilitario.Util_Log.Escribir_Log("Metodo de pago: " & cbxfpago.Value, "frmPagosMasivos")

                Try
                    rCompany.StartTransaction() 'Add 13/02/2025
                    For Each a As Entidades.PagoMasivo In ListaDePagos 'As Object In ListaAgrupada
                        Dim nErr As Long
                        Dim errMsg As String, Qr As String = ""
                        Dim LineId As String = ""
                        Dim listalineasConsolidadas As New List(Of String)

                        Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

                        pag.DocType = SAPbobsCOM.BoRcptTypes.rSupplier
                        pag.DocDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                        pag.DueDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                        pag.CardCode = a.CodPro
                        pag.TaxDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                        pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                        pag.DocCurrency = "USD"
                        pag.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value

                        pag.Remarks = $"PAGO CONSUMO DE SERVICIOS BASICOS SEGUN SOLICITUD #{txtDocEnt.Value}"
                        pag.JournalRemarks = $"PAGO CONSUMO DE SERVICIOS BASICOS SEGUN SOLICITUD #{txtDocEnt.Value}"

                        Utilitario.Util_Log.Escribir_Log($"DocEntry:{a.DocEntryFP} SumApplied:{a.Pagar} InstallmentId:{a.Cuota} InvoiceType:{a.ObjType}", "frmPagosMasivos")
                        pag.Invoices.DocEntry = a.DocEntryFP
                        pag.Invoices.SumApplied = Convert.ToDecimal(a.Pagar)
                        pag.Invoices.InstallmentId = CInt(a.Cuota)
                        If CStr(a.ObjType) = "18" Then
                            pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                        ElseIf CStr(a.ObjType) = "204" Then
                            pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment
                        End If
                        pag.Invoices.Add()

                        Utilitario.Util_Log.Escribir_Log($"TransferAccount: {cbxCuenta.Value} TransferSum: {a.Pagar}", "frmPagosMasivos")
                        pag.TransferAccount = cbxCuenta.Value
                        pag.TransferSum = Convert.ToDecimal(a.Pagar)
                        pag.TransferDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture) 'Date.Now

                        If pag.Add = 0 Then 'Si creo los pagos
                            Dim DocEntry As String = ""
                            rCompany.GetNewObjectCode(DocEntry)
                            rsboApp.StatusBar.SetText("P/E de Serv. Básico creado correctamente: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Utilitario.Util_Log.Escribir_Log("P/E de Serv. Básico creado correctamente: " & DocEntry, "frmPagosMasivos")

                            ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", "", "", "", a.LineId, DocEntry)
                        Else
                            rCompany.GetLastError(nErr, errMsg)
                            rsboApp.StatusBar.SetText($"Error al integrar P/E de Serv. Básico: {nErr}-{errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log($"Error al integrar P/E de Serv. Básico: {nErr} - {errMsg}", "frmPagosMasivos")
                            rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            Return False
                        End If
                    Next
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log($"Error campos de cabecera de pago Servicio Básico: {ex.Message} ", "frmPagosMasivos")
                    rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Return False
                End Try
            End If

            If cbxfpago.Value = "Transferencia" Then

                Utilitario.Util_Log.Escribir_Log("Metodo de pago: " & cbxfpago.Value, "frmPagosMasivos")
                ListaAgrupada = (From a In ListaDePagos Group By a.CodPro, a.Sucursal, a.Proyecto Into Group).ToList

                If Not ListaAgrupada Is Nothing Then
                    rCompany.StartTransaction()
                    If TipoSolicitudPMTransferencia = "Nomina" Then

                        Dim nErr As Long
                        Dim errMsg As String, Qr As String = ""
                        Dim acumulador As Double = 0
                        Dim LineId As String = ""

                        Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

                        pag.DocType = SAPbobsCOM.BoRcptTypes.rAccount
                        pag.DocDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture) 'Date.Now
                        pag.DueDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture) 'Date.Now
                        pag.TaxDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture) 'Date.Now
                        pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                        pag.DocCurrency = "USD"
                        pag.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value
                        'pag.Remarks = "P/c de pago de nomina de empleados cash management"

                        pag.Remarks = $"PAGO DE NOMINA CASH MANAGEMENT SEGUN SOLICITUD #{txtDocEnt.Value}"
                        pag.JournalRemarks = $"PAGO DE NOMINA CASH MANAGEMENT SEGUN SOLICITUD #{txtDocEnt.Value}"

                        For Each a As Object In ListaAgrupada
                            For Each b As Object In a.Group
                                acumulador += Convert.ToDecimal(b.Pagar)
                            Next
                        Next

                        pag.AccountPayments.AccountCode = "20107040101" 'Seteado Cuenta Sueldos por pagar
                        pag.AccountPayments.SumPaid = acumulador
                        pag.AccountPayments.Add()

                        pag.TransferAccount = cbxCuenta.Value 'Cuenta de Banco 
                        pag.TransferSum = acumulador
                        pag.TransferDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture) 'Date.Now
                        Try
                            pag.TransferReference = txtCBan.Value
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("Error ingresando referencia de la transferencia del banco! " & ex.Message, "frmPagosMasivos")
                        End Try

                        If pag.Add = 0 Then 'Si creo los pagos
                            Dim DocEntry As String = ""
                            rCompany.GetNewObjectCode(DocEntry)
                            rsboApp.StatusBar.SetText("P/E creado correctamente: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Utilitario.Util_Log.Escribir_Log("P/E creado correctamente: " & DocEntry, "frmPagosMasivos")

                            ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", txtCBan.Value, "", DocEntry, "0", "", "", FechaArchivoBanco)
                        Else
                            rCompany.GetLastError(nErr, errMsg)
                            rsboApp.StatusBar.SetText($"Error al integrar P/E: {nErr} - {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log($"Error al integrar P/E: {nErr} - {errMsg}", "frmPagosMasivos")
                            rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            Return False
                        End If

                    Else

                        For Each a As Object In ListaAgrupada

                            Dim nErr As Long
                            Dim errMsg As String, Qr As String = ""
                            Dim acumulador As Double = 0
                            Dim LineId As String = ""
                            Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                            Dim listalineasConsolidadas As New List(Of String)

                            Try
                                pag.DocType = SAPbobsCOM.BoRcptTypes.rSupplier
                                pag.DocDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                                pag.DueDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                                pag.CardCode = a.CodPro
                                pag.TaxDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                                pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                                pag.DocCurrency = "USD"
                                pag.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value

                                pag.Remarks = $"PAGO A PROVEDORES CASH MANAGEMENT SEGUN SOLICITUD #{txtDocEnt.Value}"
                                pag.JournalRemarks = $"PAGO A PROVEDORES CASH MANAGEMENT SEGUN SOLICITUD #{txtDocEnt.Value}"

                                Dim QryCodProy As String = ""
                                QryCodProy = "SELECT TOP 1 ""PrjCode"" FROM ""OPRJ"" WHERE ""PrjName"" = '" & a.Proyecto & "'"
                                Utilitario.Util_Log.Escribir_Log("Query para validar codigo de proyecto " + QryCodProy.ToString, "frmPagosMasivos")
                                QryCodProy = oFuncionesAddon.getRSvalue(QryCodProy, "PrjCode", "")

                                If Not String.IsNullOrEmpty(QryCodProy) Then pag.ProjectCode = QryCodProy

                                For Each b As Object In a.Group
                                    Utilitario.Util_Log.Escribir_Log($"DocEntry:{b.DocEntryFP} SumApplied:{b.Pagar} InstallmentId:{b.Cuota} InvoiceType:{b.ObjType}", "frmPagosMasivos")
                                    pag.Invoices.DocEntry = b.DocEntryFP
                                    pag.Invoices.SumApplied = Convert.ToDecimal(b.Pagar)
                                    pag.Invoices.InstallmentId = CInt(b.Cuota)
                                    If CStr(b.ObjType) = "18" Then
                                        pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                                    ElseIf CStr(b.ObjType) = "204" Then
                                        pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment
                                    End If

                                    pag.Invoices.Add()
                                    acumulador += Convert.ToDecimal(b.Pagar)
                                    LineId = b.LineId
                                Next

                                Utilitario.Util_Log.Escribir_Log($"Cuenta transitoria:{Functions.VariablesGlobales._CuentaTransitoriaPM} TransferSum:{acumulador}", "frmPagosMasivos")
                                pag.TransferAccount = Functions.VariablesGlobales._CuentaTransitoriaPM
                                pag.TransferSum = acumulador
                                pag.TransferDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now
                                Try
                                    pag.TransferReference = txtCBan.Value
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error ingresando referencia de la transferencia del banco! " & ex.Message, "frmPagosMasivos")
                                End Try

                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log($"Error campos de cabecera de pago Cheque/Transferencia: {ex.Message} ", "frmPagosMasivos")
                                rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                Return False
                            End Try

                            AcumuladorDeToales += acumulador

                            If pag.Add = 0 Then 'Si creo los pagos
                                Dim DocEntry As String = ""
                                rCompany.GetNewObjectCode(DocEntry)
                                rsboApp.StatusBar.SetText("P/E transitorio creado correctamente: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                Utilitario.Util_Log.Escribir_Log("P/E transitorio creado correctamente: " & DocEntry, "frmPagosMasivos")

                                ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", "", "", "", LineId, DocEntry)
                            Else
                                rCompany.GetLastError(nErr, errMsg)
                                rsboApp.StatusBar.SetText($"Error al integrar P/E transitorio: {nErr}-{errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Utilitario.Util_Log.Escribir_Log($"Error al integrar P/E transitorio: {nErr} - {errMsg}", "frmPagosMasivos")
                                rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                Return False
                            End If
                        Next

                        'Pago a cuenta de banco
                        Dim nErr2 As Long
                        Dim errMsg2 As String, Qr2 As String = ""
                        Dim pag2 As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

                        Try
                            pag2.DocType = SAPbobsCOM.BoRcptTypes.rAccount
                            pag2.DocDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now 'Date.Now
                            pag2.DueDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.NowDate.Now
                            pag2.TaxDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture)  'Date.Now Date.Now
                            pag2.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                            pag2.DocCurrency = "USD"
                            pag2.UserFields.Fields.Item("U_DE_PM").Value = txtDocEnt.Value

                            pag2.Remarks = $"PAGO A PROVEDORES CASH MANAGEMENT SEGUN SOLICITUD #{txtDocEnt.Value}"
                            pag2.JournalRemarks = $"PAGO A PROVEDORES CASH MANAGEMENT SEGUN SOLICITUD #{txtDocEnt.Value}"

                            pag2.AccountPayments.AccountCode = Functions.VariablesGlobales._CuentaTransitoriaPM
                            pag2.AccountPayments.SumPaid = AcumuladorDeToales
                            pag2.AccountPayments.Add()

                            pag2.TransferAccount = cbxCuenta.Value
                            pag2.TransferSum = AcumuladorDeToales
                            pag2.TransferDate = DateTime.ParseExact(txtFArc.Value, formato, Globalization.CultureInfo.InvariantCulture) 'Date.Now

                            Try
                                pag2.TransferReference = txtCBan.Value
                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error ingresando referencia de la transferencia del banco! " & ex.Message, "frmPagosMasivos")
                            End Try

                            If pag2.Add = 0 Then
                                Dim DocEntry2 As String = ""
                                rCompany.GetNewObjectCode(DocEntry2)
                                rsboApp.StatusBar.SetText("P/E consolidado creado correctamente:" & DocEntry2, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                Utilitario.Util_Log.Escribir_Log("Pago efectuado consolidado creado correctamente: " & DocEntry2, "frmPagosMasivos")

                                ActualizadoSolicitudDePago(txtDocEnt.Value, NivelAprobacion, "Archivo Procesado Banco", txtCBan.Value, "", DocEntry2, "0", "", "", FechaArchivoBanco)
                            Else
                                rCompany.GetLastError(nErr2, errMsg2)
                                rsboApp.StatusBar.SetText($"Error al crear P/E consolidado: {nErr2}-{errMsg2}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Utilitario.Util_Log.Escribir_Log($"Error al crear P/E consolidado: {nErr2} - {errMsg2}", "frmPagosMasivos")
                                rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                Return False
                            End If
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log($"Error creando pago consolidado: {ex.Message} ", "frmPagosMasivos")
                            rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            Return False
                        End Try
                    End If
                End If
            End If

            rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

            If cbxfpago.Value = "Cheque" Then
                Dim separadorCuentaBanco As String() = cbxCuenta.Selected.Description.Split(":")
                Dim nErr As Long
                Dim errMsg As String
                Dim oRecordSet As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("SELECT ""AbsEntry"" FROM ""DSC1"" WHERE ""BankCode"" = '" & separadorCuentaBanco(0) & "' AND ""GLAccount"" = '" & cbxCuenta.Value & "'")
                If Not oRecordSet.EoF Then
                    Dim absEntry As Integer = CInt(oRecordSet.Fields.Item("AbsEntry").Value)
                    oRecordSet = Nothing
                    GC.Collect()
                    Dim ActualizaCtaBcoPro As SAPbobsCOM.HouseBankAccounts = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oHouseBankAccounts)

                    If ActualizaCtaBcoPro.GetByKey(absEntry) Then
                        ActualizaCtaBcoPro.NextCheckNo = NumCheque
                        If ActualizaCtaBcoPro.Update() = 0 Then
                            rsboApp.StatusBar.SetText("Se actualizo el número de cheque con éxito!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Else
                            rCompany.GetLastError(nErr, errMsg)
                            rsboApp.StatusBar.SetText($"Error al actualizar número de cheque: {nErr} - {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    Else
                        rsboApp.StatusBar.SetText($"No se encontro la cuenta!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            End If

            Return True

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error metodo CrearPagoMasivo " & ex.Message, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error metodo CrearPagoMasivo..." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return False
        End Try
    End Function

    Public Function CrearPagoMasivoDesdeAprobacion(ByVal Solicitudes As List(Of Integer)) As Boolean
        Try
            rsboApp.StatusBar.SetText("Procesando Pagos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Utilitario.Util_Log.Escribir_Log("Procesando Pagos...", "frmPagosMasivos")

            Dim ListaDePagos As New List(Of Entidades.PagoMasivo)

            Try
                For Each DE As Integer In Solicitudes
                    Query = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        Query = "SELECT A.""DocEntry"", A.""U_Cuenta"", B.""U_DocEntry"", B.""U_CodProv"", B.""U_Proveedor"", B.""U_Vencimiento"", B.""U_FechaVen"", B.""U_Monto"", "
                        Query += "B.""U_Saldo"", B.""U_Pago"", A.""U_NumChe"", A.""U_MedioPago"", B.""U_Cuota"", A.""U_Banco"" FROM ""@SS_PM_CAB"" A INNER JOIN ""@SS_PM_DET1"" B "
                        Query += "ON A.""DocEntry"" = B.""DocEntry"" WHERE A.""DocEntry"" = " & DE.ToString
                    Else
                        Query = "SELECT A.DocEntry, A.U_Cuenta, B.U_DocEntry, B.U_CodProv, B.U_Proveedor, B.U_Vencimiento, B.U_FechaVen, B.U_Monto, "
                        Query += "B.U_Saldo, B.U_Pago, A.U_NumChe, A.U_MedioPago, B.U_Cuota,  A.U_Banco FROM ""@SS_PM_CAB"" A WITH(NOLOCK) INNER JOIN ""@SS_PM_DET1"" B With(NOLOCK) "
                        Query += "ON A.DocEntry = B.DocEntry WHERE A.DocEntry = " & DE.ToString
                    End If
                    Utilitario.Util_Log.Escribir_Log("Query para obtener solicitudes: " & Query.ToString, "frmPagosMasivos")

                    Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rst.DoQuery(Query)

                    If rst.RecordCount >= 1 Then
                        While rst.EoF = False
                            Dim item As New Entidades.PagoMasivo
                            item.DocEntryFP = CInt(rst.Fields.Item("U_DocEntry").Value)
                            item.CodPro = rst.Fields.Item("U_CodProv").Value
                            item.NomPro = rst.Fields.Item("U_Proveedor").Value
                            item.Vencimiento = rst.Fields.Item("U_Vencimiento").Value
                            item.FechaVencimiento = rst.Fields.Item("U_FechaVen").Value
                            item.Monto = rst.Fields.Item("U_Monto").Value
                            item.Saldo = rst.Fields.Item("U_Saldo").Value
                            item.Pagar = rst.Fields.Item("U_Pago").Value
                            item.Cuenta = rst.Fields.Item("U_Cuenta").Value
                            item.DocEntry = CInt(rst.Fields.Item("DocEntry").Value)
                            item.NumCheque = rst.Fields.Item("U_NumChe").Value
                            item.MedioPago = rst.Fields.Item("U_MedioPago").Value
                            item.Cuota = CInt(rst.Fields.Item("U_Cuota").Value)
                            item.Banco = rst.Fields.Item("U_Banco").Value
                            ListaDePagos.Add(item)
                            rst.MoveNext()
                        End While
                    End If
                Next
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error obteniendo solicitudes de pago: " & ex.Message, "frmPagosMasivos")
            End Try

            Dim ListaAgrupada = (From a In ListaDePagos Group By a.CodPro, a.DocEntry, a.Cuenta, a.NumCheque, a.MedioPago, a.Banco, a.DocEntryFP Into Group).ToList

            If Not ListaAgrupada Is Nothing Then

                rCompany.StartTransaction()

                For Each a As Object In ListaAgrupada
                    Dim nErr As Long
                    Dim errMsg As String, Qr As String = ""

                    Dim pag As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

                    Try
                        pag.DocType = SAPbobsCOM.BoRcptTypes.rSupplier
                        pag.DocDate = Date.Now
                        pag.DueDate = Date.Now
                        pag.CardCode = a.CodPro
                        pag.TaxDate = Date.Now
                        pag.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES
                        pag.DocCurrency = "USD"
                        pag.UserFields.Fields.Item("U_DE_PM").Value = a.DocEntry.ToString

                        'Dim i As Integer = 0
                        Dim acumulador As Double = 0

                        For Each b As Object In a.Group
                            'pag.Invoices.DocLine = i
                            pag.Invoices.DocEntry = a.DocEntryFP
                            pag.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                            pag.Invoices.SumApplied = Convert.ToDecimal(b.Pagar.ToString.Replace(".", ","))
                            pag.Invoices.InstallmentId = CInt(b.Cuota)
                            pag.Invoices.Add()
                            'i += 1
                            acumulador += Convert.ToDecimal(b.Pagar.ToString.Replace(".", ","))
                        Next

                        If a.MedioPago = "Cheque" Then

                            Query = ""
                            If Functions.VariablesGlobales._ManejoCuenta = "Mapeada" Then
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    Query = My.Resources.SS_CM_CONSULTAPARAMETROSBC_HANA & " Where ""U_CodBco"" = '" & a.Banco.ToString & "' AND ""U_Cuenta"" = '" & a.Cuenta.ToString & "'"
                                Else
                                    Query = My.Resources.SS_CM_CONSULTAPARAMETROSBC_SQL & " Where U_CodBco = '" & a.Banco.ToString & "' AND U_Cuenta = '" & a.Cuenta.ToString & "'"
                                End If
                            Else
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    Query = "SELECT TOP 1 ""Account"" FROM ""DSC1"" WHERE ""GLAccount"" = '" & a.Cuenta.ToString & "'"
                                Else
                                    Query = "SELECT TOP 1 Account FROM DSC1 WITH(NOLOCK) WHERE GLAccount = '" & a.Cuenta.ToString & "'"
                                End If
                            End If

                            Utilitario.Util_Log.Escribir_Log("Query para validar parametros de cheque: " & Query.ToString, "frmPagosMasivos")

                            Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rst.DoQuery(Query)
                            If rst.RecordCount >= 1 Then
                                If Functions.VariablesGlobales._ManejoCuenta = "Mapeada" Then
                                    While (rst.EoF = False)
                                        pag.Checks.AccounttNum = rst.Fields.Item("U_CtaBco").Value.ToString ' "2100042798"
                                        rst.MoveNext()
                                    End While
                                Else
                                    While (rst.EoF = False)
                                        pag.Checks.AccounttNum = rst.Fields.Item("Account").Value.ToString ' "2100042798"
                                        rst.MoveNext()
                                    End While
                                End If
                            End If

                            pag.Checks.BankCode = a.Banco
                            pag.Checks.CheckSum = a.group.sum(Function(b) b.Pagar) 'acumulador
                            pag.Checks.DueDate = Date.Now
                            pag.Checks.CheckNumber = CInt(a.NumCheque)

                            pag.Checks.Add()
                        ElseIf a.MedioPago = "Transferencia" Then
                            pag.TransferAccount = a.Cuenta.ToString '"_SYS00000000043"
                            pag.TransferSum = acumulador
                            pag.TransferDate = Date.Now
                        End If

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error campos de cabecera de pago: " + ex.Message.ToString(), "frmPagosMasivos")
                        Return False
                    End Try

                    If pag.Add = 0 Then
                        Dim DocEntry As String = ""
                        rCompany.GetNewObjectCode(DocEntry)
                        rsboApp.StatusBar.SetText("Pago efectuado creado correctamente: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Utilitario.Util_Log.Escribir_Log("Pago efectuado creado correctamente: " & DocEntry, "frmPagosMasivos")
                    Else
                        rCompany.GetLastError(nErr, errMsg)
                        rsboApp.StatusBar.SetText("Error al integrar pago: " + Str(nErr).ToString + " - " + errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Utilitario.Util_Log.Escribir_Log("Error al integrar pago: " + Str(nErr).ToString + " - " + errMsg, "frmPagosMasivos")
                        rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        Return False
                    End If
                Next
                rCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Return True
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error CrearPagoMasivoDesdeAprobacion " & ex.Message, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error CrearPagoMasivoDesdeAprobacion " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function ActualizadoSolicitudDePago(ByVal DocEntryUDOSP As String, ByVal NivelAprobacion As String, ByVal Estado As String, Optional IdCashBanco As String = "", Optional RutaArchivoDeBanco As String = "", Optional DocEntryPagCon As String = "", Optional LineId As String = "0", Optional DocEntryPagTran As String = "", Optional DocEntryNotDeb As String = "", Optional FechaArchivoBanco As String = "", Optional FechaNotaDebito As String = "", Optional RutaArchivoGenerado As String = "", Optional NumCheque As Integer = 0) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oChildren As SAPbobsCOM.GeneralDataCollection

        Try
            rsboApp.StatusBar.SetText("Actualizando Solicitud de pago: " & DocEntryUDOSP, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSMTPAGOS")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUDOSP)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            If CInt(NivelAprobacion) < CInt(NivelMaximo) And Estado = "Aprobado" Then
                oGeneralData.SetProperty("U_NivelAprob", CInt(NivelAprobacion) + 1)
            ElseIf CInt(NivelAprobacion) = CInt(NivelMaximo) And Estado = "Aprobado" Then
                oGeneralData.SetProperty("U_Estado", "Aprobado")
            ElseIf Estado = "Rechazado" Or Estado = "Modificar" Or Estado = "Archivo Generado" Or Estado = "Archivo Procesado Banco" Then

                oGeneralData.SetProperty("U_Estado", If(Not String.IsNullOrEmpty(IdCashBanco) And txtEstado.Value = "Archivo Generado", "Archivo Procesado Banco", Estado))
            End If

            If IdCashBanco <> "" Then oGeneralData.SetProperty("U_IdCashBan", IdCashBanco)
            If RutaArchivoDeBanco <> "" Then oGeneralData.SetProperty("U_RutArcBan", RutaArchivoDeBanco)
            If RutaArchivoGenerado <> "" Then oGeneralData.SetProperty("U_RutArcGen", RutaArchivoGenerado)
            If DocEntryPagCon <> "" Then oGeneralData.SetProperty("U_IdPagCon", DocEntryPagCon)

            If FechaArchivoBanco <> "" Then
                Dim FechaArchivoBco As Date = Date.ParseExact(FechaArchivoBanco, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                oGeneralData.SetProperty("U_FechaArcRec", FechaArchivoBco)
            End If

            If FechaNotaDebito <> "" Then
                Dim FechaND As Date = Date.ParseExact(FechaNotaDebito, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                oGeneralData.SetProperty("U_FechaDev", FechaND)
            End If

            oChildren = oGeneralData.Child("SS_PM_DET1")
            Dim matrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific

            If CInt(LineId) = "0" Then
                For i As Integer = 0 To oChildren.Count - 1
                    ActualizarLinea(oChildren.Item(i), matrix, i + 1)
                Next
            Else
                Dim NumLineaEditar = ObtenerNumLinea(matrix, LineId)
                If NumLineaEditar > 0 Then
                    Dim child As SAPbobsCOM.GeneralData = oChildren.Item(NumLineaEditar - 1)
                    If DocEntryPagTran <> "" Then child.SetProperty("U_IdPagTran", DocEntryPagTran)
                    If DocEntryNotDeb <> "" Then child.SetProperty("U_IdNotDeb", DocEntryNotDeb)
                    If NumCheque <> 0 Then child.SetProperty("U_NumChe", CStr(NumCheque))
                    ActualizarLinea(child, matrix, NumLineaEditar)
                End If
            End If
            oGeneralService.Update(oGeneralData)
            rsboApp.StatusBar.SetText("Solicitud de pago actualizada con éxito ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error actualizando solicitud de pago: " & DocEntryUDOSP & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Error actualizando solicitud de pago: " & DocEntryUDOSP & "-" & ex.Message, "frmPagosMasivos")
            Return False
        End Try
    End Function

    Public Function ActualizaLineaSolicitudPagoConOrdenPago(ByVal DocEntryUDOSP As String, Optional LineId As String = "0", Optional OrdPag As String = "") As Boolean
        Try
            rsboApp.StatusBar.SetText("Actualizando campo de orden de pago en la solicitud de pago masivo:" & DocEntryUDOSP, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oChildren As SAPbobsCOM.GeneralDataCollection

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSMTPAGOS")
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEntryUDOSP)
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oChildren = oGeneralData.Child("SS_PM_DET1")
            Dim child As SAPbobsCOM.GeneralData = oChildren.Item(CInt(LineId) - 1)
            child.SetProperty("U_IdOrdPag", CStr(OrdPag))
            oGeneralService.Update(oGeneralData)

            rsboApp.StatusBar.SetText("Campo actualizado con éxito!!!" & DocEntryUDOSP, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error actualizando solicitud de pago: " & DocEntryUDOSP & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Function

    Private Sub ActualizarLinea(ByRef child As SAPbobsCOM.GeneralData, ByVal matrix As SAPbouiCOM.Matrix, ByVal row As Integer)
        Dim oCheckBox As SAPbouiCOM.CheckBox = CType(matrix.Columns.Item("cbxProc").Cells.Item(row).Specific, SAPbouiCOM.CheckBox)
        child.SetProperty("U_Procesada", If(oCheckBox.Checked, "Y", "N"))
        child.SetProperty("U_Comentario", matrix.Columns.Item("U_Coment").Cells.Item(row).Specific.Value.ToString())
    End Sub

    Private Function ObtenerNumLinea(ByVal matrix As SAPbouiCOM.Matrix, ByVal LineId As String) As Integer
        For i As Integer = 1 To matrix.RowCount
            If matrix.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString() = LineId Then
                Return i
            End If
        Next
        Return 0
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

    Private Sub FormatoFacturas()
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            oGrid.Columns.Item(0).Description = "Documento"
            oGrid.Columns.Item(0).TitleObject.Caption = "Documento"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(1).Description = ""
            oGrid.Columns.Item(1).TitleObject.Caption = ""

            oGrid.Columns.Item(2).Description = "" 'DocEntry
            oGrid.Columns.Item(2).TitleObject.Caption = ""
            oGrid.Columns.Item(2).Editable = False
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = 18

            oGrid.Columns.Item(3).Description = "# Factura"
            oGrid.Columns.Item(3).TitleObject.Caption = "# Factura"
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(3).TitleObject.Sortable = True

            oGrid.Columns.Item(4).Description = "" 'DocEntry
            oGrid.Columns.Item(4).TitleObject.Caption = ""
            oGrid.Columns.Item(4).Editable = False
            Dim oEditTextColumn2 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(4)
            oEditTextColumn2.LinkedObjectType = 22

            oGrid.Columns.Item(5).Description = "# Pedido"
            oGrid.Columns.Item(5).TitleObject.Caption = "# Pedido"
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(5).TitleObject.Sortable = True

            oGrid.Columns.Item(6).Description = "Código Proveedor"
            oGrid.Columns.Item(6).TitleObject.Caption = "Código Proveedor"
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            Dim oEditTextColumn3 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(6)
            oEditTextColumn3.LinkedObjectType = 2

            oGrid.Columns.Item(7).Description = "Nombre Proveedor"
            oGrid.Columns.Item(7).TitleObject.Caption = "Nombre Proveedor"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Sortable = True

            oGrid.Columns.Item(8).Description = "Fec Fact"
            oGrid.Columns.Item(8).TitleObject.Caption = "Fec Fact"
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(8).TitleObject.Sortable = True

            'seccion de banco, cuenta pos 9, 10
            oGrid.Columns.Item(9).Description = "Bco Prov" '9
            oGrid.Columns.Item(9).TitleObject.Caption = "Bco Prov."
            oGrid.Columns.Item(9).Editable = False
            oGrid.Columns.Item(9).TitleObject.Sortable = True

            oGrid.Columns.Item(10).Description = "Cta Prov" '10
            oGrid.Columns.Item(10).TitleObject.Caption = "Cta Prov."
            oGrid.Columns.Item(10).Editable = False
            oGrid.Columns.Item(10).TitleObject.Sortable = True

            oGrid.Columns.Item(11).Description = "Sucursal"
            oGrid.Columns.Item(11).TitleObject.Caption = "Sucursal"
            oGrid.Columns.Item(11).Editable = False
            oGrid.Columns.Item(11).TitleObject.Sortable = True

            oGrid.Columns.Item(12).Description = "Proyecto"
            oGrid.Columns.Item(12).TitleObject.Caption = "Proyecto"
            oGrid.Columns.Item(12).Editable = False

            oGrid.Columns.Item(13).Description = "Cuota"
            oGrid.Columns.Item(13).TitleObject.Caption = "Cuota"
            oGrid.Columns.Item(13).Editable = False

            oGrid.Columns.Item(14).Description = "Días V."
            oGrid.Columns.Item(14).TitleObject.Caption = "Días V."
            oGrid.Columns.Item(14).Editable = False
            oGrid.Columns.Item(14).TitleObject.Sortable = True

            oGrid.Columns.Item(15).Description = "Fec Venc"
            oGrid.Columns.Item(15).TitleObject.Caption = "Fec Venc"
            oGrid.Columns.Item(15).Editable = False
            oGrid.Columns.Item(15).TitleObject.Sortable = True

            oGrid.Columns.Item(16).Description = "Total"
            oGrid.Columns.Item(16).TitleObject.Caption = "Total"
            oGrid.Columns.Item(16).Editable = False
            oGrid.Columns.Item(16).RightJustified = True
            oGrid.Columns.Item(16).TitleObject.Sortable = True

            oGrid.Columns.Item(17).Description = "Saldo"
            oGrid.Columns.Item(17).TitleObject.Caption = "Saldo"
            oGrid.Columns.Item(17).Editable = False
            oGrid.Columns.Item(17).RightJustified = True
            oGrid.Columns.Item(17).TitleObject.Sortable = True

            oGrid.Columns.Item(18).Description = "ObjType"
            oGrid.Columns.Item(18).TitleObject.Caption = "ObjType"
            oGrid.Columns.Item(18).Visible = False

            oGrid.Columns.Item(19).Description = "Observación"
            oGrid.Columns.Item(19).TitleObject.Caption = "Observación"
            oGrid.Columns.Item(19).Editable = False

            oGrid.Columns.Item(20).Description = "Cod Banco"
            oGrid.Columns.Item(20).TitleObject.Caption = "Cod Banco"
            oGrid.Columns.Item(20).Editable = False
            oGrid.Columns.Item(20).Visible = False

            oGrid.Columns.Item(21).Description = "Tipo Cuenta"
            oGrid.Columns.Item(21).TitleObject.Caption = "Tipo Cuenta"
            oGrid.Columns.Item(21).Editable = False
            oGrid.Columns.Item(21).Visible = False

            'oGrid.Columns.Item(20).Description = "Bco Prov" '9
            'oGrid.Columns.Item(20).TitleObject.Caption = "Bco Prov."
            'oGrid.Columns.Item(20).Editable = False
            'oGrid.Columns.Item(20).TitleObject.Sortable = True

            'oGrid.Columns.Item(21).Description = "Cta Prov" '10
            'oGrid.Columns.Item(21).TitleObject.Caption = "Cta Prov."
            'oGrid.Columns.Item(21).Editable = False
            'oGrid.Columns.Item(21).TitleObject.Sortable = True

            oGrid.CollapseLevel = 1
            'oGrid.AutoResizeColumns()
            oGrid.Columns.Item(2).Width = 16
            oGrid.Columns.Item(8).Width = 56
            oGrid.Columns.Item(11).Width = 40
            oGrid.Columns.Item(13).Width = 56
            oGrid.Columns.Item(20).Width = 40

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error FormatoFacturas: {ex.Message}", "frmPagosMasivos")
            rsboApp.StatusBar.SetText($"Error FormatoFacturas: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function Imprimir(ByVal DocEntryUDO As Integer)
        Dim SeccionError As String = ""
        Try
            Dim cbxfpago As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxfpago").Specific
            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim crReport As CrystalDecisions.CrystalReports.Engine.ReportDocument = Nothing
            Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific

            If cbxfpago.Value = "Transferencia" Then

                CargaFormato(DocEntryUDO)
                DocEntryUDO = 0

            ElseIf cbxfpago.Value = "Cheque" Then

                'ADD valido si existe orden de pago, caso contrario lo creo y actualizo
                Dim QueryValidaOP As String = "", NumRegistro As String = ""
                QueryValidaOP = "SELECT COUNT(*) AS ""NumRegistro"" FROM ""@SS_PM_OP_CAB"" WHERE ""U_SolicitudPago"" = '" & txtDocEnt.Value & "'"
                NumRegistro = oFuncionesAddon.getRSvalue(QueryValidaOP, "NumRegistro", "0")
                If CInt(NumRegistro) = 0 Then CrearOrdenDePago(txtDocEnt.Value)
                '
                Dim txtEstado As SAPbouiCOM.EditText = oForm.Items.Item("txtEstado").Specific
                    If (CInt(NivelUsuario) = 0 And txtEstado.Value = "Revision") Or (CInt(NivelUsuario) = CInt(NivelMaximo) And txtEstado.Value = "Aprobado") Then

                        Dim menu As Object
                        menu = oFuncionesB1.ObtenerUIDMenu("OrdCom", Functions.VariablesGlobales._RutaArchivoRPTPM) '"13056")

                        If menu <> "" Then
                            For Each f As SAPbouiCOM.Form In rsboApp.Forms
                                If f.TypeEx = "410000100" Then f.Close()
                            Next

                            rsboApp.ActivateMenuItem(menu)

                            Dim forpara As SAPbouiCOM.Form = rsboApp.Forms.GetForm("410000100", 0)
                            forpara.Select()

                            TryCast(forpara.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = DocEntryUDO.ToString
                            TryCast(forpara.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                            forpara.Visible = False

                        Else
                            rsboApp.StatusBar.SetText("No se encontró formato de ordenes de pago, por favor consulte con el administrador del sistema!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                    ElseIf CInt(NivelUsuario) = 0 And txtEstado.Value = "Archivo Procesado Banco" Then
                        ofrmImprimir.CargaFormularioImprimir(DocEntryUDO)
                    End If
                End If
        Catch ex As Exception
            rsboApp.SetStatusBarMessage($"Seccion de error: {SeccionError} excepcion: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Function

    Private Sub Event_MatrixLinkPressed(ByVal pVal As SAPbouiCOM.ItemEvent)
        Try
            If pVal.FormTypeEx = "frmPagosMasivos" Then
                Select Case pVal.ItemUID
                    Case "oGrid"
                        Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                        Dim oObjType As String = oGrid.DataTable.GetValue("ObjType", oGrid.GetDataTableRowIndex(pVal.Row))
                        Dim oColumns As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("DocEntry")
                        Select Case oObjType
                            Case "18"
                                oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                            Case "204"
                                oColumns.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments
                        End Select
                    Case "MTX_SER"
                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
                        Dim oObjType As String = oMatrix.Columns.Item("U_ObjType").Cells.Item(pVal.Row).Specific.Value.ToString
                        Dim oLinkedButton As SAPbouiCOM.LinkedButton = oMatrix.Columns.Item("U_DocEntry").ExtendedObject
                        Select Case oObjType
                            Case "18"
                                oLinkedButton.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                            Case "204"
                                oLinkedButton.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments
                        End Select
                End Select
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"ex Event_MatrixLinkPressed: {ex.Message}", "frmPagosMasivos")
            rsboApp.StatusBar.SetText($"ex Event_MatrixLinkPressed: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function GeneraArchivo(ByVal DocEntryPM As String, ByVal Cuenta As SAPbouiCOM.ComboBox) As Boolean
        Try
            Dim sQuery As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "CALL " & rCompany.CompanyDB & ".SS_CM_CONSULTAPAGOEF ('','','','" & DocEntryPM & "')"
            Else
                sQuery = "EXEC SS_CM_CONSULTAPAGOEF '','','','" & DocEntryPM & "'"
            End If

            Dim banco() = Cuenta.Selected.Description.Split(":")
            Dim Sucursal As String = ""
            Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet(sQuery)

            If oRecordSet.RecordCount > 0 Then

                Dim CompanyService As SAPbobsCOM.CompanyService = rCompany.GetCompanyService
                Dim DatosEmpresa As SAPbobsCOM.CompanyInfo = CompanyService.GetCompanyInfo()

                Dim errores As List(Of String) = ValidarDatosConErrores(CInt(banco(0).ToString), oRecordSet)

                Utilitario.Util_Log.Escribir_Log("Moviendo recordset al inicio", "frmPagosMasivos")
                oRecordSet.MoveFirst() 'Dado que ya recorrimos el recordset para la validacion, procedemos a mover el cursor al principio
                Utilitario.Util_Log.Escribir_Log("Superado proceso de movimiento de recordset", "frmPagosMasivos")

                If errores.Count > 0 Then
                    For Each ex As String In errores
                        Utilitario.Util_Log.Escribir_Log("Revisar campo: " & ex, "frmPagosMasivos")
                        rsboApp.StatusBar.SetText("Revisar campo: " & ex, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Next
                    Return False
                Else
                    Dim rutaBase As String = ""
                    Dim fecha As String = "", nombreArchivo As String = ""
                    fecha = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")

                    'Add 13/02/2025 JP
                    Dim NumeroDeCash As String = ""
                    Try
                        Dim QueryNumeroCash As String = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QueryNumeroCash = $"SELECT CAST(LPAD(COUNT(*) + 1,2,'0') AS VARCHAR) AS ""NumeroCash"" FROM ""@SS_PM_CAB"" WHERE ""U_MedioPago"" = 'Transferencia' AND ""U_Estado"" = 'Archivo Generado' AND ""CreateDate"" = '{DateTime.Now.ToString("yyyyMMdd")}'
"
                        Else
                            QueryNumeroCash = $"SELECT RIGHT(REPLICATE('0', 2) + CAST(COUNT(*) + 1 AS VARCHAR), 2) AS ""NumeroCash"" FROM ""@SS_PM_CAB"" WHERE ""U_MedioPago"" = 'Transferencia' AND ""U_Estado"" = 'Archivo Generado' AND ""CreateDate"" = '{DateTime.Now.ToString("yyyyMMdd")}'"
                        End If
                        Utilitario.Util_Log.Escribir_Log("Query validar numero de cash: " + QueryNumeroCash.ToString, "frmPagosMasivos")
                        NumeroDeCash = oFuncionesAddon.getRSvalue(QueryNumeroCash, "NumeroCash", "")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error Obteniendo Numero de cash: " + ex.Message.ToString, "frmPagosMasivos")
                    End Try


                    If TipoSolicitudPMTransferencia = "Nomina" Then
                        rutaBase = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & DatosEmpresa.CompanyName & "_ArchivoBanco"
                        fecha = DateTime.Now.ToString("yyyyMMdd")
                        nombreArchivo = "NCR" & fecha & "CB8_" & IIf(String.IsNullOrEmpty(NumeroDeCash), DocEntryPM, NumeroDeCash) & ".txt"
                    Else
                        rutaBase = Functions.VariablesGlobales._RutaReposCM & "\" & DatosEmpresa.CompanyName

                        If CInt(banco(0)) = 17 Then
                            fecha = DateTime.Now.ToString("yyyyMMdd")
                            nombreArchivo = "PAGOS_MULTICASH_" & fecha & "_" & IIf(String.IsNullOrEmpty(NumeroDeCash), DocEntryPM, NumeroDeCash) & ".txt"
                        Else
                            fecha = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")
                            nombreArchivo = DatosEmpresa.CompanyName & "_" & DocEntryPM & "_" & banco(1).ToString.Replace("*", "").Replace("?", "") & "_" & fecha.Replace(":", "") & ".txt"
                        End If

                    End If

                    Utilitario.Util_Log.Escribir_Log("Ruta base: " & rutaBase, "frmPagosMasivos")
                    Utilitario.Util_Log.Escribir_Log("Nombre archivo: " & nombreArchivo, "frmPagosMasivos")

                    Dim ruta As String = Path.Combine(rutaBase, nombreArchivo)

                    Utilitario.Util_Log.Escribir_Log("Generando archivo en la siguiente ruta: " & ruta, "frmPagosMasivos")
                    rsboApp.StatusBar.SetText("Generando archivo en la siguiente ruta: " & ruta, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)

                    If Not Directory.Exists(rutaBase) Then 'Si no existe carpeta la creamos
                        Directory.CreateDirectory(rutaBase)
                    End If

                    Dim strStreamW As Stream = Nothing
                    Dim strStreamWriter As StreamWriter = Nothing

                    If Not File.Exists(ruta) Then
                        strStreamW = File.Create(ruta) ' lo creamos
                        strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default) '
                        strStreamWriter.Close() ' cerramos
                    Else
                        File.Delete(ruta) ' lo eliminamos
                        strStreamW = File.Create(ruta) ' lo creamos
                        strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default) '
                        strStreamWriter.Close() ' cerramos
                    End If

                    Utilitario.Util_Log.Escribir_Log("Paso proceso de creacion de archivo...", "frmPagosMasivos")
                    rsboApp.StatusBar.SetText("Paso proceso de creacion de archivo... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)

                    Dim sTexto As New StringBuilder
                    Dim CampoVacio As String = ""

                    If oRecordSet.RecordCount >= 1 Then
                        While (oRecordSet.EoF = False)
                            Dim DocTotalRS As Double = Convert.ToDouble(oRecordSet.Fields.Item("DocTotal").Value)
                            Dim DocTotalConv As String = DocTotalRS.ToString("F2")
                            Dim NombreProv As String = oRecordSet.Fields.Item("CardName").Value

                            Select Case CInt(banco(0).ToString)
                                Case 10 '10 : Banco Pichincha

                                    sTexto.AppendLine(oRecordSet.Fields.Item("Metodo").Value.ToString & vbTab & 'Codigo de orientacion 1
                                                  oRecordSet.Fields.Item("CardCode").Value.ToString & vbTab & 'Contrapartida 1
                                                  oRecordSet.Fields.Item("Moneda").Value.ToString & vbTab & 'Moneda 1
                                 Right("0000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 13) & vbTab & 'Valor 1
                                 oRecordSet.Fields.Item("FormaPago").Value.ToString & vbTab & 'Forma de pago 1
                                 oRecordSet.Fields.Item("TipoCuenta").Value.ToString & vbTab & 'Tipo de cuenta 1
                                 oRecordSet.Fields.Item("NumeroCuenta").Value.ToString & vbTab & 'numero de cuenta 1
                                 oRecordSet.Fields.Item("Referencia").Value.ToString & vbTab & 'Referencia 0
                                 oRecordSet.Fields.Item("TipoCliente").Value.ToString & vbTab & 'Tipo ID Cliente 1
                                 oRecordSet.Fields.Item("NumeroID").Value.ToString & vbTab & 'Numero ID Cliente 1
                                 NombreProv.Replace("ñ", "n").Replace("Ñ", "N") & vbTab & 'Nombre del cliente 1
                                 oRecordSet.Fields.Item("CodBanco").Value.ToString) 'Codigo de banco 0

                                Case 36 '36 : Banco Produbanco

                                    sTexto.AppendLine(oRecordSet.Fields.Item("Metodo").Value.ToString & vbTab & 'Codigo de orientacion 1
                                                  Right("00000000000" & oRecordSet.Fields.Item("CodCuentaEmpresa").Value.ToString, 11) & vbTab & 'Cuenta empresa 1
                                 oRecordSet.Fields.Item("DocNum").Value.ToString & vbTab & 'Secuencial pago 1
                                 vbTab & 'Comprobante de pago 0
                                 oRecordSet.Fields.Item("CardCode").Value.ToString & vbTab & 'Contrapartida 1
                                 oRecordSet.Fields.Item("Moneda").Value.ToString & vbTab & 'Moneda 1
                                 Right("0000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 13) & vbTab & 'Valor 1
                                 oRecordSet.Fields.Item("FormaPago").Value.ToString & vbTab & 'Forma de pago 1
                                 Right("0000" & oRecordSet.Fields.Item("CodBancoEmpresa").Value.ToString, 4) & vbTab & 'Codigo de institucion financiera 1
                                 IIf(oRecordSet.Fields.Item("FormaPago").Value.ToString = "CTA", oRecordSet.Fields.Item("TipoCuenta").Value.ToString & vbTab, vbTab) & 'Tipo de cuenta 1
                                 IIf(oRecordSet.Fields.Item("FormaPago").Value.ToString = "CTA", oRecordSet.Fields.Item("NumeroCuenta").Value.ToString & vbTab, vbTab) & 'Numero de cuenta 1
                                 oRecordSet.Fields.Item("TipoCliente").Value.ToString & vbTab & 'Tipo ID Cliente 1
                                 oRecordSet.Fields.Item("NumeroID").Value.ToString & vbTab & 'Numero ID Cliente 1
                                 Right(NombreProv.Replace("ñ", "n").Replace("Ñ", "N"), 60) & vbTab & 'Nombre del cliente 1
                                 vbTab & 'Direccion 0
                                 vbTab & 'Ciudad 0
                                 vbTab & 'Telefono 0
                                 vbTab & 'Localidad de pago 0
                                 oRecordSet.Fields.Item("FactReferencia").Value.ToString & vbTab & 'Referencia 1
                                 oRecordSet.Fields.Item("Referencia").Value.ToString & IIf(oRecordSet.Fields.Item("Email").Value.ToString <> "", "| " & oRecordSet.Fields.Item("Email").Value.ToString, "")) 'Ref adicional 0

                                Case 17 '17 : Banco Guayaquil

                                    If TipoSolicitudPMTransferencia = "Nomina" Then
                                        sTexto.AppendLine(IIf(oRecordSet.Fields.Item("TipoCliente").Value.ToString = "CTA", "A", "C") & 'Tipo cuenta
                                             Right("0000000000" & (oRecordSet.Fields.Item("NumeroCuenta").Value.ToString), 10) & 'Numero cuenta
                                             Right("000000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 15) & 'Valor
                                             "XX" & 'Motivo
                                             "Y" & 'Credito
                                             "01" & 'Agencia
                                             "  " & 'Espacio destinado para codigo de otro banco
                                             "                  " & 'Espacio destinado para numero de cuento para pago a otro banco
                                             Left(NombreProv.Replace("ñ", "n").Replace("Ñ", "N"), 18) & "CB8") 'Fomrato final de nombre de titular de seruvi
                                    Else
                                        sTexto.AppendLine(
                                             oRecordSet.Fields.Item("Metodo").Value.ToString & vbTab & 'Codigo de orientacion 1
                                            Right("0000000000" & oRecordSet.Fields.Item("CodCuentaEmpresa").Value.ToString, 10) & vbTab & 'Cuenta empresa 1
                                            Right("0000000" & oRecordSet.Fields.Item("DocNum").Value.ToString, 7) & vbTab & 'Secuencial pago 1
                                            vbTab & 'Comprobante de pago 0
                                            oRecordSet.Fields.Item("CardCode").Value.ToString & vbTab & 'Contrapartida 1
                                            oRecordSet.Fields.Item("Moneda").Value.ToString & vbTab & 'Moneda 1
                                            Right("0000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 13) & vbTab & 'Valor 1
                                            oRecordSet.Fields.Item("FormaPago").Value.ToString & vbTab & 'Forma de pago 1
                                            Right("0000" & oRecordSet.Fields.Item("CodBancoEmpresa").Value.ToString, 4) & vbTab & 'Codigo de institucion financiera 1
                                            IIf(oRecordSet.Fields.Item("FormaPago").Value.ToString = "CTA", oRecordSet.Fields.Item("TipoCuenta").Value.ToString & vbTab, vbTab) & 'Tipo de cuenta 1
                                            IIf(oRecordSet.Fields.Item("FormaPago").Value.ToString = "CTA", Right("00000000000" & (oRecordSet.Fields.Item("NumeroCuenta").Value.ToString), 11) & vbTab, vbTab) & 'Numero de cuenta 1
                                            oRecordSet.Fields.Item("TipoCliente").Value.ToString & vbTab & 'Tipo ID Cliente 1
                                            oRecordSet.Fields.Item("NumeroID").Value.ToString & vbTab & 'Numero ID Cliente 1
                                            Right(NombreProv.Replace("ñ", "n").Replace("Ñ", "N"), 40) & vbTab & 'Nombre del cliente 1
                                            vbTab & 'Direccion 0
                                            vbTab & 'Ciudad 0
                                            vbTab & 'Telefono 0
                                            vbTab & 'Localidad de pago 0
                                            oRecordSet.Fields.Item("FactReferencia").Value.ToString & vbTab & 'Referencia 1
                                            Right((oRecordSet.Fields.Item("Referencia").Value.ToString & IIf(oRecordSet.Fields.Item("Email").Value.ToString <> "", "| " & oRecordSet.Fields.Item("Email").Value.ToString, "")), 200)) 'Ref adicional 0
                                    End If

                                Case Else
                                    rsboApp.StatusBar.SetText("Archivo de banco no diseñado! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Return False
                            End Select
                            oRecordSet.MoveNext()
                        End While
                    End If

                    If oRecordSet.RecordCount = 1 Then
                        Utilitario.Util_Log.Escribir_Log("Si solo devuelve un solo registro se remplaza los saltos de lineas".ToString, "frmPagosMasivos")
                        sTexto.Replace(vbLf, "").Replace(vbCrLf, "")
                    End If

                    Try
                        Dim oTextWriter As TextWriter = New StreamWriter(ruta, True)
                        oTextWriter.WriteLine(sTexto.ToString)
                        oTextWriter.Flush()
                        oTextWriter.Close()
                        oTextWriter = Nothing
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Error escribiendo archivo: " & ex.Message.ToString, "frmPagosMasivos")
                        rsboApp.StatusBar.SetText("Error escribiendo archivo: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End Try

                    oRecordSet.MoveFirst() 'Dado que ya recorrimos el recordset para la validacion, procedemos a mover el cursor al principio

                    rsboApp.StatusBar.SetText("Archivo generado en la siguiente ruta: " & ruta, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'Agregar lógica para guardar informacion en UDO
                    Dim DocEntryUDOCM As Integer
                    GuardaUDOCM(fecha, ruta, oRecordSet, DocEntryUDOCM)

                    If DocEntryUDOCM <> 0 Then
                        ActualizadoSolicitudDePago(DocEntryPM, NivelUsuario, "Archivo Generado", "", "", "", "0", "", "", "", "", ruta)
                    End If

                    Dim lblRAG As SAPbouiCOM.StaticText = oForm.Items.Item("lblRAG").Specific
                    lblRAG.Caption = ruta.ToString
                    lblRAG.Item.ForeColor = RGB(6, 69, 173)
                    lblRAG.Item.TextStyle = 4
                End If

            End If

            Return True
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error generando archivo: " & ex.Message.ToString, "frmPagosMasivos")
            Return False
        End Try
    End Function
    Private Function ValidarDatosConErrores(ByVal Banco As Integer, ByVal datos As SAPbobsCOM.Recordset) As List(Of String)
        Dim errores As New List(Of String)
        'Dim Sucursales As New List(Of String)

        Try
            Utilitario.Util_Log.Escribir_Log("Validando datos para la generación del archivo!", "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Validando datos para la generación del archivo!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Dim i As Integer = 0
            If datos.RecordCount >= 1 Then
                While (datos.EoF = False)

                    Select Case Banco
                        Case 10
                            If String.IsNullOrEmpty(datos.Fields.Item("Metodo").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("Metodo").Value.ToString & " - Código Orientacion (Metodo)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CardCode").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("CardCode").Value.ToString & " - Contrapartida 'Cod cliente, # medidor, # telefono, # Contrato' (CardCode)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("Moneda").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("Moneda").Value.ToString & " - Moneda (Moneda)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("DocTotal").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocTotal").Value.ToString & " - Valor (DocTotal)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("FormaPago").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("FormaPago").Value.ToString & " - Forma de Pago (FormaPago)")
                            Else
                                If datos.Fields.Item("FormaPago").Value.ToString = "CTA" Then
                                    If String.IsNullOrEmpty(datos.Fields.Item("TipoCuenta").Value.ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Tipo de cuenta (TipoCuenta)")
                                    End If
                                    If String.IsNullOrEmpty(datos.Fields.Item("NumeroCuenta").Value.ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Numero de cuenta (NumeroCuenta)")
                                    End If
                                End If
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("TipoCliente").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Tipo ID Cliente (TipoCliente)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("NumeroID").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Numero ID Cliente (NumeroID)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CardName").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Nombre del cliente (CardName)")
                            End If

                            'Sucursales.Add(datos.Fields.Item("Sucursal").Value.ToString)

                        Case 36

                            If String.IsNullOrEmpty(datos.Fields.Item("Metodo").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Código Orientacion (Metodo)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CodCuentaEmpresa").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Cuenta Empresa (CodCuentaEmpresa)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("DocNum").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Secuencial pago (DocNum)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("Moneda").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Moneda (Moneda)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("DocTotal").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Valor (DocTotal)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("FormaPago").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Forma de Pago (FormaPago)")
                            Else
                                If datos.Fields.Item("FormaPago").Value.ToString = "CTA" Then
                                    If String.IsNullOrEmpty(datos.Fields.Item("TipoCuenta").Value.ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Tipo de Cuenta (TipoCuenta)")
                                    End If
                                    If String.IsNullOrEmpty(datos.Fields.Item("NumeroCuenta").Value.ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Numero de cuenta (NumeroCuenta)")
                                    End If
                                End If
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CodBancoEmpresa").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Codigo de institucion financiera (CodBancoEmpresa)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("TipoCliente").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Tipo ID Cliente (TipoCliente)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("NumeroID").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Numero ID Cliente (NumeroID)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CardName").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Nombre del cliente (CardName)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("FactReferencia").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Referencia 'Numero de factura' (FactReferencia)")
                            End If

                            'Sucursales.Add(datos.Fields.Item("Sucursal").Value.ToString)

                        Case 17

                            If String.IsNullOrEmpty(datos.Fields.Item("Metodo").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("Metodo").Value.ToString & " - Código Orientacion (Metodo)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CodCuentaEmpresa").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Cuenta Empresa (CodCuentaEmpresa)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("DocNum").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Secuencial pago (DocNum)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CardCode").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("CardCode").Value.ToString & " - Código 'Cod cliente, # medidor, # telefono, # Contrato' (CardCode)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("Moneda").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Moneda (Moneda)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("DocTotal").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Valor (DocTotal)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("FormaPago").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Forma de Pago (FormaPago)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CodBancoEmpresa").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Codigo de institucion financiera (CodBancoEmpresa)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("TipoCuenta").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Tipo de Cuenta (TipoCuenta)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("NumeroCuenta").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Numero de cuenta (NumeroCuenta)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("TipoCliente").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Tipo ID Cliente (TipoCliente)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("NumeroID").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Numero ID Cliente (NumeroID)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("CardName").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Nombre del cliente (CardName)")
                            End If
                            If String.IsNullOrEmpty(datos.Fields.Item("FactReferencia").Value.ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & datos.Fields.Item("DocNum").Value.ToString & " - Referencia 'Numero de factura' (FactReferencia)")
                            End If

                            'Sucursales.Add(datos.Fields.Item("Sucursal").Value.ToString)
                    End Select
                    i += 1
                    datos.MoveNext()
                End While

                'Dim groupedSucursales = Sucursales.GroupBy(Function(suc) suc).ToList

                'If groupedSucursales.Count() = 1 Then
                '    Sucursal = CStr(groupedSucursales(0).Key)
                'Else
                '    Sucursal = "Todos"
                'End If

                Utilitario.Util_Log.Escribir_Log("Proceso de validación de datos superado con éxito!", "frmPagosMasivos")
                rsboApp.StatusBar.SetText("Proceso de validación de datos superado con éxito!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                Utilitario.Util_Log.Escribir_Log("No se presentó información para validar!", "frmPagosMasivos")
                rsboApp.StatusBar.SetText("No se presentó información para validar!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End If
            Return errores
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error validando datos para generación de archivo: {ex.Message}", "frmPagosMasivos")
            rsboApp.StatusBar.SetText($"Error validando datos para generación de archivo: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return errores
        End Try
    End Function

    Private Function GuardaUDOCM(ByVal Fecha As String, ByVal Ruta As String, ByVal datos As SAPbobsCOM.Recordset, ByRef DocEntryUDOCM As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            rsboApp.StatusBar.SetText("Creando registro de generación de archivo UDO...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Utilitario.Util_Log.Escribir_Log("Creando registro de generación de archivo UDO...", "frmPagosMasivos")

            Dim lblFac As SAPbouiCOM.StaticText = oForm.Items.Item("lblFac").Specific
            Dim txtMP As SAPbouiCOM.EditText = oForm.Items.Item("txtMP").Specific

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSCMCASH")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            Utilitario.Util_Log.Escribir_Log("Obteniendo UDO - SSCMCASH", "frmPagosMasivos")

            Utilitario.Util_Log.Escribir_Log("U_FecArc " & Fecha.ToString, "frmPagosMasivos")
            oGeneralData.SetProperty("U_FecArc", Fecha)

            Utilitario.Util_Log.Escribir_Log("U_RutaArc " & Ruta, "frmPagosMasivos")
            oGeneralData.SetProperty("U_RutaArc", Ruta)

            Utilitario.Util_Log.Escribir_Log("U_NumPagos " & lblFac.Caption, "frmPagosMasivos")
            oGeneralData.SetProperty("U_NumPagos", CInt(lblFac.Caption))

            Utilitario.Util_Log.Escribir_Log("U_TotPagos " & txtMP.Value, "frmPagosMasivos")
            oGeneralData.SetProperty("U_TotPagos", Convert.ToDouble(txtMP.Value))

            Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific
            Dim banco As String()
            banco = cbxCuenta.Selected.Description.Split(":")

            Utilitario.Util_Log.Escribir_Log("U_Banco " & banco(0), "frmPagosMasivos")
            oGeneralData.SetProperty("U_Banco", banco(0))

            oGeneralData.SetProperty("U_Estado", "Abierto")

            oChildren = oGeneralData.Child("SS_CM_DET1")

            If datos.RecordCount >= 1 Then
                While (datos.EoF = False)
                    oChild = oChildren.Add

                    Utilitario.Util_Log.Escribir_Log("U_DocEntryP " + datos.Fields.Item("DocEntry").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_DocEntryP", datos.Fields.Item("DocEntry").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_CodProv " + datos.Fields.Item("CardCode").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_CodProv", datos.Fields.Item("CardCode").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_FecPag " + datos.Fields.Item("DocDate").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_FecPag", datos.Fields.Item("DocDate").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_TotPag " + datos.Fields.Item("DocTotal").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_TotPag", Convert.ToDouble(datos.Fields.Item("DocTotal").Value.ToString))

                    Utilitario.Util_Log.Escribir_Log("U_ForPag " + datos.Fields.Item("FormaPago").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_ForPag", datos.Fields.Item("FormaPago").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_Moneda " + datos.Fields.Item("Moneda").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_Moneda", datos.Fields.Item("Moneda").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_TipCta " + datos.Fields.Item("TipoCuenta").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_TipCta", datos.Fields.Item("TipoCuenta").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_NumCta " + datos.Fields.Item("NumeroCuenta").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_NumCta", datos.Fields.Item("NumeroCuenta").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_Ref " + datos.Fields.Item("Referencia").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_Ref", datos.Fields.Item("Referencia").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_TipCli " + datos.Fields.Item("TipoCliente").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_TipCli", datos.Fields.Item("TipoCliente").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_NumID " + datos.Fields.Item("NumeroID").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_NumID", datos.Fields.Item("NumeroID").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_CodBco " + datos.Fields.Item("CodBanco").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_CodBco", datos.Fields.Item("CodBanco").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_CodCtaEm " + datos.Fields.Item("CodCuentaEmpresa").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_CodCtaEm", datos.Fields.Item("CodCuentaEmpresa").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_CodBcoEm " + datos.Fields.Item("CodBancoEmpresa").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_CodBcoEm", datos.Fields.Item("CodBancoEmpresa").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_Email " + datos.Fields.Item("Email").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_Email", datos.Fields.Item("Email").Value.ToString)

                    Utilitario.Util_Log.Escribir_Log("U_FacRef " + datos.Fields.Item("FactReferencia").Value.ToString, "frmPagosMasivos")
                    oChild.SetProperty("U_FacRef", datos.Fields.Item("FactReferencia").Value.ToString)
                    datos.MoveNext()
                End While
            End If

            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryUDOCM = oGeneralParams.GetProperty("DocEntry")
            rsboApp.StatusBar.SetText("Se creo registro de generación de archivo UDO " & DocEntryUDOCM.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error en el registro de generación de archivo UDO: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error en el registro de generación de archivo UDO: " & ex.Message, "frmPagosMasivos")
            Return False
        End Try
    End Function

    Public Function CrearND() As Boolean

        Try
            rsboApp.StatusBar.SetText("Creando devolución de pago efectuado!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Utilitario.Util_Log.Escribir_Log("Creando devolucion de pago efectuado!", "frmPagosMasivos")

            Dim DocEntryND As String = ""
            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_SER").Specific
            Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific 'cbx Cuenta
            Dim txtDocEnt As SAPbouiCOM.EditText = oForm.Items.Item("txtDocEnt").Specific 'txt id Solicitud (DocEntry)
            'Dim txtArcBco As SAPbouiCOM.EditText = oForm.Items.Item("txtArcBco").Specific 'txt ruta de archivo de banco
            Dim lblABCO As SAPbouiCOM.StaticText = oForm.Items.Item("lblABCO").Specific

            Dim txtFDev As SAPbouiCOM.EditText = oForm.Items.Item("txtFDev").Specific
            Dim FechaND As String = txtFDev.Value.ToString
            Dim FechaNDDoc As Date = Date.ParseExact(FechaND, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            btnAprob = oForm.Items.Item("btnAprob").Specific

            For i As Integer = 1 To mMatrix.RowCount
                Dim oCheckBox As SAPbouiCOM.CheckBox = CType(mMatrix.Columns.Item("cbxProc").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                Dim checkValue As Boolean = oCheckBox.Checked

                Dim DocEntryNDMatrix As String = mMatrix.Columns.Item("U_IdND").Cells.Item(i).Specific.Value.ToString()

                If checkValue And DocEntryNDMatrix = "" Then
                    Dim oDebitNote As SAPbobsCOM.Documents
                    Dim lRetCode As Integer
                    Dim sErrMsg As String = ""

                    oDebitNote = CType(rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices), SAPbobsCOM.Documents)
                    oDebitNote.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo

                    oDebitNote.CardCode = mMatrix.Columns.Item("U_CodProv").Cells.Item(i).Specific.Value.ToString()
                    oDebitNote.DocDate = FechaNDDoc 'DateTime.Now
                    oDebitNote.Comments = mMatrix.Columns.Item("U_Coment").Cells.Item(i).Specific.Value.ToString() '"ND Creada por movimiento de cuenta en banco"   
                    oDebitNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                    oDebitNote.Lines.ItemDescription = "P/c de devolución de pago masivo, factura # " & mMatrix.Columns.Item("U_NumDoc").Cells.Item(i).Specific.Value.ToString()
                    oDebitNote.Lines.Quantity = 1
                    oDebitNote.Lines.AccountCode = cbxCuenta.Value
                    oDebitNote.Lines.LineTotal = CDbl(mMatrix.Columns.Item("U_Pag").Cells.Item(i).Specific.Value.ToString())
                    oDebitNote.Lines.TaxCode = "IVA_EXO"
                    lRetCode = oDebitNote.Add()

                    If lRetCode <> 0 Then
                        rCompany.GetLastError(lRetCode, sErrMsg)
                        Utilitario.Util_Log.Escribir_Log($"Error al crear ND {lRetCode} : {sErrMsg}", "frmPagosMasivos")
                        rsboApp.StatusBar.SetText($"Error al crear ND {lRetCode} : {sErrMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        rCompany.GetNewObjectCode(DocEntryND)
                        Utilitario.Util_Log.Escribir_Log("Nota de debito creada exitosamente! " & DocEntryND, "frmPagosMasivos")
                        rsboApp.StatusBar.SetText("Nota de debito creada exitosamente! " & DocEntryND, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'ActualizadoSolicitudDePago(txtDocEnt.Value.ToString, NivelAprobacion, "Archivo Procesado Banco", "", txtArcBco.Value, "", mMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString(), "", DocEntryND, "", FechaND)
                        ActualizadoSolicitudDePago(txtDocEnt.Value.ToString, NivelAprobacion, "Archivo Procesado Banco", "", lblABCO.Caption, "", mMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString(), "", DocEntryND, "", FechaND)
                    End If
                Else
                    rsboApp.StatusBar.SetText("Ya tiene generada una nota de debito", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            Next

            mMatrix.FlushToDataSource()
            Return True
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error Creando Nota ´débito: " & ex.Message.ToString, "frmPagosMasivos")
            rsboApp.StatusBar.SetText("Error Creando Nota ´débito: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Sub CargaFormato(DocEntry As Integer)
        Try
            Dim menu As Object

            If TipoSolicitudPMTransferencia = "Nomina" Then

                menu = oFuncionesB1.ObtenerUIDMenu("RptTransNo", "16128") '"13056" Finanzas   '16128 Recursos humanos)

                If menu <> "" Then
                    For Each f As SAPbouiCOM.Form In rsboApp.Forms
                        If f.TypeEx = "410000100" Then f.Close()
                    Next

                    rsboApp.ActivateMenuItem(menu)

                    Dim forpara As SAPbouiCOM.Form = rsboApp.Forms.GetForm("410000100", 0)
                    forpara.Select()

                    TryCast(forpara.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = DocEntry.ToString
                    TryCast(forpara.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                    forpara.Visible = False
                End If

            Else
                menu = oFuncionesB1.ObtenerUIDMenu("RptDocApr", Functions.VariablesGlobales._RutaArchivoRPTPM) '"13056")

                If menu <> "" Then
                    For Each f As SAPbouiCOM.Form In rsboApp.Forms
                        If f.TypeEx = "410000100" Then f.Close()
                    Next

                    rsboApp.ActivateMenuItem(menu)

                    Dim forpara As SAPbouiCOM.Form = rsboApp.Forms.GetForm("410000100", 0)
                    forpara.Select()

                    TryCast(forpara.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = DocEntry.ToString
                    TryCast(forpara.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                    forpara.Visible = False
                End If
            End If


        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error Cargando formatos " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
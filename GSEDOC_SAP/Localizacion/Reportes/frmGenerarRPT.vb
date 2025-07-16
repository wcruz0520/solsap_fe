Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource


Public Class frmGenerarRPT

    Private oForm As SAPbouiCOM.Form
    Private rMatrix As SAPbouiCOM.Matrix
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Public num As Integer = 0


    'Dim TipoSerie As SAPbouiCOM.ComboBox = Nothing
    'Dim IdSerie As SAPbouiCOM.EditText = Nothing
    'Dim NombreSerie As SAPbouiCOM.EditText = Nothing

    'Dim txtCliente As SAPbouiCOM.EditText = Nothing
    'Dim txtNC As SAPbouiCOM.EditText = Nothing
    'Dim txtConta As SAPbouiCOM.EditText = Nothing

    'Dim txtCod As SAPbouiCOM.EditText = Nothing
    'Dim txtSerie As SAPbouiCOM.EditText = Nothing
    'Dim txtItem As SAPbouiCOM.EditText = Nothing
    'Dim txtMaquina As SAPbouiCOM.EditText = Nothing
    'Dim txtAsesor As SAPbouiCOM.EditText = Nothing

    Dim _TipoRpt As Integer = 0

    Dim _tipoManejo As String
    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioGenerarRPT()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmGenerarRPT") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmGenerarRPT.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmGenerarRPT").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            ' _TipoRpt = TipoRpt

            oForm = rsboApp.Forms.Item("frmGenerarRPT")
            oForm.Freeze(True)

            Dim ipLogoSS As SAPbouiCOM.PictureBox
            ipLogoSS = oForm.Items.Item("logo").Specific
            ipLogoSS.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"

            'Dim NomRep As String = ""
            'If TipoRpt = 1 Then
            '    NomRep = "Facturas de Compras"
            'ElseIf TipoRpt = 2 Then
            '    NomRep = "Formulario 103"
            'ElseIf TipoRpt = 3 Then
            '    NomRep = "Listado de Ventas"
            'ElseIf TipoRpt = 4 Then
            '    NomRep = "Retenciones de Clientes"
            'End If


            Dim txtFechIni As SAPbouiCOM.EditText = oForm.Items.Item("txtFechIni").Specific
            txtFechIni.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtFechFin As SAPbouiCOM.EditText = oForm.Items.Item("txtFechFin").Specific
            txtFechFin.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim cmbRPT As SAPbouiCOM.ComboBox
            cmbRPT = oForm.Items.Item("cmbRPT").Specific
            cmbRPT.ValidValues.Add("1", "Facturas de Compras")
            cmbRPT.ValidValues.Add("2", "Formulario 103")
            cmbRPT.ValidValues.Add("3", "Listado de Ventas")
            cmbRPT.ValidValues.Add("4", "Retencion de Clientes")

            cmbRPT.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue)

            'Dim titulo = oForm.Title
            'titulo += " " + NomRep
            'oForm.Title = titulo

            oForm.Visible = True
            oForm.Select()


        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla frmGenerarRPT: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormUID = "frmGenerarRPT" AndAlso pVal.BeforeAction = True AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                'Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)
                'Dim mEdit As SAPbouiCOM.EditText = Nothing
                'Dim mItem As SAPbouiCOM.IItem = Nothing
                'Try
                '    mItem = mForm.Items.Add("Code", BoFormItemTypes.it_EDIT)
                '    mItem.Left = 460
                '    mItem.Top = 10
                '    mItem.Height = 14
                '    mItem.Width = 80
                '    mItem.Enabled = False
                '    mItem.DisplayDesc = False
                '    mEdit = mItem.Specific
                '    mEdit.DataBind.SetBound(True, "@SPV_LLASER", "Code")
                '    oFuncionesB1.Release(mItem)
                '    mForm.DataBrowser.BrowseBy = "Code" 'Next


                'Catch ex As Exception
                'Finally
                '    oFuncionesB1.Release(mForm)
                '    oFuncionesB1.Release(mEdit)
                '    oFuncionesB1.Release(mItem)
                'End Try
            ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                   And pVal.FormTypeEx = "frmGenerarRPT" Then
                If pVal.BeforeAction = False And pVal.ItemUID = "btnGRep" Then

                    Dim fechai As Date
                    Dim fechaf As Date
                    Dim _fechai As String
                    Dim _fechaf As String

                    Dim txtFechIni As SAPbouiCOM.EditText = oForm.Items.Item("txtFechIni").Specific
                    Dim txtFechfin As SAPbouiCOM.EditText = oForm.Items.Item("txtFechFin").Specific

                    If Not oFuncionesB1.BobStringToDate(txtFechIni.Value, fechai) Then
                        rsboApp.SetStatusBarMessage("El formato de las fechas es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Exit Sub
                    End If

                    If Not oFuncionesB1.BobStringToDate(txtFechfin.Value, fechaf) Then
                        rsboApp.SetStatusBarMessage("El formato de las fechas es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        oForm.Freeze(False)
                        Exit Sub
                    End If

                    _fechai = Functions.FuncionesB1._FechaSql(fechai)
                    _fechaf = Functions.FuncionesB1._FechaSql(fechaf)

                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        GenerarReporteHANA(CDate(fechai), CDate(fechaf))
                    Else
                        GenerarReporte(CDate(fechai), CDate(fechaf))
                    End If

                    'rsboApp.MessageBox("Fecha Inicio: " + _fechai.ToString + " fecha fin: " + _fechaf.ToString)
                    'rsboApp.MessageBox(Format(CDate(fechai), "MM/dd/yyyy").ToString + " " + Format(CDate(fechaf), "MM/dd/yyyy").ToString)



                End If

            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex: " + ex.Message.ToString(), "frmGenerarRPT")
            System.Windows.Forms.MessageBox.Show("Error rsboApp_ItemEvent :" & ex.Message.ToString())
        End Try


    End Sub

    'Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
    '    Try

    '        If pVal.MenuUID = "1282" And rsboApp.Forms.ActiveForm.UniqueID = "frmGenerarRPT" And pVal.BeforeAction = False Then
    '            Try
    '                Dim mEdit As SAPbouiCOM.EditText = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("Code").Specific ' obtiene el campo creado anteriormente
    '                rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("Code").Enabled = False ' le dice que el campo no sea editable
    '                mEdit.Value = CDbl(oFuncionesB1.getCorrelativo("Code", "@SPV_LLASER")) 'busca cual es el ultimo secuencial para traer esa informacion
    '                'Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
    '                'mMatrix.AddRow()
    '                'mMatrix.Columns.Item("COL1").Cells.Item(mMatrix.RowCount).Specific.String = mMatrix.RowCount
    '                'rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("COL_GRUART").Enabled = True
    '            Catch ex As Exception
    '                MsgBox(ex.Message)
    '            End Try
    '        End If

    '        If rsboApp.Forms.ActiveForm.UniqueID = "frmGenerarRPT" And pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
    '            'rsboApp.Forms.ActiveForm.Freeze(True)
    '            'Try
    '            '    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
    '            '    mMatrix.AddRow()
    '            '    For i As Integer = 1 To mMatrix.RowCount
    '            '        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
    '            '    Next

    '            '    Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
    '            '    comb3.Select("FC", BoSearchKey.psk_ByValue)


    '            '    mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
    '            '    mMatrix.Columns.Item("SerDesc").Cells.Item(mMatrix.RowCount).Specific.String = ""

    '            '    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
    '            '        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
    '            '    End If
    '            'Catch ex As Exception
    '            'Finally
    '            '    rsboApp.Forms.ActiveForm.Freeze(False)
    '            'End Try

    '        End If

    '        If rsboApp.Forms.ActiveForm.UniqueID = "frmGenerarRPT" And pVal.MenuUID = "Eliminar" And pVal.BeforeAction = False Then
    '            'Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
    '            'mMatrix.DeleteRow(num)
    '            'For i As Integer = 1 To mMatrix.RowCount
    '            '    mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
    '            'Next

    '            'If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
    '            '    rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
    '            'End If
    '        End If

    '    Catch ex As Exception
    '        Utilitario.Util_Log.Escribir_Log("ex: " + ex.Message.ToString(), "frmGenerarRPT")
    '        'System.Windows.Forms.MessageBox.Show("Error rSboApp_MenuEvent :" & ex.Message.ToString())
    '    End Try

    'End Sub

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

    'Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

    '    If eventInfo.FormUID = "frmGenerarRPT" Then

    '        If eventInfo.ItemUID = "MTX_SER" Then

    '            If eventInfo.ColUID = "COL1" Then
    '                Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
    '                Dim oMenus As SAPbouiCOM.Menus = Nothing

    '                If eventInfo.BeforeAction = True Then

    '                    Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

    '                    If mForm.Mode = BoFormMode.fm_ADD_MODE Or mForm.Mode = BoFormMode.fm_OK_MODE Then

    '                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
    '                        Try
    '                            num = eventInfo.Row

    '                            oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
    '                            oMenuItem = rsboApp.Menus.Item("1280")
    '                            If oMenuItem.SubMenus.Exists("Agregar") Then
    '                                rsboApp.Menus.RemoveEx("Agregar")

    '                            End If
    '                            If oMenuItem.SubMenus.Exists("Eliminar") Then
    '                                rsboApp.Menus.RemoveEx("Eliminar")
    '                            End If
    '                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
    '                            oCreationPackage.UniqueID = "Agregar"
    '                            oCreationPackage.String = "Agregar fila"
    '                            oCreationPackage.Enabled = True
    '                            oCreationPackage.Position = 20
    '                            oMenuItem = rsboApp.Menus.Item("1280")
    '                            oMenus = oMenuItem.SubMenus
    '                            oMenus.AddEx(oCreationPackage)

    '                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
    '                            oCreationPackage.UniqueID = "Eliminar"
    '                            oCreationPackage.String = "Eliminar fila"
    '                            oCreationPackage.Enabled = True
    '                            oCreationPackage.Position = 21
    '                            oMenuItem = rsboApp.Menus.Item("1280")
    '                            oMenus = oMenuItem.SubMenus
    '                            oMenus.AddEx(oCreationPackage)

    '                        Catch ex As Exception
    '                            'MessageBox.Show(ex.Message)
    '                        End Try
    '                    End If
    '                Else
    '                    Try
    '                        oMenuItem = rsboApp.Menus.Item("1280")
    '                        If oMenuItem.SubMenus.Exists("Agregar") Then
    '                            rsboApp.Menus.RemoveEx("Agregar")

    '                        End If
    '                        If oMenuItem.SubMenus.Exists("Eliminar") Then
    '                            rsboApp.Menus.RemoveEx("Eliminar")
    '                        End If
    '                    Catch ex As Exception
    '                        rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                    End Try
    '                End If

    '            End If
    '        End If
    '        ''Else
    '        ''    Try
    '        ''        rsboApp.Menus.RemoveEx("Agregar")
    '        ''        rsboApp.Menus.RemoveEx("Eliminar")
    '        ''    Catch ex As Exception
    '        ''        rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '        ''    End Try
    '    End If
    'End Sub

    'Private Sub rSboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.FormDataEvent
    '    If BusinessObjectInfo.FormTypeEx = "frmGenerarRPT" Then
    '        Select Case BusinessObjectInfo.EventType
    '            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
    '                If BusinessObjectInfo.BeforeAction Then
    '                    Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(BusinessObjectInfo.FormUID)
    '                    Dim mEdit As SAPbouiCOM.EditText = mForm.Items.Item("code").Specific
    '                    If Integer.Parse(mEdit.Value) > 1 Then
    '                        'rsboApp.SetStatusBarMessage(NombreAddon + " - Ya existe una configuración, por favor consutarla y actualizarla de sel el caso.", BoMessageTime.bmt_Medium, True)
    '                        rsboApp.MessageBox(NombreAddon + " - Ya existe una configuración, por favor consutarla y actualizarla de sel el caso.")
    '                        BubbleEvent = False
    '                        Exit Sub
    '                    End If

    '                End If
    '        End Select
    '    End If

    'End Sub
    Public Function GenerarReporte(FechaInicial As Date, FechaFinal As Date)

        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        'BD_User = ConsultaParametro("eDoc", "PARAMETROS", "CONFIGURACION", "BD_User")
        BD_User = Functions.VariablesGlobales._gUsuarioDB
        If BD_User = "" Then
            Utilitario.Util_Log.Escribir_Log("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador", "ManejoDeDocumentos")
            rsboApp.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return Nothing
        End If


        ' BD_Pass = ConsultaParametro("eDoc", "PARAMETROS", "CONFIGURACION", "BD_Pass")
        BD_Pass = Functions.VariablesGlobales._gPasswordDB
        If BD_Pass = "" Then
            Utilitario.Util_Log.Escribir_Log("GS - No existe configuracion del Clave Base De Datos, BD_User. Contacte a su Administrador", "frmGenerarRPT")
            ' rsboApp.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return Nothing
        End If

        Dim cbxTipoRPT As SAPbouiCOM.ComboBox = oForm.Items.Item("cmbRPT").Specific
        Dim sTipoRPT As String = cbxTipoRPT.Value.Trim()
        'Dim rptFilePath As String = System.Windows.Forms.Application.StartupPath & "\ReporteLLamadaServicio.rpt"
        Dim rptFilePath As String = ""
        'If sTipoRPT = "1" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Informe facturas de compras.rpt"
        'ElseIf sTipoRPT = "2" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Informe para Formulario SRI 103.rpt"
        'ElseIf sTipoRPT = "3" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Listado de Ventas.rpt"
        'ElseIf sTipoRPT = "4" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Retenciones de clientes.rpt"
        'End If

        rptFilePath = CrearRPTLocalTMP(sTipoRPT)

        Utilitario.Util_Log.Escribir_Log("RUTA del RPT : " & rptFilePath, "frmGenerarRPT")

        rsboApp.SetStatusBarMessage("Generando el Reporte...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        Try
            Dim crReport As ReportDocument = Nothing
            'Dim crReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Try

                crReport = New ReportDocument
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al instanciar clase ReportDocument : " + ex.Message.ToString(), "frmGenerarRPT")
                Return Nothing
            End Try
            If (File.Exists(rptFilePath)) Then
                Utilitario.Util_Log.Escribir_Log("Ruta si Existe", "frmGenerarRPT")
                crReport.Load(rptFilePath)
                Utilitario.Util_Log.Escribir_Log("Cargando variables RPT", "frmGenerarRPT")
                ' set parameters for your report
                ' crReport.SetParameterValue("DocKey@", DocEntry) ' DocEntry
                'Functions.VariablesGlobales._gUsuarioDB
                'rsboApp.MessageBox(Format(CDate(FechaInicial), "MM/dd/yyyy").ToString + " " + Format(CDate(FechaFinal), "MM/dd/yyyy").ToString)

                crReport.SetParameterValue("F1", FechaInicial)
                crReport.SetParameterValue("F2", FechaFinal)

                Utilitario.Util_Log.Escribir_Log("FechaInicial: " + FechaInicial.ToString, "frmGenerarRPT")
                Utilitario.Util_Log.Escribir_Log("FechaFinal: " + FechaFinal.ToString, "frmGenerarRPT")
                'Dim _ipServidor = functions.VariablesGlobales._gIpServidorDB
                Dim _ipServidor = ""
                If String.IsNullOrEmpty(Functions.VariablesGlobales._ipServer) Then
                    _ipServidor = rCompany.Server
                Else
                    _ipServidor = Functions.VariablesGlobales._ipServer
                End If


                crReport.PrintOptions.PrinterName = ""
                'oFuncionesAddon.GuardaLOG(objecType, DocEntry, "Seteando Parametros Base de Datos...", functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
                Utilitario.Util_Log.Escribir_Log("ip servidor: " + _ipServidor.ToString, "frmGenerarRPT")
                Utilitario.Util_Log.Escribir_Log("base de datos: " + rCompany.CompanyDB.ToString, "frmGenerarRPT")
                Utilitario.Util_Log.Escribir_Log("usuario BD: " + BD_User, "frmGenerarRPT")
                Utilitario.Util_Log.Escribir_Log("pass BD: " + BD_Pass, "frmGenerarRPT")
                crReport.DataSourceConnections(0).SetConnection(_ipServidor, rCompany.CompanyDB, Functions.VariablesGlobales._gUsuarioDB, Functions.VariablesGlobales._gPasswordDB)
                crReport.DataSourceConnections(0).SetLogon(BD_User, BD_Pass)

                'crReport.SetDatabaseLogon(functions.VariablesGlobales._gUsuarioDB, functions.VariablesGlobales._gPasswordDB)

                Dim IOST As IO.Stream = Nothing
                Try
                    IOST = crReport.ExportToStream(ExportFormatType.PortableDocFormat)
                    crReport.Close() ' libero la memoria del documento crystal
                    If IOST.Length > 0 Then

                        Dim b(IOST.Length) As Byte

                        IOST.Read(b, 0, CInt(IOST.Length))


                        Dim nombreTemporal As String = Format(CDate(FechaInicial), "MM-dd-yyyy").ToString + " a " + Format(CDate(FechaFinal), "MM-dd-yyyy").ToString
                        Dim temporal As String = Path.GetTempPath() & "\" & nombreTemporal + ".pdf"
                        Utilitario.Util_Log.Escribir_Log("ruta temporal: " + temporal.ToString, "frmGenerarRPT")
                        ''If File.Exists(temporal) Then
                        ''    File.Delete(temporal)
                        ''End If

                        'Dim ms As MemoryStream = New MemoryStream(b)
                        File.WriteAllBytes(temporal, b)


                        rsboApp.SetStatusBarMessage("Abriendo Documento " & nombreTemporal & " Espere unos segundos..", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        Process.Start(temporal)

                        Return True

                    End If
                Catch ex As Exception
                    rsboApp.SetStatusBarMessage("Excepcion al exportar el PDF a Stream " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Utilitario.Util_Log.Escribir_Log("Excepcion al exportar el PDF a Stream:" + ex.Message, "frmGenerarRPT")
                    Return Nothing
                End Try
                ' oFuncionesAddon.GuardaLOG(objecType, DocEntry, "PDF creado: " + TMPPDF, functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
                '.PrintToPrinter(Copies, True, StartPage, EndPage)

            Else
                'oFuncionesAddon.GuardaLOG(objecType, DocEntry, "No existe RPT en ruta: " + rptFilePath.ToString(), functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Archivo RPT de Crystal", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
                Utilitario.Util_Log.Escribir_Log("No se encontro el Archivo RPT de Crystal : ", "frmGenerarRPT")
                Return False
            End If



        Catch ex As Exception
            'oFuncionesAddon.GuardaLOG(objecType, DocEntry, "Cath al crear PDF: " + ex.Message.ToString(), functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Solsap ,Error al generar RPT: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If

            Utilitario.Util_Log.Escribir_Log("Solsap,Error al Export el PDF desde el ERP : " + ex.Message.ToString(), "frmGenerarRPT")
            Return False

        End Try
    End Function

    Public Function GenerarReporteHANA(FechaInicial As Date, FechaFinal As Date)

        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        'BD_User = ConsultaParametro("eDoc", "PARAMETROS", "CONFIGURACION", "BD_User")
        BD_User = Functions.VariablesGlobales._gUsuarioDB
        If BD_User = "" Then
            Utilitario.Util_Log.Escribir_Log("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador", "ManejoDeDocumentos")
            rsboApp.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return Nothing
        End If


        ' BD_Pass = ConsultaParametro("eDoc", "PARAMETROS", "CONFIGURACION", "BD_Pass")
        BD_Pass = Functions.VariablesGlobales._gPasswordDB
        If BD_Pass = "" Then
            Utilitario.Util_Log.Escribir_Log("GS - No existe configuracion del Clave Base De Datos, BD_User. Contacte a su Administrador", "frmGenerarRPT")
            ' rsboApp.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return Nothing
        End If

        Dim cbxTipoRPT As SAPbouiCOM.ComboBox = oForm.Items.Item("cmbRPT").Specific
        Dim sTipoRPT As String = cbxTipoRPT.Value.Trim()
        'Dim rptFilePath As String = System.Windows.Forms.Application.StartupPath & "\ReporteLLamadaServicio.rpt"
        Dim rptFilePath As String = ""
        'If sTipoRPT = "1" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Informe facturas de compras.rpt"
        'ElseIf sTipoRPT = "2" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Informe para Formulario SRI 103.rpt"
        'ElseIf sTipoRPT = "3" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Listado de Ventas.rpt"
        'ElseIf sTipoRPT = "4" Then
        '    rptFilePath = functions.VariablesGlobales._RutaRPT & "\Retenciones de clientes.rpt"
        'End If

        rptFilePath = CrearRPTLocalTMP(sTipoRPT)

        Utilitario.Util_Log.Escribir_Log("RUTA del RPT : " & rptFilePath, "frmGenerarRPT")

        rsboApp.SetStatusBarMessage("Generando el Reporte...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        Try
            Dim crReport As ReportDocument = Nothing
            'Dim crReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Try

                crReport = New ReportDocument
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Error al instanciar clase ReportDocument : " + ex.Message.ToString(), "frmGenerarRPT")
                Return Nothing
            End Try
            If (File.Exists(rptFilePath)) Then
                Utilitario.Util_Log.Escribir_Log("Ruta si Existe", "frmGenerarRPT")
                crReport.Load(rptFilePath)
                Utilitario.Util_Log.Escribir_Log("Cargando variables RPT", "frmGenerarRPT")

                Dim _ipServidor = ""
                If String.IsNullOrEmpty(Functions.VariablesGlobales._ipServer) Then
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

                ConexionHana = String.Concat(ConexionHana, "UID=", BD_User, ";")
                ConexionHana = String.Concat(ConexionHana, "PWD=", BD_Pass, ";")


                Utilitario.Util_Log.Escribir_Log("Conexion: " + ConexionHana, "frmGenerarRPT")

                crReport.SetParameterValue("F1", FechaInicial)
                crReport.SetParameterValue("F2", FechaFinal)

                Utilitario.Util_Log.Escribir_Log("FechaInicial: " + FechaInicial.ToString, "frmGenerarRPT")
                Utilitario.Util_Log.Escribir_Log("FechaFinal: " + FechaFinal.ToString, "frmGenerarRPT")
                'Dim _ipServidor = functions.VariablesGlobales._gIpServidorDB

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
                    Return Nothing
                End Try

                Try
                    Utilitario.Util_Log.Escribir_Log("Conexion Iniciando DataSourceConnections... ", "frmGenerarRPT")
                    crReport.DataSourceConnections(0).SetLogonProperties(logonProps2)
                    If String.IsNullOrEmpty(_ipServidor) Then
                        crReport.DataSourceConnections(0).SetConnection(rCompany.Server, rCompany.CompanyDB, False)
                    Else
                        Utilitario.Util_Log.Escribir_Log("Conexion Iniciando DataSourceConnections: " + _ipServidor + " Company: " + rCompany.CompanyDB + " DB: " + BD_Pass, "frmGenerarRPT")
                        crReport.DataSourceConnections(0).SetConnection(_ipServidor, rCompany.CompanyDB, False)
                        'crReport.DataSourceConnections(0).SetConnection(Functions.VariablesGlobales._gIpServidorDB, rCompany.CompanyDB, Functions.VariablesGlobales._gPasswordDB)
                        Utilitario.Util_Log.Escribir_Log("Conexion Iniciando DataSourceConnections: " + crReport.DataSourceConnections(0).ToString(), "frmGenerarRPT")
                    End If

                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("Excepcion crReport.DataSourceConnections :" + ex.Message, "frmGenerarRPT")
                    Return Nothing
                End Try

                crReport.PrintOptions.PrinterName = ""

                Dim IOST As IO.Stream = Nothing
                Try
                    IOST = crReport.ExportToStream(ExportFormatType.PortableDocFormat)
                    crReport.Close() ' libero la memoria del documento crystal
                    If IOST.Length > 0 Then

                        Dim b(IOST.Length) As Byte

                        IOST.Read(b, 0, CInt(IOST.Length))


                        Dim nombreTemporal As String = Format(CDate(FechaInicial), "MM-dd-yyyy").ToString + " a " + Format(CDate(FechaFinal), "MM-dd-yyyy").ToString
                        Dim temporal As String = Path.GetTempPath() & "\" & nombreTemporal + ".pdf"
                        Utilitario.Util_Log.Escribir_Log("ruta temporal: " + temporal.ToString, "frmGenerarRPT")
                        ''If File.Exists(temporal) Then
                        ''    File.Delete(temporal)
                        ''End If

                        'Dim ms As MemoryStream = New MemoryStream(b)
                        File.WriteAllBytes(temporal, b)


                        rsboApp.SetStatusBarMessage("Abriendo Documento " & nombreTemporal & " Espere unos segundos..", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        Process.Start(temporal)

                        Return True

                    End If
                Catch ex As Exception
                    rsboApp.SetStatusBarMessage("Excepcion al exportar el PDF a Stream " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Utilitario.Util_Log.Escribir_Log("Excepcion al exportar el PDF a Stream:" + ex.Message, "frmGenerarRPT")
                    Return Nothing
                End Try
                ' oFuncionesAddon.GuardaLOG(objecType, DocEntry, "PDF creado: " + TMPPDF, functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
                '.PrintToPrinter(Copies, True, StartPage, EndPage)

            Else
                'oFuncionesAddon.GuardaLOG(objecType, DocEntry, "No existe RPT en ruta: " + rptFilePath.ToString(), functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
                If _tipoManejo = "A" Then
                    rsboApp.SetStatusBarMessage("No se encontro el Archivo RPT de Crystal", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
                Utilitario.Util_Log.Escribir_Log("No se encontro el Archivo RPT de Crystal : ", "frmGenerarRPT")
                Return False
            End If



        Catch ex As Exception
            'oFuncionesAddon.GuardaLOG(objecType, DocEntry, "Cath al crear PDF: " + ex.Message.ToString(), functions.FuncionesAddon.Transacciones.Creacion, functions.FuncionesAddon.TipoLog.Emision)
            If _tipoManejo = "A" Then
                rsboApp.SetStatusBarMessage("Solsap ,Error al generar RPT: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If

            Utilitario.Util_Log.Escribir_Log("Solsap,Error al Export el PDF desde el ERP : " + ex.Message.ToString(), "frmGenerarRPT")
            Return False

        End Try
    End Function



    Private Function CrearRPTLocalTMP(tipoReporte As String) As String



        Try

            Dim temporal As String = ""


            Select Case tipoReporte
                Case "1"
                    temporal = Path.GetTempPath & "\rpt1.rpt"

                    If File.Exists(temporal) Then Return temporal

                    File.WriteAllBytes(temporal, Convert.FromBase64String(Functions.VariablesGlobales._SS_RPTFacturasCompras))
                Case "2"
                    temporal = Path.GetTempPath & "\rpt2.rpt"

                    If File.Exists(temporal) Then Return temporal

                    File.WriteAllBytes(temporal, Convert.FromBase64String(Functions.VariablesGlobales._SS_RPTFormularioSRI103))
                Case "3"
                    temporal = Path.GetTempPath & "\rpt3.rpt"

                    If File.Exists(temporal) Then Return temporal

                    File.WriteAllBytes(temporal, Convert.FromBase64String(Functions.VariablesGlobales._SS_RPTListadoVentas))
                Case "4"
                    temporal = Path.GetTempPath & "\rpt4.rpt"

                    If File.Exists(temporal) Then Return temporal

                    File.WriteAllBytes(temporal, Convert.FromBase64String(Functions.VariablesGlobales._SS_RPTRetencionClientes))
                Case Else



            End Select


            If File.Exists(temporal) Then

                Return temporal


            End If




        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("No se Pudo generar el Archivo Temporal " & ex.Message, "frmGenerarRPT")

        End Try


        Return ""

    End Function
End Class

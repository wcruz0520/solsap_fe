Public Class frmMapeoCuentasCM

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private oForm As SAPbouiCOM.Form
    Dim odt As SAPbouiCOM.DataTable

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rSboApp = sboApp
    End Sub

    Public Sub CargaFormularioConfiguracionesCM()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmMapeoCuentasCM") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmMapeoCuentasCM.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmMapeoCuentasCM").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmMapeoCuentasCM")

            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            ipLogoSS.Item.Visible = True

            LlenaCombos()

            oForm.Visible = True
            oForm.Select()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Ocurrio un Error al Cargar la Pantalla frmMapeoCuentasCM: " & ex.Message.ToString, "frmMapeoCuentasCM")
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla frmMapeoCuentasCM: " + ex.Message.ToString())
        End Try
    End Sub

    Private Sub LlenaCombos()
        Try
            oForm = rsboApp.Forms.Item("frmMapeoCuentasCM")
            oForm.Freeze("true")

            Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing
            Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific

            Dim QueryCTA As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryCTA = "SELECT ""AcctCode"", ""AcctName"" FROM ""OACT"" WHERE ""Finanse"" = 'Y'" ' AND ""CfwRlvnt"" = 'Y'" 
            Else
                QueryCTA = "SELECT AcctCode, AcctName FROM OACT WITH(NOLOCK) WHERE Finanse = 'Y'" ' AND CfwRlvnt = 'Y'"
            End If

            Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet(QueryCTA)
            ValoresValidos = cbxCuenta.ValidValues

            If oRecordSet.RecordCount >= 1 Then
                While (oRecordSet.EoF = False)
                    ValoresValidos.Add(oRecordSet.Fields.Item("AcctCode").Value, oRecordSet.Fields.Item("AcctName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            Dim cbxBanco As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxBanco").Specific

            Dim QueryBCO As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryBCO = "SELECT ""BankCode"", ""BankName"" FROM ""ODSC"""
            Else
                QueryBCO = "SELECT BankCode, BankName FROM ODSC WITH(NOLOCK)"
            End If

            oRecordSet = oFuncionesB1.getRecordSet(QueryBCO)
            ValoresValidos = cbxBanco.ValidValues

            If oRecordSet.RecordCount >= 1 Then
                While (oRecordSet.EoF = False)
                    ValoresValidos.Add(oRecordSet.Fields.Item("BankCode").Value, oRecordSet.Fields.Item("BankName").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            CargaDatos()

            oForm.Freeze("false")
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error LlenaCombos: " & ex.Message.ToString, "frmMapeoCuentasCM")
            rsboApp.StatusBar.SetText("Error LlenaCombos: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub CargaDatos()
        Try
            oForm.Freeze(True)

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
                oForm.DataSources.DataTables.Add("dtConf")
            Catch ex As Exception
            End Try

            Dim Query As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Query = My.Resources.SS_CM_CONSULTAPARAMETROSBC_HANA.ToString
            Else
                Query = My.Resources.SS_CM_CONSULTAPARAMETROSBC_SQL.ToString
            End If

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                Utilitario.Util_Log.Escribir_Log("Query consulta mapeo bco-cta:" & Query, "frmMapeoCuentasCM")
                oGrid.DataTable.ExecuteQuery(Query)
                Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" & Query, "frmMapeoCuentasCM")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + Query, "frmMapeoCuentasCM")
            End Try

            oGrid.Columns.Item(0).Description = "Código Banco"
            oGrid.Columns.Item(0).TitleObject.Caption = "Código Banco"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "Nombre Banco"
            oGrid.Columns.Item(1).TitleObject.Caption = "Nombre Banco"
            oGrid.Columns.Item(1).Editable = False

            oGrid.Columns.Item(2).Description = "Cuenta Sistema"
            oGrid.Columns.Item(2).TitleObject.Caption = "Cuenta Sistema"
            oGrid.Columns.Item(2).Editable = False

            oGrid.Columns.Item(3).Description = "Nombre Cta Sistema"
            oGrid.Columns.Item(3).TitleObject.Caption = "Nombre Cta Sistema"
            oGrid.Columns.Item(3).Editable = False

            oGrid.Columns.Item(4).Description = "Cuenta Banco"
            oGrid.Columns.Item(4).TitleObject.Caption = "Cuenta Banco"
            oGrid.Columns.Item(4).Editable = True

            Query = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Query = My.Resources.SS_CM_CONSULTAPARAMETROSCONF_HANA.ToString
            Else
                Query = My.Resources.SS_CM_CONSULTAPARAMETROSCONF_SQL.ToString
            End If

            oForm.DataSources.DataTables.Item("dtConf").ExecuteQuery(Query)
            odt = oForm.DataSources.DataTables.Item("dtConf")

            Dim txtDir As SAPbouiCOM.EditText = oForm.Items.Item("txtDir").Specific
            Dim cbxMC As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxMC").Specific

            For i As Integer = 0 To odt.Rows.Count - 1
                If odt.GetValue("U_Nombre", i).ToString().Equals("RutaArchivoTxt") Then
                    txtDir.Value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ManejoCuenta") Then
                    cbxMC.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
            Next

            oForm.Freeze(False)
            rsboApp.StatusBar.SetText("Datos cargados con éxito! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error CargaDatos: " & ex.Message.ToString, "frmMapeoCuentasCM")
            rsboApp.StatusBar.SetText("Error CargaDatos: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        If pVal.FormTypeEx = "frmMapeoCuentasCM" Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If Not pVal.Before_Action Then
                        Select Case pVal.ItemUID
                            Case "obtnAdd"

                                Try
                                    Dim cbxCuenta As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxCuenta").Specific
                                    Dim cbxBanco As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxBanco").Specific
                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific

                                    If cbxBanco.Value.ToString = "" Then
                                        rsboApp.StatusBar.SetText("Seleccione un Banco! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return
                                    End If

                                    If cbxCuenta.Value.ToString = "" Then
                                        rsboApp.StatusBar.SetText("Seleccione un Cuenta! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return
                                    End If

                                    oForm.Freeze(True)

                                    If String.IsNullOrEmpty(oGrid.DataTable.GetValue("U_CodBco", oGrid.Rows.Count - 1)) Then
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_CodBco", oGrid.Rows.Count - 1, cbxBanco.Selected.Value)
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_NomBco", oGrid.Rows.Count - 1, cbxBanco.Selected.Description)
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_CtaSys", oGrid.Rows.Count - 1, cbxCuenta.Selected.Value)
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_NomCtaSys", oGrid.Rows.Count - 1, cbxCuenta.Selected.Description)
                                    Else
                                        oForm.DataSources.DataTables.Item("dtDocs").Rows.Add()
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_CodBco", oGrid.Rows.Count - 1, cbxBanco.Selected.Value)
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_NomBco", oGrid.Rows.Count - 1, cbxBanco.Selected.Description)
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_CtaSys", oGrid.Rows.Count - 1, cbxCuenta.Selected.Value)
                                        oForm.DataSources.DataTables.Item("dtDocs").SetValue("U_NomCtaSys", oGrid.Rows.Count - 1, cbxCuenta.Selected.Description)
                                    End If

                                    oForm.Freeze(False)
                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error boton agregar: " & ex.Message.ToString, "frmMapeoCuentasCM")
                                    rsboApp.StatusBar.SetText("Error boton agregar: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try

                            Case "obtnGrabar"
                                Try
                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific

                                    Dim Tabla As SAPbobsCOM.UserTable = rCompany.UserTables.Item("SS_CM_MBCOCTA")
                                    For i As Integer = 0 To oGrid.Rows.Count - 1
                                        If oGrid.DataTable.GetValue("U_CodBco", i).ToString <> "" Then
                                            Tabla.Code = CStr(i)
                                            Tabla.Name = CStr(i)
                                            Tabla.UserFields.Fields.Item("U_CodBco").Value = oGrid.DataTable.GetValue("U_CodBco", i).ToString
                                            Tabla.UserFields.Fields.Item("U_NomBco").Value = oGrid.DataTable.GetValue("U_NomBco", i).ToString
                                            Tabla.UserFields.Fields.Item("U_CtaSys").Value = oGrid.DataTable.GetValue("U_CtaSys", i).ToString
                                            Tabla.UserFields.Fields.Item("U_NomCtaSys").Value = oGrid.DataTable.GetValue("U_NomCtaSys", i).ToString
                                            Tabla.UserFields.Fields.Item("U_CtaBco").Value = oGrid.DataTable.GetValue("U_CtaBco", i).ToString

                                            If Tabla.Add() <> 0 Then
                                                Dim sErrMsg As String = ""
                                                rCompany.GetLastError(0, sErrMsg)
                                                Utilitario.Util_Log.Escribir_Log("Error registrando mapeo de banco-cuenta! " & sErrMsg, "frmMapeoCuentasCM")
                                            Else
                                                Utilitario.Util_Log.Escribir_Log("Mapeo de cuenta registrado correctamente! ", "frmMapeoCuentasCM")
                                            End If
                                        End If
                                    Next

                                    Dim oConfiguracion As Entidades.Configuracion
                                    Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                    Dim txtDir As SAPbouiCOM.EditText = oForm.Items.Item("txtDir").Specific
                                    Dim cbxMC As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxMC").Specific

                                    oConfiguracion = New Entidades.Configuracion
                                    oConfiguracion.Modulo = NombreAddonLOC
                                    oConfiguracion.Tipo = "PARAMETROS"
                                    oConfiguracion.SubTipo = "CONFIGURACION"
                                    olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RutaArchivoTxt", txtDir.Value.ToString))
                                    olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ManejoCuenta", cbxMC.Value.ToString))
                                    oConfiguracion.Detalle = olistaDetalleConfiguracion
                                    GuardaCONF(oConfiguracion)

                                    oForm.Items.Item("obtnGrabar").Visible = False
                                    oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                    Dim oB As SAPbouiCOM.Button = oForm.Items.Item("2").Specific
                                    oB.Caption = "OK"
                                    rsboApp.StatusBar.SetText("Mapeo y Parametros de configuración de Cash management guardados correctamente... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                    CargaDatos()

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al guardar parametros de banco cuenta y configuración: " & ex.Message, "frmMapeoCuentasCM")
                                    rsboApp.StatusBar.SetText("Error al guardar parametros de banco cuenta y configuración: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                        End Select
                    End If
            End Select
        End If
    End Sub


    Public Sub GuardaCONF(ByVal oConfiguracion As Entidades.Configuracion)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            Dim query As String
            Dim CodeExist As String = "0"
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@SS_CONF"" Where ""U_Modulo"" = '" + oConfiguracion.Modulo + "' AND ""U_Tipo"" = '" + oConfiguracion.Tipo + "' AND ""U_Subtipo"" = '" + oConfiguracion.SubTipo + "'"
            Else
                query = "Select DocEntry From [@SS_CONF] Where U_Modulo = '" + oConfiguracion.Modulo + "' AND U_Tipo = '" + oConfiguracion.Tipo + "' AND U_Subtipo = '" + oConfiguracion.SubTipo + "'"
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            'mRst = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CodeExist = "0" Then ' SI EXISTE, RECORRO LINEA Y ACTUALIZO

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONFLOC")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oChildren = oGeneralData.Child("SS_CONFD")

                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    Dim bandera As Boolean = False
                    For Each it As SAPbobsCOM.GeneralData In oChildren
                        Dim NombreCampo As String = it.GetProperty("U_Nombre").ToString
                        If oItem.Nombre = NombreCampo Then
                            bandera = True
                            it.SetProperty("U_Valor", oItem.Valor)
                        End If
                    Next

                    If bandera = False Then
                        oChild = oChildren.Add
                        oChild.SetProperty("U_Nombre", oItem.Nombre)
                        oChild.SetProperty("U_Valor", oItem.Valor)
                    End If
                Next
                oGeneralService.Update(oGeneralData)
            Else

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONFLOC")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("SS_CONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error guardando parametro de configuración CM: " & ex.Message, "frmMapeoCuentasCM")
            rsboApp.StatusBar.SetText("Error guardando parametro de configuración CM: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
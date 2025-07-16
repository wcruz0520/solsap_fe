Imports System.IO
Imports System.Text
Imports System.Configuration
Public Class frmCashManagemet
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Dim odt As SAPbouiCOM.DataTable
    Dim RutaGeneral As String = ""
    Dim RutaComplementaria As String = ""
    Dim FechaUDO As Date
    Dim sQuery As String = ""
    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioCashManagement(Optional ByVal DocEntryPM As String = "", Optional ByVal Cuenta As SAPbouiCOM.ComboBox = Nothing)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmCashManagement") Then Exit Sub

        strPath = System.Windows.Forms.Application.StartupPath & "\frmCashManagement.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmCashManagement").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmCashManagement")

            oForm.Freeze(True)

            InicioControles()

            oForm.Freeze(False)
            oForm.Visible = True
            oForm.Select()
        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla frmPagosMasivos: " + ex.Message.ToString())
        End Try
    End Sub

    Private Sub InicioControles()
        Try

            Dim finicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
            Dim ffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific

            finicial.Value = DateTime.Now.ToString("yyyyMMdd")
            ffinal.Value = DateTime.Now.ToString("yyyyMMdd")

            sQuery = ""

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "SELECT B.""GLAccount"", A.""AcctName"", B.""BankCode"" FROM ""OACT"" A INNER JOIN ""DSC1"" B ON A.""AcctCode"" = B.""GLAccount"""
            Else
                sQuery = "SELECT B.GLAccount, A.AcctName, B.BankCode FROM OACT A WITH(NOLOCK) INNER JOIN DSC1 B WITH(NOLOCK) ON A.AcctCode = B.GLAccount"
            End If

            Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing
            Dim cbxBan As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxBan").Specific
            Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet(sQuery)
            ValoresValidos = cbxBan.ValidValues

            If Functions.VariablesGlobales._ManejoCuenta = "Mapeada" Then
                If oRecordSet.RecordCount >= 1 Then
                    While (oRecordSet.EoF = False)
                        ValoresValidos.Add(oRecordSet.Fields.Item("U_CtaSys").Value, oRecordSet.Fields.Item("U_CodBco").Value.ToString & ":" & oRecordSet.Fields.Item("U_NomCtaSys").Value)
                        oRecordSet.MoveNext()
                    End While
                End If
            Else
                If oRecordSet.RecordCount >= 1 Then
                    While Not oRecordSet.EoF '(oRecordSet.EoF = False)
                        ValoresValidos.Add(oRecordSet.Fields.Item("GLAccount").Value, oRecordSet.Fields.Item("BankCode").Value.ToString & ":" & oRecordSet.Fields.Item("AcctName").Value)
                        oRecordSet.MoveNext()
                    End While
                End If
            End If

            Dim ruta As String = "", fecha = "" ', nombreArchivo As String = ""
            ruta = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Archivo Bancario\" 'Functions.VariablesGlobales._RutaArchivoTxt
            RutaGeneral = ruta
            fecha = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")
            FechaUDO = Date.Now

            If cbxBan.Value <> "" Then
                Dim separadorCuentaBanco As String() = cbxBan.Selected.Description.Split(":")

                Dim lblRA As SAPbouiCOM.StaticText = oForm.Items.Item("lblRA").Specific
                lblRA.Caption = ruta & separadorCuentaBanco(1).Replace("*", "").Replace("?", "") & "_" & fecha.Replace(":", "") & ".txt"
            End If

            Dim ipLogoSS As SAPbouiCOM.PictureBox = oForm.Items.Item("ipLogoSS").Specific
            ipLogoSS.Picture = Application.StartupPath & "\LogoSS.png"
            ipLogoSS.Item.Visible = True


            Dim cbxEstado As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxEstado").Specific
            cbxEstado.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            'InicializarValores(DocEntryPM)

            InicializarValores()

            BloqueaControles(True, False, False)

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oForm.Items.Item("1").Visible = False

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Error iniciando controles: " + ex.Message.ToString())
        End Try
    End Sub


    Private Sub InicializarValores(Optional ByVal DocEntry As String = "")

        CargarDatos(DocEntry)

    End Sub

    Private Sub CargarDatos(Optional ByVal DocEntry As String = "")
        Try
            oForm.Freeze(True)

            Dim finicial As SAPbouiCOM.EditText = oForm.Items.Item("finicial").Specific
            Dim ffinal As SAPbouiCOM.EditText = oForm.Items.Item("ffinal").Specific
            Dim cbxbanco As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxBan").Specific
            Dim cuenta() As String = cbxbanco.Value.Split(":")
            Dim filtro As String = ""

            Dim sQuery As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQuery = "CALL " & rCompany.CompanyDB & ".SS_CM_CONSULTAPAGOEF (" 'My.Resources.SS_CM_CONSULTAPAGOEF_HANA & filtro
                sQuery += "'" & finicial.Value.ToString & "'"
                sQuery += ", '" & ffinal.Value.ToString & "'"
                sQuery += ", '" & cuenta(0) & "'"
                sQuery += ", '" & DocEntry.ToString & "')"
            Else
                sQuery = "EXEC SS_CM_CONSULTAPAGOEF " 'My.Resources.SS_CM_CONSULTAPAGOEF_SQL & filtro
                sQuery += "'" & finicial.Value.ToString & "'"
                sQuery += ", '" & ffinal.Value.ToString & "'"
                sQuery += ", '" & cuenta(0) & "'"
                sQuery += ", '" & DocEntry.ToString & "'"
            End If

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                Utilitario.Util_Log.Escribir_Log("Query a ejecutar:" + sQuery, "frmCashManagement")
                oGrid.DataTable.ExecuteQuery(sQuery)
                Utilitario.Util_Log.Escribir_Log("Query que se ejecuto:" + sQuery, "frmCashManagement")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQuery, "frmCashManagement")
            End Try

            FormatoGrid()

            Dim lblTF As SAPbouiCOM.StaticText = oForm.Items.Item("lblTF").Specific
            lblTF.Caption = "0"

            Dim lblNF As SAPbouiCOM.StaticText = oForm.Items.Item("lblNF").Specific
            lblNF.Caption = "0"

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Error CargarDatos: " + ex.Message.ToString())
            Utilitario.Util_Log.Escribir_Log("Error CargarDatos:" & ex.Message.ToString, "frmCashManagement")
        End Try
    End Sub

    Private Sub GeneraArchivo()
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            Dim cbxBan As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxBan").Specific
            Dim separador As String()
            Dim lblTF As SAPbouiCOM.StaticText = oForm.Items.Item("lblTF").Specific


            If CInt(lblTF.Caption) <= 0 Then
                rsboApp.StatusBar.SetText("Seleccione un registro para generar archivo!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return
            End If

            If cbxBan.Value.ToString = "" Then
                rsboApp.StatusBar.SetText("Seleccione un banco... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return
            End If

            Dim banco As String() = cbxBan.Selected.Description.Split(":")

            Dim errores As List(Of String) = ValidarDatosConErrores(CInt(banco(0).ToString))

            If errores.Count > 0 Then
                For Each ex As String In errores
                    Utilitario.Util_Log.Escribir_Log("Revisar campo: " & ex, "frmCashManagement")
                    rsboApp.StatusBar.SetText("Revisar campo: " & ex, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Next
                Exit Sub
            Else
                Dim lblRA As SAPbouiCOM.StaticText = oForm.Items.Item("lblRA").Specific

                rsboApp.StatusBar.SetText("Generando archivo en la siguiente ruta: " & lblRA.Caption, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim separadorCuentaBanco As String()
                Dim fecha As String = ""
                separadorCuentaBanco = cbxBan.Selected.Description.Split(":")
                fecha = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")
                Dim nombrearchivo = separadorCuentaBanco(1).Replace("*", "").Replace("?", "") & "_" & fecha.Replace(":", "") & ".txt"

                Dim ruta As String = Path.Combine(RutaGeneral, nombrearchivo)

                Dim strStreamW As Stream = Nothing
                Dim strStreamWriter As StreamWriter = Nothing

                If Not File.Exists(ruta) Then
                    strStreamW = File.Create(ruta) ' lblRA.Caption) ' lo creamos
                    strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default) '
                    strStreamWriter.Close() ' cerramos
                Else
                    File.Delete(ruta) ' lo eliminamos
                    strStreamW = File.Create(ruta) ' lo creamos
                    strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default) '
                    strStreamWriter.Close() ' cerramos
                End If

                Dim sTexto As New StringBuilder

                Dim CampoVacio As String = ""


                For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    Dim isChecked As String = oGrid.DataTable.GetValue("Check", i)

                    If isChecked = "Y" Then

                        Dim DocTotalRS As Double = Convert.ToDouble(oGrid.DataTable.GetValue("DocTotal", i))
                        Dim DocTotalConv As String = DocTotalRS.ToString("F2")
                        Dim NombreProv As String = oGrid.DataTable.GetValue("CardName", i).ToString

                        Select Case CInt(banco(0).ToString)
                            Case 10
                                sTexto.AppendLine(oGrid.DataTable.GetValue("Metodo", i).ToString & vbTab & 'Codigo de orientacion 1
                                                  oGrid.DataTable.GetValue("CardCode", i).ToString & vbTab & 'Contrapartida 1
                                                  oGrid.DataTable.GetValue("Moneda", i).ToString & vbTab & 'Moneda 1
                                                  Right("0000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 13) & vbTab & 'Valor 1
                                                 oGrid.DataTable.GetValue("FormaPago", i).ToString & vbTab & 'Forma de pago 1
                                                 oGrid.DataTable.GetValue("TipoCuenta", i).ToString & vbTab & 'Tipo de cuenta 1
                                                 oGrid.DataTable.GetValue("NumeroCuenta", i).ToString & vbTab & 'numero de cuenta 1
                                                 oGrid.DataTable.GetValue("Referencia", i).ToString & vbTab & 'Referencia 0
                                                 oGrid.DataTable.GetValue("TipoCliente", i).ToString & vbTab & 'Tipo ID Cliente 1
                                                 oGrid.DataTable.GetValue("NumeroID", i).ToString & vbTab & 'Numero ID Cliente 1
                                                 NombreProv.Replace("ñ", "n").Replace("Ñ", "N") & vbTab & 'Nombre del cliente 1
                                                 oGrid.DataTable.GetValue("CodBanco", i).ToString) 'Codigo de banco 0

                            Case 36
                                sTexto.AppendLine(oGrid.DataTable.GetValue("Metodo", i).ToString & vbTab & 'Codigo de orientacion 1
                                              Right("00000000000" & oGrid.DataTable.GetValue("CodCuentaEmpresa", i).ToString, 11) & vbTab & 'Cuenta empresa 1
                                              oGrid.DataTable.GetValue("DocNum", i).ToString & vbTab & 'Secuencial pago 1
                                              vbTab & 'Comprobante de pago 0
                                              oGrid.DataTable.GetValue("CardCode", i).ToString & vbTab & 'Contrapartida 1
                                              oGrid.DataTable.GetValue("Moneda", i).ToString & vbTab & 'Moneda 1
                                              Right("0000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 13) & vbTab & 'Valor 1
                                              oGrid.DataTable.GetValue("FormaPago", i).ToString & vbTab & 'Forma de pago 1
                                              Right("0000" & oGrid.DataTable.GetValue("CodBancoEmpresa", i).ToString, 4) & vbTab & 'Codigo de institucion financiera 1
                                              IIf(oGrid.DataTable.GetValue("FormaPago", i).ToString = "CTA", oGrid.DataTable.GetValue("TipoCuenta", i).ToString & vbTab, vbTab) & 'Tipo de cuenta 1
                                              IIf(oGrid.DataTable.GetValue("FormaPago", i).ToString = "CTA", oGrid.DataTable.GetValue("NumeroCuenta", i).ToString & vbTab, vbTab) & 'Numero de cuenta 1
                                              oGrid.DataTable.GetValue("TipoCliente", i).ToString & vbTab & 'Tipo ID Cliente 1
                                              oGrid.DataTable.GetValue("NumeroID", i).ToString & vbTab & 'Numero ID Cliente 1
                                              Right(NombreProv.Replace("ñ", "n").Replace("Ñ", "N"), 60) & vbTab & 'Nombre del cliente 1
                                              vbTab & 'Direccion 0
                                              vbTab & 'Ciudad 0
                                              vbTab & 'Telefono 0
                                              vbTab & 'Localidad de pago 0
                                              oGrid.DataTable.GetValue("FactReferencia", i).ToString & vbTab & 'Referencia 1
                                              oGrid.DataTable.GetValue("Referencia", i).ToString & IIf(oGrid.DataTable.GetValue("Email", i).ToString <> "", "| " & oGrid.DataTable.GetValue("Email", i).ToString, "")) 'Ref adicional 0

                            Case 17
                                sTexto.AppendLine(
                                               oGrid.DataTable.GetValue("Metodo", i).ToString & vbTab & 'Codigo de orientacion 1
                                               Right("0000000000" & oGrid.DataTable.GetValue("CodCuentaEmpresa", i).ToString, 10) & vbTab & 'Cuenta empresa 1
                                               Right("0000000" & oGrid.DataTable.GetValue("DocNum", i).ToString, 7) & vbTab & 'Secuencial pago 1
                                               vbTab & 'Comprobante de pago 0
                                               oGrid.DataTable.GetValue("CardCode", i).ToString & vbTab & 'Contrapartida 1
                                               oGrid.DataTable.GetValue("Moneda", i).ToString & vbTab & 'Moneda 1
                                               Right("0000000000000" & DocTotalConv.ToString.Replace(",", "").Replace(".", ""), 13) & vbTab & 'Valor 1
                                               oGrid.DataTable.GetValue("FormaPago", i).ToString & vbTab & 'Forma de pago 1
                                               Right("0000" & oGrid.DataTable.GetValue("CodBancoEmpresa", i).ToString, 4) & vbTab & 'Codigo de institucion financiera 1
                                               IIf(oGrid.DataTable.GetValue("FormaPago", i).ToString = "CTA", oGrid.DataTable.GetValue("TipoCuenta", i).ToString & vbTab, vbTab) & 'Tipo de cuenta 1
                                               IIf(oGrid.DataTable.GetValue("FormaPago", i).ToString = "CTA", Right("00000000000" & (oGrid.DataTable.GetValue("NumeroCuenta", i).ToString), 11) & vbTab, vbTab) & 'Numero de cuenta 1
                                               oGrid.DataTable.GetValue("TipoCliente", i).ToString & vbTab & 'Tipo ID Cliente 1
                                               oGrid.DataTable.GetValue("NumeroID", i).ToString & vbTab & 'Numero ID Cliente 1
                                               Right(NombreProv.Replace("ñ", "n").Replace("Ñ", "N"), 40) & vbTab & 'Nombre del cliente 1
                                               vbTab & 'Direccion 0
                                               vbTab & 'Ciudad 0
                                               vbTab & 'Telefono 0
                                               vbTab & 'Localidad de pago 0
                                               oGrid.DataTable.GetValue("FactReferencia", i).ToString & vbTab & 'Referencia 1
                                               Right((oGrid.DataTable.GetValue("Referencia", i).ToString & IIf(oGrid.DataTable.GetValue("Email", i).ToString <> "", "| " & oGrid.DataTable.GetValue("Email", i).ToString, "")), 200)) 'Ref adicional 0

                            Case Else
                                rsboApp.StatusBar.SetText("Archivo de banco no diseñado! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Return
                        End Select
                    End If
                Next

                Try
                    Dim oTextWriter As TextWriter = New StreamWriter(ruta, True) 'New StreamWriter(lblRA.Caption, True)
                    oTextWriter.WriteLine(sTexto.ToString)
                    oTextWriter.Flush()
                    oTextWriter.Close()
                    oTextWriter = Nothing
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("Error escribiendo archivo: " & ex.Message.ToString, "frmCashManagement")
                    rsboApp.StatusBar.SetText("Error escribiendo archivo: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return
                End Try

                rsboApp.StatusBar.SetText("Archivo generado! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                'Agregar lógica para guardar informacion en UDO
                Dim DocEntryUDOCM As Integer
                If GuardaUDOCM(DocEntryUDOCM) Then
                    For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                        Dim isChecked As String = oGrid.DataTable.GetValue("Check", i)

                        If isChecked = "Y" Then
                            Dim result As Integer
                            Dim ErrCode As Long
                            Dim ErrMsg As String
                            Dim PagoEfectuado As SAPbobsCOM.Payments = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                            If PagoEfectuado.GetByKey(CInt(oGrid.DataTable.GetValue("DocEntry", i).ToString)) Then
                                PagoEfectuado.UserFields.Fields.Item("U_UDO_CM").Value = DocEntryUDOCM
                                result = PagoEfectuado.Update()
                                If result <> 0 Then
                                    rCompany.GetLastError(ErrCode, ErrMsg)
                                    rsboApp.StatusBar.SetText("Error actualizando UDF del pago efectuado " & oGrid.DataTable.GetValue("DocEntry", i).ToString & " - " & ErrCode.ToString & " - " & ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Utilitario.Util_Log.Escribir_Log("Error actualizando UDF del pago efectuado " & oGrid.DataTable.GetValue("DocEntry", i).ToString & " - " & ErrCode.ToString & " - " & ErrMsg, "frmCashManagement")
                                Else
                                    rsboApp.StatusBar.SetText("Actualización del UDF del pago efectuado con éxito! " & oGrid.DataTable.GetValue("DocEntry", i).ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Utilitario.Util_Log.Escribir_Log("Actualización del UDF del pago efectuado con éxito! " & oGrid.DataTable.GetValue("DocEntry", i).ToString, "frmCashManagement")
                                End If

                            End If
                        End If
                    Next
                    CargarDatos()
                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al GeneraArchivo:" & ex.Message.ToString, "frmCashManagement")
            rsboApp.StatusBar.SetText("Error al GeneraArchivo: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return
        End Try
    End Sub

    Private Function ValidarDatosConErrores(ByVal Banco As Integer) As List(Of String)
        Try
            Dim errores As New List(Of String)()
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific

            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                Dim isChecked As String = oGrid.DataTable.GetValue("Check", i)

                If isChecked = "Y" Then

                    Select Case Banco
                        Case 10

                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("Metodo", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Código Orientacion (Metodo)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CardCode", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Contrapartida 'Cod cliente, # medidor, # telefono, # Contrato' (CardCode)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("Moneda", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Moneda (Moneda)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("DocTotal", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Valor (DocTotal)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("FormaPago", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Forma de Pago (FormaPago)")
                            Else
                                If oGrid.DataTable.GetValue("FormaPago", i).ToString = "CTA" Then
                                    If String.IsNullOrEmpty(oGrid.DataTable.GetValue("TipoCuenta", i).ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Tipo de cuenta (TipoCuenta)")
                                    End If
                                    If String.IsNullOrEmpty(oGrid.DataTable.GetValue("NumeroCuenta", i).ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Numero de cuenta (NumeroCuenta)")
                                    End If
                                End If
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("TipoCliente", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Tipo ID Cliente (TipoCliente)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("NumeroID", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Numero ID Cliente (NumeroID)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CardName", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Nombre del cliente (CardName)")
                            End If

                        Case 36

                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("Metodo", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Código Orientacion (Metodo)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CodCuentaEmpresa", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Cuenta Empresa (CodCuentaEmpresa)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("DocNum", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Secuencial pago (DocNum)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("Moneda", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Moneda (Moneda)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("DocTotal", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Valor (DocTotal)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("FormaPago", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Forma de Pago (FormaPago)")
                            Else
                                If oGrid.DataTable.GetValue("FormaPago", i).ToString = "CTA" Then
                                    If String.IsNullOrEmpty(oGrid.DataTable.GetValue("TipoCuenta", i).ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Tipo de Cuenta (TipoCuenta)")
                                    End If
                                    If String.IsNullOrEmpty(oGrid.DataTable.GetValue("NumeroCuenta", i).ToString) Then
                                        errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Numero de cuenta (NumeroCuenta)")
                                    End If
                                End If
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CodBancoEmpresa", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Codigo de institucion financiera (CodBancoEmpresa)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("TipoCliente", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Tipo ID Cliente (TipoCliente)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("NumeroID", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Numero ID Cliente (NumeroID)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CardName", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Nombre del cliente (CardName)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("FactReferencia", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Referencia 'Numero de factura' (FactReferencia)")
                            End If

                        Case 17

                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("Metodo", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Código Orientacion (Metodo)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CodCuentaEmpresa", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Cuenta Empresa (CodCuentaEmpresa)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("DocNum", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Secuencial pago (DocNum)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CardCode", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Código 'Cod cliente, # medidor, # telefono, # Contrato' (CardCode)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("Moneda", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Moneda (Moneda)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("DocTotal", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Valor (DocTotal)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("FormaPago", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Forma de Pago (FormaPago)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CodBancoEmpresa", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Codigo de institucion financiera (CodBancoEmpresa)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("TipoCuenta", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Tipo de Cuenta (TipoCuenta)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("NumeroCuenta", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Numero de cuenta (NumeroCuenta)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("TipoCliente", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Tipo ID Cliente (TipoCliente)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("NumeroID", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Numero ID Cliente (NumeroID)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("CardName", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Nombre del cliente (CardName)")
                            End If
                            If String.IsNullOrEmpty(oGrid.DataTable.GetValue("FactReferencia", i).ToString) Then
                                errores.Add("Linea: " & i.ToString & " - Pago:" & oGrid.DataTable.GetValue("DocNum", i).ToString & " - Referencia 'Numero de factura' (FactReferencia)")
                            End If

                    End Select

                End If

            Next

            Return errores
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error ValidandoDatosAGenerar:" & ex.Message.ToString, "frmCashManagement")
            rsboApp.StatusBar.SetText("Error ValidandoDatosAGenerar: " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        If pVal.FormTypeEx = "frmCashManagement" Then
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        If Not pVal.BeforeAction Then
                            Select Case pVal.ItemUID

                                Case "cbxBan"
                                    Dim lblRA As SAPbouiCOM.StaticText = oForm.Items.Item("lblRA").Specific
                                    Dim cbxBan As SAPbouiCOM.ComboBox = oForm.Items.Item("cbxBan").Specific
                                    Dim separadorCuentaBanco As String()
                                    Dim fecha As String = ""

                                    If cbxBan.Value <> "" Then
                                        separadorCuentaBanco = cbxBan.Selected.Description.Split(":")
                                        fecha = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")
                                        FechaUDO = Date.Now
                                        lblRA.Caption = RutaGeneral & separadorCuentaBanco(1).Replace("*", "").Replace("?", "") & "_" & fecha.Replace(":", "") & ".txt"
                                    End If


                                Case "oGrid"
                                    Try
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                            If pVal.ColUID = "Check" Then
                                                If pVal.Row >= 0 Then
                                                    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
                                                    Dim isChecked As String = oGrid.DataTable.GetValue("Check", pVal.Row)

                                                    Dim lblNF As SAPbouiCOM.StaticText = oForm.Items.Item("lblNF").Specific
                                                    Dim lblTF As SAPbouiCOM.StaticText = oForm.Items.Item("lblTF").Specific
                                                    Dim valorNF As Integer = CInt(lblNF.Caption)
                                                    Dim valorTF As Double = CDbl(lblTF.Caption)

                                                    If isChecked = "Y" Then
                                                        valorNF += 1
                                                        valorTF += CDbl(oGrid.DataTable.GetValue("DocTotal", pVal.Row))
                                                    Else
                                                        valorNF -= 1
                                                        valorTF -= CDbl(oGrid.DataTable.GetValue("DocTotal", pVal.Row))
                                                    End If

                                                    lblNF.Caption = CStr(valorNF)
                                                    lblTF.Caption = CStr(valorTF)

                                                End If
                                            End If

                                        End If
                                    Catch ex As Exception
                                    End Try
                                Case "cbxEstado"
                                    BloqueaControles(False, True, True)

                                Case "lblRA"
                                    Try
                                        Dim lblRAG As SAPbouiCOM.StaticText = oForm.Items.Item("lblRA").Specific
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
                            End Select

                        Else
                            Select Case pVal.ItemUID
                                Case "cbxEstado"
                                    BloqueaControles(False, True, True)
                            End Select
                        End If
                    Catch ex As Exception

                    End Try

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If Not pVal.BeforeAction Then
                        Select Case pVal.ItemUID
                            Case "btnBuscar"
                                CargarDatos()

                            Case "btnGA"
                                GeneraArchivo()

                            Case "1"

                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then BuscaRegistros()

                        End Select

                    End If
            End Select
        End If
    End Sub

    Private Sub FormatoGrid()
        Try
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(0).Description = ""
            oGrid.Columns.Item(0).TitleObject.Caption = ""

            oGrid.Columns.Item(1).Description = "Metodo"
            oGrid.Columns.Item(1).TitleObject.Caption = "Metodo"
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(1).Visible = False

            oGrid.Columns.Item(2).Description = "DocEntry"
            oGrid.Columns.Item(2).TitleObject.Caption = "DocEntry"
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(2).Visible = False

            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = "SSMTPAGOS" '46

            oGrid.Columns.Item(3).Description = "# Solicitud" '"# Pago"
            oGrid.Columns.Item(3).TitleObject.Caption = "# Solicitud" '"# Pago"
            oGrid.Columns.Item(3).Editable = False

            oGrid.Columns.Item(4).Description = "Código proovedor"
            oGrid.Columns.Item(4).TitleObject.Caption = "Código proovedor"
            oGrid.Columns.Item(4).Editable = False
            Dim oEditTextColumn3 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item(4)
            oEditTextColumn3.LinkedObjectType = 2

            oGrid.Columns.Item(5).Description = "Nombre proveedor"
            oGrid.Columns.Item(5).TitleObject.Caption = "Nombre proveedor"
            oGrid.Columns.Item(5).Editable = False

            oGrid.Columns.Item(6).Description = "Fecha pago"
            oGrid.Columns.Item(6).TitleObject.Caption = "Fecha pago"
            oGrid.Columns.Item(6).Editable = False

            oGrid.Columns.Item(7).Description = "Total pagado"
            oGrid.Columns.Item(7).TitleObject.Caption = "Total pagado"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).RightJustified = True

            oGrid.Columns.Item(8).Description = "Forma pago"
            oGrid.Columns.Item(8).TitleObject.Caption = "Forma pago"
            oGrid.Columns.Item(8).Editable = False

            oGrid.Columns.Item(9).Description = "Moneda"
            oGrid.Columns.Item(9).TitleObject.Caption = "Moneda"
            oGrid.Columns.Item(9).Editable = False
            oGrid.Columns.Item(9).Visible = False

            oGrid.Columns.Item(10).Description = "Tipo Cuenta"
            oGrid.Columns.Item(10).TitleObject.Caption = "Tipo Cuenta"
            oGrid.Columns.Item(10).Editable = False

            oGrid.Columns.Item(11).Description = "Numero Cuenta"
            oGrid.Columns.Item(11).TitleObject.Caption = "Numero Cuenta"
            oGrid.Columns.Item(11).Editable = False

            oGrid.Columns.Item(12).Description = "Referencia"
            oGrid.Columns.Item(12).TitleObject.Caption = "Referencia"
            oGrid.Columns.Item(12).Editable = False

            oGrid.Columns.Item(13).Description = "TipoCliente"
            oGrid.Columns.Item(13).TitleObject.Caption = "TipoCliente"
            oGrid.Columns.Item(13).Editable = False
            oGrid.Columns.Item(13).Visible = False

            oGrid.Columns.Item(14).Description = "NumeroID"
            oGrid.Columns.Item(14).TitleObject.Caption = "NumeroID"
            oGrid.Columns.Item(14).Editable = False
            oGrid.Columns.Item(14).Visible = False

            oGrid.Columns.Item(15).Description = "Cod Banco"
            oGrid.Columns.Item(15).TitleObject.Caption = "Cod Banco"
            oGrid.Columns.Item(15).Editable = False

            oGrid.Columns.Item(16).Description = "CodCuentaEmpresa"
            oGrid.Columns.Item(16).TitleObject.Caption = "CodCuentaEmpresa"
            oGrid.Columns.Item(16).Editable = False
            oGrid.Columns.Item(16).Visible = False

            oGrid.Columns.Item(17).Description = "CodBancoEmpresa"
            oGrid.Columns.Item(17).TitleObject.Caption = "CodBancoEmpresa"
            oGrid.Columns.Item(17).Editable = False
            oGrid.Columns.Item(17).Visible = False

            oGrid.Columns.Item(18).Description = "Email"
            oGrid.Columns.Item(18).TitleObject.Caption = "Email"
            oGrid.Columns.Item(18).Editable = False

            oGrid.Columns.Item(19).Description = "FactReferencia"
            oGrid.Columns.Item(19).TitleObject.Caption = "FactReferencia"
            oGrid.Columns.Item(19).Editable = False

            oGrid.AutoResizeColumns()

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error FormatoGrid:" & ex.Message.ToString, "frmCashManagement")
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

    Private Function GuardaUDOCM(ByRef DocEntryUDOCM As Integer) As Boolean

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        Try
            rsboApp.StatusBar.SetText("Creando registro de generación de archivo UDO...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Utilitario.Util_Log.Escribir_Log("Creando registro de generación de archivo UDO...", "frmCashManagement")

            oForm = rsboApp.Forms.Item("frmCashManagement")

            Dim lblRA As SAPbouiCOM.StaticText = oForm.Items.Item("lblRA").Specific
            Dim lblNF As SAPbouiCOM.StaticText = oForm.Items.Item("lblNF").Specific
            Dim lblTF As SAPbouiCOM.StaticText = oForm.Items.Item("lblTF").Specific

            oCompanyService = rCompany.GetCompanyService
            oGeneralService = oCompanyService.GetGeneralService("SSCMCASH")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            Utilitario.Util_Log.Escribir_Log("Obteniendo UDO - SSCMCASH", "frmCashManagement")

            Utilitario.Util_Log.Escribir_Log("U_FecArc " & FechaUDO.ToString, "frmCashManagement")
            oGeneralData.SetProperty("U_FecArc", FechaUDO)

            Utilitario.Util_Log.Escribir_Log("U_RutaArc " & lblRA.Caption, "frmCashManagement")
            oGeneralData.SetProperty("U_RutaArc", lblRA.Caption)

            Utilitario.Util_Log.Escribir_Log("U_NumPagos " & lblNF.Caption, "frmCashManagement")
            oGeneralData.SetProperty("U_NumPagos", CInt(lblNF.Caption))

            Utilitario.Util_Log.Escribir_Log("U_TotPagos " & lblTF.Caption, "frmCashManagement")
            oGeneralData.SetProperty("U_TotPagos", Convert.ToDouble(lblTF.Caption))

            Utilitario.Util_Log.Escribir_Log("U_Banco " & oForm.Items.Item("cbxBan").Specific.Value.ToString(), "frmCashManagement")
            oGeneralData.SetProperty("U_Banco", oForm.Items.Item("cbxBan").Specific.Value.ToString())

            oGeneralData.SetProperty("U_Estado", "Abierto")

            oChildren = oGeneralData.Child("SS_CM_DET1")
            odt = oForm.DataSources.DataTables.Item("dtDocs")
            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                Dim check As String = odt.GetValue(0, i).ToString()

                If check = "Y" Then

                    oChild = oChildren.Add

                    Utilitario.Util_Log.Escribir_Log("U_DocEntryP " + odt.GetValue(2, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_DocEntryP", odt.GetValue(2, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_CodProv " + odt.GetValue(4, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_CodProv", odt.GetValue(4, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_FecPag " + odt.GetValue(6, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_FecPag", odt.GetValue(6, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_TotPag " + odt.GetValue(7, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_TotPag", Convert.ToDouble(odt.GetValue(7, i).ToString()))

                    Utilitario.Util_Log.Escribir_Log("U_ForPag " + odt.GetValue(8, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_ForPag", odt.GetValue(8, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_Moneda " + odt.GetValue(9, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_Moneda", odt.GetValue(9, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_TipCta " + odt.GetValue(10, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_TipCta", odt.GetValue(10, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_NumCta " + odt.GetValue(11, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_NumCta", odt.GetValue(11, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_Ref " + odt.GetValue(12, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_Ref", odt.GetValue(12, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_TipCli " + odt.GetValue(13, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_TipCli", odt.GetValue(13, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_NumID " + odt.GetValue(14, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_NumID", odt.GetValue(14, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_CodBco " + odt.GetValue(15, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_CodBco", odt.GetValue(15, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_CodCtaEm " + odt.GetValue(16, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_CodCtaEm", odt.GetValue(16, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_CodBcoEm " + odt.GetValue(17, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_CodBcoEm", odt.GetValue(17, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_Email " + odt.GetValue(18, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_Email", odt.GetValue(18, i).ToString())

                    Utilitario.Util_Log.Escribir_Log("U_FacRef " + odt.GetValue(19, i).ToString(), "frmCashManagement")
                    oChild.SetProperty("U_FacRef", odt.GetValue(19, i).ToString())

                End If

            Next

            oGeneralParams = oGeneralService.Add(oGeneralData)
            DocEntryUDOCM = oGeneralParams.GetProperty("DocEntry")
            rsboApp.StatusBar.SetText("Se creo registro de generación de archivo y UDO " & DocEntryUDOCM.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Ocurrio un error en el registro de generación de archivo UDO: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Utilitario.Util_Log.Escribir_Log("Ocurrio un error en el registro de generación de archivo UDO: " & ex.Message, "frmCashManagement")
            Return False
        End Try
    End Function

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            Dim typeExx, idFormm As String
            typeExx = oFuncionesB1.FormularioActivo(idFormm)

            If typeExx = "frmCashManagement" Then
                oForm.Freeze(True)

                If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") Then

                    If Not pVal.BeforeAction Then

                        BuscaRegistros()

                    End If
                ElseIf pVal.MenuUID = "1281" Then

                    oForm.Items.Item("DocEntry").Enabled = True
                    oForm.Items.Item("1").Visible = True

                ElseIf pVal.MenuUID = "1282" Then

                    oForm.Items.Item("DocEntry").Enabled = False
                    oForm.Items.Item("1").Visible = True

                    BloqueaControles(True, False, False)
                    oForm.Items.Item("1").Visible = False

                End If

                oForm.Freeze(False)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BuscaRegistros()
        Try
            Dim DocEntry As String = oForm.DataSources.DBDataSources.Item("@SS_CM_CAB").GetValue("DocEntry", 0) ' oDBDS.GetValue("U_RutaArc", 0).Trim()

            CargarDatos(DocEntry)

            Dim RutaArchivo As String = oForm.DataSources.DBDataSources.Item("@SS_CM_CAB").GetValue("U_RutaArc", 0) ' oDBDS.GetValue("U_RutaArc", 0).Trim()
            Dim NumPagos As Double = CInt(oForm.DataSources.DBDataSources.Item("@SS_CM_CAB").GetValue("U_NumPagos", 0)) ' oDBDS.GetValue("U_RutaArc", 0).Trim()
            Dim TotalPagos As Double = CDbl(oForm.DataSources.DBDataSources.Item("@SS_CM_CAB").GetValue("U_TotPagos", 0)) ' oDBDS.GetValue("U_RutaArc", 0).Trim()
            Dim Estado As String = oForm.DataSources.DBDataSources.Item("@SS_CM_CAB").GetValue("U_Estado", 0) ' oDBDS.GetValue("U_RutaArc", 0).Trim()

            Dim lblRA As SAPbouiCOM.StaticText = oForm.Items.Item("lblRA").Specific
            lblRA.Caption = CStr(RutaArchivo)

            Dim lblNF As SAPbouiCOM.StaticText = oForm.Items.Item("lblNF").Specific
            lblNF.Caption = CStr(NumPagos)

            Dim lblTF As SAPbouiCOM.StaticText = oForm.Items.Item("lblTF").Specific
            lblTF.Caption = CStr(TotalPagos)

            If Estado = "Abierto" Then
                BloqueaControles(False, True, True)
            Else
                BloqueaControles(False, True, False)
            End If

            oForm.Items.Item("oGrid").Enabled = False
            oForm.Items.Item("1").Visible = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BloqueaControles(ControlesBusquea As Boolean, ControlesActualizacion As Boolean, ByVal Optional inhabilita As Boolean = False)
        Try
            Dim focus As SAPbouiCOM.Item = oForm.Items.Item("foco")
            focus.Click()

            oForm.Items.Item("Item_1").Visible = ControlesBusquea
            oForm.Items.Item("finicial").Visible = ControlesBusquea
            oForm.Items.Item("Item_2").Visible = ControlesBusquea
            oForm.Items.Item("ffinal").Visible = ControlesBusquea
            oForm.Items.Item("Item_8").Visible = ControlesBusquea
            oForm.Items.Item("cbxBan").Visible = ControlesBusquea
            oForm.Items.Item("btnBuscar").Visible = ControlesBusquea
            oForm.Items.Item("btnGA").Visible = ControlesBusquea

            oForm.Items.Item("Item_12").Visible = ControlesActualizacion
            oForm.Items.Item("cbxEstado").Visible = ControlesActualizacion
            oForm.Items.Item("Item_5").Visible = ControlesActualizacion
            oForm.Items.Item("txtIdBan").Visible = ControlesActualizacion

            oForm.Items.Item("cbxEstado").Enabled = inhabilita
            oForm.Items.Item("txtIdBan").Enabled = inhabilita

        Catch ex As Exception

        End Try
    End Sub

End Class
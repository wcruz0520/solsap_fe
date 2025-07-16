Imports System.IO

Public Class frmDinardap

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private odt As SAPbouiCOM.DataTable


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub


    Public Sub CargaFormularioDinardap()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmDinardap") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmDinardap.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmDinardap").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmDinardap")

            oForm.Freeze(True)

            Dim txtFchIniD As SAPbouiCOM.EditText
            txtFchIniD = oForm.Items.Item("finicialD").Specific
            txtFchIniD.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim txtFchFinD As SAPbouiCOM.EditText
            txtFchFinD = oForm.Items.Item("ffinalD").Specific
            txtFchFinD.Value = DateTime.Now.ToString("yyyyMMdd")

            Dim imgSRID As SAPbouiCOM.PictureBox
            imgSRID = oForm.Items.Item("imgLogo").Specific
            imgSRID.Picture = System.Windows.Forms.Application.StartupPath & "\LogoSS.png"

            CargarGrid()

            ' CargaDatos()

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargarGrid()
        oForm.Freeze(True)


        Dim txtfinicialD As SAPbouiCOM.EditText = oForm.Items.Item("finicialD").Specific
        Dim txtffinalD As SAPbouiCOM.EditText = oForm.Items.Item("ffinalD").Specific
        'Dim schk As SAPbouiCOM.CheckBox = oForm.Items.Item("schk").Specific

        If (String.IsNullOrEmpty(txtfinicialD.Value) Or String.IsNullOrEmpty(txtffinalD.Value)) Then
            rsboApp.SetStatusBarMessage("Debe ingresar un rango de fechas a consultar!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        Dim dfechaDesdeD As Date
        Dim dfechaHastaD As Date
        Dim sQueryD As String = ""

        If Not oFuncionesB1.BobStringToDate(txtfinicialD.Value, dfechaDesdeD) Then
            rsboApp.SetStatusBarMessage("El formato de las fechas es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If

        If Not oFuncionesB1.BobStringToDate(txtffinalD.Value, dfechaHastaD) Then
            rsboApp.SetStatusBarMessage("El formato de las fechas es incorrecto", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            oForm.Freeze(False)
            Exit Sub
        End If




        'If String.IsNullOrWhiteSpace(sQueryD) Then

        '    rsboApp.SetStatusBarMessage("(GS) Pantalla Inactiva , Por Favor Revisar la Parametrizacion de los Documentos Enviados", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

        'End If

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            sQueryD = "CALL SS_DINARDAP ("
            sQueryD += Functions.FuncionesB1.FechaSql(dfechaDesdeD)
            sQueryD += "," + Functions.FuncionesB1.FechaSql(dfechaHastaD) + ")"
        Else
            sQueryD = "EXEC SS_DINARDAP "
            sQueryD += Functions.FuncionesB1.FechaSql(dfechaDesdeD)
            sQueryD += "," + Functions.FuncionesB1.FechaSql(dfechaHastaD)
        End If


        Try
            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Try
                oGrid.DataTable.ExecuteQuery(sQueryD)
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query Documentos Enviados Log:" + ex.Message().ToString() + "-QUERY: " + sQueryD, "frmDocumentosEnviados")
            End Try


            If oGrid.DataTable.Rows.Count > 0 Then

                For y As Integer = 0 To oGrid.Columns.Count - 1


                    oGrid.Columns.Item(y).Editable = False


                Next

            End If

            ''oGrid.Columns.Item(0).Description = "Tipo Documento"
            ''oGrid.Columns.Item(0).TitleObject.Caption = "Tipo Documento"
            ''oGrid.Columns.Item(0).Editable = False

            ''oGrid.Columns.Item(1).Description = "#"
            ''oGrid.Columns.Item(1).TitleObject.Caption = "#"
            ''oGrid.Columns.Item(1).Editable = False

            'oGrid.Columns.Item(2).Description = "DocEntry"
            'oGrid.Columns.Item(2).TitleObject.Caption = "DocEntry"


            'Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            'oEditTextColumn = oGrid.Columns.Item(2)
            'oEditTextColumn.LinkedObjectType = 18


            ''oGrid.Columns.Item(3).Description = "Fecha Emisión"
            ''oGrid.Columns.Item(3).TitleObject.Caption = "Fecha Emisión"
            ''oGrid.Columns.Item(3).Editable = False

            ''oGrid.Columns.Item(4).Description = "Doc. Num."
            ''oGrid.Columns.Item(4).TitleObject.Caption = "Doc. Num."
            ''oGrid.Columns.Item(4).Editable = False

            ''oGrid.Columns.Item(5).Description = "Cliente"
            ''oGrid.Columns.Item(5).TitleObject.Caption = "Cliente"
            ''oGrid.Columns.Item(5).Editable = False


            ''oGrid.Columns.Item(6).Description = "Doc. Total"
            ''oGrid.Columns.Item(6).TitleObject.Caption = "Doc. Total"
            ''oGrid.Columns.Item(6).Editable = False
            ''oGrid.Columns.Item(6).RightJustified = True


            ''oGrid.Columns.Item(7).Description = "Estado Documento"
            ''oGrid.Columns.Item(7).TitleObject.Caption = "Estado Documento"
            ''oGrid.Columns.Item(7).Editable = False

            ''oGrid.Columns.Item(8).Description = "CUF"
            ''oGrid.Columns.Item(8).TitleObject.Caption = "CUF"
            ''oGrid.Columns.Item(8).Editable = False

            ''oGrid.Columns.Item(9).Description = "EXT1"
            ''oGrid.Columns.Item(9).TitleObject.Caption = "EXT1"
            ''oGrid.Columns.Item(9).Editable = False
            ''oGrid.Columns.Item(9).Visible = False

            ''oGrid.Columns.Item(10).Description = "EXT2"
            ''oGrid.Columns.Item(10).TitleObject.Caption = "EXT2"
            ''oGrid.Columns.Item(10).Editable = False
            ''oGrid.Columns.Item(10).Visible = False

            ''oGrid.Columns.Item(11).Description = "EXT3"
            ''oGrid.Columns.Item(11).TitleObject.Caption = "EXT3"
            ''oGrid.Columns.Item(11).Editable = False
            ''oGrid.Columns.Item(11).Visible = False

            ''oGrid.Columns.Item(12).Description = "EXT4"
            ''oGrid.Columns.Item(12).TitleObject.Caption = "EXT4"
            ''oGrid.Columns.Item(12).Editable = False
            ''oGrid.Columns.Item(12).Visible = False


            'oGrid.CollapseLevel = 1
            'oGrid.AutoResizeColumns()
            'schk.Checked = False

            oForm.Freeze(False)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.FormTypeEx = "frmDinardap" Then

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                        Case "btnBuscar"
                            If pVal.BeforeAction = False Then
                                CargarGrid()
                            End If

                        Case "btnGenTXT"
                            If pVal.BeforeAction = False Then

                                Dim selectFileDialog As New SelectFileDialog("C:\", "", "|*", DialogType.FOLDER)
                                selectFileDialog.Open()
                                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFolder) Then
                                    Dim RutaArchivoD = GenerarArchivoTXT(selectFileDialog.SelectedFolder)
                                    If RutaArchivoD <> "" Then

                                        Dim respD = rsboApp.MessageBox("El archivo .txt se genero exitosamente!!", 1, "Abrir Directorio", "Abrir txt", "Cerrar")
                                        Dim ProcD As New Process()
                                        If respD = 1 Then

                                            ProcD.StartInfo.FileName = selectFileDialog.SelectedFolder
                                            ProcD.Start()
                                            ProcD.Dispose()

                                        ElseIf respD = 2 Then

                                            ProcD.StartInfo.FileName = RutaArchivoD
                                            ProcD.Start()
                                            ProcD.Dispose()

                                        End If

                                    End If

                                End If

                            End If
                        Case Else

                    End Select


                Case Else

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

    Private Function GenerarArchivoTXT(rutaD As String) As String

        Try
            oForm = rsboApp.Forms.Item("frmDinardap")
            oForm.Freeze(True)

            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            rsboApp.SetStatusBarMessage(NombreAddon + " - Este proceso puede tardar, por favor espere un momento..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            rutaD += "\Dinardap_" + DateAndTime.Year(Date.Now).ToString + "-" + DateAndTime.Month(Date.Now).ToString + "-" + DateAndTime.Day(Date.Now).ToString + ".txt"
            Dim archivoD As TextWriter = New StreamWriter(rutaD)

            Dim Linea As String = ""
            Dim Código_Entidad As String = ""
            Dim Fecha_de_datos As String = ""
            Dim Tipo_de_Identificación As String = ""
            Dim Identificación As String = ""
            Dim Nombre As String = ""
            Dim Clase_de_sujeto As String = ""
            Dim Provincia As String = ""
            Dim Cantón As String = ""
            Dim Parroquia As String = ""
            Dim Sexo As String = ""
            Dim Estado_Civil As String = ""
            Dim Origen_de_ingresos As String = ""
            Dim Número_de_Operación As String = ""
            Dim Valor_Operación As String = ""
            Dim Saldo_Operación As String = ""
            Dim Fecha_Concesión As String = ""
            Dim Fecha_de_Vencimiento As String = ""
            Dim Fecha_Exigible As String = ""
            Dim Plazo_Operación As String = ""
            Dim Periodicidad_de_Pago As String = ""
            Dim Dias_de_Morosidad As String = ""
            Dim Monto_Morosidad As String = ""
            Dim Monto_de_Interes_en_Mora As String = ""
            Dim Valor_por_Vencer_de_1_a_30 As String = ""
            Dim Valor_por_Vencer_de_31_a_90 As String = ""
            Dim Valor_por_Vencer_de_91_a_180 As String = ""
            Dim Valor_por_Vencer_de_181_a_360 As String = ""
            Dim Valor_por_Vencer_mas_de_360 As String = ""
            Dim Valor_vencido_de_1_a_30 As String = ""
            Dim Valor_vencido_de_31_a_90 As String = ""
            Dim Valor_vencido_de_91_a_180 As String = ""
            Dim Valor_vencido_de_181_a_360 As String = ""
            Dim Valor_vencido_mas_de_360 As String = ""
            Dim Valor_demanda_judicial As String = ""
            Dim Cartera_Castigada As String = ""
            Dim Cuota_Crédito As String = ""
            Dim Fecha_Cancelación As String = ""
            Dim Forma_de_Cancelación As String = ""

            Dim oGridDetD As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")
            Dim estado As String = ""
            For i As Integer = 0 To oGridDetD.Rows.Count - 1

                Código_Entidad = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Código Entidad", i)), "", oGridDetD.GetValue("Código Entidad", i))
                Fecha_de_datos = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Fecha de datos", i)), "", oGridDetD.GetValue("Fecha de datos", i))
                Tipo_de_Identificación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Tipo de Identificación", i)), "", oGridDetD.GetValue("Tipo de Identificación", i))
                Identificación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Identificación", i)), "", oGridDetD.GetValue("Identificación", i))
                Nombre = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Nombre", i)), "", oGridDetD.GetValue("Nombre", i))
                Clase_de_sujeto = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Clase de sujeto", i)), "", oGridDetD.GetValue("Clase de sujeto", i))
                Provincia = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Provincia", i)), "", oGridDetD.GetValue("Provincia", i))
                Cantón = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Cantón", i)), "", oGridDetD.GetValue("Cantón", i))
                Parroquia = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Parroquia", i)), "", oGridDetD.GetValue("Parroquia", i))
                Sexo = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Sexo", i)), "", oGridDetD.GetValue("Sexo", i))
                Estado_Civil = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Estado Civil", i)), "", oGridDetD.GetValue("Estado Civil", i))
                Origen_de_ingresos = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Origen de ingresos", i)), "", oGridDetD.GetValue("Origen de ingresos", i))
                Número_de_Operación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Número de Operación", i)), "", oGridDetD.GetValue("Número de Operación", i))
                Valor_Operación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor Operación", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor Operación", i)).ToString("###0.00"))
                Saldo_Operación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Saldo Operación", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Saldo Operación", i)).ToString("###0.00"))
                Fecha_Concesión = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Fecha Concesión", i)), "", oGridDetD.GetValue("Fecha Concesión", i))
                Fecha_de_Vencimiento = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Fecha de Vencimiento", i)), "", oGridDetD.GetValue("Fecha de Vencimiento", i))
                Fecha_Exigible = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Fecha Exigible", i)), "", oGridDetD.GetValue("Fecha Exigible", i))
                Plazo_Operación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Plazo Operación", i)), "", oGridDetD.GetValue("Plazo Operación", i))
                Periodicidad_de_Pago = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Periodicidad de Pago", i)), "", oGridDetD.GetValue("Periodicidad de Pago", i))
                Dias_de_Morosidad = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Dias de Morosidad", i)), "", oGridDetD.GetValue("Dias de Morosidad", i))
                Monto_Morosidad = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Monto Morosidad", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Monto Morosidad", i)).ToString("###0.00"))
                Monto_de_Interes_en_Mora = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Monto de Interes en Mora", i)), "0.00", Convert.ToDecimal(oGridDetD.GetValue("Monto de Interes en Mora", i)).ToString("###0.00"))
                Valor_por_Vencer_de_1_a_30 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor por Vencer de 1 a 30", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor por Vencer de 1 a 30", i)).ToString("###0.00"))
                Valor_por_Vencer_de_31_a_90 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor por Vencer de 31 a 90", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor por Vencer de 31 a 90", i)).ToString("###0.00"))
                Valor_por_Vencer_de_91_a_180 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor por Vencer de 91 a 180", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor por Vencer de 91 a 180", i)).ToString("###0.00"))
                Valor_por_Vencer_de_181_a_360 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor por Vencer de 181 a 360", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor por Vencer de 181 a 360", i)).ToString("###0.00"))
                Valor_por_Vencer_mas_de_360 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor por Vencer mas de 360", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor por Vencer mas de 360", i)).ToString("###0.00"))
                Valor_vencido_de_1_a_30 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor vencido de 1 a 30", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor vencido de 1 a 30", i)).ToString("###0.00"))
                Valor_vencido_de_31_a_90 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor vencido de 31 a 90", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor vencido de 31 a 90", i)).ToString("###0.00"))
                Valor_vencido_de_91_a_180 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor vencido de 91 a 180", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor vencido de 91 a 180", i)).ToString("###0.00"))
                Valor_vencido_de_181_a_360 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor vencido de 181 a 360", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor vencido de 181 a 360", i)).ToString("###0.00"))
                Valor_vencido_mas_de_360 = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor vencido mas de 360", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor vencido mas de 360", i)).ToString("###0.00"))
                Valor_demanda_judicial = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Valor demanda judicial", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Valor demanda judicial", i)).ToString("###0.00"))
                Cartera_Castigada = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Cartera Castigada", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Cartera Castigada", i)).ToString("###0.00"))
                Cuota_Crédito = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Cuota Crédito", i)), "", Convert.ToDecimal(oGridDetD.GetValue("Cuota Crédito", i)).ToString("###0.00"))
                Fecha_Cancelación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Fecha Cancelación", i)), "", oGridDetD.GetValue("Fecha Cancelación", i))
                Forma_de_Cancelación = IIf(String.IsNullOrEmpty(oGridDetD.GetValue("Forma de Cancelación", i)), "", oGridDetD.GetValue("Forma de Cancelación", i))
                'rsboApp.MessageBox(result)
                Linea = Código_Entidad & "|" & Fecha_de_datos & "|" & Tipo_de_Identificación & "|" & Identificación & "|"
                Linea += Nombre & "|" & Clase_de_sujeto & "|" & Provincia & "|" & Cantón & "|" & Parroquia & "|"
                Linea += Sexo & "|" & Estado_Civil & "|" & Origen_de_ingresos & "|" & Número_de_Operación & "|" & Valor_Operación & "|"
                Linea += Saldo_Operación & "|" & Fecha_Concesión & "|" & Fecha_de_Vencimiento & "|" & Fecha_Exigible & "|" & Plazo_Operación & "|"
                Linea += Periodicidad_de_Pago & "|" & Dias_de_Morosidad & "|" & Monto_Morosidad & "|" & Monto_de_Interes_en_Mora & "|" & Valor_por_Vencer_de_1_a_30 & "|"
                Linea += Valor_por_Vencer_de_31_a_90 & "|" & Valor_por_Vencer_de_91_a_180 & "|" & Valor_por_Vencer_de_181_a_360 & "|" & Valor_por_Vencer_mas_de_360 & "|"
                Linea += Valor_vencido_de_1_a_30 & "|" & Valor_vencido_de_31_a_90 & "|" & Valor_vencido_de_91_a_180 & "|" & Valor_vencido_de_181_a_360 & "|"
                Linea += Valor_vencido_mas_de_360 & "|" & Valor_demanda_judicial & "|" & Cartera_Castigada & "|" & Cuota_Crédito & "|"
                Linea += Fecha_Cancelación & "|" & Forma_de_Cancelación

                archivoD.WriteLine(Linea)

            Next

            archivoD.Close()
            oForm.Freeze(False)
            Return rutaD
            'Utilitario.Util_Log.Escribir_Log("Cantidad de Documentos Seleccionados : " + contador.ToString(), "frmDocumentosEnviados")
        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " - Error al generar txt Dinardap " & ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
            Return ""
        Finally
            oForm.Freeze(False)

        End Try
        'Return True
    End Function

End Class

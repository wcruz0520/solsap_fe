Imports System.IO

Public Class frmCargaRPT

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private odt As SAPbouiCOM.DataTable

    Private fol As SAPbouiCOM.Folder

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioCargaRPT()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmCargaRPT") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmCargaRPT.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmCargaRPT").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmCargaRPT")


            CargaDatos()

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub
    Private Sub CargaDatos()
        'oForm = rsboApp.Forms.Item("frmFindCert")
        oForm.Freeze(True)
        Try
            Dim ACTUALIZA As Integer = 0
            ' DATA TABLE CABECERA
            Try
                oForm.DataSources.DataTables.Add("odt")
            Catch ex As Exception
            End Try
            Dim QueryFC As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                QueryFC = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryFC += "FROM ""@SS_CONFD"" A INNER JOIN "
                QueryFC += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryFC += " WHERE  B.""U_Modulo"" = '" & NombreAddon & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'NORPT'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = '" & NombreAddon & "' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'NORPT'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")

            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                If odt.GetValue("U_Nombre", i).ToString().Equals("RPTFacturasCompras") Then
                    oForm.Items.Item("txtRPT1").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RPTFormularioSRI103") Then
                    oForm.Items.Item("txtRPT2").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RPTListadoVentas") Then
                    oForm.Items.Item("txtRPT3").Specific.value = odt.GetValue("U_Valor", i).ToString()
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("RPTRetencionClientes") Then
                    oForm.Items.Item("txtRPT4").Specific.value = odt.GetValue("U_Valor", i).ToString()
                End If
                ACTUALIZA = 1
            Next

            If ACTUALIZA = 1 Then
                Dim obtnGrabar As SAPbouiCOM.Button
                obtnGrabar = oForm.Items.Item("obtnGrabar").Specific
                obtnGrabar.Caption = "Actualizar"
            End If

        Catch ex As Exception
            rsboApp.MessageBox(ex.Message.ToString())
        Finally
            oForm.Freeze(False)
            ' mors = Nothing
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

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.FormTypeEx = "frmCargaRPT" Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If Not pVal.Before_Action Then
                        Select Case pVal.ItemUID
                            Case "obtnGrabar"

                                Dim oConfiguracion As Entidades.Configuracion
                                Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                oConfiguracion = New Entidades.Configuracion
                                oConfiguracion.Modulo = NombreAddon
                                oConfiguracion.Tipo = "PARAMETROS"
                                oConfiguracion.SubTipo = "NORPT"
                                olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)

                                olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RPTFacturasCompras", oForm.Items.Item("txtRPT1").Specific.value.ToString))
                                olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RPTFormularioSRI103", oForm.Items.Item("txtRPT2").Specific.value.ToString))
                                olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RPTListadoVentas", oForm.Items.Item("txtRPT3").Specific.value.ToString))
                                olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("RPTRetencionClientes", oForm.Items.Item("txtRPT4").Specific.value.ToString))

                                oConfiguracion.Detalle = olistaDetalleConfiguracion
                                GuardaCONF(oConfiguracion)

                                ' de una vez guardo en la variable global para no reiniciar addon
                                'Functions.VariablesGlobales._gCertificadoSAT = oForm.Items.Item("txtb64cer").Specific.value.ToString

                                oForm.Items.Item("obtnGrabar").Visible = False
                                oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                oForm.Items.Item("2").Specific.Caption = "OK"

                            Case "btnOpen"

                                Dim selectFileDialog As New SelectFileDialog("C:\", "", "|*.rpt", DialogType.OPEN)
                                selectFileDialog.Open()

                                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then


                                    Try



                                        If oForm.Items.Item("f1").Specific.Selected = True Then

                                            oForm.Items.Item("txtRPT1").Specific.Value = Convert.ToBase64String(File.ReadAllBytes(selectFileDialog.SelectedFile))

                                            rsboApp.SetStatusBarMessage("Cargados bytes en TXT folder: " & oForm.Items.Item("f1").Specific.Caption.ToString,, False)

                                        ElseIf oForm.Items.Item("f2").Specific.Selected = True Then

                                            oForm.Items.Item("txtRPT2").Specific.Value = Convert.ToBase64String(File.ReadAllBytes(selectFileDialog.SelectedFile))

                                            rsboApp.SetStatusBarMessage("Cargados bytes en TXT folder: " & oForm.Items.Item("f2").Specific.Caption.ToString,, False)

                                        ElseIf oForm.Items.Item("f3").Specific.Selected = True Then

                                            oForm.Items.Item("txtRPT3").Specific.Value = Convert.ToBase64String(File.ReadAllBytes(selectFileDialog.SelectedFile))

                                            rsboApp.SetStatusBarMessage("Cargados bytes en TXT folder: " & oForm.Items.Item("f3").Specific.Caption.ToString,, False)

                                        ElseIf oForm.Items.Item("f4").Specific.Selected = True Then

                                            oForm.Items.Item("txtRPT4").Specific.Value = Convert.ToBase64String(File.ReadAllBytes(selectFileDialog.SelectedFile))

                                            rsboApp.SetStatusBarMessage("Cargados bytes en TXT folder: " & oForm.Items.Item("f4").Specific.Caption.ToString,, False)
                                        End If




                                    Catch ex As Exception

                                        Me.rsboApp.MessageBox("Error al Cargar el Archivo RPT " + ex.Message)

                                    End Try


                                End If

                        End Select
                    End If
            End Select
        End If

    End Sub


    Private Sub GuardaCONF(ByVal oConfiguracion As Entidades.Configuracion)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

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
            If Not CodeExist = "0" Then ' SI EXISTE, ELIMINO Y ACTUALIZO

                ' SI EXISTE ELIMINA PARA VOLVER A CREAR
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONFLOC")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)
                oGeneralService.Delete(oGeneralParams)

                'CREA NUEVAMENTE EL REGISTRO
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

        End Try
    End Sub

End Class

Imports SAPbobsCOM

Public Class frmConsultasDbLE

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private odt As SAPbouiCOM.DataTable


    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioParametrosConsultasBD()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConsultasDbLE") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmConsultasDbLE.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConsultasDbLE").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmConsultasDbLE")


            CargaDatos()

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddonLOC + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargaDatos()
        oForm = rsboApp.Forms.Item("frmConsultasDbLE")
        oForm.Freeze(True)
        Try
            Dim ACTUALIZA As Integer = 0
            ' DATA TABLE CABECERA
            Try
                oForm.DataSources.DataTables.Add("odt")
            Catch ex As Exception
            End Try
            Dim QueryFC As String = ""
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                QueryFC = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryFC += "FROM ""@SS_CONFD"" A INNER JOIN "
                QueryFC += """@SS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryFC += " WHERE  B.""U_Modulo"" = '" & NombreAddonLOC & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'BD'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@SS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@SS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = '" & NombreAddonLOC & "' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'BD'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")

            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                If odt.GetValue("U_Nombre", i).ToString().Equals("ComprasQRY") Then
                    oForm.Items.Item("txtcompra").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("VentasQRY") Then
                    oForm.Items.Item("txtventa").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("CalculoFolioQRY") Then
                    oForm.Items.Item("txtfolio").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)


                End If
                ACTUALIZA = 1
            Next

            If ACTUALIZA = 1 Then
                Dim obtnGrabar As SAPbouiCOM.Button
                obtnGrabar = oForm.Items.Item("obtnGrabDB").Specific
                obtnGrabar.Caption = "Actualizar"
            End If

        Catch ex As Exception
            rsboApp.SetStatusBarMessage(NombreAddonLOC & ": Validar que esten Parametrizada las Consultas ")
        Finally
            oForm.Freeze(False)
            'mors = Nothing
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        Try
            If pVal.FormTypeEx = "frmConsultasDbLE" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "obtnGrabDB"

                                    Try


                                        Dim oConfiguracion As Entidades.Configuracion
                                        Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                        oForm = rsboApp.Forms.Item("frmConsultasDbLE")

                                        Dim txtcompra As SAPbouiCOM.EditText
                                        txtcompra = oForm.Items.Item("txtcompra").Specific '

                                        Dim txtventa As SAPbouiCOM.EditText
                                        txtventa = oForm.Items.Item("txtventa").Specific '

                                        Dim txtfolio As SAPbouiCOM.EditText
                                        txtfolio = oForm.Items.Item("txtfolio").Specific '


                                        'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
                                        oConfiguracion = New Entidades.Configuracion
                                        oConfiguracion.Modulo = NombreAddonLOC
                                        oConfiguracion.Tipo = "PARAMETROS"
                                        oConfiguracion.SubTipo = "BD"
                                        olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ComprasQRY", Utilitario.Util_Encriptador.Encriptar(txtcompra.Value.ToString(), sKey)))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("VentasQRY", Utilitario.Util_Encriptador.Encriptar(txtventa.Value.ToString(), sKey)))

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("CalculoFolioQRY", Utilitario.Util_Encriptador.Encriptar(txtfolio.Value.ToString(), sKey)))

                                        oConfiguracion.Detalle = olistaDetalleConfiguracion
                                        GuardaCONF(oConfiguracion)

                                        oForm.Items.Item("obtnGrabDB").Visible = False
                                        oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabDB").Left
                                        oForm.Items.Item("2").Specific.Caption = "OK"

                                    Catch ex As Exception

                                    End Try



                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmConsultasBD")
            System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try

    End Sub

    Public Sub GuardaCONF(ByVal oConfiguracion As Entidades.Configuracion)

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

End Class

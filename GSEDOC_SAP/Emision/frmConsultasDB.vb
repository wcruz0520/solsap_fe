
Imports SAPbobsCOM
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Public Class frmConsultasDB
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Private mors As SAPbobsCOM.Recordset = Nothing
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim odt As SAPbouiCOM.DataTable
    Dim Alia As String = ""
    Dim btnqry As SAPbouiCOM.ButtonCombo
    Dim cbxProveedor As SAPbouiCOM.ComboBox
    Dim cbxTipoWS As SAPbouiCOM.ComboBox

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioParametrosConsultasBD()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConsultasDB") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmConsultasDB.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConsultasDB").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmConsultasDB")



            CargaDatos()

            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargaDatos()
        '  oForm = rsboApp.Forms.Item("frmConsultasDB")
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
                QueryFC += "FROM ""@GS_CONFD"" A INNER JOIN "
                QueryFC += """@GS_CONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryFC += " WHERE  B.""U_Modulo"" = '" & NombreAddon & "' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'BD'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@GS_CONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@GS_CONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = '" & NombreAddon & "' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'BD'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")

            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                If odt.GetValue("U_Nombre", i).ToString().Equals("Query_FacturaSeccion01") Then
                    oForm.Items.Item("txtfacs01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_FacturaSeccion02") Then
                    oForm.Items.Item("txtfacs02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_FacturaAnticipoSeccion01") Then
                    oForm.Items.Item("txtfaps01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_FacturaAnticipoSeccion02") Then
                    oForm.Items.Item("txtfaps02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_NotaCreditoSeccion01") Then
                    oForm.Items.Item("txtntcs01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_NotaCreditoSeccion02") Then
                    oForm.Items.Item("txtntcs02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_NotaDebitoSeccion01") Then
                    oForm.Items.Item("txtntds01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_NotaDebitoSeccion02") Then
                    oForm.Items.Item("txtntds02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_CompleExportacion") Then
                    oForm.Items.Item("txtExport").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_CompleReembolso") Then
                    oForm.Items.Item("txtReemb").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_GuiaRemisionSeccion01") Then
                    oForm.Items.Item("txtgres01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_GuiaRemisionSeccion02") Then
                    oForm.Items.Item("txtgres02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_RetencionSeccion01") Then
                    oForm.Items.Item("txtrets01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_RetencionSeccion02") Then
                    oForm.Items.Item("txtrets02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_LiquidacionSeccion01") Then
                    oForm.Items.Item("txtliqs01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_LiquidacionSeccion02") Then
                    oForm.Items.Item("txtliqs02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_DocumentosEnviados") Then
                    oForm.Items.Item("txtDocEnv").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                    'add guias desatendidas
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_GuiasDesatendidas01") Then
                    oForm.Items.Item("txtgrde01").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Query_GuiasDesatendidas02") Then
                    oForm.Items.Item("txtgrde02").Specific.value = Utilitario.Util_Encriptador.Desencriptar(odt.GetValue("U_Valor", i).ToString().Replace("{", "").Replace("}", "").ToString(), sKey)
                End If
                ACTUALIZA = 1
            Next

            If ACTUALIZA = 1 Then
                Dim obtnGrabar As SAPbouiCOM.Button
                obtnGrabar = oForm.Items.Item("obtnGDBGT").Specific
                obtnGrabar.Caption = "Actualizar"
            End If

        Catch ex As Exception
            rsboApp.MessageBox(ex.Message.ToString())
        Finally
            oForm.Freeze(False)
            mors = Nothing
        End Try

    End Sub



    Private Sub CargarQueryBaseOrigen(nombreBD_Origen As String)
        Try
            Dim oConfiguracion As Entidades.Configuracion
            Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

            Utilitario.Util_Log.Escribir_Log("Cargando Querys de la BD Origen Nombre: " + nombreBD_Origen.ToString(), "frmParametrosAddon")

            'OBTENER QUERYS
            Dim sQUERY As String
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE'"
            End If
            Dim txtFE As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE1'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE1'"
            End If
            Dim txtFE1 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE2'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE2'"
            End If
            Dim txtFE2 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE3'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryFE3'"
            End If
            Dim txtFE3 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")


            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ'"
            End If
            Dim txtTQ As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ1'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ1'"
            End If
            Dim txtTQ1 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ2'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ2'"
            End If
            Dim txtTQ2 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + " .""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ3'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryTQ3'"
            End If
            Dim txtTQ3 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")


            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC'"
            End If
            Dim txtNC As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC1'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC1'"
            End If
            Dim txtNC1 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC2'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC2'"
            End If
            Dim txtNC2 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC3'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryNC3'"
            End If
            Dim txtNC3 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")


            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND'"
            End If
            Dim txtND As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND1'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND1'"
            End If
            Dim txtND1 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND2'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND2'"
            End If
            Dim txtND2 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")

            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + ".""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND3'"
            Else
                sQUERY = "SELECT ""U_Valor"" FROM " + nombreBD_Origen + "..""@GS_CONFD"" WHERE ""U_Nombre"" = 'QueryND3'"
            End If
            Dim txtND3 As String = oFuncionesB1.getRSvalue(sQUERY, "U_Valor", "")


            'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
            oConfiguracion = New Entidades.Configuracion
            oConfiguracion.Modulo = "eDoc"
            oConfiguracion.Tipo = "PARAMETROS"
            oConfiguracion.SubTipo = "BD"
            olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryFE", txtFE.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryFE1", txtFE1.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryFE2", txtFE2.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryFE3", txtFE3.ToString()))

            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryTQ", txtTQ.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryTQ1", txtTQ1.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryTQ2", txtTQ2.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryTQ3", txtTQ3.ToString()))

            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryNC", txtNC.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryNC1", txtNC1.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryNC2", txtNC2.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryNC3", txtNC3.ToString()))

            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryND", txtND.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryND1", txtND1.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryND2", txtND2.ToString()))
            olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("QueryND3", txtND3.ToString()))

            oConfiguracion.Detalle = olistaDetalleConfiguracion
            GuardaCONF(oConfiguracion)
        Catch ex As Exception

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
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@GS_CONF"" Where ""U_Modulo"" = '" + oConfiguracion.Modulo + "' AND ""U_Tipo"" = '" + oConfiguracion.Tipo + "' AND ""U_Subtipo"" = '" + oConfiguracion.SubTipo + "'"
            Else
                query = "Select DocEntry From [@GS_CONF] Where U_Modulo = '" + oConfiguracion.Modulo + "' AND U_Tipo = '" + oConfiguracion.Tipo + "' AND U_Subtipo = '" + oConfiguracion.SubTipo + "'"
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            'mRst = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CodeExist = "0" Then ' SI EXISTE, ELIMINO Y ACTUALIZO

                ' SI EXISTE ELIMINA PARA VOLVER A CREAR
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)
                oGeneralService.Delete(oGeneralParams)

                'CREA NUEVAMENTE EL REGISTRO
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("GS_CONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)

            Else

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("GS_CONFD")
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

            valor = oFuncionesB1.getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
        End Try
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
                    ' oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Consultas Locales y WEb"


    Function customCertValidation(ByVal sender As Object,
                                   ByVal cert As X509Certificate,
                                   ByVal chain As X509Chain,
                                   ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function


#End Region




    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        Try
            If pVal.FormTypeEx = "frmConsultasDB" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "obtnGDBGT"

                                    Try


                                        Dim oConfiguracion As Entidades.Configuracion
                                        Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)



                                        'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
                                        oConfiguracion = New Entidades.Configuracion
                                        oConfiguracion.Modulo = NombreAddon
                                        oConfiguracion.Tipo = "PARAMETROS"
                                        oConfiguracion.SubTipo = "BD"
                                        olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_FacturaSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtfacs01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_FacturaSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtfacs02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_FacturaAnticipoSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtfaps01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_FacturaAnticipoSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtfaps02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_NotaCreditoSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtntcs01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_NotaCreditoSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtntcs02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_NotaDebitoSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtntds01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_NotaDebitoSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtntds02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_CompleExportacion", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtExport").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_CompleReembolso", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtReemb").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_GuiaRemisionSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtgres01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_GuiaRemisionSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtgres02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_RetencionSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtrets01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_RetencionSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtrets02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_LiquidacionSeccion01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtliqs01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_LiquidacionSeccion02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtliqs02").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_DocumentosEnviados", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtDocEnv").Specific.Value.ToString(), sKey)))

                                        'Guias desatendidas
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_GuiasDesatendidas01", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtgrde01").Specific.Value.ToString(), sKey)))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Query_GuiasDesatendidas02", Utilitario.Util_Encriptador.Encriptar(oForm.Items.Item("txtgrde02").Specific.Value.ToString(), sKey)))


                                        oConfiguracion.Detalle = olistaDetalleConfiguracion
                                        GuardaCONF(oConfiguracion)

                                        oForm.Items.Item("obtnGDBGT").Visible = False
                                        oForm.Items.Item("2").Left = oForm.Items.Item("obtnGDBGT").Left
                                        oForm.Items.Item("2").Specific.Caption = "OK"

                                    Catch ex As Exception

                                    End Try



                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmConsultasDBGT")
            System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try


    End Sub
End Class

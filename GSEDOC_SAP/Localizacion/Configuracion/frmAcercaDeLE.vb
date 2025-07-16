Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions
Imports System.Windows.Forms

Public Class frmAcercaDeLE
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable
    Dim _sCardCode As String = ""
    Dim _fila As String
    'Dim _listaDetalleArtiulos As List(Of Entidades.DetalleArticulo)

    Private GetfileThread As Threading.Thread

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioAcercaDe()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmAcercaDeLE") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmAcercaDeLE.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)

            Catch exx As Exception
                rsboApp.Forms.Item("frmAcercaDeLE").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmAcercaDeLE")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = Application.StartupPath & "\LogoSS.png"

            Dim txtRes As SAPbouiCOM.EditText
            txtRes = oForm.Items.Item("txtRes").Specific
            txtRes.Value = "Addon que activa Funcionalidades Para cumplir las normas tributarias del SRI"

            Dim lbVersion As SAPbouiCOM.StaticText
            lbVersion = oForm.Items.Item("lbVersion").Specific
            lbVersion.Caption = "Versión : " + Functions.VariablesGlobales._vgVersionAddOn


            Dim lbCopy As SAPbouiCOM.StaticText
            lbCopy = oForm.Items.Item("lbCopy").Specific
            lbCopy.Caption = "Copyright © " & Now.Date.Year.ToString & " SOLSAP360 S.A."

            Dim lbUrl As SAPbouiCOM.StaticText
            lbUrl = oForm.Items.Item("lbUrl").Specific
            lbUrl.Item.ForeColor = RGB(6, 69, 173)
            lbUrl.Item.TextStyle = 4

            Dim lbNombre As SAPbouiCOM.StaticText
            lbNombre = oForm.Items.Item("lbNombre").Specific
            lbNombre.Item.ForeColor = RGB(6, 69, 173) 'RGB(0, 101, 184)

            Dim lbValido As SAPbouiCOM.StaticText
            lbValido = oForm.Items.Item("lbValido").Specific

            Dim lbLicencia As SAPbouiCOM.StaticText
            lbLicencia = oForm.Items.Item("lbLicencia").Specific

            If Not Functions.VariablesGlobales._vgTieneLicenciaActivaAddOn Then
                lbValido.Caption = "Su licencia esta vencida! Contactese con un asesor de SOLSAP360 S.A."
                'lbValido.Caption = "Valido Hasta : " + FechaD.ToString("MMMM dd, yyyy")
                lbValido.Item.ForeColor = RGB(204, 0, 0)
                lbLicencia.Caption = ""
            Else
                lbValido.Caption = "Su licencia esta Activa! "
                'If Functions.VariablesGlobales._gTipoLicenciaAddOn.ToLower = "full" Then

                'End If

                lbLicencia.Caption = "Su licencia actual le permite Activar la Localizacion EC"

                lbLicencia.Item.ForeColor = RGB(7, 118, 10)
                lbValido.Item.ForeColor = RGB(7, 118, 10)
            End If

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
                 And pVal.FormTypeEx = "frmAcercaDeLE" Then
                If pVal.BeforeAction = False And pVal.ItemUID = "lbUrl" Then
                    Try
                        oForm = rsboApp.Forms.Item("frmAcercaDeLE")
                        Dim lbUrl As SAPbouiCOM.StaticText
                        lbUrl = oForm.Items.Item("lbUrl").Specific
                        System.Diagnostics.Process.Start(lbUrl.Caption.ToString())
                    Catch ex As Exception

                    End Try
                    'ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnLi" Then ' LICENCIA
                    '    ProcesoBtnExaminarPart()

                ElseIf pVal.BeforeAction = False And pVal.ItemUID = "lnkConf" Then ' VALIDAR ESTRUCTURA DE LA BASE DE DATOS




                    If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                        ' ofrmConfMenuLE.CargaFormularioMenuDeConfiguraciones()

                        ofrmConfClave.CargaFormularioValidarClave(True)


                    End If


                ElseIf pVal.BeforeAction = False And pVal.ItemUID = "lnkVLE" Then ' INGRESA AL MENU DE CONFIGURACION
                    rEstructura.CreacionDeEstructura()


                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmAcercaDe")

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

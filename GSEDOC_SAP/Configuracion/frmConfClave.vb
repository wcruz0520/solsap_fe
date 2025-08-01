Public Class frmConfClave

    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Private _EsAddonLocalizacion As Boolean = False

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioValidarClave(Optional ByVal EsAddonLocalizacion As Boolean = False)
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConfClave") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmConfClave.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConfClave").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try


            _EsAddonLocalizacion = EsAddonLocalizacion

            oForm = rsboApp.Forms.Item("frmConfClave")
            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
      
        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                   And pVal.FormTypeEx = "frmConfClave" Then
            If pVal.BeforeAction = False And pVal.ItemUID = "btnIngre" Then

                oForm = rsboApp.Forms.Item("frmConfClave")

                Dim txtClave As SAPbouiCOM.EditText
                txtClave = oForm.Items.Item("txtClave").Specific
                If txtClave.Value.ToString().Equals("S0ls@p2o1f") Then

                    If _EsAddonLocalizacion Then
                        ofrmConfMenuLE.CargaFormularioMenuDeConfiguraciones()
                    Else
                        ofrmConfMenu.CargaFormularioMenuDeConfiguraciones()

                    End If


                    oForm.Close()
                Else
                    rsboApp.StatusBar.SetText(NombreAddon + " - Clave Incorrecta!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

            End If
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
End Class

Imports System.IO

Public Class Eventos
    Public WithEvents rSboApp As SAPbouiCOM.Application

    Sub New()
        rSboApp = rSboGui.GetApplication
    End Sub

    ''' <summary>
    '''  Eventos de Aplicacion
    ''' </summary>
    ''' <param name="EventType"></param>
    ''' <remarks></remarks>
    Private Sub rSboApp_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles rSboApp.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    System.Windows.Forms.Application.Exit()
                    End
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    System.Windows.Forms.Application.Exit()
                    End
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    System.Windows.Forms.Application.Exit()
                    End
                Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                    System.Windows.Forms.Application.Exit()
                    End
            End Select

        Catch ex As Exception
        End Try
    End Sub


    Private Sub rSboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.MenuEvent
        Try
            '1284
            If pVal.MenuUID = "GS30" And pVal.BeforeAction = False Then
                ' Acerca de
                ofrmAcercaDe.CargaFormularioAcercaDe()
            End If

            'If pVal.MenuUID = "1282" Or pVal.MenuUID = "1287" And pVal.BeforeAction = False Then ' NUEVO, DUPLICAR
            '    Try
            '        Dim typeExx, idFormm As String
            '        typeExx = oFuncionesB1.FormularioActivo(idFormm)
            '        If typeExx = "141" Then ' FACTURA DE PROEVEEDORES
            '            If Not pVal.BeforeAction Then
            '                Try
            '                    Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idFormm)
            '                    mForm.Items.Item("txtFE").Visible = False
            '                    mForm.Items.Item("LinkFE").Visible = False
            '                Catch ex As Exception
            '                End Try
            '            End If

            '        End If
            '    Catch ex As Exception
            '    End Try
            'End If

        Catch ex As Exception
            rSboApp.MessageBox(ex.Message)
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "eventos")
        End Try
    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent
        'Try
        '    If pVal.FormTypeEx = "141" Then ' FACTURA DE PROEVEEDORES
        '        Select Case pVal.EventType
        '            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
        '                If pVal.Before_Action Then
        '                    Select Case pVal.ItemUID
        '                        Case "LinkFE"
        '                            ConsutarPDFRecibido(IdDocumento, 1)
        '                            BubbleEvent = False

        '                    End Select
        '                End If
        '        End Select

        '    End If
        'Catch ex As Exception
        '    rSboApp.MessageBox(ex.Message)
        'End Try
    End Sub

    Private Sub rSboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rSboApp.FormDataEvent
        'Try
        '    If BusinessObjectInfo.FormTypeEx = "141" Then ' FACTURA DE PROEVEEDORES
        '        Select Case BusinessObjectInfo.EventType
        '            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
        '                If BusinessObjectInfo.BeforeAction = False Then
        '                    Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '                    oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

        '                    If oDocumento.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
        '                        AddItemBtnLink(mForm)
        '                        IdDocumento = oDocumento.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
        '                        ClaveAcceso = oDocumento.UserFields.Fields.Item("U_SSCLAVE").Value.ToString()
        '                    Else
        '                        Try
        '                            mForm.Items.Item("txtFE").Visible = False
        '                            mForm.Items.Item("LinkFE").Visible = False
        '                        Catch ex As Exception
        '                        End Try
        '                    End If
        '                End If

        '        End Select


        '    End If
        'Catch ex As Exception
        '    rSboApp.MessageBox(ex.Message)
        'End Try
    End Sub

End Class

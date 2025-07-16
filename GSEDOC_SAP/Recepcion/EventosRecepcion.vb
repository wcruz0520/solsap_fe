Imports System.IO

Public Class EventosRecepcion
    Public WithEvents rSboApp As SAPbouiCOM.Application
    Dim oDocumento As SAPbobsCOM.Documents
    Dim oDocumentoPagoRecibido As SAPbobsCOM.Payments
    Dim oUdoCompRetVentas As SAPbobsCOM.IUserObjectsMD
    Dim oLink As SAPbouiCOM.LinkedButton
    Dim IdDocumento As String
    Dim ClaveAcceso As String
    Dim _WS_RecepcionClave As String = ""
    Dim _WS_RecepcionArchivo As String = ""
    Dim mensaje As String

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential

    Sub New()
        rSboApp = rSboGui.GetApplication
    End Sub

    Private Sub rSboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.MenuEvent
        Try
            If pVal.MenuUID = "GS21" And pVal.BeforeAction = False Then
                Utilitario.Util_Log.Escribir_Log("Recepcion heison activado: " & Functions.VariablesGlobales._XMLRecepcionHeison.ToString(), "EventoRecepcion")
                If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                    ofrmDocumentosRecibidosXML.CreaFormularioDocumentosRecibidos()
                    Utilitario.Util_Log.Escribir_Log("Ingreso a funcion CreaFormularioDocumentosRecibidos de la clase ofrmDocumentosRecibidosXML", "EventoRecepcion")
                Else
                    ofrmDocumentosRecibidos.CreaFormularioDocumentosRecibidos()
                End If


            End If
            If pVal.MenuUID = "GS22" And pVal.BeforeAction = False Then
                ' Documentos Integrados
                ofrmDocumentosIntegrados.CreaFormularioDocumentosIntegrados()
            End If
            If pVal.MenuUID = "GS23" And pVal.BeforeAction = False Then
                ' Parametrización
                ofrmParametrosRecepcion.CargaFormularioParametrosRecepcion()
            End If
            If pVal.MenuUID = "GS24" And pVal.BeforeAction = False Then
                ' Parametrización
                'ofrmSubirArchivo.CargaFormularioSubirArchivo()
                Utilitario.Util_Log.Escribir_Log("Preliminar lote xml: " & Functions.VariablesGlobales._PreliminarLoteXML.ToString(), "EventoRecepcion")

                If Functions.VariablesGlobales._vgProcesoLoteManamer = "Y" Then
                    ofrmProcesoLoteManamer.CreaFormularioProcesoLoteManamer()
                ElseIf Functions.VariablesGlobales._PreliminarLoteXML = "Y" Then

                    ofrmProcesoLoteXML.CreaFormularioProcesoLote()
                Else
                    'ofrmProcesoLote.CreaFormularioProcesoLote()
                    ofrmProcesoLoteC.CreaFormularioProcesoLote()
                End If

                'ofrmProcesoLote2.CreaFormularioPL2()


            End If

            If pVal.MenuUID = "1282" Or pVal.MenuUID = "1287" And pVal.BeforeAction = False Then ' NUEVO, DUPLICAR
                Try
                    Dim typeExx, idFormm As String
#Disable Warning BC42030 ' La variable 'idFormm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    typeExx = oFuncionesB1.FormularioActivo(idFormm)
#Enable Warning BC42030 ' La variable 'idFormm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    If typeExx = "141" _
                        Or typeExx = "181" Then ' FACTURA DE PROEVEEDORES, NOTA DE CREDITO DE PROVEEDORES
                        If Not pVal.BeforeAction Then
                            Try
                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idFormm)
                                mForm.Items.Item("txtFE").Visible = False
                                mForm.Items.Item("LinkFE").Visible = False
                            Catch ex As Exception
                            End Try
                        End If

                    End If
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception
            rSboApp.MessageBox(ex.Message)
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "EventosRecepcion")

        End Try
    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "141" Or pVal.FormTypeEx = "60092" Then ' FACTURA DE PROEVEEDORES / FACTURA RESERVA DE PROVEEDORES
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "LinkFE"
                                    'ConsutarPDFRecibido(IdDocumento, 1)
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoXML.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    Else
                                        ofrmDocumento.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    End If


                            End Select
                        End If
                End Select
            ElseIf pVal.FormTypeEx = "181" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "LinkFE"
                                    'ConsutarPDFRecibido(IdDocumento, 1)
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoNCXML.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    Else
                                        ofrmDocumentoNC.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    End If


                            End Select
                        End If
                End Select
            ElseIf pVal.FormTypeEx = "146" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "LinkFE"
                                    'ConsutarPDFRecibido(IdDocumento, 1)
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoREXML.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    Else
                                        ofrmDocumentoRE.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    End If

                            End Select
                        End If
                End Select
            ElseIf pVal.FormTypeEx = "170" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "LinkFE"
                                    'ConsutarPDFRecibido(IdDocumento, 1)
                                    If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then
                                        ofrmDocumentoREXML.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    Else
                                        ofrmDocumentoRE.CargaFormularioDocumentoExistente(IdDocumento, "docFinal")
                                        BubbleEvent = False
                                    End If
                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception
            rSboApp.MessageBox(ex.Message)
        End Try
    End Sub

    Private Sub rSboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rSboApp.FormDataEvent
        Try
            If BusinessObjectInfo.FormTypeEx = "141" Or BusinessObjectInfo.FormTypeEx = "60092" Then ' FACTURA DE PROEVEEDORES / FACTURA RESERVA DE PROVEEDORES
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If BusinessObjectInfo.BeforeAction = False Then
                            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
                            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                            If BusinessObjectInfo.FormTypeEx = "60092" Then
                                oDocumento.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tYES
                            End If
                            oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                            If oDocumento.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" _
                                And oDocumento.CancelStatus = SAPbobsCOM.CancelStatusEnum.csNo Then
                                AddItemBtnLink(mForm)
                                IdDocumento = oDocumento.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                            Else
                                Try
                                    mForm.Items.Item("txtFE").Visible = False
                                    mForm.Items.Item("LinkFE").Visible = False
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                End Select
            ElseIf BusinessObjectInfo.FormTypeEx = "181" Then ' NOTA DE CREDITO DE PROVEEDORES
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If BusinessObjectInfo.BeforeAction = False Then
                            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
                            oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                            If oDocumento.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" _
                                And oDocumento.CancelStatus = SAPbobsCOM.CancelStatusEnum.csNo Then
                                AddItemBtnLink(mForm)
                                IdDocumento = oDocumento.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                            Else
                                Try
                                    mForm.Items.Item("txtFE").Visible = False
                                    mForm.Items.Item("LinkFE").Visible = False
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                End Select
            ElseIf BusinessObjectInfo.FormTypeEx = "146" Then ' RETENCION DE CLIENTES
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If BusinessObjectInfo.BeforeAction = False Then
                            ' PARA EXXIS Y ONE SOLUTIONS LA RETENCION SE MANEJA COMO TARJETA DE CREDITO FORMULARIO : 146
                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS _
                                  Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS _
                                  Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                                  Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then 'agregar seidor (syp)

                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
                                oDocumentoPagoRecibido = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                                oDocumentoPagoRecibido.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                Dim count As Integer = 0
                                For count = 0 To oDocumentoPagoRecibido.CreditCards.Count - 1
                                    oDocumentoPagoRecibido.CreditCards.SetCurrentLine(count)
                                Next
                                'Dim SSCREADAR As String = mForm.DataSources.DBDataSources.Item("RCT3").GetValue("U_SSCREADAR", 0)
                                Dim SSCREADAR As String = oDocumentoPagoRecibido.CreditCards.UserFields.Fields.Item("U_SSCREADAR").Value.ToString()
                                If LTrim(RTrim(SSCREADAR)) = "SI" Then
                                    'And oDocumentoPagoRecibido.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                    AddItemBtnLinkRetencion(mForm)
                                    IdDocumento = oDocumentoPagoRecibido.CreditCards.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                Else
                                    Try
                                        mForm.Items.Item("txtFE").Visible = False
                                        mForm.Items.Item("LinkFE").Visible = False
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If

                        End If
                End Select

            ElseIf BusinessObjectInfo.FormTypeEx = "170" Then ' PAGO RECIBIDO
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If BusinessObjectInfo.BeforeAction = False Then
                            ' PARA SYPSOFT LA RETENCION SE MANEJA COMO PAGO RECIBIDO, EFECTIVO Y TRANSFERENCIA, POR ESO SE CARGA EL LINK DIRECTO
                            ' EN LA PANTALLA DE PAGO RECIBIDO FORMULARIO: 170
                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then

                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
                                oDocumentoPagoRecibido = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                                oDocumentoPagoRecibido.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                                If oDocumentoPagoRecibido.UserFields.Fields.Item("U_SSCREADAR").Value = "SI" Then
                                    ' AddItemBtnLink(mForm)
                                    AddItemBtnLinkRetencion_PAGO(mForm)
                                    IdDocumento = oDocumentoPagoRecibido.UserFields.Fields.Item("U_SSIDDOCUMENTO").Value.ToString()
                                Else
                                    Try
                                        mForm.Items.Item("txtFE").Visible = False
                                        mForm.Items.Item("LinkFE").Visible = False
                                    Catch ex As Exception
                                    End Try
                                End If


                            End If
                        End If
                End Select
                'ElseIf BusinessObjectInfo.FormTypeEx = "UDO_FT_TM_RETV" Then 'COMPROBANTE DE RETENCION VENTAS TOPMANAGE
                '    Select Case BusinessObjectInfo.EventType
                '        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                '            If BusinessObjectInfo.BeforeAction = False Then

                '                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
                '                Dim CreadoPorAddon As String = mForm.DataSources.DBDataSources.Item("@TM_LE_RETVH").GetValue("U_SSCREADAR", 0).ToString.Trim()
                '                Dim estado As String = mForm.DataSources.DBDataSources.Item("@TM_LE_RETVH").GetValue("U_TM_STATUS", 0).ToString.Trim()
                '                'Dim CreadoPorAddon As String = mForm.Items.Item("U_SSCREADAR").Specific.value.ToString()
                '                'mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_ESTADO_AUTORIZACIO", 0).ToString.Trim()
                '                If CreadoPorAddon = "SI" Then

                '                Else
                '                    Try
                '                        '  mForm.Items.Item("txtFE").Visible = False
                '                        ' mForm.Items.Item("LinkFE").Visible = False
                '                    Catch ex As Exception
                '                    End Try
                '                End If
                '            End If
                '    End Select
            End If
        Catch ex As Exception
            rSboApp.MessageBox(ex.Message)
        End Try
    End Sub

    Public Sub AddItemBtnLink(oForm As SAPbouiCOM.Form)
        Try
            Dim oItem As SAPbouiCOM.Item

            ' BP Code link button
            Dim oLink As SAPbouiCOM.LinkedButton = Nothing
            Dim otxt As SAPbouiCOM.StaticText = Nothing

            'For Each oItemlist As SAPbouiCOM.Item In oForm.Items
            '    If oItemlist.UniqueID = "txtFE" Then
            '        Return
            '    End If
            'Next
            ' oForm.Freeze(false);
            Try
                oItem = oForm.Items.Add("txtFE", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Left = 319
                oItem.Width = 200
                oItem.Top = 113
                oItem.Height = 14
                oItem.Visible = True
                oItem.ForeColor = RGB(7, 118, 10)
                otxt = DirectCast(oItem.Specific, SAPbouiCOM.StaticText)
                otxt.Caption = "Ver Documento Electrónico Recibido"
            Catch ex As Exception
                oForm.Items.Item("txtFE").Visible = True

            End Try
           
            Try
                oItem = oForm.Items.Add("LinkFE", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                oItem.Left = 304
                oItem.Width = 12
                oItem.Top = 115
                oItem.Height = 10
                oItem.Visible = True
                ' Link the column to the BP master data system form
                oLink = DirectCast(oItem.Specific, SAPbouiCOM.LinkedButton)
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Order
                oLink.LinkedObjectType = 2
                oLink.Item.LinkTo = "txtFE"
            Catch ex As Exception
                oForm.Items.Item("LinkFE").Visible = True
            End Try
           

            '    oForm.Freeze(true);
        Catch ex As Exception

        End Try

    End Sub

    Public Sub AddItemBtnLinkRetencion(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oItem As SAPbouiCOM.Item

            ' BP Code link button
            Dim oLink As SAPbouiCOM.LinkedButton = Nothing
            Dim otxt As SAPbouiCOM.StaticText = Nothing

            'For Each oItemlist As SAPbouiCOM.Item In oForm.Items
            '    If oItemlist.UniqueID = "txtFE" Then
            '        Return
            '    End If
            'Next
            ' oForm.Freeze(false);
            Try
                oItem = oForm.Items.Add("txtFE", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Left = 319
                oItem.Width = 200
                oItem.Top = 200 '113
                oItem.Height = 14
                oItem.Visible = True
                oItem.ForeColor = RGB(7, 118, 10)
                otxt = DirectCast(oItem.Specific, SAPbouiCOM.StaticText)
                otxt.Caption = "Ver Documento Electrónico Recibido"
            Catch ex As Exception
                oForm.Items.Item("txtFE").Visible = True

            End Try

            Try
                oItem = oForm.Items.Add("LinkFE", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                oItem.Left = 304
                oItem.Width = 12
                oItem.Top = 202
                oItem.Height = 10
                oItem.Visible = True
                ' Link the column to the BP master data system form
                oLink = DirectCast(oItem.Specific, SAPbouiCOM.LinkedButton)
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Order
                oLink.LinkedObjectType = 2
                oLink.Item.LinkTo = "txtFE"
            Catch ex As Exception
                oForm.Items.Item("LinkFE").Visible = True
            End Try


            '    oForm.Freeze(true);
        Catch ex As Exception

        End Try

    End Sub

    Public Sub AddItemBtnLinkRetencion_PAGO(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oItem As SAPbouiCOM.Item

            ' BP Code link button
            Dim oLink As SAPbouiCOM.LinkedButton = Nothing
            Dim otxt As SAPbouiCOM.StaticText = Nothing

            'For Each oItemlist As SAPbouiCOM.Item In oForm.Items
            '    If oItemlist.UniqueID = "txtFE" Then
            '        Return
            '    End If
            'Next
            ' oForm.Freeze(false);
            Try
                oItem = oForm.Items.Add("txtFE", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Left = 535
                oItem.Width = 200
                oItem.Top = 450 '113
                oItem.Height = 14
                oItem.Visible = True
                oItem.ForeColor = RGB(7, 118, 10)
                otxt = DirectCast(oItem.Specific, SAPbouiCOM.StaticText)
                otxt.Caption = "Ver Documento Electrónico Recibido"
            Catch ex As Exception
                oForm.Items.Item("txtFE").Visible = True

            End Try

            Try
                oItem = oForm.Items.Add("LinkFE", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                oItem.Left = 522
                oItem.Width = 12
                oItem.Top = 452
                oItem.Height = 10
                oItem.Visible = True
                ' Link the column to the BP master data system form
                oLink = DirectCast(oItem.Specific, SAPbouiCOM.LinkedButton)
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Order
                oLink.LinkedObjectType = 2
                oLink.Item.LinkTo = "txtFE"
            Catch ex As Exception
                oForm.Items.Item("LinkFE").Visible = True
            End Try


            '    oForm.Freeze(true);
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ConsutarPDFRecibido(ByVal IdDocumento As Integer, ByVal TipoDocumento As Integer)
        Try

            '_WS_RecepcionArchivo = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsultaArchivo")
            _WS_RecepcionArchivo = Functions.VariablesGlobales._WS_RecepcionConsultaArchivo
            If _WS_RecepcionArchivo = "" Then
                rSboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End If
            '_WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")
            _WS_RecepcionClave = Functions.VariablesGlobales._WS_RecepcionClave
            rSboApp.StatusBar.SetText(NombreAddon + " - Ruta Recepcion: " + _WS_RecepcionArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            rSboApp.StatusBar.SetText(NombreAddon + " - Clave Recepcion: " + _WS_RecepcionClave, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO
            WS.Url = _WS_RecepcionArchivo
            'MANEJO DE PROXY
            Dim SALIDA_POR_PROXY As String = ""
            'SALIDA_POR_PROXY = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY")
            SALIDA_POR_PROXY = Functions.VariablesGlobales._SALIDA_POR_PROXY
            Utilitario.Util_Log.Escribir_Log("SALIDA POR PROXY : " + SALIDA_POR_PROXY.ToString, "ManejoDeDocumentos")
            Dim Proxy_puerto As String = ""
            Dim Proxy_IP As String = ""
            Dim Proxy_Usuario As String = ""
            Dim Proxy_Clave As String = ""
            If SALIDA_POR_PROXY = "Y" Then

                Proxy_puerto = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_PUERTO")
                Proxy_IP = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_IP")
                Proxy_Usuario = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_USER")
                Proxy_Clave = ConsultaParametro("SAED", "PARAMETROS", "PROXY", "PROXY_CLAVE")

                Utilitario.Util_Log.Escribir_Log("Proxy_puerto : " + Proxy_puerto.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_IP : " + Proxy_IP.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Usuario : " + Proxy_Usuario.ToString, "ManejoDeDocumentos")
                Utilitario.Util_Log.Escribir_Log("Proxy_Clave : " + Proxy_Clave.ToString, "ManejoDeDocumentos")

                If Not Proxy_puerto = "" Then
                    proxyobject = New System.Net.WebProxy(Proxy_IP, Integer.Parse(Proxy_puerto))
                Else
                    proxyobject = New System.Net.WebProxy(Proxy_IP)
                End If
                cred = New System.Net.NetworkCredential(Proxy_Usuario, Proxy_Clave)

                proxyobject.Credentials = cred

                WS.Proxy = proxyobject
                WS.Credentials = cred

            End If
            ' END MANEJO DE PROXY

            Dim filepath As String = Path.GetTempPath()
            filepath += ClaveAcceso + ".pdf"

            ' SI NO EXISTE EN LA CARPETA TEMPORAL, LO CONSULTO AL WS
            If Not File.Exists(filepath) Then
                rSboApp.SetStatusBarMessage(NombreAddon + " - Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim FS As FileStream = Nothing
                'oManejoDocumentos.SetProtocolosdeSeguridad()
                ofrmDocumentosRecibidos.SetProtocolosdeSeguridad()
                Dim dbbyte As Byte() = WS.ConsultaArchivoProveedor_PDF(_WS_RecepcionClave, 1, IdDocumento, mensaje)
                FS = New FileStream(filepath, System.IO.FileMode.Create)
                FS.Write(dbbyte, 0, dbbyte.Length)
                FS.Close()
            End If

            Dim Proc As New Process()
            Proc.StartInfo.FileName = filepath
            Proc.Start()
            Proc.Dispose()

        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Ocurrio un error al generar el PDF recibido! " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Long, False)
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

            valor = oFuncionesAddon.getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    Private Sub rSboApp_UDOEvent(ByRef udoEventArgs As SAPbouiCOM.UDOEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.UDOEvent
        'If udoEventArgs. = "TM_RETV" Then

        'End If
    End Sub
End Class

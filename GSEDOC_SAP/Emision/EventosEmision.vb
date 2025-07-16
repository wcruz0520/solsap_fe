Imports System.IO
Imports System.Drawing.Printing
'https
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports SAPbobsCOM

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Reflection
Imports Microsoft.Office.Interop.Excel

Public Class EventosEmision

    Public WithEvents rSboApp As SAPbouiCOM.Application
    Private oUserFieldsMD As SAPbobsCOM.UserFieldsMD
    'Public _Nombre_Proveedor_SAP_BO As String = ""



#Region "Variables "
    '
    Dim oAuxLocal As SAPbobsCOM.Recordset
    Dim ofils As SAPbouiCOM.EventFilters
    Dim ofil As SAPbouiCOM.EventFilter
    Dim oDBDataSource As SAPbouiCOM.DBDataSource


    Dim formID As String = ""
    Dim oForm As SAPbouiCOM.Form
    Dim item As SAPbouiCOM.Item
    Dim nombreFormulario As String = ""
    Dim oUserDataSourceFC As SAPbouiCOM.UserDataSource
    Dim oUserDataSourceNC As SAPbouiCOM.UserDataSource
    Dim oUserDataSourceND As SAPbouiCOM.UserDataSource
    Dim oUserDataSourceGR As SAPbouiCOM.UserDataSource
    Dim oUserDataSourceRE As SAPbouiCOM.UserDataSource
    Dim oUserDataSourceCE As SAPbouiCOM.UserDataSource
    Dim EsElectronico As String = ""
    Dim LQEsElectronico As String = ""

    Dim lbComentario As SAPbouiCOM.StaticText
    Dim lbEstado As SAPbouiCOM.StaticText
    Dim cbEstado As SAPbouiCOM.ComboBox
    Dim btnAccion As SAPbouiCOM.ButtonCombo


    Dim btnPruebaCombo As SAPbouiCOM.ButtonCombo
    Dim btnPrueba As SAPbouiCOM.Button

    Dim lbComentarioLQ As SAPbouiCOM.StaticText
    Dim lbComentarioGR As SAPbouiCOM.StaticText
    Dim lbEstadoLQ As SAPbouiCOM.StaticText
    Dim lbEstadoGR As SAPbouiCOM.StaticText
    Dim cbEstadoLQ As SAPbouiCOM.ComboBox
    Dim cbEstadoGR As SAPbouiCOM.ComboBox
    Dim btnAccLQ As SAPbouiCOM.ButtonCombo
    Dim btnAccGR As SAPbouiCOM.ButtonCombo
    Dim btnPruebaComboLQ As SAPbouiCOM.ButtonCombo
    Dim btnPruebaLQ As SAPbouiCOM.Button

    Dim sFolio As String = ""
    Dim sCode As String = ""
    Dim numFolio As Integer
    Dim oTabla As String = ""
    Dim oTipoTabla As String = "" ' TIPO DE DOCUMENTO, EJEMPLO SI ES FACTURA O FACTURA DE ANTICIPO NOTA DE DEBITO
    Dim oDocumento As SAPbobsCOM.Documents

    Dim oRetND As SAPbobsCOM.Documents

    Dim oTransferencia As SAPbobsCOM.StockTransfer


    Dim oActualizar As Integer = 0

    Dim CodigoUsuario As String
    Dim IDUsuario As String = 0

    Dim ActualizaSecuenciaFolio As Boolean = False

    Dim SeriesElectronicasUDF As String = "N"

    'Dim FechaEmisionDoc As String = ""
    Dim FechaEmisionDoc As Date

    Dim tipoDoc2 As String = ""

    Dim truco As Boolean = False


#End Region

#Region "Eventos SAP BO"

    Public Sub New()
        Try
            rSboApp = rSboGui.GetApplication
            oFuncionesB1 = New Functions.FuncionesB1(rCompany, rSboApp, True, False, NombreAddon)
            oFuncionesAddon = New Functions.FuncionesAddon(rCompany, rSboApp, True, False, NombreAddon)
            oFuncionesAddon.NombreAddon = NombreAddon

            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                oManejoDocumentosEcua = New Negocio.ManejoDeDocumentosEcua(rCompany, rSboApp, "A", Nombre_Proveedor_SAP_BO)
            ElseIf Functions.VariablesGlobales._ActApiSS = "Y" Then
                oManejoDocumentosSolsap = New Negocio.ManejoDeDocumentoSolsap(rCompany, rSboApp, "A", Nombre_Proveedor_SAP_BO)
            Else
                oManejoDocumentos = New Negocio.ManejoDeDocumentos(rCompany, rSboApp, "A", Nombre_Proveedor_SAP_BO)
            End If

            ofrmDocumentosEnviados = New frmDocumentosEnviados(rCompany, rSboApp)
            If Functions.VariablesGlobales._vgImpBlo = "Y" Then
                ofrmImpresionPorBloque = New frmImpresionPorBloque(rCompany, rSboApp)
            End If

            'ofrmDocumentosRecibidos = New frmDocumentosRecibidos(rCompany, rSboApp)
            'ofrmMapeo = New frmMapeo(rCompany, rSboApp)
            'ofrmDocumento = New frmDocumento(rCompany, rSboApp)
            'ofrmConsultaOrdenes = New frmConsultaOrdenes(rCompany, rSboApp)
            'ofrmDocumentosIntegrados = New frmDocumentosIntegrados(rCompany, rSboApp)

        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    '''  Eventos de Menu
    ''' </summary>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    ''' <remarks></remarks>
    ''' 


    Private Sub Eventos_Menu_Localizacion(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try
            '1284
            If pVal.MenuUID = "GSLoc_1" And pVal.BeforeAction = False Then
                '-------menu Acerca De...
                ofrmAcercaDeLE.CargaFormularioAcercaDe()

            ElseIf pVal.MenuUID = "GSLoc5" And pVal.BeforeAction = False Then

                ofrmSRI.CargaFormularioSRI()

            ElseIf pVal.MenuUID = "GSLoc30" And pVal.BeforeAction = False Then

                ofrmAnexoCompras.CargaFormularioAnexoCompras()

            ElseIf pVal.MenuUID = "GSLoc40" And pVal.BeforeAction = False Then

                ofrmAnexoVentas.CargaFormularioAnexoVentas()

            ElseIf pVal.MenuUID = "GSLoc10_1" And pVal.BeforeAction = False Then
                ofrmGeneradorATS.CargaFormularioGeneradorATS()

            ElseIf pVal.MenuUID = "GSLoc20" And pVal.BeforeAction = False Then
                '-------menu Informes Legales

                ofrmGenerarRPT.CargaFormularioGenerarRPT()

                'Informes
                'Informe facturas de compras
            ElseIf pVal.MenuUID = "GSLoc20_2" And pVal.BeforeAction = False Then

                Dim uuidMenu As String = ObtenerUIDMenu("Informe facturas de compras")

                If Not String.IsNullOrWhiteSpace(uuidMenu) Then

                    rSboApp.ActivateMenuItem(uuidMenu)

                End If

                'Informe para Formulario SRI 103
            ElseIf pVal.MenuUID = "GSLoc20_3" And pVal.BeforeAction = False Then

                Dim uuidMenu As String = ObtenerUIDMenu("Informe para Formulario SRI 103")

                If Not String.IsNullOrWhiteSpace(uuidMenu) Then

                    rSboApp.ActivateMenuItem(uuidMenu)

                End If

                'Informe Listado de Ventas
            ElseIf pVal.MenuUID = "GSLoc20_4" And pVal.BeforeAction = False Then

                Dim uuidMenu As String = ObtenerUIDMenu("Informe Listado de Ventas")

                If Not String.IsNullOrWhiteSpace(uuidMenu) Then

                    rSboApp.ActivateMenuItem(uuidMenu)

                End If

                'Informe Retenciones de clientes
            ElseIf pVal.MenuUID = "GSLoc20_5" And pVal.BeforeAction = False Then

                Dim uuidMenu As String = ObtenerUIDMenu("Informe Retenciones de clientes")

                If Not String.IsNullOrWhiteSpace(uuidMenu) Then

                    rSboApp.ActivateMenuItem(uuidMenu)

                End If

                'Informe 104
            ElseIf pVal.MenuUID = "GSLoc20_6" And pVal.BeforeAction = False Then

                Dim uuidMenu As String = ObtenerUIDMenu("Informe 104")

                If Not String.IsNullOrWhiteSpace(uuidMenu) Then

                    rSboApp.ActivateMenuItem(uuidMenu)

                End If

                'Informe DINARDAP
            ElseIf pVal.MenuUID = "GSLoc20_7" And pVal.BeforeAction = False Then

                'Dim uuidMenu As String = ObtenerUIDMenu("Informe DINARDAP")

                'If Not String.IsNullOrWhiteSpace(uuidMenu) Then

                '    rSboApp.ActivateMenuItem(uuidMenu)

                'End If
                ofrmDinardap.CargaFormularioDinardap()

            ElseIf pVal.MenuUID = "GSLoc50_1" And pVal.BeforeAction = False Then

                ofrmChequeP.CreaFormulario_frmChequeP()

            ElseIf pVal.MenuUID = "GSLoc60_1" And pVal.BeforeAction = False Then

                ofrmTransEntreCompanias.CreaFormulario_frmTransEntreCompanias()

            ElseIf pVal.MenuUID = "GSLoc60_2" And pVal.BeforeAction = False Then

                ofrmConsultaSalidaEntrada.Carga_Salidas_Entrada()

                'Nuevos Menus de Localizacion

                'pagos masivos
            ElseIf pVal.MenuUID = "GSLoc70_1" And pVal.BeforeAction = False Then

                ofrmPagosMasivos.CargaFormularioPagosMasivos()

                ' Aprobacion pagos Masivos

            ElseIf pVal.MenuUID = "GSLoc70_2" And pVal.BeforeAction = False Then

                ofrmPagosAprobacion.CargaFormularioPagosAprobacion()

                'Guias Desatendidas
            ElseIf pVal.MenuUID = "GSLoc80_1" And pVal.BeforeAction = False Then

                ofrmGuiasRemision.CargaFormularioGuia()

                'Add 03072024
            ElseIf pVal.MenuUID = "GSLoc90_1" And pVal.BeforeAction = False Then

                ofrmMapeoCuentas.CargaFormularioConfiguracionesCM()

                'Add 02072024
            ElseIf pVal.MenuUID = "GSLoc90_2" And pVal.BeforeAction = False Then

                ofrmCashManagement.CargaFormularioCashManagement()

                'Add 10/09/2024
            ElseIf pVal.MenuUID = "GSLoc_2" And pVal.BeforeAction = False Then

                ofrmServiciosBasicos.CargaFormularioServiciosBasicos()

            End If

            If pVal.MenuUID = "1287" Then '  DUPLICAR
                If pVal.BeforeAction = False Then
                    Try
                        Dim typeEx As String
                        Dim idForm As String = ""
                        Dim countx As Integer = 0
                        typeEx = oFuncionesB1.FormularioActivo(idForm, countx)

                        If typeEx = "133" Or
                            typeEx = "60090" Or
                            typeEx = "60091" Or
                            typeEx = "60092" Or
                            typeEx = "65303" Or
                            typeEx = "65300" Or
                            typeEx = "65301" Or
                            typeEx = "179" Or
                            typeEx = "140" Or
                            typeEx = "940" Or
                            typeEx = "141" Then

                            'Dim formprincipal = typeEx.Replace("-", "")

                            SetearCamposAutoFE_LOC(rSboApp.Forms.Item(idForm), "DUPLICAR", typeEx)

                        End If

                    Catch ex As Exception

                        rSboApp.SetStatusBarMessage(ex.Message)

                    End Try
                End If

            End If




            If pVal.MenuUID = "1282" Then '  NUEVO
                If pVal.BeforeAction = False Then
                    Try
                        Dim typeEx
                        Dim idForm As String = ""
                        typeEx = oFuncionesB1.FormularioActivo(idForm)
                        If typeEx = "60090" Or
                           typeEx = "133" Or
                           typeEx = "60091" Or
                            typeEx = "60092" Or
                            typeEx = "65303" Or
                            typeEx = "65300" Or
                           typeEx = "65301" Or
                            typeEx = "179" Or
                           typeEx = "140" Or
                           typeEx = "940" Or
                           typeEx = "141" Then

                            SetearCamposAutoFE_LOC(rSboApp.Forms.Item(idForm), "NUEVO", typeEx)

                        End If

                    Catch ex As Exception

                        rSboApp.SetStatusBarMessage(ex.Message)

                    End Try
                End If

            End If



        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error en el control de MenuEvent " & ex.Message, "EventosLE")
        End Try

    End Sub

    Private Sub Eventos_Menu_CM(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            'Add 03072024
            If pVal.MenuUID = "GSLoc90_1" And pVal.BeforeAction = False Then

                ofrmMapeoCuentas.CargaFormularioConfiguracionesCM()

                'Add 02072024
            ElseIf pVal.MenuUID = "GSLoc90_2" And pVal.BeforeAction = False Then

                ofrmCashManagement.CargaFormularioCashManagement()
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error en el control de Eventos_Menu_CM " & ex.Message, "EventosLE")
        End Try
    End Sub
    Private Function ObtenerUIDMenu(ByVal nombreBuscado As String) As String

        Try

            'el codigo 30338 es Informes electrónicos
            For i As Integer = 0 To rSboApp.Menus.Item("30338").SubMenus.Count - 1


                If rSboApp.Menus.Item("30338").SubMenus.Item(i).String.ToLower.Contains(nombreBuscado.ToLower) Then Return rSboApp.Menus.Item("30338").SubMenus.Item(i).UID


            Next



        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar funcion ObtenerUIDMenu " & ex.Message, "EventosLE")
        End Try

        Return ""

    End Function
    Private Sub rSboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.MenuEvent
        Try
            'If pVal.MenuUID = "GS01" And pVal.BeforeAction = False Then
            '    CreaFomularioParametrizacionUsuarios()
            'End If
            If pVal.MenuUID = "GS11" And pVal.BeforeAction = False Then
                ofrmDocumentosEnviados.CreaFormularioDocumentosEnviados()
            End If
            If pVal.MenuUID = "GS12" And pVal.BeforeAction = False Then
                ofrmImpresionPorBloque.CreaFormularioImpresionPorLote()

            End If
            If pVal.MenuUID = "SS_LOG" And pVal.BeforeAction = False Then
                Try
                    Dim typeEx As String = "", idForm As String = ""
                    typeEx = oFuncionesB1.FormularioActivo(idForm)
                    Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)

                    SeteaTipoTabla_FormTypeEx(typeEx)

                    Dim lDocEntry As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)
                    Dim lObjType As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("ObjType", 0)
                    Dim lDocSubType As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocSubType", 0)

                    ofrmLogEmision.CargaFormularioLogEmision(lDocEntry, lObjType, lDocSubType)
                Catch ex As Exception

                End Try

            End If
            If pVal.MenuUID = "1282" Then 'NUEVO
                If pVal.BeforeAction = False Then
                    Try
                        Dim typeEx = "", idForm As String = ""
                        Dim xcount = 0
                        typeEx = oFuncionesB1.FormularioActivo(idForm, xcount)
                        If typeEx.ToString.Contains("-") Then
                            rSboApp.Forms.GetForm(typeEx.ToString.Replace("-", ""), xcount).Select()
                            typeEx = oFuncionesB1.FormularioActivo(idForm, xcount)
                        End If


                        If typeEx = "133" Or
                            typeEx = "60090" Or
                            typeEx = "60091" Or
                            typeEx = "60092" Or
                            typeEx = "65303" Or
                            typeEx = "65307" Or
                            typeEx = "65300" Or
                            typeEx = "65301" Or
                            typeEx = "179" Or
                            typeEx = "140" Or
                            typeEx = "940" Or
                            typeEx = "1250000940" Or
                            typeEx = "65306" Or
                            typeEx = "141" Then

                            Dim count As Integer = 0

                            If typeEx = "65307" Then
                                Functions.VariablesGlobales._SS_FacturaExportacion = "SI"
                            Else
                                Functions.VariablesGlobales._SS_FacturaExportacion = "NO"
                            End If
                            SeteaTipoTabla_FormTypeEx(typeEx)

                            Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)

                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                CargaItemEnFormularioEcuanexus(oForm, "NUEVO", typeEx, oTabla)
                            Else
                                CargaItemEnFormulario(oForm, "NUEVO", typeEx, oTabla)
                            End If


                            If typeEx = "141" Or typeEx = "60092" Then
                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                    CargaItemEnFormulario_LiquidacionCompraEcuanexus(oForm, "NUEVO", typeEx, oTabla)
                                Else
                                    CargaItemEnFormulario_LiquidacionCompra(oForm, "NUEVO", typeEx, oTabla)
                                End If
                            End If



                        End If

                    Catch ex As Exception
                        rSboApp.MessageBox(NombreAddon + " Evento NUEVO- " + ex.Message.ToString())
                    End Try
                End If
            End If

            If pVal.MenuUID = "1287" Then '  DUPLICAR
                If pVal.BeforeAction = False Then
                    Try
                        Dim typeEx = "", idForm As String = ""
                        typeEx = oFuncionesB1.FormularioActivo(idForm)
                        If typeEx = "133" Or
                            typeEx = "60090" Or
                            typeEx = "60091" Or
                            typeEx = "60092" Or
                            typeEx = "65303" Or
                             typeEx = "65307" Or
                            typeEx = "65300" Or
                            typeEx = "65301" Or
                            typeEx = "179" Or
                            typeEx = "140" Or
                            typeEx = "940" Or
                            typeEx = "1250000940" Or
                            typeEx = "65306" Or
                            typeEx = "141" Then

                            Utilitario.Util_Log.Escribir_Log("typeEx: " + typeEx.ToString, "EventosEmision")
                            Dim count As Integer = 0
                            Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)
                            'Dim oFrmUser As SAPbouiCOM.Form
                            'Try
                            '    oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                            'Catch ex As Exception
                            '    rSboApp.SendKeys("^+U")
                            '    oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                            'End Try
                            If typeEx = "65307" Then
                                Functions.VariablesGlobales._SS_FacturaExportacion = "SI"
                            Else
                                Functions.VariablesGlobales._SS_FacturaExportacion = "NO"
                            End If
                            Try

                                SetearCamposAutoFE(oForm, "", typeEx)


                            Catch ex As Exception
                                Utilitario.Util_Log.Escribir_Log("Error al limpiar los campos " + ex.Message.ToString, "EventosEmision")
                            End Try
                            Utilitario.Util_Log.Escribir_Log("campos limpios", "EventosEmision")
                            SeteaTipoTabla_FormTypeEx(typeEx)

                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                CargaItemEnFormularioEcuanexus(oForm, "NUEVO", typeEx, oTabla)
                            Else
                                CargaItemEnFormulario(oForm, "NUEVO", typeEx, oTabla)
                            End If


                            If typeEx = "141" Or typeEx = "60092" Then
                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                    CargaItemEnFormulario_LiquidacionCompraEcuanexus(oForm, "NUEVO", typeEx, oTabla)
                                Else
                                    CargaItemEnFormulario_LiquidacionCompra(oForm, "NUEVO", typeEx, oTabla)
                                End If

                            End If




                        End If

                    Catch ex As Exception
                        rSboApp.MessageBox(NombreAddon + " - " + ex.Message.ToString())
                        Utilitario.Util_Log.Escribir_Log("error en evento duplicar: " + ex.Message.ToString, "EventosEmision")
                    End Try
                End If

            End If

            If pVal.MenuUID = "1284" Then '  cancelar documento
                If pVal.BeforeAction = True Then

                    If Functions.VariablesGlobales._AnuDocNumAtCard = "Y" Then
                        Dim resp As Integer = 0
                        resp = rSboApp.MessageBox("DESEA CANCELAR EL DOCUMENTO..?", 1, "SI", "NO")
                        Select Case resp
                            Case 1
                                Dim typeEx = "", idForm As String = ""
                                typeEx = oFuncionesB1.FormularioActivo(idForm)
                                If typeEx = "133" Or
                                    typeEx = "60090" Or
                                    typeEx = "60091" Or
                                    typeEx = "65303" Or
                                    typeEx = "65307" Or
                                    typeEx = "179" Or
                                    typeEx = "140" Or
                                    typeEx = "940" Or
                                    typeEx = "65306" Or
                                    typeEx = "60092" Or
                                    typeEx = "141" Then

                                    Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)

                                    Dim _docentry = LTrim(RTrim(oForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                    Dim objtype = LTrim(RTrim(oForm.DataSources.DBDataSources.Item(oTabla).GetValue("ObjType", 0)))

                                    ProcesoActualizarParaCancelacion(_docentry, objtype)
                                    oForm.Refresh()

                                    'Dim count As Integer = 0
                                    'Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)
                                    'Dim oFrmUser As SAPbouiCOM.Form
                                    'Try
                                    '    oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                                    'Catch ex As Exception
                                    '    rSboApp.SendKeys("^+U")
                                    '    oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                                    'End Try

                                    'Dim est As String = ""
                                    'Dim pemi As String = ""
                                    'Dim folio As String = ""
                                    'Dim numeracion As String = ""
                                    'Dim RetValAnu As Long
                                    'Dim ErrCodeAnu As Long
                                    'Dim ErrMsgAnu As String = ""



                                    'If typeEx = "133" Or typeEx = "60090" Or typeEx = "60091" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    'ElseIf typeEx = "65303" Then
                                    '    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                                    '    oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                                    'ElseIf typeEx = "179" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                    '    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                                    'ElseIf typeEx = "140" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                    '    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes
                                    'ElseIf typeEx = "940" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                                    '    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                                    'ElseIf typeEx = "141" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                    'ElseIf typeEx = "65306" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                    '    oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                                    'ElseIf typeEx = "60092" Then
                                    '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                    '    oDocumento.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tYES
                                    'End If


                                    'Dim _docentry = LTrim(RTrim(oForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                    'Dim objtype = LTrim(RTrim(oForm.DataSources.DBDataSources.Item(oTabla).GetValue("ObjType", 0)))
                                    'oDocumento.GetByKey(_docentry)

                                    'If oDocumento.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then

                                    '    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                    '        est = oDocumento.UserFields.Fields.Item("U_SS_Est").Value
                                    '        pemi = oDocumento.UserFields.Fields.Item("U_SS_Pemi").Value
                                    '        folio = oDocumento.FolioNumber.ToString
                                    '    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                    '        est = oDocumento.UserFields.Fields.Item("U_SER_EST").Value
                                    '        pemi = oDocumento.UserFields.Fields.Item("U_SER_PE").Value
                                    '        folio = oDocumento.FolioNumber.ToString
                                    '    End If
                                    '    If Not folio.Length.Equals("9") Then
                                    '        folio = folio.PadLeft(9, "0")
                                    '    End If

                                    '    If CInt(folio) > 0 Then
                                    '        numeracion = est + "-" + pemi + "-" + folio

                                    '        oDocumento.FolioPrefixString = ""
                                    '        oDocumento.FolioNumber = 0
                                    '        oDocumento.NumAtCard = numeracion.ToString
                                    '        oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = "11"
                                    '        If typeEx = "141" Then
                                    '            oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value = "11"
                                    '        End If


                                    '        RetValAnu = oDocumento.Update()

                                    '        If RetValAnu <> 0 Then
                                    '            rCompany.GetLastError(ErrCodeAnu, ErrMsgAnu)
                                    '            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualzar campo NumAtCard..!! - " + ErrMsgAnu.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    '        Else
                                    '            rSboApp.SetStatusBarMessage(NombreAddon + " - Campo NumAtCard actualizo con éxito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    '        End If
                                    '    End If

                                    'End If

                                End If
                        End Select



                    End If
                End If
            End If

        Catch ex As Exception
            rSboApp.MessageBox(NombreAddon + " - " + ex.Message.ToString())
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "EventosEmision")

        End Try

        If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

            Eventos_Menu_Localizacion(pVal, BubbleEvent)

        End If

        If Functions.VariablesGlobales._ActivarCMFML = "Y" Then Eventos_Menu_CM(pVal, BubbleEvent)

    End Sub

    Private Sub rSboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent
        Try
            ' 133 'FACTURA DE CLIENTES
            ' 60090 'FACTURA DE DEUDOR + PAGO
            ' 60091 'FACTURA DE RESERVA CLIENTES
            ' 60092 'FACTURA DE RESERVA PROVEEDOR/RETENCION
            ' 65303 'NOTA DE DEBITO DE CLIENTES
            ' 65300 'FACTURA DE ANTICIPO DE CLIENTES
            ' 65301 'FACTURA DE ANTICIPO DE PROVEEDORES
            ' 179 'NOTA DE CREDITO DE CLIENTES
            ' 140 'GUIA DE REMISION
            ' 940 ' GUIA DE REMISION TRANSFERENCIA
            ' 141 'FACTURA DE PROVEEDOR/RETENCION

            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                If pVal.FormTypeEx = "134" Then

                    Select Case pVal.ItemUID

                        Case "1"

                            If pVal.BeforeAction Then

                                If Functions.VariablesGlobales._SS_ValidarSocioNegociosUDF = "Y" Then
                                    ValidacionesPrevioCreacionMaestro(rSboApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount), BubbleEvent)
                                End If

                            End If

                    End Select

                End If

            End If

            'SE AÑADE BOTON PARA CARGAR EXTRACTO BANCARIO EN FORMATO XLS 2025-06-09
            If pVal.FormTypeEx = "385" Then
                If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Try

                                If pVal.BeforeAction = False Then
                                    Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                                    creaBotonCargaExtractoBnacario(mForm)
                                End If
                            Catch ex As Exception
                                rSboApp.SetStatusBarMessage("et_FORM_LOAD 385: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            End Try

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            If pVal.ItemUID = "btnCargaEB" And pVal.BeforeAction = True Then

                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                                Dim rutaXLS = ""
                                'Dim selectFileDialog As New SelectFileDialog("C:\", "", "CSV files (*.csv)|*.csv|All files (*.*)|*.*", DialogType.OPEN)
                                Dim selectFileDialog As New SelectFileDialog("C:\", "", "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*", DialogType.OPEN)
                                selectFileDialog.Open()

                                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                                    rutaXLS = selectFileDialog.SelectedFile
                                    Acciones_BotonCargaEB(mForm, rutaXLS) 'BOTON DE EXTRACTO BANCARIO
                                End If

                            End If



                    End Select
                End If
            End If


            Dim typeEx, idForm As String


            If pVal.FormTypeEx = "133" Or
                pVal.FormTypeEx = "60090" Or
                pVal.FormTypeEx = "60091" Or
                pVal.FormTypeEx = "60092" Or
                pVal.FormTypeEx = "65303" Or
                pVal.FormTypeEx = "65307" Or
                pVal.FormTypeEx = "65300" Or
                pVal.FormTypeEx = "65301" Or
                pVal.FormTypeEx = "179" Or
                pVal.FormTypeEx = "140" Or
                pVal.FormTypeEx = "940" Or
                pVal.FormTypeEx = "1250000940" Or
                pVal.FormTypeEx = "65306" Or
                pVal.FormTypeEx = "-141" Or
                pVal.FormTypeEx = "141" Or
                pVal.FormTypeEx = "-60092" Or
                pVal.FormTypeEx = "720" Then

                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Try
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            If oCFLEvento.BeforeAction = False Then

                                Dim oUserDataSource As SAPbouiCOM.UserDataSource
                                Dim sCFL_ID As String
                                Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim val As String = String.Empty
                                Dim val1 As String = String.Empty


                                If truco = False Then

                                    If (pVal.ItemUID = "etssloc63" And oCFL.UniqueID = "CFL1") Then
                                        sCFL_ID = oCFLEvento.ChooseFromListUID
                                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                        oDataTable = oCFLEvento.SelectedObjects
                                        val = oDataTable.GetValue(0, 0)




                                        Try
                                            truco = True

                                            Try
                                                Dim txtnomCli As SAPbouiCOM.EditText
                                                txtnomCli = oForm.Items.Item("etssloc63").Specific
                                                txtnomCli.Value = val

                                            Catch ex As Exception
                                            End Try


                                        Catch ex As Exception
                                        Finally
                                            ' Rehabilitar el evento después de setear el valor
                                            truco = False

                                        End Try


                                    End If
                                End If


                                'Else
                                'BubbleEvent = False
                            End If

                        Catch ex As Exception

                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Try

                            If pVal.BeforeAction = False Then

                                'guardo el estado del indicador para saber si se controlara la serie electronica por serie o campo

                                If pVal.FormTypeEx = "141" Or pVal.FormTypeEx = "60092" Then

                                    'SeriesElectronicasUDF = oManejoDocumentos.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SeriesFEUDF")
                                    SeriesElectronicasUDF = Functions.VariablesGlobales._vgSerieUDF
                                End If

                                'fin de la verificacion
                                If pVal.FormTypeEx = "65307" Then
                                    Functions.VariablesGlobales._SS_FacturaExportacion = "SI"
                                Else
                                    Functions.VariablesGlobales._SS_FacturaExportacion = "NO"
                                End If
                                SeteaTipoTabla(pVal)

                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)

                                If Not nombreFormulario.Contains("ELECTRONICO") Then
                                    nombreFormulario = mForm.Title
                                End If

                                creaBotonPrueba(mForm)
                                If mForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        CargaItemEnFormularioEcuanexus(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                    Else
                                        CargaItemEnFormulario(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                    End If
                                Else
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        CargaItemEnFormularioEcuanexus(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                    Else

                                        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                            creaBotonPrueba_Factura_GuiaRemision(mForm)
                                            If (pVal.FormTypeEx = "133" Or pVal.FormTypeEx = "720" Or pVal.FormTypeEx = "60091") And ValidarGuiaEnFacturaHeison(mForm, pVal.FormTypeEx) Then
                                                'creaBotonPrueba_Factura_GuiaRemision(mForm)
                                                CargaItemEnFormulario_Factura_GuiaRemision(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                                'ElseIf pVal.FormTypeEx = "720" And ValidarGuiaEnFacturaHeison(mForm, pVal.FormTypeEx) Then
                                                ' CargaItemEnFormulario_Factura_GuiaRemision(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                            Else
                                                CargaItemEnFormulario_Factura_GuiaRemision(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                            End If
                                        End If

                                        CargaItemEnFormulario(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla, pVal.FormUID)
                                    End If

                                End If


                                If pVal.FormTypeEx = "141" Or pVal.FormTypeEx = "60092" Then
                                    creaBotonPrueba_LiquidacionCompra(mForm)
                                    If mForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            CargaItemEnFormulario_LiquidacionCompraEcuanexus(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                        Else
                                            CargaItemEnFormulario_LiquidacionCompra(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                        End If

                                    Else

                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            CargaItemEnFormulario_LiquidacionCompraEcuanexus(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                        Else
                                            CargaItemEnFormulario_LiquidacionCompra(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                        End If

                                    End If

                                End If


                                'se intenta crear un tab para representar os UDF
                                Try

                                    crearTabPrueba_Facturacion(mForm)


                                Catch ex As Exception

                                End Try


                                'se intenta crear tab Localizacion

                                If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                                    Try

                                        crearTabPrueba_Localizacion2(mForm)
                                    Catch ex As Exception

                                    End Try


                                    If Not mForm.Title.ToLower.Contains("cancel") Then

                                        SetearCamposAutoFE_LOC(mForm, "FORMLOAD", pVal.FormTypeEx)

                                    End If


                                End If


                            End If
                        Catch ex As Exception
                            rSboApp.SetStatusBarMessage("et_FORM_LOAD2", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        ' CAMBIA DE SERIE
                        If pVal.ItemUID = "88" Or pVal.ItemUID = "40" Or pVal.ItemUID = "1250000068" Or pVal.ItemUID = "U_DocEmision" Then '40 es id de la serie en transferencias
                            If pVal.BeforeAction = False AndAlso pVal.ItemChanged Then

                                Dim count As Integer = 0

                                SeteaTipoTabla(pVal)

                                Try
                                    Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)

                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        CargaItemEnFormularioEcuanexus(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                    Else
                                        CargaItemEnFormulario(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                    End If


                                    If pVal.FormTypeEx = "-141" Or pVal.FormTypeEx = "141" Or pVal.FormTypeEx = "-60092" Or pVal.FormTypeEx = "60092" Then

                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            CargaItemEnFormulario_LiquidacionCompraEcuanexus(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                        Else
                                            CargaItemEnFormulario_LiquidacionCompra(mForm, "NUEVO", mForm.TypeEx, oTabla)
                                        End If

                                    End If


                                    If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then


                                        Try

                                            ' Dim xforv = rSboApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                                            ' si es cancelacion no realiza nada

                                            If Not mForm.Title.ToLower.Contains("cancel") Then

                                                SetearCamposAutoFE_LOC(mForm, "COMBO", pVal.FormTypeEx)

                                            End If

                                            'Form_LoadEx(xforv.UDFFormUID)
                                            'ValidarCategoriaCamposUsuario(xforv)


                                        Catch ex As Exception
                                            rSboApp.SetStatusBarMessage(ex.Message)
                                        End Try


                                    End If



                                Catch ex As Exception
                                    rSboApp.SetStatusBarMessage("et_COMBO_SELECT", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                End Try


                            End If
                            'ElseIf pVal.ItemUID = "10000329" Then ' COPIAR A 
                            '    Try
                            '        Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                            '        Dim oUDFForm As SAPbouiCOM.Form = rSboApp.Forms.Item(mForm.UDFFormUID)
                            '        oUDFForm.Items.Item("U_CLAVE_ACCESO").Specific.String = ""
                            '        oUDFForm.Items.Item("U_NUM_AUTO_FAC").Specific.String = ""
                            '        oUDFForm.Items.Item("U_OBSERVACION_FACT").Specific.String = ""
                            '        cbEstado = oUDFForm.Items.Item("U_ESTADO_AUTORIZACIO").Specific
                            '        cbEstado.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            '    Catch ex As Exception
                            '    End Try
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        ' CLICK CREAR
                        If pVal.ItemUID = "1" Then

                            If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                                '720 es Salida de Mercacia el addon solo debe actuar Para HEISON para el resto salir
                                If pVal.FormTypeEx = "720" And Functions.VariablesGlobales._ProveedorSAP <> Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then

                                    Exit Sub

                                End If

                                Dim xforv = rSboApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                                ' si es cancelacion no realiza nada

                                If Not xforv.Title.ToLower.Contains("cancel") Then


                                    If pVal.BeforeAction Then

                                        If Functions.VariablesGlobales._SS_ValidarDocumentosUDF = "Y" Then
                                            ValidacionesPrevioCreacionDocumento2(xforv, BubbleEvent)
                                        End If

                                    Else


                                        SetearCamposAutoFE_LOC(xforv, "DespuesPulsar", pVal.FormTypeEx)



                                    End If


                                End If


                            End If

                        ElseIf pVal.ItemUID = "btprints" And pVal.BeforeAction = False Then


                            Dim xforv = rSboApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                            Dim DocEntryfrm = xforv.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).Trim

                            If DocEntryfrm <> "" Then

                                ProcesoImpresion(pVal.FormTypeEx, DocEntryfrm)

                                rSboApp.SetStatusBarMessage("PROCESO IMPRESION CONCLUIDO!",, False)

                            Else

                                rSboApp.SetStatusBarMessage("Opcion no Disponible!",, False)

                            End If


                            'Manejo clik enlaces QR
                        ElseIf (pVal.ItemUID = "pboxQR" Or pVal.ItemUID = "pboxQRL") And pVal.BeforeAction = True Then

                            Try

                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)

                                Dim EnlaceQR As String = ""

                                Select Case pVal.ItemUID
                                    Case "pboxQR"
                                        EnlaceQR = mForm.Items.Item("etssut6").Specific.string
                                    Case "pboxQRL"
                                        EnlaceQR = mForm.Items.Item("etssut25").Specific.string
                                End Select



                                If EnlaceQR <> "" Then

                                    oManejoDocumentos.AbrirEnlaceExterno(EnlaceQR)

                                End If


                            Catch ex As Exception

                                rSboApp.SetStatusBarMessage(NombreAddon & " No es Posible Abrir el enlace", True)

                            End Try

                            'Manejo de Tabs
                        ElseIf (pVal.ItemUID = "TabGenFE" Or pVal.ItemUID = "TabGenLOC") And pVal.BeforeAction = True Then

                            Try
                                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)

                                Select Case pVal.ItemUID
                                    Case "TabGenFE"
                                        mForm.PaneLevel = 50

                                        ObtenerEnlacesURLyGenerarRQ(mForm, pVal.FormUID)

                                    Case "TabGenLOC"

                                        mForm.PaneLevel = 51

                                End Select




                            Catch ex As Exception



                            End Try



                        ElseIf pVal.ItemUID = "btnAccLQ" And pVal.BeforeAction = True Then
                            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                            btnAccLQ = mForm.Items.Item("btnAccLQ").Specific
                            'SeteaTipoTabla(pVal)
                            Acciones_Liquidacion(btnAccLQ, mForm)

                        ElseIf pVal.ItemUID = "btnAccGR" And pVal.BeforeAction = True Then
                            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                            btnAccGR = mForm.Items.Item("btnAccGR").Specific
                            'SeteaTipoTabla(pVal)
                            Acciones_Factura_GuiaRemision(btnAccGR, mForm)

                            'CLICK(ACCION)
                        ElseIf pVal.ItemUID = "btnAccion" And pVal.BeforeAction = True Then
                            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                            btnAccion = mForm.Items.Item("btnAccion").Specific

                            SeteaTipoTabla(pVal)

                            If btnAccion.Caption = "(GS) Ver RIDE" Then
                                Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_CLAVE_ACCESO", 0)
                                'Dim oObject As Object = GetBusinessObjectForm(oForm)
                                'strClaveAcceso = oObject.UserFields.Fields.Item("U_EXX_FE_ClaAcc").Value.Trim()
                                Try

                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Consultando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                    '' TENGO LA RUTA DE LA CARPETA TEMPORAL   
                                    'Dim filepath As String = Path.GetTempPath()
                                    'filepath += ClaveAcceso + ".pdf"

                                    '' SI NO EXISTE EN LA CARPETA TEMPORAL, LO CONSULTO AL WS
                                    'If Not File.Exists(filepath) Then
                                    '    rSboApp.SetStatusBarMessage("Generando el documento, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                                    '    Dim FS As FileStream = Nothing
                                    '    Dim dbbyte As Byte() = oConsultaEmision.ConsultarDocumento(ClaveAcceso, "PDF")
                                    '    FS = New FileStream(filepath, System.IO.FileMode.Create)
                                    '    FS.Write(dbbyte, 0, dbbyte.Length)
                                    '    FS.Close() 
                                    'End If
                                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                                    'If Functions.VariablesGlobales._vgHttps = "Y" Then
                                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                                    'End If
                                    'oManejoDocumentos.SetProtocolosdeSeguridad()
                                    Dim docentry As String = ""
                                    docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAcceso, docentry, oTipoTabla, "pdf")
                                    Else
                                        oManejoDocumentos.ConsultaPDF(ClaveAcceso)
                                    End If


                                Catch x As TimeoutException
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del documento! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Catch ex As Exception
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            ElseIf btnAccion.Caption = "(GS) Ver XML" Then
                                Dim ClaveAccesoXML As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_CLAVE_ACCESO", 0)
                                Try
                                    rSboApp.SetStatusBarMessage(NombreAddon + " -Consultando el XML, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "SalidaporHttps") = "Y" Then
                                    '    ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
                                    'End If
                                    'oManejoDocumentos.SetProtocolosdeSeguridad()
                                    Dim docentry As String = ""
                                    docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAccesoXML, docentry, oTipoTabla, "xml")
                                    Else
                                        oManejoDocumentos.ConsultaXML(ClaveAccesoXML)
                                    End If

                                Catch x As TimeoutException
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del XML! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Catch ex As Exception
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try
                            ElseIf btnAccion.Caption = "(GS) Consultar AUT" Then

                                Try
                                    Dim docentry As String = ""
                                    docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla, True)
                                    Else
                                        oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla, True)
                                    End If

                                    Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                                        mForm.Freeze(True)
                                        rSboApp.ActivateMenuItem("1289")
                                        rSboApp.ActivateMenuItem("1288")
                                    Catch ex As Exception
                                    Finally
                                        mForm.Freeze(False)
                                    End Try

                                Catch ex As Exception
                                    rSboApp.SetStatusBarMessage("Error al intentar Consultar la autorizacion desde eDoc " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                End Try

                            ElseIf btnAccion.Caption = "(GS) Reenviar SRI" Then
                                Dim FechaSalidaEnVivo As String = Functions.VariablesGlobales._vgFechaSalidaEnVivo
                                Utilitario.Util_Log.Escribir_Log("Fecha de salida: " + FechaSalidaEnVivo.ToString, "EventosEmision")
                                If (FechaSalidaEnVivo = "") Or (FechaSalidaEnVivo <> "") Then

                                    Dim ff As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocDate", 0)
                                    ff = CInt(ff)
                                    If FechaSalidaEnVivo = "" Then
                                        FechaSalidaEnVivo = ff
                                    End If
                                    FechaSalidaEnVivo = CInt(FechaSalidaEnVivo)
                                    Utilitario.Util_Log.Escribir_Log("Fecha de salida: " + FechaSalidaEnVivo.ToString, "EventosEmision")
                                    Utilitario.Util_Log.Escribir_Log("Fecha de emision: " + ff.ToString, "EventosEmision")

                                    If ff < FechaSalidaEnVivo Then
                                        'If (DateTime.Compare(_FechaEmisionDoc, _FechaSalidaEnVivo) < 0) Then

                                        rSboApp.SetStatusBarMessage(NombreAddon + " - No se puede reenviar el documento debido a que la fecha de emision es menor a la configurada ! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Else
                                        If Functions.VariablesGlobales._vgNoEnviarRT = "Y" Then
                                            Utilitario.Util_Log.Escribir_Log("Paramero activo no reenciar sri rt: " + Functions.VariablesGlobales._vgNoEnviarRT.ToString, "EventosEmision")
                                            Dim estLQ As String = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_ESTADO", 0)))
                                            If LQEsElectronico = "FE" And (estLQ = "3" Or estLQ = "6" Or estLQ = "4") Then
                                                Utilitario.Util_Log.Escribir_Log("LQEsElectronico: " + LQEsElectronico + " Estado: " + estLQ.ToString, "EventosEmision")
                                                rSboApp.SetStatusBarMessage(NombreAddon + " - No se emitirá la Retención hasta a que la Liquidación se encuentre Autorizada", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                Exit Sub
                                            End If
                                        End If

                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Reenviando Documento! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        Utilitario.Util_Log.Escribir_Log("Reenviando Documento..", "EventosEmision")
                                        ' VALIDAR SI TIENE FOLIO, Y PREGUNTAR SI QUIERE AGREGARLE Y AVANZAR

                                        Utilitario.Util_Log.Escribir_Log("Tabla: " + oTabla, "EventosEmision")

                                        Dim FolioNum As String = ""
                                        Dim docentry As String = ""
                                        Dim objType As String = ""
                                        Dim DocSubType As String = ""
                                        Dim Series As String = ""

                                        Try
                                            FolioNum = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0)
                                            Utilitario.Util_Log.Escribir_Log("Folio: " + FolioNum, "EventosEmision")
                                            docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                                            Utilitario.Util_Log.Escribir_Log("DocEntry: " + docentry, "EventosEmision")
                                            objType = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("ObjType", 0)))
                                            Utilitario.Util_Log.Escribir_Log("ObjType: " + objType, "EventosEmision")
                                            DocSubType = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocSubType", 0)))
                                            Utilitario.Util_Log.Escribir_Log("DocSubType: " + DocSubType, "EventosEmision")

                                            Series = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("Series", 0)))
                                            Utilitario.Util_Log.Escribir_Log("Series: " + Series, "EventosEmision")

                                            If DocSubType = "IX" Then
                                                Functions.VariablesGlobales._SS_FacturaExportacion = "SI"
                                            Else
                                                Functions.VariablesGlobales._SS_FacturaExportacion = "NO"
                                            End If
                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Catch Recupera Variables: " + ex.Message.ToString(), "EventosEmision")
                                        End Try

                                        Try

                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Recuperando variables - Reenviar al SRI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Accediendo a libreria FuncionesAddon - Catch : " + ex.Message.ToString(), "EventosEmision")
                                        End Try

                                        If objType = "18" Or objType = "204" Then '' ** LOGICA CUANDO ES UNA RETENCION, SE SEGMENTO DEBIDO A QUE ES OTRO CAMPO DE USUARIO Y DEPENDE DEL PROVEEDOR SAP BO

                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                Utilitario.Util_Log.Escribir_Log("Proceso ingresando por Proveedor HEINSOHN", "EventosEmision")
                                                Try
                                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                        oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                    Else
                                                        oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                    End If

                                                Catch ex As Exception
                                                    Utilitario.Util_Log.Escribir_Log("Accediendo a libreria Negocio.ManejoDocumentos - Catch : " + ex.Message.ToString(), "EventosEmision")
                                                End Try
                                            Else
                                                Dim RetVal As Long
                                                Dim ErrCode As Long
                                                Dim ErrMsg As String

                                                If objType = "18" Then  'FACTURA DE PROVEEDOR/RETENCION                             
                                                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Then
                                                        Dim sSQL As String = ""
                                                        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                sSQL = "SELECT ""Name"" FROM ""@GS_SERIESE"" INNER JOIN NNM1 B ON A.""Name""= B.""SeriesName"" WHERE B.""SeriesName"" like 'LQE%' AND B.""Series"" =" + Series.ToString
                                                            Else
                                                                sSQL = "SELECT Name FROM [@GS_SERIESE] A WITH(NOLOCK) INNER JOIN NNM1 B WITH(NOLOCK) ON A.Name= B.SeriesName WHERE B.SeriesName like 'LQE%' AND B.Series =" + Series.ToString
                                                            End If
                                                            Dim tipotabla = oFuncionesB1.getRSvalue(sSQL, "Code", "")
                                                            Utilitario.Util_Log.Escribir_Log("Query: " + sSQL.ToString(), "EventosEmision")
                                                            If Not String.IsNullOrEmpty(tipotabla) Then
                                                                oTipoTabla = "LQE"
                                                            Else
                                                                oTipoTabla = "REE"
                                                            End If
                                                            Utilitario.Util_Log.Escribir_Log("tipo tabla: " + tipotabla.ToString(), "EventosEmision")

                                                        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                sSQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE ""U_TipoD"" IN ('RT','LQRT') AND A.""Series"" = " + Series.ToString
                                                            Else
                                                                sSQL = "SELECT ISNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE ""U_TipoD"" IN ('RT','LQRT') AND A.""Series"" = " + Series.ToString
                                                            End If
                                                            Utilitario.Util_Log.Escribir_Log("Consulta: " + sSQL.ToString(), "EventosEmision")
                                                            Dim tipotabla = oFuncionesB1.getRSvalue(sSQL, "U_FE_TipoEmision", "")
                                                            If tipotabla = "RT" Then '' RETENCION
                                                                oTipoTabla = "REE"
                                                            End If
                                                        Else
                                                            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                sSQL = "SELECT ""U_TIPO_DOC"" FROM ""@EXX_DOCUM_LEG_INTER"" A INNER JOIN ""NNM1"" B ON A.""U_NOMBRE"" = B.""SeriesName"" WHERE A.""Code""=B.""SeriesName"" AND B.""Series""= " + Series.ToString
                                                            Else
                                                                sSQL = "SELECT U_TIPO_DOC FROM [@EXX_DOCUM_LEG_INTER] A WITH(NOLOCK) INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_NOMBRE = B.SeriesName WHERE A.Code=B.SeriesName AND B.Series =" + Series.ToString
                                                            End If

                                                            Utilitario.Util_Log.Escribir_Log("Query: " + sSQL.ToString(), "EventosEmision")
                                                            Dim tipotabla = oFuncionesB1.getRSvalue(sSQL, "U_TIPO_DOC", "")
                                                            'If tipotabla = "LC" Then '' LIQUIDACION DE COMPRA
                                                            '    oTipoTabla = "LQE"
                                                            'Else
                                                            If tipotabla = "RT" Then '' RETENCION
                                                                oTipoTabla = "REE"
                                                            End If
                                                            Utilitario.Util_Log.Escribir_Log("tipo tabla: " + tipotabla.ToString(), "EventosEmision")
                                                        End If
                                                        Utilitario.Util_Log.Escribir_Log("Query: " + oTipoTabla.ToString(), "EventosEmision")
                                                        'oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                        'oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                                                        'oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_None
                                                    Else
                                                        oTipoTabla = "RDM"
                                                        'oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                        'oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                                                        oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo
                                                    End If

                                                ElseIf objType = "204" Then  'FACTURA DE ANTICIPO PROVEEDOR/RETENCION                             
                                                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                    oTipoTabla = "REA"
                                                End If
                                                oDocumento.GetByKey(docentry)

                                                Dim TieneFolio As Boolean = False
                                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                    If Not oDocumento.UserFields.Fields.Item("U_COMP_RET").Value = "" Then
                                                        TieneFolio = True
                                                    End If
                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                    If Not oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value = "" Then
                                                        TieneFolio = True
                                                    End If
                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                    If Not oDocumento.UserFields.Fields.Item("U_BPP_MDCD").Value = "" Then
                                                        TieneFolio = True
                                                    End If
                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                    If Not oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value = "" Then
                                                        TieneFolio = True
                                                    End If

                                                    '''''''''   PENDIENTE
                                                End If

                                                'SIN SON SERIES QUE SE DEBE EXCLUIR PONEMOS TIENE FOLIO = TRUE
                                                Dim SerieExcluir As String = ""
                                                SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + Series + "' ", "Code", "")
                                                If Not String.IsNullOrEmpty(SerieExcluir) Then
                                                    TieneFolio = True
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Serie excluida para folear, el documento no se folieara...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If
                                                'END SIN SON SERIES QUE SE DEBE EXCLUIR PONEMOS TIENE FOLIO = TRUE

                                                'SE REVISARA SI ESTA ACTIVADO EL PARAMETRO DE FOLIACION POR POSTN
                                                'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "FoliacionPOSTN") = "Y" Then
                                                'If Functions.VariablesGlobales._FoliacionPostin = "Y" Then
                                                '    TieneFolio = False 'se ajusto a false ya que cuando se tiene esta cionfiguracion hay que ingresar al post dm 20250319
                                                '    rSboApp.SetStatusBarMessage(NombreAddon + " - Foliacion por POSTN ACTIVADA", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                'End If

                                                If Functions.VariablesGlobales._AsignarFolioalReenviarSolsap = "Y" Then
                                                    TieneFolio = False
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Parametro activo Asignar Folio", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If

                                                If TieneFolio = False Then
                                                    ' Utilitario.Util_Log.Escribir_Log("Proceso ingresando por IF No tiene Folio..", "EventosEmision")
                                                    Dim iReturnValue As Integer
                                                    iReturnValue = rSboApp.MessageBox(NombreAddon + " - El documento NO tiene # de Retención, se le asignará desea continuar?", 1, "&SI", "&NO")

                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene # de Retención, se le asignará desea continuar?", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                    If iReturnValue = 2 Then
                                                        BubbleEvent = False

                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene # de Retención, se le asignará desea continuar?, Contesto: NO", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                    ElseIf iReturnValue = 1 Then

                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene # de Retención, se le asignará desea continuar?, Contesto: SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                        Try

                                                            Dim sSQL As String = ""
                                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                    sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@EXX_DOCUM_LEG_INTER"" A "
                                                                    sSQL += " INNER JOIN ""NNM1"" B ON A.""U_NOMBRE"" = B.""SeriesName"" "
                                                                Else
                                                                    sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@EXX_DOCUM_LEG_INTER] A WITH(NOLOCK) "
                                                                    sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_NOMBRE = B.SeriesName "
                                                                End If
                                                                'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                                '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                '        sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@SERIES"" A "
                                                                '        sSQL += " INNER JOIN ""NNM1"" B ON A.""U_SERIE"" = B.""Series"" "
                                                                '    Else
                                                                '        sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@SERIES] A WITH(NOLOCK) "
                                                                '        sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_SERIE = B.Series "
                                                                '    End If
                                                                'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                                '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                '        sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@GS_SERIESE"" A "
                                                                '        sSQL += " INNER JOIN ""NNM1"" B ON A.""Code"" = B.""Series"" "
                                                                '    Else
                                                                '        sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@GS_SERIESE] A WITH(NOLOCK) "
                                                                '        sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.Code = B.Series "
                                                                '    End If
                                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                    sSQL += "WHERE B.""Series"" = " + oDocumento.Series.ToString
                                                                Else
                                                                    sSQL += "WHERE B.Series = " + oDocumento.Series.ToString
                                                                End If

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Ejecutando consulta para obtener # Retención :" + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)



                                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "U_ULT_SECUEN")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Folio :" + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                sCode = oFuncionesB1.getRSvalue(sSQL, "Code")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Code :" + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                numFolio = Integer.Parse(sFolio) + 1

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Sumando 1 al Folio se asiganará :" + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                                                sSQL = Functions.VariablesGlobales._ConsultaFolioSS

                                                                oDocumento.GetByKey(docentry)

                                                                If sSQL = "" Then
                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Por favor agregar consulta para asignacion de Folio", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                    Exit Sub
                                                                End If

                                                                sSQL = sSQL.Replace("TABLA", oTabla)
                                                                sSQL = sSQL.Replace("IDENTIFICADOR", docentry)
                                                                sSQL = sSQL.Replace("TIPDOCE", "('RT','LQRT')")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Ejecutando consulta para obtener # Retención :" + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "Secuencial")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Folio :" + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                sCode = oFuncionesB1.getRSvalue(sSQL, "TipoDoc")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Code :" + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                If sFolio = "" Then
                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro campo Secuencial en la consulta, por favor verificar..!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                    Exit Sub
                                                                Else
                                                                    numFolio = sFolio
                                                                End If

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Sumando 1 al Folio se asiganará :" + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                            End If


                                                            Dim iReturnValue2 As Integer
                                                            iReturnValue2 = rSboApp.MessageBox(NombreAddon + " - El # de Retención que se asiganará es: " + numFolio.ToString() + ", desea continuar?", 1, "&SI", "&NO")

                                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Preguntando: El # de Retención que se asiganará es: " + numFolio.ToString() + ", desea continuar?", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                            If iReturnValue2 = 2 Then

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Contestó NO", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                BubbleEvent = False
                                                            ElseIf iReturnValue2 = 1 Then

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Contestó SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                If objType = "18" Or objType = "204" Then

                                                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                        'oDocumento.FolioNumber = numFolio
                                                                        oDocumento.UserFields.Fields.Item("U_COMP_RET").Value = numFolio.ToString()

                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Contestó SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "# Retención Asignada: " + oDocumento.UserFields.Fields.Item("U_COMP_RET").Value.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                                        oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value = numFolio.ToString()

                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "# Retención Asignada: " + oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)



                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                                        'oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value = numFolio.ToString()
                                                                        'oFuncionesAddon.GuardaLOG(objType, docentry, "# Retención Asignada: " + oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        ''''  PENDIENTE

                                                                        'oDocumento.FolioNumber = numFolio
                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                                        oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value = numFolio.ToString()
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Contestó SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "# Retención Asignada: " + oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                                    End If
                                                                    Try
                                                                        RetVal = oDocumento.Update()
                                                                    Catch ex As Exception
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar folio.!!" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                    End Try

                                                                End If

                                                                If RetVal <> 0 Then

                                                                    rCompany.GetLastError(ErrCode, ErrMsg)

                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar el # de Retención..!!" + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)

                                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "Error al asignar el # de Retención..!!" + ErrCode.ToString() + "-" + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    Exit Sub
                                                                Else
                                                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                        Try
                                                                            Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia (antes del try)", "ManejoDeDocumentos")
                                                                        Catch ex As Exception
                                                                            Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia : " + ex.Message.ToString, "ManejoDeDocumentos")

                                                                        End Try

                                                                        If oFuncionesAddon.ActualizaSecuencia(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            Try
                                                                                Utilitario.Util_Log.Escribir_Log("Dentro del if ActualizaSecuencia linea 551 : ", "ManejoDeDocumentos")
                                                                            Catch ex As Exception
                                                                                Utilitario.Util_Log.Escribir_Log("ERROR Dentro del if ActualizaSecuencia linea 551 : " + ex.Message.ToString, "ManejoDeDocumentos")
                                                                            End Try
                                                                            Try
                                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                End If

                                                                            Catch ex As Exception
                                                                                Utilitario.Util_Log.Escribir_Log("ERROR ACTUALIZAR SECUENCIA : " + ex.Message.ToString, "ManejoDeDocumentos")
                                                                            End Try

                                                                            Exit Sub
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            Exit Sub
                                                                        End If
                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Actualizando la secuencia en la tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                        If oFuncionesAddon.ActualizaSecuencia_ONE(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en la tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                            Exit Sub
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            Exit Sub
                                                                        End If
                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Actualizando la secuencia en la tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        If oFuncionesAddon.ActualizaSecuencia_SYPSOFT(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en la tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            ElseIf Functions.VariablesGlobales._ActApiSS = "Y" Then

                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                            Exit Sub
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            Exit Sub
                                                                        End If

                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                                        'rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en (SS) Documentos Legales..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                        'If oFuncionesAddon.ActualizaSecuenciaSS(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                        Try
                                                                            'oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en (SS) Documentos Legales..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                        Catch ex As Exception
                                                                            Utilitario.Util_Log.Escribir_Log("ERROR ACTUALIZAR SECUENCIA (SS) Documentos Legales : " + ex.Message.ToString, "ManejoDeDocumentos")
                                                                        End Try
                                                                        Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                                                                            mForm.Freeze(True)
                                                                            rSboApp.ActivateMenuItem("1289")
                                                                            rSboApp.ActivateMenuItem("1288")
                                                                        Catch ex As Exception
                                                                        Finally
                                                                            mForm.Freeze(False)
                                                                        End Try

                                                                        Exit Sub
                                                                        'Else
                                                                        '    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                        '    Exit Sub
                                                                        'End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Catch ex As Exception
                                                        End Try
                                                    End If
                                                Else
                                                    Utilitario.Util_Log.Escribir_Log("Proceso ingresando por IF SI tiene # de Retención..", "EventosEmision")
                                                    Try
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                        End If

                                                    Catch ex As Exception
                                                        Utilitario.Util_Log.Escribir_Log("Accediendo a libreria Negocio.ManejoDocumentos - Catch : " + ex.Message.ToString(), "EventosEmision")
                                                    End Try
                                                End If
                                            End If
                                        Else ' LOGICA CUANDO ES OTRO DOCUMENTO QUE NO ES RETENCION
                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                                                Try
                                                    If objType = "13" Then  ' FACTURA DE CLIENTE or  NOTA DE DEBITO
                                                        If DocSubType = "--" Then
                                                            oTipoTabla = "FCE"
                                                        ElseIf DocSubType = "IX" Then
                                                            oTipoTabla = "FCE"
                                                        Else
                                                            oTipoTabla = "NDE"
                                                        End If
                                                    ElseIf objType = "203" Then ' NOTA DE DEBITO
                                                        oTipoTabla = "FAE"
                                                    ElseIf objType = "14" Then 'NOTA DE CREDITO DE CLIENTES
                                                        oTipoTabla = "NCE"
                                                    ElseIf objType = "15" Then 'GUIA DE REMISION - ENTREGA
                                                        oTipoTabla = "GRE"
                                                    ElseIf objType = "67" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
                                                        oTipoTabla = "TRE"
                                                    ElseIf objType = "1250000001" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
                                                        oTipoTabla = "TLE"

                                                    End If

                                                    Try
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                        End If

                                                    Catch ex As Exception
                                                        Utilitario.Util_Log.Escribir_Log("Accediendo a libreria Negocio.ManejoDocumentos - Catch : " + ex.Message.ToString(), "EventosEmision")
                                                    End Try

                                                Catch ex As Exception
                                                    Utilitario.Util_Log.Escribir_Log("Obteniendo tipos de documentos : " + ex.Message.ToString(), "EventosEmision")
                                                End Try

                                            Else ' CUANDO EL PROVEEDOR NO ES HEINSONH SE MANEJA LA FOLIACIÓN
                                                ' CUANDO ES UNA FACTURA, NC, GUIA, O NOTA DE DEBITO

                                                'SIN SON SERIES QUE SE DEBE EXCLUIR PONEMOS TIENE FOLIO = TRUE
                                                Dim SerieExcluir As String = ""
                                                SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + Series + "' ", "Code", "")
                                                If Not String.IsNullOrEmpty(SerieExcluir) Then
                                                    FolioNum = "99"
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Serie excluida para folear, el documento no se folieara...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If
                                                'END SIN SON SERIES QUE SE DEBE EXCLUIR PONEMOS TIENE FOLIO = TRUE
                                                'SE REVISARA SI ESTA ACTIVADO EL PARAMETRO DE FOLIACION POR POSTN
                                                'If ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "FoliacionPOSTN") = "Y" Then

                                                'If Functions.VariablesGlobales._FoliacionPostin = "Y" Then
                                                '    FolioNum = "0"  'EENIA 99 se ajusto a false ya que cuando se tiene esta cionfiguracion hay que ingresar al post dm 20250319
                                                '    rSboApp.SetStatusBarMessage(NombreAddon + " - Foliacion por POSTN ACTIVADA", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                'End If

                                                If Functions.VariablesGlobales._AsignarFolioalReenviarSolsap = "Y" Then
                                                    FolioNum = ""
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Parametro activo Asignar Folio", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If

                                                If FolioNum.ToString().Equals("0") Or FolioNum.ToString.Equals("") Then
                                                    ' Utilitario.Util_Log.Escribir_Log("Proceso ingresando por IF No tiene Folio..", "EventosEmision")
                                                    Dim iReturnValue As Integer
                                                    iReturnValue = rSboApp.MessageBox(NombreAddon + " - El documento NO tiene Folio, se le asignará desea continuar?", 1, "&SI", "&NO")

                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene Folio, se le asignará desea continuar?", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)



                                                    If iReturnValue = 2 Then
                                                        BubbleEvent = False

                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene Folio, se le asignará desea continuar?, Contesto: NO", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                    ElseIf iReturnValue = 1 Then

                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene Folio, se le asignará desea continuar?, Contesto: SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                        Try
                                                            Dim RetVal As Long
                                                            Dim ErrCode As Long
                                                            Dim ErrMsg As String

                                                            If objType = "13" Then  ' FACTURA DE CLIENTE or  NOTA DE DEBITO
                                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                                                If DocSubType = "--" Then
                                                                    oTipoTabla = "FCE"
                                                                Else
                                                                    oTipoTabla = "NDE"
                                                                End If
                                                            ElseIf objType = "203" Then ' NOTA DE DEBITO
                                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                                                                oTipoTabla = "FAE"
                                                            ElseIf objType = "14" Then 'NOTA DE CREDITO DE CLIENTES
                                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                                                oTipoTabla = "NCE"
                                                            ElseIf objType = "15" Then 'GUIA DE REMISION - ENTREGA
                                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                                                oTipoTabla = "GRE"
                                                            ElseIf objType = "67" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
                                                                oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                                                                oTipoTabla = "TRE"
                                                            ElseIf objType = "1250000001" Then 'GUIA DE REMISION - SOLICITUD TRANSFERENCIAS                                            
                                                                oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
                                                                oTipoTabla = "TLE"
                                                            End If

                                                            Dim sSQL As String = ""
                                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                    sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@EXX_DOCUM_LEG_INTER"" A "
                                                                    sSQL += " INNER JOIN ""NNM1"" B ON A.""U_NOMBRE"" = B.""SeriesName"" "
                                                                Else
                                                                    sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@EXX_DOCUM_LEG_INTER] A WITH(NOLOCK) "
                                                                    sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_NOMBRE = B.SeriesName "
                                                                End If
                                                                'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                                '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                '        sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@SERIES"" A "
                                                                '        sSQL += " INNER JOIN ""NNM1"" B ON A.""U_SERIE"" = B.""Series"" "
                                                                '    Else
                                                                '        sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@SERIES] A WITH(NOLOCK) "
                                                                '        sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_SERIE = B.Series "
                                                                '    End If
                                                                'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                                '    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                '        sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@GS_SERIESE"" A "
                                                                '        sSQL += " INNER JOIN ""NNM1"" B ON A.""Code"" = B.""Series"" "
                                                                '    Else
                                                                '        sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@GS_SERIESE] A WITH(NOLOCK) "
                                                                '        sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.Code = B.Series "
                                                                '    End If
                                                                If objType = "67" Or objType = "1250000001" Then
                                                                    oTransferencia.GetByKey(docentry)
                                                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                        sSQL += "WHERE B.""Series"" = " + oTransferencia.Series.ToString
                                                                    Else
                                                                        sSQL += "WHERE B.Series = " + oTransferencia.Series.ToString
                                                                    End If
                                                                Else
                                                                    oDocumento.GetByKey(docentry)
                                                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                        sSQL += "WHERE B.""Series"" = " + oDocumento.Series.ToString
                                                                    Else
                                                                        sSQL += "WHERE B.Series = " + oDocumento.Series.ToString
                                                                    End If
                                                                End If

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Ejecutando consulta para obtener Folio :" + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)



                                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "U_ULT_SECUEN")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Folio :" + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                sCode = oFuncionesB1.getRSvalue(sSQL, "Code")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Code :" + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                numFolio = Integer.Parse(sFolio) + 1

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Sumando 1 al Folio se asiganará :" + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)



                                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then


                                                                sSQL = Functions.VariablesGlobales._ConsultaFolioSS

                                                                If objType = "67" Or objType = "1250000001" Then
                                                                    oTransferencia.GetByKey(docentry)
                                                                Else
                                                                    oDocumento.GetByKey(docentry)
                                                                End If


                                                                If sSQL = "" Then
                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Por favor agregar consulta para asignacion de Folio", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                    Exit Sub
                                                                End If

                                                                If oTipoTabla = "FCE" Or oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then

                                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                                    sSQL = sSQL.Replace("IDENTIFICADOR", docentry)
                                                                    sSQL = sSQL.Replace("TIPDOCE", "('FV')")

                                                                ElseIf oTipoTabla = "NCE" Then

                                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                                    sSQL = sSQL.Replace("IDENTIFICADOR", docentry)
                                                                    sSQL = sSQL.Replace("TIPDOCE", "('NC')")

                                                                ElseIf oTipoTabla = "NDE" Then

                                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                                    sSQL = sSQL.Replace("IDENTIFICADOR", docentry)
                                                                    sSQL = sSQL.Replace("TIPDOCE", "('ND')")

                                                                ElseIf oTipoTabla = "GRE" Or oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then

                                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                                    sSQL = sSQL.Replace("IDENTIFICADOR", docentry)
                                                                    sSQL = sSQL.Replace("TIPDOCE", "('GR')")

                                                                End If

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Ejecutando consulta para obtener Folio :" + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "Secuencial")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Folio :" + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                                sCode = oFuncionesB1.getRSvalue(sSQL, "TipoDoc")

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Code :" + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                                If sFolio = "" Then
                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro campo Secuencial en la consulta, por favor verificar..!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                    Exit Sub
                                                                Else
                                                                    numFolio = sFolio
                                                                End If


                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Folio que se asiganará :" + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                                            End If


                                                            Dim iReturnValue2 As Integer
                                                            iReturnValue2 = rSboApp.MessageBox(NombreAddon + " - El documento # de Folio que se asiganará es: " + numFolio.ToString() + ", desea continuar?", 1, "&SI", "&NO")

                                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Preguntando: El documento # de Folio que se asiganará es: " + numFolio.ToString() + ", desea continuar?", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                            If iReturnValue2 = 2 Then

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Contestó NO", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                BubbleEvent = False
                                                            ElseIf iReturnValue2 = 1 Then

                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Contestó SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                If objType = "67" Or objType = "1250000001" Then
                                                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                        numFolio = Integer.Parse(sFolio) + 1
                                                                    End If

                                                                    oTransferencia.FolioNumber = numFolio

                                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "Folio Asignado: " + oTransferencia.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    oTransferencia.FolioPrefixString = oTipoTabla '"/*Prefijo del folio*/"  

                                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "Prefijo Asignado: " + oTransferencia.FolioPrefixString.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    RetVal = oTransferencia.Update()
                                                                Else
                                                                    ' Utilitario.Util_Log.Escribir_Log("Actualizando Folio", "EventosEmision")
                                                                    oDocumento.FolioNumber = numFolio

                                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "Folio Asignado: " + oDocumento.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    oDocumento.FolioPrefixString = oTipoTabla '"/*Prefijo del folio*/"

                                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "Prefijo Asignado: " + oDocumento.FolioPrefixString.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    oDocumento.Printed = SAPbobsCOM.PrintStatusEnum.psYes

                                                                    ' VALIDO SI EXISTE EL CAMPO NUMFOLIO QUE ESTABA EN EL VERSION PI, PARA SI EXISTE ACTUALIZAR EL FOLIO
                                                                    ' CENTURIOSA USA ESE CAMPO PARA UN REPORTE
                                                                    If oFuncionesB1.checkCampoBD("OINV", "NUMFOLIO") Then
                                                                        oDocumento.UserFields.Fields.Item("U_NUMFOLIO").Value = numFolio.ToString()
                                                                    End If

                                                                    RetVal = oDocumento.Update()

                                                                End If

                                                                If RetVal <> 0 Then

                                                                    rCompany.GetLastError(ErrCode, ErrMsg)

                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar el numero de folio..!!" + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)

                                                                    oFuncionesAddon.GuardaLOG(objType, docentry, "Error al asignar el numero de folio..!!" + ErrCode.ToString() + "-" + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                    Exit Sub
                                                                Else
                                                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Actualizando la secuencia en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Antes de ingresar a la funcion actualiza secuencia", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                        If objType = "67" Then
                                                                            If oFuncionesAddon.ActualizaSecuencia(sCode, numFolio, oTransferencia.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                End If

                                                                            Else
                                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                                Exit Sub
                                                                            End If
                                                                        Else
                                                                            If oFuncionesAddon.ActualizaSecuencia(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                                oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                End If

                                                                            Else
                                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                                Exit Sub
                                                                            End If
                                                                        End If
                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Actualizando la secuencia en la tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        If oFuncionesAddon.ActualizaSecuencia_ONE(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en la tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If objType = "67" Or objType = "1250000001" Then
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                End If

                                                                                Exit Sub
                                                                            Else
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                End If

                                                                                Exit Sub
                                                                            End If
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            Exit Sub
                                                                        End If
                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Actualizando la secuencia en la tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        If oFuncionesAddon.ActualizaSecuencia_SYPSOFT(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en la tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If objType = "67" Then
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                End If

                                                                                Exit Sub
                                                                            Else
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                End If

                                                                                Exit Sub
                                                                            End If
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            Exit Sub
                                                                        End If
                                                                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                                        'rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en (SS) Documentos Legales..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                                        'oFuncionesAddon.GuardaLOG(objType, docentry, "Actualizando la secuencia en (SS) Documentos Legales..!!", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        'oFuncionesAddon.GuardaLOG(objType, docentry, "Antes de ingresar a la funcion actualiza secuencia", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)


                                                                        If objType = "67" Then
                                                                            'If oFuncionesAddon.ActualizaSecuenciaSS(sCode, numFolio, oTransferencia.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            'oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                            End If

                                                                            'Else
                                                                            '    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            '    Exit Sub
                                                                            'End If
                                                                        Else
                                                                            'If oFuncionesAddon.ActualizaSecuenciaSS(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                                            '    oFuncionesAddon.GuardaLOG(objType, docentry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                            'Else
                                                                            '    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            '    Exit Sub
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If

                                                        Catch ex As Exception
                                                            Utilitario.Util_Log.Escribir_Log("actualiza secuencia error : " + ex.Message.ToString(), "Ma")
                                                        End Try
                                                    End If
                                                Else
                                                    Utilitario.Util_Log.Escribir_Log("Proceso ingresando por IF SI tiene Folio..", "EventosEmision")
                                                    Try
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla)
                                                        End If

                                                    Catch ex As Exception
                                                        Utilitario.Util_Log.Escribir_Log("error al actualizar la secuencia" + ex.Message.ToString, "ManejoDeDocumentos")
                                                    End Try
                                                End If
                                            End If
                                        End If


                                        ' END VALIDAR SI TIENE FOLIO, Y PREGUNTAR SI QUIERE AGREGARLE Y AVANZAR

                                        'mForm.Freeze(True)
                                        'Dim Docnum As SAPbouiCOM.EditText = CType(mForm.Items.Item("5").Specific, SAPbouiCOM.EditText)
                                        'Dim ValueDocNum As String = Docnum.Value

                                        'mForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        'Dim field As SAPbouiCOM.EditText = CType(mForm.Items.Item("5").Specific, SAPbouiCOM.EditText)

                                        'field.String = ValueDocNum

                                        Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                                            mForm.Freeze(True)
                                            rSboApp.ActivateMenuItem("1289")
                                            rSboApp.ActivateMenuItem("1288")
                                        Catch ex As Exception
                                        Finally
                                            mForm.Freeze(False)
                                        End Try
                                    End If
                                End If
                            ElseIf btnAccion.Caption = "(GS) Reenviar MAIL" Then
                                mForm.Freeze(True)
                                Try
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Re Enviando Mail... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_CLAVE_ACCESO", 0)
                                    Dim docentry As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)
                                    Dim SQUERY As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", oTabla), "IDENTIFICADOR", docentry)
                                        'SQUERY = String.Format("SELECT A.""E_Mail"" FROM ""{0}"" O INNER JOIN ""OCRD"" A ON O.""CardCode"" = A.""CardCode"" WHERE O.""DocEntry"" =  {1}", oTabla, docentry)
                                    Else
                                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", oTabla), "IDENTIFICADOR", docentry)
                                        'SQUERY = String.Format("SELECT A.E_Mail FROM {0} O WITH(NOLOCK) INNER JOIN OCRD A WITH(NOLOCK) ON O.CardCode = A.CardCode WHERE O.DocEntry = {1} ", oTabla, docentry)
                                    End If
                                    Dim sCorreoNuevo As String = oFuncionesB1.getRSvalue(SQUERY, "Email", "")

                                    If String.IsNullOrEmpty(sCorreoNuevo) Then
                                        rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro Email, verificar consulta..!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        mForm.Freeze(False)
                                    Else
                                        Try
                                            Utilitario.Util_Log.Escribir_Log("Consulta: " + SQUERY.ToString() + " - Respuesta: " + sCorreoNuevo.ToString, "EventosEmision")
                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                'oManejoDocumentosEcua.ReenvioMail(sCorreoNuevo, ClaveAcceso)
                                            Else
                                                If oManejoDocumentos.ReenvioMail(sCorreoNuevo, ClaveAcceso) Then
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Mail Re Enviado, Listo!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If
                                            End If


                                            mForm.Freeze(False)
                                        Catch ex As Exception
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al al reenviar el Mail.!!: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Utilitario.Util_Log.Escribir_Log("error al reenviar el e-mail: " + ex.Message.ToString, "EventosEmision")
                                        End Try
                                    End If
                                Catch ex As Exception
                                    mForm.Freeze(False)
                                End Try

                            End If


                        End If
                End Select

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Acciones_BotonCargaEB(mForm As SAPbouiCOM.Form, ByRef selectedFile As String)

        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim sheet As Worksheet = Nothing
        Dim range As Range = Nothing

        Try

            Dim ctaContable As String = CType(mForm.Items.Item("17").Specific, SAPbouiCOM.EditText).Value

            If ctaContable = "" Then
                Throw New Exception("Debe primero seleccionar una cuenta de mayor.")

            End If

            ' Buscar el nombre del banco asociado a la cuenta contable
            'Dim sQueryBanco = $"SELECT A.""AcctName"" FROM ""OACT"" A INNER JOIN ""DSC1"" B ON A.""AcctCode"" = B.""GLAccount"" WHERE A.""Segment_0"" + '-' + A.""Segment_1"" + '-' + A.""Segment_2"" = '{ctaContable}'"
            Dim sQueryBanco = $"SELECT A.""AcctName"" FROM ""OACT"" A INNER JOIN ""DSC1"" B ON A.""AcctCode"" = B.""GLAccount"" WHERE A.""FormatCode"" = '{Replace(ctaContable, "-", "")}'"
            Dim rs As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Utilitario.Util_Log.Escribir_Log("Consulta nombre cuenta extracto: " + sQueryBanco, "EventosEmision")
            rs.DoQuery(sQueryBanco)

            Dim nombreBanco As String = ""
            If rs.RecordCount > 0 Then
                nombreBanco = rs.Fields.Item("AcctName").Value
            End If

            Dim oMatrix As SAPbouiCOM.Matrix = CType(mForm.Items.Item("5").Specific, SAPbouiCOM.Matrix)









            rSboApp.SetStatusBarMessage("Archivo seleccionado: " & selectedFile, SAPbouiCOM.BoMessageTime.bmt_Short, False)

            excelApp = New Application
            workbook = excelApp.Workbooks.Open(selectedFile)
            sheet = workbook.Sheets(1)
            range = sheet.UsedRange

            Dim rows As Integer = range.Rows.Count
            Dim cols As Integer = range.Columns.Count

            ' Detectar tipo de banco y configurar encabezados esperados
            Dim tipoBanco As String = ""
            Dim columnasEsperadas As String() = {}
            Dim filaInicio As Integer = 2
            Dim columnasMapeo As New Dictionary(Of String, String)

            If nombreBanco.ToUpperInvariant().Contains("PICHINCHA") Then
                tipoBanco = "PICHINCHA"
                columnasEsperadas = {"Fecha", "Codigo", "Concepto", "Tipo", "Documento", "Oficina", "Monto", "Saldo"}
                filaInicio = 2
                columnasMapeo.Add("Fecha", "Fecha")
                columnasMapeo.Add("Codigo", "Codigo")
                columnasMapeo.Add("Concepto", "Concepto")
                columnasMapeo.Add("Tipo", "Tipo")
                columnasMapeo.Add("Documento", "Documento")
                columnasMapeo.Add("Oficina", "Oficina")
                columnasMapeo.Add("Monto", "Monto")
                'columnasMapeo.Add("Saldo", "Saldo")
            ElseIf nombreBanco.ToUpperInvariant().Contains("PRODUBANCO") Then
                tipoBanco = "PRODUBANCO"
                columnasEsperadas = {"FECHA", "REFERENCIA", "DESCRIPCION", "+/-", "VALOR", "SALDO CONTABLE", "SALDO DISPONIBLE", "OFICINA"}
                filaInicio = 11
                columnasMapeo.Add("Fecha", "FECHA")
                columnasMapeo.Add("Codigo", "REFERENCIA")
                columnasMapeo.Add("Concepto", "DESCRIPCION")
                columnasMapeo.Add("Tipo", "+/-")
                columnasMapeo.Add("Documento", "REFERENCIA")
                columnasMapeo.Add("Oficina", "OFICINA")
                columnasMapeo.Add("Monto", "VALOR")
                'columnasMapeo.Add("Saldo", "SALDO DISPONIBLE")
            ElseIf nombreBanco.ToUpperInvariant().Contains("PACIFICO") Then
                tipoBanco = "PACIFICO"
                columnasEsperadas = {"Estado", "FechaContable", "Lugar", "Caja", "TipoMov", "Nut", "Valor", "Numero", "Concepto", "SaldoDespMov", "Descripcion", "FechaReal"}
                filaInicio = 2
                columnasMapeo.Add("Fecha", "FechaContable")
                columnasMapeo.Add("Codigo", "Caja")
                columnasMapeo.Add("Concepto", "Concepto")
                columnasMapeo.Add("Tipo", "TipoMov")
                columnasMapeo.Add("Documento", "Numero")
                columnasMapeo.Add("Oficina", "Lugar")
                columnasMapeo.Add("Monto", "Valor")
                'columnasMapeo.Add("Saldo", "SALDO DISPONIBLE")
            Else
                Throw New Exception("El nombre del banco no coincide con Pichincha o Produbanco.")
            End If

            ' Leer encabezados dinámicamente
            Dim headers As New Dictionary(Of String, Integer)
            For col As Integer = 1 To cols
                Dim header As String = Convert.ToString(range.Cells(filaInicio - 1, col).Value)?.Trim().ToUpperInvariant()
                If Not String.IsNullOrEmpty(header) AndAlso Not headers.ContainsKey(header) Then
                    headers.Add(header, col)
                End If
            Next

            ' Validar columnas requeridas
            Dim columnasFaltantes As New List(Of String)
            For Each COL As String In columnasEsperadas
                If Not headers.ContainsKey(COL.ToUpperInvariant()) Then
                    columnasFaltantes.Add(COL)
                End If
            Next

            If columnasFaltantes.Count > 0 Then
                rSboApp.MessageBox("Faltan columnas: " & String.Join(", ", columnasFaltantes))
                Exit Sub
            End If

            ' Leer datos desde la fila correspondiente
            For i As Integer = filaInicio To rows
                Dim rowIndex As Integer = oMatrix.RowCount

                ' Extraer valores usando mapeo dinámico
                Dim fechaCruda = range.Cells(i, headers(columnasMapeo("Fecha").ToUpperInvariant())).Value
                Dim fechaConvertida As Date = oFuncionesB1.ParseFechaDesdeExcel(fechaCruda, nombreBanco)

                If oMatrix.Columns.Item("2").Editable Then
                    CType(oMatrix.Columns.Item("2").Cells.Item(rowIndex).Specific, SAPbouiCOM.EditText).Value = fechaConvertida.ToString("yyyyMMdd")
                End If

                Dim codigo = Convert.ToString(range.Cells(i, headers(columnasMapeo("Codigo").ToUpperInvariant())).Value)
                Dim concepto = Convert.ToString(range.Cells(i, headers(columnasMapeo("Concepto").ToUpperInvariant())).Value)
                Dim tipo = Convert.ToString(range.Cells(i, headers(columnasMapeo("Tipo").ToUpperInvariant())).Value)
                Dim documento = Convert.ToString(range.Cells(i, headers(columnasMapeo("Documento").ToUpperInvariant())).Value)
                Dim oficina = Convert.ToString(range.Cells(i, headers(columnasMapeo("Oficina").ToUpperInvariant())).Value)
                Dim monto = Convert.ToString(range.Cells(i, headers(columnasMapeo("Monto").ToUpperInvariant())).Value)
                'Dim saldo = Convert.ToString(range.Cells(i, headers(columnasMapeo("Saldo").ToUpperInvariant())).Value)

                ' Mostrar mensaje de avance
                rSboApp.SetStatusBarMessage($"Fila {i - 1}: Fecha={fechaConvertida}, Código={codigo}, Concepto={concepto}, Monto={monto}", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                ' Asignar valores a la matriz
                'If oMatrix.Columns.Item("2").Editable Then oMatrix.Columns.Item("2").Cells.Item(rowIndex).Specific.Value = fechaFormateada
                If oMatrix.Columns.Item("3").Editable Then oMatrix.Columns.Item("3").Cells.Item(rowIndex).Specific.Value = documento
                If oMatrix.Columns.Item("4").Editable Then oMatrix.Columns.Item("4").Cells.Item(rowIndex).Specific.Value = concepto

                If tipoBanco = "PICHINCHA" Then
                    If tipo = "C" AndAlso oMatrix.Columns.Item("6").Editable Then
                        oMatrix.Columns.Item("6").Cells.Item(rowIndex).Specific.Value = monto
                    ElseIf tipo = "D" AndAlso oMatrix.Columns.Item("5").Editable Then
                        oMatrix.Columns.Item("5").Cells.Item(rowIndex).Specific.Value = monto
                    End If
                ElseIf tipoBanco = "PRODUBANCO" Then
                    If tipo = "+" AndAlso oMatrix.Columns.Item("6").Editable Then
                        oMatrix.Columns.Item("6").Cells.Item(rowIndex).Specific.Value = monto
                    ElseIf tipo = "-" AndAlso oMatrix.Columns.Item("5").Editable Then
                        oMatrix.Columns.Item("5").Cells.Item(rowIndex).Specific.Value = monto
                    End If
                ElseIf tipoBanco = "PACIFICO" Then
                    If tipo = "N/C" AndAlso oMatrix.Columns.Item("6").Editable Then
                        oMatrix.Columns.Item("6").Cells.Item(rowIndex).Specific.Value = monto
                    ElseIf tipo = "N/D" AndAlso oMatrix.Columns.Item("5").Editable Then
                        oMatrix.Columns.Item("5").Cells.Item(rowIndex).Specific.Value = monto
                    End If
                End If

                'If oMatrix.Columns.Item("7").Editable Then oMatrix.Columns.Item("7").Cells.Item(rowIndex).Specific.Value = saldo
            Next


        Catch ex As Exception
            rSboApp.MessageBox("Error durante el proceso de carga: " & ex.Message)
        Finally
            If Not IsNothing(workbook) Then workbook.Close(False)
            If Not IsNothing(excelApp) Then excelApp.Quit()
            If Not IsNothing(range) Then Runtime.InteropServices.Marshal.ReleaseComObject(range)
            If Not IsNothing(sheet) Then Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            If Not IsNothing(workbook) Then Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            If Not IsNothing(excelApp) Then Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Private Sub Acciones_Liquidacion(btnAccLQ As SAPbouiCOM.ButtonCombo, mForm As SAPbouiCOM.Form)
        Try
            oTabla = "OPCH"
            oTipoTabla = "LQE"
            Dim _DocEntry As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)
            Dim _Serie As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("Series", 0)
            If btnAccLQ.Caption = "(GS) Ver RIDE" Then
                Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_CLAVE", 0)
                Dim docentry As String = ""
                docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                Try

                    rSboApp.SetStatusBarMessage(NombreAddon + " -Consultando La Lquidación de Compra, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    'oManejoDocumentos.SetProtocolosdeSeguridad()
                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                        oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAcceso, docentry, oTipoTabla, "pdf")
                    Else
                        oManejoDocumentos.ConsultaPDF(ClaveAcceso)
                    End If


                Catch x As TimeoutException
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del documento! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            ElseIf btnAccLQ.Caption = "(GS) Ver XML" Then
                Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_CLAVE", 0)
                Dim docentry As String = ""
                docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                Try

                    rSboApp.SetStatusBarMessage(NombreAddon + " -Consultando XML LIquidación de Compra, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    'SetProtocolosdeSeguridad()
                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                        oManejoDocumentosEcua.Consulta_PDF_XML(ClaveAcceso, docentry, oTipoTabla, "xml")
                    Else
                        oManejoDocumentos.ConsultaXML(ClaveAcceso)
                    End If


                Catch x As TimeoutException
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del XML! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try

            ElseIf btnAccLQ.Caption = "(GS) Consultar AUT" Then
                Try
                    Dim docentry As String = ""
                    docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                        oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla, True)
                    Else
                        oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla, True)
                    End If

                    Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                        mForm.Freeze(True)
                        rSboApp.ActivateMenuItem("1289")
                        rSboApp.ActivateMenuItem("1288")
                    Catch ex As Exception
                    Finally
                        mForm.Freeze(False)

                    End Try
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Error al intentar Consultar la autorizacion desde eDoc " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try

            ElseIf btnAccLQ.Caption = "(GS) Reenviar SRI" Then

                rSboApp.SetStatusBarMessage(NombreAddon + " - Reenviando Liquidación de Compra! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Utilitario.Util_Log.Escribir_Log("Reenviando Documento..", "EventosEmision")
                ' VALIDAR SI TIENE FOLIO, Y PREGUNTAR SI QUIERE AGREGARLE Y AVANZAR


                Utilitario.Util_Log.Escribir_Log("Tabla: " + oTabla, "EventosEmision")
                Dim SerieExcluir As String = "N"
                'SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + mForm.Items.Item("88").Specific.Value.ToString.Trim + "' ", "Code", "")
                SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + _Serie + "' ", "Code", "")
                If String.IsNullOrEmpty(SerieExcluir) Then
                    SerieExcluir = "N"
                Else
                    SerieExcluir = "Y"
                End If
                Utilitario.Util_Log.Escribir_Log("Verificar Serie Excluir: " + SerieExcluir, "ManejoDeDocumentos")
                If Functions.VariablesGlobales._FoliacionPostin = "Y" Then
                    SerieExcluir = "Y"
                End If
                If Functions.VariablesGlobales._AsignarFolioalReenviarSolsap = "Y" Then
                    SerieExcluir = "N"
                End If
                If SerieExcluir = "Y" And (Functions.VariablesGlobales._vgSerieUDF = "N" Or Functions.VariablesGlobales._vgSerieUDF = "Y") Then
                    'If SerieExcluir = "Y" And (Functions.VariablesGlobales._vgSerieUDF = "N" And Functions.VariablesGlobales._vgSerieUDF = "Y") Then

                    Dim EnviaDocumentosEnBackGround As String = "N"
                    'EnviaDocumentosEnBackGround = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EnviaDocumentosEnBackGround")
                    EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung
                    ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA
                    'If Not EnviaDocumentosEnBackGround = "Y" Then ' se comenta ya que cuando la serie esta excluida y activo parametro para envia servicio no se reenvia
                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                        oManejoDocumentosEcua.ProcesaEnvioDocumento(_DocEntry, oTipoTabla)
                    Else
                        oManejoDocumentos.ProcesaEnvioDocumento(_DocEntry, oTipoTabla)
                    End If

                    '    Utilitario.Util_Log.Escribir_Log("EnviaDocumentosEnBackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                    'Else
                    '    Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                    '    rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, _DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'End If
                    'oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                Else
                    Dim FolioNum As String = ""
                    Dim docentry As String = ""
                    Dim objType As String = ""
                    Dim DocSubType As String = ""
                    Dim Series As String = ""

                    Try
                        If Functions.VariablesGlobales._vgFolioLQUDF = "Y" Then
                            FolioNum = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_BPP_MDCD", 0)
                            'FolioNum = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0)

                        Else
                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                FolioNum = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_N_DOCSRI", 0) 'U_N_DOCSRI
                                Utilitario.Util_Log.Escribir_Log("Folio: " + FolioNum, "EventosEmision")
                            Else
                                FolioNum = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0)
                                Utilitario.Util_Log.Escribir_Log("Folio: " + FolioNum, "EventosEmision")
                            End If
                        End If

                        docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                        Utilitario.Util_Log.Escribir_Log("DocEntry: " + docentry, "EventosEmision")
                        objType = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("ObjType", 0)))
                        Utilitario.Util_Log.Escribir_Log("ObjType: " + objType, "EventosEmision")
                        DocSubType = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocSubType", 0)))
                        Utilitario.Util_Log.Escribir_Log("DocSubType: " + DocSubType, "EventosEmision")

                        Series = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("Series", 0)))
                        Utilitario.Util_Log.Escribir_Log("Series: " + Series, "EventosEmision")

                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Catch Recupera Variables: " + ex.Message.ToString(), "EventosEmision")
                    End Try

                    Try
                        oFuncionesAddon.GuardaLOG(objType, docentry, "Recuperando variables LQ - Reenviar al SRI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("Accediendo a libreria FuncionesAddon - Catch : " + ex.Message.ToString(), "EventosEmision")
                    End Try

                    If FolioNum.ToString().Equals("0") Or FolioNum.ToString.Equals("") Then
                        ' Utilitario.Util_Log.Escribir_Log("Proceso ingresando por IF No tiene Folio..", "EventosEmision")
                        Dim iReturnValue As Integer
                        iReturnValue = rSboApp.MessageBox(NombreAddon + " - La Liquidación NO tiene Folio, se le asignará desea continuar?", 1, "&SI", "&NO")
                        oFuncionesAddon.GuardaLOG(objType, docentry, "El documento NO tiene Folio, se le asignará desea continuar?", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                        If iReturnValue = 2 Then
                            'BubbleEvent = False
                            oFuncionesAddon.GuardaLOG(objType, docentry, "La Liquidación NO tiene Folio, se le asignará desea continuar?, Contesto: NO", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                            Exit Sub
                        ElseIf iReturnValue = 1 Then
                            oFuncionesAddon.GuardaLOG(objType, docentry, "La Liquidación NO tiene Folio, se le asignará desea continuar?, Contesto: SI", Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)
                            Try
                                Dim sSQL As String = ""
                                Dim RetVal As Long
                                Dim ErrCode As Long
                                Dim ErrMsg As String

                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                oDocumento.GetByKey(docentry)

                                If Functions.VariablesGlobales._vgFolioLQUDF = "Y" Then
                                    Dim est As String = "00"
                                    'Dim typeEx = "", idForm As String = ""
                                    'typeEx = oFuncionesB1.FormularioActivo(idForm)
                                    'If typeEx = "141" Then
                                    '    Dim count As Integer = 0
                                    '    Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)
                                    '    Dim oFrmUser As SAPbouiCOM.Form
                                    '    Try
                                    '        oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                                    '    Catch ex As Exception
                                    '        rSboApp.SendKeys("^+U")
                                    '        oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                                    '    End Try
                                    est = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_Establecimiento", 0)
                                    sSQL = "SELECT ISNULL(""U_Sec"",'0') AS U_ULT_SECUEN, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_Est"" = " + est.ToString
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener Secuencial LQ: " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    'est = oFrmUser.Items.Item("U_Establecimiento").Specific
                                    'End If
                                Else

                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                        sSQL = Functions.VariablesGlobales._ConsultaFolioSS

                                        If sSQL = "" Then
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Por favor agregar consulta para asignacion de Folio", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Exit Sub
                                        End If

                                        sSQL = sSQL.Replace("TABLA", oTabla)
                                        sSQL = sSQL.Replace("IDENTIFICADOR", docentry)
                                        sSQL = sSQL.Replace("TIPDOCE", "('LQ','LQRT')")

                                        sFolio = oFuncionesB1.getRSvalue(sSQL, "Secuencial")

                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Folio :" + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                        sCode = oFuncionesB1.getRSvalue(sSQL, "TipoDoc")

                                        oFuncionesAddon.GuardaLOG(objType, docentry, "Code :" + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Reenvío, Functions.FuncionesAddon.TipoLog.Emision)

                                        If sFolio = "" Then
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro campo Secuencial en la consulta, por favor verificar..!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Exit Sub
                                        Else
                                            numFolio = sFolio
                                        End If

                                    Else
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            sSQL = "SELECT IFNULL(""U_Sec"",'0') AS U_ULT_SECUEN, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                        Else
                                            sSQL = "SELECT ISNULL(""U_Sec"",'0') AS U_ULT_SECUEN, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                        End If
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener Secuencial : " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                        sFolio = oFuncionesB1.getRSvalue(sSQL, "U_ULT_SECUEN")
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se obtuvo el siguiente # de Folio : " + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        If sFolio = "" Then
                                            sFolio = 0
                                        End If
                                        sCode = oFuncionesB1.getRSvalue(sSQL, "Code")
                                        If sCode = "" Then
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Al NO Obtenerse el Code, implica que no esta registrado en la tabla Liquidación de Compra la serie del documento : " + oDocumento.Series.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se Omitió el envío al SRI, BubbleEvent = False", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            'BubbleEvent = False
                                            Exit Sub
                                        End If

                                        numFolio = Integer.Parse(sFolio) + 1
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se sumo 1 al # de Folio: " + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                    End If

                                End If


                                If Functions.VariablesGlobales._vgFolioLQUDF = "Y" Then
                                    oDocumento.UserFields.Fields.Item("U_BPP_MDCD").Value = numFolio.ToString.PadLeft(9, "0")
                                Else
                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                        Dim _est As String = ""
                                        Dim Est As String = ""
                                        Dim _puntoemi As String = ""
                                        Dim PuntoEmi As String = ""
                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            _est = "SELECT IFNULL(""U_Estable"",'0') AS Establecimiento, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                            _puntoemi = "SELECT IFNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                        Else
                                            _est = "SELECT ISNULL(""U_Estable"",'0') AS Establecimiento, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                            _puntoemi = "SELECT ISNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                        End If
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener establecimiento : " + _est.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener punto de emision : " + _puntoemi.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        Est = oFuncionesB1.getRSvalue(_est, "Establecimiento")
                                        PuntoEmi = oFuncionesB1.getRSvalue(_puntoemi, "PuntoEmision")
                                        'U_N_DOCSRI
                                        oDocumento.UserFields.Fields.Item("U_N_DOCSRI").Value = Est.ToString + PuntoEmi.ToString + numFolio.ToString.PadLeft(9, "0")
                                    Else
                                        oDocumento.FolioNumber = numFolio
                                        oDocumento.FolioPrefixString = oTipoTabla
                                    End If
                                    'oDocumento.FolioNumber = numFolio
                                    'oDocumento.FolioPrefixString = oTipoTabla
                                End If
                                '"/*Prefijo del folio*/"
                                RetVal = oDocumento.Update()
                                If RetVal <> 0 Then

                                    rCompany.GetLastError(ErrCode, ErrMsg)

                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar el numero de folio a la Liquidación de Compra..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al asignar el numero de folio a la Liquidación de Compra..!!" + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    Exit Sub
                                Else
                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                        'rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en tabla de Liquidación de Compra..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        'If oFuncionesAddon.ActualizaSecuencia_LiquidacionDeCompra(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                        Else
                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                        End If

                                        'End If
                                    Else
                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en tabla de Liquidación de Compra..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If oFuncionesAddon.ActualizaSecuencia_LiquidacionDeCompra(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                            Else
                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                            End If

                                        End If
                                    End If

                                End If
                            Catch ex As Exception
                            End Try

                        End If

                    Else
                        Utilitario.Util_Log.Escribir_Log("Proceso ingresando por IF SI tiene Folio..", "EventosEmision")
                        Try
                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                oManejoDocumentosEcua.ProcesaEnvioDocumento(docentry, oTipoTabla)
                            Else
                                oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla)
                            End If


                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("error al actualizar la secuencia" + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try

                    End If
                End If

                Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                    mForm.Freeze(True)
                    rSboApp.ActivateMenuItem("1289")
                    rSboApp.ActivateMenuItem("1288")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("error generar evento de refrescar LQ" + ex.Message.ToString, "ManejoDeDocumentos")
                Finally
                    mForm.Freeze(False)
                End Try

            ElseIf btnAccLQ.Caption = "(GS) Reenviar MAIL" Then
                mForm.Freeze(True)
                Try
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Re Enviando Mail... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_CLAVE", 0)
                    Dim docentry As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)
                    Dim SQUERY As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", oTabla), "IDENTIFICADOR", docentry)
                        'SQUERY = String.Format("SELECT A.""E_Mail"" FROM ""{0}"" O INNER JOIN ""OCRD"" A ON O.""CardCode"" = A.""CardCode"" WHERE O.""DocEntry"" =  {1}", oTabla, docentry)
                    Else
                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", oTabla), "IDENTIFICADOR", docentry)
                        'SQUERY = String.Format("SELECT A.E_Mail FROM {0} O WITH(NOLOCK) INNER JOIN OCRD A WITH(NOLOCK) ON O.CardCode = A.CardCode WHERE O.DocEntry = {1} ", oTabla, docentry)
                    End If
                    Dim sCorreoNuevo As String = oFuncionesB1.getRSvalue(SQUERY, "Email", "")
                    If String.IsNullOrEmpty(sCorreoNuevo) Then
                        rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro Email, verificar consulta..!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    Else
                        Try
                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                'oManejoDocumentosEcua.ReenvioMail(sCorreoNuevo, ClaveAcceso)
                            Else
                                oManejoDocumentos.ReenvioMail(sCorreoNuevo, ClaveAcceso)
                            End If

                            rSboApp.SetStatusBarMessage(NombreAddon + " - Mail Re Enviado, Listo!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            mForm.Freeze(False)
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("error al reenviar el e-mail: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try
                    End If
                Catch ex As Exception
                    mForm.Freeze(False)
                End Try
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Acciones_Factura_GuiaRemision(_btnAccGR As SAPbouiCOM.ButtonCombo, mForm As SAPbouiCOM.Form)
        Try
            'oTabla = "OINV"
            'oTipoTabla = "FCE"
            Dim _DocEntry As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)
            Dim _Serie As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("Series", 0)
            If _btnAccGR.Caption = "(GS) Ver RIDE" Then
                Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_GR_CLAVE", 0)
                Dim docentry As String = ""
                docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                Try

                    rSboApp.SetStatusBarMessage(NombreAddon + " -Consultando PDF Guia Remision, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    oManejoDocumentos.ConsultaPDF(ClaveAcceso)



                Catch x As TimeoutException
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del documento! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try

            ElseIf _btnAccGR.Caption = "(GS) Ver XML" Then
                Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_GR_CLAVE", 0)
                Dim docentry As String = ""
                docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))
                Try

                    rSboApp.SetStatusBarMessage(NombreAddon + " -Consultando XML Guia Remision, por favor espere..! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    oManejoDocumentos.ConsultaXML(ClaveAcceso)



                Catch x As TimeoutException
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Se excedio el tiempo de consulta del XML! ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Existio un error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try

            ElseIf _btnAccGR.Caption = "(GS) Consultar AUT" Then
                Try
                    Dim docentry As String = ""
                    docentry = LTrim(RTrim(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)))

                    oManejoDocumentos.ProcesaEnvioDocumento(docentry, oTipoTabla, True)


                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Error al intentar Consultar la autorizacion desde eDoc " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try

            ElseIf _btnAccGR.Caption = "(GS) Reenviar SRI" Then

                rSboApp.SetStatusBarMessage(NombreAddon + " - Reenviando Guia Remision! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Utilitario.Util_Log.Escribir_Log("Reenviando Documento..", "EventosEmision")
                ' VALIDAR SI TIENE FOLIO, Y PREGUNTAR SI QUIERE AGREGARLE Y AVANZAR


                'Utilitario.Util_Log.Escribir_Log("Tabla: " + oTabla, "EventosEmision")
                'If oTabla = "FCE" Then
                oManejoDocumentos.ProcesaEnvioDocumento(_DocEntry, "GRE")
                'Else
                'oManejoDocumentos.ProcesaEnvioDocumento(_DocEntry, "GRESM")
                'End If



                Try ' RETROCEDO Y AVANZO PARA ACTUALIZAR EL FORMULARIO
                    mForm.Freeze(True)
                    rSboApp.ActivateMenuItem("1289")
                    rSboApp.ActivateMenuItem("1288")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("error generar evento de refrescar LQ" + ex.Message.ToString, "ManejoDeDocumentos")
                Finally
                    mForm.Freeze(False)
                End Try

            ElseIf _btnAccGR.Caption = "(GS) Reenviar MAIL" Then
                mForm.Freeze(True)
                Try
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Re Enviando Mail... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Dim ClaveAcceso As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_CLAVE", 0)
                    Dim docentry As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0)
                    Dim SQUERY As String = ""
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", oTabla), "IDENTIFICADOR", docentry)
                        'SQUERY = String.Format("SELECT A.""E_Mail"" FROM ""{0}"" O INNER JOIN ""OCRD"" A ON O.""CardCode"" = A.""CardCode"" WHERE O.""DocEntry"" =  {1}", oTabla, docentry)
                    Else
                        SQUERY = Replace(Replace(Functions.VariablesGlobales._vgQueryCorreo, "TABLA", oTabla), "IDENTIFICADOR", docentry)
                        'SQUERY = String.Format("SELECT A.E_Mail FROM {0} O WITH(NOLOCK) INNER JOIN OCRD A WITH(NOLOCK) ON O.CardCode = A.CardCode WHERE O.DocEntry = {1} ", oTabla, docentry)
                    End If
                    Dim sCorreoNuevo As String = oFuncionesB1.getRSvalue(SQUERY, "Email", "")
                    If String.IsNullOrEmpty(sCorreoNuevo) Then
                        rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro Email, verificar consulta..!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    Else
                        Try
                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                'oManejoDocumentosEcua.ReenvioMail(sCorreoNuevo, ClaveAcceso)
                            Else
                                oManejoDocumentos.ReenvioMail(sCorreoNuevo, ClaveAcceso)
                            End If

                            rSboApp.SetStatusBarMessage(NombreAddon + " - Mail Re Enviado, Listo!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            mForm.Freeze(False)
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("error al reenviar el e-mail: " + ex.Message.ToString, "ManejoDeDocumentos")
                        End Try
                    End If
                Catch ex As Exception
                    mForm.Freeze(False)
                End Try
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub rSboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rSboApp.FormDataEvent
        Try  'SOLO INGRESO PARA LOS SIGUIENTES FORMULARIOS ,el 1250000940 es solicitud de traslado
            'If BusinessObjectInfo.FormTypeEx = "85" Then
            '    Select Case BusinessObjectInfo.EventType
            '        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
            '            If Not BusinessObjectInfo.BeforeAction Then
            '                Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
            '                AgregaBotonLP(mForm)
            '            End If
            '    End Select
            'End If



            If BusinessObjectInfo.FormTypeEx = "133" Or
                BusinessObjectInfo.FormTypeEx = "60090" Or
                BusinessObjectInfo.FormTypeEx = "60091" Or
                BusinessObjectInfo.FormTypeEx = "60092" Or
                BusinessObjectInfo.FormTypeEx = "65303" Or
                BusinessObjectInfo.FormTypeEx = "65307" Or
                BusinessObjectInfo.FormTypeEx = "65300" Or
                BusinessObjectInfo.FormTypeEx = "65301" Or
                BusinessObjectInfo.FormTypeEx = "179" Or
                BusinessObjectInfo.FormTypeEx = "140" Or
                BusinessObjectInfo.FormTypeEx = "940" Or
                BusinessObjectInfo.FormTypeEx = "1250000940" Or
                BusinessObjectInfo.FormTypeEx = "65306" Or
                BusinessObjectInfo.FormTypeEx = "141" Or
                BusinessObjectInfo.FormTypeEx = "720" Then


                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If Not BusinessObjectInfo.BeforeAction Then
                            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
                            If Not nombreFormulario.Contains("ELECTRONICO") Then
                                nombreFormulario = mForm.Title
                            End If
                            Try
                                Dim oFormMode As Integer = 1
                                oFormMode = mForm.Mode
                                If BusinessObjectInfo.FormTypeEx = "65307" Then
                                    Functions.VariablesGlobales._SS_FacturaExportacion = "SI"
                                Else
                                    Functions.VariablesGlobales._SS_FacturaExportacion = "NO"
                                End If
                                SeteaTipoTabla_FormTypeEx(BusinessObjectInfo.FormTypeEx)

                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                    CargaItemEnFormularioEcuanexus(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                Else

                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                                        creaBotonPrueba_Factura_GuiaRemision(mForm)
                                        Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = ""
                                        Functions.VariablesGlobales._FacturaGuiaRemision = ""
                                        If (BusinessObjectInfo.FormTypeEx = "133" Or BusinessObjectInfo.FormTypeEx = "720") And ValidarGuiaEnFacturaHeison(mForm, BusinessObjectInfo.FormTypeEx) Then

                                            CargaItemEnFormulario_Factura_GuiaRemision(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                        Else
                                            CargaItemEnFormulario_Factura_GuiaRemision(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                        End If
                                    End If
                                    CargaItemEnFormulario(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla, BusinessObjectInfo.FormUID)
                                End If

                                If BusinessObjectInfo.FormTypeEx = "141" Or BusinessObjectInfo.FormTypeEx = "60092" Then
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        CargaItemEnFormulario_LiquidacionCompraEcuanexus(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                    Else
                                        CargaItemEnFormulario_LiquidacionCompra(mForm, "ACTUALIZAR", mForm.TypeEx, oTabla)
                                    End If

                                End If
                                mForm.Mode = oFormMode

                                'se coloco esta validacion dentro de la funcion cargaritemformulario pero dejando el boton visiblw false 
                                ' ya que dentro no lo bloquea por motivo de que al actualizar informacion se vuelve habilitar el boton
                                'If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Then
                                '    Dim estado As String = cbEstado.Selected.Value
                                '    Dim _btnAccion As SAPbouiCOM.ButtonCombo = mForm.Items.Item("btnAccion").Specific
                                '    If _btnAccion.Caption = "(GS) Reenviar SRI" And estado = "0" Then
                                '        _btnAccion.Item.Visible = False
                                '        ' _btnAccion.Item.AffectsFormMode = False

                                '    End If
                                '    If BusinessObjectInfo.FormTypeEx = "141" Then
                                '        Dim estadoLQ As String = cbEstadoLQ.Selected.Value
                                '        Dim _btnAccionLQ As SAPbouiCOM.ButtonCombo = mForm.Items.Item("btnAccLQ").Specific
                                '        If _btnAccionLQ.Caption = "(GS) Reenviar SRI" And estadoLQ = "0" Then
                                '            _btnAccionLQ.Item.Enabled = False
                                '        End If
                                '    End If

                                'End If


                                'mForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                ''mForm.Update()
                            Catch ex As Exception
                                rSboApp.SetStatusBarMessage("et_FORM_DATA_LOAD", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            End Try

                            'If mForm.TypeEx = "133" Then
                            'End If
                        End If
                        'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        '    If Not BusinessObjectInfo.BeforeAction Then
                        '        Select Case BusinessObjectInfo.ActionSuccess
                        '            Case True
                        '                'oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                        '                'agregar funcion parametros= id del formulario
                        '                'If oManejoDocumentos.AnularDocumento(BusinessObjectInfo.FormTypeEx) Then
                        '                Dim numeracion As String
                        '                Dim RetValAnu As Long
                        '                Dim ErrCodeAnu As Long
                        '                Dim ErrMsgAnu As String = ""

                        '                If BusinessObjectInfo.FormTypeEx = "133" Or BusinessObjectInfo.FormTypeEx = "60090" Or BusinessObjectInfo.FormTypeEx = "60091" Then
                        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        '                ElseIf BusinessObjectInfo.FormTypeEx = "65303" Then
                        '                    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                        '                    oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                        '                ElseIf BusinessObjectInfo.FormTypeEx = "179" Then
                        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                        '                    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                        '                ElseIf BusinessObjectInfo.FormTypeEx = "140" Then
                        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                        '                    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes
                        '                ElseIf BusinessObjectInfo.FormTypeEx = "940" Then
                        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        '                    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
                        '                ElseIf BusinessObjectInfo.FormTypeEx = "141" Then
                        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        '                ElseIf BusinessObjectInfo.FormTypeEx = "65306" Then
                        '                    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        '                    oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
                        '                End If
                        '                oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                        '                If oDocumento.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                        '                    Dim est As String = ""
                        '                    Dim pemi As String = ""
                        '                    Dim folio As String = ""

                        '                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                        '                        est = oDocumento.UserFields.Fields.Item("U_SS_Est").Value
                        '                        pemi = oDocumento.UserFields.Fields.Item("U_SS_Pemi").Value
                        '                        folio = oDocumento.FolioNumber.ToString
                        '                    ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                        '                        est = oDocumento.UserFields.Fields.Item("U_SER_EST").Value
                        '                        pemi = oDocumento.UserFields.Fields.Item("U_SER_PE").Value
                        '                        folio = oDocumento.FolioNumber.ToString
                        '                        'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                        '                        '    est = oDocumento.UserFields.Fields.Item("U_SYP_SERIESUC").Value
                        '                        '    pemi = oDocumento.UserFields.Fields.Item("U_SYP_MDSD").Value
                        '                        '    folio = oDocumento.UserFields.Fields.Item("U_SYP_MDCD").Value
                        '                        'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                        '                        '    est = oDocumento.UserFields.Fields.Item("U_SS_Est").Value
                        '                        '    pemi = oDocumento.UserFields.Fields.Item("U_SS_Pemi").Value
                        '                        '    folio = oDocumento.FolioNumber.ToString
                        '                        'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                        '                        '    est = oDocumento.UserFields.Fields.Item("U_N_DOCSRI").Value
                        '                        '    est = est.Substring(1, 3)
                        '                        '    pemi = oDocumento.UserFields.Fields.Item("U_N_DOCSRI").Value
                        '                        '    est = est.Substring(4, 3)
                        '                        '    folio = Right(oDocumento.UserFields.Fields.Item("U_N_DOCSRI").Value, 9)
                        '                    End If
                        '                    If Not folio.Length.Equals("9") Then
                        '                        folio = folio.PadLeft(9, "0")
                        '                    End If

                        '                    numeracion = est + "-" + pemi + "-" + folio
                        '                    oDocumento.FolioNumber = ""
                        '                    oDocumento.NumAtCard = numeracion
                        '                    RetValAnu = oDocumento.Update()

                        '                    If RetValAnu <> 0 Then
                        '                        rCompany.GetLastError(ErrCodeAnu, ErrMsgAnu)
                        '                        rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualzar campo NumAtCard..!! - " + ErrMsgAnu.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        '                    Else
                        '                        rSboApp.SetStatusBarMessage(NombreAddon + " - Campo NumAtCard actualizo con éxito..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        '                    End If
                        '                End If


                        '        End Select


                        '    End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If Not BusinessObjectInfo.BeforeAction Then
                            Select Case BusinessObjectInfo.ActionSuccess
                                Case True

                                    If BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                                        rSboApp.SetStatusBarMessage(NombreAddon + " - El documento creado es un documento Preliminar, no se procesará..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    Else
                                        Try
                                            Dim form As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)

                                            Dim s As String = BusinessObjectInfo.ObjectKey.ToString()

                                            ''If EsElectronico = "" Then
                                            ''    Dim feLQ As String = ""
                                            ''    feLQ = "SELECT ""U_IdSerie""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = '" + oDocumento.Series
                                            ''    'SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + Series + "' ", "Code", "")
                                            ''    Dim fe As String = oFuncionesB1.getRSvalue(feLQ, "U_IdSerie", "")
                                            ''    If Not fe = "" Then
                                            ''        EsElectronico = "FE"
                                            ''    End If
                                            ''End If
                                            'If EsElectronico = "" Then
                                            '    Dim oSerie As String
                                            '    oSerie = oForm.Items.Item("88").Specific.value.ToString()
                                            '    Dim SQL As String = ""
                                            '    SQL = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + oSerie
                                            '    Dim esElectronico = oFuncionesB1.getRSvalue(SQL, "Code", "")
                                            '    If Not String.IsNullOrEmpty(esElectronico) Then
                                            '        esElectronico = "FE"
                                            '    Else
                                            '        LQEsElectronico = ""
                                            '    End If
                                            'End If
                                            If BusinessObjectInfo.FormTypeEx = "141" Or BusinessObjectInfo.FormTypeEx = "60092" Then

                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

                                                If ValidarDocSerieElectronica("LQE", oDocumento.Series, BusinessObjectInfo.FormTypeEx, form) Then



                                                    If LQEsElectronico = "FE" Then
                                                        oTipoTabla = "LQE"
                                                        Dim RetVal As Long
                                                        Dim ErrCode As Long
                                                        Dim ErrMsg As String
                                                        Dim queryestadoLQ As String = ""
                                                        Dim estadoLQ As String = ""
                                                        Dim LQexterna As String
                                                        'oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                        'oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                        If Not oDocumento.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation Then
                                                            If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                                                                'validar si efectivamente es LQ electronica
                                                                'DM 2023--03-01 Se comento la seccion siguiente debido a que no es necesaria ya que al momento de crear el documento se verifica si es electronica o no, en caso de no serlo no realizara ningun proceso
                                                                'If Functions.VariablesGlobales._vgSerieUDF = "Y" Then
                                                                '    Dim SerieDocUDF As String = LTrim(RTrim(form.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0)))
                                                                '    If SerieDocUDF = "02" Or SerieDocUDF = "03" Then
                                                                '        Utilitario.Util_Log.Escribir_Log("Verificacion LQ electronica UDF: " + SerieDocUDF.ToString(), "ManejoDeDocumentos")
                                                                '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Verificacion LQ electronica UDF: " + SerieDocUDF.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                '    Else
                                                                '        Utilitario.Util_Log.Escribir_Log("vLQ NO Electronica UDF: " + SerieDocUDF.ToString(), "ManejoDeDocumentos")
                                                                '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "LQ NO Electronica UDF: " + SerieDocUDF.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                '        Exit Sub
                                                                '    End If
                                                                'Else

                                                                '    Dim oSerie As String = ""
                                                                '    oSerie = LTrim(RTrim(form.DataSources.DBDataSources.Item(oTabla).GetValue("Series", 0)))
                                                                '    Dim SQL As String = ""
                                                                '    Dim LQELEC As String = ""
                                                                '    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                                '        SQL = "SELECT ""U_TipoD"" FROM ""@SS_SERD"" where ""U_SerId"" = " + oSerie
                                                                '        LQELEC = oFuncionesB1.getRSvalue(SQL, "U_TipoD", "")
                                                                '        If LQELEC = "LQ" Or LQELEC = "LQRT" Then
                                                                '            LQELEC = "FE"
                                                                '        Else

                                                                '            LQELEC = ""
                                                                '        End If
                                                                '    Else
                                                                '        SQL = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + oSerie
                                                                '        LQELEC = oFuncionesB1.getRSvalue(SQL, "Code", "")
                                                                '    End If

                                                                '    If Not String.IsNullOrEmpty(LQELEC) Then
                                                                '        Utilitario.Util_Log.Escribir_Log("Verificacion LQ electronica Serie: query " + SQL.ToString() + " Repuesta: " + LQELEC.ToString, "ManejoDeDocumentos")
                                                                '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Verificacion LQ electronica Serie: query " + SQL.ToString() + " Repuesta: " + LQELEC.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                '    Else
                                                                '        Utilitario.Util_Log.Escribir_Log("Serie LQ NO Electronica : query " + SQL.ToString() + " Repuesta: " + LQELEC.ToString, "ManejoDeDocumentos")
                                                                '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Serie LQ NO Electronica : query " + SQL.ToString() + " Repuesta: " + LQELEC.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                '        Exit Sub
                                                                '    End If
                                                                'End If

                                                                'queryestadoLQ = oFuncionesB1.getRSvalue("SELECT ""U_LQ_ESTADO"" FROM ""OPCH"" WHERE ""DocEntry"" = '" + oDocumento.DocEntry.ToString() + "' ", "U_LQ_ESTADO", "")
                                                                queryestadoLQ = LTrim(RTrim(form.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_ESTADO", 0)))
                                                                Utilitario.Util_Log.Escribir_Log("estado del documento: " + queryestadoLQ.ToString, "ManejoDeDocumentos")
                                                                'If queryestadoLQ = "0" Or queryestadoLQ = "" Then 'se comento debido a que cuando se duplica y no hay proceso que cambie el estado no se procesa 04012023
                                                                Utilitario.Util_Log.Escribir_Log("estado del documento 1: " + queryestadoLQ.ToString, "ManejoDeDocumentos")
                                                                'Dim folioLQ As String = oFuncionesB1.getRSvalue("SELECT ""FolioNum"" FROM ""OPCH"" WHERE ""DocEntry"" = '" + oDocumento.DocEntry.ToString() + "' ", "FolioNum", "")
                                                                Dim folioLQ As String = form.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0)
                                                                Utilitario.Util_Log.Escribir_Log("folio del documento: " + folioLQ.ToString, "ManejoDeDocumentos")
                                                                'folioLQ = form.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0)
                                                                If folioLQ = "0" Or folioLQ = "" Then
                                                                    'Utilitario.Util_Log.Escribir_Log("folio del documento 1: " + folioLQ.ToString, "ManejoDeDocumentos")
                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Liquidación de Compra Electronica...!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                    Dim SerieExcluir As String = "N"
                                                                    SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + oDocumento.Series.ToString() + "' ", "Code", "")
                                                                    If String.IsNullOrEmpty(SerieExcluir) Then
                                                                        SerieExcluir = "N"
                                                                    Else
                                                                        SerieExcluir = "Y"
                                                                    End If
                                                                    Utilitario.Util_Log.Escribir_Log("Verificar Serie Excluir LQ: " + SerieExcluir, "ManejoDeDocumentos")

                                                                    If Functions.VariablesGlobales._FoliacionPostin = "Y" Then

                                                                        SerieExcluir = "Y"
                                                                        Utilitario.Util_Log.Escribir_Log("Foliacion LQ: " + SerieExcluir, "ManejoDeDocumentos")

                                                                    End If
                                                                    If SerieExcluir = "Y" And (Functions.VariablesGlobales._vgSerieUDF = "N" Or Functions.VariablesGlobales._vgSerieUDF = "Y") Then
                                                                        'If SerieExcluir = "Y" And (Functions.VariablesGlobales._vgSerieUDF = "N" And Functions.VariablesGlobales._vgSerieUDF = "Y") Then 'esta linea se activa cuando el cliente es ENTRIX
                                                                        Dim EnviaDocumentosEnBackGround As String = "N"
                                                                        'EnviaDocumentosEnBackGround = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EnviaDocumentosEnBackGround")
                                                                        EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung
                                                                        ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA
                                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                            Utilitario.Util_Log.Escribir_Log("EnviaDocumentosEnBackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                                                                        Else
                                                                            Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        End If
                                                                        'oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                    Else
                                                                        Dim sSQL As String = ""
                                                                        If Functions.VariablesGlobales._vgFolioLQUDF = "Y" Then

                                                                            Dim est As String = "00"
                                                                            Dim typeEx = "", idForm As String = ""
                                                                            typeEx = oFuncionesB1.FormularioActivo(idForm)
                                                                            If typeEx = "141" Then
                                                                                'Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(pVal.FormUID)
                                                                                Dim count As Integer = 0
                                                                                Dim oForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)
                                                                                'Dim oFrmUser As SAPbouiCOM.Form
                                                                                'Try
                                                                                '    oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                                                                                'Catch ex As Exception
                                                                                '    rSboApp.SendKeys("^+U")
                                                                                '    oFrmUser = rSboApp.Forms.GetForm("-" + typeEx, count)
                                                                                'End Try
                                                                                est = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_Establecimiento", 0).Trim
                                                                                sSQL = "SELECT ISNULL(""U_Sec"",'0') AS U_ULT_SECUEN, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_Est"" = " + est.ToString
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener Secuencial LQ: " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                'est = oFrmUser.Items.Item("U_Establecimiento").Specific
                                                                            End If
                                                                        Else
                                                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                                                                sSQL = Functions.VariablesGlobales._ConsultaFolioSS
                                                                                If sSQL = "" Then
                                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Por favor agregar consulta para asignacion de Folio", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                                    Exit Sub
                                                                                End If

                                                                                sSQL = sSQL.Replace("TABLA", oTabla)
                                                                                sSQL = sSQL.Replace("IDENTIFICADOR", oDocumento.DocEntry)
                                                                                sSQL = sSQL.Replace("TIPDOCE", "('LQ','LQRT')")

                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener Secuencial : " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "Secuencial")

                                                                                sCode = oFuncionesB1.getRSvalue(sSQL, "TipoDoc")

                                                                                If sFolio = "" Then
                                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro campo Secuencial en la consulta, por favor verificar..!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                                    Exit Sub
                                                                                Else
                                                                                    numFolio = sFolio
                                                                                End If

                                                                            Else
                                                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                                    sSQL = "SELECT IFNULL(""U_Sec"",'0') AS U_ULT_SECUEN, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                                                                Else
                                                                                    sSQL = "SELECT ISNULL(""U_Sec"",'0') AS U_ULT_SECUEN, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                                                                End If
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener Secuencial : " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                                                    LQexterna = oDocumento.UserFields.Fields.Item("U_DOC_DECLARABLE").Value

                                                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Parametro activo de Desactivar Pregunta Declarable, se envia igual al SRI..  ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                    'ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                                                    '    LQexterna = oDocumento.UserFields.Fields.Item("U_SS_Declarable").Value
                                                                                    '    If LQexterna = "NO" Then
                                                                                    '        LQexterna = "N"
                                                                                    '    End If
                                                                                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Parametro activo de Desactivar Pregunta Declarable, se envia igual al SRI..  ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                Else
                                                                                    LQexterna = "S"

                                                                                End If
                                                                                If LQexterna = "N" Then
                                                                                    Dim iReturnValue As Integer
                                                                                    iReturnValue = rSboApp.MessageBox(NombreAddon + " - El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?", 1, "&SI", "&NO")
                                                                                    If iReturnValue = 2 Then
                                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Respuesta igual NO, no se envió el documento al SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : NO", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                        Exit Sub
                                                                                    Else
                                                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : SI", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                    End If
                                                                                End If
                                                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "U_ULT_SECUEN")
                                                                                Utilitario.Util_Log.Escribir_Log("Obteniendo Ultima secuencia: query" + sSQL.ToString + "Resultado" + sFolio.ToString, "ManejoDeDocumentos")
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se obtuvo el siguiente # de Folio : " + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                If sFolio = "" Then
                                                                                    sFolio = 0
                                                                                End If
                                                                                sCode = oFuncionesB1.getRSvalue(sSQL, "Code")
                                                                                If sCode = "" Then
                                                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Al NO Obtenerse el Code, implica que no esta registrado en la tabla Liquidación de Compra la serie del documento : " + oDocumento.Series.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se Omitió el envío al SRI, BubbleEvent = False", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                    BubbleEvent = False
                                                                                    Exit Sub
                                                                                End If

                                                                                numFolio = Integer.Parse(sFolio) + 1
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se sumo 1 al # de Folio: " + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                                                                            End If

                                                                        End If

                                                                        If Functions.VariablesGlobales._vgFolioLQUDF = "Y" Then
                                                                            oDocumento.UserFields.Fields.Item("U_BPP_MDCD").Value = numFolio.ToString.PadLeft(9, "0")
                                                                        Else
                                                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                                                Dim _est As String = ""
                                                                                Dim Est As String = ""
                                                                                Dim _puntoemi As String = ""
                                                                                Dim PuntoEmi As String = ""
                                                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                                                    _est = "SELECT IFNULL(""U_Estable"",'0') AS Establecimiento, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                                                                    _puntoemi = "SELECT IFNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code"" FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                                                                Else
                                                                                    _est = "SELECT ISNULL(""U_Estable"",'0') AS Establecimiento, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                                                                    _puntoemi = "SELECT ISNULL(""U_PtoEmi"",'0') AS PuntoEmision, ""Code""  FROM ""@GS_LIQUI"" WHERE ""U_IdSerie"" = " + oDocumento.Series.ToString
                                                                                End If
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener establecimiento : " + _est.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener punto de emision : " + _puntoemi.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                Est = oFuncionesB1.getRSvalue(_est, "Establecimiento")
                                                                                PuntoEmi = oFuncionesB1.getRSvalue(_puntoemi, "PuntoEmision")
                                                                                'U_N_DOCSRI
                                                                                oDocumento.UserFields.Fields.Item("U_N_DOCSRI").Value = Est.ToString + PuntoEmi.ToString + numFolio.ToString.PadLeft(9, "0")
                                                                            Else
                                                                                oDocumento.FolioNumber = numFolio
                                                                                oDocumento.FolioPrefixString = oTipoTabla
                                                                            End If

                                                                        End If

                                                                        '"/*Prefijo del folio*/"
                                                                        RetVal = oDocumento.Update()
                                                                        If RetVal <> 0 Then

                                                                            rCompany.GetLastError(ErrCode, ErrMsg)

                                                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar el numero de folio a la Liquidación de Compra..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al asignar el numero de folio a la Liquidación de Compra..!!" + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            Exit Sub
                                                                        Else
                                                                            Dim EnviaDocumentosEnBackGround As String = "N"
                                                                            EnviaDocumentosEnBackGround = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EnviaDocumentosEnBackGround")
                                                                            EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung

                                                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                                                                                If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                        oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                    Else
                                                                                        oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                    End If

                                                                                    Utilitario.Util_Log.Escribir_Log("Proceso terminado envio de LQ", "ManejoDeDocumentos")
                                                                                    'Exit Sub

                                                                                Else
                                                                                    rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                                                End If
                                                                            Else
                                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en tabla de Liquidación de Compra..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                                                If oFuncionesAddon.ActualizaSecuencia_LiquidacionDeCompra(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then

                                                                                    Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")

                                                                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                        Else
                                                                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                        End If

                                                                                        Utilitario.Util_Log.Escribir_Log("Proceso terminado envio de LQ", "ManejoDeDocumentos")
                                                                                        'Exit Sub

                                                                                    Else
                                                                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                                    End If

                                                                                End If
                                                                            End If

                                                                        End If

                                                                    End If
                                                                Else
                                                                    Dim SerieExcluir As String = "N"
                                                                    Utilitario.Util_Log.Escribir_Log("Ya obtiene folio:" + folioLQ.ToString, "ManejoDeDocumentos")

                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Liquidación de Compra Electronica...!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                                                    SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + oDocumento.Series.ToString() + "' ", "Code", "")
                                                                    If String.IsNullOrEmpty(SerieExcluir) Then
                                                                        SerieExcluir = "N"
                                                                    Else
                                                                        SerieExcluir = "Y"
                                                                    End If
                                                                    Utilitario.Util_Log.Escribir_Log("Verificar Serie Excluir LQ: " + SerieExcluir, "ManejoDeDocumentos")


                                                                    If Functions.VariablesGlobales._FoliacionPostin = "Y" Then
                                                                        SerieExcluir = "Y"
                                                                    End If

                                                                    If SerieExcluir = "Y" Then

                                                                        Dim EnviaDocumentosEnBackGround As String = "N"
                                                                        EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung

                                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                            Utilitario.Util_Log.Escribir_Log("EnviaDocumentosEnBackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                                                                        Else
                                                                            Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        End If

                                                                    End If
                                                                End If


                                                                'Else
                                                                '    Utilitario.Util_Log.Escribir_Log("Ya obtiene ESTADO: " + queryestadoLQ.ToString, "ManejoDeDocumentos")
                                                                '    Exit Sub
                                                                'End If



                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    Utilitario.Util_Log.Escribir_Log("Serie LQ no electronica", "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                                                    oDocumento = Nothing
                                                    'Exit Sub
                                                End If

                                            End If

                                            oDocumento = Nothing
                                            'If EsElectronico = "FE" Then

                                            If BusinessObjectInfo.FormTypeEx = "133" Or BusinessObjectInfo.FormTypeEx = "60090" Or BusinessObjectInfo.FormTypeEx = "65307" Then  ' FACTURA DE CLIENTE - FACTURA DEUDOR + PAGO - FACTURA DE EXPORTACION
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                                oTipoTabla = "FCE"
                                                If BusinessObjectInfo.FormTypeEx = "65307" Then
                                                    Functions.VariablesGlobales._SS_FacturaExportacion = "SI"
                                                Else
                                                    Functions.VariablesGlobales._SS_FacturaExportacion = "NO"
                                                End If
                                            ElseIf BusinessObjectInfo.FormTypeEx = "60091" Then ' FACTURA DE RESERVA
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                                oTipoTabla = "FRE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "65303" Then ' NOTA DE DEBITO
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                                oTipoTabla = "NDE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "65300" Then ''FACTURA DE ANTICIPO DE CLIENTES
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                                                oTipoTabla = "FAE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "179" Then 'NOTA DE CREDITO DE CLIENTES
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                                oTipoTabla = "NCE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "140" Then 'GUIA DE REMISION - ENTREGA
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                                oTipoTabla = "GRE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "940" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
                                                oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                                                oTipoTabla = "TRE"

                                            ElseIf BusinessObjectInfo.FormTypeEx = "1250000940" Then 'GUIA DE REMISION - SOLICITUD TRANSLADO                                            
                                                oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
                                                oTipoTabla = "TLE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "141" Then  'FACTURA DE PROVEEDOR/RETENCION                             
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                oTipoTabla = "REE"
                                            ElseIf BusinessObjectInfo.FormTypeEx = "65306" Then  'NOTA DE DEBITO PROVEEDOR/RETENCION                             
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                oTipoTabla = "RDM"
                                                Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
                                            ElseIf BusinessObjectInfo.FormTypeEx = "65301" Then  'FACTURA DE ANTICIPO DE PROVEEDOR/RETENCION                             
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                oTipoTabla = "REA"
                                                Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
                                            ElseIf BusinessObjectInfo.FormTypeEx = "60092" Then  'FACTURA DE RESERVA PROVEEDOR/RETENCION                           
                                                oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                                oTipoTabla = "RER"
                                                Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
                                            Else
                                                Exit Sub
                                            End If

                                            If oTipoTabla = "TRE" Then
                                                oTransferencia.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                If ValidarDocSerieElectronica(oTipoTabla, oTransferencia.Series) = False Then
                                                    oDocumento = Nothing
                                                    Exit Sub
                                                End If

                                            ElseIf oTipoTabla = "TLE" Then
                                                oTransferencia.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                If ValidarDocSerieElectronica(oTipoTabla, oTransferencia.Series) = False Then
                                                    oDocumento = Nothing
                                                    Exit Sub
                                                End If
                                                'ElseIf oTipoTabla <> "REE" And oTipoTabla <> "RDM" And oTipoTabla <> "REA" And oTipoTabla <> "RER" Then
                                            Else
                                                oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                If ValidarDocSerieElectronica(oTipoTabla, oDocumento.Series, BusinessObjectInfo.FormTypeEx, form) = False Then
                                                    oDocumento = Nothing
                                                    Exit Sub
                                                End If

                                                'ElseIf (oTipoTabla = "REE" Or oTipoTabla = "RDM" Or oTipoTabla = "REA" Or oTipoTabla = "RER") And Nombre_Proveedor_SAP_BO <> Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                '    'Else
                                                '    oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                '    If ValidarDocSerieElectronica(oTipoTabla, oDocumento.Series, BusinessObjectInfo.FormTypeEx, form) = False Then
                                                '        oDocumento = Nothing
                                                '        Exit Sub
                                                '    End If
                                            End If


                                            If EsElectronico = "FE" Then
                                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then

                                                    'VALIDAR SI EFECTIVAMENTE ES ELECTRONICO 
                                                    'DM 2023--03-01 Se comento la seccion siguiente debido a que no es necesaria ya que al momento de crear el documento se verifica si es electronica o no, en caso de no serlo no realizara ningun proceso
                                                    'If ValidarSerieDocElec(BusinessObjectInfo) = False Then
                                                    '    If oTipoTabla = "TRE" Then
                                                    '        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Finaliza proceso debido a que la serie no es electronica", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    '    Else
                                                    '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Finaliza proceso debido a que la serie no es electronica", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    '    End If

                                                    '    Exit Sub
                                                    'End If
                                                    'FIN VALIDACION SI SERIE ES ELECTRONICA

                                                    If oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then
                                                        Try
                                                            'oTransferencia.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Obteniendo información del objeto creado", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                            If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                ' PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA
                                                                Dim EnviaDocumentosEnBackGround As String = "N"
                                                                'EnviaDocumentosEnBackGround = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EnviaDocumentosEnBackGround")
                                                                EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung
                                                                Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")

                                                                If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                        oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                    Else
                                                                        oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                    End If

                                                                Else
                                                                    rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                End If

                                                            End If

                                                        Catch ex As Exception ' CUANDO PIDE AUTORIZACION SE CAE AQUI
                                                        End Try
                                                    Else
                                                        'oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Obteniendo información del objeto creado", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        If Not oDocumento.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation Then ' SI EL DOCUMENTO NO ES UNA CANCELACION
                                                            'If Not oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then                                               
                                                            If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                                                                Dim estLQ As String = oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value
                                                                If Functions.VariablesGlobales._vgNoEnviarRT = "Y" Then
                                                                    Utilitario.Util_Log.Escribir_Log("Paramero activo no reenviar sri rt: " + Functions.VariablesGlobales._vgNoEnviarRT.ToString, "EventosEmision")
                                                                    If LQEsElectronico = "FE" And (estLQ = "3" Or estLQ = "6" Or estLQ = "4") Then
                                                                        Utilitario.Util_Log.Escribir_Log("LQEsElectronico: " + LQEsElectronico + " Estado: " + estLQ.ToString, "EventosEmision")
                                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - No se emitirá la Retención hasta a que la Liquidación se encuentre Autorizada", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                        Exit Sub
                                                                    End If
                                                                End If


                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                Try

                                                                    ' PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA
                                                                    Dim EnviaDocumentosEnBackGround As String = "N"
                                                                    'EnviaDocumentosEnBackGround = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EnviaDocumentosEnBackGround")
                                                                    EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung
                                                                    Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
                                                                    ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA

                                                                    If oTipoTabla = "REE" Or
                                                                        oTipoTabla = "REA" Or
                                                                        oTipoTabla = "RDM" Or
                                                                        oTipoTabla = "RER" Then

                                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            Else
                                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                            End If

                                                                            Utilitario.Util_Log.Escribir_Log("Tipo Tabla envio docuemnto SRI: " + oTipoTabla.ToString, "ManejoDeDocumentos")
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                        End If

                                                                    Else ' LOGICA CUANDO NO ES UNA RETENCION

                                                                        ' INCIALIZO LOS CAMPOS DE ADDON QUE GUARDAN LA AUTORIZACIÓN
                                                                        If oFuncionesB1.checkCampoBD("OINV", "NUM_AUTO_FAC") Then
                                                                            oDocumento.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = ""
                                                                        End If
                                                                        If oFuncionesB1.checkCampoBD("OINV", "CLAVE_ACCESO") Then
                                                                            oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = ""
                                                                        End If
                                                                        If oFuncionesB1.checkCampoBD("OINV", "ESTADO_AUTORIZACIO") Then
                                                                            oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = "0"
                                                                        End If
                                                                        If oFuncionesB1.checkCampoBD("OINV", "OBSERVACION_FACT") Then
                                                                            oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = ""
                                                                        End If
                                                                        ' END INCIALIZO LOS CAMPOS DE ADDON QUE GUARDAN LA AUTORIZACIÓN

                                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                                            If oTipoTabla = "TRE" Then
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                                                End If

                                                                            Else
                                                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                    'ElseIf Functions.VariablesGlobales._ApiAutSS = "Y" Then
                                                                                    '    oManejoDocumentosSolsap.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                Else
                                                                                    oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                                                End If

                                                                            End If
                                                                        Else
                                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                            If oTipoTabla = "TRE" Then
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            Else
                                                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                            End If

                                                                        End If

                                                                    End If

                                                                Catch ex As Exception
                                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Error: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error, Catch: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                                    Exit Sub
                                                                End Try

                                                            End If


                                                        End If
                                                    End If



                                                Else

                                                    Try
                                                        Procesa_EXXIS_ONESOLUTIONS_SYPSOFT(BusinessObjectInfo, BubbleEvent)

                                                    Catch ex As Exception
                                                        Utilitario.Util_Log.Escribir_Log("Ingresando por try Procesa_EXXIS_ONESOLUTIONS_SYPSOFT  ", "ManejoDeDocumentos")
                                                        BubbleEvent = False
                                                    End Try


                                                    ' EMPIEZA EL PROCESO DE LIQUIDACIO DE COMPRA


                                                End If

                                            End If


                                        Catch ex As Exception
                                            Utilitario.Util_Log.Escribir_Log("Error: " + ex.Message.ToString, "ManejoDeDocumentos")
                                        Finally
                                            Utilitario.Util_Log.Escribir_Log("Ingresando por finally: ", "ManejoDeDocumentos")
                                            BubbleEvent = False
                                        End Try
                                    End If
                            End Select

                        End If

                End Select

            End If
        Catch ex As Exception
            rSboApp.MessageBox(NombreAddon + " - Error:" + ex.ToString())



        End Try
    End Sub

    Public Sub serir() ' NO SE ESTA USANDO
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim oSeriesService As SAPbobsCOM.SeriesService
        Dim oSeries As SAPbobsCOM.Series
        Dim oSeriesParams As SAPbobsCOM.SeriesParams

        'get company service
        oCmpSrv = rCompany.GetCompanyService

        'get series service
        oSeriesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)

        'get series params
        oSeriesParams = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeriesParams)

        'set the number of an existing series
        oSeriesParams.Series = 84

        'get the series
        oSeries = oSeriesService.GetSeries(oSeriesParams)

        oSeries.NextNumber = 3
        oSeriesService.UpdateSeries(oSeries)

        'print the series name
        Debug.WriteLine(oSeries.Name)


    End Sub

    Public Sub Procesa_EXXIS_ONESOLUTIONS_SYPSOFT(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)

        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String
            Dim queryestadoRT As String = ""
            ActualizaSecuenciaFolio = False

            ' PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA
            Dim EnviaDocumentosEnBackGround As String = "N"
            'EnviaDocumentosEnBackGround = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "EnviaDocumentosEnBackGround")
            EnviaDocumentosEnBackGround = Functions.VariablesGlobales._EnviarBackGroung
            Utilitario.Util_Log.Escribir_Log("Envía Documentos En BackGround: " + EnviaDocumentosEnBackGround, "ManejoDeDocumentos")
            ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA

            'SIN SON SERIES QUE SE DEBE EXCLUIR PONEMOS TIENE FOLIO = TRUE
            Dim SerieExcluir As String = "N"
            'END SIN SON SERIES QUE SE DEBE EXCLUIR PONEMOS TIENE FOLIO = TRUE

            If oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then 'SOLICITUD DE TraSLADO AGREGADO
                Try
                    oTransferencia.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                    If ValidarSerieDocElec(BusinessObjectInfo) = False Then
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Serie no es electronica", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        Exit Sub
                    End If
                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Obteniendo información del objeto creado", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + oTransferencia.Series.ToString() + "' ", "Code", "")
                    If String.IsNullOrEmpty(SerieExcluir) Then
                        SerieExcluir = "N"
                    End If
                    'SE REVISARA SI ESTA ACTIVADO EL PARAMETRO DE FOLIACION POR POSTN
                    If Functions.VariablesGlobales._FoliacionPostin = "Y" Then
                        SerieExcluir = "S"
                    End If
                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Serie excluir " + SerieExcluir.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                Catch ex As Exception ' CUANDO PIDE AUTORIZACION SE CAE AQUI
                End Try
            Else
                oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)
                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Obteniendo información del objeto creado", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                SerieExcluir = oFuncionesB1.getRSvalue("SELECT ""Code"" FROM ""@GS_SERIEXLUIR"" WHERE ""Code"" = '" + oDocumento.Series.ToString() + "' ", "Code", "")
                If String.IsNullOrEmpty(SerieExcluir) Then
                    SerieExcluir = "N"
                End If

                'SE REVISARA SI ESTA ACTIVADO EL PARAMETRO DE FOLIACION POR POSTN
                If Functions.VariablesGlobales._FoliacionPostin = "Y" Then
                    SerieExcluir = "S"
                End If


            End If

            '******** VERIFICAR SI EL DOC ES ELECTRONICO***************
            'If oTipoTabla <> "TRE" And oTipoTabla <> "TLE" Then
            '    If ValidarSerieDocElec(BusinessObjectInfo) = False Then
            '        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Serie no es electronica", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '        Exit Sub
            '    End If
            'End If

            'Dim form As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
            'Dim odbds As SAPbouiCOM.DBDataSource = CType(form.DataSources.DBDataSources.Item(0), SAPbouiCOM.DBDataSource)
            'Dim SerieDocElec As String = odbds.GetValue("Series", odbds.Offset).Trim
            'If Functions.VariablesGlobales._vgSerieUDF = "Y" And oTipoTabla = "REE" Then
            '    Dim SerieDocUDF = odbds.GetValue("U_DocEmision", odbds.Offset).Trim
            '    If SerieDocUDF = "01" Or SerieDocUDF = "03" Then
            '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Verificacion RT electronica UDF: " + SerieDocUDF.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '    Else
            '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "RT NO Electronica UDF: " + SerieDocUDF.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '        Exit Sub
            '    End If
            'Else

            '    Dim SerieElec As String = ValidarDocElec(SerieDocElec)
            '    If Not String.IsNullOrEmpty(SerieElec) Then
            '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Serie Documento Electronico: SI", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '    Else
            '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Serie Documento NO Electronica: NO", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
            '        Exit Sub
            '    End If
            'End If
            '*************FIN************

            ' **** CONTROL POR CAMPO DE USUARIO "EXTERNA"
            Dim externa As String = "S"
            Dim Procesar As Boolean = True
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If oTipoTabla = "TRE" Then
                    Try
                        externa = oTransferencia.UserFields.Fields.Item("U_DOC_DECLARABLE").Value
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Campo U_DOC_DECLARABLE =  " + externa.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Catch ex As Exception
                        rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Campo Error TRY CATCH  " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End Try
                    If externa = "N" Then

                        ' PARAMETRO PARA DESACTIVAR PREGUNTA POR EL CAMPO U_DOCDECLARABLE Y NO LO ENVIA AL SRI
                        Dim DesactivaPreguntaDeclaraable As String = "N"
                        DesactivaPreguntaDeclaraable = ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "DesactivaPreguntaDeclaraable")
                        Utilitario.Util_Log.Escribir_Log("DesactivaPreguntaDeclaraable: " + DesactivaPreguntaDeclaraable, "ManejoDeDocumentos")
                        ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA

                        If DesactivaPreguntaDeclaraable = "Y" Then
                            Procesar = True
                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Parametro activo de Desactivar Pregunta Declarable, se envia igual al SRI..  ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        Else
                            Dim iReturnValue As Integer
                            iReturnValue = rSboApp.MessageBox(NombreAddon + " - El Documento tiene marcado como EXTERNA/DECLARABLE = 'NO', desea igual enviarlo al SRI?", 1, "&SI", "&NO")
                            If iReturnValue = 2 Then
                                rSboApp.SetStatusBarMessage("Respuesta igual NO, no se envió el documento al SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                Procesar = False
                                BubbleEvent = False
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como EXTERNA/DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : NO", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            Else
                                Procesar = True
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como EXTERNA/DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : SI", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            End If
                        End If
                    End If
                Else
                    Try
                        externa = oDocumento.UserFields.Fields.Item("U_DOC_DECLARABLE").Value
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Campo U_DOC_DECLARABLE =  " + externa.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Catch ex As Exception
                        rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Campo Error TRY CATCH  " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    End Try
                    If externa = "N" Then
                        ' PARAMETRO PARA DESACTIVAR PREGUNTA POR EL CAMPO U_DOCDECLARABLE Y NO LO ENVIA AL SRI
                        Dim DesactivaPreguntaDeclaraable As String = "N"
                        DesactivaPreguntaDeclaraable = ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "DesactivaPreguntaDeclaraable")
                        Utilitario.Util_Log.Escribir_Log("DesactivaPreguntaDeclaraable: " + DesactivaPreguntaDeclaraable, "ManejoDeDocumentos")
                        ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA

                        If DesactivaPreguntaDeclaraable = "Y" Then
                            Procesar = True
                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Parametro activo de Desactivar Pregunta Declarable, se envia igual al SRI..  ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        Else
                            Dim iReturnValue As Integer
                            iReturnValue = rSboApp.MessageBox(NombreAddon + " - El Documento tiene marcado como EXTERNA/DECLARABLE = 'NO', desea igual enviarlo al SRI?", 1, "&SI", "&NO")
                            If iReturnValue = 2 Then
                                rSboApp.SetStatusBarMessage(NombreAddon + " - Respuesta igual NO, no se envió el documento al SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                Procesar = False
                                BubbleEvent = False
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como EXTERNA/DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : NO", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            Else
                                Procesar = True
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como EXTERNA/DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : SI", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            End If
                        End If

                    End If
                End If
            End If
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If oTipoTabla = "TRE" Then
                    'Try
                    '    externa = oTransferencia.UserFields.Fields.Item("U_SS_Declarable").Value
                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Campo U_SS_Declarable =  " + externa.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'Catch ex As Exception
                    '    rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Campo Error TRY CATCH  " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'End Try
                    'If externa = "NO" Then

                    '    ' PARAMETRO PARA DESACTIVAR PREGUNTA POR EL CAMPO U_DOCDECLARABLE Y NO LO ENVIA AL SRI
                    '    Dim DesactivaPreguntaDeclaraable As String = "N"
                    '    DesactivaPreguntaDeclaraable = ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "DesactivaPreguntaDeclaraable")
                    '    Utilitario.Util_Log.Escribir_Log("DesactivaPreguntaDeclaraable: " + DesactivaPreguntaDeclaraable, "ManejoDeDocumentos")
                    '    ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA

                    '    If DesactivaPreguntaDeclaraable = "Y" Then
                    '        Procesar = True
                    '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Parametro activo de Desactivar Pregunta Declarable, se envia igual al SRI..  ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    '    Else
                    '        Dim iReturnValue As Integer
                    '        iReturnValue = rSboApp.MessageBox(NombreAddon + " - El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?", 1, "&SI", "&NO")
                    '        If iReturnValue = 2 Then
                    '            rSboApp.SetStatusBarMessage("Respuesta igual NO, no se envió el documento al SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '            Procesar = False
                    '            BubbleEvent = False
                    '            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : NO", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    '        Else
                    '            Procesar = True
                    '            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : SI", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    '        End If
                    '    End If
                    'End If
                    Procesar = True
                Else
                    'Try
                    '    externa = oDocumento.UserFields.Fields.Item("U_SS_Declarable").Value
                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Campo U_SS_Declarable =  " + externa.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'Catch ex As Exception
                    '    rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Campo Error TRY CATCH  " + ex.Message.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    'End Try
                    'If externa = "NO" Then
                    '    ' PARAMETRO PARA DESACTIVAR PREGUNTA POR EL CAMPO U_DOCDECLARABLE Y NO LO ENVIA AL SRI
                    '    Dim DesactivaPreguntaDeclaraable As String = "N"
                    '    DesactivaPreguntaDeclaraable = ConsultaParametro(Functions.VariablesGlobales._vgNombreAddOn, "PARAMETROS", "CONFIGURACION", "DesactivaPreguntaDeclaraable")
                    '    Utilitario.Util_Log.Escribir_Log("DesactivaPreguntaDeclaraable: " + DesactivaPreguntaDeclaraable, "ManejoDeDocumentos")
                    '    ' END PARAMETRO PARA VALIDAR SI SE ENVÍA EL DOCUMENTO AL MOMENTO DE CREAR LA FACTURA

                    '    If DesactivaPreguntaDeclaraable = "Y" Then
                    '        Procesar = True
                    '        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Parametro activo de Desactivar Pregunta Declarable, se envia igual al SRI..  ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    '    Else
                    '        Dim iReturnValue As Integer
                    '        iReturnValue = rSboApp.MessageBox(NombreAddon + " - El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?", 1, "&SI", "&NO")
                    '        If iReturnValue = 2 Then
                    '            rSboApp.SetStatusBarMessage(NombreAddon + " - Respuesta igual NO, no se envió el documento al SRI..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    '            Procesar = False
                    '            BubbleEvent = False
                    '            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : NO", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    '        Else
                    '            Procesar = True
                    '            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando Respuesta a 'El Documento tiene marcado como DECLARABLE = 'NO', desea igual enviarlo al SRI?' " + ", Contesto : SI", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    '        End If
                    '    End If

                    'End If
                    Procesar = True
                End If
            End If
            ' **** FIN CONTROL POR CAMPO DE USUARIO "EXTERNA"

            If Procesar Then

                If oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then ' SI ES TRANSFERENCIA ES OTRO OBJETO, NO ES DOCUMENTS
                    'If Not oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then
                    If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                        'If Not oTransferencia.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation Then
                        rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        If SerieExcluir = "N" Then
                            Try

                                rSboApp.SetStatusBarMessage(NombreAddon + " - Consultando siguiente numero de Folio..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Consultando siguiente numero de Folio..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                Dim sSQL As String = ""
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@EXX_DOCUM_LEG_INTER"" A "
                                        sSQL += " INNER JOIN ""NNM1"" B ON A.""U_NOMBRE"" = B.""SeriesName"" "
                                        sSQL += "WHERE B.""Series"" = " + oTransferencia.Series.ToString
                                    Else
                                        sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@EXX_DOCUM_LEG_INTER] A WITH(NOLOCK) "
                                        sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_NOMBRE = B.SeriesName "
                                        sSQL += "WHERE B.Series = " + oTransferencia.Series.ToString
                                    End If

                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Ejecutado Consulta para obtener Folio : " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                    sFolio = oFuncionesB1.getRSvalue(sSQL, "U_ULT_SECUEN")
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Se obtuvo el siguiente # de Folio : " + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                    sCode = oFuncionesB1.getRSvalue(sSQL, "Code")
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Se obtuvo el Code: " + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                    '' VALIDACION: Se agrego este control por Centuriosa, ya que se le estaban reiniciando la secuencia
                                    ''             la misma que se debía a que en algun punto cuando se procesaba el documento, no encontraba el registro
                                    ''             en la tabla/UDO de usuario @EXX_DOCUM_LEG_INTER
                                    ''             Si no encuentra registro en la tabla significa que no es un documento electronico, esta validacion se encuntra 
                                    ''             en una seccion previa. Pero se refuerza la validación en este punto.
                                    If sCode = "" Then
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Al NO Obtenerse el Code, implica que no esta registrado en la tabla DocLegalInterno la serie del documento : " + oTransferencia.Series.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Se Omitió el envío al SRI, BubbleEvent = False", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    '' END VALIDACION: Se agrego este control por Centuriosa, ya que se le estaban reiniciando la secuencia
                                    ''             la misma que se debía a que en algun punto cuando se procesaba el documento, no encontraba el registro
                                    ''             en la tabla/UDO de usuario @EXX_DOCUM_LEG_INTER
                                    ''             Si no encuentra registro en la tabla significa que no es un documento electronico, esta validacion se encuntra 
                                    ''             en una seccion previa. Pero se refuerza la validación en este punto.

                                    numFolio = Integer.Parse(sFolio) + 1
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Se sumo 1 al # de Folio: " + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                    ' VALIDAR SI EL DOCUMENTO YA TIENE UN FOLIO ASIGNADO
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Validando si ya tiene Folio Asignado ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Folio Actual: " + oTransferencia.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                    'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    '    sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@SERIES"" A "
                                    '    sSQL += " INNER JOIN ""NNM1"" B ON A.""U_SERIE"" = B.""Series"" "
                                    '    sSQL += "WHERE B.""Series"" = " + oTransferencia.Series.ToString
                                    'Else
                                    '    sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@SERIES] A WITH(NOLOCK) "
                                    '    sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_SERIE = B.Series "
                                    '    sSQL += "WHERE B.Series = " + oTransferencia.Series.ToString
                                    'End If

                                    Try

                                        If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then

                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                            If Not EnviaDocumentosEnBackGround = "Y" Then
                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                Else
                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                End If

                                            Else
                                                rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            End If

                                        End If

                                    Catch ex As Exception


                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Error One SoLutions Procesamiento, Folio Actual: " + oTransferencia.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                    End Try




                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                    sSQL = Functions.VariablesGlobales._ConsultaFolioSS

                                    If sSQL = "" Then
                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Por favor agregar consulta para asignacion de Folio", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        Exit Sub
                                    End If

                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                    sSQL = sSQL.Replace("IDENTIFICADOR", oTransferencia.DocEntry)
                                    sSQL = sSQL.Replace("TIPDOCE", "('GR')")

                                    sFolio = oFuncionesB1.getRSvalue(sSQL, "Secuencial")
                                    sCode = oFuncionesB1.getRSvalue(sSQL, "TipoDoc")

                                    If sFolio = "" Then
                                        rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro campo Secuencial en la consulta, por favor verificar..!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        Exit Sub
                                    Else
                                        numFolio = sFolio
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Se obtuvo el siguiente # de Folio : " + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    End If

                                End If


                                If oTransferencia.FolioNumber > 0 Then
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - El documento ya tiene asigando un Número de Folio, se obviará la foliación..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    oTransferencia.UserFields.Fields.Item("U_OBSERVACION_FACT").Value += "El documento ya tiene asigando un Número de Folio, se obviará la foliación..!!"
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento ya tiene asigando un Número de Folio, se obviará la foliación..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    ActualizaSecuenciaFolio = False
                                Else
                                    ActualizaSecuenciaFolio = True
                                    oTransferencia.FolioNumber = numFolio
                                    oTransferencia.FolioPrefixString = oTipoTabla '"/*Prefijo del folio*/"
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Folio Asignado :" + oTransferencia.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Asignando Prefijo :" + oTransferencia.FolioPrefixString.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                End If

                                'oTransferencia.Printed = SAPbobsCOM.PrintStatusEnum.psYes
                            Catch ex As Exception
                                rSboApp.SetStatusBarMessage(NombreAddon + " - Error al consultar siguiente folio..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Error al consultar siguiente folio..!!" + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                Exit Sub
                            End Try

                            RetVal = oTransferencia.Update()
                            If RetVal <> 0 Then

                                rCompany.GetLastError(ErrCode, ErrMsg)

                                rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar el numero de folio..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Error al asignar el numero de folio..!" + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                Exit Sub
                            Else
                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Actualizando la secuencia en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    If ActualizaSecuenciaFolio = True Then
                                        If oFuncionesAddon.ActualizaSecuencia(sCode, numFolio, oTransferencia.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            If Not EnviaDocumentosEnBackGround = "Y" Then
                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                Else
                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                End If

                                            Else
                                                rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            End If
                                        Else
                                            ' VUELVE A DEJAR EL FOLIO EN 0, EN CASO QUE DE ERROR LA ACTUALIZACIÓN DE DOC LEGAL INTERNO
                                            oTransferencia.FolioNumber = 0
                                            RetVal = oTransferencia.Update()

                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Error al actualizar la secuencia en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            Exit Sub
                                        End If
                                    ElseIf ActualizaSecuenciaFolio = False Then
                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                            Else
                                                oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                            End If

                                        Else
                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End If

                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la Tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Actualizando la secuencia en la Tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    If ActualizaSecuenciaFolio = True Then
                                        If oFuncionesAddon.ActualizaSecuencia_ONE(sCode, numFolio, oTransferencia.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Secuencia Actualizada en la Tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            If Not EnviaDocumentosEnBackGround = "Y" Then
                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                Else
                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                End If

                                            Else
                                                rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            End If
                                        ElseIf ActualizaSecuenciaFolio = False Then
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la Tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Error al actualizar la secuencia en la Tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            Exit Sub
                                        End If
                                    Else
                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                            Else
                                                oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                            End If

                                        Else
                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End If
                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la Tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Actualizando la secuencia en la Tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    If ActualizaSecuenciaFolio = True Then
                                        If oFuncionesAddon.ActualizaSecuencia_SYPSOFT(sCode, numFolio, oTransferencia.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Secuencia Actualizada en la Tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            If Not EnviaDocumentosEnBackGround = "Y" Then
                                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                Else
                                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                                End If

                                            Else
                                                rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            End If
                                        ElseIf ActualizaSecuenciaFolio = False Then
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la Tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Error al actualizar la secuencia en la Tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            Exit Sub
                                        End If
                                    Else
                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                            Else
                                                oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                            End If

                                        Else
                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End If

                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                    'rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en (SS) Documentos Legales..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    'oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Actualizando la secuencia en (SS) Documentos Legales..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    'If ActualizaSecuenciaFolio = True Then
                                    '    If oFuncionesAddon.ActualizaSecuenciaSS(sCode, numFolio, oTransferencia.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                    '        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Secuencia Actualizada en (SS) Documentos Legales..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                        Else
                                            oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                        End If

                                    Else
                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    End If


                                End If


                            End If
                        Else
                            oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Serie Excluida, No se folea, pero se envia al SRI...", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            If Not EnviaDocumentosEnBackGround = "Y" Then
                                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                    oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                Else
                                    oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                End If

                            Else
                                rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            End If
                        End If
                        'End If
                    End If

                ElseIf oTipoTabla = "TLE" Then ' SSOLICITUD DE TRASLADO
                    'If Not oTransferencia.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then
                    If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                        rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                        Try

                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then

                                If Not EnviaDocumentosEnBackGround = "Y" Then
                                    If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                        oManejoDocumentosEcua.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                    Else
                                        oManejoDocumentos.ProcesaEnvioDocumento(oTransferencia.DocEntry, oTipoTabla)
                                    End If

                                Else
                                    rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oTransferencia.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                End If


                            End If

                        Catch ex As Exception

                        End Try
                    End If

                Else
                    If Not oDocumento.CancelStatus = SAPbobsCOM.CancelStatusEnum.csCancellation Then ' SI EL DOCUMENTO NO ES UNA CANCELACION
                        'If Not oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDrafts Then                                               
                        If Not BusinessObjectInfo.Type = SAPbobsCOM.BoObjectTypes.oDrafts Then
                            rSboApp.SetStatusBarMessage(NombreAddon + " - Procesando Documento Electronico..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Procesando Documento Electronico..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                            If oTipoTabla = "REE" Or oTipoTabla = "REA" Or oTipoTabla = "RDM" Or oTipoTabla = "RER" Then
                                'queryestadoRT = oFuncionesB1.getRSvalue("SELECT ""U_ESTADO_AUTORIZACIO"" FROM ""OPCH"" WHERE ""DocEntry"" = '" + oDocumento.DocEntry.ToString() + "' ", "U_ESTADO_AUTORIZACIO", "")
                                queryestadoRT = oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value
                                'LTrim(RTrim(Form.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0)))
                                If queryestadoRT <> "0" Or queryestadoRT = "" Then
                                    queryestadoRT = "Y"
                                Else

                                    queryestadoRT = "N"
                                End If
                                Utilitario.Util_Log.Escribir_Log("estado del documento: " + queryestadoRT.ToString, "ManejoDeDocumentos")
                            Else
                                queryestadoRT = "N"
                            End If
                            'Dim estLQ As String = oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value
                            'If Functions.VariablesGlobales._vgNoEnviarRT = "Y" Then
                            '    Utilitario.Util_Log.Escribir_Log("Paramero activo no reenciar sri rt: " + Functions.VariablesGlobales._vgNoEnviarRT.ToString, "EventosEmision")
                            '    If LQEsElectronico = "FE" And (estLQ = "3" Or estLQ = "6" Or estLQ = "4") Then
                            '        Utilitario.Util_Log.Escribir_Log("LQEsElectronico: " + LQEsElectronico + " Estado: " + estLQ.ToString, "EventosEmision")
                            '        rSboApp.SetStatusBarMessage(NombreAddon + " - No se emitirá la Retención hasta a que la Liquidación se encuentre Autorizada", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            '        Exit Sub
                            '    End If
                            'End If
                            If queryestadoRT = "N" Then
                                Utilitario.Util_Log.Escribir_Log("estado del documento1: " + queryestadoRT.ToString, "ManejoDeDocumentos")
                                If SerieExcluir = "N" Then
                                    Dim TieneFolio As Boolean = False
                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                        If oTipoTabla = "REE" Or
                                            oTipoTabla = "REA" Or
                                            oTipoTabla = "RDM" Or
                                            oTipoTabla = "RER" Then
                                            If Not oDocumento.UserFields.Fields.Item("U_COMP_RET").Value = "" Then
                                                TieneFolio = True
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "numero de Folio encontrado en la retencion al crear: " + oDocumento.UserFields.Fields.Item("U_COMP_RET").Value.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            End If
                                        End If
                                    End If
                                    If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                        If oTipoTabla = "REE" Or
                                            oTipoTabla = "REA" Or
                                            oTipoTabla = "RDM" Or
                                            oTipoTabla = "RER" Then
                                            If Not oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value = "" Then
                                                TieneFolio = True
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "numero de Folio encontrado en la retencion al crear: " + oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            End If
                                        End If
                                    End If
                                    If TieneFolio = False Then
                                        Try
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Consultando siguiente numero de Folio..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Consultando siguiente numero de Folio..!! ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                            Dim sSQL As String = ""
                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                    sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@EXX_DOCUM_LEG_INTER"" A "
                                                    sSQL += " INNER JOIN ""NNM1"" B ON A.""U_NOMBRE"" = B.""SeriesName"" "
                                                    sSQL += "WHERE B.""Series"" = " + oDocumento.Series.ToString
                                                Else
                                                    sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@EXX_DOCUM_LEG_INTER] A WITH(NOLOCK) "
                                                    sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_NOMBRE = B.SeriesName "
                                                    sSQL += "WHERE B.Series = " + oDocumento.Series.ToString
                                                End If

                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Ejecutado Consulta para obtener Folio : " + sSQL.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "U_ULT_SECUEN")
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se obtuvo el siguiente # de Folio : " + sFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                If sFolio = "" Then
                                                    sFolio = 0
                                                End If
                                                sCode = oFuncionesB1.getRSvalue(sSQL, "Code")
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se obtuvo el Code: " + sCode.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                '' VALIDACION: Se agrego este control por Centuriosa, ya que se le estaban reiniciando la secuencia
                                                ''             la misma que se debía a que en algun punto cuando se procesaba el documento, no encontraba el registro
                                                ''             en la tabla/UDO de usuario @EXX_DOCUM_LEG_INTER
                                                ''             Si no encuentra registro en la tabla significa que no es un documento electronico, esta validacion se encuntra 
                                                ''             en una seccion previa. Pero se refuerza la validación en este punto.
                                                If sCode = "" Then
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Al NO Obtenerse el Code, implica que no esta registrado en la tabla DocLegalInterno la serie del documento : " + oDocumento.Series.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se Omitió el envío al SRI, BubbleEvent = False", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                                '' END VALIDACION: Se agrego este control por Centuriosa, ya que se le estaban reiniciando la secuencia
                                                ''             la misma que se debía a que en algun punto cuando se procesaba el documento, no encontraba el registro
                                                ''             en la tabla/UDO de usuario @EXX_DOCUM_LEG_INTER
                                                ''             Si no encuentra registro en la tabla significa que no es un documento electronico, esta validacion se encuntra 
                                                ''             en una seccion previa. Pero se refuerza la validación en este punto.

                                                numFolio = Integer.Parse(sFolio) + 1
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Se sumo 1 al # de Folio: " + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)


                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                    sSQL = "SELECT ""Code"",IFNULL(""U_ULT_SECUEN"",0) AS U_ULT_SECUEN FROM ""@SERIES"" A "
                                                    sSQL += " INNER JOIN ""NNM1"" B ON A.""U_SERIE"" = B.""Series"" "
                                                    sSQL += "WHERE B.""Series"" = " + oDocumento.Series.ToString
                                                Else
                                                    sSQL = "SELECT Code,ISNULL(U_ULT_SECUEN,0) AS U_ULT_SECUEN FROM [@SERIES] A WITH(NOLOCK) "
                                                    sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.U_SERIE = B.Series "
                                                    sSQL += "WHERE B.Series = " + oDocumento.Series.ToString
                                                End If
                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                    sSQL = "SELECT ""Code"",0 AS U_ULT_SECUEN FROM ""@GS_SERIESE"" A "
                                                    sSQL += " INNER JOIN ""NNM1"" B ON A.""Code"" = B.""Series"" "
                                                    sSQL += "WHERE A.""Code"" = " + oDocumento.Series.ToString
                                                Else
                                                    sSQL = "SELECT Code,0 AS U_ULT_SECUEN FROM [@GS_SERIESE] A WITH(NOLOCK) "
                                                    sSQL += " INNER JOIN NNM1 B WITH(NOLOCK) ON A.Code = B.Series "
                                                    sSQL += "WHERE A.Code = " + oDocumento.Series.ToString
                                                End If
                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                sSQL = Functions.VariablesGlobales._ConsultaFolioSS

                                                If sSQL = "" Then
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - Por favor agregar consulta para asignacion de Folio", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                    Exit Sub
                                                End If


                                                If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Or oTipoTabla = "REA" Then

                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                    sSQL = sSQL.Replace("IDENTIFICADOR", oDocumento.DocEntry)
                                                    sSQL = sSQL.Replace("TIPDOCE", "('RT','LQRT')")

                                                ElseIf oTipoTabla = "FCE" Or oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then

                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                    sSQL = sSQL.Replace("IDENTIFICADOR", oDocumento.DocEntry)
                                                    sSQL = sSQL.Replace("TIPDOCE", "('FV')")

                                                ElseIf oTipoTabla = "NCE" Then

                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                    sSQL = sSQL.Replace("IDENTIFICADOR", oDocumento.DocEntry)
                                                    sSQL = sSQL.Replace("TIPDOCE", "('NC')")

                                                ElseIf oTipoTabla = "NDE" Then

                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                    sSQL = sSQL.Replace("IDENTIFICADOR", oDocumento.DocEntry)
                                                    sSQL = sSQL.Replace("TIPDOCE", "('ND')")

                                                ElseIf oTipoTabla = "GRE" Or oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then

                                                    sSQL = sSQL.Replace("TABLA", oTabla)
                                                    sSQL = sSQL.Replace("IDENTIFICADOR", oDocumento.DocEntry)
                                                    sSQL = sSQL.Replace("TIPDOCE", "('GR')")

                                                End If


                                                sFolio = oFuncionesB1.getRSvalue(sSQL, "Secuencial")
                                                sCode = oFuncionesB1.getRSvalue(sSQL, "TipoDoc")

                                                If sFolio = "" Then
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - No se encontro campo Secuencial en la consulta, por favor verificar..!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                    Exit Sub
                                                Else
                                                    numFolio = sFolio
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "# de Folio calculado: " + numFolio.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                End If


                                            End If


                                            'REE - RETENCION - FACTURA DE PROVEEDOR
                                            'REA - RETENCION - FACTURA DE ANTICIPO DE PROVEEDOR
                                            'RER - RETENCION - FACTURA DE RESERVA DE PROVEEDOR
                                            If oTipoTabla = "REE" Or
                                                oTipoTabla = "REA" Or
                                                oTipoTabla = "RDM" Or
                                                oTipoTabla = "RER" Then
                                                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                    oDocumento.UserFields.Fields.Item("U_COMP_RET").Value = numFolio.ToString()
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "# Retención Asignado :" + oDocumento.UserFields.Fields.Item("U_COMP_RET").Value.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    ActualizaSecuenciaFolio = True
                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                    ' OBTENER ESTABLECIMIENTO Y PUNTO DE EMISION DE LA NNM1, PARA LUEGO ASIGNARLO EN EL NUMERO DE RETENCION
                                                    Dim sEstablecimiento As String = ""
                                                    Dim sPtoEmision As String = ""
                                                    sEstablecimiento = oFuncionesB1.getRSvalue("SELECT ""BeginStr"" FROM ""NNM1"" WHERE ""Series"" = " + oDocumento.Series.ToString, "BeginStr", "")
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Establecimiento Obtenido :" + sEstablecimiento, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                    sPtoEmision = oFuncionesB1.getRSvalue("SELECT ""EndStr"" FROM ""NNM1"" WHERE ""Series"" = " + oDocumento.Series.ToString, "EndStr", "")
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Punto de Emisión Obtenido :" + sPtoEmision, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                    'END OBTENER ESTABLECIMIENTO Y PUNTO DE EMISION DE LA NNM1, PARA LUEGO ASIGNARLO EN EL NUMERO DE RETENCION

                                                    oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value = sEstablecimiento + sPtoEmision + numFolio.ToString().PadLeft(9, "0")

                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "# Retención Asignado :" + oDocumento.UserFields.Fields.Item("U_RETENCION_NO").Value.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    ActualizaSecuenciaFolio = True

                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then

                                                    ''''PENDIENTE CONOCER DONDE SE GUARDA EL NUMERO DE FOLIO DE LA RETENCION
                                                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                    oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value = numFolio.ToString()
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "# Retención Asignado :" + oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    ActualizaSecuenciaFolio = True


                                                End If
                                            Else
                                                ' VALIDAR SI EL DOCUMENTO YA TIENE UN FOLIO ASIGNADO
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Validando si ya tiene Folio Asignado ", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Folio Actual: " + oDocumento.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                If oDocumento.FolioNumber > 0 Then
                                                    rSboApp.SetStatusBarMessage(NombreAddon + " - El documento ya tiene asigando un Número de Folio, se obviará la foliación..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                    oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value += "El documento ya tiene asigando un Número de Folio, se obviará la foliación..!!"
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento ya tiene asigando un Número de Folio, se obviará la foliación..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    ActualizaSecuenciaFolio = False

                                                Else
                                                    oDocumento.FolioNumber = numFolio
                                                    oDocumento.FolioPrefixString = oTipoTabla '"/*Prefijo del folio*/"
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Folio Asignado :" + oDocumento.FolioNumber.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Asignando Prefijo :" + oDocumento.FolioPrefixString.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    ActualizaSecuenciaFolio = True
                                                End If

                                                oDocumento.Printed = SAPbobsCOM.PrintStatusEnum.psYes

                                                ' INCIALIZO LOS CAMPOS DE ADDON QUE GUARDAN LA AUTORIZACIÓN
                                                If oFuncionesB1.checkCampoBD("OINV", "NUM_AUTO_FAC") Then
                                                    oDocumento.UserFields.Fields.Item("U_NUM_AUTO_FAC").Value = ""
                                                End If
                                                If oFuncionesB1.checkCampoBD("OINV", "CLAVE_ACCESO") Then
                                                    oDocumento.UserFields.Fields.Item("U_CLAVE_ACCESO").Value = ""
                                                End If
                                                If oFuncionesB1.checkCampoBD("OINV", "ESTADO_AUTORIZACIO") Then
                                                    oDocumento.UserFields.Fields.Item("U_ESTADO_AUTORIZACIO").Value = "0"
                                                End If
                                                If oFuncionesB1.checkCampoBD("OINV", "OBSERVACION_FACT") Then
                                                    oDocumento.UserFields.Fields.Item("U_OBSERVACION_FACT").Value = ""
                                                End If
                                                ' END INCIALIZO LOS CAMPOS DE ADDON QUE GUARDAN LA AUTORIZACIÓN

                                                ' VALIDO SI EXISTE EL CAMPO NUMFOLIO QUE ESTABA EN EL VERSION PI, PARA SI EXISTE ACTUALIZAR EL FOLIO
                                                ' CENTURIOSA USA ESE CAMPO PARA UN REPORTE
                                                If oFuncionesB1.checkCampoBD("OINV", "NUMFOLIO") Then
                                                    oDocumento.UserFields.Fields.Item("U_NUMFOLIO").Value = numFolio.ToString()
                                                End If
                                            End If

                                        Catch ex As Exception
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al consultar siguiente folio..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al consultar siguiente folio..!!, Catch: " + ex.Message.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            Exit Sub
                                        End Try

                                        RetVal = oDocumento.Update()

                                        'If oDocumento.FolioNumber = 0 Then validar si filtarlo por documento
                                        '    Try
                                        '        oDocumento.FolioNumber = numFolio
                                        '        RetVal = oDocumento.Update()
                                        '    Catch ex As Exception
                                        '        Utilitario.Util_Log.Escribir_Log("Error al actualizar nuevamente el folio: " + ex.ToString, "ManejoDeDocumentos")
                                        '    End Try

                                        'End If
                                        If RetVal <> 0 Then
                                            rCompany.GetLastError(ErrCode, ErrMsg)
                                            '    MsgBox(ErrCode & " " & ErrMsg)
                                            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al asignar el numero de folio..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al asignar el numero de folio..!!" + ErrMsg.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                            Exit Sub
                                        Else
                                            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Actualizando la secuencia en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                If ActualizaSecuenciaFolio = True Then
                                                    If oFuncionesAddon.ActualizaSecuencia(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Secuencia Actualizada en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                        Dim estLQ As String = oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value
                                                        If Functions.VariablesGlobales._vgNoEnviarRT = "Y" Then
                                                            Utilitario.Util_Log.Escribir_Log("Paramero activo no reenciar sri rt: " + Functions.VariablesGlobales._vgNoEnviarRT.ToString, "EventosEmision")
                                                            If LQEsElectronico = "FE" And (estLQ = "3" Or estLQ = "6" Or estLQ = "4") Then
                                                                Utilitario.Util_Log.Escribir_Log("LQEsElectronico: " + LQEsElectronico + " Estado: " + estLQ.ToString, "EventosEmision")
                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - No se emitirá la Retención hasta a que la Liquidación se encuentre Autorizada", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                Exit Sub
                                                            End If
                                                        End If

                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                            Else
                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                            End If

                                                        Else
                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        End If

                                                    Else
                                                        ' VUELVE A DEJAR EL FOLIO EN 0, EN CASO QUE DE ERROR LA ACTUALIZACIÓN DE DOC LEGAL INTERNO
                                                        If oTipoTabla = "REE" Or
                                                             oTipoTabla = "REA" Or
                                                                 oTipoTabla = "RDM" Or
                                                                     oTipoTabla = "RER" Then

                                                            oDocumento.UserFields.Fields.Item("U_COMP_RET").Value = ""
                                                        Else
                                                            oDocumento.FolioNumber = 0
                                                        End If
                                                        RetVal = oDocumento.Update()
                                                        ' END VUELVE A DEJAR EL FOLIO EN 0, EN CASO QUE DE ERROR LA ACTUALIZACIÓN DE DOC LEGAL INTERNO

                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al actualizar la secuencia en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        Exit Sub
                                                    End If
                                                ElseIf ActualizaSecuenciaFolio = False Then

                                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        End If

                                                    Else
                                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    End If

                                                End If

                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la Tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Actualizando la secuencia en la Tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                If ActualizaSecuenciaFolio = True Then
                                                    If oFuncionesAddon.ActualizaSecuencia_ONE(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Secuencia Actualizada en la Tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                            Else
                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                            End If

                                                        Else
                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        End If

                                                    Else
                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la Tabla SERIES..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al actualizar la secuencia en la Tabla SERIES..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        Exit Sub
                                                    End If
                                                ElseIf ActualizaSecuenciaFolio = False Then

                                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        End If

                                                    Else
                                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    End If

                                                End If

                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                                                rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en la Tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Actualizando la secuencia en la Tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                If ActualizaSecuenciaFolio = True Then
                                                    If oFuncionesAddon.ActualizaSecuencia_SYPSOFT(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Secuencia Actualizada en la Tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                            Else
                                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                            End If

                                                        Else
                                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        End If
                                                    Else
                                                        rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en la Tabla GS_SERIESE..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al actualizar la secuencia en la Tabla GS_SERIESE..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                        Exit Sub
                                                    End If

                                                ElseIf ActualizaSecuenciaFolio = False Then

                                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        End If

                                                    Else
                                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    End If

                                                End If
                                            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                                                'rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizando la secuencia en (SS) Documentos Legales..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                'oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Actualizando la secuencia en (SS) Documentos Legales..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                If ActualizaSecuenciaFolio = True Then
                                                    'If oFuncionesAddon.ActualizaSecuenciaSS(sCode, numFolio, oDocumento.DocEntry, oTipoTabla, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision) Then
                                                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Secuencia Actualizada en (SS) Documentos Legales..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    'AQUI CAMBIAR LA VALIDACION PARA QUE NO SE ENVIE LA RETENCION 
                                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                                        Dim estLQ As String = oDocumento.UserFields.Fields.Item("U_LQ_ESTADO").Value
                                                        If Functions.VariablesGlobales._vgNoEnviarRT = "Y" Then
                                                            Utilitario.Util_Log.Escribir_Log("Paramero activo no reenciar sri rt: " + Functions.VariablesGlobales._vgNoEnviarRT.ToString, "EventosEmision")
                                                            If LQEsElectronico = "FE" And (estLQ = "3" Or estLQ = "6" Or estLQ = "4") Then
                                                                Utilitario.Util_Log.Escribir_Log("LQEsElectronico: " + LQEsElectronico + " Estado: " + estLQ.ToString, "EventosEmision")
                                                                rSboApp.SetStatusBarMessage(NombreAddon + " - No se emitirá la Retención hasta a que la Liquidación se encuentre Autorizada", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                                Exit Sub
                                                            End If
                                                        End If
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        End If

                                                    Else
                                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    End If

                                                    'Else
                                                    '    ' VUELVE A DEJAR EL FOLIO EN 0, EN CASO QUE DE ERROR LA ACTUALIZACIÓN DE DOC LEGAL INTERNO
                                                    '    If oTipoTabla = "REE" Or
                                                    '         oTipoTabla = "REA" Or
                                                    '             oTipoTabla = "RDM" Or
                                                    '                 oTipoTabla = "RER" Then

                                                    '        oDocumento.UserFields.Fields.Item("U_SS_SecRet").Value = ""
                                                    '    Else
                                                    '        oDocumento.FolioNumber = 0
                                                    '    End If
                                                    '    RetVal = oDocumento.Update()
                                                    '    ' END VUELVE A DEJAR EL FOLIO EN 0, EN CASO QUE DE ERROR LA ACTUALIZACIÓN DE DOC LEGAL INTERNO

                                                    '    rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar la secuencia en Documentos Legales Internos..!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                                    '    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Error al actualizar la secuencia en Documentos Legales Internos..!!", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    '    Exit Sub
                                                    'End If
                                                ElseIf ActualizaSecuenciaFolio = False Then

                                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        Else
                                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                                        End If

                                                    Else
                                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                                    End If

                                                End If
                                            End If

                                        End If
                                    Else
                                        If Not EnviaDocumentosEnBackGround = "Y" Then
                                            If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                                oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                            Else
                                                oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                            End If

                                        Else
                                            rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                        End If
                                    End If

                                Else
                                    oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Serie Excluida, No se folea, pero se envia al SRI...", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    If Not EnviaDocumentosEnBackGround = "Y" Then
                                        If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                                            oManejoDocumentosEcua.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                        Else
                                            oManejoDocumentos.ProcesaEnvioDocumento(oDocumento.DocEntry, oTipoTabla)
                                        End If

                                    Else
                                        rSboApp.SetStatusBarMessage("El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "El documento se enviará al SRI a través de servicio windows, revise el estado del documentos en unos minutos..", Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                                    End If
                                End If
                            Else
                                Utilitario.Util_Log.Escribir_Log("estado del documento exit sub: " + queryestadoRT.ToString, "ManejoDeDocumentos")
                                Exit Sub
                            End If


                        End If


                    End If

                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("estado del documento exit sub: " + ex.ToString, "ManejoDeDocumentos")
        End Try
    End Sub

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rSboApp.RightClickEvent

        Try
            Dim typeEx, idForm As String
#Disable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            typeEx = oFuncionesB1.FormularioActivo(idForm)
#Enable Warning BC42030 ' La variable 'idForm' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idForm)
            If mForm.TypeEx = "133" Or
                mForm.TypeEx = "60090" Or
                mForm.TypeEx = "60091" Or
                 mForm.TypeEx = "60092" Or
                  mForm.TypeEx = "65303" Or
                  mForm.TypeEx = "65307" Or
                   mForm.TypeEx = "65300" Or
                    mForm.TypeEx = "65301" Or
                     mForm.TypeEx = "179" Or
                       mForm.TypeEx = "140" Or
                         mForm.TypeEx = "940" Or
                          mForm.TypeEx = "1250000940" Or
                           mForm.TypeEx = "141" Or
                           mForm.TypeEx = "85" Then

                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                If eventInfo.BeforeAction Then
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    oMenuItem = rSboApp.Menus.Item("1280")
                    If oMenuItem.SubMenus.Exists("SS_LOG") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("SS_LOG"))
                    End If

                    oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "SS_LOG"
                    oCreationPackage.String = "(SS) Log Emisión de Documentos..."
                    oCreationPackage.Enabled = True
                    oCreationPackage.Position = 20
                    oMenuItem = rSboApp.Menus.Item("1280")
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)

                Else
                    oMenuItem = rSboApp.Menus.Item("1280")
                    If oMenuItem.SubMenus.Exists("SS_LOG") Then
                        oMenuItem.SubMenus.Remove(oMenuItem.SubMenus.Item("SS_LOG"))
                    End If
                End If

            End If


        Catch ex As Exception
            rSboApp.MessageBox("Error: " & ex.Message)
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

    Private Sub SeteaTipoTabla(pVal As SAPbouiCOM.ItemEvent)
        Try
            If pVal.FormTypeEx = "133" Or pVal.FormTypeEx = "60090" Or pVal.FormTypeEx = "65307" Then
                oTabla = "OINV"
                oTipoTabla = "FCE"
            ElseIf pVal.FormTypeEx = "60091" Then
                oTabla = "OINV"
                oTipoTabla = "FRE"
            ElseIf pVal.FormTypeEx = "65303" Then
                oTabla = "OINV"
                oTipoTabla = "NDE"
            ElseIf pVal.FormTypeEx = "65300" Then
                oTabla = "ODPI"
                oTipoTabla = "FAE"
            ElseIf pVal.FormTypeEx = "179" Then
                oTabla = "ORIN"
                oTipoTabla = "NCE"
            ElseIf pVal.FormTypeEx = "140" Then
                oTabla = "ODLN"
                oTipoTabla = "GRE"
            ElseIf pVal.FormTypeEx = "940" Then
                oTabla = "OWTR"
                oTipoTabla = "TRE"
            ElseIf pVal.FormTypeEx = "1250000940" Then
                oTabla = "OWTQ"
                oTipoTabla = "TLE"
            ElseIf pVal.FormTypeEx = "141" Then
                oTabla = "OPCH"
                oTipoTabla = "REE"
            ElseIf pVal.FormTypeEx = "65306" Then 'NOTA DE DEBITO DE PROVEEDORES
                oTabla = "OPCH"
                oTipoTabla = "RDM"
            ElseIf pVal.FormTypeEx = "60092" Then   'FACTURA DE RESERVA PROVEEDOR/RETENCION
                oTabla = "OPCH"
                oTipoTabla = "RER"
            ElseIf pVal.FormTypeEx = "65301" Then ' 'FACTURA DE ANTICIPO DE PROVEEDORES
                oTabla = "ODPO"
                oTipoTabla = "REA"
            ElseIf pVal.FormTypeEx = "720" Then 'Salida de mercancias
                oTabla = "OIGE"
                oTipoTabla = "GRSM"
            End If
        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al Setear Tipo Tabla: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub SeteaTipoTabla_FormTypeEx(FormTypeEx As String)
        Try
            If FormTypeEx = "133" Or FormTypeEx = "60090" Or FormTypeEx = "65307" Then
                oTabla = "OINV"
                oTipoTabla = "FCE"
            ElseIf FormTypeEx = "60091" Then
                oTabla = "OINV"
                oTipoTabla = "FRE"
            ElseIf FormTypeEx = "65303" Then
                oTabla = "OINV"
                oTipoTabla = "NDE"
            ElseIf FormTypeEx = "65300" Then
                oTabla = "ODPI"
                oTipoTabla = "FAE"
            ElseIf FormTypeEx = "179" Then
                oTabla = "ORIN"
                oTipoTabla = "NCE"
            ElseIf FormTypeEx = "140" Then
                oTabla = "ODLN"
                oTipoTabla = "GRE"
            ElseIf FormTypeEx = "940" Then
                oTabla = "OWTR"
                oTipoTabla = "TRE"
            ElseIf FormTypeEx = "1250000940" Then
                oTabla = "OWTQ"
                oTipoTabla = "TLE"
            ElseIf FormTypeEx = "141" Then
                oTabla = "OPCH"
                oTipoTabla = "REE"

            ElseIf FormTypeEx = "65306" Then
                oTabla = "OPCH"
                oTipoTabla = "RDM"

            ElseIf FormTypeEx = "60092" Then   'FACTURA DE RESERVA PROVEEDOR/RETENCION
                oTabla = "OPCH"
                oTipoTabla = "RER"
            ElseIf FormTypeEx = "65301" Then ' 'FACTURA DE ANTICIPO DE PROVEEDORES
                oTabla = "ODPO"
                oTipoTabla = "REA"
            ElseIf FormTypeEx = "720" Then ' SOLICITUD DE MERCANCIAS
                oTabla = "OIGE"
                oTipoTabla = "GRSM"
            End If
        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al Setear Tipo Tabla: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub


    Public Function ValidarSerieDocElec(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo) As Boolean
        Dim form As SAPbouiCOM.Form = rSboApp.Forms.Item(BusinessObjectInfo.FormUID)
        Dim odbds As SAPbouiCOM.DBDataSource = CType(form.DataSources.DBDataSources.Item(0), SAPbouiCOM.DBDataSource)
        Dim SerieDocElec As String = odbds.GetValue("Series", odbds.Offset).Trim
        If Functions.VariablesGlobales._vgSerieUDF = "Y" And (oTipoTabla = "REE" Or oTipoTabla = "RER") Then
            Dim SerieDocUDF = odbds.GetValue("U_DocEmision", odbds.Offset).Trim
            If SerieDocUDF = "01" Or SerieDocUDF = "03" Then
                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "Verificacion RT electronica UDF: " + SerieDocUDF.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)

                Return True
            Else
                oFuncionesAddon.GuardaLOG(oTipoTabla, oDocumento.DocEntry, "RT NO Electronica UDF: " + SerieDocUDF.ToString(), Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                Return False

            End If
        Else

            'Dim SerieElec As String = ValidarDocElec(SerieDocElec)
            Dim SerieDoc As String = ""
            Dim _esElectronico As String = ""
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SerieDoc = "SELECT ""U_FE_TipoEmision"" AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE A.""Series"" = " + SerieDocElec
                Else
                    SerieDoc = "SELECT U_FE_TipoEmision AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE Series = " + SerieDocElec
                End If
                _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_FE_TipoEmision", "")

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SerieDoc = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + SerieDocElec
                Else
                    SerieDoc = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + SerieDocElec
                End If
                _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_DIGITAL", "")

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SerieDoc = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + SerieDocElec
                Else
                    SerieDoc = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + SerieDocElec
                End If
                _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_ELECTRONICA", "")

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SerieDoc = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + SerieDocElec
                Else
                    SerieDoc = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + SerieDocElec
                End If
                _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_FE_TipoEmision", "")
                If _esElectronico = "" Or _esElectronico = "NAN" Then
                    _esElectronico = ""
                End If

            End If
            If Not String.IsNullOrEmpty(_esElectronico) Then
                If oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then
                    oFuncionesAddon.GuardaLOG(oTipoTabla.ToString, oTransferencia.DocEntry.ToString, " Serie " + oTipoTabla.ToString + " Verificacion serie electronica: query: " + SerieDoc.ToString + " Resultado: " + _esElectronico.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Return True
                Else
                    oFuncionesAddon.GuardaLOG(oTipoTabla.ToString, oDocumento.DocEntry.ToString, " Serie " + oTipoTabla.ToString + " Verificacion serie electronica: " + SerieDoc.ToString + " Resultado: " + _esElectronico.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Return True
                End If

            Else
                If oTipoTabla = "TRE" Or oTipoTabla = "TLE" Then
                    oFuncionesAddon.GuardaLOG(oTipoTabla.ToString, oTransferencia.DocEntry.ToString, "Serie " + oTipoTabla.ToString + " Documento NO Electronico: query: " + SerieDoc.ToString + " Resultado: " + _esElectronico.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Return False
                Else
                    oFuncionesAddon.GuardaLOG(oTipoTabla.ToString, oDocumento.DocEntry.ToString, "Serie " + oTipoTabla.ToString + " Documento NO Electronico: query: " + SerieDoc.ToString + " Resultado: " + _esElectronico.ToString, Functions.FuncionesAddon.Transacciones.Creacion, Functions.FuncionesAddon.TipoLog.Emision)
                    Return False
                End If

            End If
        End If
    End Function

    Public Function ValidarDocElec(ByVal SerieDocElec As String)
        Dim SerieDoc As String = ""
        Dim _esElectronico As String = ""
        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                SerieDoc = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE A.""Series"" = " + SerieDocElec
            Else
                SerieDoc = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE Series = " + SerieDocElec
            End If
            _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_FE_TipoEmision", "")

        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                SerieDoc = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + SerieDocElec
            Else
                SerieDoc = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + SerieDocElec
            End If
            _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_DIGITAL", "")

        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                SerieDoc = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + SerieDocElec
            Else
                SerieDoc = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + SerieDocElec
            End If
            _esElectronico = oFuncionesB1.getRSvalue(SerieDoc, "U_ELECTRONICA", "")

        End If
        Return _esElectronico

    End Function
#End Region

#Region "FUNCIONES Y METODOS"


    Public Sub CreaFomularioParametrizacionUsuarios()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rSboApp, "Param") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\Param.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rSboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rSboApp.Forms.Item("Param").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            CargaFormularioParametrizacionUsuarios()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CargaFormularioParametrizacionUsuarios()

        Dim sQuery As String
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            ' PROBAR SENTENCIA HANA
            sQuery = "SELECT ""USER_CODE"" As Codigo,IFNULL(""U_name"",""USER_CODE"") As Nombre,""USERID"" FROM ""OUSR"" WHERE ""GROUPS"" <> 99  AND ""LOCKED"" = 'N' "
        Else
            sQuery = "SELECT USER_CODE As Codigo,ISNULL(U_name,USER_CODE) As Nombre,USERID FROM OUSR WITH (NOLOCK) WHERE GROUPS <> 99  AND LOCKED = 'N' "
        End If


        Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item("Param")
        mForm.Freeze(True)
        mForm.DataSources.DataTables.Add("dtUser")
        mForm.DataSources.DataTables.Add("dtDpto")
        mForm.DataSources.DataTables.Item("dtUser").ExecuteQuery(sQuery)

        Dim oGrid As SAPbouiCOM.Grid
        oGrid = mForm.Items.Item("gUser").Specific

        oGrid.DataTable = mForm.DataSources.DataTables.Item("dtUser")
        oGrid.Columns.Item(0).Editable = False
        oGrid.Columns.Item(0).TitleObject.Sortable = True
        oGrid.Columns.Item(1).Editable = False
        oGrid.Columns.Item(1).TitleObject.Sortable = True
        oGrid.Columns.Item(2).Visible = False

        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oGrid.AutoResizeColumns()
        'oGrid.Rows.SelectedRows.Add(0)

        sQuery = " SELECT ' TODOS' AS DEPARTAMENTO ,9999 as ID "
        sQuery += "UNION "
        sQuery += "SELECT DISTINCT UPPER(B.Name) AS DEPARTAMENTO,B.Code as ID FROM OUSR A WITH (NOLOCK) "
        sQuery += "INNER JOIN OUDP B WITH (NOLOCK)  ON A.Department = B.Code "
        sQuery += "WHERE A.GROUPS <> 99  AND A.LOCKED = 'N'"

        mForm.DataSources.DataTables.Item("dtDpto").ExecuteQuery(sQuery)
        oGrid = mForm.Items.Item("gDpto").Specific
        oGrid.DataTable = mForm.DataSources.DataTables.Item("dtDpto")
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oGrid.Columns.Item(0).Editable = False
        oGrid.Columns.Item(0).TitleObject.Sortable = True
        oGrid.Columns.Item(1).Editable = False
        oGrid.Columns.Item(1).Visible = False
        oGrid.AutoResizeColumns()
        oGrid.Rows.SelectedRows.Add(0)

        mForm.DataSources.UserDataSources.Add("tckFC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        Dim ckFC As SAPbouiCOM.CheckBox
        ckFC = mForm.Items.Item("ckFC").Specific
        ckFC.ValOn = "S"
        ckFC.ValOff = "N"
        ckFC.DataBind.SetBound(True, "", "tckFC")
        oUserDataSourceFC = mForm.DataSources.UserDataSources.Item("tckFC")
        oUserDataSourceFC.ValueEx = "N"

        mForm.DataSources.UserDataSources.Add("tckNC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        Dim ckNC As SAPbouiCOM.CheckBox
        ckNC = mForm.Items.Item("ckNC").Specific
        ckNC.ValOn = "S"
        ckNC.ValOff = "N"
        ckNC.DataBind.SetBound(True, "", "tckNC")
        oUserDataSourceNC = mForm.DataSources.UserDataSources.Item("tckNC")
        oUserDataSourceNC.ValueEx = "N"

        mForm.DataSources.UserDataSources.Add("tckND", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        Dim ckND As SAPbouiCOM.CheckBox
        ckND = mForm.Items.Item("ckND").Specific
        ckND.ValOn = "S"
        ckND.ValOff = "N"
        ckND.DataBind.SetBound(True, "", "tckND")
        oUserDataSourceND = mForm.DataSources.UserDataSources.Item("tckND")
        oUserDataSourceND.ValueEx = "N"

        mForm.DataSources.UserDataSources.Add("tckGR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        Dim ckGR As SAPbouiCOM.CheckBox
        ckGR = mForm.Items.Item("ckGR").Specific
        ckGR.ValOn = "S"
        ckGR.ValOff = "N"
        ckGR.DataBind.SetBound(True, "", "tckGR")
        oUserDataSourceGR = mForm.DataSources.UserDataSources.Item("tckGR")
        oUserDataSourceGR.ValueEx = "N"

        mForm.DataSources.UserDataSources.Add("tckRE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        Dim ckRE As SAPbouiCOM.CheckBox
        ckRE = mForm.Items.Item("ckRE").Specific
        ckRE.ValOn = "S"
        ckRE.ValOff = "N"
        ckRE.DataBind.SetBound(True, "", "tckRE")
        oUserDataSourceRE = mForm.DataSources.UserDataSources.Item("tckRE")
        oUserDataSourceRE.ValueEx = "N"

        mForm.DataSources.UserDataSources.Add("tckCE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
#Disable Warning BC42024 ' Variable local sin usar: 'ckCE'.
        Dim ckCE As SAPbouiCOM.CheckBox
#Enable Warning BC42024 ' Variable local sin usar: 'ckCE'.
        ckRE = mForm.Items.Item("ckCE").Specific
        ckRE.ValOn = "S"
        ckRE.ValOff = "N"
        ckRE.DataBind.SetBound(True, "", "tckCE")
        oUserDataSourceCE = mForm.DataSources.UserDataSources.Item("tckCE")
        oUserDataSourceCE.ValueEx = "N"

        mForm.Freeze(False)
        mForm.Visible = True

    End Sub

    Private Sub ActualizarParametrizacionUsuario(idUsuario As String, oForm As SAPbouiCOM.Form)
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim oUser As SAPbobsCOM.Users

        Try
            rSboApp.SetStatusBarMessage("Actualizando el Registro, por favor espere..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            oUser = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
            oUser.GetByKey(idUsuario)
            oUser.UserFields.Fields.Item("U_GS_FC").Value = oUserDataSourceFC.Value
            oUser.UserFields.Fields.Item("U_GS_NC").Value = oUserDataSourceNC.Value
            oUser.UserFields.Fields.Item("U_GS_ND").Value = oUserDataSourceND.Value
            oUser.UserFields.Fields.Item("U_GS_GR").Value = oUserDataSourceGR.Value
            oUser.UserFields.Fields.Item("U_GS_RE").Value = oUserDataSourceRE.Value 'IIf(oUserDataSourceRE.Value = "S", "SI", "NO")
            oUser.UserFields.Fields.Item("U_GS_CE").Value = oUserDataSourceCE.Value

            RetVal = oUser.Update()

            'Check the result
            If RetVal <> 0 Then
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                rCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                'MsgBox(ErrCode & " " & ErrMsg)
                rSboApp.SetStatusBarMessage(ErrCode.ToString() + " - " + ErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Else
                Dim obt As SAPbouiCOM.Button
                obt = oForm.Items.Item("btnAct").Specific
                obt.Caption = "OK"
                rSboApp.SetStatusBarMessage("Se ha actualizado el registro..!!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If

        Catch ex As Exception
            rSboApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
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
                    oForm.Select()
                    oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub CargaItemEnFormulario(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String, Optional formUID As String = "")
        Dim codDoc As String = "00"

        Try
            'Dim oFormMode As Integer = 1
            'oFormMode = oForm.Mode

            Dim oSerie As String = ""
            ' OBTENGO EL NUMERO DE SERIE
            If (SeriesElectronicasUDF = "Y" And typeEx = "141") Or (SeriesElectronicasUDF = "Y" And typeEx = "-141") Or (SeriesElectronicasUDF = "Y" And typeEx = "60092") Or (SeriesElectronicasUDF = "Y" And typeEx = "-60092") Then ' si es formulario de factura de proveedores 141 o su formulario de usuario -141
                codDoc = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0).Trim

                If codDoc = "01" Or codDoc = "03" Then
                    EsElectronico = "FE"
                Else
                    EsElectronico = "NA"
                End If

                Dim mycount As Integer = 0
                mycount = oForm.TypeCount
                'oForm = rSboApp.Forms.GetForm("141", mycount)
                If typeEx = "-141" Or typeEx = "141" Then
                    oForm = rSboApp.Forms.GetForm("141", mycount)
                Else
                    oForm = rSboApp.Forms.GetForm("60092", mycount)
                End If

            Else
                If oTipoTabla = "TRE" Then
                    oSerie = oForm.Items.Item("40").Specific.value.ToString()
                ElseIf oTipoTabla = "TLE" Then
                    oSerie = oForm.Items.Item("1250000068").Specific.value.ToString()
                Else
                    oSerie = oForm.Items.Item("88").Specific.value.ToString()
                End If


                Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")
                Dim SQL As String = ""
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            SQL = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE B.""U_TIPO_DOC""='RT' and A.""Series"" = " + oSerie
                        Else
                            SQL = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE B.""U_TIPO_DOC""='RT' and Series = " + oSerie
                        End If
                    Else
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            SQL = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE A.""Series"" = " + oSerie
                        Else
                            SQL = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE Series = " + oSerie
                        End If
                    End If


                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + oSerie
                    Else
                        SQL = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + oSerie
                    End If
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + oSerie
                    Else
                        SQL = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + oSerie
                    End If
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                        Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + oSerie
                    Else
                        SQL = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + oSerie
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                            SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE ""U_TipoD"" IN ('RT','LQRT') AND A.""Series"" = " + oSerie
                        Else
                            SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + oSerie
                        End If

                    Else
                        If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                            SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE ""U_TipoD"" IN ('RT','LQRT') AND Series = " + oSerie
                        Else
                            SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + oSerie
                        End If

                    End If
                End If

                Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + SQL, "EventosEmision")
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_DIGITAL", "")

                    If EsElectronico = "Y" Then
                        EsElectronico = "FE"
                    Else
                        EsElectronico = "NA"
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
                    If EsElectronico = "SI" Then
                        EsElectronico = "FE"
                    Else
                        EsElectronico = "NA"
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
                    If EsElectronico = "SI" Then
                        EsElectronico = "FE"
                    Else
                        EsElectronico = "NA"
                    End If
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")
                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                        If EsElectronico = "RT" Or EsElectronico = "LQRT" Then
                            EsElectronico = "FE"
                        ElseIf EsElectronico = "NAN" Or EsElectronico = "" Then
                            EsElectronico = "NA"
                        Else
                            EsElectronico = "NA"
                        End If
                    Else
                        If EsElectronico = "NAN" Or EsElectronico = "" Then
                            EsElectronico = "NA"
                        Else
                            EsElectronico = "FE"
                        End If
                    End If
                End If
                Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "EventosEmision")
            End If

            'FRIZEO ANTES DE CUALQUIER DIBUJADO

            oForm.Freeze(True)

            If evento = "NUEVO" Then

                lbEstado = oForm.Items.Item("lbEstado").Specific
                lbEstado.Item.ForeColor = RGB(204, 0, 0)
                lbEstado.Caption = "Estado: NO ENVIADO"
                lbEstado.Item.Visible = True

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                lbComentario = oForm.Items.Item("lbComen").Specific

                If EsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(7, 118, 10)
                    ' lbComentario.Caption = "ELECTRONICO"
                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO NO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(204, 0, 0)
                End If
                lbComentario.Item.Visible = True
                If (oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM") And EsElectronico <> "FE" Then
                    lbComentario.Item.Visible = False
                    lbEstado.Item.Visible = False
                End If
                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("NUEVO-Eliminacion valores validos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                ' se trata de crear un codigo QR





            ElseIf evento = "ACTUALIZAR" Then

                lbComentario = oForm.Items.Item("lbComen").Specific
                lbEstado = oForm.Items.Item("lbEstado").Specific
                'cbEstado = oForm.Items.Item("cbEstado").Specific
                'btnAccion = oForm.Items.Item("btnAccion").Specific

                'Obtengo valor almacenado en base
                Dim UDFEA As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_ESTADO_AUTORIZACIO", 0)

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEA) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEA = "0"
                End If
                UDFEA = UDFEA.Trim

                Try
                    If EsElectronico = "FE" Then
                        If UDFEA = "2" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: AUTORIZADO"

                        ElseIf UDFEA = "5" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: RECIBIDO"

                        ElseIf UDFEA = "7" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: ERROR RECEPCION SRI"

                        ElseIf UDFEA = "4" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: VALIDAR DATOS"

                        ElseIf UDFEA = "3" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO AUTORIZADA"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEA = "11" Then
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = False
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: ANULADO"

                        Else
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO ENVIADO"

                        End If

                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Actualizar-FE ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If EsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(7, 118, 10)

                    Select Case oForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            lbComentario.Item.Visible = True
                            lbEstado.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Then
                                If UDFEA = "0" Then
                                    btnAccion.Item.Visible = False
                                End If
                            End If

                        Case Else
                            btnAccion.Item.Visible = False

                    End Select

                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO NO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(204, 0, 0)

                    lbEstado.Caption = "Estado: NO ENVIADO"
                    lbEstado.Item.ForeColor = RGB(204, 0, 0)
                    btnAccion.Item.Visible = False

                    If oTipoTabla = "REE" Then
                        lbComentario.Item.Visible = False
                        lbEstado.Item.Visible = False
                    End If


                End If


                'se intenta crear imagenes QR

                Try

                    If oForm.PaneLevel = 50 Then

                        ObtenerEnlacesURLyGenerarRQ(oForm, formUID)

                    End If
                Catch ex As Exception

                End Try


                If Functions.VariablesGlobales._ActivarLocalizacionEC = "Y" Then

                    'se oculta folios generados si no es modo ADD

                    Try

                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            oForm.Items.Item("ptexto").Visible = False
                            oForm.Items.Item("pfolio").Visible = False

                        End If

                    Catch ex As Exception

                    End Try

                    'Solo para mostrar el folio de la solicitud
                    Try

                        If oForm.TypeEx = "1250000940" Then

                            Dim foliosolicitud As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("FolioNum", 0).Trim

                            oForm.Items.Item("etssloc80").Enabled = False
                            oForm.Items.Item("etssloc80").Specific.value = foliosolicitud

                        End If


                    Catch ex As Exception

                    End Try

                End If



            End If

        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oFrmUser.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Private Sub CargaItemEnFormularioEcuanexus(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String)
        Dim codDoc As String = "00"

        Try
            'Dim oFormMode As Integer = 1
            'oFormMode = oForm.Mode

            Dim oSerie As String = ""
            ' OBTENGO EL NUMERO DE SERIE
            If (SeriesElectronicasUDF = "Y" And typeEx = "141") Or (SeriesElectronicasUDF = "Y" And typeEx = "-141") Or (SeriesElectronicasUDF = "Y" And typeEx = "60092") Or (SeriesElectronicasUDF = "Y" And typeEx = "-60092") Then ' si es formulario de factura de proveedores 141 o su formulario de usuario -141
                codDoc = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0).Trim

                If codDoc = "01" Or codDoc = "03" Then
                    EsElectronico = "FE"
                Else
                    EsElectronico = "NA"
                End If

                Dim mycount As Integer = 0
                mycount = oForm.TypeCount
                'oForm = rSboApp.Forms.GetForm("141", mycount)
                If typeEx = "-141" Or typeEx = "141" Then
                    oForm = rSboApp.Forms.GetForm("141", mycount)
                Else
                    oForm = rSboApp.Forms.GetForm("60092", mycount)
                End If

            Else
                If oTipoTabla = "TRE" Then
                    oSerie = oForm.Items.Item("40").Specific.value.ToString()
                ElseIf oTipoTabla = "TLE" Then
                    oSerie = oForm.Items.Item("1250000068").Specific.value.ToString()
                Else
                    oSerie = oForm.Items.Item("88").Specific.value.ToString()
                End If


                Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")
                Dim SQL As String = ""
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            SQL = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE B.""U_TIPO_DOC""='RT' and A.""Series"" = " + oSerie
                        Else
                            SQL = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE B.""U_TIPO_DOC""='RT' and Series = " + oSerie
                        End If
                    Else
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            SQL = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE A.""Series"" = " + oSerie
                        Else
                            SQL = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE Series = " + oSerie
                        End If
                    End If


                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + oSerie
                    Else
                        SQL = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + oSerie
                    End If
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + oSerie
                    Else
                        SQL = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + oSerie
                    End If
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                        Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + oSerie
                    Else
                        SQL = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + oSerie
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                            SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE ""U_TipoD"" IN ('RT','LQRT') AND A.""Series"" = " + oSerie
                        Else
                            SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + oSerie
                        End If

                    Else
                        If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                            SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE ""U_TipoD"" IN ('RT','LQRT') AND Series = " + oSerie
                        Else
                            SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + oSerie
                        End If

                    End If
                End If

                Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + SQL, "EventosEmision")
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_DIGITAL", "")

                    If EsElectronico = "Y" Then
                        EsElectronico = "FE"
                    Else
                        EsElectronico = "NA"
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
                    If EsElectronico = "SI" Then
                        EsElectronico = "FE"
                    Else
                        EsElectronico = "NA"
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
                    If EsElectronico = "SI" Then
                        EsElectronico = "FE"
                    Else
                        EsElectronico = "NA"
                    End If
                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "NAN")
                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                        If EsElectronico = "RT" Or EsElectronico = "LQRT" Then
                            EsElectronico = "FE"
                        ElseIf EsElectronico = "NAN" Or EsElectronico = "" Then
                            EsElectronico = "NA"
                        Else
                            EsElectronico = "NA"
                        End If
                    Else
                        If EsElectronico = "NAN" Or EsElectronico = "" Then
                            EsElectronico = "NA"
                        Else
                            EsElectronico = "FE"
                        End If
                    End If

                End If
                Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "EventosEmision")
            End If

            'FRIZEO ANTES DE CUALQUIER DIBUJADO

            oForm.Freeze(True)

            If evento = "NUEVO" Then

                lbEstado = oForm.Items.Item("lbEstado").Specific
                lbEstado.Item.ForeColor = RGB(204, 0, 0)
                lbEstado.Caption = "Estado: NO ENVIADO"
                lbEstado.Item.Visible = True

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                lbComentario = oForm.Items.Item("lbComen").Specific

                If EsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(7, 118, 10)
                    ' lbComentario.Caption = "ELECTRONICO"
                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO NO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(204, 0, 0)
                End If
                lbComentario.Item.Visible = True
                If (oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM") And EsElectronico <> "FE" Then
                    lbComentario.Item.Visible = False
                    lbEstado.Item.Visible = False
                End If
                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("NUEVO-Eliminacion valores validos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

            ElseIf evento = "ACTUALIZAR" Then

                lbComentario = oForm.Items.Item("lbComen").Specific
                lbEstado = oForm.Items.Item("lbEstado").Specific
                'cbEstado = oForm.Items.Item("cbEstado").Specific
                'btnAccion = oForm.Items.Item("btnAccion").Specific

                'Obtengo valor almacenado en base
                Dim UDFEA As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_ESTADO_AUTORIZACIO", 0)

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEA) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEA = "0"
                End If
                UDFEA = UDFEA.Trim

                Try
                    If EsElectronico = "FE" Then
                        If UDFEA = "1" Then
                            'btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            'btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Estado")
                            btnAccion.Select("(GS) Consultar AUT", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            'lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: EN PROCESO"

                        ElseIf UDFEA = "2" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            'btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar Documento", "(GS) Reenviar Documento")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: AUTORIZADA"

                        ElseIf UDFEA = "4" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar Doc")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO ENCONTRADO"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar Doc")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: CON ERROR"

                        ElseIf UDFEA = "11" Then
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = False
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: ANULADO"

                        Else
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar Doc")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO ENVIADO"

                        End If

                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Actualizar-FE ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If EsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(7, 118, 10)

                    Select Case oForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            lbComentario.Item.Visible = True
                            lbEstado.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Then
                                If UDFEA = "0" Then
                                    btnAccion.Item.Visible = False
                                End If
                            End If

                        Case Else
                            btnAccion.Item.Visible = False

                    End Select

                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO NO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(204, 0, 0)

                    lbEstado.Caption = "Estado: NO ENVIADO"
                    lbEstado.Item.ForeColor = RGB(204, 0, 0)
                    btnAccion.Item.Visible = False

                    If (oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM") Then
                        lbComentario.Item.Visible = False
                        lbEstado.Item.Visible = False
                    End If

                End If

            End If

        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oFrmUser.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub


    Private Sub CargaItemEnFormulario_LiquidacionCompra(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String)
        Dim codDoc As String = "00"

        Try

            If (SeriesElectronicasUDF = "Y" And typeEx = "141") Or (SeriesElectronicasUDF = "Y" And typeEx = "-141") Or (SeriesElectronicasUDF = "Y" And typeEx = "60092") Or (SeriesElectronicasUDF = "Y" And typeEx = "-60092") Then ' si es formulario de factura de proveedores 141 o su formulario de usuario -141

                codDoc = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0).Trim

                If codDoc = "02" Or codDoc = "03" Then
                    LQEsElectronico = "FE"
                Else
                    LQEsElectronico = "NA"
                End If

                Dim mycount As Integer = 0
                mycount = oForm.TypeCount
                If typeEx = "-141" Or typeEx = "141" Then
                    oForm = rSboApp.Forms.GetForm("141", mycount)
                Else
                    oForm = rSboApp.Forms.GetForm("60092", mycount)
                End If
                'oForm = rSboApp.Forms.GetForm(typeEx, mycount)

            Else

                Dim oSerie As String = ""
                oSerie = oForm.Items.Item("88").Specific.value.ToString()
                Dim SQL As String = ""

                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                    Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE B.""U_TipoD"" IN ('LQ','LQRT') AND A.""Series"" = " + oSerie
                    Else
                        SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE B.U_TipoD IN ('LQ','LQRT') AND Series = " + oSerie
                    End If

                    LQEsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")

                    If LQEsElectronico = "NAN" Or LQEsElectronico = "" Then
                        LQEsElectronico = "NA"
                    ElseIf LQEsElectronico = "LQ" Or LQEsElectronico = "LQRT" Then
                        LQEsElectronico = "FE"
                    Else
                        LQEsElectronico = "NA"
                    End If

                    Utilitario.Util_Log.Escribir_Log("QUERY LQ: " + SQL.ToString + " RESPUESTA: " + LQEsElectronico.ToString, "EventosEmision")

                Else
                    Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")


                    SQL = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + oSerie

                    LQEsElectronico = oFuncionesB1.getRSvalue(SQL, "Code", "")

                    If Not String.IsNullOrEmpty(LQEsElectronico) Then
                        LQEsElectronico = "FE"
                    Else
                        LQEsElectronico = "NA"
                    End If

                    Utilitario.Util_Log.Escribir_Log("QUERY LQ: " + SQL.ToString + " RESPUESTA: " + LQEsElectronico.ToString, "EventosEmision")
                End If

            End If

            oForm.Freeze(True)
            If evento = "NUEVO" Then

                lbEstadoLQ = oForm.Items.Item("lbEstadoLQ").Specific
                lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                lbEstadoLQ.Caption = "Estado: NO ENVIADO"
                lbEstadoLQ.Item.Visible = True

                'U_LQ_ESTADO

                btnAccLQ = oForm.Items.Item("btnAccLQ").Specific
                btnAccLQ.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccLQ.Item.AffectsFormMode = False
                btnAccLQ.Item.Visible = False


                lbComentarioLQ = oForm.Items.Item("lbComenLQ").Specific
                If LQEsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    lbComentarioLQ.Caption = "(GS) LIQUIDACIÓN ELECTRÓNICA"
                    lbComentarioLQ.Item.ForeColor = RGB(7, 118, 10)
                    ' lbComentario.Caption = "ELECTRONICO"
                ElseIf LQEsElectronico <> "FE" Then
                    lbComentarioLQ.Caption = "(GS) LIQUIDACIÓN NO ELECTRÓNICA"
                    lbComentarioLQ.Item.ForeColor = RGB(204, 0, 0)


                Else
                    oForm.Title = nombreFormulario
                    lbComentarioLQ.Caption = "(GS) LIQUIDACIÓN NO ELECTRÓNICA"
                    lbComentarioLQ.Item.ForeColor = RGB(204, 0, 0)
                End If
                lbComentarioLQ.Item.Visible = True
                If LQEsElectronico <> "FE" Then
                    lbComentarioLQ.Item.Visible = False
                    lbEstadoLQ.Item.Visible = False
                End If
                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccLQ.ValidValues.Count > 0 Then
                        For i As Integer = btnAccLQ.ValidValues.Count - 1 To 0 Step -1
                            btnAccLQ.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("NUEVO-Eliminacion valores validos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try



            ElseIf evento = "ACTUALIZAR" Then

                lbComentarioLQ = oForm.Items.Item("lbComenLQ").Specific
                lbEstadoLQ = oForm.Items.Item("lbEstadoLQ").Specific
                'cbEstadoLQ = oForm.Items.Item("cbEstLQ").Specific
                'btnAccLQ = oForm.Items.Item("btnAccLQ").Specific

                'Obtengo valor almacenado en base
                Dim UDFEALQ As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_ESTADO", 0)

                btnAccLQ = oForm.Items.Item("btnAccLQ").Specific
                btnAccLQ.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccLQ.Item.AffectsFormMode = False
                btnAccLQ.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccLQ.ValidValues.Count > 0 Then
                        For i As Integer = btnAccLQ.ValidValues.Count - 1 To 0 Step -1
                            btnAccLQ.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEALQ) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEALQ = "0"
                End If
                UDFEALQ = UDFEALQ.Trim

                Try
                    If LQEsElectronico = "FE" Then
                        If UDFEALQ = "2" Then
                            btnAccLQ.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccLQ.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccLQ.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            lbEstadoLQ.Item.ForeColor = RGB(7, 118, 10)
                            lbEstadoLQ.Caption = "Estado: AUTORIZADO"

                        ElseIf UDFEALQ = "5" Then
                            btnAccLQ.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccLQ.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccLQ.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccLQ.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            lbEstadoLQ.Item.ForeColor = RGB(7, 118, 10)
                            lbEstadoLQ.Caption = "Estado: RECIBIDO"

                        ElseIf UDFEALQ = "7" Then
                            btnAccLQ.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccLQ.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccLQ.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccLQ.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            lbEstadoLQ.Item.ForeColor = RGB(7, 118, 10)
                            lbEstadoLQ.Caption = "Estado: ERROR RECEPCION SRI"

                        ElseIf UDFEALQ = "4" Then
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: VALIDAR DATOS"

                        ElseIf UDFEALQ = "3" Then
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: NO AUTORIZADA"

                        ElseIf UDFEALQ = "6" Then
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEALQ = "11" Then
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccLQ.Item.Visible = False
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: ANULADO"

                        Else
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: NO ENVIADO"

                        End If
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Actualizar-FE ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If LQEsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    lbComentarioLQ.Caption = "GS LQ ELECTRÓNICO"
                    lbComentarioLQ.Item.ForeColor = RGB(7, 118, 10)

                    Select Case oForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            btnAccLQ.Item.Visible = True
                            lbComentarioLQ.Item.Visible = True
                            lbEstadoLQ.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Or UDFEALQ = "11" Then
                                If UDFEALQ = "0" Or UDFEALQ = "11" Then
                                    btnAccLQ.Item.Visible = False
                                End If
                            End If

                        Case Else
                            btnAccLQ.Item.Visible = False

                    End Select


                Else
                    lbComentarioLQ.Caption = "GS DOCUMENTO NO ELECTRÓNICO"
                    lbComentarioLQ.Item.ForeColor = RGB(204, 0, 0)

                    lbEstadoLQ.Caption = "Estado: NO ENVIADO"
                    lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                    btnAccLQ.Item.Visible = False


                    lbEstadoLQ.Item.Visible = False
                    lbComentarioLQ.Item.Visible = False


                End If

            End If

        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Private Sub CargaItemEnFormulario_LiquidacionCompraEcuanexus(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String)
        Dim codDoc As String = "00"

        Try

            If (SeriesElectronicasUDF = "Y" And typeEx = "141") Or (SeriesElectronicasUDF = "Y" And typeEx = "-141") Or (SeriesElectronicasUDF = "Y" And typeEx = "60092") Or (SeriesElectronicasUDF = "Y" And typeEx = "-60092") Then ' si es formulario de factura de proveedores 141 o su formulario de usuario -141

                codDoc = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0).Trim

                If codDoc = "02" Or codDoc = "03" Then
                    LQEsElectronico = "FE"
                Else
                    LQEsElectronico = "NA"
                End If

                Dim mycount As Integer = 0
                mycount = oForm.TypeCount
                If typeEx = "-141" Or typeEx = "141" Then
                    oForm = rSboApp.Forms.GetForm("141", mycount)
                Else
                    oForm = rSboApp.Forms.GetForm("60092", mycount)
                End If
                'oForm = rSboApp.Forms.GetForm(typeEx, mycount)

            Else
                Dim oSerie As String = ""
                oSerie = oForm.Items.Item("88").Specific.value.ToString()
                Dim SQL As String = ""

                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                    Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE B.""U_TipoD"" IN ('LQ','LQRT') AND A.""Series"" = " + oSerie
                    Else
                        SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE B.U_TipoD IN ('LQ','LQRT') AND Series = " + oSerie
                    End If

                    LQEsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")

                    If LQEsElectronico = "NAN" Or LQEsElectronico = "" Then
                        LQEsElectronico = "NA"
                    ElseIf LQEsElectronico = "LQ" Or LQEsElectronico = "LQRT" Then
                        LQEsElectronico = "FE"
                    Else
                        LQEsElectronico = "NA"
                    End If

                    Utilitario.Util_Log.Escribir_Log("QUERY LQ: " + SQL.ToString + " RESPUESTA: " + LQEsElectronico.ToString, "EventosEmision")

                Else
                    Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")


                    SQL = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + oSerie

                    LQEsElectronico = oFuncionesB1.getRSvalue(SQL, "Code", "")

                    If Not String.IsNullOrEmpty(LQEsElectronico) Then
                        LQEsElectronico = "FE"
                    Else
                        LQEsElectronico = "NA"
                    End If

                    Utilitario.Util_Log.Escribir_Log("QUERY LQ: " + SQL.ToString + " RESPUESTA: " + LQEsElectronico.ToString, "EventosEmision")
                End If
            End If

            oForm.Freeze(True)
            If evento = "NUEVO" Then

                lbEstadoLQ = oForm.Items.Item("lbEstadoLQ").Specific
                lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                lbEstadoLQ.Caption = "Estado: NO ENVIADO"
                lbEstadoLQ.Item.Visible = True

                'U_LQ_ESTADO

                btnAccLQ = oForm.Items.Item("btnAccLQ").Specific
                btnAccLQ.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccLQ.Item.AffectsFormMode = False
                btnAccLQ.Item.Visible = False


                lbComentarioLQ = oForm.Items.Item("lbComenLQ").Specific
                If LQEsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    lbComentarioLQ.Caption = "(GS) LIQUIDACIÓN ELECTRÓNICA"
                    lbComentarioLQ.Item.ForeColor = RGB(7, 118, 10)
                    ' lbComentario.Caption = "ELECTRONICO"
                ElseIf EsElectronico = "FE" Then
                    'oForm.Title += " - ELECTRONICO"
                    If LQEsElectronico <> "FE" Then
                        lbComentarioLQ.Caption = "(GS) LIQUIDACIÓN NO ELECTRÓNICA"
                        lbComentarioLQ.Item.ForeColor = RGB(204, 0, 0)
                    End If

                Else
                    oForm.Title = nombreFormulario
                    lbComentarioLQ.Caption = "(GS) LIQUIDACIÓN NO ELECTRÓNICA"
                    lbComentarioLQ.Item.ForeColor = RGB(204, 0, 0)
                End If
                lbComentarioLQ.Item.Visible = True
                If LQEsElectronico <> "FE" Then
                    lbComentarioLQ.Item.Visible = False
                    lbEstadoLQ.Item.Visible = False
                End If
                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccLQ.ValidValues.Count > 0 Then
                        For i As Integer = btnAccLQ.ValidValues.Count - 1 To 0 Step -1
                            btnAccLQ.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("NUEVO-Eliminacion valores validos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try



            ElseIf evento = "ACTUALIZAR" Then

                lbComentarioLQ = oForm.Items.Item("lbComenLQ").Specific
                lbEstadoLQ = oForm.Items.Item("lbEstadoLQ").Specific
                'cbEstadoLQ = oForm.Items.Item("cbEstLQ").Specific
                'btnAccLQ = oForm.Items.Item("btnAccLQ").Specific

                'Obtengo valor almacenado en base
                Dim UDFEALQ As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_ESTADO", 0)

                btnAccLQ = oForm.Items.Item("btnAccLQ").Specific
                btnAccLQ.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccLQ.Item.AffectsFormMode = False
                btnAccLQ.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccLQ.ValidValues.Count > 0 Then
                        For i As Integer = btnAccLQ.ValidValues.Count - 1 To 0 Step -1
                            btnAccLQ.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEALQ) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEALQ = "0"
                End If
                UDFEALQ = UDFEALQ.Trim

                Try
                    If LQEsElectronico = "FE" Then
                        If UDFEALQ = "1" Then
                            'btnAccLQ.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            'btnAccLQ.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccLQ.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Estado")
                            btnAccLQ.Select("(GS) Consultar AUT", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: EN PROCESO"

                        ElseIf UDFEALQ = "2" Then
                            btnAccLQ.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            'btnAccLQ.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccLQ.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccLQ.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            lbEstadoLQ.Item.ForeColor = RGB(7, 118, 10)
                            lbEstadoLQ.Caption = "Estado: AUTORIZADO"

                        ElseIf UDFEALQ = "4" Then
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: NO ENCONTRADO"

                        ElseIf UDFEALQ = "6" Then
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: CON ERROR"

                        ElseIf UDFEALQ = "11" Then
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccLQ.Item.Visible = False
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: ANULADO"

                        Else
                            btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                            lbEstadoLQ.Caption = "Estado: NO ENVIADO"

                        End If
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Actualizar-FE ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If LQEsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    lbComentarioLQ.Caption = "GS LQ ELECTRÓNICO"
                    lbComentarioLQ.Item.ForeColor = RGB(7, 118, 10)

                    Select Case oForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            btnAccLQ.Item.Visible = True
                            lbComentarioLQ.Item.Visible = True
                            lbEstadoLQ.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Or UDFEALQ = "11" Then
                                If UDFEALQ = "0" Or UDFEALQ = "11" Then
                                    btnAccLQ.Item.Visible = False
                                End If
                            End If

                        Case Else
                            btnAccLQ.Item.Visible = False

                    End Select


                Else
                    lbComentarioLQ.Caption = "GS LQ NO ELECTRÓNICO"
                    lbComentarioLQ.Item.ForeColor = RGB(204, 0, 0)

                    lbEstadoLQ.Caption = "Estado: NO ENVIADO"
                    lbEstadoLQ.Item.ForeColor = RGB(204, 0, 0)
                    btnAccLQ.Item.Visible = False

                    lbComentarioLQ.Item.Visible = False
                    lbEstadoLQ.Item.Visible = False

                End If

            End If

        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Private Sub CargaItemEnFormulario_Factura_GuiaRemision(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String)


        Try


            oForm.Freeze(True)
            If evento = "ACTUALIZAR" Then

                lbComentarioGR = oForm.Items.Item("lbComenGR").Specific
                lbEstadoGR = oForm.Items.Item("lbEstadoGR").Specific
                'cbEstadoLQ = oForm.Items.Item("cbEstLQ").Specific
                'btnAccLQ = oForm.Items.Item("btnAccLQ").Specific

                'Obtengo valor almacenado en base
                Dim UDFEAGR As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_GR_ESTADO", 0)

                btnAccGR = oForm.Items.Item("btnAccGR").Specific
                btnAccGR.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccGR.Item.AffectsFormMode = False
                btnAccGR.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccGR.ValidValues.Count > 0 Then
                        For i As Integer = btnAccGR.ValidValues.Count - 1 To 0 Step -1
                            btnAccGR.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEAGR) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEAGR = "0"
                End If
                UDFEAGR = UDFEAGR.Trim

                Try

                    If UDFEAGR = "2" Then
                        btnAccGR.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                        btnAccGR.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                        btnAccGR.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                        btnAccGR.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        lbEstadoGR.Item.ForeColor = RGB(7, 118, 10)
                        lbEstadoGR.Caption = "Estado: AUTORIZADO"

                    ElseIf UDFEAGR = "5" Then
                        btnAccGR.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                        btnAccGR.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        btnAccGR.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                        'btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                        btnAccGR.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        lbEstadoGR.Item.ForeColor = RGB(7, 118, 10)
                        lbEstadoGR.Caption = "Estado: RECIBIDO"

                    ElseIf UDFEAGR = "7" Then
                        btnAccGR.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                        btnAccGR.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        btnAccGR.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                        'btnAccLQ.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                        btnAccGR.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        lbEstadoGR.Item.ForeColor = RGB(7, 118, 10)
                        lbEstadoGR.Caption = "Estado: ERROR RECEPCION SRI"

                    ElseIf UDFEAGR = "4" Then
                        btnAccGR.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        btnAccGR.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        lbEstadoGR.Item.ForeColor = RGB(204, 0, 0)
                        lbEstadoGR.Caption = "Estado: VALIDAR DATOS"

                    ElseIf UDFEAGR = "3" Then
                        btnAccGR.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        btnAccGR.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        lbEstadoGR.Item.ForeColor = RGB(204, 0, 0)
                        lbEstadoGR.Caption = "Estado: NO AUTORIZADA"

                    ElseIf UDFEAGR = "6" Then
                        btnAccGR.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        btnAccGR.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        lbEstadoGR.Item.ForeColor = RGB(204, 0, 0)
                        lbEstadoGR.Caption = "Estado: DEVUELTA"

                    ElseIf UDFEAGR = "11" Then
                        'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        btnAccGR.Item.Visible = False
                        lbEstadoGR.Item.ForeColor = RGB(204, 0, 0)
                        lbEstadoGR.Caption = "Estado: ANULADO"

                    Else
                        btnAccGR.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                        btnAccGR.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        lbEstadoGR.Item.ForeColor = RGB(204, 0, 0)
                        lbEstadoGR.Caption = "Estado: NO ENVIADO"

                    End If

                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Actualizar-FE GR", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try




                Select Case oForm.Mode
                    Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        If Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Or Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI" Then
                            btnAccGR.Item.Visible = True
                            lbComentarioGR.Caption = "GS GR ELECTRÓNICO"
                            lbComentarioGR.Item.ForeColor = RGB(7, 118, 10)
                            lbComentarioGR.Item.Visible = True
                            lbEstadoGR.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Or UDFEAGR = "11" Then
                                If UDFEAGR = "0" Or UDFEAGR = "11" Then
                                    btnAccGR.Item.Visible = False
                                End If
                            End If
                        Else
                            Try

                                lbComentarioGR.Item.Visible = False
                                lbEstadoGR.Item.Visible = False
                                btnAccGR.Item.Visible = False
                            Catch ex As Exception

                            End Try

                        End If


                    Case Else
                        btnAccLQ.Item.Visible = False

                End Select

            End If


        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Private Sub creaBotonPrueba(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try
            Dim _left As Integer = 0, _top As Integer = 0, _espaciado As Integer = 15

            Dim ItemParaLinkeo As String = "5"

            If oTipoTabla = "TRE" Then
                _left = oForm.Items.Item("9").Left + 135

            ElseIf oTipoTabla = "TLE" Then
                _left = oForm.Items.Item("9").Left + 135
            Else
                _left = oForm.Items.Item("14").Left
            End If


            ' COMENTARIO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbComen", "(GS) Documento Electrónico", _left, 85, 250, 14, 0, False)
                oForm.Items.Item("lbComen").LinkTo = ItemParaLinkeo

            Catch ex As Exception
                rSboApp.SetStatusBarMessage("No se pudo crear la label de Comentario", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try


            ' ESTADO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbEstado", "(GS) Estado", _left, 100, 250, 14, 0, False)
                oForm.Items.Item("lbEstado").LinkTo = ItemParaLinkeo

            Catch ex As Exception
                rSboApp.SetStatusBarMessage("No se pudo crear la label de estado", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

            ' Boton COmbo
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO, "btnAccion", "", _left, 115, 110, 19, 0, False)
                oForm.Items.Item("btnAccion").LinkTo = ItemParaLinkeo
            Catch ex As Exception
                rSboApp.SetStatusBarMessage("No se pudo crear El boton combo Accion", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try



        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Botones Prueba ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)

    End Sub

    Private Sub creaBotonPrueba_LiquidacionCompra(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)

        Try
            Dim _left As Integer = 0, _top As Integer = 0, _espaciado As Integer = 15


            _left = oForm.Items.Item("14").Left + 160


            ' Boton COmbo
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO, "btnAccLQ", "", _left, 115, 110, 19, 0, False)
            Catch ex As Exception
                rSboApp.SetStatusBarMessage("No se pudo crear El boton combo Accion", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

            ' COMENTARIO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbComenLQ", "(GS) Liquidación Electrónica", _left, 85, 250, 14, 0, False)
            Catch ex As Exception
            End Try
            lbComentarioLQ = oForm.Items.Item("lbComenLQ").Specific

            ' ESTADO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbEstadoLQ", "(GS) Estado Liquidación", _left, 100, 250, 14, 0, False)
            Catch ex As Exception
                rSboApp.SetStatusBarMessage("NUEVO-Creacion label estado", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try


        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Botones Prueba ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Private Sub creaBotonPrueba_Factura_GuiaRemision(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)

        Try
            Dim _left As Integer = 0, _top As Integer = 0, _espaciado As Integer = 15


            _left = oForm.Items.Item("14").Left + 160

            ' Boton COmbo
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO, "btnAccGR", "", _left, 115, 110, 19, 0, False)
            Catch ex As Exception
                rSboApp.SetStatusBarMessage("No se pudo crear El boton combo Accion GR", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

            ' COMENTARIO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbComenGR", "(GS) Guia Remisión Electrónica", _left, 85, 250, 14, 0, False)
            Catch ex As Exception
            End Try
            lbComentarioGR = oForm.Items.Item("lbComenGR").Specific

            ' ESTADO
            Try
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lbEstadoGR", "(GS) Estado Guia Remisión", _left, 100, 250, 14, 0, False)
            Catch ex As Exception
                rSboApp.SetStatusBarMessage("NUEVO-Creacion label estado GR", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try


        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Botones Prueba GR", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Function Fechas(ByVal fecha As Date) As String
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Try
            mRst = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sqlstr = "select ""DateFormat"" from ""OADM"""
            Else
                sqlstr = "select DateFormat from OADM"
            End If

            mRst.DoQuery(sqlstr)
            If mRst.EoF = False Then

                Select Case mRst.Fields.Item("DateFormat").Value
                    Case 0
                        Return Format(fecha, "dd/MM/yy")
                    Case 1
                        Return Format(fecha, "dd/MM/YYYY")
                    Case 2
                        Return Format(fecha, "MM/dd/yy")
                    Case 3
                        Return Format(fecha, "MM/dd/yyyy")
                    Case 4
                        Return Format(fecha, "yyyy/MM/dd")
                    Case 5
                        Return Format(fecha, "dd/MM/yyyy")
                    Case 6
                        Return Format(fecha, "yy/MM/dd")
                End Select

            End If

        Catch ex As Exception
            Return ""
        End Try

#Disable Warning BC42105 ' La función 'Fechas' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.
    End Function
#Enable Warning BC42105 ' La función 'Fechas' no devuelve un valor en todas las rutas de acceso de código. Puede producirse una excepción de referencia NULL en tiempo de ejecución cuando se use el resultado.

    Private Sub SetearCamposAutoFE(ByVal pform As SAPbouiCOM.Form, ByVal ItemInvocador As String, Optional ByRef IdForm As String = "")
        If pform.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_PRINT_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then Return

        Try
            Dim usform = rSboApp.Forms.Item(pform.UDFFormUID)

            If IdForm <> "141" Then
                Utilitario.Util_Log.Escribir_Log("Inicio de seteo ", "EventosEmision")
                usform.Items.Item("U_ESTADO_AUTORIZACIO").Specific.Select("0")
                usform.Items.Item("U_CLAVE_ACCESO").Specific.String = ""
                usform.Items.Item("U_NUM_AUTO_FAC").Specific.String = ""
                usform.Items.Item("U_OBSERVACION_FACT").Specific.String = ""
                usform.Items.Item("U_FECHA_AUT_FACT").Specific.String = ""
                Utilitario.Util_Log.Escribir_Log("Inicio de seteo ", "EventosEmision")
            Else
                Utilitario.Util_Log.Escribir_Log("Inicio de seteo ", "EventosEmision")
                usform.Items.Item("U_ESTADO_AUTORIZACIO").Specific.Select("0")
                usform.Items.Item("U_CLAVE_ACCESO").Specific.String = ""
                usform.Items.Item("U_NUM_AUTO_FAC").Specific.String = ""
                usform.Items.Item("U_OBSERVACION_FACT").Specific.String = ""
                usform.Items.Item("U_FECHA_AUT_FACT").Specific.String = ""
                usform.Items.Item("U_LQ_ESTADO").Specific.Select("0")
                usform.Items.Item("U_LQ_CLAVE").Specific.String = ""
                usform.Items.Item("U_LQ_NUM_AUTO").Specific.String = ""
                usform.Items.Item("U_LQ_OBSERVACION").Specific.String = ""
                usform.Items.Item("U_LQ_FECHA_AUT").Specific.String = ""
                Utilitario.Util_Log.Escribir_Log("Fin de seteo ", "EventosEmision")
            End If


        Catch ex As Exception
            rSboApp.SetStatusBarMessage(ex.Message.ToString)
            Utilitario.Util_Log.Escribir_Log("Error al setear campos de autorizacion a vacio" & ex.Message.ToString, "EventosEmision")
        End Try

        Try
            pform.Refresh()
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al Liberar o Refrescar from " & ex.Message.ToString, "EventosEmision")

        End Try
    End Sub


    Private Sub SetearCamposAutoFE_LOC(ByRef pform As SAPbouiCOM.Form, ByVal ItemInvocador As String, ByVal IdFormDM As String)

        If Functions.VariablesGlobales._SS_SetearCamposUsuario = "Y" Then


            If pform.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_PRINT_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or pform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then Return
            'If pform.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE And pform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pform.Mode = SAPbouiCOM.BoFormMode.fm_PRINT_MODE And pform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE And pform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then Return

            Try

                Dim est As String = ""
                Dim pm As String = ""
                Dim serie As String = ""
                Dim tipoComp As String = ""
                Dim docmun As String = ""
                Dim SerieDoc As String = ""

                Dim rs As SAPbobsCOM.Recordset
                Dim rs1 As SAPbobsCOM.Recordset

                est = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Est", 0).Trim
                pm = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Pemi", 0).Trim

                serie = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                tipoComp = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim

                SerieDoc = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                If ItemInvocador = "DUPLICAR" Then

                    Dim m2query As String = "Select ""NextNumber"" as sigiente from ""NNM1"" where ""Series"" = '" & serie & "'"

                    rs = oFuncionesB1.getRecordSet(m2query)

                    If rs.RecordCount > 0 Then

                        docmun = rs.Fields.Item("sigiente").Value

                    End If

                    Try
                        pform.Items.Add("pfolio", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                    Catch ex As Exception

                    End Try
                    Try
                        pform.Items.Add("ptexto", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                    Catch ex As Exception

                    End Try


                Else

                    docmun = pform.DataSources.DBDataSources.Item(0).GetValue("DocNum", 0).Trim

                    If pform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And (ItemInvocador <> "FORMLOAD" And ItemInvocador <> "DUPLICAR") Then

                        Try
                            pform.Items.Add("pfolio", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                        Catch ex As Exception

                        End Try

                        Try
                            pform.Items.Add("ptexto", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                        Catch ex As Exception

                        End Try

                    End If

                End If




                If ItemInvocador = "FORMLOAD" Then

                    Try
                        pform.Items.Add("pfolio", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                    Catch ex As Exception

                    End Try

                    Try
                        pform.Items.Add("ptexto", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                    Catch ex As Exception

                    End Try


                    ' If est <> "" AndAlso pm <> "" AndAlso tipoComp <> "" Then Exit Sub

                End If


                ' dibujado del Folio Previo




                '------------------------



                Dim query As String = ""

                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                    query = "Select IFNULL(""U_Establec"",'') as EST,IFNULL(""U_PuntoEmi"",'') as PDE,IFNULL(""U_TipoD"",'') as TDOC FROM ""@SS_SERD"" WHERE ""U_SerId"" = '" + serie + "'"

                Else

                    query = "Select ISNULL(""U_Establec"",'') as EST,ISNULL(""U_PuntoEmi"",'') as PDE,ISNULL(""U_TipoD"",'') as TDOC FROM ""@SS_SERD"" WHERE ""U_SerId"" = '" + serie + "'"

                End If
                Utilitario.Util_Log.Escribir_Log("query serie electronica: " & query.ToString, "EventosLE")

                Dim leye As SAPbouiCOM.StaticText = pform.Items.Item("pfolio").Specific
                Dim leye2 As SAPbouiCOM.StaticText = pform.Items.Item("ptexto").Specific

                rs1 = oFuncionesB1.getRecordSet(query)

                Dim canReg = rs1.RecordCount
                If rs1.RecordCount > 0 Then

                    While (rs1.EoF = False)

                        'End While

                        'usform = rSboApp.Forms.Item(pform.UDFFormUID)

                        tipoDoc2 = rs1.Fields.Item("TDOC").Value

                        'usform.Items.Item("U_SS_SerieRet").Specific.String = ""
                        Utilitario.Util_Log.Escribir_Log("query serie electronica: " & query.ToString & " resultado: " & tipoDoc2.ToString, "EventosLE")
                        Select Case tipoDoc2
                            Case "FV"



                                'usform.Items.Item("U_SS_Est").Specific.String = rs1.Fields.Item("EST").Value
                                'usform.Items.Item("U_SS_Pemi").Specific.String = rs1.Fields.Item("PDE").Value

                                'usform.Items.Item("U_SS_TipCom").Specific.String = "18"

                                'usform.Items.Item("U_SS_FormaPagos").Specific.Select("20")

                                'ajuste para que sean leidos de los form  principales
                                'kingArtur

                                pform.Items.Item("etssloc1").Specific.String = rs1.Fields.Item("EST").Value
                                pform.Items.Item("etssloc2").Specific.String = rs1.Fields.Item("PDE").Value

                                pform.Items.Item("etssloc3").Specific.String = "18"

                                pform.Items.Item("etssloc6").Specific.Select("20")



                            Case "ND"
                                'usform.Items.Item("U_SS_Est").Specific.String = rs1.Fields.Item("EST").Value
                                'usform.Items.Item("U_SS_Pemi").Specific.String = rs1.Fields.Item("PDE").Value
                                'usform.Items.Item("U_SS_TipCom").Specific.String = "05"
                                'usform.Items.Item("U_SS_FormaPagos").Specific.Select("20")
                                'usform.Items.Item("U_SS_TipDocAplica").Specific.String = "18"

                                'nueva forma por tabs
                                pform.Items.Item("etssloc1").Specific.String = rs1.Fields.Item("EST").Value
                                pform.Items.Item("etssloc2").Specific.String = rs1.Fields.Item("PDE").Value
                                pform.Items.Item("etssloc3").Specific.String = "05"
                                pform.Items.Item("etssloc6").Specific.Select("20")
                                pform.Items.Item("etssloc55").Specific.String = "18"


                            Case "NC"
                                'usform.Items.Item("U_SS_Est").Specific.String = rs1.Fields.Item("EST").Value
                                'usform.Items.Item("U_SS_Pemi").Specific.String = rs1.Fields.Item("PDE").Value
                                'usform.Items.Item("U_SS_TipCom").Specific.String = "04"
                                'usform.Items.Item("U_SS_FormaPagos").Specific.Select("20")
                                'usform.Items.Item("U_SS_TipDocAplica").Specific.String = "18"

                                'nueva forma por tabs
                                pform.Items.Item("etssloc1").Specific.String = rs1.Fields.Item("EST").Value
                                pform.Items.Item("etssloc2").Specific.String = rs1.Fields.Item("PDE").Value
                                pform.Items.Item("etssloc3").Specific.String = "04"

                                'se comenta la el pago en las notas credito en el diccionario no se encuentra
                                'actualizado 20042024
                                'pform.Items.Item("etssloc6").Specific.Select("20")
                                pform.Items.Item("etssloc55").Specific.String = "18"

                            Case "GR", "GRT", "GRST"
                                'usform.Items.Item("U_SS_Est").Specific.String = rs1.Fields.Item("EST").Value
                                'usform.Items.Item("U_SS_Pemi").Specific.String = rs1.Fields.Item("PDE").Value
                                'usform.Items.Item("U_SS_TipCom").Specific.String = "06"

                                'nueva forma tabs

                                pform.Items.Item("etssloc1").Specific.String = rs1.Fields.Item("EST").Value
                                pform.Items.Item("etssloc2").Specific.String = rs1.Fields.Item("PDE").Value
                                pform.Items.Item("etssloc3").Specific.String = "06"

                            Case "LQ"
                                pform.Items.Item("etssloc1").Specific.String = rs1.Fields.Item("EST").Value
                                pform.Items.Item("etssloc2").Specific.String = rs1.Fields.Item("PDE").Value

                                pform.Items.Item("etssloc3").Specific.String = "03"
                                pform.Items.Item("etssloc6").Specific.Select("20")

                            Case "LQRT"
                                pform.Items.Item("etssloc1").Specific.String = rs1.Fields.Item("EST").Value
                                pform.Items.Item("etssloc2").Specific.String = rs1.Fields.Item("PDE").Value

                                pform.Items.Item("etssloc3").Specific.String = "03"
                                pform.Items.Item("etssloc59").Specific.String = rs1.Fields.Item("EST").Value & rs1.Fields.Item("PDE").Value
                                pform.Items.Item("etssloc6").Specific.Select("20")

                            Case "RT"

                                pform.Items.Item("etssloc59").Specific.String = rs1.Fields.Item("EST").Value & rs1.Fields.Item("PDE").Value
                                pform.Items.Item("etssloc3").Specific.String = "01"
                                pform.Items.Item("etssloc6").Specific.Select("20")
                        End Select



                        '-------------

                        If pform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then



                            Try

                                If tipoDoc2 = "GRT" Or tipoDoc2 = "GRST" Then
                                    leye.Item.Left = pform.Items.Item("10000056").Left
                                    leye.Item.Top = pform.Items.Item("10000056").Top + 20
                                    leye.Item.Width = pform.Items.Item("10000056").Width


                                    leye2.Item.Left = pform.Items.Item("10000053").Left
                                    leye2.Item.Top = pform.Items.Item("10000053").Top + 20
                                    leye2.Item.Width = pform.Items.Item("10000053").Width

                                Else
                                    leye.Item.Left = pform.Items.Item("211").Left
                                    leye.Item.Top = pform.Items.Item("211").Top + 20
                                    leye.Item.Width = pform.Items.Item("211").Width

                                    leye2.Item.Left = pform.Items.Item("84").Left
                                    leye2.Item.Top = pform.Items.Item("84").Top + 20
                                    leye2.Item.Width = pform.Items.Item("84").Width
                                End If



                                query = Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._SS_CalculoFolioQRY, sKey)
                                query = query.Replace("@SERIE", "'" & serie & "'")
                                query = query.Replace("@DOCNUM", docmun)
                                query = query.Replace("@TIPDOC", tipoDoc2)

                                Utilitario.Util_Log.Escribir_Log("query serie electronica secuencial : " & query.ToString, "EventosLE")

                                rs = oFuncionesB1.getRecordSet(query)

                                If rs.RecordCount > 0 Then

                                    leye2.Caption = "Secuencia Generada:"
                                    leye.Item.Visible = False

                                    If tipoDoc2 = "LQ" Then
                                        leye.Caption = ""
                                        pform.Items.Item("etssloc60").Specific.String = ""
                                        pform.Items.Item("etssloc59").Specific.String = ""
                                        pform.Items.Item("etssloc61").Specific.String = ""


                                        leye.Caption = rs.Fields.Item("Secuencial").Value

                                    ElseIf tipoDoc2 = "RT" Then

                                        pform.Items.Item("etssloc60").Specific.String = rs.Fields.Item("Secuencial").Value

                                        If canReg <= 1 Then
                                            leye.Caption = ""
                                        Else
                                            pform.Items.Item("etssloc3").Specific.String = "03"
                                        End If
                                    ElseIf tipoDoc2 = "LQRT" Then

                                        leye.Caption = ""
                                        pform.Items.Item("etssloc60").Specific.String = rs.Fields.Item("Secuencial").Value

                                        leye.Caption = rs.Fields.Item("Secuencial").Value

                                    Else
                                        leye.Caption = ""
                                        leye.Caption = rs.Fields.Item("Secuencial").Value

                                    End If

                                    'If tipoDoc2 = "RT" Then

                                    '    usform.Items.Item("U_SS_SecRet").Specific.String = rs.Fields.Item("Secuencial").Value
                                    '    leye.Item.Click()
                                    '    If canReg <= 1 Then
                                    '        leye.Caption = ""

                                    '    End If

                                    'Else


                                    '    If tipoDoc2 = "LQRT" Then

                                    '        usform.Items.Item("U_SS_SecRet").Specific.String = rs.Fields.Item("Secuencial").Value


                                    '    ElseIf tipoDoc2 = "LQ" Then

                                    '        usform.Items.Item("U_SS_SecRet").Specific.String = ""
                                    '        usform.Items.Item("U_SS_SerieRet").Specific.String = ""
                                    '        usform.Items.Item("U_SS_NumAutRet").Specific.String = ""
                                    '        usform.Items.Item("U_SS_FecRet").Specific.String = ""

                                    '    End If


                                    '    leye.Caption = rs.Fields.Item("Secuencial").Value


                                    'End If


                                    leye2.Item.Visible = True

                                    leye.Item.Visible = True
                                    oFuncionesB1.Release(rs)

                                End If

                            Catch ex As Exception

                                ' leye.Caption = "Error al Obtener el Folio!"

                                rSboApp.SetStatusBarMessage(ex.Message)
                            End Try

                        End If ' modo add


                        rs1.MoveNext()

                    End While
                Else

                    If IdFormDM = "-141" Or IdFormDM = "141" Then

                        pform.Items.Item("etssloc59").Specific.String = ""
                        pform.Items.Item("etssloc60").Specific.String = ""
                        'pform.Items.Item("etssloc3").Specific.String = ""


                    Else

                        pform.Items.Item("etssloc1").Specific.String = ""
                        pform.Items.Item("etssloc2").Specific.String = ""

                        pform.Items.Item("etssloc3").Specific.String = ""

                        Try
                            pform.Items.Item("etssloc6").Specific.Select("")

                        Catch ex As Exception

                        End Try

                        'usform.Items.Item("U_SS_SerieRet").Specific.String = ""
                        'usform.Items.Item("U_SS_TipCom").Specific.String = ""

                    End If




                    leye2.Item.Visible = False

                    leye.Item.Visible = False

                End If


                Try
                    pform.Refresh()
                    leye.Item.Refresh()

                    'pform.Refresh()

                    oFuncionesB1.Release(rs1)

                Catch ex As Exception

                    Utilitario.Util_Log.Escribir_Log("Error al Liberar o Refrescar from " & ex.Message, "EventosLE")

                End Try






                'End If


            Catch ex As Exception

                rSboApp.SetStatusBarMessage(ex.Message)

            End Try
        End If

    End Sub



    Public Function AnularDocumento(ByVal AidForm As String) As Boolean
        'Dim _AidForm As String = AidForm
        'If AidForm = "133" Or AidForm = "60091" Or AidForm = "60090" Then

        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)


        'ElseIf AidForm = "65303" Then
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oDocumento.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
        '    oDocumento.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_DebitMemo
        'End If
        'oDocumento.Browser.GetByKeys(BusinessObjectInfo.ObjectKey)

        'If BusinessObjectInfo.FormTypeEx = "133" Or BusinessObjectInfo.FormTypeEx = "60090" Then  ' FACTURA DE CLIENTE - FACTURA DEUDOR + PAGO
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oTipoTabla = "FCE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "60091" Then ' FACTURA DE RESERVA
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oTipoTabla = "FRE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "65303" Then ' NOTA DE DEBITO
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '    oTipoTabla = "NDE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "65300" Then ''FACTURA DE ANTICIPO DE CLIENTES
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
        '    oTipoTabla = "FAE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "179" Then 'NOTA DE CREDITO DE CLIENTES
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
        '    oTipoTabla = "NCE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "140" Then 'GUIA DE REMISION - ENTREGA
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
        '    oTipoTabla = "GRE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "940" Then 'GUIA DE REMISION - TRANSFERENCIAS                                            
        '    oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
        '    oTipoTabla = "TRE"

        'ElseIf BusinessObjectInfo.FormTypeEx = "1250000940" Then 'GUIA DE REMISION - SOLICITUD TRANSLADO                                            
        '    oTransferencia = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
        '    oTipoTabla = "TLE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "141" Then  'FACTURA DE PROVEEDOR/RETENCION                             
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "REE"
        'ElseIf BusinessObjectInfo.FormTypeEx = "65306" Then  'NOTA DE DEBITO PROVEEDOR/RETENCION                             
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "RDM"
        '    Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
        'ElseIf BusinessObjectInfo.FormTypeEx = "65301" Then  'FACTURA DE ANTICIPO DE PROVEEDOR/RETENCION                             
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "REA"
        '    Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")
        'ElseIf BusinessObjectInfo.FormTypeEx = "60092" Then  'FACTURA DE RESERVA PROVEEDOR/RETENCION                           
        '    oDocumento = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        '    oTipoTabla = "RER"
        '    Utilitario.Util_Log.Escribir_Log("Tipo Tabla: " + oTipoTabla.ToString, "ManejoDeDocumentos")



        Return True
    End Function

    Public Function ValidarDocSerieElectronica(TipoDoc As String, ByVal IdSerie As String, Optional ByVal typeEx As String = "", Optional oForm As SAPbouiCOM.Form = Nothing) As Boolean
        'Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia_LiquidacionDeCompra (antes del try)", "ManejoDeDocumentos")
        Try
            'Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS
            Dim serie As String = ""
            Dim QrySerie As String = ""
            Dim QrySerieLQ As String = ""

            Dim codDoc As String = "00"
            If (SeriesElectronicasUDF = "Y" And typeEx = "141") Or (SeriesElectronicasUDF = "Y" And typeEx = "-141") Or (SeriesElectronicasUDF = "Y" And typeEx = "60092") Or (SeriesElectronicasUDF = "Y" And typeEx = "-60092") Then
                codDoc = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_DocEmision", 0).Trim

                If codDoc = "01" And TipoDoc <> "LQE" Then

                    EsElectronico = "FE"
                    Return True

                ElseIf codDoc = "02" And TipoDoc = "LQE" Then

                    EsElectronico = ""
                    LQEsElectronico = "FE"
                    Return True

                ElseIf codDoc = "03" Then

                    EsElectronico = "FE"
                    LQEsElectronico = "FE"
                    Return True

                Else

                    EsElectronico = ""
                    LQEsElectronico = ""
                    Return False
                End If


            Else

                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then

                    If TipoDoc = "REE" Then

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QrySerie = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE B.""U_FE_TipoEmision""='FE' AND B.""U_TIPO_DOC""='RT' and A.""Series"" = " + IdSerie
                        Else
                            QrySerie = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE B.""U_FE_TipoEmision""='FE' AND B.""U_TIPO_DOC""='RT' and Series = " + IdSerie
                        End If

                    ElseIf TipoDoc = "LQE" Then

                        QrySerieLQ = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + IdSerie

                    Else
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QrySerie = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE B.""U_FE_TipoEmision""='FE' AND A.""Series"" = " + IdSerie
                        Else
                            QrySerie = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE B.U_FE_TipoEmision='FE' AND Series = " + IdSerie
                        End If
                    End If


                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then

                    If TipoDoc = "LQE" Then
                        QrySerieLQ = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + IdSerie
                    Else
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QrySerie = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + IdSerie
                        Else
                            QrySerie = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + IdSerie
                        End If
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then

                    If TipoDoc = "LQE" Then
                        QrySerieLQ = "SELECT ""Code"" FROM ""@GS_LIQUI"" where ""U_IdSerie"" = " + IdSerie
                    Else
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            QrySerie = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + IdSerie
                        Else
                            QrySerie = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + IdSerie
                        End If
                    End If

                    Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + QrySerie, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                    Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES LQ ELECTRÓNICO: " + QrySerieLQ, "FUNCION_VALIDAR_SERIE_ELECTRONICA")


                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                        If TipoDoc = "REE" Or TipoDoc = "RER" Or TipoDoc = "RDM" Then

                            QrySerie = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE ""U_TipoD"" IN ('RT','LQRT') AND A.""Series"" = " + IdSerie

                        ElseIf TipoDoc = "LQE" Then

                            QrySerie = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE B.""U_TipoD"" IN ('LQ','LQRT') AND A.""Series"" = " + IdSerie

                        Else

                            QrySerie = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + IdSerie

                        End If

                    Else

                        If TipoDoc = "REE" Or TipoDoc = "RER" Or TipoDoc = "RDM" Then
                            QrySerie = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE ""U_TipoD"" IN ('RT','LQRT') AND Series = " + IdSerie
                        ElseIf TipoDoc = "LQE" Then
                            QrySerie = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE B.U_TipoD IN ('LQ','LQRT') and Series = " + IdSerie
                        Else
                            QrySerie = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + IdSerie
                        End If

                    End If
                End If

                Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + QrySerie, "FUNCION_VALIDAR_SERIE_ELECTRONICA")

                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                    If TipoDoc <> "LQE" Then
                        EsElectronico = oFuncionesB1.getRSvalue(QrySerie, "U_FE_TipoEmision", "")
                        Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                        If EsElectronico = "FE" Then
                            Return True
                        Else
                            Return False
                        End If
                    Else

                        LQEsElectronico = oFuncionesB1.getRSvalue(QrySerieLQ, "Code", "")
                        Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY LQ: " + LQEsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                        If LQEsElectronico <> "" Then
                            LQEsElectronico = "FE"
                            Return True
                        Else
                            Return False
                        End If

                    End If



                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    EsElectronico = oFuncionesB1.getRSvalue(QrySerie, "U_DIGITAL", "")

                    If TipoDoc = "LQE" Then
                        LQEsElectronico = oFuncionesB1.getRSvalue(QrySerieLQ, "Code", "")
                        If LQEsElectronico <> "" Then
                            LQEsElectronico = "FE"
                            Return True
                        Else
                            LQEsElectronico = ""
                            Return False
                        End If
                    Else
                        If EsElectronico = "Y" Then
                            EsElectronico = "FE"
                            Return True
                        Else
                            EsElectronico = ""
                            Return False
                        End If
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                        Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Or
                         Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                    EsElectronico = oFuncionesB1.getRSvalue(QrySerie, "U_ELECTRONICA", "")

                    If TipoDoc = "LQE" Then
                        LQEsElectronico = oFuncionesB1.getRSvalue(QrySerieLQ, "Code", "")
                        Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                        If LQEsElectronico <> "" Then
                            LQEsElectronico = "FE"
                            Return True
                        Else
                            LQEsElectronico = ""
                            Return False
                        End If
                    Else
                        Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                        If EsElectronico = "SI" Then
                            EsElectronico = "FE"
                            Return True
                        Else
                            EsElectronico = ""
                            Return False
                        End If
                    End If

                ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    EsElectronico = oFuncionesB1.getRSvalue(QrySerie, "U_FE_TipoEmision", "")
                    If TipoDoc = "REE" Or TipoDoc = "RER" Or TipoDoc = "RDM" Or TipoDoc = "LQE" Then
                        If EsElectronico = "RT" Then
                            EsElectronico = "FE"
                            If TipoDoc = "LQE" Then
                                LQEsElectronico = ""
                                Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                                Return False
                            End If
                            Return True
                        ElseIf EsElectronico = "LQRT" Then
                            EsElectronico = "FE"
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            LQEsElectronico = "FE"
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY LQ: " + LQEsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            Return True
                        ElseIf EsElectronico = "LQ" Then
                            EsElectronico = ""
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY LQ: " + LQEsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            LQEsElectronico = "FE"
                            Return True
                        ElseIf EsElectronico = "NAN" Or EsElectronico = "" Then
                            EsElectronico = "NA"
                            LQEsElectronico = "NA"
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            Return False
                        Else
                            EsElectronico = "NA"
                            LQEsElectronico = "NA"
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            Return False
                        End If

                    Else
                        If EsElectronico = "NAN" Or EsElectronico = "" Then
                            EsElectronico = "NA"
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            Return False
                        Else
                            EsElectronico = "FE"
                            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "FUNCION_VALIDAR_SERIE_ELECTRONICA")
                            Return True
                        End If
                    End If
                End If

            End If




        Catch ex As Exception
            rCompany.SetStatusBarMessage(NombreAddon + "Error al validar serie electrónica", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
            Return False
        End Try

    End Function

#End Region

    Shared Function customCertValidation(ByVal sender As Object,
                                             ByVal cert As X509Certificate,
                                             ByVal chain As X509Chain,
                                             ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function

    Private Sub AgregaBotonLP(ByVal oForm As SAPbouiCOM.Form)

        Try
            Dim oItem As SAPbouiCOM.Item
            Dim oItem1 As SAPbouiCOM.Item
            Dim oButton As SAPbouiCOM.ButtonCombo

            Try
                oItem = oForm.Items.Add("btnAccionP", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
            Catch ex As Exception
                oItem = oForm.Items.Item("btnAccionP")
            End Try

            oItem1 = oForm.Items.Item("52")
            oItem.Enabled = True
            oItem.Left = oItem1.Left + oItem1.Width + 5
            oItem.Width = oItem1.Width + 15
            oItem.Top = oItem1.Top
            oItem.Height = oItem1.Height
            oItem.AffectsFormMode = False
            oButton = oItem.Specific
            'oButton.Caption = "(SS) Renviar SRI"
            oButton.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oButton.ValidValues.Add("(SS) Renviar SRI", "(SS) Renviar SRI")
            oButton.Select("(SS) Renviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oButton.Item.Visible = False

        Catch ex As Exception
            rSboApp.MessageBox("ERROR: " & ex.Message)
        End Try

    End Sub

    Private Sub CargaItemEnFormularioLP(oForm As SAPbouiCOM.Form, evento As String, typeEx As String, oTabla As String)
        Dim codDoc As String = "00"

        Try
            'Dim oFormMode As Integer = 1
            'oFormMode = oForm.Mode

            Dim oSerie As String = ""
            ' OBTENGO EL NUMERO DE SERIE
            If oTipoTabla = "TRE" Then
                oSerie = oForm.Items.Item("40").Specific.value.ToString()
            ElseIf oTipoTabla = "TLE" Then
                oSerie = oForm.Items.Item("1250000068").Specific.value.ToString()
            Else
                oSerie = oForm.Items.Item("88").Specific.value.ToString()
            End If


            Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "EventosEmision")
            Dim SQL As String = ""
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE B.""U_TIPO_DOC""='RT' and A.""Series"" = " + oSerie
                    Else
                        SQL = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE B.""U_TIPO_DOC""='RT' and Series = " + oSerie
                    End If
                Else
                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        SQL = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE A.""Series"" = " + oSerie
                    Else
                        SQL = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE Series = " + oSerie
                    End If
                End If


            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SQL = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + oSerie
                Else
                    SQL = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + oSerie
                End If
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SQL = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + oSerie
                Else
                    SQL = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + oSerie
                End If
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    SQL = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + oSerie
                Else
                    SQL = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + oSerie
                End If

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then

                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                        SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE ""U_TipoD"" IN ('RT','LQRT') AND A.""Series"" = " + oSerie
                    Else
                        SQL = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + oSerie
                    End If

                Else
                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                        SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE ""U_TipoD"" IN ('RT','LQRT') AND Series = " + oSerie
                    Else
                        SQL = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + oSerie
                    End If

                End If
            End If

            Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + SQL, "EventosEmision")
            If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
                EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_DIGITAL", "")

                If EsElectronico = "Y" Then
                    EsElectronico = "FE"
                Else
                    EsElectronico = "NA"
                End If

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
                EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
                If EsElectronico = "SI" Then
                    EsElectronico = "FE"
                Else
                    EsElectronico = "NA"
                End If

            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
                Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
                If EsElectronico = "SI" Then
                    EsElectronico = "FE"
                Else
                    EsElectronico = "NA"
                End If
            ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")
                If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
                    If EsElectronico = "RT" Or EsElectronico = "LQRT" Then
                        EsElectronico = "FE"
                    ElseIf EsElectronico = "NAN" Or EsElectronico = "" Then
                        EsElectronico = "NA"
                    Else
                        EsElectronico = "NA"
                    End If
                Else
                    If EsElectronico = "NAN" Or EsElectronico = "" Then
                        EsElectronico = "NA"
                    Else
                        EsElectronico = "FE"
                    End If
                End If
            End If
            Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "EventosEmision")
            'End If

            'FRIZEO ANTES DE CUALQUIER DIBUJADO

            oForm.Freeze(True)

            If evento = "NUEVO" Then

                lbEstado = oForm.Items.Item("lbEstado").Specific
                lbEstado.Item.ForeColor = RGB(204, 0, 0)
                lbEstado.Caption = "Estado: NO ENVIADO"
                lbEstado.Item.Visible = True

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                lbComentario = oForm.Items.Item("lbComen").Specific

                If EsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(7, 118, 10)
                    ' lbComentario.Caption = "ELECTRONICO"
                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "(GS) RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "(GS) DOCUMENTO NO ELECTRÓNICO"
                    End If

                    lbComentario.Item.ForeColor = RGB(204, 0, 0)
                End If
                lbComentario.Item.Visible = True
                If (oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM") And EsElectronico <> "FE" Then
                    lbComentario.Item.Visible = False
                    lbEstado.Item.Visible = False
                End If
                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("NUEVO-Eliminacion valores validos", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

            ElseIf evento = "ACTUALIZAR" Then

                lbComentario = oForm.Items.Item("lbComen").Specific
                lbEstado = oForm.Items.Item("lbEstado").Specific
                'cbEstado = oForm.Items.Item("cbEstado").Specific
                'btnAccion = oForm.Items.Item("btnAccion").Specific

                'Obtengo valor almacenado en base
                Dim UDFEA As String = oForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_ESTADO_AUTORIZACIO", 0)

                btnAccion = oForm.Items.Item("btnAccion").Specific
                btnAccion.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                btnAccion.Item.AffectsFormMode = False
                btnAccion.Item.Visible = False

                ' ELIMINO LOS VALORES VALIDOS PARA LUEGO AGREGARLO DEPENDIENDO
                Try
                    If btnAccion.ValidValues.Count > 0 Then
                        For i As Integer = btnAccion.ValidValues.Count - 1 To 0 Step -1
                            btnAccion.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                        Next
                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage(NombreAddon + " - Actualizar-Eliminacion datos combo", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If String.IsNullOrWhiteSpace(UDFEA) Then
                    'Entra aqui solo si el udf no tiene un valor por defecto
                    'si tiene seteado a 0 por default no deberia hacer esto
                    UDFEA = "0"
                End If
                UDFEA = UDFEA.Trim

                Try
                    If EsElectronico = "FE" Then
                        If UDFEA = "2" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: AUTORIZADO"

                        ElseIf UDFEA = "5" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: RECIBIDO"

                        ElseIf UDFEA = "7" Then
                            btnAccion.ValidValues.Add("(GS) Ver RIDE", "(GS) Ver Ride")
                            btnAccion.ValidValues.Add("(GS) Consultar AUT", "(GS) Cons. Autorizacion")
                            btnAccion.ValidValues.Add("(GS) Ver XML", "(GS) Ver XML")
                            'btnAccion.ValidValues.Add("(GS) Reenviar MAIL", "(GS) Reenviar Mail")
                            btnAccion.Select("(GS) Ver RIDE", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(7, 118, 10)
                            lbEstado.Caption = "Estado: ERROR RECEPCION SRI"

                        ElseIf UDFEA = "4" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: VALIDAR DATOS"

                        ElseIf UDFEA = "3" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO AUTORIZADA"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEA = "6" Then
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: DEVUELTA"

                        ElseIf UDFEA = "11" Then
                            'btnAccLQ.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            'btnAccLQ.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = False
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: ANULADO"

                        Else
                            btnAccion.ValidValues.Add("(GS) Reenviar SRI", "(GS) Reenviar SRI")
                            btnAccion.Select("(GS) Reenviar SRI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            btnAccion.Item.Visible = True
                            lbEstado.Item.ForeColor = RGB(204, 0, 0)
                            lbEstado.Caption = "Estado: NO ENVIADO"

                        End If

                    End If
                Catch ex As Exception
                    rSboApp.SetStatusBarMessage("Actualizar-FE ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

                If EsElectronico = "FE" Then
                    oForm.Title += " - ELECTRONICO"
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(7, 118, 10)

                    Select Case oForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE, SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            lbComentario.Item.Visible = True
                            lbEstado.Item.Visible = True
                            If Functions.VariablesGlobales._vgBloquearReenviarSRI = "Y" Then
                                If UDFEA = "0" Then
                                    btnAccion.Item.Visible = False
                                End If
                            End If

                        Case Else
                            btnAccion.Item.Visible = False

                    End Select

                Else
                    oForm.Title = nombreFormulario
                    If oTipoTabla = "REE" Then
                        lbComentario.Caption = "GS RETENCIÓN NO ELECTRÓNICA"
                    Else
                        lbComentario.Caption = "GS DOCUMENTO NO ELECTRÓNICO"
                    End If
                    lbComentario.Item.ForeColor = RGB(204, 0, 0)

                    lbEstado.Caption = "Estado: NO ENVIADO"
                    lbEstado.Item.ForeColor = RGB(204, 0, 0)
                    btnAccion.Item.Visible = False

                    If oTipoTabla = "REE" Then
                        lbComentario.Item.Visible = False
                        lbEstado.Item.Visible = False
                    End If


                End If

            End If

        Catch ex As Exception
            ' rSboApp.SetStatusBarMessage("General ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
        'oFrmUser.Freeze(False)
        'oForm.Update()
        'oForm.Refresh()
    End Sub

    Public Function ValidarGuiaEnFacturaHeison(oForm As SAPbouiCOM.Form, idForm As String) As Boolean

        Dim DocEntryFG As String = "00"
        Dim QryBotonFacGuia = ""
        Dim _QryBotonFacGuia = ""

        If idForm = "133" Or idForm = "60091" Then

            Try
                DocEntryFG = oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).Trim

                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    QryBotonFacGuia = "select distinct ifnull(T0.""U_HBT_DocEntry"",0) as ""HBT_DocEntry"" FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.""Code""=T0.""U_HBT_IdGuiaRemision"" inner join OINV ON T1.""U_HBT_NumeroDesde1""=OINV.""DocNum"" where OINV.""DocEntry"" =" + DocEntryFG.ToString

                Else
                    QryBotonFacGuia = "select distinct ISNULL(T0.U_HBT_DocEntry,0)as HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OINV ON T1.U_HBT_NumeroDesde1=OINV.""DocNum"" where OINV.""DocEntry"" =" + DocEntryFG.ToString

                End If
                _QryBotonFacGuia = oFuncionesB1.getRSvalue(QryBotonFacGuia, "HBT_DocEntry", "")

                If _QryBotonFacGuia = "" Or _QryBotonFacGuia = "0" Then
                    Functions.VariablesGlobales._FacturaGuiaRemision = ""

                    Return False
                Else
                    Functions.VariablesGlobales._FacturaGuiaRemision = "SI"
                    Return True
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query boton guia factura: " + QryBotonFacGuia + "ERROR: " + ex.Message.ToString, "EventosEmision")
                Return False
            End Try

        Else

            Try
                DocEntryFG = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0).Trim

                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    QryBotonFacGuia = "select distinct ifnull(T0.""U_HBT_DocEntry"",0) as ""HBT_DocEntry"" FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.""Code""=T0.""U_HBT_IdGuiaRemision"" inner join OIGE ON T1.""U_HBT_NumeroDesde4""=OIGE.""DocNum"" where OIGE.""DocEntry"" =" + DocEntryFG.ToString

                Else
                    QryBotonFacGuia = "select distinct ISNULL(T0.U_HBT_DocEntry,0)as HBT_DocEntry FROM ""@HBT_GUIAREMDETALLE"" T0 inner join ""@HBT_GUIAREMISION"" T1 ON T1.Code=T0.U_HBT_IdGuiaRemision inner join OIGE ON T1.U_HBT_NumeroDesde4=OIGE.""DocNum"" where OIGE.""DocEntry"" =" + DocEntryFG.ToString

                End If
                _QryBotonFacGuia = oFuncionesB1.getRSvalue(QryBotonFacGuia, "HBT_DocEntry", "")

                If _QryBotonFacGuia = "" Or _QryBotonFacGuia = "0" Then
                    Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = ""

                    Return False
                Else
                    Functions.VariablesGlobales._SalidaMercanciasGuiaRemision = "SI"
                    Return True
                End If
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("Query boton guia factura: " + QryBotonFacGuia + "ERROR: " + ex.Message.ToString, "EventosEmision")
                Return False
            End Try

        End If




    End Function

    Private Sub ProcesoActualizarParaCancelacion(Docenrty As String, objtype As Integer)

        Try
            Dim Sentencia As String = ""

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sentencia = "call " & rCompany.CompanyDB & ".SS_ACTUALIZARCAMPOSCANCELACION(" + Docenrty.ToString + "," + objtype.ToString + ") "
            Else
                Sentencia = "EXEC SS_ACTUALIZARCAMPOSCANCELACION " + Docenrty.ToString + "," + objtype.ToString
            End If

            Dim mrst As SAPbobsCOM.Recordset = Nothing
            mrst = rCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            mrst.DoQuery(Sentencia)

        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Error al actualizar los campos" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try

    End Sub


#Region "Generacion Controles TAB y Operaciones QR"

    Private Sub CargarQRenPantalla(oForm As SAPbouiCOM.Form, dato As String, nombreImg As String, Optional esQRLIQ As Boolean = False)

        Try

            Dim referenciaimagenQR As SAPbouiCOM.PictureBox

            If esQRLIQ Then

                referenciaimagenQR = oForm.Items.Item("pboxQRL").Specific
            Else

                referenciaimagenQR = oForm.Items.Item("pboxQR").Specific
            End If

            If dato = "" Then

                referenciaimagenQR.Picture = ""

            Else

                Dim respimg = Utilitario.ClsGeneradorQR.SaveStringToQR(dato, nombreImg, 300, 300)

                If respimg <> String.Empty Then

                    referenciaimagenQR.Picture = respimg

                Else

                    referenciaimagenQR.Picture = ""

                End If


            End If


        Catch ex As Exception

        End Try

    End Sub

    Private Function GenerarEnlaceQR(clave As String) As String



        If Functions.VariablesGlobales._wsConsultaEmision <> "" And clave <> "" Then

            Dim x() = Functions.VariablesGlobales._wsConsultaEmision.Split("/")


            If Functions.VariablesGlobales._wsConsultaEmision.ToLower.Contains("edocnube.com") Then
                Return $"{x(0)}//{x(2)}/WSEDOC_FILES/files/{clave}.pdf"
            Else
                Return $"{x(0)}//{x(2)}/eDocEcuador/WSEDOC_FILES/files/{clave}.pdf"
            End If



        End If



        Return String.Empty


    End Function

    Private Sub crearTabPrueba_Facturacion(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try

            oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)


            Const NumeroPane As Integer = 50

            'CREO EL NUEVO TAB
            Dim tabreferencia As SAPbouiCOM.Item
            Dim itemreferencia As SAPbouiCOM.Item

            Dim campoTexto As SAPbouiCOM.EditText
            Dim carpetaEdoc As SAPbouiCOM.Folder
            Dim campoCombo As SAPbouiCOM.ComboBox

            'variables para tabs

            Dim IdTabReferencia As String = ""
            Dim ItemToLinkear As String = ""


            Dim lef1 As Integer = 0
            Dim top1 As Integer = 0
            Dim lef2 As Integer = 0
            Dim top2 As Integer = 0

            Dim anchoGenerico = 0
            Dim altoGenerico = 0


            If oForm.TypeEx = "1250000940" Or oForm.TypeEx = "940" Then


                Dim valorLeft As Integer = oForm.Items.Item("5").Left
                Dim valorTop As Integer = 260
                Dim valorAncho As Integer = oForm.Items.Item("3").Width
                Dim valorAlto As Integer = oForm.Items.Item("3").Height

                lef1 = valorLeft
                top1 = valorTop

                lef2 = 130 + lef1
                top2 = valorTop

                anchoGenerico = valorAncho
                altoGenerico = valorAlto

                IdTabReferencia = "1320000081"
                ItemToLinkear = "5"

            Else


                Dim valorLeft As Integer
                Dim valorTop As Integer

                If Functions.VariablesGlobales._PosicionItemTabX <> "" AndAlso Functions.VariablesGlobales._PosicionItemTabY <> "" Then

                    Try
                        valorLeft = CInt(Functions.VariablesGlobales._PosicionItemTabX)
                        valorTop = CInt(Functions.VariablesGlobales._PosicionItemTabY)
                    Catch ex As Exception

                        valorLeft = oForm.Items.Item("19").Left
                        valorTop = oForm.Items.Item("19").Top

                    End Try

                Else

                    valorLeft = oForm.Items.Item("19").Left
                    valorTop = oForm.Items.Item("19").Top

                End If

                'Dim valorLeft As Integer = oForm.Items.Item("19").Left
                'Dim valorTop As Integer = oForm.Items.Item("19").Top



                Dim valorAncho As Integer = oForm.Items.Item("18").Width
                Dim valorAlto As Integer = oForm.Items.Item("18").Height

                'ya se encontraba
                ' mejor ajustar con nuevas variables en parte superior

                lef1 = valorLeft
                top1 = valorTop

                lef2 = 130 + lef1
                top2 = valorTop

                anchoGenerico = valorAncho
                altoGenerico = valorAlto

                IdTabReferencia = "138"
                ItemToLinkear = "5"


            End If

            tabreferencia = oForm.Items.Item(IdTabReferencia) ' tab referencia


            If oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_FOLDER, "TabGenFE", "SAED FE", tabreferencia.Left + tabreferencia.Width, tabreferencia.Top, tabreferencia.Width, tabreferencia.Height) Then

                carpetaEdoc = oForm.Items.Item("TabGenFE").Specific
                carpetaEdoc.Pane = NumeroPane
                carpetaEdoc.DataBind.SetBound(True, "", "FolderDS")

                carpetaEdoc.GroupWith(IdTabReferencia)

                ''Aqui empieza la columna 1


                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut1", "**ESCANEAR / CLICK EN EL QR**", lef1, top1, anchoGenerico + 70, altoGenerico, NumeroPane,, True)
                oForm.Items.Item("lssaut1").LinkTo = ItemToLinkear

                'Cuadro Para el QR
                top1 += 15
                top2 += 15

                'DibujaQRenPantalla(oForm)

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_PICTURE, "pboxQR", "QR", lef1, top1, 150, 150, NumeroPane)
                oForm.Items.Item("pboxQR").LinkTo = ItemToLinkear
                'Crear los uuidCertificacion
                top1 += 15 + 135
                top2 += 15 + 135

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut2", "Clave de Acceso", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("lssaut2").LinkTo = ItemToLinkear

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssut2", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("etssut2").LinkTo = ItemToLinkear

                campoTexto = oForm.Items.Item("etssut2").Specific
                campoTexto.Item.Enabled = False
                'campoTexto.DataBind.SetBound(True, "", "etssut1")
                campoTexto.DataBind.SetBound(True, oTabla, "U_CLAVE_ACCESO")


                '(GSE) Fecha Certificcion
                top1 += 15
                top2 += 15
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut3", "Fecha.Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("lssaut3").LinkTo = ItemToLinkear

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssut3", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("etssut3").LinkTo = ItemToLinkear

                campoTexto = oForm.Items.Item("etssut3").Specific
                campoTexto.Item.Enabled = False
                campoTexto.DataBind.SetBound(True, oTabla, "U_FECHA_AUT_FACT")

                '(GSE) Estado
                top1 += 15
                top2 += 15
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut4", "Estado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("lssaut4").LinkTo = ItemToLinkear

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssut4", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("etssut4").LinkTo = ItemToLinkear

                campoCombo = oForm.Items.Item("etssut4").Specific
                campoCombo.Item.DisplayDesc = True
                campoCombo.DataBind.SetBound(True, oTabla, "U_ESTADO_AUTORIZACIO")


                '(GSE) Fecha Certificcion
                top1 += 15
                top2 += 15
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut5", "Número.Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("lssaut5").LinkTo = ItemToLinkear

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssut5", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("etssut5").LinkTo = ItemToLinkear

                campoTexto = oForm.Items.Item("etssut5").Specific
                campoTexto.Item.Enabled = False
                campoTexto.DataBind.SetBound(True, oTabla, "U_NUM_AUTO_FAC")

                '(GSE) URL QR
                top1 += 15
                top2 += 15
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut6", "Url QR", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("lssaut6").LinkTo = ItemToLinkear

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT, "etssut6", "", lef2, top2, anchoGenerico, altoGenerico * 2, NumeroPane)
                oForm.Items.Item("etssut6").LinkTo = ItemToLinkear


                campoTexto = oForm.Items.Item("etssut6").Specific
                campoTexto.Item.Enabled = False
                ' campoTexto.DataBind.SetBound(True, oTabla, "U_QR")
                campoTexto.String = "" ' GenerarEnlaceQR(claveRetencion)


                '(GSE) Observación
                top1 += 5 + (altoGenerico * 2)
                top2 += 5 + (altoGenerico * 2)
                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut7", "Observación", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                oForm.Items.Item("lssaut7").LinkTo = ItemToLinkear

                oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT, "etssut7", "", lef2, top2, anchoGenerico, altoGenerico * 2, NumeroPane)
                oForm.Items.Item("etssut7").LinkTo = ItemToLinkear

                campoTexto = oForm.Items.Item("etssut7").Specific
                campoTexto.Item.Enabled = False
                campoTexto.DataBind.SetBound(True, oTabla, "U_OBSERVACION_FACT")

                '------------------------------------------------------------------------------
                'Aqui empieza la columna 2

                If oForm.TypeEx = "141" Or oForm.TypeEx = "60092" Then

                    If oForm.TypeEx = "1250000940" Or oForm.TypeEx = "940" Then


                        Dim valorLeft = oForm.Items.Item("1470000099").Left + 5
                        Dim valorTop = 260

                        lef1 = valorLeft
                        top1 = valorTop

                        lef2 = lef1 + 100
                        top2 = valorTop

                        ItemToLinkear = "1470000099"

                    Else

                        itemreferencia = oForm.Items.Item("156")

                        lef1 = itemreferencia.Left
                        top1 = itemreferencia.Top

                        lef2 = itemreferencia.Left + 100
                        top2 = itemreferencia.Top

                        ItemToLinkear = "11"

                    End If


                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut20", "**ESCANEAR / CLICK EN EL QR** (LIQ)", lef1, top1, anchoGenerico + 75, altoGenerico, NumeroPane,, True)
                    'oForm.Items.Item("lssaut20").LinkTo = "TabGenFE"
                    oForm.Items.Item("lssaut20").LinkTo = ItemToLinkear

                    'Cuadro Para el QR
                    top1 += 15
                    top2 += 15

                    ''DibujaQRenPantalla(oForm)

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_PICTURE, "pboxQRL", "QR", lef1, top1, 150, 150, NumeroPane)
                    oForm.Items.Item("pboxQRL").LinkTo = ItemToLinkear
                    'Crear los uuidCertificacion
                    top1 += 15 + 135
                    top2 += 15 + 135

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut21", "Clave de Acceso", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("lssaut21").LinkTo = ItemToLinkear

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssut21", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("etssut21").LinkTo = ItemToLinkear

                    campoTexto = oForm.Items.Item("etssut21").Specific
                    campoTexto.Item.Enabled = False
                    'campoTexto.DataBind.SetBound(True, "", "etssut1")
                    campoTexto.DataBind.SetBound(True, oTabla, "U_LQ_CLAVE")


                    '(GSE) Fecha Certificcion
                    top1 += 15
                    top2 += 15
                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut22", "Fecha.Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("lssaut22").LinkTo = ItemToLinkear

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssut22", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("etssut22").LinkTo = ItemToLinkear

                    campoTexto = oForm.Items.Item("etssut22").Specific
                    campoTexto.Item.Enabled = False
                    campoTexto.DataBind.SetBound(True, oTabla, "U_LQ_FECHA_AUT")

                    '(GSE) Estado
                    top1 += 15
                    top2 += 15
                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut23", "Estado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("lssaut23").LinkTo = ItemToLinkear

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssut23", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("etssut23").LinkTo = ItemToLinkear

                    campoCombo = oForm.Items.Item("etssut23").Specific
                    campoCombo.Item.DisplayDesc = True
                    campoCombo.DataBind.SetBound(True, oTabla, "U_LQ_ESTADO")


                    '(GSE) Fecha Certificcion
                    top1 += 15
                    top2 += 15
                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut24", "Número.Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("lssaut24").LinkTo = ItemToLinkear

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssut24", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("etssut24").LinkTo = ItemToLinkear

                    campoTexto = oForm.Items.Item("etssut24").Specific
                    campoTexto.Item.Enabled = False
                    campoTexto.DataBind.SetBound(True, oTabla, "U_LQ_NUM_AUTO")

                    '(GSE) URL QR
                    top1 += 15
                    top2 += 15
                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut25", "Url QR", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("lssaut25").LinkTo = ItemToLinkear

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT, "etssut25", "", lef2, top2, anchoGenerico, altoGenerico * 2, NumeroPane)
                    oForm.Items.Item("etssut25").LinkTo = ItemToLinkear

                    campoTexto = oForm.Items.Item("etssut25").Specific
                    campoTexto.Item.Enabled = False
                    ' campoTexto.DataBind.SetBound(True, oTabla, "U_QR")
                    campoTexto.String = "" ' GenerarEnlaceQR(claveRetencion)


                    '(GSE) Observación
                    top1 += 5 + (altoGenerico * 2)
                    top2 += 5 + (altoGenerico * 2)
                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssaut26", "Observación", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                    oForm.Items.Item("lssaut26").LinkTo = ItemToLinkear

                    oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT, "etssut26", "", lef2, top2, anchoGenerico, altoGenerico * 2, NumeroPane)
                    oForm.Items.Item("etssut26").LinkTo = ItemToLinkear

                    campoTexto = oForm.Items.Item("etssut26").Specific
                    campoTexto.Item.Enabled = False
                    campoTexto.DataBind.SetBound(True, oTabla, "U_LQ_OBSERVACION")


                End If ' solo si son formulario de Factura Prov se muestra el de liq





            End If

            ' oForm.PaneLevel = 1

        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Tab de pruebas ", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
    End Sub


    Private Function LlenadoCombosporUDT(ByRef combo As SAPbouiCOM.ComboBox, nombreTabla As String) As Boolean

        Dim rst As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

        Try

            rst.DoQuery($"Select ""Code"",""Name"" from ""{nombreTabla}"" ")

            If rst.RecordCount > 0 Then

                While Not rst.EoF

                    combo.ValidValues.Add(rst.Fields.Item("Code").Value.ToString, rst.Fields.Item("Name").Value.ToString)

                    rst.MoveNext()
                End While


            End If

            oFuncionesB1.Release(rst)

            Return True
        Catch ex As Exception

        End Try

        oFuncionesB1.Release(rst)

        Return False

    End Function

    Private Sub crearTabPrueba_Localizacion2(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try
            oForm.DataSources.UserDataSources.Add("FolderDS2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            Const NumeroPane As Integer = 51

            'CREO EL NUEVO TAB
            Dim tabreferencia As SAPbouiCOM.Item
            Dim itemreferencia As SAPbouiCOM.Item

            Dim campoTexto As SAPbouiCOM.EditText
            Dim carpetaEdoc As SAPbouiCOM.Folder
            Dim campoCombo As SAPbouiCOM.ComboBox

            Dim ItemLinkedButton As SAPbouiCOM.LinkedButton

            Dim IdTabReferencia As String = ""
            Dim ItemToLinkear As String = ""


            Dim lef1 As Integer = 0
            Dim top1 As Integer = 0
            Dim lef2 As Integer = 0
            Dim top2 As Integer = 0

            Dim anchoGenerico = 0
            Dim altoGenerico = 0

            If oForm.TypeEx = "1250000940" Or oForm.TypeEx = "940" Then

                Dim valorLeft As Integer = oForm.Items.Item("5").Left
                Dim valorTop As Integer = 260
                Dim valorAncho As Integer = oForm.Items.Item("3").Width
                Dim valorAlto As Integer = oForm.Items.Item("3").Height

                lef1 = valorLeft
                top1 = valorTop

                lef2 = 130 + lef1
                top2 = valorTop

                anchoGenerico = valorAncho
                altoGenerico = valorAlto

                IdTabReferencia = "1320000081"
                ItemToLinkear = "5"

            Else

                Dim valorLeft As Integer
                Dim valorTop As Integer

                If Functions.VariablesGlobales._PosicionItemTabX <> "" AndAlso Functions.VariablesGlobales._PosicionItemTabY <> "" Then

                    Try
                        valorLeft = CInt(Functions.VariablesGlobales._PosicionItemTabX)
                        valorTop = CInt(Functions.VariablesGlobales._PosicionItemTabY)
                    Catch ex As Exception

                        valorLeft = oForm.Items.Item("19").Left
                        valorTop = oForm.Items.Item("19").Top

                    End Try

                Else

                    valorLeft = oForm.Items.Item("19").Left
                    valorTop = oForm.Items.Item("19").Top

                End If


                Dim valorAncho As Integer = oForm.Items.Item("18").Width
                Dim valorAlto As Integer = oForm.Items.Item("18").Height

                '------------------------------------------------------------------------------
                'esta logica ya se encontraba


                lef1 = valorLeft
                top1 = valorTop

                lef2 = 130 + lef1
                top2 = valorTop

                anchoGenerico = valorAncho
                altoGenerico = valorAlto

                IdTabReferencia = "138"
                ItemToLinkear = "5"
            End If

            tabreferencia = oForm.Items.Item(IdTabReferencia) ' tab finanzas


            If oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_FOLDER, "TabGenLOC", "SAED Localizacion", tabreferencia.Left + tabreferencia.Width, tabreferencia.Top, tabreferencia.Width, tabreferencia.Height) Then

                carpetaEdoc = oForm.Items.Item("TabGenLOC").Specific
                carpetaEdoc.Pane = NumeroPane
                carpetaEdoc.DataBind.SetBound(True, "", "FolderDS2")

                carpetaEdoc.GroupWith(IdTabReferencia)

                'Aqui empieza la columna 1

                'Establecimiento
                'top1 += 20
                'top2 += 20


                Select Case oForm.TypeEx
                    Case "133", "65300", "60091", "65307"
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc1", "Establecimiento *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc1").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc1", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc1").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc1").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Est")

                        'Punto de Emisión
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc2", "Punto de Emisión *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc2").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc2", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc2").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc2").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Pemi")

                        'Número de Autorización
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc5", "Número de Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc5").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc5", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc5").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc5").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAut")

                        'Tipo de Comprobante
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc3", "Tipo de Comprobante *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc3").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc3", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc3").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc3").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_TipCom")

                        'Forma de Pago
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc6", "Forma de Pago *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc6").LinkTo = ItemToLinkear


                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssloc6", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane, True)
                        oForm.Items.Item("etssloc6").LinkTo = ItemToLinkear


                        campoCombo = oForm.Items.Item("etssloc6").Specific
                        'campoCombo.Item.DisplayDesc = True
                        LlenadoCombosporUDT(campoCombo, "@SS_FORMASDEPAGO") '"@SS_FORMAS_DE_PAGOS")
                        campoCombo.DataBind.SetBound(True, oTabla, "U_SS_FormaPagos")

                        'Remblsos campo vinculo con EL udo de Reembolso

                        'Reembolsos
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc63", "Reembolsos", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc63").LinkTo = ItemToLinkear


                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc63", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc63").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc63").Specific

                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Reembolsos")

                        AddChooseFromList(oForm)

                        campoTexto.ChooseFromListUID = "CFL1"
                        campoTexto.ChooseFromListAlias = "Code"


                        '-----linked object a reembolso
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON, "lbssloc63", "", lef2 - 10, top2, 10, 10, NumeroPane)
                        oForm.Items.Item("lbssloc63").LinkTo = "etssloc63"
                        ItemLinkedButton = oForm.Items.Item("lbssloc63").Specific
                        ItemLinkedButton.LinkedObjectType = "SS_REEMCAB"
                        '----- fin linked
                        '------------------------------------------------------------------------------
                        'Aqui empieza la columna 3 156

                        If oForm.TypeEx = "1250000940" Or oForm.TypeEx = "940" Then


                            Dim valorLeft = oForm.Items.Item("1470000099").Left + 5
                            Dim valorTop = 260

                            lef1 = valorLeft
                            top1 = valorTop

                            lef2 = lef1 + 100
                            top2 = valorTop

                            ItemToLinkear = "1470000099"

                        Else

                            itemreferencia = oForm.Items.Item("156")

                            lef1 = itemreferencia.Left
                            top1 = itemreferencia.Top

                            lef2 = itemreferencia.Left + 100
                            top2 = itemreferencia.Top

                            ItemToLinkear = "11"

                        End If


                        'exportacion

                        'Tipo Exportación

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc70", "Tipo Exportación *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc70").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssloc70", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc70").LinkTo = ItemToLinkear

                        campoCombo = oForm.Items.Item("etssloc70").Specific
                        campoCombo.Item.DisplayDesc = True
                        campoCombo.DataBind.SetBound(True, oTabla, "U_SS_TipoExpor")

                        'Comercio Exterior
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc71", "Comercio Exterior *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc71").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc71", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc71").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc71").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_ComercioExt")

                        'Inicio Neg Exp
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc72", "Inicio Neg Exp *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc72").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc72", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc72").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc72").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_IncoTermFac")


                        'Lugar Neg Exp
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc73", "Lugar Neg Exp *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc73").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc73", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc73").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc73").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_LugIncoTerm")

                        'País Origen
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc74", "País Origen *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc74").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc74", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc74").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc74").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PaisOrigen")


                        'Puerto Embarque
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc75", "Puerto Embarque *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc75").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc75", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc75").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc75").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PuertoEmb")


                        'Puerto Destino
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc76", "Puerto Destino *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc76").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc76", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc76").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc76").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PuertoDestino")

                        'Pais Destino
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc77", "Pais Destino *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc77").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc77", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc77").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc77").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PaisDestino")

                        'País Adquisición
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc78", "País Adquisición *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc78").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc78", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc78").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc78").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PaisAdqui")


                        'Term Total Sin Impuestos
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc79", "Term Total Sin Impuestos *", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc79").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc79", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc79").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc79").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Incotermto")

                        'Tipo de Ingresos Ext
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc80", "Tipo de Ingresos Ext", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc80").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc80", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc80").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc80").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_TipIngExt")

                        'Ingreso Ext. Fue Grav. IR
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc81", "Ingreso Ext. Fue Grav. IR", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc81").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc81", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc81").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc81").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_IngFueGra")

                        'Valor FOB Aduana
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc82", "Valor FOB Aduana", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc82").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc82", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc82").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc82").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_ValFob")

                        'Refrendo Año
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc83", "Refrendo Año", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc83").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc83", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc83").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc83").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_RefrendoAnio")

                        'Refrendo Regímen
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc84", "Refrendo Regímen", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc84").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc84", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc84").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc84").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_RefrendoReg")

                        'Fecha Embarque
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc85", "Fecha Embarque", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc85").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc85", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc85").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc85").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_FechaEmb")


                        'Distrito Aduanero
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc86", "Distrito Aduanero", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc86").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc86", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc86").LinkTo = ItemToLinkear


                        campoTexto = oForm.Items.Item("etssloc86").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_DisAduanero")




                    Case "140", "1250000940", "940"

                        'solo solicitud de traslado
                        'se dibujara el folio asignado ya que no se muestra nativamente en este doc
                        If oForm.TypeEx = "1250000940" Then

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc80", "Folio", oForm.Items.Item("1470000099").Left, oForm.Items.Item("1470000099").Top + 15, oForm.Items.Item("1470000099").Width, oForm.Items.Item("1470000099").Height, 0,, True)
                            oForm.Items.Item("lssloc80").LinkTo = "1470000099"

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc80", "", oForm.Items.Item("1470000101").Left, oForm.Items.Item("1470000101").Top + 15, oForm.Items.Item("1470000101").Width, oForm.Items.Item("1470000101").Height, 0)
                            oForm.Items.Item("etssloc80").LinkTo = "1470000099"
                            oForm.Items.Item("etssloc80").Enabled = False


                        End If


                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc1", "Establecimiento", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc1").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc1", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc1").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc1").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Est")

                        'Punto de Emisión
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc2", "Punto de Emisión", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc2").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc2", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc2").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc2").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Pemi")

                        'Número de Autorización
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc5", "Número de Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc5").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc5", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc5").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc5").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAut")

                        'Tipo de Comprobante
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc3", "Tipo de Comprobante", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc3").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc3", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc3").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc3").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_TipCom")


                        'Fecha Inicio Traslado
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc8", "Fecha Inicio Traslado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc8").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc8", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc8").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc8").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_FecIniTra")

                        'Fecha Fin Traslado
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc9", "Fecha Fin Traslado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc9").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc9", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc9").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc9").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_FecFinTra")

                        ''Hora de Salida
                        'top1 += 15
                        'top2 += 15
                        'oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc10", "Hora de Salida", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        'oForm.Items.Item("lssloc10").LinkTo = ItemToLinkear

                        'oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc10", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        'oForm.Items.Item("etssloc10").LinkTo = ItemToLinkear

                        'campoTexto = oForm.Items.Item("etssloc10").Specific
                        'campoTexto.DataBind.SetBound(True, oTabla, "U_SS_HoraSal")

                        ''Hora de LLegada
                        'top1 += 15
                        'top2 += 15
                        'oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc11", "Hora de LLegada", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        'oForm.Items.Item("lssloc11").LinkTo = ItemToLinkear

                        'oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc11", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        'oForm.Items.Item("etssloc11").LinkTo = ItemToLinkear

                        'campoTexto = oForm.Items.Item("etssloc11").Specific
                        'campoTexto.DataBind.SetBound(True, oTabla, "U_SS_HoraLLeg")

                        'Punto de Partida
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc12", "Punto de Partida", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc12").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc12", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc12").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc12").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PunPart")

                        ' Código de Transporte
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc13", "Código de Transporte", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc13").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc13", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc13").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc13").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_CodTra")


                        'Transportista
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc14", "Transportista", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc14").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc14", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc14").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc14").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Transportista")


                        'Motivo Traslado
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc15", "Motivo Traslado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc15").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssloc15", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc15").LinkTo = ItemToLinkear

                        campoCombo = oForm.Items.Item("etssloc15").Specific
                        campoCombo.Item.DisplayDesc = True
                        campoCombo.DataBind.SetBound(True, oTabla, "U_SS_MotTraslado")

                    Case "141", "65301", "60092"
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc1", "Establecimiento", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc1").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc1", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc1").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc1").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Est")

                        'Punto de Emisión
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc2", "Punto de Emisión", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc2").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc2", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc2").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc2").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Pemi")

                        'Número de Autorización
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc5", "Número de Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc5").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc5", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc5").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc5").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAut")

                        'Tipo de Comprobante
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc3", "Tipo de Comprobante", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc3").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc3", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc3").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc3").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_TipCom")

                        'Forma de Pago
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc6", "Forma de Pago", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc6").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssloc6", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc6").LinkTo = ItemToLinkear

                        campoCombo = oForm.Items.Item("etssloc6").Specific
                        campoCombo.Item.DisplayDesc = True

                        LlenadoCombosporUDT(campoCombo, "@SS_FORMASDEPAGO") '"@SS_FORMAS_DE_PAGOS")
                        campoCombo.DataBind.SetBound(True, oTabla, "U_SS_FormaPagos")

                        'RETENCION
                        'Serie Retención
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc59", "Serie Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc59").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc59", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc59").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc59").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_SerieRet")

                        'Secuencial Retención
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc60", "Sec Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc60").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc60", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc60").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc60").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_SecRet")


                        'Num.Autor.Retención
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc61", "Num.Aut.Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc61").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc61", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc61").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc61").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAutRet")

                        'Fecha Retención
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc62", "Fecha Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc62").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc62", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc62").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc62").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_FecRet")

                        'Sustento Tributario
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc4", "Sustento Tributario", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc4").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc4", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc4").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc4").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_SUSTRIB")


                        'Remblsos campo vinculo con EL udo de Reembolso

                        'Reembolsos
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc63", "Reembolsos", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc63").LinkTo = ItemToLinkear


                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc63", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc63").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc63").Specific

                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Reembolsos")

                        AddChooseFromList(oForm)

                        campoTexto.ChooseFromListUID = "CFL1"
                        campoTexto.ChooseFromListAlias = "Code"


                        '-----linked object a reembolso
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON, "lbssloc63", "", lef2 - 10, top2, 10, 10, NumeroPane)
                        oForm.Items.Item("lbssloc63").LinkTo = "etssloc63"
                        ItemLinkedButton = oForm.Items.Item("lbssloc63").Specific
                        ItemLinkedButton.LinkedObjectType = "SS_REEMCAB"
                        '----- fin linked


                    Case "179", "65303", "181"
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc1", "Establecimiento", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc1").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc1", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc1").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc1").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Est")


                        'Punto de Emisión
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc2", "Punto de Emisión", lef1, top1, anchoGenerico, altoGenerico, NumeroPane,, True)
                        oForm.Items.Item("lssloc2").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc2", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc2").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc2").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_Pemi")

                        'Número de Autorización
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc5", "Número de Autorización", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc5").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc5", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc5").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc5").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAut")

                        'Tipo de Comprobante
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc3", "Tipo de Comprobante", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc3").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc3", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc3").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc3").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_TipCom")

                        '------------Info Doc Relacionado--------------------------
                        'Establecimiento fac relacionado
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc50", "Estab.Relacionado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc50").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc50", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc50").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc50").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_EstFacRel")

                        'Punto Emisión.Fact.Relac
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc51", "P.Emi.Relacionado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc51").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc51", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc51").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc51").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_PunEmiFacRel")

                        'Num.Autoriz.Fact.Relac
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc52", "Num.Aut.Relacionado", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc52").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc52", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc52").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc52").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAutFacRel")

                        'Número Fact.Relac
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc53", "Número.Fact.Relac", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc53").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc53", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc53").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc53").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumFacRel")


                        'Fecha Emisión Doc.Vtas
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc54", "Fecha.Emi.Doc.Vtas", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc54").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc54", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc54").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc54").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_FecEmiDocRel")

                        'Tipo Doc Aplica
                        top1 += 15
                        top2 += 15
                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc55", "Tipo Doc Aplica", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("lssloc55").LinkTo = ItemToLinkear

                        oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc55", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                        oForm.Items.Item("etssloc55").LinkTo = ItemToLinkear

                        campoTexto = oForm.Items.Item("etssloc55").Specific
                        campoTexto.DataBind.SetBound(True, oTabla, "U_SS_TipDocAplica")


                        If oForm.TypeEx = "179" Then
                            'Motivo Nota de Credito
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc56", "Motivo de NC", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc56").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc56", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane, True)
                            oForm.Items.Item("etssloc56").LinkTo = ItemToLinkear
                            campoTexto = oForm.Items.Item("etssloc56").Specific
                            campoTexto.DataBind.SetBound(True, oTabla, "U_SS_MOTIVO_NC")
                            'campoCombo = oForm.Items.Item("etssloc56").Specific
                            ''campoCombo.Item.DisplayDesc = True
                            'LlenadoCombosporUDT(campoCombo, "@SS_MOTIVOS_NC")
                            'campoCombo.DataBind.SetBound(True, oTabla, "U_SS_MOTIVO_NC")


                        ElseIf oForm.TypeEx = "65303" Then

                            'Forma de Pago
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc6", "Forma de Pago", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc6").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssloc6", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("etssloc6").LinkTo = ItemToLinkear

                            campoCombo = oForm.Items.Item("etssloc6").Specific
                            campoCombo.Item.DisplayDesc = True

                            LlenadoCombosporUDT(campoCombo, "@SS_FORMASDEPAGO") ', "@SS_FORMAS_DE_PAGOS")
                            campoCombo.DataBind.SetBound(True, oTabla, "U_SS_FormaPagos")


                        ElseIf oForm.TypeEx = "181" Then
                            'Forma de Pago
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc6", "Forma de Pago", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc6").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "etssloc6", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("etssloc6").LinkTo = ItemToLinkear

                            campoCombo = oForm.Items.Item("etssloc6").Specific
                            campoCombo.Item.DisplayDesc = True

                            LlenadoCombosporUDT(campoCombo, "@SS_FORMASDEPAGO") ' "@SS_FORMAS_DE_PAGOS")
                            campoCombo.DataBind.SetBound(True, oTabla, "U_SS_FormaPagos")

                            'RETENCION
                            'Serie Retención
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc59", "Serie Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc59").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc59", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("etssloc59").LinkTo = ItemToLinkear

                            campoTexto = oForm.Items.Item("etssloc59").Specific
                            campoTexto.DataBind.SetBound(True, oTabla, "U_SS_SerieRet")

                            'Secuencial Retención
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc60", "Sec Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc60").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc60", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("etssloc60").LinkTo = ItemToLinkear

                            campoTexto = oForm.Items.Item("etssloc60").Specific
                            campoTexto.DataBind.SetBound(True, oTabla, "U_SS_SecRet")


                            'Num.Autor.Retención
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc61", "Num.Aut.Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc61").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc61", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("etssloc61").LinkTo = ItemToLinkear

                            campoTexto = oForm.Items.Item("etssloc61").Specific
                            campoTexto.DataBind.SetBound(True, oTabla, "U_SS_NumAutRet")

                            'Fecha Retención
                            top1 += 15
                            top2 += 15
                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_STATIC, "lssloc62", "Fecha Retención", lef1, top1, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("lssloc62").LinkTo = ItemToLinkear

                            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_EDIT, "etssloc62", "", lef2, top2, anchoGenerico, altoGenerico, NumeroPane)
                            oForm.Items.Item("etssloc62").LinkTo = ItemToLinkear

                            campoTexto = oForm.Items.Item("etssloc62").Specific
                            campoTexto.DataBind.SetBound(True, oTabla, "U_SS_FecRet")


                        End If


                    Case Else




                End Select





            End If

            'oForm.PaneLevel = 1

        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Tab de pruebas " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)
    End Sub


    Private Sub ObtenerEnlacesURLyGenerarRQ(ByRef mForm As SAPbouiCOM.Form, FormularioID As String)



        Dim EnlaceQR = GenerarEnlaceQR(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_CLAVE_ACCESO", 0).Trim)

        Dim EnlaceQRLIQ = GenerarEnlaceQR(mForm.DataSources.DBDataSources.Item(oTabla).GetValue("U_LQ_CLAVE", 0).Trim)

        Dim NombreImagen As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0).Trim & "_" & FormularioID

        Dim NombreImagenLIQ As String = mForm.DataSources.DBDataSources.Item(oTabla).GetValue("DocEntry", 0).Trim & "_" & FormularioID & "_LIQ"

        Try
            'item creado para guardar la URL del pdf de documentos
            mForm.Items.Item("etssut6").Specific.string = EnlaceQR
        Catch ex As Exception

        End Try

        Try
            'item creado para guardar la URL de documentos de Liquidacion
            mForm.Items.Item("etssut25").Specific.string = EnlaceQRLIQ
        Catch ex As Exception

        End Try

        CargarQRenPantalla(mForm, EnlaceQR, NombreImagen)
        CargarQRenPantalla(mForm, EnlaceQRLIQ, NombreImagenLIQ, True)

    End Sub

    Private Sub AddChooseFromList(ByRef oForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            'Dim oCons As SAPbouiCOM.Conditions
            'Dim oCon As SAPbouiCOM.Condition

            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            ' oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.ObjectType = "SS_REEMCAB"
            oCFLCreationParams.UniqueID = "CFL1"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            'oCons = oCFL.GetConditions()

            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)

            'oCFLCreationParams.UniqueID = "CFL2"
            'oCFL = oCFLs.Add(oCFLCreationParams)

        Catch
            MsgBox(Err.Description)
        End Try
    End Sub


    Private Sub ValidacionesPrevioCreacionMaestro(pform As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)

        If pform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And pform.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Return



        Try
            Dim cvalid As String = ""
            Dim identifi As String = ""

            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipoId", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo Id] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False
                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipoSN", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo S.N] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False
                Exit Sub
            End If

            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipoCon", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo Contribuyente] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False
                Exit Sub
            End If


            ' Validaciones para El Tipo

            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipoId", 0).Trim
            identifi = pform.DataSources.DBDataSources.Item(0).GetValue("LicTradNum", 0).Trim

            If cvalid = "C" And identifi.Length <> 10 Then
                rSboApp.SetStatusBarMessage(NombreAddon & ": El Número de Identificacion Ingresado No Corresponde a una CEDULA , Validar por Favor")

                BubbleEvent = False
                Exit Sub

            ElseIf cvalid = "R" And identifi.Length <> 13 Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El Número de Identificacion Ingresado No Corresponde a un RUC , Validar por Favor")

                BubbleEvent = False
                Exit Sub

            End If

            ' Campos Nuevos
            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_ParteRel", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Parte Relacionada] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipoRegi", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tip Regímen Fiscal Ext] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_AplDobTri", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Aplica Doble Tributación] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PagoLocExt", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Pago Residente] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisEfecPago", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) País se Efectua Pago] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PagExtSujRetNorLeg", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Pag.Sujeto a Ret N.Legal] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If


            cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PagoRegFis", 0).Trim

            If cvalid = "" Then

                rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Pag. Regimen Fiscal] Se encuentra Vacio, Validar por Favor")

                BubbleEvent = False

                Exit Sub
            End If



        Catch ex As Exception

            rSboApp.SetStatusBarMessage(ex.Message)

            BubbleEvent = False

            Utilitario.Util_Log.Escribir_Log(" Error Func ValidacionesPrevioCreacionMaestro " & ex.Message, "EventosLE")

        End Try

    End Sub


    Private Sub CrearBotonImpresion(oFormx As SAPbouiCOM.Form)


        Dim item222 = oFormx.Items.Item("222")


        Dim _left = item222.Left
        Dim _top = item222.Top + 20
        Dim _alto = item222.Height + 5
        Dim _ancho = item222.Width

        ' COMENTARIO
        Try
            oFuncionesB1.creaControl(oFormx, SAPbouiCOM.BoFormItemTypes.it_BUTTON, "btprints", "Imprimir Documentos", _left, _top, _ancho, _alto)
        Catch ex As Exception
        End Try

        ' btnBotonImprimir = oFormx.Items.Item("btprints").Specific

    End Sub

    Private Sub ValidacionesPrevioCreacionDocumento(pform As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)


        If pform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Return


        Dim SerieDoc As String = ""
        SerieDoc = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

        Dim QrySerieDoc As String = "Select ""Code"" from ""@SS_SERD"" where ""U_SerId"" = '" & SerieDoc & "'"
        Dim rsSerieDoc = oFuncionesB1.getRecordSet(QrySerieDoc)

        If rsSerieDoc.RecordCount > 0 Then



            Try

                SeteaTipoTabla_FormTypeEx(pform.TypeEx)


                Try

                    If oTipoTabla = "FEC" Then

                        Dim query As String = ""

                        Dim serie2 = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                            query = "Select IFNULL(""U_Establec"",'') as EST,IFNULL(""U_PuntoEmi"",'') as PDE,IFNULL(""U_TipoD"",'') as TDOC FROM ""@SS_SERD"" WHERE ""U_SerId"" = '" + serie2 + "'"

                        Else

                            query = "Select ISNULL(""U_Establec"",'') as EST,ISNULL(""U_PuntoEmi"",'') as PDE,ISNULL(""U_TipoD"",'') as TDOC FROM ""@SS_SERD"" WHERE ""U_SerId"" = '" + serie2 + "'"

                        End If


                        Dim rs = oFuncionesB1.getRecordSet(query)

                        If rs.RecordCount > 0 Then

                            tipoDoc2 = rs.Fields.Item("TDOC").Value

                        End If

                        oFuncionesB1.Release(rs)


                    End If

                Catch ex As Exception

                    Utilitario.Util_Log.Escribir_Log("Error al Validar Tipo de Documento desde la pantalla de Serie " & ex.Message, "ValidacionPrevioCreacionDocumento")

                End Try

                'primera validacion, establecimientos y punto de emision llenos

                Dim est As String = ""
                Dim pm As String = ""
                Dim serie As String = ""
                Dim tipoComp As String = ""
                Dim cvalid As String = ""


                est = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Est", 0).Trim
                pm = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Pemi", 0).Trim

                tipoComp = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim

                'serie = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                'tipoComp = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim


                If est = "" Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Establecimiento] Se encuentra Vacio, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub

                ElseIf est.Length <> 3 Then


                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Establecimiento] No Cumple con el Formato Correcto!, Validar por Favor")

                    BubbleEvent = False

                    Exit Sub

                End If

                If pm = "" Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto de Emisión] Se encuentra Vacio, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub

                ElseIf pm.Length <> 3 Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto de Emisión]  No Cumple con el Formato Correcto!, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub


                End If

                If tipoComp = "" Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo de Comprobante] Se encuentra Vacio, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub
                End If


                ' valido Folio

                'Try

                '    cvalid = pform.Items.Item("pfolio").Specific.Caption

                '    If cvalid = "" Then

                '        rSboApp.SetStatusBarMessage(NombreAddon & ": No se Genero un Folio con formato Correcto, Validar por Favor")

                '        BubbleEvent = False
                '        Exit Sub
                '    End If


                'Catch ex As Exception
                '    Utilitario.Util_Log.Escribir_Log("Error al no tener Folio " & ex.Message, "EventosLE")
                'End Try

                'Validacion de casos especificos de Facturas
                'Casos Reembolsos y Exportacion

                ' si es factura de deudores, deudor + pago , reserva
                If oTipoTabla = "FCE" OrElse oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then


                    ' 01 es exportacion
                    If tipoComp = "01" Then

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_ComercioExt", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Comercio Exterior] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_IncoTermFac", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Inicio Neg Exp] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_LugIncoTerm", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Lugar Neg Exp] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisOrigen", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) País Origen] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PuertoEmb", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Puerto Embarque] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PuertoDestino", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Puerto Destino] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If



                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisDestino", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Pais Destino] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If



                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisAdqui", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) País Adquisición] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Incotermto", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Term Total Sin Impuestos] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If



                    End If




                End If


                ' El campo de Reembolso se validara en facturas de Compras y Ventas

                If oTipoTabla = "FEC" Or oTipoTabla = "FCE" OrElse oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then


                    ' 41 reembolsos
                    If tipoComp = "41" Then

                        ' validacion ue este asignado

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Reembolsos", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Reembolsos] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        ' validacion que sea unico
                        Dim totaldetectados = 0
                        Dim qry As String = ""

                        qry = "Select COUNT(*) as tot FROM OINV Where ""U_SS_Reembolsos"" = '" + cvalid + "'"

                        totaldetectados = CInt(oFuncionesB1.getRSvalue(qry, "tot", "0"))

                        If (totaldetectados > 0) Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El Codigo de Reembolso (" + cvalid + ") ya esta asignado a un Documento de Ventas, Por Favor Genere uno Nuevo")

                            BubbleEvent = False
                            Exit Sub

                        End If

                        totaldetectados = 0

                        qry = "Select COUNT(*) as tot FROM OPCH Where ""U_SS_Reembolsos"" = '" + cvalid + "'"

                        totaldetectados = CInt(oFuncionesB1.getRSvalue(qry, "tot", "0"))

                        If (totaldetectados > 0) Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El Codigo de Reembolso (" + cvalid + ") ya esta asignado a un Documento de Compras, Por Favor Genere uno Nuevo")

                            BubbleEvent = False
                            Exit Sub

                        End If



                    End If


                End If



                ' se validara la nota credito/ debito

                If oTipoTabla = "NDE" OrElse oTipoTabla = "NCE" Then



                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_EstFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Estab.Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    ' se valida en logica que el tamano sea correcto

                    If cvalid.Length <> 3 Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Estab.Fact.Relac] No cumple el formato Correcto!, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If



                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PunEmiFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto Emisión.Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    If cvalid.Length <> 3 Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto Emisión.Fact.Relac] No cumple el formato Correcto!, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumAutFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Num.Autoriz.Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    If cvalid.Length <> 10 AndAlso cvalid.Length <> 49 Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Num.Autoriz.Fact.Relac] No cumple el formato Correcto!, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If




                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Número Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecEmiDocRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Emisión Doc.Vtas] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipDocAplica", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo Doc Aplica] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    If oTipoTabla = "NCE" Then

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_MOTIVO_NC", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Motivo Nota de Credito] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                    End If ' motivo nc


                End If ' notas C/D


                ' forma de pago

                If oTipoTabla = "FCE" OrElse oTipoTabla = "NDE" OrElse oTipoTabla = "FEC" Or oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FormaPagos", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Forma de Pago] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If

                End If

                ' Para Entregas

                If oTipoTabla = "GR" OrElse oTipoTabla = "GRT" Then


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecIniTra", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Inicio Traslado] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If



                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecFinTra", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Fin Traslado] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    'cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_HoraSal", 0).Trim

                    'If cvalid = "" Then

                    '    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Hora de Salida] Se encuentra Vacio, Validar por Favor")

                    '    BubbleEvent = False

                    '    Exit Sub
                    'End If


                    'cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_HoraLLeg", 0).Trim

                    'If cvalid = "" Then

                    '    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Hora de LLegada] Se encuentra Vacio, Validar por Favor")

                    '    BubbleEvent = False

                    '    Exit Sub
                    'End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PunPart", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto de Partida] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_CodTra", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS)  Código de Transporte] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Transportista", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Transportista] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_MotTraslado", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Motivo Traslado] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                End If



                If (oTipoTabla = "FEC" And tipoDoc2 = "LQ") Or (oTipoTabla = "FEC" And tipoDoc2 = "LQRT") Then

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("FolioNum", 0).Trim

                    If cvalid <> "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo Número de Folio debe estar Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                End If


                If (oTipoTabla = "FEC" And tipoDoc2 = "RT") Then


                    Dim docmun As String = ""
                    Dim query As String = ""
                    Dim serie2 As String = ""
                    ' Dim rs As SAPbobsCOM.Recordset

                    docmun = pform.DataSources.DBDataSources.Item(0).GetValue("DocNum", 0).Trim
                    serie2 = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                    query = Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._SS_CalculoFolioQRY, sKey)
                    query = query.Replace("@SERIE", "'" & serie2 & "'")
                    query = query.Replace("@DOCNUM", docmun)
                    query = query.Replace("@TIPDOC", tipoDoc2)

                    Dim rs = oFuncionesB1.getRecordSet(query)

                    Dim secCalculadoret As String = ""

                    If rs.RecordCount > 0 Then

                        secCalculadoret = rs.Fields.Item("Secuencial").Value

                    End If



                    Dim secRetencion = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SecRet", 0).Trim

                    If secRetencion = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] No debe ser Modificado, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub


                    ElseIf secRetencion <> secCalculadoret Then


                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] Generado no debe Modificarse, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub

                    End If


                    '-------------------
                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("FolioNum", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [Número de Folio] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumAut", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Número de Autorización] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If




                    ' que el folio no se repita por cada socio de Negocios
                    'se requiere socio negoio y folio

                    Dim fnum = pform.DataSources.DBDataSources.Item(0).GetValue("FolioNum", 0).Trim

                    Dim ccode = pform.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim

                    Dim _est = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Est", 0).Trim

                    Dim _ptoemi = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Pemi", 0).Trim

                    Dim _TipComSus = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim


                    Dim xquery As String = ""

                    xquery = "Select count(*) as ctot from ""OPCH"" where ""CardCode""='" & ccode & "' AND ""U_SS_TipCom"" ='" & _TipComSus & "' AND ""U_SS_Est"" ='" & _est & "' AND ""U_SS_Pemi"" ='" & _ptoemi & "' AND ""FolioNum""=" & fnum.ToString

                    rs = oFuncionesB1.getRecordSet(xquery)

                    If rs.RecordCount > 0 Then

                        Dim va = rs.Fields.Item("ctot").Value

                        If CInt(va) > 0 Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El Número de Folio ya se Encuentra Ocupado para el Socio de Negocios, Validar por Favor")

                            BubbleEvent = False

                            Exit Sub

                        End If


                    End If



                    ' se libera todo el RecordSet
                    oFuncionesB1.Release(rs)



                End If


                If (oTipoTabla = "FEC" And tipoDoc2 = "LQRT") Then


                    Try

                        Dim secRetencion = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SecRet", 0).Trim

                        Dim valorItem = pform.Items.Item("pfolio").Specific.Caption

                        If secRetencion <> valorItem Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] Generado no debe Modificarse, Validar por Favor")

                            BubbleEvent = False

                            Exit Sub
                        End If


                    Catch ex As Exception


                        Utilitario.Util_Log.Escribir_Log("Error validar folio Item + Secuencial Retencion " & ex.Message, "EventosLE")

                    End Try


                End If


                If (oTipoTabla = "FEC" And tipoDoc2 = "LQ") Then


                    Try

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SUSTRIB", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Sustento Tributario] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False

                            Exit Sub
                        End If


                    Catch ex As Exception


                        Utilitario.Util_Log.Escribir_Log("Error al Setear Sustento Tributario " & ex.Message, "EventosLE")

                    End Try


                End If


                If (oTipoTabla = "FEC" And tipoDoc2 = "RT") Or (oTipoTabla = "FEC" And tipoDoc2 = "LQRT") Then




                    '-------------------------


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SerieRet", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Serie Retención] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SecRet", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    'cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumAutRet", 0).Trim

                    'If cvalid = "" Then

                    '    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Num.Autor.Retención] Se encuentra Vacio, Validar por Favor")

                    '    BubbleEvent = False

                    '    Exit Sub
                    'End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecRet", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Retención] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    ' Nuevos Campos





                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SUSTRIB", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Sustento Tributario] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If





                End If


            Catch ex As Exception

                rSboApp.SetStatusBarMessage(ex.Message)

                BubbleEvent = False

            End Try

        End If


    End Sub

    Private Sub ValidacionesPrevioCreacionDocumento2(pform As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)


        If pform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Return


        Dim SerieDoc As String = ""
        SerieDoc = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

        Dim QrySerieDoc As String = "Select ""Code"" from ""@SS_SERD"" where ""U_SerId"" = '" & SerieDoc & "'"
        Dim rsSerieDoc = oFuncionesB1.getRecordSet(QrySerieDoc)

        If rsSerieDoc.RecordCount > 0 Then



            Try

                SeteaTipoTabla_FormTypeEx(pform.TypeEx)


                Try

                    If oTipoTabla = "REE" Or oTipoTabla = "RER" Then

                        Dim query As String = ""

                        Dim serie2 = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                            query = "Select IFNULL(""U_Establec"",'') as EST,IFNULL(""U_PuntoEmi"",'') as PDE,IFNULL(""U_TipoD"",'') as TDOC FROM ""@SS_SERD"" WHERE ""U_SerId"" = '" + serie2 + "'"

                        Else

                            query = "Select ISNULL(""U_Establec"",'') as EST,ISNULL(""U_PuntoEmi"",'') as PDE,ISNULL(""U_TipoD"",'') as TDOC FROM ""@SS_SERD"" WHERE ""U_SerId"" = '" + serie2 + "'"

                        End If


                        Dim rs = oFuncionesB1.getRecordSet(query)

                        If rs.RecordCount > 0 Then

                            tipoDoc2 = rs.Fields.Item("TDOC").Value

                        End If

                        oFuncionesB1.Release(rs)


                    End If

                Catch ex As Exception

                    Utilitario.Util_Log.Escribir_Log("Error al Validar Tipo de Documento desde la pantalla de Serie " & ex.Message, "ValidacionPrevioCreacionDocumento")

                End Try

                'primera validacion, establecimientos y punto de emision llenos

                Dim est As String = ""
                Dim pm As String = ""
                Dim serie As String = ""
                Dim tipoComp As String = ""
                Dim cvalid As String = ""


                est = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Est", 0).Trim
                pm = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Pemi", 0).Trim

                tipoComp = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim

                'serie = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                'tipoComp = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim


                If est = "" Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Establecimiento] Se encuentra Vacio, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub

                ElseIf est.Length <> 3 Then


                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Establecimiento] No Cumple con el Formato Correcto!, Validar por Favor")

                    BubbleEvent = False

                    Exit Sub

                End If

                If pm = "" Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto de Emisión] Se encuentra Vacio, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub

                ElseIf pm.Length <> 3 Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto de Emisión]  No Cumple con el Formato Correcto!, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub


                End If

                If tipoComp = "" Then

                    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo de Comprobante] Se encuentra Vacio, Validar por Favor")

                    BubbleEvent = False
                    Exit Sub
                End If


                ' valido Folio

                'Try

                '    cvalid = pform.Items.Item("pfolio").Specific.Caption

                '    If cvalid = "" Then

                '        rSboApp.SetStatusBarMessage(NombreAddon & ": No se Genero un Folio con formato Correcto, Validar por Favor")

                '        BubbleEvent = False
                '        Exit Sub
                '    End If


                'Catch ex As Exception
                '    Utilitario.Util_Log.Escribir_Log("Error al no tener Folio " & ex.Message, "EventosLE")
                'End Try

                'Validacion de casos especificos de Facturas
                'Casos Reembolsos y Exportacion

                ' si es factura de deudores, deudor + pago , reserva
                If oTipoTabla = "FCE" OrElse oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then


                    ' 01 es exportacion
                    If tipoComp = "01" Then

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_ComercioExt", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Comercio Exterior] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_IncoTermFac", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Inicio Neg Exp] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_LugIncoTerm", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Lugar Neg Exp] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisOrigen", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) País Origen] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PuertoEmb", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Puerto Embarque] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PuertoDestino", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Puerto Destino] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If



                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisDestino", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Pais Destino] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If



                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PaisAdqui", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) País Adquisición] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Incotermto", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Term Total Sin Impuestos] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If



                    End If




                End If


                ' El campo de Reembolso se validara en facturas de Compras y Ventas

                If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "FCE" OrElse oTipoTabla = "FRE" Or oTipoTabla = "FAE" Then


                    ' 41 reembolsos
                    If tipoComp = "41" Then

                        ' validacion ue este asignado

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Reembolsos", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Reembolsos] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                        ' validacion que sea unico
                        Dim totaldetectados = 0
                        Dim qry As String = ""

                        qry = "Select COUNT(*) as tot FROM OINV Where ""U_SS_Reembolsos"" = '" + cvalid + "'"

                        totaldetectados = CInt(oFuncionesB1.getRSvalue(qry, "tot", "0"))

                        If (totaldetectados > 0) Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El Codigo de Reembolso (" + cvalid + ") ya esta asignado a un Documento de Ventas, Por Favor Genere uno Nuevo")

                            BubbleEvent = False
                            Exit Sub

                        End If

                        totaldetectados = 0

                        qry = "Select COUNT(*) as tot FROM OPCH Where ""U_SS_Reembolsos"" = '" + cvalid + "'"

                        totaldetectados = CInt(oFuncionesB1.getRSvalue(qry, "tot", "0"))

                        If (totaldetectados > 0) Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El Codigo de Reembolso (" + cvalid + ") ya esta asignado a un Documento de Compras, Por Favor Genere uno Nuevo")

                            BubbleEvent = False
                            Exit Sub

                        End If



                    End If


                End If



                ' se validara la nota credito/ debito

                If oTipoTabla = "NDE" OrElse oTipoTabla = "NCE" Then



                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_EstFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Estab.Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    ' se valida en logica que el tamano sea correcto

                    If cvalid.Length <> 3 Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Estab.Fact.Relac] No cumple el formato Correcto!, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If



                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PunEmiFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto Emisión.Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    If cvalid.Length <> 3 Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto Emisión.Fact.Relac] No cumple el formato Correcto!, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumAutFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Num.Autoriz.Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    If cvalid.Length <> 10 AndAlso cvalid.Length <> 49 Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Num.Autoriz.Fact.Relac] No cumple el formato Correcto!, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If




                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumFacRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Número Fact.Relac] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecEmiDocRel", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Emisión Doc.Vtas] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipDocAplica", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Tipo Doc Aplica] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If


                    If oTipoTabla = "NCE" Then

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_MOTIVO_NC", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Motivo Nota de Credito] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False
                            Exit Sub
                        End If


                    End If ' motivo nc


                End If ' notas C/D


                ' forma de pago

                If oTipoTabla = "FCE" OrElse oTipoTabla = "NDE" OrElse oTipoTabla = "REE" OrElse oTipoTabla = "FRE" OrElse oTipoTabla = "FAE" OrElse oTipoTabla = "RER" Then

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FormaPagos", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Forma de Pago] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False
                        Exit Sub
                    End If

                End If

                ' Para Entregas

                If oTipoTabla = "GRE" OrElse oTipoTabla = "TRE" OrElse oTipoTabla = "TLE" Then


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecIniTra", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Inicio Traslado] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If



                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecFinTra", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Fin Traslado] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    'cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_HoraSal", 0).Trim

                    'If cvalid = "" Then

                    '    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Hora de Salida] Se encuentra Vacio, Validar por Favor")

                    '    BubbleEvent = False

                    '    Exit Sub
                    'End If


                    'cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_HoraLLeg", 0).Trim

                    'If cvalid = "" Then

                    '    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Hora de LLegada] Se encuentra Vacio, Validar por Favor")

                    '    BubbleEvent = False

                    '    Exit Sub
                    'End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_PunPart", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Punto de Partida] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_CodTra", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS)  Código de Transporte] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Transportista", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Transportista] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_MotTraslado", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Motivo Traslado] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                End If



                If ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "LQ") Or ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "LQRT") Then

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("FolioNum", 0).Trim

                    If cvalid <> "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo Número de Folio debe estar Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                End If


                If ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "RT") Then


                    Dim docmun As String = ""
                    Dim query As String = ""
                    Dim serie2 As String = ""
                    ' Dim rs As SAPbobsCOM.Recordset

                    docmun = pform.DataSources.DBDataSources.Item(0).GetValue("DocNum", 0).Trim
                    serie2 = pform.DataSources.DBDataSources.Item(0).GetValue("Series", 0).Trim

                    query = Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._SS_CalculoFolioQRY, sKey)
                    query = query.Replace("@SERIE", "'" & serie2 & "'")
                    query = query.Replace("@DOCNUM", docmun)
                    query = query.Replace("@TIPDOC", tipoDoc2)

                    Dim rs = oFuncionesB1.getRecordSet(query)

                    Dim secCalculadoret As String = ""

                    If rs.RecordCount > 0 Then

                        secCalculadoret = rs.Fields.Item("Secuencial").Value

                    End If



                    Dim secRetencion = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SecRet", 0).Trim

                    If secRetencion = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] No debe ser Modificado, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub


                    ElseIf secRetencion <> secCalculadoret Then


                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] Generado no debe Modificarse, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub

                    End If


                    '-------------------
                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("FolioNum", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [Número de Folio] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If

                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumAut", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Número de Autorización] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If




                    ' que el folio no se repita por cada socio de Negocios
                    'se requiere socio negoio y folio

                    Dim fnum = pform.DataSources.DBDataSources.Item(0).GetValue("FolioNum", 0).Trim

                    Dim ccode = pform.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim

                    Dim _est = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Est", 0).Trim

                    Dim _ptoemi = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_Pemi", 0).Trim

                    Dim _TipComSus = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_TipCom", 0).Trim


                    Dim xquery As String = ""

                    xquery = "Select count(*) as ctot from ""OPCH"" where ""CardCode""='" & ccode & "' AND ""U_SS_TipCom"" ='" & _TipComSus & "' AND ""U_SS_Est"" ='" & _est & "' AND ""U_SS_Pemi"" ='" & _ptoemi & "' AND ""FolioNum""=" & fnum.ToString

                    'xquery = "Select count(*) as ctot from ""OPCH"" where ""CardCode""='" & ccode & "' AND ""FolioNum""=" & fnum.ToString

                    rs = oFuncionesB1.getRecordSet(xquery)

                    If rs.RecordCount > 0 Then

                        Dim va = rs.Fields.Item("ctot").Value

                        If CInt(va) > 0 Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El Número de Folio ya se Encuentra Ocupado para el Socio de Negocios, Validar por Favor")

                            BubbleEvent = False

                            Exit Sub

                        End If


                    End If



                    ' se libera todo el RecordSet
                    oFuncionesB1.Release(rs)



                End If


                If ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "LQRT") Then


                    Try

                        Dim secRetencion = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SecRet", 0).Trim

                        Dim valorItem = pform.Items.Item("pfolio").Specific.Caption

                        If secRetencion <> valorItem Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] Generado no debe Modificarse, Validar por Favor")

                            BubbleEvent = False

                            Exit Sub
                        End If


                    Catch ex As Exception


                        Utilitario.Util_Log.Escribir_Log("Error validar folio Item + Secuencial Retencion " & ex.Message, "EventosLE")

                    End Try


                End If


                If ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "LQ") Then


                    Try

                        cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SUSTRIB", 0).Trim

                        If cvalid = "" Then

                            rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Sustento Tributario] Se encuentra Vacio, Validar por Favor")

                            BubbleEvent = False

                            Exit Sub
                        End If


                    Catch ex As Exception


                        Utilitario.Util_Log.Escribir_Log("Error al Setear Sustento Tributario " & ex.Message, "EventosLE")

                    End Try


                End If


                If ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "RT") Or ((oTipoTabla = "REE" Or oTipoTabla = "RER") And tipoDoc2 = "LQRT") Then




                    '-------------------------


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SerieRet", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Serie Retención] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SecRet", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Secuencial Retención] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    'cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_NumAutRet", 0).Trim

                    'If cvalid = "" Then

                    '    rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Num.Autor.Retención] Se encuentra Vacio, Validar por Favor")

                    '    BubbleEvent = False

                    '    Exit Sub
                    'End If


                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_FecRet", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Fecha Retención] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If


                    ' Nuevos Campos





                    cvalid = pform.DataSources.DBDataSources.Item(0).GetValue("U_SS_SUSTRIB", 0).Trim

                    If cvalid = "" Then

                        rSboApp.SetStatusBarMessage(NombreAddon & ": El campo [(SS) Sustento Tributario] Se encuentra Vacio, Validar por Favor")

                        BubbleEvent = False

                        Exit Sub
                    End If





                End If


            Catch ex As Exception

                rSboApp.SetStatusBarMessage(ex.Message)

                BubbleEvent = False

            End Try

        End If


    End Sub


#Region "Logica para Impresion"

    Private Sub ProcesoImpresion(Formulario As String, Docenrty As Integer)


        Dim ImpresoraIndvCORRECTA = ""
        Dim ImpresoraCompCORRECTA = ""



        Select Case Formulario
            Case "140"
                ' entregas
                Dim confentrega = (From k In Functions.VariablesGlobales._SS_ConfImpresoras Where k.TipoDocumento = Formulario).FirstOrDefault

                If Not IsNothing(confentrega) Then

                    rSboApp.SetStatusBarMessage("Recuperando Impresora..",, False)
                    ImpresoraCompCORRECTA = ObtenerImpresoraPorNombreParcial(confentrega.ImpresoraCompartida)
                    ' rSboApp.SetStatusBarMessage("Comfigurando Impresora.." & ImpresoraCompCORRECTA,, False)
                    '  SetearImpresora(confentrega.IdReporte, ImpresoraCompCORRECTA)
                    rSboApp.SetStatusBarMessage("Imprimiendo ..",, False)
                    ' ImprimirCR(confentrega.IdReporte, Docenrty)
                    ImprimirViaCrystal(confentrega.IdReporte, Docenrty, ImpresoraCompCORRECTA)

                End If


            Case "133", "60091"
                'facturas deudor
                Dim conffacturaDeudor = (From k In Functions.VariablesGlobales._SS_ConfImpresoras Where k.TipoDocumento = Formulario And (k.Usuario = "0" Or k.Usuario = rCompany.UserSignature.ToString)).ToList

                If Not IsNothing(conffacturaDeudor) Then

                    For Each x As Functions.DatosImpresora In conffacturaDeudor

                        ' si hay usuario
                        If x.Usuario <> 0 Then

                            rSboApp.SetStatusBarMessage("Recuperando Impresora..",, False)
                            ImpresoraIndvCORRECTA = ObtenerImpresoraPorNombreParcial(x.ImpresoraIndividual)
                            'rSboApp.SetStatusBarMessage("Comfigurando Impresora.." & ImpresoraIndvCORRECTA,, False)
                            'SetearImpresora(x.IdReporte, ImpresoraIndvCORRECTA)
                            rSboApp.SetStatusBarMessage("Imprimiendo ..",, False)
                            ' ImprimirCR(x.IdReporte, Docenrty)

                            ImprimirViaCrystal(x.IdReporte, Docenrty, ImpresoraIndvCORRECTA)

                        Else

                            rSboApp.SetStatusBarMessage("Recuperando Impresora..",, False)
                            ImpresoraCompCORRECTA = ObtenerImpresoraPorNombreParcial(x.ImpresoraCompartida)
                            'rSboApp.SetStatusBarMessage("Comfigurando Impresora.." & ImpresoraCompCORRECTA,, False)
                            'SetearImpresora(x.IdReporte, ImpresoraCompCORRECTA)
                            rSboApp.SetStatusBarMessage("Imprimiendo ..",, False)
                            ' ImprimirCR(x.IdReporte, Docenrty)
                            ImprimirViaCrystal(x.IdReporte, Docenrty, ImpresoraCompCORRECTA)

                        End If


                        rSboApp.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)

                    Next



                End If

            Case ""



            Case Else

        End Select






    End Sub


    Private Sub ImprimirViaCrystal(NombreCR As String, doc As Integer, Impresora As String)

        If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then

            Dim RutaCrystal = Functions.VariablesGlobales._RutaRPT & "\" & NombreCR & ".rpt"
            Dim CadenaConexion As String = Functions.VariablesGlobales._SS_SConexionHana
            Dim Proveedor As String = Functions.VariablesGlobales._SS_DriverHana
            Dim ParametroRTP As String = Functions.VariablesGlobales._SS_ParametroRPT

            ImprimirCrystalHANA(RutaCrystal, CadenaConexion, ParametroRTP, doc, Impresora, Proveedor)

        Else

            Dim RutaCrystal = Functions.VariablesGlobales._RutaRPT & "\" & NombreCR & ".rpt"
            Dim ParametroRTP As String = Functions.VariablesGlobales._SS_ParametroRPT

            ImprimirCrystalSQL(RutaCrystal, ParametroRTP, doc, Impresora)

        End If


    End Sub



    Public Function ImprimirCrystalSQL(RutaArchivo As String, Parametro As String, DocEntry As String, Impresora As String)



        If Not File.Exists(RutaArchivo) Then

            rSboApp.SetStatusBarMessage("No existe la Ruta " & RutaArchivo)

            Utilitario.Util_Log.Escribir_Log("No existe la Ruta " & RutaArchivo, "ImpresionDocumentos")

            Return False

        End If


        If String.IsNullOrWhiteSpace(Parametro) Then

            rSboApp.SetStatusBarMessage("Parametro No Configurado  o Erroneo " & Parametro)

            Utilitario.Util_Log.Escribir_Log("Parametro No Configurado  o Erroneo " & Parametro, "ImpresionDocumentos")

            Return False

        End If

        If String.IsNullOrWhiteSpace(DocEntry) Then

            rSboApp.SetStatusBarMessage("DocEntry Vacio! ")

            Utilitario.Util_Log.Escribir_Log("DocEntry Vacio " & DocEntry, "ImpresionDocumentos")

            Return False

        End If

        If String.IsNullOrWhiteSpace(Impresora) Then

            rSboApp.SetStatusBarMessage("Nombre Impresora Vacio! ")

            Utilitario.Util_Log.Escribir_Log("Nombre Impresora Vacio! " & Impresora, "ImpresionDocumentos")

            Return False

        End If

        Try
            Dim crReport As ReportDocument = Nothing

            Try

                crReport = New ReportDocument

            Catch ex As Exception
                rSboApp.SetStatusBarMessage("Error al instanciar clase ReportDocument! ")
                Utilitario.Util_Log.Escribir_Log("Error al instanciar clase ReportDocument : " + ex.Message.ToString(), "ImpresionDocumentos")
                Return False
            End Try



            Utilitario.Util_Log.Escribir_Log("Ruta RPT: " + RutaArchivo, "ImpresionDocumentos")

            Utilitario.Util_Log.Escribir_Log("Patametro: " + Parametro.ToString, "ImpresionDocumentos")

            Utilitario.Util_Log.Escribir_Log("Valor: " + DocEntry.ToString, "ImpresionDocumentos")


            Try

                crReport.Load(RutaArchivo)

            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion Load :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try



            Try


                crReport.SetParameterValue(Parametro, DocEntry)


            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion SetParameterValue :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try


            Try

                crReport.DataSourceConnections(0).SetLogon(Functions.VariablesGlobales._gUsuarioDB, Functions.VariablesGlobales._gPasswordDB)

            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion SetLogon :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try



            Try


                If String.IsNullOrEmpty(Functions.VariablesGlobales._ipServer) Then

                    crReport.DataSourceConnections(0).SetConnection(rCompany.Server, rCompany.CompanyDB, False)

                Else

                    crReport.DataSourceConnections(0).SetConnection(Functions.VariablesGlobales._ipServer, rCompany.CompanyDB, False)


                End If

            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion SetConnection :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try





            Try

                '----------Seteamos a impresora
                'crReport.PrintOptions.PrinterName = Impresora
                'crReport.PrintToPrinter(1, False, 0, 0) '//se genera y se imprime el reporte

                '---nueva logica 18072023

                Dim PrinterSettings = New System.Drawing.Printing.PrinterSettings()
                PrinterSettings.PrinterName = Impresora
                PrinterSettings.Copies = 1
                PrinterSettings.Collate = False
                crReport.PrintToPrinter(PrinterSettings, New System.Drawing.Printing.PageSettings(), False)


            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion PrintToPrinter :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try



            crReport.Close() '//se cierra el reporte  

            crReport.Dispose() '// se libera el reporte


            Return True

        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("Error General para Proceso de Impresion : " + ex.Message.ToString(), "ImpresionDocumentos")

            Return False

        End Try
    End Function

    Public Function ImprimirCrystalHANA(RutaArchivo As String, CadenaConexion As String, Parametro As String, DocEntry As String, Impresora As String, Proveedor As String)



        If Not File.Exists(RutaArchivo) Then

            rSboApp.SetStatusBarMessage("No existe la Ruta " & RutaArchivo)

            Utilitario.Util_Log.Escribir_Log("No existe la Ruta " & RutaArchivo, "ImpresionDocumentos")

            Return False

        End If

        If String.IsNullOrWhiteSpace(CadenaConexion) Then

            rSboApp.SetStatusBarMessage("Cadena de Conexion no configurada o Erroneo ")

            Utilitario.Util_Log.Escribir_Log("Cadena de Conexion no configurada o Erroneo " & CadenaConexion, "ImpresionDocumentos")

            Return False

        End If

        If String.IsNullOrWhiteSpace(Proveedor) Then

            rSboApp.SetStatusBarMessage("Proveedor de Driver no configurada o Erroneo ")

            Utilitario.Util_Log.Escribir_Log("Proveedor de Driver no configurada o Erroneo " & Proveedor, "ImpresionDocumentos")

            Return False

        End If


        If String.IsNullOrWhiteSpace(Parametro) Then

            rSboApp.SetStatusBarMessage("Parametro No Configurado  o Erroneo " & Parametro)

            Utilitario.Util_Log.Escribir_Log("Parametro No Configurado  o Erroneo " & Parametro, "ImpresionDocumentos")

            Return False

        End If

        If String.IsNullOrWhiteSpace(DocEntry) Then

            rSboApp.SetStatusBarMessage("DocEntry Vacio! ")

            Utilitario.Util_Log.Escribir_Log("DocEntry Vacio " & DocEntry, "ImpresionDocumentos")

            Return False

        End If

        If String.IsNullOrWhiteSpace(Impresora) Then

            rSboApp.SetStatusBarMessage("Nombre Impresora Vacio! ")

            Utilitario.Util_Log.Escribir_Log("Nombre Impresora Vacio! " & Impresora, "ImpresionDocumentos")

            Return False

        End If

        Try
            Dim crReport As ReportDocument = Nothing

            Try

                crReport = New ReportDocument

            Catch ex As Exception
                rSboApp.SetStatusBarMessage("Error al instanciar clase ReportDocument! ")
                Utilitario.Util_Log.Escribir_Log("Error al instanciar clase ReportDocument : " + ex.Message.ToString(), "ImpresionDocumentos")
                Return False
            End Try


            Utilitario.Util_Log.Escribir_Log("Conexion: " + CadenaConexion, "ImpresionDocumentos")

            Utilitario.Util_Log.Escribir_Log("Ruta RPT: " + RutaArchivo, "ImpresionDocumentos")

            Utilitario.Util_Log.Escribir_Log("Patametro: " + Parametro.ToString, "ImpresionDocumentos")

            Utilitario.Util_Log.Escribir_Log("Valor: " + DocEntry.ToString, "ImpresionDocumentos")

            Utilitario.Util_Log.Escribir_Log("Nombre impresora: " + Impresora.ToString, "ImpresionDocumentos")


            Try

                crReport.Load(RutaArchivo)

            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion Load :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try



            Try


                crReport.SetParameterValue(Parametro, DocEntry)


            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion SetParameterValue :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try



            Dim logonProps2 As NameValuePairs2 = crReport.DataSourceConnections(0).LogonProperties

            Try

                If (IntPtr.Size = 8) Then
                    logonProps2.Set("Provider", Proveedor)
                    logonProps2.Set("Server Type", Proveedor)
                Else
                    logonProps2.Set("Provider", Proveedor & "32")
                    logonProps2.Set("Server Type", Proveedor & "32")
                End If

                logonProps2.Set("Connection String", CadenaConexion)

                crReport.DataSourceConnections(0).SetLogonProperties(logonProps2)


            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion SetLogonProperties :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try


            Try


                If String.IsNullOrEmpty(Functions.VariablesGlobales._ipServer) Then

                    crReport.DataSourceConnections(0).SetConnection(rCompany.Server, rCompany.CompanyDB, False)

                Else

                    crReport.DataSourceConnections(0).SetConnection(Functions.VariablesGlobales._ipServer, rCompany.CompanyDB, False)


                End If

            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion SetConnection :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try





            Try
                '----------Seteamos a impresora
                crReport.PrintOptions.PrinterName = Impresora
                crReport.PrintToPrinter(1, False, 0, 0) '//se genera y se imprime el reporte

                '---nueva logica 18072023

                'Dim PrinterSettings = New System.Drawing.Printing.PrinterSettings()
                'PrinterSettings.PrinterName = Impresora
                'PrinterSettings.Copies = 1
                'PrinterSettings.Collate = False
                'crReport.PrintToPrinter(PrinterSettings, New System.Drawing.Printing.PageSettings(), False)


            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Excepcion PrintToPrinter :" + ex.Message, "ImpresionDocumentos")

                Return False

            End Try



            crReport.Close() '//se cierra el reporte  

            crReport.Dispose() '// se libera el reporte


            Return True

        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("Error General para Proceso de Impresion : " + ex.Message.ToString(), "ImpresionDocumentos")

            Return False

        End Try
    End Function

    Private Sub ImprimirCR(codigoCR As String, doc As Integer)

        Dim oReportLayoutService As ReportLayoutsService = TryCast(rCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), ReportLayoutsService)
        Dim oReportPrintParams As ReportLayoutPrintParams = TryCast(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), ReportLayoutPrintParams)

        oReportPrintParams.LayoutCode = codigoCR
        oReportPrintParams.DocEntry = doc

        Try
            oReportLayoutService.Print(oReportPrintParams)
        Catch ex As Exception
            rSboApp.SetStatusBarMessage(ex.Message)
            Utilitario.Util_Log.Escribir_Log("Error codigo layout " & codigoCR & " Docentry " & doc.ToString & " - " & ex.Message, "Impresion")
        End Try


    End Sub

    Private Sub SetearImpresora(codigoCR As String, Impresora As String)

        Dim oReportLayoutService As ReportLayoutsService = TryCast(rCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), ReportLayoutsService)
        Dim oReportLayoutParams As ReportLayoutParams = TryCast(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams), ReportLayoutParams)

        oReportLayoutParams.LayoutCode = codigoCR

        Dim datosreporte = oReportLayoutService.GetReportLayout(oReportLayoutParams)

        If datosreporte.Printer <> Impresora Then

            datosreporte.Printer = Impresora

            oReportLayoutService.UpdatePrinterSettings(datosreporte)

        End If

    End Sub


    Private Function ObtenerImpresoraPorNombreParcial(nombreprinter As String) As String

        Try
            For i As Integer = 0 To PrinterSettings.InstalledPrinters.Count - 1
                Dim nombreX = PrinterSettings.InstalledPrinters.Item(i)

                If nombreX.ToLower.Contains(nombreprinter.ToLower) Then

                    Return nombreX

                End If

            Next

        Catch ex As Exception

        End Try

        Return String.Empty

    End Function

#End Region


#End Region

    Private Sub creaBotonCargaExtractoBnacario(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Try
            Dim UbicacionBotones As String = ""

            'item = oForm.Items.Add("btnAccion", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oForm.DataSources.UserDataSources.Add("btnAccion", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

            Dim item_1 As SAPbouiCOM.Button = oForm.Items.Item("2").Specific


            Dim lef As Integer = item_1.Item.Left + item_1.Item.Width + 5
            Dim top As Integer = item_1.Item.Top
            Dim Width As Integer = item_1.Item.Width + 15
            Dim Height As Integer = item_1.Item.Height

            oFuncionesB1.creaControl(oForm, SAPbouiCOM.BoFormItemTypes.it_BUTTON, "btnCargaEB", "Cargar valores", lef, top, Width, Height, 0, False)

        Catch ex As Exception
            rSboApp.SetStatusBarMessage(NombreAddon + " - Erro al crear boton carga extracto bancario: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
        oForm.Freeze(False)

    End Sub



End Class

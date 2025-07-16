Option Strict Off
Option Explicit On

Public Class Estructura

    Public WithEvents rSboApp As SAPbouiCOM.Application

    Public VersionAddon As String = "3.8.1"
    Public LicenciaAddon As String = "" 'EMISION, RECEPCION, FULL
    Dim N_Categoria As String = "BF_SOLSAP"
    ''' <summary>
    '''  Inicializaciòn de la Clase
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try


            rSboApp = rSboGui.GetApplication
            oFuncionesB1 = New Functions.FuncionesB1(rCompany, rSboApp)
            oFuncionesB1.mostrarMensajesError = False
            oFuncionesB1.mostrarMensajesExito = True
            oFuncionesB1.mantenerLogErrores = True
            oFuncionesB1.validarVersion_SoloCrearTabla()
                


        Catch ex As Exception
        End Try
    End Sub

    Public Sub CreacionDeEstructura(Optional generalEstructuraLocalizacion As Boolean = False)
        Try
            'If Not oFuncionesB1.validarVersion(NombreAddon, VersionAddon) Then

            Nombre_Proveedor_SAP_BO = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "ProveedorSAP")
            If Nombre_Proveedor_SAP_BO = "" Then
                rSboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización de Proveedor SAP BO", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Sub
            End If
            Utilitario.Util_Log.Escribir_Log("NOMBRE PROVEEDOR SAP BO: " + Nombre_Proveedor_SAP_BO, "Estructura")



            If Functions.VariablesGlobales._XMLRecepcionHeison = "Y" Then

                rSboApp.StatusBar.SetText(NombreAddon + " - Creando estructura para leer XML recibidos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                oFuncionesB1.creaTablaMD("GS_FC", "(SS) FACTURA CABECERA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("GS_FCDET", "(SS) FACTURA DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                'oFuncionesB1.creaTablaMD("GS1_FCDET", "(SS) FACTURA DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                'oFuncionesB1.creaTablaMD("GS2_FCDETIMP", "(SS) FACTURA DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                CreaCamposFacturaXML()
                CreaUDOFacturaXML()

                oFuncionesB1.creaTablaMD("GS_NC", "(SS) NCREDITO CABECERA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("GS_NCDET", "(SS) NCREDITO DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                CreaCamposNCreditoXML()
                CreaUDONCreditoXML()

                oFuncionesB1.creaTablaMD("GS_RT", "(SS) RETENCION CABECERA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("GS_RTDET", "(SS) RETENCION DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                CreaCamposRetencionXML()
                CreaUDORetencionXML()

            End If

            rSboApp.StatusBar.SetText(NombreAddon + " - Validando la estructura necesaria para el correcto funcionamiento del Addon.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            FUN_CreaTablas()
            rSboApp.StatusBar.SetText(NombreAddon + " - TABLAS CREADAS...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            FUN_CreaCampos()
            rSboApp.StatusBar.SetText(NombreAddon + " - CAMPOS CREADOS...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '    oFuncionesB1.InsertaConfiguracion() ' TABLA DE USUARIO INSERTA RUTAS WS, ETC. LOS DATOS DE ESTA TABLA DEBEN SER MIGRADOS AL UDO DE CONFIGURACION

            FUN_CreaUDO_LOG() ' EMISION Y RECEPCION
            rSboApp.StatusBar.SetText(NombreAddon + " - UDO LOG CREADO...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'oFuncionesB1.InsertaConfiguracionUDO()

            If Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "recepcion" Or Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "full" Then
                CrearBusquedasFormateadas()
            End If



            ' Estructura Localizacion -------------add 19022024 Artur
            If generalEstructuraLocalizacion Then

                CrearEstructuraLocalizacion()

            End If



            'Fin Estructura----------------------------------------


            oFuncionesB1.confirmarVersion(NombreAddon, VersionAddon)
            rSboApp.StatusBar.SetText(NombreAddon + " - Validación de Estructura de Base Finalizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'End If
        Catch ex As Exception

        End Try
    End Sub

#Region "Estructuras Localizacion"

    Private Sub InseratInfoMesAño(ByVal tablaUsuario As String, recursoData As String, campoCodigo As String, campoDescripcion As String)

        Dim oUserTable As SAPbobsCOM.UserTable

        Try

            oUserTable = rCompany.UserTables.Item(tablaUsuario)

            Dim valorRecursoPM() As String = recursoData.Split(New Char() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim ultimoinsercion As Integer = 0, i As Integer
            Dim lErrCode As Integer = 0
            Dim sErrMsg As String = ""
            Dim parte As String = ""
            ultimoinsercion = CInt(oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1))

            If ultimoinsercion > 0 And ultimoinsercion < valorRecursoPM.Length Then

                For i = ultimoinsercion - 1 To valorRecursoPM.Length - 1

                    lErrCode = 0
                    sErrMsg = ""
                    parte = ""
                    Try

                        parte = valorRecursoPM(i)
                        oUserTable.Code = parte.Split(";")(0)
                        oUserTable.Name = parte.Split(";")(1)
                        'oUserTable.UserFields.Fields.Item(campoCodigo).Value = parte.Split(";")(0)
                        'oUserTable.UserFields.Fields.Item(campoDescripcion).Value = parte.Split(";")(1)
                        oUserTable.Add()
                        rCompany.GetLastError(lErrCode, sErrMsg)
                        If lErrCode <> 0 Then
                            rSboApp.StatusBar.SetText(NombreAddon + " - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log(" - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1) + "- " + sErrMsg.ToString(), "Estructura")
                        Else
                            Utilitario.Util_Log.Escribir_Log(NombreAddon + " -  Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1) + "- " + sErrMsg.ToString(), "Estructura")
                            rSboApp.StatusBar.SetText(NombreAddon + " - Registro en la " + tablaUsuario + " id= " & parte.Split(";")(0) & " des= " & parte.Split(";")(1), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'escribelog("GuardaDireccion , Provincia: " & Provincia.ToString() & "- Canton:" & Canton.ToString() & "- Distrito:" & Distrito.ToString() & "- Barrio:" & Barrio.ToString())
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" Error al poblar Tabla: " + tablaUsuario + " ...  " + ex.Message.ToString + " ! ", "Estructura")
                    End Try

                Next

                rSboApp.StatusBar.SetText(NombreAddon + " - " + tablaUsuario + ":Insertados Correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Else

                rSboApp.StatusBar.SetText(NombreAddon + "La informacion ya se encuentra Insertada para la Tabla: " + tablaUsuario, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log(" Error2 al poblar Tabla: " + tablaUsuario + " ...  " + ex.Message.ToString + " ! ", "Estructura")

        Finally

            oUserTable = Nothing
            System.GC.Collect()

        End Try

    End Sub


    Private Sub InseratInfoSapESQCOD(ByVal tablaUsuario As String, recursoData As String, campoCodigo As String, CodPro As String, CodCan As String, CodPar As String, NomPro As String, NomCan As String, NomPar As String)

        Dim oUserTable As SAPbobsCOM.UserTable

        Try

            oUserTable = rCompany.UserTables.Item(tablaUsuario)

            Dim valorRecursoPM() As String = recursoData.Split(New Char() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim ultimoinsercion As Integer = 0, i As Integer
            Dim lErrCode As Integer = 0
            Dim sErrMsg As String = ""
            Dim parte As String = ""

            If campoCodigo = "Code" Then
                ultimoinsercion = CInt(oFuncionesB1.getCorrelativoCount("@" + tablaUsuario)) + 1
            Else

                ultimoinsercion = CInt(oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1))
            End If



            If ultimoinsercion > 0 And ultimoinsercion < valorRecursoPM.Length Then

                For i = ultimoinsercion - 1 To valorRecursoPM.Length - 1

                    lErrCode = 0
                    sErrMsg = ""
                    parte = ""
                    Try

                        parte = valorRecursoPM(i)


                        oUserTable.Code = oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1)
                        oUserTable.Name = oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1)

                        oUserTable.UserFields.Fields.Item(CodPro).Value = parte.Split(";")(0)
                        oUserTable.UserFields.Fields.Item(CodCan).Value = parte.Split(";")(1)
                        oUserTable.UserFields.Fields.Item(CodPar).Value = parte.Split(";")(2)

                        oUserTable.UserFields.Fields.Item(NomPro).Value = parte.Split(";")(3)
                        oUserTable.UserFields.Fields.Item(NomCan).Value = parte.Split(";")(4)
                        oUserTable.UserFields.Fields.Item(NomPar).Value = parte.Split(";")(5)



                        oUserTable.Add()
                        rCompany.GetLastError(lErrCode, sErrMsg)
                        If lErrCode <> 0 Then
                            rSboApp.StatusBar.SetText(NombreAddon + " - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log(" - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1) + "- " + sErrMsg.ToString(), "Estructura")
                        Else
                            Utilitario.Util_Log.Escribir_Log(" - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1) + "- " + sErrMsg.ToString(), "Estructura")
                            rSboApp.StatusBar.SetText(NombreAddon + " - Registro en la " + tablaUsuario + " id= " & parte.Split(";")(0) & " des= " & parte.Split(";")(1), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'escribelog("GuardaDireccion , Provincia: " & Provincia.ToString() & "- Canton:" & Canton.ToString() & "- Distrito:" & Distrito.ToString() & "- Barrio:" & Barrio.ToString())
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" Error al poblar Tabla: " + tablaUsuario + " ...  " + ex.Message.ToString + " ! ", "Estructura")
                    End Try

                Next

                rSboApp.StatusBar.SetText(NombreAddon + " - " + tablaUsuario + ":Insertados Correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Else

                rSboApp.StatusBar.SetText(NombreAddon + "La informacion ya se encuentra Insertada para la Tabla: " + tablaUsuario, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log(" Error2 al poblar Tabla: " + tablaUsuario + " ...  " + ex.Message.ToString + " ! ", "Estructura")

        Finally

            oUserTable = Nothing
            System.GC.Collect()

        End Try

    End Sub

    Private Sub InseratInfoSap(ByVal tablaUsuario As String, recursoData As String, campoCodigo As String, campoDescripcion As String)

        Dim oUserTable As SAPbobsCOM.UserTable

        Try

            oUserTable = rCompany.UserTables.Item(tablaUsuario)

            Dim valorRecursoPM() As String = recursoData.Split(New Char() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim ultimoinsercion As Integer = 0, i As Integer
            Dim lErrCode As Integer = 0
            Dim sErrMsg As String = ""
            Dim parte As String = ""

            If campoCodigo = "Code" Then
                ultimoinsercion = CInt(oFuncionesB1.getCorrelativoCount("@" + tablaUsuario)) + 1
            Else

                ultimoinsercion = CInt(oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1))
            End If



            If ultimoinsercion > 0 And ultimoinsercion < valorRecursoPM.Length Then

                For i = ultimoinsercion - 1 To valorRecursoPM.Length - 1

                    lErrCode = 0
                    sErrMsg = ""
                    parte = ""
                    Try

                        parte = valorRecursoPM(i)

                        If campoCodigo = "Code" Then

                            oUserTable.Code = parte.Split(";")(0)
                            oUserTable.Name = parte.Split(";")(1)

                        Else
                            oUserTable.Code = oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1)
                            oUserTable.Name = oFuncionesB1.getCorrelativo("Code", "@" + tablaUsuario, , 1)
                            oUserTable.UserFields.Fields.Item(campoCodigo).Value = parte.Split(";")(0)
                            oUserTable.UserFields.Fields.Item(campoDescripcion).Value = parte.Split(";")(1)

                        End If


                        oUserTable.Add()
                        rCompany.GetLastError(lErrCode, sErrMsg)
                        If lErrCode <> 0 Then
                            rSboApp.StatusBar.SetText(NombreAddon + " - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log(" - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1) + "- " + sErrMsg.ToString(), "Estructura")
                        Else
                            Utilitario.Util_Log.Escribir_Log(" - No se pudo insertar el Registro en la " + tablaUsuario + " id=" & parte.Split(";")(0) & " des=" & parte.Split(";")(1) + "- " + sErrMsg.ToString(), "Estructura")
                            rSboApp.StatusBar.SetText(NombreAddon + " - Registro en la " + tablaUsuario + " id= " & parte.Split(";")(0) & " des= " & parte.Split(";")(1), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'escribelog("GuardaDireccion , Provincia: " & Provincia.ToString() & "- Canton:" & Canton.ToString() & "- Distrito:" & Distrito.ToString() & "- Barrio:" & Barrio.ToString())
                        End If
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log(" Error al poblar Tabla: " + tablaUsuario + " ...  " + ex.Message.ToString + " ! ", "Estructura")
                    End Try

                Next

                rSboApp.StatusBar.SetText(NombreAddon + " - " + tablaUsuario + ":Insertados Correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Else

                rSboApp.StatusBar.SetText(NombreAddon + "La informacion ya se encuentra Insertada para la Tabla: " + tablaUsuario, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log(" Error2 al poblar Tabla: " + tablaUsuario + " ...  " + ex.Message.ToString + " ! ", "Estructura")

        Finally

            oUserTable = Nothing
            System.GC.Collect()

        End Try

    End Sub

    Private Sub AsociarBusquedasFormateadas()
        Dim querys As String = ""
        Dim lista_form_tablas As New List(Of String)

        ' Se asigna las busquedas de paises
        oFuncionesB1.creaQueryCat(N_Categoria)

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140") 'guia
        lista_form_tablas.Add("OWTR-940") 'guia
        lista_form_tablas.Add("OINV-133") 'guia
        lista_form_tablas.Add("ORIN-179")
        lista_form_tablas.Add("OINV-65303")
        lista_form_tablas.Add("OPCH-141")

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = "Select ""U_SS_CodPais"",""U_SS_NomPais"" from ""@SS_PAIS"" "
        Else
            querys = "Select ""U_SS_CodPais"",""U_SS_NomPais"" from ""@SS_PAIS"" "
        End If
        CrearBusquedasFormateadas("SS_COD_PAISES", querys, lista_form_tablas, "SS_PaisOrigen", "U_SS_PaisOrigen")
        CrearBusquedasFormateadas("SS_COD_PAISES", querys, lista_form_tablas, "SS_PaisDestino", "U_SS_PaisDestino")
        CrearBusquedasFormateadas("SS_COD_PAISES", querys, lista_form_tablas, "SS_PaisAdqui", "U_SS_PaisAdqui")
        CrearBusquedasFormateadas("SS_COD_PAISES", querys, lista_form_tablas, "SS_PaisEfecPago", "U_SS_PaisEfecPago")


        '-----------------------------------
        lista_form_tablas.Clear()
        lista_form_tablas.Add("OCRD-134") 'guia
        CrearBusquedasFormateadas("SS_COD_PAISES", querys, lista_form_tablas, "SS_PaisEfecPago", "U_SS_PaisEfecPago")

        'para las guias de trasferencias

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OWTR-940") 'guia
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_EST_GRT_HANA
        Else
            querys = My.Resources.SS_EST_GRT_SQL
        End If
        CrearBusquedasFormateadas("SS_EST_GRT", querys, lista_form_tablas, "SS_Est", "U_SS_Est")

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PUNTOEMISION_GRT_HANA
        Else
            querys = My.Resources.SS_PUNTOEMISION_GRT_SQL
        End If
        CrearBusquedasFormateadas("SS_PUNTOEMISION_GRT", querys, lista_form_tablas, "SS_Pemi", "U_SS_Pemi")

        ' para entregas

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140") 'guia
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_EST_GR_HANA
        Else
            querys = My.Resources.SS_EST_GR_SQL
        End If
        CrearBusquedasFormateadas("SS_EST_GR", querys, lista_form_tablas, "SS_Est", "U_SS_Est")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PUNTOEMISION_GR_HANA
        Else
            querys = My.Resources.SS_PUNTOEMISION_GR_SQL
        End If
        CrearBusquedasFormateadas("SS_PUNTOEMISION", querys, lista_form_tablas, "SS_Pemi", "U_SS_Pemi")



        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_GR_ESTFAC_HANA
        Else
            querys = My.Resources.SS_GR_ESTFAC_SQL
        End If
        CrearBusquedasFormateadas("SS_GR_ESTFAC_REL", querys, lista_form_tablas, "SS_EstFacRel", "U_SS_EstFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_GR_PTOEMIFAC_HANA
        Else
            querys = My.Resources.SS_GR_PUNTOEMISION_SQL
        End If
        CrearBusquedasFormateadas("SS_GR_PUNTOEMISIONFC_REL", querys, lista_form_tablas, "SS_PunEmiFacRel", "U_SS_PunEmiFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_GR_NUMAUTFAC_HANA
        Else
            querys = My.Resources.SS_GR_NUMAUTOFC_REL_SQL
        End If
        CrearBusquedasFormateadas("SS_GR_NUMAUTOFC_REL", querys, lista_form_tablas, "SS_NumAutFacRel", "U_SS_NumAutFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_GR_FOLIOFAC_HANA
        Else
            querys = My.Resources.SS_GR_FOLIOFAC_SQL
        End If
        CrearBusquedasFormateadas("SS_GR_FOLIOFAC", querys, lista_form_tablas, "SS_NumFacRel", "U_SS_NumFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ODLN-140")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_GR_FECHAFAC_HANA
        Else
            querys = My.Resources.SS_GR_FECHAFAC_SQL
        End If
        CrearBusquedasFormateadas("SS_GR_FECHAFAC", querys, lista_form_tablas, "SS_FecEmiDocRel", "U_SS_FecEmiDocRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-133")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_EST_FV_HANA
        Else
            querys = My.Resources.SS_EST_FV_SQL
        End If
        CrearBusquedasFormateadas("SS_EST_FV", querys, lista_form_tablas, "SS_Est", "U_SS_Est")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-133")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PUNTOEMISION_FV_HANA
        Else
            querys = My.Resources.SS_PUNTOEMISION_FV_SQL
        End If
        CrearBusquedasFormateadas("SS_PUNTOEMISION_FV", querys, lista_form_tablas, "SS_Pemi", "U_SS_Pemi")

        'lista_form_tablas.Clear()
        'lista_form_tablas.Add("OINV-133")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    querys = My.Resources.SS_NUM_AUTO_FV_HANA
        'Else
        '    querys = My.Resources.SS_NUM_AUTO_FV_SQL
        'End If
        'CrearBusquedasFormateadas("SS_NUM_AUTO_FV", querys, lista_form_tablas, "SS_NumAut", "U_SS_NumAut")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_EST_NC_HANA
        Else
            querys = My.Resources.SS_EST_NC_SQL
        End If
        CrearBusquedasFormateadas("SS_EST_NC", querys, lista_form_tablas, "SS_Est", "U_SS_Est")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PUNTOEMISION_NC_HANA
        Else
            querys = My.Resources.SS_PUNTOEMISION_NC_SQL
        End If
        CrearBusquedasFormateadas("SS_PUNTOEMISION_NC", querys, lista_form_tablas, "SS_Pemi", "U_SS_Pemi")

        'lista_form_tablas.Clear()
        'lista_form_tablas.Add("ORIN-179")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    querys = My.Resources.SS_NUMAUTO_NC_HANA
        'Else
        '    querys = My.Resources.SS_NUMAUTO_NC_SQL
        'End If
        'CrearBusquedasFormateadas("SS_NUMAUTO_NC", querys, lista_form_tablas, "SS_NumAut", "U_SS_NumAut")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_NC_EST_FAC_REL_HANA
        Else
            querys = My.Resources.SS_NC_EST_FAC_REL_SQL
        End If
        CrearBusquedasFormateadas("SS_NC_EST_FAC_REL", querys, lista_form_tablas, "SS_EstFacRel", "U_SS_EstFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_NC_PUNTOEMISION_FC_REL_HANA
        Else
            querys = My.Resources.SS_NC_PUNTOEMISION_FC_REL_SQL
        End If
        CrearBusquedasFormateadas("SS_NC_PUNTOEMISION_FC_REL", querys, lista_form_tablas, "SS_PunEmiFacRel", "U_SS_PunEmiFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_NC_NUMAUTO_FAC_REL_HANA
        Else
            querys = My.Resources.SS_NC_NUMAUTO_FAC_REL_SQL
        End If
        CrearBusquedasFormateadas("SS_NC_NUMAUTO_FAC_REL", querys, lista_form_tablas, "SS_NumAutFacRel", "U_SS_NumAutFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_NC_FOLIO_FC_REL_HANA
        Else
            querys = My.Resources.SS_NC_FOLIO_FC_REL_SQL
        End If
        CrearBusquedasFormateadas("SS_NC_FOLIO_FC_REL", querys, lista_form_tablas, "SS_NumFacRel", "U_SS_NumFacRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_NC_FECHA_FC_REL_HANA
        Else
            querys = My.Resources.SS_NC_FECHA_FC_REL_SQL
        End If
        CrearBusquedasFormateadas("SS_NC_FECHA_FC_REL", querys, lista_form_tablas, "SS_FecEmiDocRel", "U_SS_FecEmiDocRel")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_TIPOCOMPROBANTE_APLICA_NC_HANA
        Else
            querys = My.Resources.SS_TIPOCOMPROBANTE_APLICA_NC_SQL
        End If
        CrearBusquedasFormateadas("SS_TIPOCOMPROBANTE_APLICA_NC", querys, lista_form_tablas, "SS_TipDocAplica", "U_SS_TipDocAplica")


        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-65303") 'NOTA DE DEBITO
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_EST_ND_HANA
        Else
            querys = My.Resources.SS_EST_ND_SQL
        End If
        CrearBusquedasFormateadas("SS_EST_ND", querys, lista_form_tablas, "SS_Est", "U_SS_Est")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-65303")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PUNTOEMISION_ND_HANA
        Else
            querys = My.Resources.SS_PUNTOEMISION_ND_SQL
        End If
        CrearBusquedasFormateadas("SS_PUNTOEMISION_ND", querys, lista_form_tablas, "SS_Pemi", "U_SS_Pemi")

        'lista_form_tablas.Clear()
        'lista_form_tablas.Add("OINV-65303")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    querys = My.Resources.SS_NUMAUTO_ND_HANA
        'Else
        '    querys = My.Resources.SS_NUMAUTO_ND_SQL
        'End If
        'CrearBusquedasFormateadas("SS_NUMAUTO_ND", querys, lista_form_tablas, "SS_NumAut", "U_SS_NumAut")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OPCH-141") 'RETENCION
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_SERIE_RETENCION_HANA
        Else
            querys = My.Resources.SS_SERIE_RETENCION_SQL
        End If
        CrearBusquedasFormateadas("SS_SERIE_RETENCION", querys, lista_form_tablas, "SS_SerieRet", "U_SS_SerieRet")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OPCH-141")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_FECHA_RT_HANA
        Else
            querys = My.Resources.SS_FECHA_RT_SQL
        End If
        CrearBusquedasFormateadas("SS_FECHA_RT", querys, lista_form_tablas, "SS_FecRet", "U_SS_FecRet")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OPCH-141")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_EST_LQE_HANA
        Else
            querys = My.Resources.SS_EST_LQE_SQL
        End If
        CrearBusquedasFormateadas("SS_EST_LQE", querys, lista_form_tablas, "SS_Est", "U_SS_Est")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OPCH-141")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PUNTOEMISION_LQE_HANA
        Else
            querys = My.Resources.SS_PUNTOEMISION_LQE_SQL
        End If
        CrearBusquedasFormateadas("SS_PUNTOEMISION_LQE", querys, lista_form_tablas, "SS_Pemi", "U_SS_Pemi")

        'lista_form_tablas.Clear()
        'lista_form_tablas.Add("OPCH-141")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    querys = My.Resources.SS_NUMAUTO_LQE_HANA
        'Else
        '    querys = My.Resources.SS_NUMAUTO_LQE_SQL
        'End If
        'CrearBusquedasFormateadas("SS_NUMAUTO_LQE", querys, lista_form_tablas, "SS_NumAut", "U_SS_NumAut")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-133")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_FORMAPAGO_MAX_VENTAS_HANA
        Else
            querys = My.Resources.SS_FORMAPAGO_MAX_VENTAS_SQL
        End If
        CrearBusquedasFormateadas("SS_FORMAPAGO_MAX_VENTAS", querys, lista_form_tablas, "SS_ForPagVen", "U_SS_ForPagVen")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OPCH-141")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_FORMAPAGO_MAX_COMPRAS_HANA
        Else
            querys = My.Resources.SS_FORMAPAGO_MAX_COMRPAS_SQL
        End If
        CrearBusquedasFormateadas("SS_FORMAPAGO_MAX_COMRPAS", querys, lista_form_tablas, "SS_ForPagCompras", "U_SS_ForPagCompras")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-133")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_TIPO_COMPROBANTE_FV_HANA
        Else
            querys = My.Resources.SS_TIPO_COMPROBANTE_FV_SQL
        End If
        CrearBusquedasFormateadas("SS_TIPO_COMPROBANTE_FV", querys, lista_form_tablas, "SS_TipCom", "U_SS_TipCom")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-133")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_FORMASPAGO_HANA
        Else
            querys = My.Resources.SS_FORMASPAGO_SQL
        End If
        CrearBusquedasFormateadas("SS_SS_FORMAPAGO", querys, lista_form_tablas, "SS_FormaPago", "U_SS_FormaPago")


        lista_form_tablas.Clear()
        lista_form_tablas.Add("ORIN-179")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_TIPO_COMPROBANTE_NC_HANA
        Else
            querys = My.Resources.SS_TIPO_COMPROBANTE_NC_SQL
        End If
        CrearBusquedasFormateadas("SS_TIPO_COMPROBANTE_NC", querys, lista_form_tablas, "SS_TipCom", "U_SS_TipCom")

        'lista_form_tablas.Clear()
        'lista_form_tablas.Add("OPCH-141")
        'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '    querys = My.Resources.SS_NUMAUTO_RT_HANA
        'Else
        '    querys = My.Resources.SS_NUM_AUTO_RT_SQL
        'End If
        'CrearBusquedasFormateadas("SS_NUMAUTO_RT", querys, lista_form_tablas, "SS_NumAutRet", "U_SS_NumAutRet")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OPCH-141")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_TIPO_COMPROBANTES_SRI_COMPRAS_HANA
        Else
            querys = My.Resources.SS_TIPO_COMPROBANTES_SRI_COMPRAS_SQL
        End If
        CrearBusquedasFormateadas("SS_TIPO_COMPROBANTES_SRI_COMPRAS", querys, lista_form_tablas, "SS_TipCom", "U_SS_TipCom")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OINV-133")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_REEMBOLSO_MAX_HANA
        Else
            querys = My.Resources.SS_REEMBOLSO_MAX_SQL
        End If
        CrearBusquedasFormateadas("SS_REEMBOLSO_MAX", querys, lista_form_tablas, "SS_Reembolsos", "U_SS_Reembolsos")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OCRD-134")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PROVINCIAS_HANA
        Else
            querys = My.Resources.SS_PROVINCIAS_SQL
        End If
        CrearBusquedasFormateadas("SS_PROVINCIA", querys, lista_form_tablas, "SS_Provincia", "U_SS_Provincia")


        lista_form_tablas.Clear()
        lista_form_tablas.Add("OCRD-134")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_CANTON_HANA
        Else
            querys = My.Resources.SS_CANTON_SQL
        End If
        CrearBusquedasFormateadas("SS_CANTON", querys, lista_form_tablas, "SS_Canton", "U_SS_Canton")

        lista_form_tablas.Clear()
        lista_form_tablas.Add("OCRD-134")
        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            querys = My.Resources.SS_PARROQUIA_HANA
        Else
            querys = My.Resources.SS_PARROQUIA_SQL
        End If
        CrearBusquedasFormateadas("SS_PARROQUIA", querys, lista_form_tablas, "SS_Parroquia", "U_SS_Parroquia")
        ' busqueda para paises campos



    End Sub
    Private Sub InsertarProvinciaCantonParroquia()

        InseratInfoSapESQCOD("SS_ESQCOD", My.Resources.SS_PROVINCIA_CANTON_PARROQUIA.ToString, "Code", "U_SS_CodPro", "U_SS_CodCan", "U_SS_CodPar", "U_SS_NomPro", "U_SS_NomCan", "U_SS_NomPar")

    End Sub

    Private Sub InsertaDatosSustentoTributario()
        ' validar con tabla de usuario
        InseratInfoSap("SS_SUSTRI", My.Resources.SS_SUSTENTOTRIBUTARIO.ToString, "Code", "Name")

    End Sub

    Private Sub InsertaDatosTiposComprobantes()

        'InseratInfoTiposComprobantes("SS_TIPCOMAUT", My.Resources.SS_TIPOS_COMPROBANTES, "Code", "Name", "U_SS_Estado", "U_SS_TipoTra", "U_SS_Orden")

        InseratInfoSap("SS_TIPCOMAUT", My.Resources.SS_TIPOS_COMPROBANTES.ToString, "Code", "Name")

    End Sub
    Private Sub InsertarTiposIngresosExterior()

        InseratInfoSap("SS_TIPO_ING_EXT", My.Resources.SS_TIPOS_INGRESOS_EXTERIOR.ToString, "Code", "Name")

    End Sub

    Private Sub InsertaDatosTiposTransaccion()

        InseratInfoSap("SS_TIPOTRANS", My.Resources.SS_TIPOTRANSACCION.ToString, "Code", "Name")

    End Sub

    Private Sub InsertaDatosTiposIdentificacion()

        InseratInfoSap("SS_TIPOIDENT", My.Resources.SS_TIPOSIDENTIFICACION.ToString, "Code", "Name")

    End Sub
    Private Sub InsertarProvincias()

        InseratInfoMesAño("SS_PROVINCIA", My.Resources.SS_PROVINCIAS.ToString, "Code", "Name")

    End Sub
    Private Sub InsertarCiudades()

        InseratInfoSap("SS_CIUDAD", My.Resources.SS_CIUDADES.ToString, "Code", "Name")

    End Sub
    Private Sub InsertaPaises()

        InseratInfoSap("SS_PAIS", My.Resources.SS_PAISES, "U_SS_CodPais", "U_SS_NomPais")

    End Sub
    Private Sub InsertarCodigoRegimen()

        InseratInfoSap("SS_COD_REGIMEN", My.Resources.SS_CODIGOS_REGIMEN.ToString, "Code", "Name")

    End Sub

    Private Sub InsertarDistritoAduanero()

        InseratInfoSap("SS_DISTRITO_ADU", My.Resources.SS_DISTRITO_ADUANERO.ToString, "Code", "Name")

    End Sub
    Private Sub InsertarMotivoNotaCredito()

        InseratInfoSap("SS_MOTIVOS_NC", My.Resources.SS_MOTIVOS_NC.ToString, "Code", "Name")

    End Sub
    Private Sub InsertarAnio()

        InseratInfoSap("SS_ANIO", My.Resources.SS_ANIO.ToString, "Code", "Name")

    End Sub
    Private Sub InsertarMes()

        InseratInfoSap("SS_MES", My.Resources.SS_MES.ToString, "Code", "Name")

    End Sub
    Private Sub InsertaFormasdePagos()


        InseratInfoSap("SS_FORMASDEPAGO", My.Resources.FORMASDEPAGO.ToString, "Code", "Name") 'SS_FORMAS_DE_PAGOS




    End Sub

    Public Sub CrearEstructuraGuiaDesatendida()

        Try

            'guias de remision desatendida

            Try

                oFuncionesB1.creaTablaMD("SS_GRCAB", "(SS) SS GR CABECERA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("SS_GRDET", "(SS) SS GR Contenido", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("SS_GRDET1", "(SS) SS GR Info Adicional", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Catch ex As Exception

            End Try

            'Guias Desatendidas
            'cabecera
            oFuncionesB1.creaCampoMD("SS_GRCAB", "CardCode", "Codigo Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "CardName", "Nombre Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Est", "Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Pemi", "Punto de Emisión", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Sec", "Secuencial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_NumAut", "Número de Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_TipCom", "Tipo de Comprobante", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            ' oFuncionesB1.creaCampoMD("OINV", "SS_FechaEmb", "(SS) Fecha Embarque", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_FecIniTra", "Fecha Inicio Traslado", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_FecFinTra", "Fecha Fin Traslado", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_PunPart", "Punto de Partida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_CodTra", "Código de Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Transportista", "Transportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdValMT() As String = {"V", "T", "C"}
            Dim StrcbdDesMT() As String = {"VENTA", "TRANSFERENCIA", "CONSIGNACIÓN"}

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_MotTraslado", "(SS) Motivo Traslado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValMT, StrcbdDesMT, "V")


            oFuncionesB1.creaCampoMD("SS_GRCAB", "CLAVE_ACCESO", "Clave de Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "FECHA_AUT_FACT", "Fecha.Autorización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdVal() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "11"}
            Dim StrcbdDes() As String = {"NO ENVIADO", "EN PROCESO", "AUTORIZADA", "NO AUTORIZADA", "VALIDAR DATOS", "EN PROCESO SRI", "DEVUELTA", "ERROR EN RECEPCION", "ANULADO"}

            oFuncionesB1.creaCampoMD("SS_GRCAB", "ESTADO_AUTORIZACIO", "(SSE) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "0")

            oFuncionesB1.creaCampoMD("SS_GRCAB", "NUM_AUTO_FAC", "Número.Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "OBSERVACION_FACT", "(SSE) Observación", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "UrlQR", "Url.QR", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_IMAG", "(SS) Imagen", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

            'Detalle

            oFuncionesB1.creaCampoMD("SS_GRDET", "Codigo", "Codigo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Cantidad", "Cantidad", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 16, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Descripcion", "Descripcion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Adicional1", "Adicional 1", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Adicional2", "Adicional 2", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Adicional3", "Adicional 3", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'Ifo Adicional
            oFuncionesB1.creaCampoMD("SS_GRDET1", "Clave", "Clave", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET1", "Valor", "Valor", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            FUN_CreaUDO_GUIAS_DESATENDIDAS()

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("CreaarEstructuraGuiaDesatendida , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub
    Private Sub FUN_CreaCampos_a_TablasDeUsuario_LOCALIZACION()
        Try

            '--- udos

            '' CONFIGURACION ADDON
            oFuncionesB1.creaCampoMD("SS_CONF", "Modulo", "(SS) Modulo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CONF", "Tipo", "(SS) Tipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CONF", "Subtipo", "(SS) Subtipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CONFD", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CONFD", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            ' Se agrega campos de Blokeo en la serie
            Dim StrcbdValblok() As String = {"Y", "N"}
            Dim StrcbdDesblok() As String = {"Y", "N"}



            ' Se agrega los campos para el UDO
            Dim StrcbdVal8() As String = {"FV", "ND", "NC", "GR", "GRT", "RT", "LQ", "LQRT", "GRST"}
            Dim StrcbdDes8() As String = {"Factura", "Nota de Debito", "Nota de Crédito", "Guía Remision", "Guía Remision Transferencia", "Retención", "Liquidación en Compras", "Liquidación + Retención", "Guía Remision Solicitud de Traslado"}

            'electronico
            oFuncionesB1.creaCampoMD("SS_SERD", "TipoD", "(SS) Tipo Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal8, StrcbdDes8, "FV")
            oFuncionesB1.creaCampoMD("SS_SERD", "SerId", "(SS) SerieID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "SerN", "(SS) SerieName", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "Establec", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "PuntoEmi", "(SS) Punto Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "SecInicio", "(SS) Secuencial Inicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "UltimoSec", "(SS) Ultimo Secuencial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "Dire", "(SS) Direccion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERD", "Bloqueado", "(SS) Bloqueado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValblok, StrcbdDesblok, "N")

            'Preimpreso
            oFuncionesB1.creaCampoMD("SS_SERDP", "TipoD", "(SS) Tipo Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal8, StrcbdDes8, "FV")
            oFuncionesB1.creaCampoMD("SS_SERDP", "SerId", "(SS) SerieID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "SerN", "(SS) SerieName", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_SERDP", "NuInicial", "(SS) Número Inicial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "NuFinal", "(SS) Número Final", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "NumAut", "(SS) Número Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "FiniD", "(SS) F.Inicio", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "FfinD", "(SS) F.Caducidad", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "Dire", "(SS) Direccion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SERDP", "Bloqueado", "(SS) Bloqueado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValblok, StrcbdDesblok, "N")

            'Se eliminara el UDO de Documentos Legles

            '*** SS_DOCLEGALES



            '*** END SS_FORMAS_DE_PAGO
            oFuncionesB1.creaCampoMD("SS_PAIS", "SS_CodPais", "(SS) Codigo Pais", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAIS", "SS_NomPais", "(SS) Nombre Pais", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 70, SAPbobsCOM.BoYesNoEnum.tNO)



            '************REEMBOLSO
            oFuncionesB1.creaCampoMD("SS_REEMCAB", "SS_CodProv", "(SS) Codigo Proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMCAB", "SS_Estab", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMCAB", "SS_PtoEmi", "(SS) Punto Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMCAB", "SS_NumDoc", "(SS) Numero de Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 220, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdValTIDReem() As String = {"C", "P", "R", "F"}
            Dim StrcbdDesTIDReem() As String = {"CEDULA", "PASAPORTE", "RUC", "CONSUMIDOR FINAL"}

            Dim StrcbdValTCA() As String = {"CP", "VT", "EX", "TC", "RF", "FF"}
            Dim StrcbdDesTCA() As String = {"COMPRA", "VENTA", "EXPORTACION", "TARJETA DE CREDITO", "RENDIMIENTO FINANCIERO", "FONDOS Y FIDEICOMISOS"}


            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_TipoId", "(SS) Tipo ID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTIDReem, StrcbdDesTIDReem, "C")
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_IdProv", "(SS) Id Proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)

            ' se enlazara con la tabla de Tipos de Comprobantes
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_TipoComp", "(SS) Tipo Comprobante", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 40, SAPbobsCOM.BoYesNoEnum.tNO,,,, "SS_TIPCOMAUT")

            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_Est", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_PtoEmi", "(SS) Punto Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_NumDoc", "(SS) Numero Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_NumAut", "(SS) Numero Autorizacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_FecEmi", "(SS) Facha Emision", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_IVA0", "(SS) Base 0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_IvaDif0", "(SS) Base 12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_NoObjIVA", "(SS) Base No objeto IVA", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_MontoIVA", "(SS) Monto IVA", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_MontoICE", "(SS) Monto ICE", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_REEMDET", "SS_IvaExe", "(SS) Base Exenta", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 30, SAPbobsCOM.BoYesNoEnum.tNO)



            '***********TRANSPORTISTA
            oFuncionesB1.creaCampoMD("SS_TRANSPORTISTA", "SS_Transportista", "(SS) ID Transportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdValTRANE() As String = {"A", "I"}
            Dim StrcbdDesTRANE() As String = {"ACTIVO", "INACTIVO"}
            oFuncionesB1.creaCampoMD("SS_TRANSPORTISTA", "SS_Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTRANE, StrcbdDesTRANE, "A")
            Dim StrcbdValTRANTT() As String = {"I", "E"}
            Dim StrcbdDesTRANTT() As String = {"INTERNO", "EXTERNO"}
            oFuncionesB1.creaCampoMD("SS_TRANSPORTISTA", "SS_TipoTrans", "(SS) Tipo Transportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTRANTT, StrcbdDesTRANTT, "I")
            Dim StrcbdValTRANTI() As String = {"04", "05", "06", "07", "08", "09"}
            Dim StrcbdDesTRANTI() As String = {"RUC", "CÉDULA", "PASAPORTE", "VENTA A CONSUMIDOR FINAL", "IDENTIFICACION DEL EXTERIOR", "PLACA"}
            oFuncionesB1.creaCampoMD("SS_TRANSPORTISTA", "SS_TipoId", "(SS) Tipo ID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTRANTI, StrcbdDesTRANTI, "05")
            Dim StrcbdValTRANAR() As String = {"SI", "NO"}
            Dim StrcbdDesTRANAR() As String = {"SI", "NO"}
            oFuncionesB1.creaCampoMD("SS_TRANSPORTISTA", "SS_AfectoRise", "(SS) Afecto a Rise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTRANAR, StrcbdDesTRANAR, "NO")

            '************

            '***********TRANSPORTE
            oFuncionesB1.creaCampoMD("SS_TRANSPORTE", "SS_Placa", "(SS) Placa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_TRANSPORTE", "SS_TipoTransporte", "(SS) Tipo Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTRANTT, StrcbdDesTRANTT, " ")
            oFuncionesB1.creaCampoMD("SS_TRANSPORTE", "SS_Capacidad", "(SS) Capacidad", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Measurement, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_TRANSPORTE", "SS_Refrigerado", "(SS) Refrigerado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTRANAR, StrcbdDesTRANAR, " ")
            '************

            ''*************SUSTENTO TRIBUTARIO
            'oFuncionesB1.creaCampoMD("SS_SUSTRI", "SS_FechaIni", "(SS) Fecha Inicio", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SS_SUSTRI", "SS_FechaFin", "(SS) Fecha Fin", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            ''***************************************
            oFuncionesB1.creaCampoMD("SS_ESQCOD", "SS_CodPro", "(SS) Codigo Provincia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_ESQCOD", "SS_CodCan", "(SS) Codigo Canton", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_ESQCOD", "SS_CodPar", "(SS) Codigo Parroquia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_ESQCOD", "SS_NomPro", "(SS) Nombre Provincia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 65, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_ESQCOD", "SS_NomCan", "(SS) Nombre Canton", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 65, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_ESQCOD", "SS_NomPar", "(SS) Nombre Parroquia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 65, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_DINARDAP_TL", "SS_NumDoc", "(SS) Numero Documento Sap", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_DINARDAP_TL", "SS_VALDEM", "(SS) Valor Demanada Judicial", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 65, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_DINARDAP_TL", "SS_CARCAS", "(SS) Cartera Castigada", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 65, SAPbobsCOM.BoYesNoEnum.tNO)


            'Guias Desatendidas
            'cabecera
            oFuncionesB1.creaCampoMD("SS_GRCAB", "CardCode", "Codigo Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "CardName", "Nombre Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Est", "Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Pemi", "Punto de Emisión", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_NumAut", "Número de Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_TipCom", "Tipo de Comprobante", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            ' oFuncionesB1.creaCampoMD("OINV", "SS_FechaEmb", "(SS) Fecha Embarque", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_FecIniTra", "Fecha Inicio Traslado", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_FecFinTra", "Fecha Fin Traslado", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_PunPart", "Punto de Partida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_CodTra", "Código de Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Transportista", "Transportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdValMT() As String = {"V", "T", "C"}
            Dim StrcbdDesMT() As String = {"VENTA", "TRANSFERENCIA", "CONSIGNACIÓN"}

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_MotTraslado", "(SS) Motivo Traslado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValMT, StrcbdDesMT, "V")


            oFuncionesB1.creaCampoMD("SS_GRCAB", "CLAVE_ACCESO", "Clave de Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "FECHA_AUT_FACT", "Fecha.Autorización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdVal() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "11"}
            Dim StrcbdDes() As String = {"NO ENVIADO", "EN PROCESO", "AUTORIZADA", "NO AUTORIZADA", "VALIDAR DATOS", "EN PROCESO SRI", "DEVUELTA", "ERROR EN RECEPCION", "ANULADO"}

            oFuncionesB1.creaCampoMD("SS_GRCAB", "ESTADO_AUTORIZACIO", "(SSE) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "0")

            oFuncionesB1.creaCampoMD("SS_GRCAB", "NUM_AUTO_FAC", "Número.Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "OBSERVACION_FACT", "(SSE) Observación", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "UrlQR", "Url.QR", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Sec", "Secuencial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_Destino", "Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_GRCAB", "SS_IMAG", "(SS) Imagen", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)


            'Detalle

            oFuncionesB1.creaCampoMD("SS_GRDET", "Codigo", "Codigo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Cantidad", "Cantidad", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 16, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Descripcion", "Descripcion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 240, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Adicional1", "Adicional 1", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Adicional2", "Adicional 2", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET", "Adicional3", "Adicional 3", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'Ifo Adicional
            oFuncionesB1.creaCampoMD("SS_GRDET1", "Clave", "Clave", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_GRDET1", "Valor", "Valor", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas_Catch , Error: " & ex.Message.ToString(), "Estructura")

        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub FUN_CreaCampos_LOCALIZACION()
        Try

            'Se quitan lo del socio de NEgocios al Validar la Estructura

            Dim StrcbdValTID() As String = {"C", "P", "R", "F"}
            Dim StrcbdDesTID() As String = {"CEDULA", "PASAPORTE", "RUC", "CONSUMIDOR FINAL"}
            oFuncionesB1.creaCampoMD("OCRD", "SS_TipoId", "(SS) Tipo Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTID, StrcbdDesTID, "R")

            Dim StrcbdValTSN() As String = {"01", "02"}
            Dim StrcbdDesTSN() As String = {"PERSONA NATURAL", "SOCIEDAD"}
            oFuncionesB1.creaCampoMD("OCRD", "SS_TipoSN", "(SS) Tipo S.N", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTSN, StrcbdDesTSN, "02")



            Dim StrcbdValTC() As String = {"01", "02", "03", "04", "05", "06", "07", "09", "10", "99"}
            Dim StrcbdDesTC() As String = {"CONTRIBUYENTE ESPECIAL", "SECTOR PUBLICO", "OTRAS SOCIEDADES", "PERSONAS NATURALES OBLIGADAS A LLEVAR CONTABILIDAD", "PERSONAS NATURALES NO OBLIGADAS A LLEVAR CONTABILIDAD FACTURA", "PERSONAS NATURALES NO OBLIGADAS A LLEVAR CONTABILIDAD LIQ COMPRAS", "PERSONAS NATURALES NO OBLIGADAS A LLEVAR CONTABILIDAD PROFESIONALES", "RIMPE-NEGOCIOS POPULARES", "RIMPE-EMPRENDEDORES", "EXTERIOR"}


            oFuncionesB1.creaCampoMD("OCRD", "SS_TipoCon", "(SS) Tipo Contribuyente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTC, StrcbdDesTC, "04")


            Dim StrcbdValSex() As String = {"NA", "M", "F"}
            Dim StrcbdDesSex() As String = {"Ninguno", "Masculino", "Femenino"}
            oFuncionesB1.creaCampoMD("OCRD", "SS_Sexo", "(SS) Sexo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValSex, StrcbdDesSex, "NA")

            Dim StrcbdValEC() As String = {"NA", "S", "C", "D", "U", "V"}
            Dim StrcbdDesEC() As String = {"Ninguno", "Soltero", "Casado", "Divorciado", "Unión Libre", "Viudo"}
            oFuncionesB1.creaCampoMD("OCRD", "SS_EstCivil", "(SS) Estado Civil", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValEC, StrcbdDesEC, "NA")

            Dim StrcbdValOI() As String = {"NA", "B", "V", "I", "A", "R", "H", "M"}
            Dim StrcbdDesOI() As String = {"Ninguno", "Empleado Público", "Empleado Privado", "Independiente", "Ama de casa o estudiante", "Rentista", "Jubilado", "Remesas del Exterior"}
            oFuncionesB1.creaCampoMD("OCRD", "SS_OrigIng", "(SS) Origen de Ingresos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValOI, StrcbdDesOI, "NA")


            oFuncionesB1.creaCampoMD("OCRD", "SS_Provincia", "(SS) Provincia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 65, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OCRD", "SS_Canton", "(SS) Cantón", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 65, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OCRD", "SS_Parroquia", "(SS) Parroquia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 65, SAPbobsCOM.BoYesNoEnum.tNO)


            '*******************GUIAS DE REMISION*********************

            'LOC F,NC,ND,FC,ENTREGAS,TRASF
            Dim StrcbdValDEC() As String = {"SI", "NO"}
            Dim StrcbdDesDEC() As String = {"SI", "NO"}
            '  oFuncionesB1.creaCampoMD("OINV", "SS_Declarable", "(SS) Declarable", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValDEC, StrcbdDesDEC, "SI")

            ' este se llena con udo y cambiara a tabla validar insercion de datos
            'LOC FC

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_SUSTRIB", "(SS) Sustento Tributario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_SUSTRI")

            Catch ex As Exception

                oFuncionesB1.creaCampoMD("OINV", "SS_SUSTRIB", "(SS) Sustento Tributario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, , , , )

            End Try


            'LOC F,NC,ND,FC,ENTREGAS,TRASF
            oFuncionesB1.creaCampoMD("OINV", "SS_Est", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC F,NC,ND,FC,ENTREGAS,TRASF
            oFuncionesB1.creaCampoMD("OINV", "SS_Pemi", "(SS) Punto de Emisión", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC F,NC,ND,FC,ENTREGAS,TRASF
            oFuncionesB1.creaCampoMD("OINV", "SS_NumAut", "(SS) Número de Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            'validar que sea asignada la bf a una tabla y no al udo (crear la tabla TIPCOMPROBANTE)
            'LOC F,NC,ND,FC,ENTREGAS,TRASF
            oFuncionesB1.creaCampoMD("OINV", "SS_TipCom", "(SS) Tipo de Comprobante", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC NC,ND
            oFuncionesB1.creaCampoMD("OINV", "SS_EstFacRel", "(SS) Estab.Fact.Relac", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC NC,ND
            oFuncionesB1.creaCampoMD("OINV", "SS_PunEmiFacRel", "(SS) Punto Emisión.Fact.Relac", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC NC,ND
            oFuncionesB1.creaCampoMD("OINV", "SS_NumAutFacRel", "(SS) Num.Autoriz.Fact.Relac", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC NC,ND
            oFuncionesB1.creaCampoMD("OINV", "SS_NumFacRel", "(SS) Número Fact.Relac", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC NC,ND
            oFuncionesB1.creaCampoMD("OINV", "SS_FecEmiDocRel", "(SS) Fecha Emisión Doc.Vtas", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC NC,ND
            'RELACIONADA A TABLA TIPCOMPROBANTE
            oFuncionesB1.creaCampoMD("OINV", "SS_TipDocAplica", "(SS) Tipo Doc Aplica", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC F,ND,FC
            ' Revisar la insercion de los datos a la  tabla de formas de Pago.

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_FormaPagos", "(SS) Forma de Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_FORMASDEPAGO") ' "SS_FORMAS_DE_PAGOS")

            Catch ex As Exception

                oFuncionesB1.creaCampoMD("OINV", "SS_FormaPagos", "(SS) Forma de Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , )
            End Try


            'oFuncionesB1.creaCampoMD("OINV", "SS_FormaPago", "(SS) Forma de Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC NC
            'Lo mismo que en LA forma de Pago  se creo tabla Pero no se agregaron los datos

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_MOTIVO_NC", "(SS) Motivo Nota de Credito", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_MOTIVOS_NC")

            Catch ex As Exception
                oFuncionesB1.creaCampoMD("OINV", "SS_MOTIVO_NC", "(SS) Motivo Nota de Credito", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , )

            End Try


            'LOC F
            ' tiene que tener una busqueda formateada apuntando a las entregas cod 15
            oFuncionesB1.creaCampoMD("OINV", "SS_NumGuia", "(SS) Numero de Guia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC 'ENTREGA,TRANS

            oFuncionesB1.creaCampoMD("OINV", "SS_FecIniTra", "(SS) Fecha Inicio Traslado", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC 'ENTREGA,TRANS

            oFuncionesB1.creaCampoMD("OINV", "SS_FecFinTra", "(SS) Fecha Fin Traslado", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC 'ENTREGA,TRANS
            ' formato debe ser hora
            oFuncionesB1.creaCampoMD("OINV", "SS_HoraSal", "(SS) Hora de Salida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC 'ENTREGA,TRANS
            ' formato debe ser hora
            oFuncionesB1.creaCampoMD("OINV", "SS_HoraLLeg", "(SS) Hora de LLegada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC 'ENTREGA,TRANS
            oFuncionesB1.creaCampoMD("OINV", "SS_PunPart", "(SS) Punto de Partida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC 'ENTREGA,TRANS

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_CodTra", "(SS)  Código de Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_TRANSPORTE")

            Catch ex As Exception

                oFuncionesB1.creaCampoMD("OINV", "SS_CodTra", "(SS)  Código de Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, , , )


            End Try


            'LOC 'ENTREGA,TRANS

            Try

                oFuncionesB1.creaCampoMD("OINV", "SS_Transportista", "(SS) Transportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_TRANSPORTISTA")


            Catch ex As Exception

                oFuncionesB1.creaCampoMD("OINV", "SS_Transportista", "(SS) Transportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO, , , , )


            End Try

            'LOC 'ENTREGA,TRANS
            Dim StrcbdValMT() As String = {"V", "T", "C"}
            Dim StrcbdDesMT() As String = {"VENTA", "TRANSFERENCIA", "CONSIGNACIÓN"}
            oFuncionesB1.creaCampoMD("OINV", "SS_MotTraslado", "(SS) Motivo Traslado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValMT, StrcbdDesMT, "V")

            '****************************************************************

            '*****************FACTURA VENTA***************************

            'ATS
            oFuncionesB1.creaCampoMD("OINV", "SS_DenoFiscal", "(SS) Denominación Rg. Fiscal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("OINV", "SS_ParFiscal", "(SS) Paraíso Fiscal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            '************EXPORTACION**********************
            Dim StrcbdValPREL() As String = {"SI", "NO"}
            Dim StrcbdDesPREL() As String = {"SI", "NO"}
            'oFuncionesB1.creaCampoMD("OINV", "SS_ParteRel", "(SS) Parte Relacionada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPREL, StrcbdDesPREL, "NO")

            'Dim StrcbdValTREG() As String = {"01", "02", "03"}
            'Dim StrcbdDesTREG() As String = {"Régimen general", "Paraíso fiscal", "Régimen fiscal preferente o jurisdicción de menor imposición"}
            'oFuncionesB1.creaCampoMD("OINV", "SS_TipoRegFis", "(SS) Tipo de Regimen Fiscal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 75, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTREG, StrcbdDesTREG, "")

            Dim StrcbdValTE() As String = {"01", "02", "03"}
            Dim StrcbdDesTE() As String = {"Con Refrendo", "Sin Refrendo", "Exportaciones de Servicios"}
            oFuncionesB1.creaCampoMD("OINV", "SS_TipoExpor", "(SS) Tipo Exportación", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTE, StrcbdDesTE, "")

            'LOC F EX  Cuando tipo Comp es 01
            oFuncionesB1.creaCampoMD("OINV", "SS_ComercioExt", "(SS) Comercio Exterior", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_IncoTermFac", "(SS) Inicio Neg Exp", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_LugIncoTerm", "(SS) Lugar Neg Exp", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            ' busqueda a tabla apuntando a la tabla paises
            oFuncionesB1.creaCampoMD("OINV", "SS_PaisOrigen", "(SS) País Origen", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_PuertoEmb", "(SS) Puerto Embarque", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_PuertoDestino", "(SS) Puerto Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            ' busqueda a tabla apuntando a la tabla paises
            oFuncionesB1.creaCampoMD("OINV", "SS_PaisDestino", "(SS) Pais Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            ' busqueda a tabla apuntando a la tabla paises
            oFuncionesB1.creaCampoMD("OINV", "SS_PaisAdqui", "(SS) País Adquisición", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_Incotermto", "(SS) Term Total Sin Impuestos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            'Revisar el motivo de no llenarse tabla 

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_TipIngExt", "(SS) Tipo de Ingresos Ext", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_TIPO_ING_EXT")
            Catch ex As Exception
                oFuncionesB1.creaCampoMD("OINV", "SS_TipIngExt", "(SS) Tipo de Ingresos Ext", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , )
            End Try

            'oFuncionesB1.creaCampoMD("OINV", "SS_TipIngExt", "(SS) Tipo de Ingresos Ext", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("OINV", "SS_IngFueGra", "(SS) Ingreso Ext. Fue Grav. IR", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_ValFob", "(SS) Valor FOB Aduana", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("OINV", "SS_RefrendoAnio", "(SS) Refrendo Año", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("OINV", "SS_RefrendoReg", "(SS) Refrendo Regímen", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            'oFuncionesB1.creaCampoMD("OINV", "SS_NumDocTra", "(SS) No. Documento Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            'LOC F EX
            oFuncionesB1.creaCampoMD("OINV", "SS_FechaEmb", "(SS) Fecha Embarque", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            'LOC F EX
            'revisar llenado de tablas
            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_DisAduanero", "(SS) Distrito Aduanero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_DISTRITO_ADU")
            Catch ex As Exception
                oFuncionesB1.creaCampoMD("OINV", "SS_DisAduanero", "(SS) Distrito Aduanero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, , , ,)
            End Try


            'revisar llenado de tablas

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_CodRegimen", "(SS) Codigo Regimen", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, , , , "SS_COD_REGIMEN")
            Catch ex As Exception
                oFuncionesB1.creaCampoMD("OINV", "SS_CodRegimen", "(SS) Codigo Regimen", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, , , ,)
            End Try


            'campo ats
            oFuncionesB1.creaCampoMD("OINV", "SS_Correlativo", "(SS) Correlativo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8, SAPbobsCOM.BoYesNoEnum.tNO)
            'ats
            oFuncionesB1.creaCampoMD("OINV", "SS_Doc_Transporte", "(SS) Num.Doc Transporte", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 13, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("OINV", "SS_ValIrExt", "(SS) Valor del IR del Exterior", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            '*************************************************************************

            '****************************RETENCIÓN*********************************
            ' FC
            oFuncionesB1.creaCampoMD("OINV", "SS_SerieRet", "(SS) Serie Retención", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 6, SAPbobsCOM.BoYesNoEnum.tNO)
            ' FC
            oFuncionesB1.creaCampoMD("OINV", "SS_SecRet", "(SS) Secuencial Retención", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
            ' FC
            oFuncionesB1.creaCampoMD("OINV", "SS_NumAutRet", "(SS) Num.Autor.Retención", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            ' FC
            oFuncionesB1.creaCampoMD("OINV", "SS_FecRet", "(SS) Fecha Retención", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            ' FC y F  Cuanto tipo de Documento es 41

            Try
                oFuncionesB1.creaCampoMD("OINV", "SS_Reembolsos", "(SS) Reembolsos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , , "SS_REEMCAB")
            Catch ex As Exception
                oFuncionesB1.creaCampoMD("OINV", "SS_Reembolsos", "(SS) Reembolsos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , , )
            End Try


            'oFuncionesB1.creaCampoMD("OINV", "SS_Reembolso", "(SS) Reembolsos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdValIP() As String = {"01", "02"}
            Dim StrcbdDesIP() As String = {"LOCAL", "EXTERIOR"}

            Dim StrcbdValPSRNL() As String = {"S", "N"}
            Dim StrcbdDesPSRNL() As String = {"SI", "NO"}


            oFuncionesB1.creaCampoMD("OINV", "SS_InfPago", "(SS) Inf. Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValIP, StrcbdDesIP, "01")



            'oFuncionesB1.creaCampoMD("OINV", "SS_ForPagCompras", "(SS) Forma de Pago Compras", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, , , , , "SS_FPAGCOM")
            '*************************PAGO RECIBIDO*************************************
            oFuncionesB1.creaCampoMD("RCT3", "SS_EstPtoRetRec", "(SS) Est y Pto. Retención", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("RCT3", "SS_SecRetRec", "(SS) Sec. Retencion Recibida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("RCT3", "SS_AutRetRec", "(SS) Num.Autori.Reten.Recibida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdValTCV() As String = {"18", "05", "48"}
            Dim StrcbdDesTCV() As String = {"Documentos autorizados utilizados en ventas excepto NC / ND", "Nota de Débito", "Nota de Débito por Reembolso emitida por intermediario"}
            oFuncionesB1.creaCampoMD("RCT3", "SS_TipoCompVenta", "(SS) Tipo Comprobante Venta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTCV, StrcbdDesTCV, "18")

            'oFuncionesB1.creaCampoMD("RCT3", "SS_MontoBaseImp", "(SS) Base Imponible Retenc", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Price, 15, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("RCT3", "SS_MontoBase", "(SS) Monto Base Imponible", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("RCT3", "SS_TipoFinanSN", "(SS) SN de Tipo Financiera", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("RCT3", "SS_NombreSN", "(SS) Nombre Socio Negocio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)


            'validar en RT y LQRT
            ' pasara a SOCIO de NEGOCIO

            oFuncionesB1.creaCampoMD("OCRD", "SS_AplDobTri", "(SS) Aplica Doble Tributación", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPSRNL, StrcbdDesPSRNL, "N")

            'FC
            Dim StrcbdValPagoLocExt() As String = {"01", "02"}
            Dim StrcbdDesPagoLocExt() As String = {"Residente", "No Residente"}
            oFuncionesB1.creaCampoMD("OCRD", "SS_PagoLocExt", "(SS) Pago Residente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPagoLocExt, StrcbdDesPagoLocExt, "01")

            oFuncionesB1.creaCampoMD("OCRD", "SS_ParteRel", "(SS) Parte Relacionada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPSRNL, StrcbdDesPSRNL, "N")

            Dim StrcbdValTipoRegiL() As String = {"01", "02", "03"}
            Dim StrcbdDesTipoRegiL() As String = {"General", "Paraíso Fiscal", "Fiscal Preferente o Juridicción"}

            oFuncionesB1.creaCampoMD("OCRD", "SS_TipoRegi", "(SS) Tip Regímen Fiscal Ext", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTipoRegiL, StrcbdDesTipoRegiL, "01")

            'tabla paises BF
            oFuncionesB1.creaCampoMD("OCRD", "SS_PaisEfecPago", "(SS) País se Efectua Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OCRD", "SS_PagExtSujRetNorLeg", "(SS) Pag.Sujeto a Ret N.Legal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPSRNL, StrcbdDesPSRNL, "N")
            oFuncionesB1.creaCampoMD("OCRD", "SS_PagoRegFis", "(SS) Pag. Regimen Fiscal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPSRNL, StrcbdDesPSRNL, "N")


            '' PARA EL MANEJO DE PROTESTO, CREA CAMPO DE USUARIO EN CHEQUE DONDE SE ALOJARA EL DOCENTRY DE LA NOTA DE DEBIDO GENERADA POR EL PROTESTO
            oFuncionesB1.creaCampoMD("RCT1", "SS_IDND", "(SS) ND PROTESTO", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValPSRNL, StrcbdDesPSRNL, "N", "", "", "13")
            Dim StrcbdValProtesto() As String = {"SI", "NO"}
            Dim StrcbdDesProtesto() As String = {"SI", "NO"}
            oFuncionesB1.creaCampoMD("OINV", "SS_PROTESTO", "(SS) PROTESTO", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValProtesto, StrcbdDesProtesto, "NO")
            'oFuncionesB1.creaCampoMD("OINV", "SS_IdEntMer", "(SS) Id Entrada Mercancias", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            'SALIDA
            oFuncionesB1.creaCampoMD("OINV", "SS_SalTrans", "(SS) Salida Transf Entre Compañias", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValProtesto, StrcbdDesProtesto, "NO")
            oFuncionesB1.creaCampoMD("OINV", "SS_BaseDestino", "(SS) Base Destino de Entrada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OINV", "SS_IdBodegaEnt", "(SS) Id Bodega Entrada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OINV", "SS_IdEntMer", "(SS) Id Entrada Mercancias", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            'ENTRADA
            oFuncionesB1.creaCampoMD("OINV", "SS_EntTrans", "(SS) Entrada Transf Entre Compañias", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValProtesto, StrcbdDesProtesto, "NO")
            oFuncionesB1.creaCampoMD("OINV", "SS_BaseOrigen", "(SS) Base Origen de Salida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OINV", "SS_IdBodegaSal", "(SS) Id Bodega Salida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OINV", "SS_IdSalMer", "(SS) Id Salida Mercancias", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)




        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_Catch , Error: " & ex.Message.ToString(), "Estructura")
        Finally
            GC.Collect()
        End Try

    End Sub

    Private Sub FUN_CreaTablas_LOCALIZACION()
        Try
            Try
                '' CONFIGURACIÓN - CATALOGO
                oFuncionesB1.creaTablaMD("SS_CONF", "(GS) CONFIGURACION", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("SS_CONFD", "(GS) CONFIGURACION DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Error al crear tabla SS_CONF SS_CONFD " & ex.Message, "Estructura")

            End Try

            Try

                ' CONFIGURACION SERIES DIAN
                oFuncionesB1.creaTablaMD("SS_SER", "(SS) DOC LEGALES SRI", SAPbobsCOM.BoUTBTableType.bott_MasterData)
                oFuncionesB1.creaTablaMD("SS_SERD", "(SS) LEGALES ELECTRONICO", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
                oFuncionesB1.creaTablaMD("SS_SERDP", "(SS) LEGALES PREIMPRESO", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Error al crear tabla SS_SER SS_SERD SS_SERDP " & ex.Message, "Estructura")

            End Try


            Try
                oFuncionesB1.creaTablaMD("SS_FORMASDEPAGO", "(SS) Formas de Pago", SAPbobsCOM.BoUTBTableType.bott_NoObject) 'SS_FORMAS_DE_PAGOS
            Catch ex As Exception
            End Try


            Try
                oFuncionesB1.creaTablaMD("SS_MES", "(SS) Mes", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_ANIO", "(SS) Año", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_MOTIVOS_NC", "(SS) Motivo Nota Credito", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try


            Try
                oFuncionesB1.creaTablaMD("SS_TIPO_ING_EXT", "(SS) Tipos de Ing del Exterior", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_CIUDAD", "(SS) Ciudad", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_PROVINCIA", "(SS) Provincia", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_PAIS", "(SS) Pais", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_DISTRITO_ADU", "(SS) Distrito Aduanero", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_COD_REGIMEN", "(SS) Codigo Regimen", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try



            Try
                oFuncionesB1.creaTablaMD("SS_REEMCAB", "(SS) Reembolso Cabecera", SAPbobsCOM.BoUTBTableType.bott_MasterData)
                oFuncionesB1.creaTablaMD("SS_REEMDET", "(SS) Reembolso Detalle", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            Catch ex As Exception
            End Try



            Try

                'Se cambio de Master table a tabla Normal
                oFuncionesB1.creaTablaMD("SS_TIPCOMAUT", "(SS) Tipo Comp. Autorizados", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try


            Try
                oFuncionesB1.creaTablaMD("SS_TRANSPORTISTA", "(SS) Transportista", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_TRANSPORTE", "(SS) Transporte", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            Catch ex As Exception
            End Try

            ' Se cambiara a Tipo Tabla Normal , borrar tabla y volver a Ingresarla
            Try
                oFuncionesB1.creaTablaMD("SS_SUSTRI", "(SS) Sustento Tributario", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try


            ' Seagregara la Tabla Para el udo de SERIES

            '10/02/2023 DM: se agrega para almacenar provincia, canton, parroquia para dinardap 
            Try
                oFuncionesB1.creaTablaMD("SS_ESQCOD", "(GS) ESQUEMA CODIFICACION", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try


            Try
                oFuncionesB1.creaTablaMD("SS_DINARDAP_TL", "(SS) DINARDAP Trans.Leg", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            Try
                oFuncionesB1.creaTablaMD("SS_BASES", "(SS) BASES PARA TRANSFERENCIA", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            Catch ex As Exception
            End Try

            'guias de remision desatendida

            Try

                oFuncionesB1.creaTablaMD("SS_GRCAB", "(SS) SS GR CABECERA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("SS_GRDET", "(SS) SS GR Contenido", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("SS_GRDET1", "(SS) SS GR Info Adicional", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Catch ex As Exception

            End Try


        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas_Catch , Error: " & ex.Message.ToString(), "Estructura")
        Finally
            GC.Collect()
        End Try
    End Sub

#Region "Definicion UDOS Localizacion"

    Private Sub FUN_CreaUDO_GUIAS_DESATENDIDAS()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SSGRNEW") Then
                oUdo.Code = "SSGRNEW"
                oUdo.Name = "SS GUIAS DESATENDIDAS"
                oUdo.TableName = "SS_GRCAB"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SS_GRCAB"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "SS_GRDET"
                oUdo.ChildTables.Add()
                oUdo.ChildTables.TableName = "SS_GRDET1"



                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SSGRNEW , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO SSGRNEW" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SSGRNEW, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SSGRNEW_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub FUN_CreaUDO_CONF_LOC()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_CONFLOC") Then
                oUdo.Code = "SS_CONFLOC"
                oUdo.Name = "SS SS_CONFLOC"
                oUdo.TableName = "SS_CONF"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                'oUdo.LogTableName = "A_SS_CONF"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "SS_CONFD"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_CONF , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO SS_CONFLOC" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_CONF, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_CONF_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub FUN_CreaUDO_SER_SRI()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_SER") Then
                oUdo.Code = "SS_SER"
                oUdo.Name = "SS SS_SER"
                oUdo.TableName = "SS_SER"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SS_SER"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData

                oUdo.ChildTables.TableName = "SS_SERD"
                oUdo.ChildTables.Add()
                oUdo.ChildTables.TableName = "SS_SERDP"



                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SER_SRI , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO SER" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_SER, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SER_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub CrearUDOTransporte()

        'oFuncionesB1.creaTablaMD("SS_CODFIS", "(SS) Codigos Fiscales", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        'oFuncionesB1.creaCampoMD("SS_CODFIS", "SS_Indmp", "(SS) Indicador de Impuesto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
        'oFuncionesB1.creaCampoMD("SS_CODFIS", "SS_DescImp", "(SS) Descripcion de Impuesto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)


        Dim oUdo As SAPbobsCOM.UserObjectsMD
        'Dim oUDOEnhancedForm As SAPbobsCOM.UserObjectMD_EnhancedFormColumns
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_TRANSPORTE") Then
                oUdo.Code = "SS_TRANSPORTE"
                oUdo.Name = "(SS) Transporte"
                oUdo.TableName = "SS_TRANSPORTE"
                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData

                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO
                'oUdo.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SS_TRANSPORTE"

                oUdo.FormColumns.FormColumnAlias = "Code"
                oUdo.FormColumns.FormColumnDescription = "Código"
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "Name"
                oUdo.FormColumns.FormColumnDescription = "Descripción"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_SS_Placa"
                oUdo.FormColumns.FormColumnDescription = "(SS) Placa"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_TipoTransporte"
                oUdo.FormColumns.FormColumnDescription = "(SS) Tipo Transporte"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_Capacidad"
                oUdo.FormColumns.FormColumnDescription = "(SS) Capacidad"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_Refrigerado"
                oUdo.FormColumns.FormColumnDescription = "(SS) Refrigerado"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_ObligacionesFiscales , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO FUN_CreaUDO_CodigosFiscales" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_CODFIS, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SER_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try


    End Sub

    Private Sub CrearUDOTransportista()


        Dim oUdo As SAPbobsCOM.UserObjectsMD
        'Dim oUDOEnhancedForm As SAPbobsCOM.UserObjectMD_EnhancedFormColumns
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_TRANSPORTISTA") Then
                oUdo.Code = "SS_TRANSPORTISTA"
                oUdo.Name = "(SS) Transportista"
                oUdo.TableName = "SS_TRANSPORTISTA"
                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData

                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO
                'oUdo.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SS_TRANSPORTISTA"

                oUdo.FormColumns.FormColumnAlias = "Code"
                oUdo.FormColumns.FormColumnDescription = "Código"
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "Name"
                oUdo.FormColumns.FormColumnDescription = "Descripción"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_Transportista"
                oUdo.FormColumns.FormColumnDescription = "(SS) ID Transportista"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_Estado"
                oUdo.FormColumns.FormColumnDescription = "(SS) Estado"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_TipoTrans"
                oUdo.FormColumns.FormColumnDescription = "(SS) Tipo Transportista"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_TipoId"
                oUdo.FormColumns.FormColumnDescription = "(SS) Tipo ID"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()
                oUdo.FormColumns.FormColumnAlias = "U_SS_AfectoRise"
                oUdo.FormColumns.FormColumnDescription = "(SS) Afecto a Rise"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.FormColumns.Add()

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SS_TRANSPORTISTA , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO FUN_CreaUDO_SS_TRANSPORTISTA" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_TRANSPORTISTA, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SS_TRANSPORTISTA_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub crearUdoReembolsos()



        Dim oUdo As SAPbobsCOM.UserObjectsMD
        'Dim oUDOEnhancedForm As SAPbobsCOM.UserObjectMD_EnhancedFormColumns
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_REEMCAB") Then
                oUdo.Code = "SS_REEMCAB"
                oUdo.Name = "(SS) Reembolsos"
                oUdo.TableName = "SS_REEMCAB"
                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData

                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SS_REEMCAB"
                'Busqueda para Choosefrom List y Poder vincular Udo a UDT

                'BUSQUEDA
                oUdo.FindColumns.ColumnAlias = "Code"
                oUdo.FindColumns.ColumnDescription = "Code"
                oUdo.FindColumns.Add()

                'CABECERA

                oUdo.FormColumns.FormColumnAlias = "Code"
                oUdo.FormColumns.FormColumnDescription = "Code"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                'DETALLES

                oUdo.ChildTables.TableName = "SS_REEMDET"

                'tablas hijas con formulario nuevo

                oUdo.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_TipoId"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Tipo ID"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 1
                oUdo.EnhancedFormColumns.Add()


                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_IdProv"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) ID Proveedor"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 2
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_TipoComp"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Tipo Comprobante"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 3
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_Est"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Establecimiento"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 4
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_PtoEmi"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Punto Emision"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 5
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_NumDoc"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Secuencial"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 6
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_NumAut"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Num Autorizacion"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 7
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_FecEmi"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Facha Emision"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 8
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_IVA0"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Base 0"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 9
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_IvaDif0"
                oUdo.EnhancedFormColumns.ColumnDescription = "SS) Base 12"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 10
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_NoObjIVA"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Base No objeto IVA"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 11
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_IvaExe"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Base Exenta"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 12
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_MontoIVA"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Monto IVA"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 13
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_SS_MontoICE"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Monto ICE"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 14
                oUdo.EnhancedFormColumns.Add()


                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_Reembolsos , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO FUN_CreaUDO_Reembolsos " + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_REEMCAB, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SER_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try


    End Sub

    Private Sub CreaUdos_Localizacion()

        FUN_CreaUDO_CONF_LOC() 'CATALOGO

        FUN_CreaUDO_SER_SRI()

        CrearUDOTransporte()

        CrearUDOTransportista()

        crearUdoReembolsos()

        FUN_CreaUDO_GUIAS_DESATENDIDAS()

    End Sub
#End Region

    Private Sub CrearEstructuraLocalizacion()

        Try
            rSboApp.StatusBar.SetText(NombreAddon + " - Creando Estructura de Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'crea tablas

            ' bgw.ReportProgress(5)
            rSboApp.StatusBar.SetText(NombreAddon + " - Creando TABLAS de Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            FUN_CreaTablas_LOCALIZACION()

            ' bgw.ReportProgress(10)
            rSboApp.StatusBar.SetText(NombreAddon + " - Creando CAMPOS de UDT para Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'crea campos
            FUN_CreaCampos_a_TablasDeUsuario_LOCALIZACION() '  a tablas de usuario

            ' bgw.ReportProgress(25)
            rSboApp.StatusBar.SetText(NombreAddon + " - Creando UDOS para Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            CreaUdos_Localizacion()

            ' bgw.ReportProgress(50)
            rSboApp.StatusBar.SetText(NombreAddon + " - Creando CAMPOS de MARKETING para Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            FUN_CreaCampos_LOCALIZACION() ' a tablas nativas

            ' bgw.ReportProgress(60)

            rSboApp.StatusBar.SetText(NombreAddon + " - Insertando Datos en UDTs para Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'inserta info
            Utilitario.Util_Log.Escribir_Log("Inicio InsertaFormasdePagos ", "Estructura")
            InsertaFormasdePagos()
            Utilitario.Util_Log.Escribir_Log("Fin InsertaFormasdePagos ", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarMes", "Estructura")
            InsertarMes()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarMes", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarAnio", "Estructura")
            InsertarAnio()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarAnio", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarMotivoNotaCredito", "Estructura")
            InsertarMotivoNotaCredito()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarMotivoNotaCredito", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarDistritoAduanero", "Estructura")
            InsertarDistritoAduanero()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarDistritoAduanero", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarCodigoRegimen", "Estructura")
            InsertarCodigoRegimen()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarCodigoRegimen", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertaPaises", "Estructura")
            InsertaPaises()
            Utilitario.Util_Log.Escribir_Log("Fin InsertaPaises", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarCiudades", "Estructura")
            InsertarCiudades()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarCiudades", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarProvincias", "Estructura")
            InsertarProvincias()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarProvincias", "Estructura")

            'bgw.ReportProgress(70)
            ''Anteriores esta OK------------------------------

            Utilitario.Util_Log.Escribir_Log("Inicio InsertaDatosTiposIdentificacion", "Estructura")
            InsertaDatosTiposIdentificacion()
            Utilitario.Util_Log.Escribir_Log("Fin InsertaDatosTiposIdentificacion", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertaDatosTiposTransaccion", "Estructura")
            InsertaDatosTiposTransaccion()
            Utilitario.Util_Log.Escribir_Log("Fin InsertaDatosTiposTransaccion", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarTiposIngresosExterior", "Estructura")
            InsertarTiposIngresosExterior()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarTiposIngresosExterior", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertaDatosTiposComprobantes", "Estructura")
            InsertaDatosTiposComprobantes()
            Utilitario.Util_Log.Escribir_Log("Fin InsertaDatosTiposComprobantes", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertaDatosSustentoTributario", "Estructura")
            InsertaDatosSustentoTributario()
            Utilitario.Util_Log.Escribir_Log("Fin InsertaDatosSustentoTributario", "Estructura")

            Utilitario.Util_Log.Escribir_Log("Inicio InsertarProvinciaCantonParroquia", "Estructura")
            InsertarProvinciaCantonParroquia()
            Utilitario.Util_Log.Escribir_Log("Fin InsertarProvinciaCantonParroquia", "Estructura")

            'bgw.ReportProgress(80)
            ' Se asocia BF
            rSboApp.StatusBar.SetText(NombreAddon + " - Asociando Busquedas Formateadas para Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            AsociarBusquedasFormateadas()

            'CrearBF()

            ' bgw.ReportProgress(100)

            rSboApp.StatusBar.SetText(NombreAddon + " - Finalizada la  Estructura de Localizacion EC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

        End Try



    End Sub


#End Region



    Private Sub CreaSPs()
        Dim ofrmProcedimientos As New ProcedimientosAlmacenados
        Dim str As String
        Dim mrst As SAPbobsCOM.Recordset = Nothing

        Try
            'lb1
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb1.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt1.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb2
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb2.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt2.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb3
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb3.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt3.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb4
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb4.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt4.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb5
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb5.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt5.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb6
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb6.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt6.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb7
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb7.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt7.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb8
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb8.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt8.Text)
            End If
        Catch ex As Exception

        End Try
        Try
            'lb9
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb9.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt9.Text)
            End If

        Catch ex As Exception

        End Try

        Try
            'lb10
            str = "select specific_name, ROUTINE_DEFINITION from information_schema.routines where specific_name ='" & ofrmProcedimientos.lb10.Text & "'"
            mrst = oFuncionesB1.getRecordSet(str)
            If Not mrst.EoF = False Then
                mrst.DoQuery(ofrmProcedimientos.txt10.Text)
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Crea las tablas de usuario 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FUN_CreaTablas()
        Try
            Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas_Catch , Licencia : " & Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToString(), "Estructura")
            ' oFuncionesB1.creaTablaMD("GS_CONFIGURACION", "(SS) Configuracion", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            ' LOG DE DOCUMENTOS ENVIADOS AL SRI ( LLEGARON A SAP POR IVEND )
            If Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "emision" Or Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "full" Then
                oFuncionesB1.creaTablaMD("GS_DOCUMENTOSTRANS", "(SS) Estado de Docs Inte", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Or _
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or _
                    Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    oFuncionesB1.creaTablaMD("GS_SERIESE", "(SS) Conf. Series Electrónicas", SAPbobsCOM.BoUTBTableType.bott_NoObject)
                End If
            End If

            ' LOG
            oFuncionesB1.creaTablaMD("GS_LOG", "(SS) LOG", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("GS_LOGD", "(SS) LOG DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            ' CONFIGURACIÓN - UDO
            oFuncionesB1.creaTablaMD("GS_CONF", "(SS) CONFIGURACION", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("GS_CONFD", "(SS) CONFIGURACION DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            'LIQUIDACION DE COMPRA
            oFuncionesB1.creaTablaMD("GS_LIQUI", "(SS) Series Liquidación Compra", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            If Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "recepcion" Or Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "full" Then
                'If rEstructura.LicenciaAddon = "FULL" Then
                ' DOCUMENTO RECIBIDO - FACTURA
                oFuncionesB1.creaTablaMD("GS_FVR", "(SS) FACTURA RECIBIDA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("GS0_FVR", "(SS) FACTURA RECIBIDA DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("GS1_FVR", "(SS) FACTURA RECIBIDA RELACION", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("GS2_FVR", "(SS) FACTURA RECIBIDA DATOS AD", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                ' END DOCUMENTO RECIBIDO - FACTURA
                ' DOCUMENTO RECIBIDO - NOTA DE CREDITO
                oFuncionesB1.creaTablaMD("GS_NCR", "(SS) NC RECIBIDA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("GS0_NCR", "(SS) NC RECIBIDA DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("GS1_NCR", "(SS) NC RECIBIDA RELACION", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("GS2_NCR", "(SS) NC RECIBIDA DATOS AD", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

                ' END DOCUMENTO RECIBIDO - NOTA DE CREDITO
                ' DOCUMENTO RECIBIDO - PAGO RECIBIDO RETENCION
                oFuncionesB1.creaTablaMD("GS_RER", "(SS) RE RECIBIDA", SAPbobsCOM.BoUTBTableType.bott_Document)
                oFuncionesB1.creaTablaMD("GS0_RER", "(SS) RE RECIBIDA DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                oFuncionesB1.creaTablaMD("GS1_RER", "(SS) RE RECIBIDA DATOS AD", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
                ' END DOCUMENTO RECIBIDO - PAGO RECIBIDO RETENCION

                ' RECEPCION - TABLA DE MAPEO IMPUESTO - IVA
                oFuncionesB1.creaTablaMD("GS_MAPEO_IVA", "(SS)Mapeo SRI Cod de Imp IVA", SAPbobsCOM.BoUTBTableType.bott_NoObject)
                oFuncionesB1.creaTablaMD("GS_MAPEO_RENTA", "(SS)Mapeo SRI Cod de Imp RENTA", SAPbobsCOM.BoUTBTableType.bott_NoObject)
                oFuncionesB1.creaTablaMD("GS_MAPEO_ISD", "(SS)Mapeo SRI Cod de Imp ISD", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                oFuncionesB1.creaTablaMD("GS_MAPEO_TC", "(SS)Mapeo Codigo de Impuesto", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                oFuncionesB1.creaTablaMD("GS_COMEN_PRCB", "(SS)Comentarios PagoRecibidoCB", SAPbobsCOM.BoUTBTableType.bott_NoObject)

                'END RECEPCION - TABLA DE MAPEO IMPUESTO - IVA

            End If



            oFuncionesB1.creaTablaMD("GS_SERIEXLUIR", "(SS) Conf. Series Excluir", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oFuncionesB1.creaTablaMD("SS_DIAS_PARAM", "(SS) Dias Parametrizables", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas_Catch , Error: " & ex.Message.ToString(), "Estructura")
        Finally
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' Crea los campos de Usuario (UDF)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FUN_CreaCampos()
        Try
            'oFuncionesB1.creaCampoMD("OUSR", "GS_FC", "FC Electronico Default", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            'oFuncionesB1.creaCampoMD("OUSR", "GS_NC", "NC Electronico Default", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            'oFuncionesB1.creaCampoMD("OUSR", "GS_ND", "ND Electronico Default", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            'oFuncionesB1.creaCampoMD("OUSR", "GS_GR", "GR Electronico Default", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            'oFuncionesB1.creaCampoMD("OUSR", "GS_RE", "RE Electronico Default", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            'oFuncionesB1.creaCampoMD("OUSR", "GS_CE", "RE Electronico Default", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            'oFuncionesB1.creaCampoMD("OINV", "GS_Elec", "Electronico", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "N")
            Utilitario.Util_Log.Escribir_Log("FUN_Creacamposh , Licencia : " & Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToString(), "Estructura")
            'If oLicencia.Opcion = "Emisión" Or oLicencia.Opcion = "FULL" Then
            If Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "emision" Or Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "full" Then
                Dim StrcbdVal() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "11"}
                Dim StrcbdDes() As String = {"NO ENVIADO", "EN PROCESO", "AUTORIZADA", "NO AUTORIZADA", "VALIDAR DATOS", "EN PROCESO SRI", "DEVUELTA", "ERROR EN RECEPCION", "ANULADO"}
                Dim StrcbdValEcau() As String = {"0", "1", "2", "4", "6", "11"}
                Dim StrcbdDesEcau() As String = {"NO ENVIADO", "EN PROCESO", "AUTORIZADA", "NO ENCONTRADO", "CON ERROR", "ANULADO"}
                Dim StrcbdVal2() As String = {"SI", "NO"}
                Dim StrcbdDes2() As String = {"SI", "NO"}
                oFuncionesB1.creaCampoMD("OINV", "CLAVE_ACCESO", "(SSE) Clave de Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)

                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                    oFuncionesB1.creaCampoMD("OINV", "ESTADO_AUTORIZACIO", "(SSE) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValEcau, StrcbdDesEcau, "0")
                Else
                    oFuncionesB1.creaCampoMD("OINV", "ESTADO_AUTORIZACIO", "(SSE) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "0")
                End If

                oFuncionesB1.creaCampoMD("OINV", "NUM_AUTO_FAC", "(SSE) Número de Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("OINV", "FECHA_AUT_FACT", "(SSE) Fecha de Autorización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("OINV", "OBSERVACION_FACT", "(SSE) Observación", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

                If Functions.VariablesGlobales._vgSerieUDF = "Y" Then
                    Dim StrcbdValSeriesUDF() As String = {"01", "02", "03", "04"}
                    Dim StrcbdDesSeriesUDF() As String = {"RETENCION", "LIQUIDACION", "LIQUIDACION Y RETENCION", "DOCUMENTO NO ELECTRONICO"}
                    oFuncionesB1.creaCampoMD("OINV", "DocEmision", "(SSE) Documento a Emitir", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValSeriesUDF, StrcbdDesSeriesUDF, "01")
                End If
                'CAMPOS PARA GUARDAR LA INFO DE AUTORIZACION DE LA LIQUIDACION DE COMPRA
                Dim StrcbdValLQ() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "11"}
                Dim StrcbdDesLQ() As String = {"NO ENVIADO", "EN PROCESO", "AUTORIZADA", "NO AUTORIZADA", "ERROR AL FIRMAR", "EN PROCESO SRI", "DEVUELTA", "ERROR EN RECEPCION", "ANULADO"}
                Dim StrcbdVal2LQ() As String = {"SI", "NO"}
                Dim StrcbdDes2LQ() As String = {"SI", "NO"}
                Dim StrcbdValLQEcau() As String = {"0", "1", "2", "4", "6", "11"}
                Dim StrcbdDesLQEcau() As String = {"NO ENVIADO", "EN PROCESO", "AUTORIZADA", "NO ENCONTRADO", "CON ERROR", "ANULADO"}

                oFuncionesB1.creaCampoMD("OINV", "LQ_CLAVE", "(SSE) LQ Clave de Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                If Functions.VariablesGlobales._IntegracionEcuanexus = "Y" Then
                    oFuncionesB1.creaCampoMD("OINV", "LQ_ESTADO", "(SSE) LQ Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValLQEcau, StrcbdDesLQEcau, "0")
                Else
                    oFuncionesB1.creaCampoMD("OINV", "LQ_ESTADO", "(SSE) LQ Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValLQ, StrcbdDesLQ, "0")
                End If

                oFuncionesB1.creaCampoMD("OINV", "LQ_NUM_AUTO", "(SSE) LQ Número de Autorizació", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("OINV", "LQ_FECHA_AUT", "(SSE) LQ Fecha de Autorización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("OINV", "LQ_OBSERVACION", "(SSE) LQ Observación", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

                'oFuncionesB1.ActualizaCampos("OINV", "LQ_CLAVE", "(SSE) LQ Clave de Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                'oFuncionesB1.ActualizaCampos("OINV", "LQ_ESTADO", "(SSE) LQ Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValLQ, StrcbdDesLQ, "0")
                'oFuncionesB1.ActualizaCampos("OINV", "LQ_NUM_AUTO", "(SSE) LQ Número de Autorizació", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                'oFuncionesB1.ActualizaCampos("OINV", "LQ_FECHA_AUT", "(SSE) LQ Fecha de Autorización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                'oFuncionesB1.ActualizaCampos("OINV", "LQ_OBSERVACION", "(SSE) LQ Observación", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

                ' TABLA INTERMEDIA DE IVEND
                oFuncionesB1.creaCampoMD("GS_DocumentosTrans", "DocEntry", "(GS) DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_DocumentosTrans", "ObjectType", "(GS) ObjectType", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_DocumentosTrans", "DocSubType", "(GS) DocSubType", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_DocumentosTrans", "SRI_Code", "(GS) Codigo SRI", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_DocumentosTrans", "Procesado", "(GS) Procesado", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_DocumentosTrans", "Oberva", "(GS) Observacion", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

                If Not Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Or Not Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
                    If Functions.VariablesGlobales._vgSerieUDF = "Y" Then
                        Dim ValUDF() As String = {"01", "02", "03", "04"}
                        Dim DesUDF() As String = {"RETENCION", "LIQUIDACION", "LIQUIDACION / RETENCION", "DOCUMENTO NO ELECTRONICO"}
                        oFuncionesB1.creaCampoMD("OINV", "DocEmision", "(SSE) Documento a Emitir", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, ValUDF, DesUDF, "01")
                    End If

                End If

                '' LA TABLA SERIES ES DE LA LOCALIZACION DE ONE SOLUTION, SI EXISTE CREO EL CAMPO ULT_SECUEN
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
                    oFuncionesB1.creaCampoMD("SERIES", "ULT_SECUEN", "Última Secuencia", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
                End If

                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN _
                   Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT _
                   Or Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    Dim StrcbdValELE() As String = {"SI", "NO"}
                    Dim StrcbdDesELE() As String = {"SI", "NO"}
                    oFuncionesB1.creaCampoMD("GS_SERIESE", "ELECTRONICA", "(SSE) Es Electrónica", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValELE, StrcbdDesELE, "")
                End If
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    oFuncionesB1.creaCampoMD("GS_SERIESE", "ULT_SECUEN", "Última Secuencia", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
                End If

                ' VALIDA SI ACTUALIZA LOS CAMPOS
                Dim ActualizaCamposDeUsuario As String = ""
                ActualizaCamposDeUsuario = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "ActualizaCamposDeUsuario")
                If ActualizaCamposDeUsuario = "Y" Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                    oFuncionesB1.ActualizaCampos("OINV", "ESTADO_AUTORIZACIO", "(SSE) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "0")
                    oFuncionesB1.ActualizaCampos("OINV", "CLAVE_ACCESO", "(SSE) Clave de Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("OINV", "ESTADO_AUTORIZACIO", "(SSE) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "0")
                    oFuncionesB1.ActualizaCampos("OINV", "NUM_AUTO_FAC", "(SSE) Número de Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("OINV", "FECHA_AUT_FACT", "(SSE) Fecha de Autorización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("OINV", "OBSERVACION_FACT", "(SSE) Observación", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

                    'CAMPOS PARA TABLA DE CONFIGRUACION DE LIQUIDACION DE COMPRA
                    oFuncionesB1.ActualizaCampos("GS_LIQUI", "IdSerie", "(SS) Id Serie", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("GS_LIQUI", "Serie", "(SS) Serie", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("GS_LIQUI", "Estable", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("GS_LIQUI", "PtoEmi", "(SS) Punto de Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("GS_LIQUI", "Sec", "(SS) Secuencial", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
                    oFuncionesB1.ActualizaCampos("GS_LIQUI", "Est", "(SS) Est. Usuario", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

                End If

                'CAMPOS PARA TABLA DE CONFIGRUACION DE LIQUIDACION DE COMPRA
                oFuncionesB1.creaCampoMD("GS_LIQUI", "IdSerie", "(SS) Id Serie", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_LIQUI", "Serie", "(SS) Serie", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_LIQUI", "Estable", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_LIQUI", "PtoEmi", "(SS) Punto de Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_LIQUI", "Sec", "(SS) Secuencial", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_LIQUI", "Est", "(SS) Est. Usuario", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 9, SAPbobsCOM.BoYesNoEnum.tNO)

                ' VALIDO SI EL MOTOR NO ES HANA PARA CREAR LOS SPs
                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Else
                    CreaSPs() 'SQL - EMISION
                End If

            End If

            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_FC", "(GS) WS Factura", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_ND", "(GS) WS Nota de Debito", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_NC", "(GS) WS Nota de Credito", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_GR", "(GS) WS Guia de Remision", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_RE", "(GS) WS Retencion", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_CE", "(GS) WS Consulta Emision", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_RM", "(GS) WS Reenvio Mail", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_Cl", "(GS) WS Clave", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            '' TIPO DE EMISION, 1 EN LINEA, 2 SEMIENLINEA
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_TP", "(GS) Tipo Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, SAPbobsCOM.BoYesNoEnum.tNO)
            'If rEstructura.LicenciaAddon = "FULL" Then 'Or oLicencia.Opcion = "FULL" Then

            'If oLicencia.Opcion = "Recepción" Or oLicencia.Opcion = "FULL" Then
            If Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "recepcion" Or Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "full" Then

                ' RECEPCION
                Dim StrcbdValRE() As String = {"SI", "NO"}
                Dim StrcbdDesRE() As String = {"SI", "NO"}
                oFuncionesB1.creaCampoMD("OINV", "SSCREADAR", "(SSR) Creada por Addon", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValRE, StrcbdDesRE, "NO")
                oFuncionesB1.creaCampoMD("OINV", "SSIDDOCUMENTO", "(SSR) Id Documento Recibido", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("RCT3", "SSCREADAR", "(SSR) Creada por Addon", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValRE, StrcbdDesRE, "NO")
                oFuncionesB1.creaCampoMD("RCT3", "SSIDDOCUMENTO", "(SSR) Id Documento Recibido", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("ORCT", "SSCREADAR", "(SSR) Creada por Addon", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValRE, StrcbdDesRE, "NO")
                oFuncionesB1.creaCampoMD("ORCT", "SSIDDOCUMENTO", "(SSR) Id Documento Recibido", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
                    Try
                        oFuncionesB1.creaCampoMD("TM_LE_RETVH", "SSCREADAR", "(SSR) Creada por Addon", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValRE, StrcbdDesRE, "NO")
                        oFuncionesB1.creaCampoMD("TM_LE_RETVH", "SSIDDOCUMENTO", "(SSR) Id Documento Recibido", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                    Catch ex As Exception
                    End Try
                End If
                'U_FX_AUTO_RETENCION -- CAMPO PARA GUARDAR EL NUMERO DE AUTORIZACION DE LA RETENCION, ESTE CAMPO LO CREO CARLOS FIGUERA EN DIBEAL, Y SE USARA PARA LA LOCALIZACION SYPSOFT
                If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Then
                    oFuncionesB1.creaCampoMD("ORCT", "FX_AUTO_RETENCION", "(SSR) No Autorizacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                End If

                oFuncionesB1.creaCampoMD("OSCN", "ItemName", "(SSR) Descripción Item", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

                oFuncionesB1.creaCampoMD("OCRD", "SSCUENTA", "(SSR) Cuenta de Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("OCRD", "SSCEN_COS", "(SSR) Centro de costo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("OCRD", "SSMARCA", "(SSR) Marca", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
                Dim FechaValRE() As String = {"SI", "NO"}
                Dim FechadDesRE() As String = {"SI", "NO"}
                oFuncionesB1.creaCampoMD("OCRD", "SSFECH_EMI", "(SSR) Fecha Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, FechaValRE, FechadDesRE, "NO")


                ' CAMPOS TABLA MAPEO IMPUESTO - IVA
                Dim StrcbdValIVA() As String = {"9", "10", "1", "11", "2", "3"}
                Dim StrcbdDesIVA() As String = {"10 %", "20 %", "30 %", "50 %", "70 %", "100 %"}
                oFuncionesB1.creaCampoMD("GS_MAPEO_IVA", "SSCOD", "(SSR) Codigo SRI", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValIVA, StrcbdDesIVA, "9")
                oFuncionesB1.creaCampoMD("GS_MAPEO_IVA", "SSID", "(SSR) ID Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_MAPEO_IVA", "SSDES", "(SSR) DESC Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
                CreaCampoMapeoImpuestoRenta()
                'END CAMPOS TABLA MAPEO IMPUESTO - IVA

                ' CAMPOS TABLA MAPEO ISD - IVA GS_MAPEO_ISD
                Dim StrcbdValISD() As String = {"4580"}
                Dim StrcbdDesISD() As String = {"5 %"}
                oFuncionesB1.creaCampoMD("GS_MAPEO_ISD", "SSCOD", "(SSR) Codigo SRI", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValISD, StrcbdDesISD, "4580")
                oFuncionesB1.creaCampoMD("GS_MAPEO_ISD", "SSID", "(SSR) ID Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_MAPEO_ISD", "SSDES", "(SSR) DESC Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
                ' END CAMPOS TABLA MAPEO ISD - IVA GS_MAPEO_ISD

                ' CAMPOS TABLA MAPEO TC - CODIGO DE IMPUESTO
                Dim StrcbdValTC() As String = {"0", "2", "3", "6", "7", "5", "8", "15"}
                Dim StrcbdDesTC() As String = {"0 %", "12 %", "14 %", "No Objeto de Impuesto", "Exento de IVA", "5%", "8%", "15%"}
                oFuncionesB1.creaCampoMD("GS_MAPEO_TC", "SSCOD", "(SSR) Codigo SRI", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValTC, StrcbdDesTC, "0")
                oFuncionesB1.creaCampoMD("GS_MAPEO_TC", "SSID", "(SSR) ID Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_MAPEO_TC", "SSDES", "(SSR) Desc Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
                ' END CAMPOS TABLA MAPEO TC - CODIGO DE IMPUESTO

                'manamer comentarios pago recibido cliente bancario
                oFuncionesB1.creaCampoMD("GS_COMEN_PRCB", "Code", "(SS) Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_COMEN_PRCB", "SocioNegocio", "(SS) Socio de Negocio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
                oFuncionesB1.creaCampoMD("GS_COMEN_PRCB", "Comentario", "(SS) Comentario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
                'oFuncionesB1.creaCampoMD("GS_COMENTARIOS_PR_CB", "PtoEmi", "(SS) Punto de Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
                'fin manamer comentarios pago recibido cliente bancario
                'campo para pago recibido clientes bancos               
                Try
                    Dim StrcbdValRECB() As String = {"SI", "NO", "PROVEEDOR"}
                    Dim StrcbdDesRECB() As String = {"SI", "NO", "PROVEEDOR"}
                    oFuncionesB1.creaCampoMD("OCRD", "SSCLIENTEBANCO", "(SSE) Cliente Bancario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValRECB, StrcbdDesRECB, "NO")
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_Catch , Error263: " & ex.Message.ToString(), "Estructura")
                End Try

                CreaCamposFacturaRecibida()
                CreaCamposNotaDeCreditoRecibida()
                CreaCamposRetencionRecibida()

                CreaUDOFacturaRecibida() ' RECEPCION
                CreaUDONotaDeCredito() ' RECEPCION
                'CreaUDORetencion() ' RECEPCION

                crearUdoRegistradoRetencionRecibida()
                oFuncionesB1.creaCampoMD("OINV", "SS_IDRETCERO", "(SS) Id Retencion cero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, , , , , "GS_RER")

            End If

            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_C", "(GS) WS Recepcion Consulta", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_CES", "(GS) WS Recepcion Estado", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_CC", "(GS) WS Recepcion Clave", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_OC", "(GS) WS Recepcion Campo OC", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "ws_RA", "(GS) WS Recepcion Archivo", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'Dim StrcbdValP() As String = {"EXXIS", "ONESOLUTIONS", "SYPSOFT", "HEINSOHN"}
            'Dim StrcbdDesP() As String = {"EXXIS", "ONESOLUTIONS", "SYPSOFT", "HEINSOHN"}
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "PRO", "(GS) Proveedor SAP BO", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValP, StrcbdDesP, "EXXIS")

            'Dim StrcbdValE() As String = {"LOCAL", "NUBE"}
            'Dim StrcbdDesE() As String = {"LOCAL", "NUBE"}
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "WS", "(GS) TIPO WS", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValE, StrcbdDesE, "LOCAL")

            ''GS_DocumentosRecibidos Integrados
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "ObjType", "(GS) ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "DocSubType", "(GS) DocSubType", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "DocEntry", "(GS) DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "Folio", "(GS) Folio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "CardCode", "(GS) CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "CardName", "(GS) CardName", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "Valor", "(GS) Valor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "ClaAcce", "(GS) ClaveAcceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            'Dim StrcbdVal3() As String = {"docPreliminar", "docFinal", "docSincronizado", "docReAbierto", "docCancelado", "Error"}
            'Dim StrcbdDes3() As String = {"docPreliminar", "docFinal", "docSincronizado", "docReAbierto", "docCancelado", "Error"}  ' Si esta o No sincronizado con GS
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "Fecha", "(GS) Fecha", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "Estado", "(GS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "Tipo", "(GS) Tipo Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "Log", "(GS) Log Desc", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "ObjTypeR", "ObjType Relacionado", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRec", "DocEntryR", "DocEntrys Relacionados", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            ''
            ' RECEPCION PARAMETRIZACIÓN
            'oFuncionesB1.creaCampoMD("GS_DocumentosRecP", "Pref", "Prefijo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRecP", "CuentaS", "Cuenta Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRecP", "CuentaDS", "Des Cuenta Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRecP", "PrefN", "Prefijo NC", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_DocumentosRecP", "CuentaSN", "Cuenta NC S", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_Configuracion", "actCa", "(GS) Actualiza Campos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal2, StrcbdDes2, "NO")

            ' LOG
            oFuncionesB1.creaCampoMD("GS_LOG", "Clave", "(SS) DocEntry/Clav", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_LOG", "ObjType", "(SS) ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_LOG", "SubType", "(SS) DocSubType", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 5, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal3() As String = {"Emision", "Recepcion"}
            Dim StrcbdDes3() As String = {"Emision", "Recepcion"}
            oFuncionesB1.creaCampoMD("GS_LOG", "Tipo", "(SS) Tipo LOG", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            oFuncionesB1.creaCampoMD("GS_LOGD", "Transacc", "(SS) Transaccion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_LOGD", "Detalle", "(SS) Detalle", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_LOGD", "Fecha", "(SS) Fecha", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 40, SAPbobsCOM.BoYesNoEnum.tNO)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_Catch , Error: " & ex.Message.ToString(), "Estructura")
        Finally
            GC.Collect()
        End Try

    End Sub

    Private Sub CreaCampoMapeoImpuestoRenta()
        ' CAMPOS TABLA MAPEO IMPUESTO - RENTA
        Dim StrcbdValRENTA() As String = {"303", "303A", "304", "304A ", "304B ", "304C ", "304D", "304E ", "307", "308", "309", "310", "311", "312", "312A ", "314A", "314C", "314D", "319", "320", "322",
                                          "323", "323A", "323B1", "323E", "323E2", "323F", "323G", "323H", "323I", "323 M", "323 N", "323 O", "323 P", "323Q", "323R", "324A", "324B ", "325", "325A",
                                          "326", "327", "328", "329", "330", "331", "332", "332A", "332B", "332C", "332D", "332E", "332F", "332G", "332H", "332I", "333", "334", "335", "336", "337", "338", "339", "340",
                                          "341", "342", "342A", "342B ", "343A", "344", "344A", "346A", "347", "348", "349", "500", "501", "502", "503", "504", "504A", "504B", "504C", "504D", "504F", "504G", "504H",
                                          "505", "505A", "505B", "505C", "505D", "505E", "505F", "509", "509A", "510", "511", "512", "513", "513A", "514", "515", "516", "517", "518", "519", "520", "520A",
                                          "520B", "520D", "520E", "520F", "520G ", "521", "522A", "523A", "524", "525", "3440", "343", "312C", "314B", "323S", "323T", "323U", "324C", "343B", "343C", "344B",
                                          "345", "346"}

        Dim StrcbdDesRENTA() As String = {"Honorarios profesionales y demás pagos por servicios relacionados con el título profesional ", "Servicios profesionales prestados por sociedades residentes", "Servicios predomina el intelecto no relacionados con el título profesional", "Comisiones y demás pagos por servicios predomina intelecto no relacionados con el título profesional", "Pagos a notarios y registradores de la propiedad y mercantil por sus actividades ejercidas como tales",
                                          "Pagos a deportistas, entrenadores, árbitros, miembros del cuerpo técnico por sus actividades ejercidas como tales", "Pagos a artistas por sus actividades ejercidas como tales", "Honorarios y demás pagos por servicios de docencia", "Servicios predomina la mano de obra", "Utilización o aprovechamiento de la imagen o renombre", "Servicios prestados por medios de comunicación y agencias de publicidad",
                                          "Servicio de transporte privado de pasajeros o transporte público o privado de carga", "Por pagos a través de liquidación de compra (nivel cultural o rusticidad)", "Transferencia de bienes muebles de naturaleza corporal", "Compra de bienes de origen agrícola, avícola, pecuario, apícola, cunícula, bioacuático, y forestal", "Regalías por concepto de franquicias de acuerdo a Ley de Propiedad Intelectual - pago a personas naturales 314B Cánones, derechos de autor,  marcas, patentes y similares de acuerdo a Ley de Propiedad Intelectual – pago a personas naturales",
                                          "Regalías por concepto de franquicias de acuerdo a Ley de Propiedad Intelectual  - pago a sociedades", "Cánones, derechos de autor,  marcas, patentes y similares de acuerdo a Ley de Propiedad Intelectual – pago a sociedades ", "Cuotas de arrendamiento mercantil, inclusive la de opción de compra", "Por arrendamiento bienes inmuebles", "Seguros y reaseguros (primas y cesiones)", "Por rendimientos financieros pagados a naturales y sociedades  (No a IFIs)", "Por RF: depósitos Cta. Corriente",
                                          "Por RF:  depósitos Cta. Ahorros Sociedades", "Por RF: depósito a plazo fijo  gravados", "Por RF: depósito a plazo fijo exentos", "Por rendimientos financieros: operaciones de reporto - repos", "Por RF: inversiones (captaciones) rendimientos distintos de aquellos pagados a IFIs", "Por RF: obligaciones", "Por RF: bonos convertible en acciones", "Por RF: Inversiones en títulos valores en renta fija gravados", "Por RF: Inversiones en títulos valores en renta fija exentos", "Por RF: Intereses pagados a bancos y otras entidades sometidas al control de la Superintendencia de Bancos y de la Economía Popular y Solidaria",
                                          "Por RF: Intereses pagados por entidades del sector público a favor de sujetos pasivos", "Por RF: Otros intereses y rendimientos financieros gravados", "Por RF: Otros intereses y rendimientos financieros exentos", "Por RF: Intereses en operaciones de crédito entre instituciones del sistema financiero y entidades economía popular y solidaria.", "Por RF: Por inversiones entre instituciones del sistema financiero y entidades economía popular y solidaria.", "Anticipo dividendos", "Dividendos anticipados préstamos accionistas, beneficiarios o partícipes", "Dividendos distribuidos que correspondan al impuesto a la renta único establecido en el art. 27 de la lrti",
                                          "Dividendos distribuidos a personas naturales residentes cuando la sociedad que distribuye aplicó tarifa del 22% IR", "Dividendos distribuidos a sociedades residentes", "Dividendos distribuidos a fideicomisos residentes", "Dividendos gravados distribuidos en acciones (reinversión de utilidades sin derecho a reducción tarifa IR) cuando la sociedad que distribuye aplicó tarifa del 22% IR", "Dividendos exentos distribuidos en acciones (reinversión de utilidades con derecho a reducción tarifa IR) ", "Otras compras de bienes y servicios no sujetas a retención", "Por la enajenación ocasional de acciones o participaciones y títulos valores",
                                          "Compra de bienes inmuebles", "Transporte público de pasajeros", "Pagos en el país por transporte de pasajeros o transporte internacional de carga, a compañías nacionales o extranjeras de aviación o marítimas", "Valores entregados por las cooperativas de transporte a sus socios", "Compraventa de divisas distintas al dólar de los Estados Unidos de América", "Pagos con tarjeta de crédito ", "Pago al exterior tarjeta de crédito reportada por la Emisora de tarjeta de crédito, solo recap ", "Pago a través de convenio de debito (Clientes IFI`s) ", "Enajenación de derechos representativos de capital y otros derechos cotizados en bolsa ecuatoriana ",
                                          "Enajenación de derechos representativos de capital y otros derechos no cotizados en bolsa ecuatoriana", "Por loterías, rifas, apuestas y similares", "Por venta de combustibles a comercializadoras", "Por venta de combustibles a distribuidores", "Compra local de banano a productor", "Liquidación impuesto único a la venta local de banano de producción propia", "Impuesto único a la exportación de banano de producción propia - componente 1", "Impuesto único a la exportación de banano de producción propia - componente 2", "Impuesto único a la exportación de banano producido por terceros", "Impuesto único a la exportación de banano producido por terceros de Asociaciones de micro y pequeños productores hasta 1000 cajas por semana por cada socio.",
                                          "Impuesto único a la exportación de banano producido por terceros de Asociaciones de micro, pequeños y medianos productores", "Por energía eléctrica 343B Por actividades de construcción de obra material inmueble, urbanización, lotización o actividades similares", "Otras retenciones aplicables el 2%", "Pago local tarjeta de crédito /debito reportada por la Emisora de tarjeta de crédito /debito , solo recap", "Ganancias de capital", "Donaciones en dinero -Impuesto a la donaciones", "Retención a cargo del propio sujeto pasivo por la exportación de concentrados y/o elementos metálicos", "Retención a cargo del propio sujeto pasivo por la comercialización de productos forestales", "Pago al exterior - Rentas Inmobiliarias", "Pago al exterior - Beneficios Empresariales",
                                          "Pago al exterior - Servicios Empresariales", "Pago al exterior - Navegación Marítima y/o aérea", "Pago al exterior- Dividendos distribuidos a personas naturales", "Pago al exterior - Dividendos a sociedades", "Pago al exterior - Anticipo dividendos (excepto paraisos fiscales o de regimen de menor imposición)", "Pago al exterior - Dividendos anticipados préstamos accionistas, beneficiarios o partìcipes (paraísos fiscales o regímenes de menor imposición)", "Pago al exterior - Dividendos a fideicomisos", "Pago al exterior - Dividendos a sociedades  (paraísos fiscales)", "Pago al exterior - Anticipo dividendos  (paraísos fiscales)", "Pago al exterior - Dividendos a fideicomisos  (paraísos fiscales)", "Pago al exterior - Rendimientos financieros",
                                          "Pago al exterior – Intereses de créditos de Instituciones Financieras del exterior", "Pago al exterior – Intereses de créditos de gobierno a gobierno", "Pago al exterior – Intereses de créditos de organismos multilaterales", "Pago al exterior - Intereses por financiamiento de proveedores externos", "Pago al exterior - Intereses de otros créditos externos", "Pago al exterior - Otros Intereses y Rendimientos Financieros", "Pago al exterior - Cánones, derechos de autor,  marcas, patentes y similares", "Pago al exterior - Regalías por concepto de franquicias", "Pago al exterior - Ganancias de capital", "Pago al exterior - Servicios profesionales independientes", "Pago al exterior - Servicios profesionales dependientes", "Pago al exterior - Artistas",
                                          "Pago al exterior - Deportistas", "Pago al exterior - Participación de consejeros", "Pago al exterior - Entretenimiento Público", "Pago al exterior - Pensiones", "Pago al exterior - Reembolso de Gastos", "Pago al exterior - Funciones Públicas", "Pago al exterior - Estudiantes", "Pago al exterior - Otros conceptos de ingresos gravados", "Pago al exterior - Pago a proveedores de servicios hoteleros y turísticos en el exterior", "Pago al exterior - Arrendamientos mercantil internacional", "Pago al exterior - Comisiones por exportaciones y por promoción de turismo receptivo", "Pago al exterior - Por las empresas de transporte marítimo o aéreo y por empresas pesqueras de alta mar, por su actividad.", "Pago al exterior - Por las agencias internacionales de prensa",
                                          "Pago al exterior - Contratos de fletamento de naves para empresas de transporte aéreo o marítimo internacional", "Pago al exterior - Enajenación de derechos representativos de capital y otros derechos", "Pago al exterior - Servicios técnicos, administrativos o de consultoría y regalías con convenio de doble tributación", "Pago al exterior - Seguros y reaseguros (primas y cesiones)  con convenio de doble tributación", "Pago al exterior - Otros pagos al exterior no sujetos a retención", "Pago al exterior - Donaciones en dinero -Impuesto a la donaciones", "Otras Retenciones aplicables al 2.75%", "Otras retenciones aplicables el 1% (incluye régimen RIMPE - Emprendedores)",
                                          "COMPRAS AL COMERCIALIZADOR: de bienes de origen bioacuático, forestal y los descritos  el art.27.1 de LRTI", "Cánones, derechos de autor,  marcas, patentes y similares de acuerdo  al Código INGENIOS (COESCCI) – pago a personas naturales", "Pagos y créditos en cuenta efectuados por el BCE y los depósitos centralizados de valores, en calidad de intermediarios, a instituciones del sistema financiero por cuenta de otras personas naturales y sociedades", "Rendimientos financieros originados en la deuda pública ecuatoriana", "Rendimientos financieros originados en títulos valores de obligaciones de 360 días o más para el financiamiento de proyectos públicos en asociación público-privada",
                                          "Pagos y créditos en cuenta efectuados por el BCE y los depósitos centralizados de valores, en calidad de intermediarios, a instituciones del sistema financiero por cuenta de otras instituciones del sistema financiero", "Actividades de construcción de obra material inmueble, urbanización, lotización o actividades similares", "Recepción de botellas plásticas no retornables de PET", "Adquisición de sustancias minerales dentro del territorio nacional", "Otras retenciones aplicables el 8%", "Otras retenciones aplicables a otros porcentajes"}

        oFuncionesB1.creaCampoMD("GS_MAPEO_RENTA", "SSID", "(SSR) ID Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
        oFuncionesB1.creaCampoMD("GS_MAPEO_RENTA", "SSDES", "(SSR) DESC Codigo SAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
        oFuncionesB1.creaCampoMD("GS_MAPEO_RENTA", "SSCOD", "(SSR) Codigo SRI", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValRENTA, StrcbdDesRENTA, "303")
        'END CAMPOS TABLA MAPEO IMPUESTO - RENTA

    End Sub

    Public Sub CreacionDeCampos_Y_UDO_Configuracion()
        ' CONFIGURACION - UDO
        ' CONFIGURACIÓN - UDO
        oFuncionesB1.creaTablaMD("GS_CONF", "(SS) CONFIGURACION", SAPbobsCOM.BoUTBTableType.bott_Document)
        oFuncionesB1.creaTablaMD("GS_CONFD", "(SS) CONFIGURACION DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        oFuncionesB1.creaCampoMD("GS_CONF", "Modulo", "(SS) Modulo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
        oFuncionesB1.creaCampoMD("GS_CONF", "Tipo", "(SS) Tipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
        oFuncionesB1.creaCampoMD("GS_CONF", "Subtipo", "(SS) Subtipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
        oFuncionesB1.creaCampoMD("GS_CONFD", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
        oFuncionesB1.creaCampoMD("GS_CONFD", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

        FUN_CreaUDO_CONF() ' EMISION Y RECEPCION select * from "@GS_CONF" select * from "@GS_CONFD"

    End Sub

    Private Sub FUN_CreaUDO_LOG()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_LOG") Then
                oUdo.Code = "SS_LOG"
                oUdo.Name = "SS LOG"
                oUdo.TableName = "GS_LOG"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_GS_LOG"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "GS_LOGD"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO LOG" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_LOG , Error: " + sErrMsg.ToString(), "Estructura")
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_LOG, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_LOG_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
#Disable Warning BC42104 ' La variable 'oUdo' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
#Enable Warning BC42104 ' La variable 'oUdo' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            oUdo = Nothing
            GC.Collect()
        End Try

    End Sub
    Private Sub FUN_CreaUDO_CONF()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_CONF") Then
                oUdo.Code = "SS_CONF"
                oUdo.Name = "SS SS_CONF"
                oUdo.TableName = "GS_CONF"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_GS_CONF"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "GS_CONFD"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_CONF , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO CONF" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_CONF, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_CONF_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
#Disable Warning BC42104 ' La variable 'oUdo' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
#Enable Warning BC42104 ' La variable 'oUdo' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
            oUdo = Nothing
            GC.Collect()
        End Try

    End Sub

    Private Sub CreaCamposFacturaRecibida()
        Try

            ' DOCUMENTO RECIBIDO FACTURA
            oFuncionesB1.creaCampoMD("GS_FVR", "RUC", "(SS) RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "CardCode", "(SS) CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Mapeado", "(SS) Mapeado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "ClaAcc", "(SS) Clave Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "NumAut", "(SS) Numero de Auto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "FecAut", "(SS) Fecha de Auto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "NumDoc", "(SS) Numero de Doc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "FPrelim", "(SS) Factura Prel", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ' DOCUMENTO RECIBIDO FACTURA - Totales
            oFuncionesB1.creaCampoMD("GS_FVR", "SubTot", "(SS) SubTotal", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "SubTot5", "(SS) SubTotal5", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Sub0", "(SS) SubTotal 0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "SubNO", "(SS) SubTotalNoOb", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "SubEx", "(SS) SubTotalExen", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "SubSI", "(SS) SubTotalSinIm", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Desc", "(SS) Descuento", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "ICE", "(SS) ICE", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "IVA", "(SS) IVA", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "IVA5", "(SS) IVA5", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "vTotal", "(SS) Valor Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ' DOCUMENTO RECIBIDO FACTURA - Totales Documentos Relacionados
            oFuncionesB1.creaCampoMD("GS_FVR", "rTades", "(SS) TotAntDes R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "rPDesc", "(SS) % Descuento R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "rDesc", "(SS) Descuento R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "rGast", "(SS) Gastos Adic R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "rImp", "(SS) Impuesto R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "rTotal", "(SS) ValorTotal R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Tipo", "(SS)Tipo Fact Grabd", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "IdGS", "(SS) Id Doc GS", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Sincro", "(SS) Sincronizado", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "SincroE", "(SS) Sincro EDOC", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal3() As String = {"docPreliminar", "docFinal", "docCancelado", "docMarcado", "docReAbierto", "docPrelXML"}
            Dim StrcbdDes3() As String = {"docPreliminar", "docFinal", "docCancelado", "docMarcado", "docReAbierto", "docPrelXML"}
            oFuncionesB1.creaCampoMD("GS_FVR", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            ' CAMPO FECHA QUE SE LO USARÁ PARA EL REPORTE DE DOCUMENTOS INTEGRADOS
            oFuncionesB1.creaCampoMD("GS_FVR", "FechaS", "(SS)Fecha DocFinal", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 8, SAPbobsCOM.BoYesNoEnum.tNO)

            'rutacompartida
            oFuncionesB1.creaCampoMD("GS_FVR", "Ruta_xml", "(SS) Ruta_xml", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FVR", "Ruta_pdf", "(SS) Ruta_pdf", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)



            ' DOCUMENTO RECIBIDO FACTURA - Detalle
            oFuncionesB1.creaCampoMD("GS0_FVR", "CodPrin", "(SS) CodPrin", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "CodAuxi", "(SS) CodAuxi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "CodSAP", "(SS) CodSAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "Descripc", "(SS) Descripc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "Cantid", "(SS) Cantid", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "Precio", "(SS) Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "Desc", "(SS) Desc", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_FVR", "Total", "(SS) Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ''xml y factura 
            
            '   ' DOCUMENTO RECIBIDO FACTURA - Detalle Doc Relacionados
            oFuncionesB1.creaCampoMD("GS1_FVR", "DocEntr", "(SS) DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "LineNu", "(SS) LineNum", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "ItemCode", "(SS) ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "Descripc", "(SS) Dscription", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "Cantid", "(SS) Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "Precio", "(SS) Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "DiscPr", "(SS) DiscPrcnt", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "TaxCode", "(SS) TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "lTotal", "(SS) LineTotal", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_FVR", "ObjType", "(SS) ObjType", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS2_FVR", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS2_FVR", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_FACTURA , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub
    Private Sub CreaUDOFacturaRecibida()
        Try
            Dim Child3() As String = {"GS0_FVR", "GS1_FVR", "GS2_FVR"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("GS_FVR", "(SS) Factura Recibida", "GS_FVR", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CreaCamposNotaDeCreditoRecibida()
        Try

            ' DOCUMENTO RECIBIDO FACTURA
            oFuncionesB1.creaCampoMD("GS_NCR", "RUC", "(SS) RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "CardCode", "(SS) CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Mapeado", "(SS) Mapeado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "ClaAcc", "(SS) Clave Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "NumAut", "(SS) Numero de Auto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "FecAut", "(SS) Fecha de Auto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "NumDoc", "(SS) Numero de Doc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "FPrelim", "(SS) Factura Prel", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ' DOCUMENTO RECIBIDO FACTURA - Totales
            oFuncionesB1.creaCampoMD("GS_NCR", "SubTot", "(SS) SubTotal", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "SubTot5", "(SS) SubTotal5", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Sub0", "(SS) SubTotal 0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "SubNO", "(SS) SubTotalNoOb", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "SubEx", "(SS) SubTotalExen", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "SubSI", "(SS) SubTotalSinIm", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Desc", "(SS) Descuento", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "ICE", "(SS) ICE", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "IVA", "(SS) IVA", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "IVA5", "(SS) IVA5", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "vTotal", "(SS) Valor Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ' DOCUMENTO RECIBIDO FACTURA - Totales Documentos Relacionados
            oFuncionesB1.creaCampoMD("GS_NCR", "rTades", "(SS) TotAntDes R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "rPDesc", "(SS) % Descuento R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "rDesc", "(SS) Descuento R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "rGast", "(SS) Gastos Adic R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "rImp", "(SS) Impuesto R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "rTotal", "(SS) ValorTotal R", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Tipo", "(SS)Tipo Fact Grabd", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "IdGS", "(SS) Id Doc GS", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Sincro", "(SS) Sincronizado", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "SincroE", "(SS) Sincro EDOC", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal3() As String = {"docPreliminar", "docFinal", "docCancelado", "docMarcado", "docReAbierto", "docPrelXML"}
            Dim StrcbdDes3() As String = {"docPreliminar", "docFinal", "docCancelado", "docMarcado", "docReAbierto", "docPrelXML"}
            oFuncionesB1.creaCampoMD("GS_NCR", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            ' CAMPO FECHA QUE SE LO USARÁ PARA EL REPORTE DE DOCUMENTOS INTEGRADOS
            oFuncionesB1.creaCampoMD("GS_NCR", "FechaS", "(SS)Fecha DocFinal", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 8, SAPbobsCOM.BoYesNoEnum.tNO)
            'xml y pdf
            oFuncionesB1.creaCampoMD("GS_NCR", "Ruta_xml", "(SS) Ruta_xml", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCR", "Ruta_pdf", "(SS) Ruta_pdf", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)


            ' DOCUMENTO RECIBIDO FACTURA - Detalle
            oFuncionesB1.creaCampoMD("GS0_NCR", "CodPrin", "(SS) CodPrin", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "CodAuxi", "(SS) CodAuxi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "CodSAP", "(SS) CodSAP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "Descripc", "(SS) Descripc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "Cantid", "(SS) Cantid", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "Precio", "(SS) Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "Desc", "(SS) Desc", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_NCR", "Total", "(SS) Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            
            '   ' DOCUMENTO RECIBIDO FACTURA - Detalle Doc Relacionados
            oFuncionesB1.creaCampoMD("GS1_NCR", "DocEntr", "(SS) DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "LineNu", "(SS) LineNum", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "ItemCode", "(SS) ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "Descripc", "(SS) Dscription", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "Cantid", "(SS) Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "Precio", "(SS) Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "DiscPr", "(SS) DiscPrcnt", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "TaxCode", "(SS) TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "lTotal", "(SS) LineTotal", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_NCR", "ObjType", "(SS) ObjType", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS2_NCR", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS2_NCR", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_NOTA_CREDITO_Catch , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub
    Private Sub CreaUDONotaDeCredito()
        Try
            Dim Child3() As String = {"GS0_NCR", "GS1_NCR", "GS2_NCR"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("GS_NCR", "(SS) NC Recibida", "GS_NCR", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CreaCamposRetencionRecibida()
        Try

            ' DOCUMENTO RECIBIDO FACTURA
            oFuncionesB1.creaCampoMD("GS_RER", "RUC", "(SS) RUC", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "CardCode", "(SS) CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_NCR", "Mapeado", "(SS) Mapeado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "ClaAcc", "(SS) Clave Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "NumAut", "(SS) Numero de Auto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "FecAut", "(SS) Fecha de Auto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "NumDoc", "(SS) Numero de Doc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "FPrelim", "(SS) Factura Prel", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ' DOCUMENTO RECIBIDO FACTURA - Totales
            oFuncionesB1.creaCampoMD("GS_RER", "vTotal", "(SS) Valor Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            ' DOCUMENTO RECIBIDO FACTURA - Totales Documentos Relacionados
            oFuncionesB1.creaCampoMD("GS_RER", "Tipo", "(SS)Tipo Fact Grabd", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "IdGS", "(SS) Id Doc GS", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "Sincro", "(SS) Sincronizado", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "SincroE", "(SS) Sincro EDOC", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal3() As String = {"docPreliminar", "docFinal", "docCancelado", "docMarcado", "docReAbierto", "docPrelXML"}
            Dim StrcbdDes3() As String = {"docPreliminar", "docFinal", "docCancelado", "docMarcado", "docReAbierto", "docPrelXML"}
            oFuncionesB1.creaCampoMD("GS_RER", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            ' CAMPO FECHA QUE SE LO USARÁ PARA EL REPORTE DE DOCUMENTOS INTEGRADOS
            oFuncionesB1.creaCampoMD("GS_RER", "FechaS", "(SS)Fecha DocFinal", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 8, SAPbobsCOM.BoYesNoEnum.tNO)
            'xml y pdf
            oFuncionesB1.creaCampoMD("GS_RER", "Ruta_xml", "(SS) Ruta_xml", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RER", "Ruta_pdf", "(SS) Ruta_pdf", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)

            ' DOCUMENTO RECIBIDO FACTURA - Detalle
            oFuncionesB1.creaCampoMD("GS0_RER", "CodRet", "(SS) CodRetencion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "NumDocRe", "(SS) NumDocRe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "Fecha", "(SS) Fecha", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "pFiscal", "(SS) pFiscal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS0_NCR", "Cantid", "(SS) Cantid", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "Base", "(SS) Base", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "Impuesto", "(SS) Impuesto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "Porcent", "(SS) Porcent", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS0_RER", "valorR", "(SS) valorR", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS1_RER", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS1_RER", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_RETENCION_Catch , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub
    Private Sub CreaUDORetencion()
        Try
            Dim Child3() As String = {"GS0_RER", "GS1_RER"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("GS_RER", "(SS) RE Recibida", "GS_RER", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CreaCamposFacturaXML()
        Try

            ' DOCUMENTO RECIBIDO FACTURA
            oFuncionesB1.creaCampoMD("GS_FC", "NumAut", "(SS) Numero Aut", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "FechaAut", "(SS) Fecha Aut", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "RazSoc", "(SS) Razon Social", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Ruc", "(SS) Ruc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "ClaAcc", "(SS) Clave Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Est", "(SS) Est", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "PuntoEmi", "(SS) Punto Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Sec", "(SS) Secuencial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "FecEmi", "(SS) Fecha Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_FC", "DirEst", "(SS) Direccion Est", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "ConEsp", "(SS) Contri Especial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "RazSocComp", "(SS) Razon Soc Comp", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "IdenComp", "(SS) Ident Comprador", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "DirComp", "(SS) Direccion Comp", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_FC", "TotSinImp", "(SS) Total sin Imp", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "TotDesc", "(SS) Total Desc", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "ImpTotal", "(SS) Importe Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_FC", "FormaPago", "(SS) Forma de Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "TotalPago", "(SS) Total Pago", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "PlazoPago", "(SS) Plazo Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "UniTiempo", "(SS) Unidad Tiempo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base 0
            'oFuncionesB1.creaCampoMD("GS_FC", "Base0", "(SS) Base0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Cod0", "(SS) Codigo0", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodPorc0", "(SS) Cod Porcentaje0", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "BaseImp0", "(SS) Base Imp0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Tarifa0", "(SS) Tarifa0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Valor0", "(SS) Valor0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base 8
            'oFuncionesB1.creaCampoMD("GS_FC", "Base8", "(SS) Base8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Cod8", "(SS) Codigo8", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodPorc8", "(SS) Cod Porcentaje8", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "BaseImp8", "(SS) Base Imp8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Tarifa8", "(SS) Tarifa8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Valor8", "(SS) Valor8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base 12
            'oFuncionesB1.creaCampoMD("GS_FC", "Base12", "(SS) Base12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Cod12", "(SS) Codigo12", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodPorc12", "(SS) Cod Porcentaje12", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "BaseImp12", "(SS) Base Imp12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Tarifa12", "(SS) Tarifa12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "Valor12", "(SS) Valor12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base Noi
            'oFuncionesB1.creaCampoMD("GS_FC", "BaseNoi", "(SS) BaseNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodNoi", "(SS) CodigoNoi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodPorcNoi", "(SS) Cod PorcNoi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "BaseImpNoi", "(SS) Base ImpNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "TarifaNoi", "(SS) TarifaNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "ValorNoi", "(SS) ValorNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base Exen
            'oFuncionesB1.creaCampoMD("GS_FC", "BaseExe", "(SS) BaseExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodExe", "(SS) CodigoExe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodPorcExe", "(SS) Cod PorcExe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "BaseImpExe", "(SS) Base ImpExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "TarifaExe", "(SS) TarifaExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "ValorExe", "(SS) ValorExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base Ice
            'oFuncionesB1.creaCampoMD("GS_FC", "BaseIce", "(SS) BaseIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodIce", "(SS) CodigoIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "CodPorcIce", "(SS) Cod PorcIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "BaseImpIce", "(SS) Base ImpIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "TarifaIce", "(SS) TarifaIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FC", "ValorIce", "(SS) ValorIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdVal3() As String = {"Importado", "Exportado", "Contabilizado", "Cancelado", "Marcado", "Desmarcado"}
            Dim StrcbdDes3() As String = {"Importado", "Exportado", "Contabilizado", "Cancelado", "Marcado", "Desmarcado"}
            oFuncionesB1.creaCampoMD("GS_FC", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            oFuncionesB1.creaCampoMD("GS_FC", "FechaFin", "(SS)Fecha DocFinal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)


            ' DOCUMENTO RECIBIDO FACTURA - Detalle
            oFuncionesB1.creaCampoMD("GS_FCDET", "CodPrin", "(SS) CodPrin", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "CodAuxi", "(SS) CodAuxi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "Descripc", "(SS) Descripc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "Cantid", "(SS) Cantid", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "Precio", "(SS) Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "Desc", "(SS) Desc", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "TotSinImp", "(SS) Total Sin Imp", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            ' DOCUMENTO RECIBIDO FACTURA - Impuestos
            oFuncionesB1.creaCampoMD("GS_FCDET", "Cod", "(SS) Codigo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "CodPorc", "(SS) Cod Porc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "BaseImp", "(SS) Base Imp", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "Tarifa", "(SS) Tarifa", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_FCDET", "CodIce", "(SS) CodigoIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "CodPorcIce", "(SS) Cod PorcIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "BaseImpIce", "(SS) Base ImpIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "TarifaIce", "(SS) TarifaIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_FCDET", "ValorIce", "(SS) ValorIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)


        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_FACTURA , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub

    Private Sub CreaCamposNCreditoXML()
        Try

            ' DOCUMENTO RECIBIDO FACTURA
            oFuncionesB1.creaCampoMD("GS_NC", "NumAut", "(SS) Numero Aut", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "FechaAut", "(SS) Fecha Aut", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "RazSoc", "(SS) Razon Social", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Ruc", "(SS) Ruc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ClaAcc", "(SS) Clave Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Est", "(SS) Est", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "PuntoEmi", "(SS) Punto Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Sec", "(SS) Secuencial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "FecEmi", "(SS) Fecha Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_NC", "DirEst", "(SS) Direccion Est", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ConEsp", "(SS) Contri Especial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "RazSocComp", "(SS) Razon Soc Comp", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "IdenComp", "(SS) Ident Comprador", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "DirComp", "(SS) Direccion Comp", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_NC", "TotSinImp", "(SS) Total sin Imp", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "TotDesc", "(SS) Total Desc", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ImpTotal", "(SS) Importe Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'oFuncionesB1.creaCampoMD("GS_NC", "FormaPago", "(SS) Forma de Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_NC", "TotalPago", "(SS) Total Pago", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_NC", "PlazoPago", "(SS) Plazo Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("GS_NC", "UniTiempo", "(SS) Unidad Tiempo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_NC", "CodDocMod", "(SS) Cod Doc Mod", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "NumDocMod", "(SS) Num Doc Mod", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "FecDocMod", "(SS) Fecha Doc Mod", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ValorMod", "(SS) Valor Modificacion", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Motivo", "(SS) Motivo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base 0
            'oFuncionesB1.creaCampoMD("GS_FC", "Base0", "(SS) Base0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Cod0", "(SS) Codigo0", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodPorc0", "(SS) Cod Porcentaje0", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "BaseImp0", "(SS) Base Imp0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Tarifa0", "(SS) Tarifa0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Valor0", "(SS) Valor0", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base 8
            'oFuncionesB1.creaCampoMD("GS_FC", "Base8", "(SS) Base8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Cod8", "(SS) Codigo8", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodPorc8", "(SS) Cod Porcentaje8", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "BaseImp8", "(SS) Base Imp8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Tarifa8", "(SS) Tarifa8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Valor8", "(SS) Valor8", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base 12
            'oFuncionesB1.creaCampoMD("GS_FC", "Base12", "(SS) Base12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Cod12", "(SS) Codigo12", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodPorc12", "(SS) Cod Porcentaje12", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "BaseImp12", "(SS) Base Imp12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Tarifa12", "(SS) Tarifa12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "Valor12", "(SS) Valor12", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base Noi
            'oFuncionesB1.creaCampoMD("GS_FC", "BaseNoi", "(SS) BaseNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodNoi", "(SS) CodigoNoi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodPorcNoi", "(SS) Cod PorcNoi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "BaseImpNoi", "(SS) Base ImpNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "TarifaNoi", "(SS) TarifaNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ValorNoi", "(SS) ValorNoi", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base Exen
            'oFuncionesB1.creaCampoMD("GS_FC", "BaseExe", "(SS) BaseExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodExe", "(SS) CodigoExe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodPorcExe", "(SS) Cod PorcExe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "BaseImpExe", "(SS) Base ImpExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "TarifaExe", "(SS) TarifaExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ValorExe", "(SS) ValorExe", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            'IMPUESTOS CABECERA Base Ice
            'oFuncionesB1.creaCampoMD("GS_FC", "BaseIce", "(SS) BaseIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodIce", "(SS) CodigoIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "CodPorcIce", "(SS) Cod PorcIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "BaseImpIce", "(SS) Base ImpIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "TarifaIce", "(SS) TarifaIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NC", "ValorIce", "(SS) ValorIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdVal3() As String = {"Importado", "Exportado", "Contabilizado", "Cancelado", "Marcado", "Desmarcado"}
            Dim StrcbdDes3() As String = {"Importado", "Exportado", "Contabilizado", "Cancelado", "Marcado", "Desmarcado"}
            oFuncionesB1.creaCampoMD("GS_NC", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            oFuncionesB1.creaCampoMD("GS_NC", "FechaFin", "(SS)Fecha DocFinal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)


            ' DOCUMENTO RECIBIDO FACTURA - Detalle
            oFuncionesB1.creaCampoMD("GS_NCDET", "CodPrin", "(SS) CodPrin", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "CodAuxi", "(SS) CodAuxi", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "Descripc", "(SS) Descripc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "Cantid", "(SS) Cantid", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "Precio", "(SS) Precio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "Desc", "(SS) Desc", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "TotSinImp", "(SS) Total Sin Imp", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            ' DOCUMENTO RECIBIDO FACTURA - Impuestos
            oFuncionesB1.creaCampoMD("GS_NCDET", "Cod", "(SS) Codigo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "CodPorc", "(SS) Cod Porc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "BaseImp", "(SS) Base Imp", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "Tarifa", "(SS) Tarifa", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_NCDET", "CodIce", "(SS) CodigoIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "CodPorcIce", "(SS) Cod PorcIce", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "BaseImpIce", "(SS) Base ImpIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "TarifaIce", "(SS) TarifaIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_NCDET", "ValorIce", "(SS) ValorIce", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)


        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_nota credito , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub

    Private Sub CreaCamposRetencionXML()
        Try

            ' DOCUMENTO RECIBIDO retencion
            oFuncionesB1.creaCampoMD("GS_RT", "NumAut", "(SS) Numero Aut", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "FechaAut", "(SS) Fecha Aut", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "RazSoc", "(SS) Razon Social", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "Ruc", "(SS) Ruc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "ClaAcc", "(SS) Clave Acceso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "Est", "(SS) Est", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "PuntoEmi", "(SS) Punto Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "Sec", "(SS) Secuencial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "FecEmi", "(SS) Fecha Emision", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("GS_RT", "DirEst", "(SS) Direccion Est", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "ConEsp", "(SS) Contri Especial", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "RazSocRet", "(SS) Razon Social Ret", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "IdenRet", "(SS) Identificacion Ret", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RT", "Periodo", "(SS) Periodo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, SAPbobsCOM.BoYesNoEnum.tNO)

            Dim StrcbdVal3() As String = {"Importado", "Exportado", "Contabilizado", "Cancelado", "Marcado", "Desmarcado"}
            Dim StrcbdDes3() As String = {"Importado", "Exportado", "Contabilizado", "Cancelado", "Marcado", "Desmarcado"}
            oFuncionesB1.creaCampoMD("GS_RT", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            oFuncionesB1.creaCampoMD("GS_RT", "FechaFin", "(SS)Fecha DocFinal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)


            ' DOCUMENTO RECIBIDO FACTURA - Detalle
            oFuncionesB1.creaCampoMD("GS_RTDET", "Codigo", "(SS) Codigo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "CodRet", "(SS) Codigo Retencion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "BaseImp", "(SS) Base Imponible", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "PorcRet", "(SS) Porcentaje Ret", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "ValorRet", "(SS) Valor Retenido", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "CodDocSus", "(SS) Cod Doc Sustento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "NumDocSus", "(SS) Num Doc Sustento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("GS_RTDET", "FemiDocSus", "(SS) Fecha Emi Doc Sus", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO)


        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_nota credito , Error: " & ex.Message.ToString(), "Estructura")
        End Try
    End Sub

    Private Sub CreaUDOFacturaXML()
        Try
            Dim Child3() As String = {"GS_FCDET"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("GS_FC", "(SS) FACTURA CABECERA", "GS_FC", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CreaUDONCreditoXML()
        Try
            Dim Child3() As String = {"GS_NCDET"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("GS_NC", "(SS) NCREDITO CABECERA", "GS_NC", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CreaUDORetencionXML()
        Try
            Dim Child3() As String = {"GS_RTDET"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("GS_RT", "(SS) RETENCION CABECERA", "GS_RT", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CrearBusquedasFormateadas()
        Dim N_Categoria As String = "SOLSAP_FE"
        oFuncionesB1.creaQueryCat(N_Categoria)
        oFuncionesB1.creaQuery("GetIndicadoresImpuestos", My.Resources.GetIndicadoresImpuestos, N_Categoria, False)
        oFuncionesB1.creaQuery("GetTarjetasCredito", My.Resources.GetTartejaCredito, N_Categoria, False)
        oFuncionesB1.creaQuery("GetCuentaServicio", My.Resources.GetCuentasServicios, N_Categoria, False)

    End Sub

    Private Sub CrearBusquedasFormateadas(Nombrebf As String, Query As String, TABLAS_FORMID As List(Of String), AliasCampo As String, ItemId As String, Optional columnaid As String = "")

        'MEDIOS PAGO

        'Dim N_Categoria As String = "BF_SOLSAP"
        Dim N_Query As String = Nombrebf
        Dim idcat As Integer = 0
        Dim idquery As Integer = 0
        Dim idcampo As Integer = 0

        'oFuncionesB1.creaQueryCat(N_Categoria)
        oFuncionesB1.creaQuery(N_Query, Query, N_Categoria, False)

        If Not String.IsNullOrWhiteSpace(AliasCampo) And Not String.IsNullOrWhiteSpace(ItemId) Then

            idcat = oFuncionesB1.getIdQueryCat(N_Categoria)
            Utilitario.Util_Log.Escribir_Log("idcat:  " & idcat.ToString(), "FuncionesB1")

            If idcat <> -1 Then

                idquery = oFuncionesB1.getIdQuery(N_Query, idcat)
                Utilitario.Util_Log.Escribir_Log("idquery:  " & idquery.ToString(), "FuncionesB1")

                If idquery <> -1 Then

                    For Each dato As String In TABLAS_FORMID

                        idcampo = oFuncionesB1.getIdUserField(dato.Split("-")(0), AliasCampo)
                        Utilitario.Util_Log.Escribir_Log("idcampo:  " & idcampo.ToString(), "FuncionesB1")

                        If idcampo <> -1 Then

                            Dim fUserBusFor2 As SAPbobsCOM.FormattedSearches
                            fUserBusFor2 = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                            Dim existe As Boolean = False
                            If columnaid = "" Then
                                existe = fUserBusFor2.GetByKey(oFuncionesB1.getIdBusquedaF(dato.Split("-")(1), ItemId))
                            Else

                                existe = fUserBusFor2.GetByKey(oFuncionesB1.getIdBusquedaF(dato.Split("-")(1), ItemId, columnaid))
                            End If
                            Utilitario.Util_Log.Escribir_Log("existe:  " & existe.ToString(), "FuncionesB1")

                            If existe Then
                                fUserBusFor2.Remove()
                            End If
                            oFuncionesB1.Release(fUserBusFor2)

                            Dim FMS As SAPbobsCOM.FormattedSearches = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                            FMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery

                            FMS.Refresh = SAPbobsCOM.BoYesNoEnum.tYES
                            FMS.ByField = SAPbobsCOM.BoYesNoEnum.tYES
                            FMS.FormID = dato.Split("-")(1) ' Para el formulario de Factura de Deudores
                            FMS.FieldID = CStr(idcampo)
                            FMS.QueryID = idquery
                            FMS.ItemID = ItemId
                            If columnaid <> "" Then
                                FMS.ColumnID = columnaid
                            End If

                            'FMS.ColumnID=
                            Dim ms As String = ""
                            Dim ret As Integer = FMS.Add
                            If ret <> 0 Then
                                rCompany.GetLastError(ret, ms)
                                Utilitario.Util_Log.Escribir_Log(" - Ocurrio un Error al Asociar la Busqueda Formateada en el campo= " + AliasCampo + " tabla= " + dato.Split("-")(0) + " : " & rCompany.GetLastErrorDescription + " ms:" + ms.ToString(), "FuncionesB1")
                                rSboApp.SetStatusBarMessage(NombreAddon + " - Ocurrio un Error al Asociar la Busqueda Formateada en el campo= " + AliasCampo + " tabla= " + dato.Split("-")(0) + " : " & rCompany.GetLastErrorDescription)
                            Else
                                rSboApp.SetStatusBarMessage(NombreAddon + " - Busqueda Formateada Asociada Correctamente al campo= " + AliasCampo + " tabla= " + dato.Split("-")(0), SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                            End If

                            'If FMS.Add <> 0 Then

                            '    rSboApp.SetStatusBarMessage(NombreAddon + " - Ocurrio un Error al Asociar la Busqueda Formateada en el campo= " + AliasCampo + " tabla= " + dato.Split("-")(0) + " : " & rCompany.GetLastErrorDescription)
                            'Else
                            '    rSboApp.SetStatusBarMessage(NombreAddon + " - Busqueda Formateada Asociada Correctamente al campo= " + AliasCampo + " tabla= " + dato.Split("-")(0), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            'End If

                        End If
                    Next
                End If
            End If
        End If

    End Sub

    'Add 15/07/2024
    Public Sub CreacionEstructuraCM()
        Try
            oFuncionesB1.creaTablaMD("SS_PM_CAB", "(SS) Pagos masivos", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SS_PM_DET1", "(SS) Detalle pagos masivos", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oFuncionesB1.creaTablaMD("SS_PAG_PERMISOS", "(SS) Pagos masivos permisos", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oFuncionesB1.creaTablaMD("SS_PAG_CC3PRPFC", "(SS) Relacion Suc Pro Par", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oFuncionesB1.creaTablaMD("SS_PM_CONTROLCHEQUE", "(SS) Control de num cheques PM", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            oFuncionesB1.creaTablaMD("SS_PM_OP_CAB", "(SS) Ord Pag PM CAB", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SS_PM_OP_DET1", "(SS) Ord Pag PM DET1 ", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            Dim StrcbdVal1() As String = {"Revision", "Aprobado", "Rechazado", "Edicion", "Procesado", "Modificar", "Archivo Generado", "Archivo Procesado Banco"}
            Dim StrcbdDes1() As String = {"Revision", "Aprobado", "Rechazado", "Edicion", "Procesado", "Modificar", "Archivo Generado", "Archivo Procesado Banco"}
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal1, StrcbdDes1, "")
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "NivelAprob", "(SS) NivelAprob", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "Cuenta", "(SS) Cuenta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal2() As String = {"Cheque", "Transferencia", "Servicios Basicos"}
            Dim StrcbdDes2() As String = {"Cheque", "Transferencia", "Servicios Basicos"}
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "MedioPago", "(SS) MedioPago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal2, StrcbdDes2, "")
            Dim StrcbdVal3() As String = {"Standard", "Nomina"}
            Dim StrcbdDes3() As String = {"Standard", "Nomina"}
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "Tipo", "(SS) Tipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal3, StrcbdDes3, "")
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "TotalPagado", "(SS) TotalPagado", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "FacProcesadas", "(SS) FacProcesadas", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "Banco", "(SS) Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "IdCashBan", "(SS) Id Cash Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "RutArcBan", "(SS) Ruta Archivo de banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "RutArcGen", "(SS) Ruta Archivo generado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "IdPagCon", "(SS) Id Pago Consolidado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "FechaArcRec", "(SS) Fecha Archivo Rec", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "FechaDev", "(SS) Fecha Devolucion", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_PM_DET1", "CodProv", "(SS) CodProv", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Proveedor", "(SS) Proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Vencimiento", "(SS) Vencimiento", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "FechaVen", "(SS) FechaVen", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Monto", "(SS) Monto", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Saldo", "(SS) Saldo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Pago", "(SS) Pago", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "DocEntry", "(SS) DocEntry FP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Cuota", "(SS) Cuota", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "ObjType", "(SS) Tipo obj", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "NumDoc", "(SS) Numero doc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "NumLinea", "(SS) Num Linea Base", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Sucursal", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Proyecto", "(SS) Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "CtaBcoPr", "(SS) Cta Bco Prov", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "NumChe", "(SS) Número de ch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Procesada", "(SS) Procesada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Comentario", "(SS) Comentario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "BcoPr", "(SS) Banco Prov", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "TipCtaPr", "(SS) Tipo cuenta prov", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "IdPagTran", "(SS) Id Pago Transitorio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "IdNotDeb", "(SS) Id Nota Debito", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "ComentarioFac", "(SS) Comentario Factura", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Consolidado", "(SS) Consolidado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "Impreso", "(SS) Impreso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "IdOrdPag", "(SS) Orden Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_PAG_PERMISOS", "Usuario", "(SS) Usuario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAG_PERMISOS", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal4() As String = {"0", "1", "2", "3", "4", "5"}
            Dim StrcbdDes4() As String = {"Agente", "Aprobacion1", "Aprobacion2", "Aprobacion3", "Aprobacion4", "Aprobacion5"}
            oFuncionesB1.creaCampoMD("SS_PAG_PERMISOS", "Nivel", "(SS) Nivel", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal4, StrcbdDes4)

            oFuncionesB1.creaCampoMD("SS_PAG_CC3PRPFC", "Cod_Sucursal", "(SS) Código Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAG_CC3PRPFC", "Nombre_Sucursal", "(SS) Nombre Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAG_CC3PRPFC", "Cod_Proyecto", "(SS) Código Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAG_CC3PRPFC", "Nom_Proyecto", "(SS) Nombre Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAG_CC3PRPFC", "Cod_Partida", "(SS) Código Partida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PAG_CC3PRPFC", "Nom_Partida", "(SS) Nombre Partida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_PM_CONTROLCHEQUE", "NumChe", "(SS) N° Cheque Sig", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            'ADD Seccion archivo CashManagement
            oFuncionesB1.creaTablaMD("SS_CM_CAB", "(SS) Archivo Cabecera", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SS_CM_DET1", "(SS) Archivo Detalle", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oFuncionesB1.creaTablaMD("SS_CM_MBCOCTA", "(SS) Mapeo Bco-Cta", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            oFuncionesB1.creaCampoMD("SS_CM_CAB", "FecArc", "(SS) Fecha Archivo", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_CAB", "RutaArc", "(SS) Ruta Archivo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_CAB", "NumPagos", "(SS) Número de pagos", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_CAB", "TotPagos", "(SS) Total de pagos", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_CAB", "Banco", "(SS) Banco/Cuenta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal5() As String = {"Abierto", "Cerrado"}
            Dim StrcbdDes5() As String = {"Abierto", "Cerrado"}
            oFuncionesB1.creaCampoMD("SS_CM_CAB", "Estado", "(SS) Estado CM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal5, StrcbdDes5) 'add 30/09/2024
            oFuncionesB1.creaCampoMD("SS_CM_CAB", "IdCashMan", "(SS) Id Cash Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_CM_DET1", "DocEntryP", "(SS) DocEntry Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "CodProv", "(SS) Código proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "FecPag", "(SS) Fecha Pago", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "TotPag", "(SS) Total Pagado", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "ForPag", "(SS) Forma Pago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_CM_DET1", "Moneda", "(SS) Moneda", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "TipCta", "(SS) Tipo cuenta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "NumCta", "(SS) Numero cuenta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "Ref", "(SS) Referencia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "TipCli", "(SS) Tipo Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "NumID", "(SS) Tipo Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "CodBco", "(SS) Codigo Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "CodCtaEm", "(SS) Cod Cuenta Empresa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "CodBcoEm", "(SS) Cod Banco Empresa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "Email", "(SS) Email", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_DET1", "FacRef", "(SS) Factura Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_CM_MBCOCTA", "CodBco", "(SS) Cod Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_MBCOCTA", "NomBco", "(SS) Nom Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_MBCOCTA", "CtaSys", "(SS) Cta Sistema", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_MBCOCTA", "NomCtaSys", "(SS) Nom Cta Sist", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CM_MBCOCTA", "CtaBco", "(SS) Cta Bco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)


            oFuncionesB1.creaCampoMD("OVPM", "UDO_CM", "(SS) Id Archivo CM", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OVPM", "DE_PM", "(SS) DE Solicitud PM", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            'Seccion destinada para campos de integracion con Odoo
            oFuncionesB1.creaCampoMD("OHEM", "ID_NOMINA", "(SS) Id Nomina", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OHEM", "NOMINA_LAST_SYNC", "(SS) Nomina Last Sync Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OHEM", "NOMINA_LAST_SYNC_HORA", "(SS) Nomina Last Sync Time", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_PM_CAB", "ID_NOMINA", "(SS) Id Nomina", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "NOMINA_LAST_SYNC", "(SS) Nomina Last Sync Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_CAB", "NOMINA_LAST_SYNC_HORA", "(SS) Nomina Last Sync Time", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("OJDT", "ID_NOMINA", "(SS) Id Nomina", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OJDT", "NOMINA_LAST_SYNC", "(SS) Nomina Last Sync Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("OJDT", "NOMINA_LAST_SYNC_HORA", "(SS) Nomina Last Sync Time", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, 11, SAPbobsCOM.BoYesNoEnum.tNO)


            'NEW JP 06022025
            oFuncionesB1.creaCampoMD("SS_PM_DET1", "IdOrdPag", "(SS) Id Ord Pag", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "NivelAprob", "(SS) NivelAprob", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "Cuenta", "(SS) Cuenta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "MedioPago", "(SS) MedioPago", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "Tipo", "(SS) Tipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "TotalPagado", "(SS) TotalPagado", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "FacProcesadas", "(SS) FacProcesadas", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "Banco", "(SS) Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "IdCashBan", "(SS) Id Cash Banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "RutArcBan", "(SS) Ruta Archivo de banco", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "RutArcGen", "(SS) Ruta Archivo generado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "IdPagCon", "(SS) Id Pago Consolidado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "FechaArcRec", "(SS) Fecha Archivo Rec", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "FechaDev", "(SS) Fecha Devolucion", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_CAB", "SolicitudPago", "(SS) Solicitud PM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "CodProv", "(SS) CodProv", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Proveedor", "(SS) Proveedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Vencimiento", "(SS) Vencimiento", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "FechaVen", "(SS) FechaVen", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Monto", "(SS) Monto", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Saldo", "(SS) Saldo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Pago", "(SS) Pago", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "DocEntry", "(SS) DocEntry FP", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Cuota", "(SS) Cuota", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "ObjType", "(SS) Tipo obj", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "NumDoc", "(SS) Numero doc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "NumLinea", "(SS) Num Linea Base", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Sucursal", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Proyecto", "(SS) Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "CtaBcoPr", "(SS) Cta Bco Prov", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "NumChe", "(SS) Número de ch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Procesada", "(SS) Procesada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Comentario", "(SS) Comentario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "BcoPr", "(SS) Banco Prov", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "TipCtaPr", "(SS) Tipo cuenta prov", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "IdPagTran", "(SS) Id Pago Transitorio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "IdNotDeb", "(SS) Id Nota Debito", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "ComentarioFac", "(SS) Comentario Factura", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Consolidado", "(SS) Consolidado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_PM_OP_DET1", "Impreso", "(SS) Impreso", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            CreaUDOOrdenPago()
            '

            CreaUDOPagoMasivo()
            CreaUDOCashManagement()

        Catch ex As Exception
            rSboApp.StatusBar.SetText(NombreAddon + " - Error CreacionEstructuraCM " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    'Add 15/07/2024
    Private Sub CreaUDOPagoMasivo()
        Try
            Dim Child3() As String = {"SS_PM_DET1"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("SSMTPAGOS", "(SS) Pagos Masivos", "SS_PM_CAB", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            rSboApp.StatusBar.SetText(NombreAddon + " - Error CreaUDOPagoMasivo " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub CreaUDOOrdenPago()
        Try
            Dim Child3() As String = {"SS_PM_OP_DET1"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("SSOPPMPAGOS", "(SS) Orden Pago PM", "SS_PM_OP_CAB", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            rSboApp.StatusBar.SetText(NombreAddon + " - Error CreaUDOOrdenPago " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub CreaUDOCashManagement()
        Try
            Dim Child3() As String = {"SS_CM_DET1"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("SSCMCASH", "SS Cash Management", "SS_CM_CAB", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            rSboApp.StatusBar.SetText(NombreAddon + " - Error CreaUDOCashManagement " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub crearUdoRegistradoRetencionRecibida()


        Dim oUdo As SAPbobsCOM.UserObjectsMD
        'Dim oUDOEnhancedForm As SAPbobsCOM.UserObjectMD_EnhancedFormColumns
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("GS_RER") Then
                oUdo.Code = "GS_RER"
                oUdo.Name = "(SS) Retención Recibida"
                oUdo.TableName = "GS_RER"
                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_GS_RER"
                'Busqueda para Choosefrom List y Poder vincular Udo a UDT

                'BUSQUEDA
                oUdo.FindColumns.ColumnAlias = "DocEntry"
                oUdo.FindColumns.ColumnDescription = "DocEntry"
                oUdo.FindColumns.Add()

                'CABECERA

                oUdo.FormColumns.FormColumnAlias = "DocEntry"
                oUdo.FormColumns.FormColumnDescription = "DocEntry"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_Nombre"
                oUdo.FormColumns.FormColumnDescription = "(SS) Nombre"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_RUC"
                oUdo.FormColumns.FormColumnDescription = "(SS) Ruc"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_CardCode"
                oUdo.FormColumns.FormColumnDescription = "(SS) Codigo Cliente"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_ClaAcc"
                oUdo.FormColumns.FormColumnDescription = "DocEntry"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_NumAut"
                oUdo.FormColumns.FormColumnDescription = "(SS) Numero de Aut."
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_FecAut"
                oUdo.FormColumns.FormColumnDescription = "SS) Fecha Aut."
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_NumDoc"
                oUdo.FormColumns.FormColumnDescription = "(SS) Numero de Doc"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_Estado"
                oUdo.FormColumns.FormColumnDescription = "(SS) Estado"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_vTotal"
                oUdo.FormColumns.FormColumnDescription = "(SS) Valor Total"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_IdGS"
                oUdo.FormColumns.FormColumnDescription = "(SS) Id Doc GS"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_Sincro"
                oUdo.FormColumns.FormColumnDescription = "(SS) Sincronizado"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()

                oUdo.FormColumns.FormColumnAlias = "U_SincroE"
                oUdo.FormColumns.FormColumnDescription = "(SS) Sincro EDOC"
                oUdo.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.FormColumns.Add()


                'DETALLES

                oUdo.ChildTables.TableName = "GS0_RER"

                'tablas hijas con formulario nuevo

                oUdo.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.EnhancedFormColumns.ColumnAlias = "U_CodRet"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) CodRetencion"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 1
                oUdo.EnhancedFormColumns.Add()


                oUdo.EnhancedFormColumns.ColumnAlias = "U_NumDocRe"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) NumDocRe"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 2
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_Fecha"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Fecha"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 3
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_Base"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Base"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 4
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_Impuesto"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Impuesto"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 5
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_Porcent"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Porcentaje"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 6
                oUdo.EnhancedFormColumns.Add()

                oUdo.EnhancedFormColumns.ColumnAlias = "U_valorR"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) valorR"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 1
                oUdo.EnhancedFormColumns.ColumnNumber = 7
                oUdo.EnhancedFormColumns.Add()

                oUdo.ChildTables.TableName = "GS1_RER"

                oUdo.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES

                oUdo.EnhancedFormColumns.ColumnAlias = "U_Nombre"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Nombre"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 2
                oUdo.EnhancedFormColumns.ColumnNumber = 1
                oUdo.EnhancedFormColumns.Add()


                oUdo.EnhancedFormColumns.ColumnAlias = "U_Valor"
                oUdo.EnhancedFormColumns.ColumnDescription = "(SS) Valor"
                oUdo.EnhancedFormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.EnhancedFormColumns.ChildNumber = 2
                oUdo.EnhancedFormColumns.ColumnNumber = 2
                oUdo.EnhancedFormColumns.Add()

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_RetencionRecibida , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO FUN_CreaUDO_RetencionRecibida " + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : GS_RER, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_SER_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try


    End Sub


    'Add JP 13/11/2024
    Public Sub CreacionEstructuraSB()
        Try
            rSboApp.StatusBar.SetText(NombreAddon + " - Iniciando creacion de estructura!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oFuncionesB1.creaTablaMD("SS_SB_CAB", "(SS) Servicios Basicos cab", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SS_SB_DET1", "(SS) Servicios Basicos det1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oFuncionesB1.creaTablaMD("SS_SB_DET2", "(SS) Servicios Basicos det2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oFuncionesB1.creaTablaMD("SS_SB_DET3", "(SS) Servicios Basicos det3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oFuncionesB1.creaTablaMD("SS_SB_MED_ADM", "(SS) Medidores Administrativos", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oFuncionesB1.creaTablaMD("SS_SB_USR_SUC", "(SS) Mapeo usuario-sucursal", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            oFuncionesB1.creaCampoMD("SS_SB_CAB", "IniPer", "(SS) Inicio Periodo", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "FinPer", "(SS) Fin Periodo", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "NivCC3", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "TipSer", "(SS) Tipo Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "UM", "(SS) Unidad Medida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "Factor", "(SS) Factor", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "Concepto", "(SS) Concepto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdValest() As String = {"Borrador", "Aprobado"}
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "Estado", "(SS) Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValest, StrcbdValest, "")
            oFuncionesB1.creaCampoMD("SS_SB_CAB", "Ruta", "(SS) Ruta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Contrato", "(SS) Contrato", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Denominacion", "(SS) Denominacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Locales", "(SS) Locales", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Nivel", "(SS) Nivel", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "LecIni", "(SS) Lectura Inicial", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "LecFin", "(SS) Lectura Final", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Consumo", "(SS) Consumo", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Factor", "(SS) Factor", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Costo", "(SS) Costo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Total", "(SS) Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Observacion", "(SS) Observacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "SujetoFact", "(SS) SujetoFact", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "KilCon", "(SS) KilCon", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 250, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_SB_DET1", "Facturado", "(SS) Facturado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "FechaFact", "(SS) Fecha Fact", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 49, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "NumeroFac", "(SS) Numero Fact", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET1", "EntryFac", "(SS) Entry Fact", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_SB_DET2", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET2", "Consumo", "(SS) Consumo", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET2", "Costo", "(SS) Costo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET2", "DocEntryFac", "(SS) DocEntryFac", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_SB_DET3", "Nivel", "(SS) Nivel", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET3", "Consumo", "(SS) Consumo", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET3", "Costo", "(SS) Costo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET3", "Total", "(SS) Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_DET3", "Porcentaje", "(SS) Porcentaje", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 11, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SS_SB_MED_ADM", "Contrato", "(SS) Contrato", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_MED_ADM", "Denominacion", "(SS) Denominacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_MED_ADM", "Sucursal", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_MED_ADM", "LecIni", "(SS) Lectura Inicial", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdValets() As String = {"Gas", "Agua", "Energia"}
            oFuncionesB1.creaCampoMD("SS_SB_MED_ADM", "TipSer", "(SS) Tipo Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValets, StrcbdValets, "")

            oFuncionesB1.creaCampoMD("SS_SB_USR_SUC", "Usuario", "(SS) Usuario", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_USR_SUC", "Sucursal", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_USR_SUC", "DesSuc", "(SS) Desc sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_SB_USR_SUC", "TipSer", "(SS) Tipo Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValets, StrcbdValets, "")

            CreaUDOServiciosBasicos()

            oFuncionesB1.creaCampoMD("OCRD", "TipoServicio", "(SS) Tipo Servicio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdValets, StrcbdValets, "")

            oFuncionesB1.creaCampoMD("OINV", "SS_ServicioBasico", "(SS) Servicio Basico", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO)

            rSboApp.StatusBar.SetText(NombreAddon + " - Creacion de estructura realizada con éxito!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            rSboApp.StatusBar.SetText(NombreAddon + " - Error CreacionEstructuraSB " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub CreaUDOServiciosBasicos()
        Try
            Dim Child3() As String = {"SS_SB_DET1", "SS_SB_DET2", "SS_SB_DET3"}
            Dim find3() As String = {}
            oFuncionesB1.creaUDOC("SSSB", "(SS) Servicios Basicos", "SS_SB_CAB", find3, Child3, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            rSboApp.StatusBar.SetText(NombreAddon + " - Error CreaUDOServiciosBasicos " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class



'Imports Entidades
Imports Microsoft.Office.Interop.Excel
Public Class FuncionesAddon

    Dim mCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company
    Private fCompany As SAPbobsCOM.Company
    Private fPappl As SAPbouiCOM.Application

    Public mostrarMensajesExito As Boolean = False
    Public mostrarMensajesError As Boolean = False

    Public NombreAddon As String = ""

    Private oForm As SAPbouiCOM.Form
    Public Sub New(ByVal objetfCompany As SAPbobsCOM.Company, ByVal objectApplication As SAPbouiCOM.Application, Optional ByVal mostrarErrores As Boolean = True, Optional ByVal mostrarExito As Boolean = False, Optional ByVal sNombreAddon As String = "")
        Try
            fCompany = objetfCompany
            fPappl = objectApplication
            mostrarMensajesError = mostrarErrores
            mostrarMensajesExito = mostrarExito
            NombreAddon = sNombreAddon

        Catch ex As Exception
        Finally
            If mostrarMensajesExito Then
                If isInSpanish() Then
                    fPappl.StatusBar.SetText("Funciones Addon Iniciado", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    fPappl.StatusBar.SetText("Addon Functions connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If
        End Try
    End Sub

    Public Function isInSpanish() As Boolean
        Dim miBool As Boolean = False
        Try
            If fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish Or fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish_Ar Or fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish_La Or fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish_Pa Then miBool = True
        Catch ex As Exception
        End Try
        Return miBool
    End Function

    Private Sub CreateModalForm(ByVal Texto As String)
        Dim cp As SAPbouiCOM.FormCreationParams
        Dim oItem As SAPbouiCOM.Item
        Dim oStatic As SAPbouiCOM.StaticText
#Disable Warning BC42024 ' Variable local sin usar: 'oPicture'.
        Dim oPicture As SAPbouiCOM.PictureBox
#Enable Warning BC42024 ' Variable local sin usar: 'oPicture'.

        ' Create the form
        cp = fPappl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)

        cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed
        cp.FormType = "Modal"
        cp.UniqueID = "Modal"

        oForm = fPappl.Forms.AddEx(cp)
        oForm.ClientHeight = 130
        oForm.ClientWidth = 300

        oForm.Left = 350
        oForm.Top = 200

        ' Create the form GUI elements
        oForm.AutoManaged = False
        oForm.SupportedModes = 0
        'oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        'oItem.AffectsFormMode = False
        'oItem.Top = 100
        'oItem.Left = 10

        oItem = oForm.Items.Add("MyStatic", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oStatic = oItem.Specific()
        oStatic.Caption = Texto
        oStatic.Item.Width = 300

        'Dim strPath As String
        'strPath = System.Windows.Forms.Application.StartupPath & "\prueba.png"

        'oPicture = oForm.Items.Add("ImgLogo", SAPbouiCOM.BoFormItemTypes.it_PICTURE)
        ''oItem.AffectsFormMode = False        
        'oPicture.Picture = strPath

        oForm.Visible = True
        ' bModal = True
    End Sub

    Public Sub CloseModalForm(ByVal formulario As String)
        Try
            For Each oForm In fPappl.Forms
                Select Case oForm.UniqueID
                    Case formulario
                        oForm.Select()
                        oForm.Close()
                        ' bModal = False
                End Select
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Function ActualizaSecuencia(ByVal code As String, ByVal UltimaSecuencia As Integer, ByVal DocEntry As String, ByVal Tipotabla As String, ByVal Transaccion As String, ByVal TipoLog As String) As Boolean
        Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia (antes del try)", "ManejoDeDocumentos")
        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            Dim ActualizaSecuenc As Boolean = True
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
            Dim oUserTable As SAPbobsCOM.UserTable = Nothing
            GC.Collect()
            oUserObjectMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = fCompany.GetCompanyService
            Utilitario.Util_Log.Escribir_Log("antes del if ", "ManejoDeDocumentos")
            If oUserObjectMD.GetByKey("EXX_DOCUM_LEG_INTER") Then ' PREGUNTO SI ES UN UDO, YA QUE ALGUNOS CLIENTES NO TIENEN REGISTRADO EL UDO
                GuardaLOG(Tipotabla, DocEntry, "'EXX_DOCUM_LEG_INTER' es un UDO: ", Transaccion, TipoLog)
                oGeneralService = sCmp.GetGeneralService("EXX_DOCUM_LEG_INTER")
                Utilitario.Util_Log.Escribir_Log("IEXX_DOCUM_LEG_INTER oGeneralService", "ManejoDeDocumentos")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("Code", code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'EXX_DOCUM_LEG_INTER' por el Code: " + code.ToString(), Transaccion, TipoLog)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                Utilitario.Util_Log.Escribir_Log("oGeneralData error", "ManejoDeDocumentos")
                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oGeneralData.GetProperty("U_ULT_SECUEN").ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oGeneralData.GetProperty("U_ULT_SECUEN").ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception
                    GuardaLOG(Tipotabla, DocEntry, "Cath Ultima Secuencia: " + ex.Message.ToString(), Transaccion, TipoLog)
                End Try

                If ActualizaSecuenc Then
                    oGeneralData.SetProperty("U_ULT_SECUEN", UltimaSecuencia)
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'EXX_DOCUM_LEG_INTER' en el campo 'U_ULT_SECUEN' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    oGeneralService.Update(oGeneralData)
                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If


                Return True
            Else
                GuardaLOG(Tipotabla, DocEntry, "'EXX_DOCUM_LEG_INTER' es una TABLA DE USUARIO: ", Transaccion, TipoLog)
                oUserTable = fCompany.UserTables.Item("EXX_DOCUM_LEG_INTER")
                oUserTable.GetByKey(code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'EXX_DOCUM_LEG_INTER' por el Code: " + code.ToString(), Transaccion, TipoLog)

                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception

                End Try

                If ActualizaSecuenc Then
                    oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value = UltimaSecuencia
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'EXX_DOCUM_LEG_INTER' en el campo 'U_ULT_SECUEN' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    'oUserTable.Update()
                    RetVal = oUserTable.Update()
                    If RetVal <> 0 Then
                        'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        fCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        GuardaLOG(Tipotabla, DocEntry, "ERROR en 'EXX_DOCUM_LEG_INTER' al actualizar el campo 'U_ULT_SECUEN' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                    End If


                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If

                Return True
            End If


        Catch ex As Exception
            fPappl.SetStatusBarMessage(NombreAddon + "Error al actualizar la secuencia", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia" + ex.Message.ToString(), Transaccion, TipoLog)
            ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
            Return False
        End Try

    End Function

    Public Function ActualizaSecuenciaSS(ByVal code As String, ByVal UltimaSecuencia As Integer, ByVal DocEntry As String, ByVal Tipotabla As String, ByVal Transaccion As String, ByVal TipoLog As String) As Boolean
        Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia (antes del try)", "ManejoDeDocumentos")
        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            Dim ActualizaSecuenc As Boolean = True
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
            Dim oUserTable As SAPbobsCOM.UserTable = Nothing
            GC.Collect()
            oUserObjectMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = fCompany.GetCompanyService
            Utilitario.Util_Log.Escribir_Log("antes del if ", "ManejoDeDocumentos")
            If oUserObjectMD.GetByKey("SS_DOCLEGALES") Then ' PREGUNTO SI ES UN UDO, YA QUE ALGUNOS CLIENTES NO TIENEN REGISTRADO EL UDO
                GuardaLOG(Tipotabla, DocEntry, "'SS_DOCLEGALES' es un UDO: ", Transaccion, TipoLog)
                oGeneralService = sCmp.GetGeneralService("SS_DOCLEGALES")
                Utilitario.Util_Log.Escribir_Log("SS_DOCLEGALES oGeneralService", "ManejoDeDocumentos")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("Code", code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'SS_DOCLEGALES' por el Code: " + code.ToString(), Transaccion, TipoLog)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                Utilitario.Util_Log.Escribir_Log("oGeneralData error", "ManejoDeDocumentos")
                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oGeneralData.GetProperty("U_UltimaSec").ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oGeneralData.GetProperty("U_UltimaSec").ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception
                    GuardaLOG(Tipotabla, DocEntry, "Cath Ultima Secuencia: " + ex.Message.ToString(), Transaccion, TipoLog)
                End Try

                If ActualizaSecuenc Then
                    oGeneralData.SetProperty("U_UltimaSec", UltimaSecuencia)
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'SS_DOCLEGALES' en el campo 'U_UltimaSec' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    oGeneralService.Update(oGeneralData)
                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If


                Return True
            Else
                GuardaLOG(Tipotabla, DocEntry, "'SS_DOCLEGALES' es una TABLA DE USUARIO: ", Transaccion, TipoLog)
                oUserTable = fCompany.UserTables.Item("SS_DOCLEGALES")
                oUserTable.GetByKey(code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'SS_DOCLEGALES' por el Code: " + code.ToString(), Transaccion, TipoLog)

                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oUserTable.UserFields.Fields.Item("U_UltimaSec").Value.ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception

                End Try

                If ActualizaSecuenc Then
                    oUserTable.UserFields.Fields.Item("U_UltimaSec").Value = UltimaSecuencia
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'UltimaSec' en el campo 'U_UltimaSec' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    'oUserTable.Update()
                    RetVal = oUserTable.Update()
                    If RetVal <> 0 Then
                        'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        fCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        GuardaLOG(Tipotabla, DocEntry, "ERROR en 'UltimaSec' al actualizar el campo 'U_UltimaSec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                    End If


                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If

                Return True
            End If


        Catch ex As Exception
            fPappl.SetStatusBarMessage(NombreAddon + "Error al actualizar la secuencia", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia" + ex.Message.ToString(), Transaccion, TipoLog)
            ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
            Return False
        End Try

    End Function

    Public Function ActualizaSecuencia_LiquidacionDeCompra(ByVal code As String, ByVal UltimaSecuencia As Integer, ByVal DocEntry As String, ByVal Tipotabla As String, ByVal Transaccion As String, ByVal TipoLog As String) As Boolean
        Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia_LiquidacionDeCompra (antes del try)", "ManejoDeDocumentos")
        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            Dim ActualizaSecuenc As Boolean = True
         
            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
            Dim oUserTable As SAPbobsCOM.UserTable = Nothing
            GC.Collect()
            oUserObjectMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = fCompany.GetCompanyService
            Utilitario.Util_Log.Escribir_Log("antes del if ", "ManejoDeDocumentos")
          
            oUserTable = fCompany.UserTables.Item("GS_LIQUI")
            oUserTable.GetByKey(code)
            GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'GS_LIQUI' por el Code: " + code.ToString(), Transaccion, TipoLog)

            Try
                GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oUserTable.UserFields.Fields.Item("U_Sec").Value.ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                If Integer.Parse(oUserTable.UserFields.Fields.Item("U_Sec").Value.ToString()) > UltimaSecuencia Then
                    ActualizaSecuenc = False
                End If

            Catch ex As Exception

            End Try

            If ActualizaSecuenc Then
                oUserTable.UserFields.Fields.Item("U_Sec").Value = UltimaSecuencia
                GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'GS_LIQUI' en el campo 'U_Sec' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                'oUserTable.Update()
                RetVal = oUserTable.Update()
                If RetVal <> 0 Then
                    'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    fCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_LIQUI' al actualizar el campo 'U_Sec' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                End If


            Else
                GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
            End If

            Return True


        Catch ex As Exception
            fPappl.SetStatusBarMessage(NombreAddon + "Error al actualizar la secuencia de Liquidacion de Compra", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
            ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
            Return False
        End Try

    End Function

    Public Function ActualizaDatosAutorizacion_TablaTM(ByVal code As String, ByVal UltimaSecuencia As Integer, ByVal DocEntry As String, ByVal Tipotabla As String, ByVal Transaccion As String, ByVal TipoLog As String) As Boolean

    End Function

    Public Function ActualizaSecuencia_ONE(ByVal Code As String, ByVal UltimaSecuencia As Integer, ByVal DocEntry As String, ByVal Tipotabla As String, ByVal Transaccion As String, ByVal TipoLog As String) As Boolean
        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            Dim ActualizaSecuenc As Boolean = True
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
            Dim oUserTable As SAPbobsCOM.UserTable = Nothing
            GC.Collect()
            oUserObjectMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = fCompany.GetCompanyService

            If oUserObjectMD.GetByKey("SERIES") Then ' PREGUNTO SI ES UN UDO, YA QUE ALGUNOS CLIENTES NO TIENEN REGISTRADO EL UDO
                GuardaLOG(Tipotabla, DocEntry, "'SERIES' es un UDO: ", Transaccion, TipoLog)
                oGeneralService = sCmp.GetGeneralService("SERIES")

                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("Code", Code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'SERIES' por el Code: " + Code.ToString(), Transaccion, TipoLog)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oGeneralData.GetProperty("U_ULT_SECUEN").ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oGeneralData.GetProperty("U_ULT_SECUEN").ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception
                End Try

                If ActualizaSecuenc Then
                    oGeneralData.SetProperty("U_ULT_SECUEN", UltimaSecuencia)
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'SERIES' en el campo 'U_ULT_SECUEN' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    oGeneralService.Update(oGeneralData)
                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If


                Return True
            Else
                GuardaLOG(Tipotabla, DocEntry, "'SERIES' es una TABLA DE USUARIO: ", Transaccion, TipoLog)
                oUserTable = fCompany.UserTables.Item("SERIES")
                oUserTable.GetByKey(code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'SERIES' por el Code: " + Code.ToString(), Transaccion, TipoLog)

                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception
                End Try

                If ActualizaSecuenc Then
                    oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value = UltimaSecuencia
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'SERIES' en el campo 'U_ULT_SECUEN' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    'oUserTable.Update()
                    RetVal = oUserTable.Update()
                    If RetVal <> 0 Then
                        'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        fCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        GuardaLOG(Tipotabla, DocEntry, "ERROR en 'SERIES' al actualizar el campo 'U_ULT_SECUEN' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                    End If


                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If

                Return True
            End If


        Catch ex As Exception
            fPappl.SetStatusBarMessage(NombreAddon + "Error al actualizar la secuencia", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia" + ex.Message.ToString(), Transaccion, TipoLog)
            ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
            Return False
        End Try

    End Function

    Public Function ActualizaSecuencia_SYPSOFT(ByVal Code As String, ByVal UltimaSecuencia As Integer, ByVal DocEntry As String, ByVal Tipotabla As String, ByVal Transaccion As String, ByVal TipoLog As String) As Boolean
        Try
            Dim RetVal As Long
            Dim ErrCode As Long
            Dim ErrMsg As String

            Dim ActualizaSecuenc As Boolean = True
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
            Dim oUserTable As SAPbobsCOM.UserTable = Nothing
            GC.Collect()
            oUserObjectMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            Dim sCmp As SAPbobsCOM.CompanyService
            sCmp = fCompany.GetCompanyService

            If oUserObjectMD.GetByKey("GS_SERIESE") Then ' PREGUNTO SI ES UN UDO, YA QUE ALGUNOS CLIENTES NO TIENEN REGISTRADO EL UDO
                GuardaLOG(Tipotabla, DocEntry, "'GS_SERIESE' es un UDO: ", Transaccion, TipoLog)
                oGeneralService = sCmp.GetGeneralService("GS_SERIESE")

                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("Code", Code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'GS_SERIESE' por el Code: " + Code.ToString(), Transaccion, TipoLog)

                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oGeneralData.GetProperty("U_ULT_SECUEN").ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oGeneralData.GetProperty("U_ULT_SECUEN").ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception
                End Try

                If ActualizaSecuenc Then
                    oGeneralData.SetProperty("U_ULT_SECUEN", UltimaSecuencia)
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'GS_SERIESE' en el campo 'U_ULT_SECUEN' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    oGeneralService.Update(oGeneralData)
                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If


                Return True
            Else
                GuardaLOG(Tipotabla, DocEntry, "'GS_SERIESE' es una TABLA DE USUARIO: ", Transaccion, TipoLog)
                oUserTable = fCompany.UserTables.Item("GS_SERIESE")
                oUserTable.GetByKey(Code)
                GuardaLOG(Tipotabla, DocEntry, "Obteniendo Registro a actualizar en 'GS_SERIESE' por el Code: " + Code.ToString(), Transaccion, TipoLog)

                Try
                    GuardaLOG(Tipotabla, DocEntry, "Validando SI 'Ultima Secuencia' Actual:" + oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString() + " MAYOR QUE 'Ultima Secuencia' Nueva: " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    If Integer.Parse(oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value.ToString()) > UltimaSecuencia Then
                        ActualizaSecuenc = False
                    End If

                Catch ex As Exception
                End Try

                If ActualizaSecuenc Then
                    oUserTable.UserFields.Fields.Item("U_ULT_SECUEN").Value = UltimaSecuencia
                    GuardaLOG(Tipotabla, DocEntry, "Actualizando en 'GS_SERIESE' en el campo 'U_ULT_SECUEN' a : " + UltimaSecuencia.ToString(), Transaccion, TipoLog)
                    'oUserTable.Update()
                    RetVal = oUserTable.Update()
                    If RetVal <> 0 Then
                        'rsboApp.theAppl.MessageBox(rsboApp.diCompany.GetLastErrorDescription())
#Disable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        fCompany.GetLastError(ErrCode, ErrMsg)
#Enable Warning BC42030 ' La variable 'ErrMsg' se ha pasado como referencia antes de haberle asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                        GuardaLOG(Tipotabla, DocEntry, "ERROR en 'GS_SERIESE' al actualizar el campo 'U_ULT_SECUEN' : " + ErrCode.ToString() + " - " + ErrMsg.ToString(), Transaccion, TipoLog)
                    End If


                Else
                    GuardaLOG(Tipotabla, DocEntry, "No se actualiza la secuencia, debido a que la nueva secuencia es MAYOR a la actual!!", Transaccion, TipoLog)
                End If

                Return True
            End If


        Catch ex As Exception
            fPappl.SetStatusBarMessage(NombreAddon + "Error al actualizar la secuencia", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia" + ex.Message.ToString(), Transaccion, TipoLog)
            ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
            Return False
        End Try

    End Function


    Public Sub GuardaLOG(ByVal TipoTabla As String, ByVal DocEntry_Clave As String, ByVal DescripcionLOG As String, ByVal Transaccion As String, ByVal TipoLog As String)
        If Functions.VariablesGlobales._vgGuardarLog = "Y" Then
            Utilitario.Util_Log.Escribir_Log("Guardando LOG a BASE..Transaccion: " + Transaccion + ", Descripcion: " + DescripcionLOG, "FuncionesAddonn")
            If Functions.VariablesGlobales._FacturaGuiaRemision = "SI" Then
                TipoTabla = "FCE"
            End If
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

            Dim ObjType As String = TipoTabla_TO_ObjType(TipoTabla)
            Dim DocSubType As String = TipoTabla_TO_DocSubType(TipoTabla)
            ''Utilitario.Util_Log.Escribir_Log("ObjType : " + ObjType + ", DocSubType: " + DocSubType, "FuncionesAddonn")

            Try
                Dim query As String
                Dim CodeExist As String = "0"
                If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    query = "Select ""DocEntry"" From """ & fCompany.CompanyDB & """.""@GS_LOG"" Where ""U_Clave"" = '" + DocEntry_Clave + "' AND ""U_ObjType"" = '" + ObjType + "'"
                Else
                    query = "Select DocEntry From [@GS_LOG] Where U_Clave = '" + DocEntry_Clave + "' AND U_ObjType = '" + ObjType + "'"
                End If
                CodeExist = getRSvalue(query, "DocEntry")

                ''Utilitario.Util_Log.Escribir_Log("Query : " + query, "FuncionesAddonn")
                ''Utilitario.Util_Log.Escribir_Log("CodeExist : " + CodeExist, "FuncionesAddonn")

                'mRst = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If CodeExist = "0" Then ' SI NO EXISTE EL UDO PARA EL DOCENTRY LO CREO CASO CONTRARIO LO ACTUALIZO
                    'conta = ConsultaID_LOG()
                    '.Util_Log.Escribir_Log("EXIST", "FuncionesAddonn")
                    oCompanyService = fCompany.GetCompanyService
                    Utilitario.Util_Log.Escribir_Log("GetCompanyService OK", "FuncionesAddonn")

                    Try
                        'oGeneralService = oCompanyService.GetGeneralService("SS_LOG")
                        oGeneralService = oCompanyService.GetGeneralService("SS_LOG")
                        'Utilitario.Util_Log.Escribir_Log("SS_LOG OK", "FuncionesAddonn")
                    Catch ex As Exception
                        Utilitario.Util_Log.Escribir_Log("SS_LOG CATH", "FuncionesAddonn")
                    End Try

#Disable Warning BC42104 ' La variable 'oGeneralService' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
#Enable Warning BC42104 ' La variable 'oGeneralService' se usa antes de que se le haya asignado un valor. Podría darse una excepción de referencia NULL en tiempo de ejecución.
                    'Utilitario.Util_Log.Escribir_Log("GetDataInterface OK", "FuncionesAddonn")

                    'oGeneralData.SetProperty("Code", conta)
                    oGeneralData.SetProperty("U_Clave", DocEntry_Clave)
                    'Utilitario.Util_Log.Escribir_Log("U_Clave OK " + DocEntry_Clave, "FuncionesAddonn")

                    oGeneralData.SetProperty("U_ObjType", ObjType)
                    'Utilitario.Util_Log.Escribir_Log("ObjType OK " + ObjType, "FuncionesAddonn")

                    oGeneralData.SetProperty("U_SubType", DocSubType)
                    'Utilitario.Util_Log.Escribir_Log("DocSubType OK " + DocSubType, "FuncionesAddonn")

                    oGeneralData.SetProperty("U_Tipo", TipoLog)
                    'Utilitario.Util_Log.Escribir_Log("TipoLog OK " + TipoLog, "FuncionesAddonn")

                    oChildren = oGeneralData.Child("GS_LOGD")
                    'Utilitario.Util_Log.Escribir_Log("oChildren OK ", "FuncionesAddonn")

                    oChild = oChildren.Add

                    oChild.SetProperty("U_Transacc", Transaccion)
                    'Utilitario.Util_Log.Escribir_Log("Transaccion OK " + Transaccion, "FuncionesAddonn")

                    oChild.SetProperty("U_Detalle", DescripcionLOG)
                    'Utilitario.Util_Log.Escribir_Log("U_Detalle OK " + DescripcionLOG, "FuncionesAddonn")

                    oChild.SetProperty("U_Fecha", Date.Now.ToString())
                    'Utilitario.Util_Log.Escribir_Log("U_Fecha OK " + Date.Now.ToString(), "FuncionesAddonn")

                    oGeneralParams = oGeneralService.Add(oGeneralData)
                    'Utilitario.Util_Log.Escribir_Log("Add OK " + oGeneralParams.ToString(), "FuncionesAddonn")

                Else
                    'Utilitario.Util_Log.Escribir_Log("NO EXIST", "FuncionesAddonn")
                    oCompanyService = fCompany.GetCompanyService
                    oGeneralService = oCompanyService.GetGeneralService("SS_LOG")
                    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    oGeneralParams.SetProperty("DocEntry", CodeExist)
                    '  Utilitario.Util_Log.LogEmisión(DirDelLog, "Obteniendo Registro a actualizar en 'EXX_DOCUM_LEG_INTER' por el Code: " + code.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                    oChildren = oGeneralData.Child("GS_LOGD")
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Transacc", Transaccion)
                    oChild.SetProperty("U_Detalle", DescripcionLOG)
                    oChild.SetProperty("U_Fecha", Date.Now.ToString())
                    'oGeneralParams = oGeneralService.Add(oGeneralData)
                    oGeneralService.Update(oGeneralData)

                End If
                'Utilitario.Util_Log.Escribir_Log("FIN", "FuncionesAddonn")

                '' Referencia1 = dpsParamAddCash.DepositNumber.ToString()
                'sDocEntry = oGeneralParams.GetProperty("Code")
            Catch ex As Exception
                Utilitario.Util_Log.Escribir_Log("GuardaLOG Catch: " + ex.Message.ToString, "FuncionesAddon")
            End Try
        End If
        
    End Sub

    Public Function ConsultaID_LOG() As String
        Dim ds As New DataSet
        Dim ID As String = ""
        Dim query As String

        If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            query = "Select ifnull(MAX(""DocEntry""),0)+1 AS ""MAX"" From """ & fCompany.CompanyDB & """.""@GS_LOG"""
        Else
            query = "Select isnull(MAX(DocEntry),0)+1 AS MAX From [@GS_LOG]"
        End If

        ID = getRSvalue(query, "MAX")

        Return ID
    End Function
    Structure Transacciones
        Const Creacion = "Creacion"
        Const Reenvío = "Reenvío"
        Const Consulta_PDF = "Consulta_PDF"
        Const Envío_Mail = "Envío_Mail"

        Const CreacionPreliminar = "CreacionPreliminar"
        Const CreacionFinal = "CreacionFinal"
        Const Cancelacion = "Cancelacion"

        Const EliminarDocMarcado = "EliminarRegistroDocMarcado"

    End Structure

    Structure TipoLog
        Const Emision = "Emision"
        Const Recepcion = "Recepcion"
    End Structure

    Structure PROVEEDOR_DE_SAPBO
        Const EXXIS = "EXXIS"
        Const ONESOLUTIONS = "ONESOLUTIONS"
        Const SYPSOFT = "SYPSOFT"
        Const HEINSOHN = "HEINSOHN"
        Const TOPMANAGE = "TOPMANAGE"
        Const SOLSAP = "SOLSAP"
    End Structure


    Public Function TipoTabla_TO_ObjType(ByVal TipoTabla As String) As String

        If TipoTabla = "FCE" Then
            Return "13"
        ElseIf TipoTabla = "FRE" Then
            Return "13"
        ElseIf TipoTabla = "NDE" Then
            Return "13"
        ElseIf TipoTabla = "FAE" Then
            Return "203"
        ElseIf TipoTabla = "NCE" Then
            Return "14"
        ElseIf TipoTabla = "GRE" Then
            Return "15"
        ElseIf TipoTabla = "TRE" Then
            Return "67"
        ElseIf TipoTabla = "REE" Then
            Return "18"
        ElseIf TipoTabla = "REA" Then
            Return "18"
        ElseIf TipoTabla = "RER" Then
            Return "18"
        ElseIf TipoTabla = "PRR" Then ' PAGO RECIBIDO RETENCION - RECEPCION
            Return "24"
        ElseIf TipoTabla = "NCR" Then ' NOTA DE CREDITO - RECEPCION
            Return "19"
        ElseIf TipoTabla = "TLE" Then ' SOLICITUD TRASLADO
            Return "1250000001"
        Else
            Return TipoTabla
        End If
    End Function
    Public Function TipoTabla_TO_DocSubType(ByVal TipoTabla As String) As String

        If TipoTabla = "FCE" Then
            Return "--"
        ElseIf TipoTabla = "FRE" Then
            Return "--"
        ElseIf TipoTabla = "NDE" Then
            Return "DN"
        ElseIf TipoTabla = "FAE" Then
            Return "--"
        ElseIf TipoTabla = "NCE" Then
            Return "--"
        ElseIf TipoTabla = "GRE" Then
            Return "--"
        ElseIf TipoTabla = "TRE" Then
            Return "--"
        ElseIf TipoTabla = "REE" Then
            Return "--"
        ElseIf TipoTabla = "REA" Then
            Return "--"
        ElseIf TipoTabla = "RER" Then
            Return "--"
        ElseIf Functions.VariablesGlobales._SS_FacturaExportacion = "SI" Then
            Return "IX"
        Else
            Return "--"
        End If

    End Function

    Public Shared Sub CreaComboEnFormulario(ByVal oFrm As SAPbouiCOM.Form)

        Try
            oFrm.Freeze(True)

            Dim itmCombo As SAPbouiCOM.Item

            Try
                itmCombo = oFrm.Items.Add("cmbElectronico", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)

                Dim iReferencia As SAPbouiCOM.Item = oFrm.Items.Item("84")

                itmCombo.AffectsFormMode = False
                itmCombo.Top = iReferencia.Top - 20
                itmCombo.Left = iReferencia.Left
                itmCombo.Width = 100
                itmCombo.Height = iReferencia.Height
                itmCombo.DisplayDesc = True
                itmCombo.LinkTo = 2

            Catch ex As Exception
                itmCombo = oFrm.Items.Item("cmbElectronico")
            End Try

            Dim cmbFeOpc As SAPbouiCOM.ButtonCombo = itmCombo.Specific

            If cmbFeOpc.ValidValues.Count > 0 Then
                For i As Integer = cmbFeOpc.ValidValues.Count - 1 To 0 Step -1
                    cmbFeOpc.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            cmbFeOpc.ValidValues.Add("0", "Opciones FEX")
            cmbFeOpc.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oFrm.Freeze(False)

        Catch ex As Exception
            oFrm.Freeze(False)
            Throw ex
        End Try

    End Sub

    Public Shared Sub CreaStaticTextEnFormularioSap(ByVal oFrm As SAPbouiCOM.Form, ByVal NombreNewItem As String, ByVal ItemsEnlace As String, ByVal ItemTop As Integer, ByVal ItemLeft As Integer, ByVal ItemWidth As Integer, ByVal blnvisible As Boolean, ByVal Caption As String)

        Try
            Dim itm As SAPbouiCOM.Item
            Dim itmC, itmText As SAPbouiCOM.Item
            Dim sta As SAPbouiCOM.StaticText

            itm = oFrm.Items.Add(NombreNewItem, SAPbouiCOM.BoFormItemTypes.it_STATIC)
            itmC = oFrm.Items.Item(ItemsEnlace)
            itmText = oFrm.Items.Item("5")

            itm.LinkTo = ItemsEnlace
            itm.Left = itmC.Left + ItemLeft
            itm.Top = itmC.Top + ItemTop
            itm.Width = ItemWidth
            itm.FontSize = itmText.FontSize

            'If CompanyVersion = eVersion.Sap_9 Then
            '    itm.TextStyle = itmText.TextStyle
            'End If

            itm.Visible = blnvisible
            sta = itm.Specific
            sta.Caption = Caption

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' CREA ITEM COMBO EN UN FORMULARIO SAP CON ENLACE A UN ITEM X
    ''' </summary>
    Public Shared Sub CreaComboBoxEnFormularioSap(ByVal oFrm As SAPbouiCOM.Form, ByVal NombreNewItem As String, ByVal ItemsEnlace As String, ByVal ItemTop As Integer, ByVal ItemLeft As Integer, ByVal ItemWidth As Integer, ByVal blnvisible As Boolean)

        Try
            Dim itm As SAPbouiCOM.Item
            Dim itmC As SAPbouiCOM.Item
            Dim cmb As SAPbouiCOM.ComboBox

            itm = oFrm.Items.Add(NombreNewItem, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            itmC = oFrm.Items.Item(ItemsEnlace)

            itm.LinkTo = ItemsEnlace
            itm.Left = itmC.Left + ItemLeft
            itm.Top = itmC.Top + ItemTop
            itm.Width = ItemWidth

            itm.Visible = blnvisible
            cmb = itm.Specific

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' Devuelve el valor de un campo de una consulta en formato String
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <param name="columnaRet">Columna de la consulta a retornar</param>
    ''' <param name="valorNulo">Valor a retornar en caso de error/nulo</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRSvalue(ByVal query As String, ByVal columnaRet As String, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRSvalue Catch: " + ex.Message.ToString() + " - Query: " + query, "FuncionesAddon")
        End Try
        Return ret
    End Function
    ''' <summary>
    ''' Libera un objeto de la memoria. Se recomienda usar con objetos de meta-datos.
    ''' </summary>
    ''' <param name="myObject">Objeto a liberar</param>
    ''' <remarks></remarks>
    Public Sub Release(ByVal myObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myObject)
            myObject = Nothing
            GC.Collect()
        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Devuelve un string. Si el parámetro es nulo, devuelve una cadena vacía.
    ''' </summary>
    ''' <param name="unString">Valor a convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzString(ByVal unString As String, Optional ByVal formatoSQL As Boolean = False, Optional ByVal valorSiNulo As String = "") As String
        Try
            If Not IsDBNull(unString) Then
                If formatoSQL Then
                    unString = unString.Replace("'", "' + CHAR(39) + '")
                End If
                valorSiNulo = unString
            End If
        Catch ex As Exception
        End Try
        Return valorSiNulo
    End Function
    ''' <summary>
    ''' Devuelve un Recordset a partir de un query
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRecordSet(ByVal query As String) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
        End Try
        Return fRS
    End Function

    'Public Function ValidarSerieElectronica(ByVal PartnerSap As String, ByVal TipoDoc As String, ByVal IdSerie As String) As Boolean
    '    'Utilitario.Util_Log.Escribir_Log("Ingresando a la funcion ActualizaSecuencia_LiquidacionDeCompra (antes del try)", "ManejoDeDocumentos")
    '    Try
    '        'Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS
    '        Dim serie As String = ""
    '        If PartnerSap = "EXXIS" Then

    '            If TipoDoc = "REE" Or TipoDoc = "RER" Or TipoDoc = "RDM" Then

    '                If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '                    serie = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE B.""U_TIPO_DOC""='RT' and A.""Series"" = " + IdSerie
    '                Else
    '                    serie = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE B.""U_TIPO_DOC""='RT' and Series = " + IdSerie
    '                End If
    '            Else
    '                If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '                    serie = "SELECT IFNULL(""U_FE_TipoEmision"",'NA') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@EXX_DOCUM_LEG_INTER"" B ON A.""SeriesName"" = B.""U_NOMBRE"" WHERE A.""Series"" = " + IdSerie
    '                Else
    '                    serie = "SELECT ISNULL(U_FE_TipoEmision,'NA') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@EXX_DOCUM_LEG_INTER] B WITH(NOLOCK) ON A.SeriesName = B.U_NOMBRE WHERE Series = " + IdSerie
    '                End If
    '            End If



    '        ElseIf PartnerSap = "ONESOLUTIONS" Then
    '            If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '                serie = "SELECT ""U_DIGITAL"" FROM ""@SERIES"" WHERE ""U_SERIE"" = " + IdSerie
    '            Else
    '                serie = "SELECT U_DIGITAL FROM ""@SERIES"" WHERE U_SERIE = " + IdSerie
    '            End If
    '        ElseIf PartnerSap = "HEINSOHN" Or PartnerSap = "SYPSOFT" Or PartnerSap = "TOPMANAGE" Then

    '            If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '                serie = "SELECT ""U_ELECTRONICA"" FROM ""@GS_SERIESE"" WHERE ""Code"" = " + IdSerie
    '            Else
    '                serie = "SELECT U_ELECTRONICA FROM ""@GS_SERIESE"" WHERE Code = " + IdSerie
    '            End If

    '        ElseIf PartnerSap = "SOLSAP" Then

    '            If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
    '                serie = "SELECT IFNULL(""U_TipoD"",'NAN') AS U_FE_TipoEmision FROM ""NNM1"" A INNER JOIN ""@SS_SERD"" B ON A.""SeriesName"" = B.""U_SerN"" WHERE A.""Series"" = " + IdSerie
    '            Else
    '                serie = "SELECT ISNULL(U_TipoD,'NAN') AS U_FE_TipoEmision FROM NNM1 A WITH(NOLOCK) INNER JOIN [@SS_SERD] B WITH(NOLOCK) ON A.SeriesName = B.U_SerN WHERE Series = " + IdSerie
    '            End If
    '        End If

    '        Utilitario.Util_Log.Escribir_Log("QUERY A EJECUTAR PARA VALIDAR SI ES ELECTRÓNICO: " + serie, "FuncionesAddon")
    '        If Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.EXXIS Then
    '            EsElectronico = getRSvalue(Sql, "U_FE_TipoEmision", "")

    '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.ONESOLUTIONS Then
    '            EsElectronico = oFuncionesB1.getRSvalue(Sql, "U_DIGITAL", "")

    '            If EsElectronico = "Y" Then
    '                EsElectronico = "FE"
    '            Else
    '                EsElectronico = "NA"
    '            End If

    '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.HEINSOHN Then
    '            EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
    '            If EsElectronico = "SI" Then
    '                EsElectronico = "FE"
    '            Else
    '                EsElectronico = "NA"
    '            End If

    '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SYPSOFT Or
    '                Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.TOPMANAGE Then
    '            EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_ELECTRONICA", "")
    '            If EsElectronico = "SI" Then
    '                EsElectronico = "FE"
    '            Else
    '                EsElectronico = "NA"
    '            End If
    '        ElseIf Nombre_Proveedor_SAP_BO = Functions.FuncionesAddon.PROVEEDOR_DE_SAPBO.SOLSAP Then
    '            EsElectronico = oFuncionesB1.getRSvalue(SQL, "U_FE_TipoEmision", "")
    '            If oTipoTabla = "REE" Or oTipoTabla = "RER" Or oTipoTabla = "RDM" Then
    '                If EsElectronico = "RT" Or EsElectronico = "LQRT" Then
    '                    EsElectronico = "FE"
    '                ElseIf EsElectronico = "NAN" Or EsElectronico = "" Then
    '                    EsElectronico = "NA"
    '                Else
    '                    EsElectronico = "NA"
    '                End If
    '            Else
    '                If EsElectronico = "NAN" Or EsElectronico = "" Then
    '                    EsElectronico = "NA"
    '                Else
    '                    EsElectronico = "FE"
    '                End If
    '            End If
    '        End If
    '        Utilitario.Util_Log.Escribir_Log("RESPUESTA DEL QUERY: " + EsElectronico, "EventosEmision")
    '        End If

    '    Catch ex As Exception
    '        fPappl.SetStatusBarMessage(NombreAddon + "Error al actualizar la secuencia de Liquidacion de Compra", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
    '        'GuardaLOG(Tipotabla, DocEntry, "Error al actualizar la secuencia de Liquidacion de Compra" + ex.Message.ToString(), Transaccion, TipoLog)
    '        ' Utilitario.Util_Log.LogEmisión(DirDelLog, "Error al actualizar la secuencia TRY CATCH: " + ex.Message.ToString(), Utilitario.Util_Log.Transacciones.Creacion, oTipoTabla, DocNum)
    '        Return False
    '    End Try

    'End Function

    'Public Shared Sub ActualizaButtonComboEnFormularioSap(ByVal valCombo As String, ByVal oFrm As SAPbouiCOM.Form)

    '    Try
    '        Dim itmCombo As SAPbouiCOM.Item = oFrm.Items.Item("cmbFeOpc")
    '        Dim cmbFeOpc As SAPbouiCOM.ButtonCombo
    '        cmbFeOpc = itmCombo.Specific()

    '        If Not valCombo.Equals(String.Empty) Then

    '            Dim valores As String()

    '            valores = {eEstadoDocumento.Autorizado
    '                      }

    '            If valores.Contains(valCombo) Then
    '                cmbFeOpc.ValidValues.Add("1", "Ver XML")
    '                cmbFeOpc.ValidValues.Add("2", "Ver PDF")
    '            End If

    '            valores = {eEstadoDocumento.Contingencia
    '                      }

    '            If valores.Contains(valCombo) Then
    '                cmbFeOpc.ValidValues.Add("2", "Ver PDF")
    '            End If

    '            valores = {eEstadoDocumento.Autorizado,
    '                       eEstadoDocumento.Contingencia
    '                      }

    '            If valores.Contains(valCombo) Then
    '                cmbFeOpc.ValidValues.Add("3", "Enviar Mail")
    '            End If

    '            valores = {eEstadoDocumento.Devuelto,
    '                       eEstadoDocumento.No_Autorizado,
    '                       eEstadoDocumento.Proceso_SRI
    '                      }

    '            If valores.Contains(valCombo) Then
    '                cmbFeOpc.ValidValues.Add("4", "Reenviar SRI")
    '            End If

    '            valores = {eEstadoDocumento.Autorizado,
    '                       eEstadoDocumento.Contingencia
    '                      }

    '            If valores.Contains(valCombo) Then
    '                cmbFeOpc.ValidValues.Add("5", "Regenerar PDF")
    '            End If

    '        End If

    '        itmCombo.Visible = False

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Sub

    'Public Sub ExportGridToExcel(ByVal grid As SAPbouiCOM.Grid)

    '    Try

    '        fPappl.SetStatusBarMessage("Generado visualizacion en Excel, por favor espere un momento..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

    '        ' Crear una instancia de Excel
    '        Dim excelApp As New Microsoft.Office.Interop.Excel.Application
    '        Dim workbook As Workbook = excelApp.Workbooks.Add()
    '        Dim worksheet As Worksheet = workbook.Sheets(1)

    '        ' Encabezados de las columnas del Grid
    '        For col As Integer = 0 To grid.Columns.Count - 1
    '            Dim cellValue As String = grid.Columns.Item(col).TitleObject.Caption
    '            worksheet.Cells(1, col + 1) = grid.Columns.Item(col).TitleObject.Caption
    '        Next
    '        'Dim oGridDet As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

    '        ' Datos del Grid
    '        'For row As Integer = 0 To grid.Rows.Count
    '        '    For col As Integer = 0 To grid.Columns.Count
    '        '        worksheet.Cells(row + 2, col + 1) = grid.DataTable.GetValue(col, row).ToString()
    '        '    Next
    '        'Next
    '        Dim k As Integer = 0
    '        For i As Integer = 0 To grid.Rows.Count - 1
    '            ' Verificar si la fila es una fila de agrupación
    '            If grid.Rows.IsLeaf(i) Then
    '                ' Saltar la fila de agrupación
    '                'Continue For
    '                For j As Integer = 0 To grid.Columns.Count - 1
    '                    'obtener el nombre de la colunma para setearlo como formato texto
    '                    Dim NombreColumna As String = grid.Columns.Item(j).TitleObject.Caption
    '                    If NombreColumna.ToLower = "autorizacion" Or NombreColumna.ToLower = "Clave acceso" Or NombreColumna.ToLower = "ruc" Or NombreColumna.ToLower = "identificacion" Or NombreColumna.ToLower = "tipo comprobante" _
    '                        Or NombreColumna.ToLower = "tipo identificacion" Or NombreColumna.ToLower = "sustento tributario" Or NombreColumna.ToLower = "tipo sujeto retenido" Then
    '                        worksheet.Cells(k + 2, j + 1).NumberFormat = "@"
    '                    End If
    '                    ' Dim cellValue As String = grid.DataTable.GetValue(j, i).ToString()
    '                    'rsboApp.MessageBox("Columna :" + j.ToString + " fila: " + i.ToString + "Dato: " + grid.DataTable.GetValue(j, i - 1).ToString())
    '                    worksheet.Cells(k + 2, j + 1) = grid.DataTable.GetValue(j, k).ToString()
    '                    ' Realizar tus operaciones con cellValue
    '                Next
    '                k += 1
    '            End If

    '            ' Procesar la fila de datos (no agrupada)
    '            'For j As Integer = 0 To grid.Columns.Count - 1
    '            '    ' Acceder a las celdas del Grid que contienen los datos

    '            '    ' Dim cellValue As String = grid.DataTable.GetValue(j, i).ToString()
    '            '    worksheet.Cells(i + 2, j + 1) = grid.DataTable.GetValue(j, i).ToString()
    '            '    ' Realizar tus operaciones con cellValue
    '            'Next
    '        Next
    '        ' Mostrar Excel al usuario
    '        fPappl.SetStatusBarMessage("Abriendo Excel...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    '        excelApp.Visible = True

    '        ' Liberar objetos COM
    '        Release(worksheet)
    '        Release(workbook)
    '        Release(excelApp)


    '    Catch ex As Exception

    '        fPappl.SetStatusBarMessage("Error al enviar informacion a excel " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)


    '    End Try

    'End Sub

    Public Sub ExportGridToExcel(ByVal grid As SAPbouiCOM.Grid)

        Try

            fPappl.SetStatusBarMessage("Generado visualizacion en Excel, por favor espere un momento..!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            ' Crear una instancia de Excel
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application
            Dim workbook As Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Worksheet = workbook.Sheets(1)

            ' Encabezados de las columnas del Grid
            For col As Integer = 0 To grid.Columns.Count - 1
                Dim cellValue As String = grid.Columns.Item(col).TitleObject.Caption
                worksheet.Cells(1, col + 1) = grid.Columns.Item(col).TitleObject.Caption
            Next
            Dim ListNomColumna As New List(Of String)

            For Each NomColumna As String In VariablesGlobales._NombreColumnbasAnexo.Split(New Char() {";"c})
                ListNomColumna.Add(NomColumna)
            Next

            Dim k As Integer = 0
            For i As Integer = 0 To grid.Rows.Count - 1
                ' Verificar si la fila es una fila de agrupación
                If grid.Rows.IsLeaf(i) Then
                    ' Saltar la fila de agrupación
                    'Continue For
                    For j As Integer = 0 To grid.Columns.Count - 1
                        'obtener el nombre de la colunma para setearlo como formato texto
                        Dim NombreColumna As String = grid.Columns.Item(j).TitleObject.Caption
                        'If NombreColumna.ToLower = "autorizacion" Or NombreColumna.ToLower = "Clave acceso" Or NombreColumna.ToLower = "ruc" Or NombreColumna.ToLower = "identificacion" Or NombreColumna.ToLower = "tipo comprobante" _
                        '    Or NombreColumna.ToLower = "tipo identificacion" Or NombreColumna.ToLower = "sustento tributario" Or NombreColumna.ToLower = "tipo sujeto retenido" Then
                        '    worksheet.Cells(k + 2, j + 1).NumberFormat = "@"
                        'End If
                        If ListNomColumna.Contains(NombreColumna) Then
                            worksheet.Cells(k + 2, j + 1).NumberFormat = "@"

                        End If
                        worksheet.Cells(k + 2, j + 1) = grid.DataTable.GetValue(j, k).ToString()
                    Next
                    k += 1
                End If

                ' Procesar la fila de datos (no agrupada)
                'For j As Integer = 0 To grid.Columns.Count - 1
                '    ' Acceder a las celdas del Grid que contienen los datos

                '    ' Dim cellValue As String = grid.DataTable.GetValue(j, i).ToString()
                '    worksheet.Cells(i + 2, j + 1) = grid.DataTable.GetValue(j, i).ToString()
                '    ' Realizar tus operaciones con cellValue
                'Next
            Next
            ' Mostrar Excel al usuario
            fPappl.SetStatusBarMessage("Abriendo Excel...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            excelApp.Visible = True

            ' Liberar objetos COM
            Release(worksheet)
            Release(workbook)
            Release(excelApp)
            ListNomColumna.Clear()
            ListNomColumna = Nothing

        Catch ex As Exception

            fPappl.SetStatusBarMessage("Error al enviar informacion a excel " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)


        End Try

    End Sub

End Class

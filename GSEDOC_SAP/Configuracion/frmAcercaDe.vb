Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions

Public Class frmAcercaDe
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
        If RecorreFormulario(rsboApp, "frmAcercaDe") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmAcercaDe.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)

            Catch exx As Exception
                rsboApp.Forms.Item("frmAcercaDe").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmAcercaDe")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            'ipLogo.Picture = Application.StartupPath & "\imagen.jpg"
            ipLogo.Picture = Application.StartupPath & "\LogoSS.png"

            Dim txtRes As SAPbouiCOM.EditText
            txtRes = oForm.Items.Item("txtRes").Specific
            txtRes.Value = "Addon que integra la solucón EDOC de manera nativa a SAP de tal forma que maneja la emision y recepción de los documentos electrónicos."

            'Dim sQueryVersion As String = ""
            'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '    sQueryVersion = "SELECT TOP 1  ""U_SS_VERS"" FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" + NombreAddon + "' ORDER BY 1 DESC"
            'Else
            '    sQueryVersion = "SELECT TOP 1  U_SS_VERS FROM ""@SS_SETUP"" WHERE U_SS_ADDN = '" + NombreAddon + "' ORDER BY 1 DESC"
            'End If
            'Dim Version As String = oFuncionesB1.getRSvalue(sQueryVersion, "U_SS_VERS", "")
            Dim lbVersion As SAPbouiCOM.StaticText
            lbVersion = oForm.Items.Item("lbVersion").Specific
            lbVersion.Caption = "Versión : " + Functions.VariablesGlobales._vgVersionAddOn

            Dim lbUrl As SAPbouiCOM.StaticText
            lbUrl = oForm.Items.Item("lbUrl").Specific
            lbUrl.Item.ForeColor = RGB(6, 69, 173)
            lbUrl.Item.TextStyle = 4

            Dim lbValido As SAPbouiCOM.StaticText
            lbValido = oForm.Items.Item("lbValido").Specific

            ' TIENE LICENCIA
            'Dim sQueryLicencia As String = ""
            'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '    sQueryLicencia = "SELECT TOP 1  ""U_SS_LICEN"" FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" + NombreAddon + "' ORDER BY ""Code"" DESC"
            'Else
            '    sQueryLicencia = "SELECT TOP 1  U_SS_LICEN FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" + NombreAddon + "' ORDER BY Code DESC"
            'End If
            'Dim Fecha As Integer = 0
            'Dim Result As String = oFuncionesB1.getRSvalue(sQueryLicencia, "U_SS_LICEN", "").ToString()
            'If Result = "" Then
            '    ' NO TIENE LICENCIA
            '    lbValido.Caption = "No tiene licencia asignada, contactese con un asesor de SOLSAP S.A."
            '    lbValido.Item.ForeColor = RGB(204, 0, 0)
            'Else
            '    Dim sLicenciaDesencriptada = Utilitario.Util_Encriptador.Desencriptar(Result, sKey) ' DESENCRIPTO
            '    Dim slicenciaXML As New XmlDocument()
            '    slicenciaXML.LoadXml(sLicenciaDesencriptada)
            '    Dim xmlnode As XmlNodeList
            '    xmlnode = slicenciaXML.GetElementsByTagName("Licencia")
            '    Dim x As Integer
            '    'Dim oLicencia As Licencia = Nothing
            '    If xmlnode.Count > 0 Then ' RECORRO XML CARGO CLASE LICENCIA
            '        oLicencia = New Licencia
            '        For x = 0 To xmlnode.Count - 1
            '            oLicencia.NombreBaseSAP = xmlnode(x).ChildNodes.Item(0).InnerText.Trim()
            '            oLicencia.Opcion = xmlnode(x).ChildNodes.Item(1).InnerText.Trim()
            '            oLicencia.validoHasta = xmlnode(x).ChildNodes.Item(2).InnerText.Trim()
            '        Next
            '    End If
            '    Dim FechaD As Date = DateSerial(oLicencia.validoHasta.ToString.Substring(0, 4), oLicencia.validoHasta.ToString.Substring(4, 2), oLicencia.validoHasta.ToString.Substring(6, 2))
            '    If oLicencia.validoHasta < Integer.Parse(Date.Now.ToString("yyyyMMdd")) Then
            '        lbValido.Caption = "Su licencia esta vencida! Contactese con un asesor de SOLSAP S.A."
            '        lbValido.Caption = "Valido Hasta : " + FechaD.ToString("MMMM dd, yyyy")
            '        lbValido.Item.ForeColor = RGB(204, 0, 0)
            '    Else
            '        lbValido.Caption = "Valido Hasta : " + FechaD.ToString("MMMM dd, yyyy")
            '        lbValido.Item.ForeColor = RGB(7, 118, 10)
            '    End If
            'End If
            If Functions.VariablesGlobales._vgTieneLicenciaActivaAddOn = False Then
                lbValido.Caption = "Su licencia esta vencida! Contactese con un asesor de SOLSAP S.A."
                lbValido.Item.ForeColor = RGB(204, 0, 0)

            Else
                'lbValido.Caption = "Su licencia esta Activa! "
                If Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "emision" Then
                    lbValido.Caption = "Su licencia Tipo Emisión esta Activa!"
                ElseIf Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "recepcion" Then
                    lbValido.Caption = "Su licencia Tipo Recepción esta Activa!"
                ElseIf Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToLower = "full" Then
                    lbValido.Caption = "Su licencia Tipo Full esta Activa!"
                End If
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

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
                   And pVal.FormTypeEx = "frmAcercaDe" Then
            If pVal.BeforeAction = False And pVal.ItemUID = "lbUrl" Then
                Try
                    oForm = rsboApp.Forms.Item("frmAcercaDe")
                    Dim lbUrl As SAPbouiCOM.StaticText
                    lbUrl = oForm.Items.Item("lbUrl").Specific
                    System.Diagnostics.Process.Start(lbUrl.Caption.ToString())
                Catch ex As Exception

                End Try
            ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnLi" Then
                'ProcesoBtnExaminarPart()
                rsboApp.StatusBar.SetText(NombreAddon + " - Creando UDO de Configuracíón, favor espere...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                rEstructura.CreacionDeCampos_Y_UDO_Configuracion()
                rsboApp.StatusBar.SetText(NombreAddon + " Felicitaciones - UDO de Configuracíón creado con exito!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            ElseIf pVal.BeforeAction = False And pVal.ItemUID = "lnkConf" Then
                'ofrmConfClave.CargaFormularioValidarClave()
                Utilitario.Util_Log.Escribir_Log("tipo licencia: " + Functions.VariablesGlobales._vgTipoLicenciaAddOn.ToString, "frmAcercaDe")
                Try
                    If IsNothing(Functions.VariablesGlobales._vgTipoLicenciaAddOn) Or Functions.VariablesGlobales._vgTipoLicenciaAddOn = "" Or Functions.VariablesGlobales._vgPruebaSinWSLIC = "Pruebas" Then
                        Try
                            'ofrmParametrosAddon.CargaFormularioParametrosADDON()
                            ofrmConfClave.CargaFormularioValidarClave()
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("cargar formulario parametros addon: " + ex.Message.ToString, "frmAcercaDe")
                        End Try

                    Else
                        Try
                            ofrmValidarUsuario.CargaFormularioValidarUsuario()
                        Catch ex As Exception
                            Utilitario.Util_Log.Escribir_Log("cargar formulario validar usuario: " + ex.Message.ToString, "frmAcercaDe")
                        End Try

                    End If
                Catch ex As Exception
                    Utilitario.Util_Log.Escribir_Log("validar que ventana mostrar: " + ex.Message.ToString, "frmAcercaDe")
                End Try




            ElseIf pVal.BeforeAction = False And pVal.ItemUID = "lnkEst" Then ' VALIDAR ESTRUCTURA BD

                Dim resp = rsboApp.MessageBox("Que Estructura Nesecita Crear?", 1, "FE", "FE + LOC", "GR-DES")

                If resp = 1 Then

                    rEstructura.CreacionDeEstructura()

                ElseIf resp = 2 Then

                    rEstructura.CreacionDeEstructura(True)
                ElseIf resp = 3 Then
                    rEstructura.CrearEstructuraGuiaDesatendida()
                End If


                '
            End If
        End If

    End Sub

    Private Sub ProcesoBtnExaminarPart()

        rsboApp.StatusBar.SetText(NombreAddon + " Revise las ventanas abiertas (ALT+TAB), puede que la ventana no quedo como principal.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        GetfileThread = New Threading.Thread(AddressOf GetNombreArchivoPart, 1)
        GetfileThread.SetApartmentState(Threading.ApartmentState.STA)
        GetfileThread.Start()

    End Sub

    Private Sub GetNombreArchivoPart()

        Try
            Dim busca As New OpenFileDialog
            Dim path As String

            busca.DefaultExt = ".txt"
            busca.Filter = "Texto|*.txt|Todos los tipos (*.*)|*.*"
            path = GetSetting("SBOCYP", "software", "FILEPATH", System.Environment.CurrentDirectory)
            busca.InitialDirectory = path
            busca.FilterIndex = 1

            If busca.ShowDialog = DialogResult.OK Then

                SaveSetting("SBOCYP", "software", "FILEPATH", IO.Path.GetDirectoryName(busca.FileName))

                'DateTime.Now.ToString("HH.mm.ss")

                Dim fi As New System.IO.FileInfo(busca.FileName)

                Dim sLic As New StringBuilder()
                Dim file As New System.IO.StreamReader(fi.FullName)
                sLic.Append(file.ReadLine())
                file.Close()

                Dim sLicenciaDesencriptada = Utilitario.Util_Encriptador.Desencriptar(sLic.ToString().Replace("{", "").Replace("}", "").ToString(), sKey) ' DESENCRIPTO
                Dim slicenciaXML As New XmlDocument()
                slicenciaXML.LoadXml(sLicenciaDesencriptada)
                Dim xmlnode As XmlNodeList
                xmlnode = slicenciaXML.GetElementsByTagName("Licencia")
                Dim x As Integer
                ' Dim oLicencia As Licencia = Nothing
                If xmlnode.Count > 0 Then ' RECORRO XML CARGO CLASE LICENCIA
                    For x = 0 To xmlnode.Count - 1
                        oLicencia.NombreBaseSAP = xmlnode(x).ChildNodes.Item(0).InnerText.Trim()
                        oLicencia.Opcion = xmlnode(x).ChildNodes.Item(1).InnerText.Trim()
                        oLicencia.validoHasta = xmlnode(x).ChildNodes.Item(2).InnerText.Trim()
                    Next
                End If

                If oLicencia.NombreBaseSAP <> rCompany.CompanyDB Then
                    rsboApp.StatusBar.SetText(NombreAddon + " La licencia no corresponde a la compañia que seleccionada! - " + oLicencia.NombreBaseSAP.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                If oLicencia.validoHasta < Integer.Parse(Date.Now.ToString("yyyyMMdd")) Then
                    rsboApp.StatusBar.SetText(NombreAddon + " La licencia seleccionada tiene fecha: " + oLicencia.validoHasta + " la cual se encuentra vencida!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
               
                ' DESPUES QUE SE VALIDA SE MANDA A GUARDAR LA LICENCIA
                Dim Result As Boolean = GuardarLicencia(sLic.ToString())
                'Dim Result As Boolean = GuardarLicencia(Convert.ToBase64String(ArchivoBytes))
                If Result Then
                    Dim lbValido As SAPbouiCOM.StaticText
                    lbValido = oForm.Items.Item("lbValido").Specific
                    Dim FechaD As Date = DateSerial(oLicencia.validoHasta.ToString.Substring(0, 4), oLicencia.validoHasta.ToString.Substring(4, 2), oLicencia.validoHasta.ToString.Substring(6, 2))
                    lbValido.Caption = "Valido Hasta :" + FechaD.ToString("MMMM dd, yyyy")
                    lbValido.Item.ForeColor = RGB(7, 118, 10)

                    'VALIDANDO UDO DE CONFIGURACION
                    rsboApp.StatusBar.SetText(NombreAddon + " - Validando UDO Configuracíón, favor espere...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    rEstructura.CreacionDeCampos_Y_UDO_Configuracion()

                    rsboApp.MessageBox(NombreAddon + " Licencia Actualizada Correctamente, por favor volver a ingresar a SAP para que se inicie el addon con su nueva licencia!")

                    rsboApp.StatusBar.SetText(NombreAddon + " Felicitaciones!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If


            End If

        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " Error al cargar Licencia, " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    'Private Function LeerLicencia(ByVal sInputFilename As String) As Licencia
    '    Try
    '        Dim oLicencia As Licencia
    '        Dim DES As New DESCryptoServiceProvider()
    '        'Se requiere una clave de 64 bits y IV para este proveedor.
    '        'Establecer la clave secreta para el algoritmo DES.
    '        DES.Key() = ASCIIEncoding.ASCII.GetBytes(sKey)
    '        'Establecer el vector de inicialización.
    '        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

    '        'crear la secuencia de archivos para volver a leer el archivo cifrado
    '        Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
    '        'crear descriptor DES a partir de nuestra instancia de DES
    '        Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
    '        'crear conjunto de secuencias de cifrado para leer y realizar 
    '        'una transformación de descifrado DES en los bytes entrantes
    '        Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
    '        'imprimir el contenido de archivo descifrado
    '        Dim XMLDesencriptado As String = Path.GetTempPath() + "\LicenciaDesencriptada.xml"
    '        Dim fsDecrypted As New StreamWriter(XMLDesencriptado)
    '        fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd)
    '        fsDecrypted.Flush()
    '        fsDecrypted.Close()

    '        Dim doc As New XmlDocument()
    '        Dim fs As FileStream
    '        Dim xmlnode As XmlNodeList
    '        fs = New FileStream(XMLDesencriptado, FileMode.Open, FileAccess.Read)
    '        doc.Load(fs)
    '        xmlnode = doc.GetElementsByTagName("Licencia")
    '        Dim x As Integer
    '        If xmlnode.Count > 0 Then
    '            oLicencia = New Licencia
    '            For x = 0 To xmlnode.Count - 1
    '                oLicencia.NombreBaseSAP = xmlnode(x).ChildNodes.Item(0).InnerText.Trim()
    '                oLicencia.Opcion = xmlnode(x).ChildNodes.Item(1).InnerText.Trim()
    '                oLicencia.validoHasta = xmlnode(x).ChildNodes.Item(2).InnerText.Trim()
    '            Next
    '        End If

    '        Return oLicencia

    '    Catch ex As Exception

    '        Return Nothing
    '    End Try


    'End Function

    Private Function GuardarLicencia(Contenido As String) As Boolean
        Try
            ' OBTENER ULTIMO CODE DE VERSION ADDON
            Dim sQueryCode As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryCode = "SELECT TOP 1  ""Code"" FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" + NombreAddon + "' ORDER BY 1 DESC"
            Else
                sQueryCode = "SELECT TOP 1  Code FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" + NombreAddon + "' ORDER BY 1 DESC"
            End If
            Dim Code As String = oFuncionesB1.getRSvalue(sQueryCode, "Code", "").ToString()

            Dim lErrCode As Integer = 0
            Dim sErrMsg As String = ""
            Dim oUserTable As SAPbobsCOM.UserTable
            Try
                '// set the object with the requested table
                oUserTable = rCompany.UserTables.Item("SS_SETUP")
                If oUserTable.GetByKey(Code) Then

                    oUserTable.UserFields.Fields.Item("U_SS_LICEN").Value = Contenido

                    oUserTable.Update()
                    '// Check for errors
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    If lErrCode <> 0 Then
                        rsboApp.StatusBar.SetText(NombreAddon + " Error al Guardar Licencia: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Return True
                Else
                    oUserTable.Code = oFuncionesB1.getCorrelativo("Code", """@SS_SETUP""", , 1)
                    oUserTable.Name = oFuncionesB1.getCorrelativo("Code", """@SS_SETUP""", , 1)
                    oUserTable.UserFields.Fields.Item("U_SS_ADDN").Value = NombreAddon.ToString()
                    oUserTable.UserFields.Fields.Item("U_SS_VERS").Value = "0"

                    oUserTable.UserFields.Fields.Item("U_SS_LICEN").Value = Contenido

                    oUserTable.Add()
                    '// Check for errors
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    If lErrCode <> 0 Then
                        rsboApp.StatusBar.SetText(NombreAddon + " Error al Guardar Licencia: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Return True
                End If
            Catch ex As Exception
                Return False
            Finally
                oUserTable = Nothing
                System.GC.Collect()
            End Try

        Catch ex As Exception
            Return False
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

End Class

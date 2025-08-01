Imports System.IO
Imports System.Text

Public Class frmSubirArchivo
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Dim odt As SAPbouiCOM.DataTable
    Dim ClaveAcceso As String = ""
    Dim _WS_RecepcionClave As String = ""
    Dim _WS_RecepcionArchivo As String = ""

    Dim Claves As List(Of String)

    Private GetfileThread As Threading.Thread

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioSubirArchivo()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmSubirArchivo") Then
            Exit Sub
        End If
     
        strPath = System.Windows.Forms.Application.StartupPath & "\frmSubirArchivo.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmSubirArchivo").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmSubirArchivo")
            oForm.Freeze(True)

            Try
                oForm.DataSources.DataTables.Add("dtDocs")
            Catch ex As Exception
            End Try

            oForm.DataSources.DataTables.Item("dtDocs").Clear()
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Count", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            oForm.DataSources.DataTables.Item("dtDocs").Columns.Add("Clave", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)

            oForm.Width = 394
            oForm.Height = 294

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                   And pVal.FormTypeEx = "frmSubirArchivo" Then
            If pVal.BeforeAction = False And pVal.ItemUID = "btnArchivo" Then
                ProcesoBtnExaminarPart()
            ElseIf pVal.BeforeAction = False And pVal.ItemUID = "obtnPro" Then
                Try

                    rsboApp.StatusBar.SetText(NombreAddon + " - Recuperando información de Web Services...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    _WS_RecepcionArchivo = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "WS_RecepcionConsultaArchivo")
                    If _WS_RecepcionArchivo = "" Then
                        rsboApp.SetStatusBarMessage(NombreAddon + " - No existe parametrización del Web Services de Recepcion, Revisar Parametrización", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    End If
                    _WS_RecepcionClave = ofrmParametrosAddon.ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "RecepcionClave")

                    Dim WS As New Entidades.wsEDoc_ConsultaRecepcionArchivo.WSRAD_KEY_ARCHIVO
                    WS.Url = _WS_RecepcionArchivo

                    rsboApp.StatusBar.SetText(NombreAddon + " - Procesando Claves de Acceso...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    mensaje = ""
                    If WS.CargarClavesAcceso(_WS_RecepcionClave, Claves.ToArray, mensaje) Then
                        rsboApp.StatusBar.SetText(NombreAddon + " - Proceso terminado Correctamente, verifique sus documentos recibidos despues de un momento..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        oForm.Items.Item("obtnPro").Visible = False
                        oForm.Items.Item("2").Left = oForm.Items.Item("obtnPro").Left
                        Dim oB As SAPbouiCOM.Button
                        oB = oForm.Items.Item("2").Specific
                        oB.Caption = "OK"

                    Else
                        If Not mensaje = "" Then
                            rsboApp.StatusBar.SetText(NombreAddon + " - Ocurrio un error al procesar las claves de acceso.. Error: " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                Catch ex As Exception

                End Try
               
            End If
        End If

    End Sub

    Private Sub ProcesoBtnExaminarPart()

        Try
            rsboApp.StatusBar.SetText(NombreAddon + " Revise las ventanas abiertas (ALT+TAB), puede que la ventana no quedo como principal.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            GetfileThread = New Threading.Thread(AddressOf GetNombreArchivoPart, 1)
            GetfileThread.SetApartmentState(Threading.ApartmentState.STA)
            GetfileThread.Start()
        Catch ex As Exception

        End Try


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

                rsboApp.StatusBar.SetText(NombreAddon + " - Leyendo archivo, por favor espere!...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                mensaje = ""
                Claves = ListarClaveAcceso_ArchivoTXT(fi.FullName, vbTab, mensaje)
                If Claves Is Nothing Then
                    rsboApp.StatusBar.SetText(NombreAddon + " - No se pudo leer el archivo.. Error: " + mensaje.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    oForm.DataSources.DataTables.Item("dtDocs").Rows.Clear()
                    llenaGrid(Claves)
                End If


            End If

        Catch ex As Exception
            rsboApp.StatusBar.SetText(NombreAddon + " Error al cargar Licencia, " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Function ListarClaveAcceso_ArchivoTXT(ByVal RutaArchivo As String, ByVal CaracterSPLIT As Char, ByRef mensaje As String) As List(Of String)
        Dim arrText As New List(Of String)
        Try
            Dim objReader As New StreamReader(RutaArchivo)
            Dim sLine As String = String.Empty
            Do
                sLine = objReader.ReadLine()
                If sLine IsNot Nothing Then
                    sLine = sLine.ToLower
                    For Each ClaveAcceso In sLine.Split(CaracterSPLIT)
                        If ClaveAcceso.Length = 49 Then ClaveAcceso = ExtraerSoloNumeros(ClaveAcceso)
                        If ClaveAcceso.Length = 49 Then If Not arrText.Contains(ClaveAcceso) Then arrText.Add(ClaveAcceso)

                    Next
                End If
            Loop Until sLine Is Nothing
            objReader.Dispose()
            objReader.Close()
            Return arrText
        Catch ex As Exception
            mensaje = ex.Message
            Return Nothing
        End Try
    End Function

    Private Function ExtraerSoloNumeros(ByVal strCadena As String) As String
        Dim sbSoloNumero As New StringBuilder
        If strCadena IsNot Nothing Then
            For Each caracter As Char In strCadena.Trim
                If IsNumeric(caracter) Then sbSoloNumero.Append(caracter)
            Next
        End If
        Return sbSoloNumero.ToString
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

    Private Sub llenaGrid(Claves As List(Of String))
        Try
            oForm.Freeze(True)
          
            oForm.DataSources.DataTables.Item("dtDocs").Rows.Add(Claves.Count)
            Dim i As Integer = 0
            For Each ClaveAcceso In Claves
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Count", i, (i + 1).ToString())
                oForm.DataSources.DataTables.Item("dtDocs").SetValue("Clave", i, ClaveAcceso)
                i += 1
            Next

            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("oGrid").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDocs")

            oGrid.Columns.Item(0).Description = "#"
            oGrid.Columns.Item(0).TitleObject.Caption = "#"
            oGrid.Columns.Item(0).Editable = False

            oGrid.Columns.Item(1).Description = "Clave de Acceso"
            oGrid.Columns.Item(1).TitleObject.Caption = "Clave de Acceso"
            oGrid.Columns.Item(1).Editable = False

            rsboApp.StatusBar.SetText(NombreAddon + " - " + Claves.Count.ToString() + " Claves de Acceso Cargadas, Dele click en Procesar para subir los documentos!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

        Finally
            oForm.Freeze(False)
        End Try
    End Sub


End Class

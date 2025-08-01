Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Windows.Forms

Public Class FuncionesB1

#Region ">>> Variables Globales de Opciones <<<"

    Dim mCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company
    Private fCompany As SAPbobsCOM.Company
    Private fPappl As SAPbouiCOM.Application

    ''' <summary>
    ''' Versión de la biblioteca de funciones
    ''' </summary>
    ''' <remarks></remarks>
    Public Const VersionDeFunctions As String = "20150308"

    Dim oForm As SAPbouiCOM.Form

    ' CARPETAS

    ''' <summary>
    ''' Carpeta dentro de la ruta del add-on que contiene los archivos XML de los formularios
    ''' </summary>
    ''' <remarks></remarks>
    Public carpetaFormularios As String = "Forms"
    ''' <summary>
    ''' Carpeta dentro de la ruta del add-on que contiene los archivos XML de los reportes
    ''' </summary>
    ''' <remarks></remarks>
    Public carpetaReportes As String = "Forms"
    ''' <summary>
    ''' Carpeta dentro de la ruta del add-on que contiene los archivos de imágen a utilizar
    ''' </summary>
    ''' <remarks></remarks>
    Public carpetaImagenes As String = "Forms"
    ''' <summary>
    ''' Nombre del grupo de query en el que se guardarán los querys de Búsquedas Formateadas
    ''' </summary>
    ''' <remarks></remarks>
    Public grupoQueryBusqF As String = "Busquedas Formateadas"


    ' MENSAJES

    ''' <summary>
    ''' Establece si se deben registrar errores en un archivo de texto en C:"
    ''' </summary>
    ''' <remarks></remarks>
    Public mantenerLogErrores As Boolean = False
    ''' <summary>
    ''' Si está activo, se mostrarán mensajes al tener éxito en las operaciones.
    ''' </summary>
    ''' <remarks></remarks>
    Public mostrarMensajesExito As Boolean = False
    ''' <summary>
    ''' Si está activo, se mostrarán mensajes al presentarse errores en las operaciones.
    ''' </summary>
    ''' <remarks></remarks>
    Public mostrarMensajesError As Boolean = False

    Public BD As String = Nothing
    Public TipoServer As String = Nothing
    Public NombreAddon As String = ""
#End Region

#Region ">>> Enumeradores <<<"

    ''' <summary>
    ''' Enumerador de tipo de documento. Indica el tipo de documento que se debe evaluar.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum T_Doc
        FacturaReserva = 1
        FacturaInmediata = 2
        NotaDebito = 3
        Cotizacion = 4
        NotaCredito = 5
        DocumentoPreliminar = 6
    End Enum
    ''' <summary>
    ''' Enumerador de tipo de formulario. indica si un formulario es Estándar o de Usuario.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum T_Form
        Standard = 1
        Usuario = 2
    End Enum
    ''' <summary>
    ''' Enumerador de tipos de código de cuenta
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum fCodigosDeCuenta
        AcctCode = 0
        AcctName = 1
        FormatCode = 2
        SegmentedCode = 3
    End Enum

#End Region

    ''' <summary>
    ''' Instancia la Biblioteca de funciones
    ''' </summary>
    ''' <param name="objetfCompany">Un objeto SAPbobsCOM.Company instanciado para ser usado en las funciones</param>
    ''' <param name="mostrarErrores">Indica si las funciones deben mostrar mensajes al presentar errores</param>
    ''' <param name="mostrarExito">Indica si las funciones deben mostrar mensajes al tener éxito en los procesos</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal objetfCompany As SAPbobsCOM.Company, ByVal objectApplication As SAPbouiCOM.Application, Optional ByVal mostrarErrores As Boolean = True, Optional ByVal mostrarExito As Boolean = False, Optional ByVal sNombreAddon As String = "")
        Try
            fCompany = objetfCompany
            fPappl = objectApplication
            NombreAddon = sNombreAddon

            mostrarMensajesError = mostrarErrores
            mostrarMensajesExito = mostrarExito

            BD = fCompany.CompanyDB
            TipoServer = fCompany.DbServerType
        Catch ex As Exception
        Finally
            If mostrarMensajesExito Then
                If isInSpanish() Then
                    fPappl.StatusBar.SetText("SOLSAP Functions conectó con éxito")
                Else
                    fPappl.StatusBar.SetText("SOLSAP Functions connected succesfully")
                End If
            End If
        End Try
    End Sub


    ' OBJETOS

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
            Utilitario.Util_Log.Escribir_Log("Release Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
    End Sub


    ''' <summary>
    ''' Indica si la clase del add-on se encuentra instalada en la BD, en base a la tabla CMP_SETUP.
    ''' </summary>
    ''' <param name="addOnName">Nombre que identifica la clase</param>
    ''' <param name="addOnVersion">Versión de la clase</param>
    ''' <returns>Devuelve verdadero si el addon ya está instalado y falso si no se encuentra</returns>
    ''' <remarks></remarks>
    Public Function validarVersion(ByVal addOnName As String, ByVal addOnVersion As String) As Boolean
        Dim retorno As Boolean = False
        Try
            '1. Si LA TABLA no existe la creo
            If Not checkCampoBD("@SS_SETUP", "SS_VERS") Then
                creaTablaMD("SS_SETUP", "Setup de AddOn's de SOLSAP", BoUTBTableType.bott_NoObject)
                creaCampoMD("SS_SETUP", "SS_ADDN", "Nombre del AddOn", BoFieldTypes.db_Alpha, , 100)
                creaCampoMD("SS_SETUP", "SS_VERS", "Version del AddOn", BoFieldTypes.db_Alpha, , 100)
                creaCampoMD("SS_SETUP", "SS_LICEN", "Licencia", BoFieldTypes.db_Memo, , 1000)
            Else
                '2. Valido que los datos de add-on y versión coincidan
                Dim VRS As SAPbobsCOM.Recordset
                If fCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                    VRS = getRecordSet("SELECT * FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" & addOnName & "' ORDER BY ""U_SS_VERS"" DESC")
                Else
                    VRS = getRecordSet("SELECT * FROM [@SS_SETUP] WHERE U_SS_ADDN = '" & addOnName & "' ORDER BY U_SS_VERS DESC")
                End If

                '3. Si coinciden retorno true, de lo contrario false
                If VRS.EoF Then
                    fPappl.MessageBox("Se creará la estructura de datos para el Add-On " & addOnName)
                Else
                    If VRS.Fields.Item("U_SS_VERS").Value.ToString < addOnVersion Then
                        fPappl.MessageBox("Se actualizará la estructura de datos para el Add-On " & addOnName & " de versión " & VRS.Fields.Item("U_SS_VERS").Value.ToString & " a " & addOnVersion)
                    ElseIf VRS.Fields.Item("U_SS_VERS").Value.ToString > addOnVersion Then
                        fPappl.MessageBox("Se detectó una versión del Add-On " & addOnName & " más avanzada (" & VRS.Fields.Item("U_SS_VERS").Value.ToString & ") instalada previamente. No se recomienda el uso de la versión que está intentando ejecutar (" & addOnVersion & ")")
                        retorno = True
                    ElseIf VRS.Fields.Item("U_SS_VERS").Value.ToString = addOnVersion Then
                        retorno = True
                    End If
                End If
                Release(VRS)
            End If
        Catch ex As Exception
            If mostrarMensajesError Then fPappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
        Return retorno
    End Function

    Public Sub validarVersion_SoloCrearTabla()

        Try
            '1. Si LA TABLA no existe la creo
            If Not checkCampoBD("@SS_SETUP", "SS_VERS") Then
                creaTablaMD("SS_SETUP", "Setup de AddOn's de SOLSAP", BoUTBTableType.bott_NoObject)
                creaCampoMD("SS_SETUP", "SS_ADDN", "Nombre del AddOn", BoFieldTypes.db_Alpha, , 100)
                creaCampoMD("SS_SETUP", "SS_VERS", "Version del AddOn", BoFieldTypes.db_Alpha, , 100)
                creaCampoMD("SS_SETUP", "SS_LICEN", "Licencia", BoFieldTypes.db_Memo, , 1000)
            End If
        Catch ex As Exception
            If mostrarMensajesError Then fPappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub


#Disable Warning BC42300 ' El bloque de comentario XML debe preceder directamente al elemento de lenguaje al que se aplica. Se omitirá el comentario XML.
    ''' <summary>
    ''' Ingresa la versión de la clase a la tabla CMP_SETUP.
    ''' </summary>
    ''' <param name="addOnName">Nombre de la clase</param>
    ''' <param name="addOnVersion">Versión de la clase</param>
    ''' <remarks></remarks>
    'Public Sub confirmarVersion(ByVal addOnName As String, ByVal addOnVersion As String)
    '    Try
    '        ' Ejecuto insert a la tabla anexando data de add-on, versión y éxito al crear.
    '        ' De esta forma cuando se vuelva a ejecutar el add-on no creará los campos.
    '        Dim strSQL As String = ""
    '        If fCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
    '            strSQL += "INSERT INTO ""@SS_SETUP"" "
    '            strSQL += "(""Code"", ""Name"", ""U_SS_ADDN"" ,""U_SS_VERS"") VALUES "
    '            strSQL += "(" & getCorrelativo("Code", "[@SS_SETUP]", , 1) & ", '" & getCorrelativo("Code", "[@SS_SETUP]", , 1) & "', '" & addOnName & "','" & addOnVersion & "')"
    '        Else
    '            strSQL += "INSERT INTO [@SS_SETUP] "
    '            strSQL += "(Code, Name, [U_SS_ADDN] ,[U_SS_VERS]) VALUES "
    '            strSQL += "(" & getCorrelativo("Code", "[@SS_SETUP]", , 1) & ", '" & getCorrelativo("Code", "[@SS_SETUP]", , 1) & "', '" & addOnName & "','" & addOnVersion & "')"
    '        End If

    '        Dim RSV As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
    '        RSV.DoQuery(strSQL)
    '        Release(RSV)
    '    Catch ex As Exception
    '        If mostrarMensajesError Then fPappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    'End Sub
    Public Sub confirmarVersion(ByVal addOnName As String, ByVal addOnVersion As String)
#Enable Warning BC42300 ' El bloque de comentario XML debe preceder directamente al elemento de lenguaje al que se aplica. Se omitirá el comentario XML.
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim oUserTable As SAPbobsCOM.UserTable
        Try
            '// set the object with the requested table
            oUserTable = fCompany.UserTables.Item("SS_SETUP")

            '// set the two default fields 
            oUserTable.Code = getCorrelativo("Code", """@SS_SETUP""", , 1)
            oUserTable.Name = getCorrelativo("Code", """@SS_SETUP""", , 1)

            oUserTable.UserFields.Fields.Item("U_SS_ADDN").Value = addOnName.ToString()
            oUserTable.UserFields.Fields.Item("U_SS_VERS").Value = addOnVersion.ToString()

            ' OBTENGO LA ULTIMA VERSIÓN QUE ESTUVO REGISTRADA EN EL ADDON
            Dim sQueryLicencia As String = ""
            If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryLicencia = "SELECT TOP 1  ""U_SS_LICEN"" FROM ""@SS_SETUP"" WHERE ""U_SS_ADDN"" = '" + addOnName + "' ORDER BY ""Code"" DESC"
            Else
                sQueryLicencia = "SELECT TOP 1  U_SS_LICEN FROM ""@SS_SETUP"" WHERE U_SS_ADDN = '" + addOnName + "' ORDER BY Code DESC"
            End If

            Dim Result As String = getRSvalue(sQueryLicencia, "U_SS_LICEN", "").ToString()
            If Not Result = "" Then
                oUserTable.UserFields.Fields.Item("U_SS_LICEN").Value = Result.ToString()
            End If

            oUserTable.Add()
            '// Check for errors
            fCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                fPappl.StatusBar.SetText(addOnName + " Error al confirmar Version: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                fPappl.StatusBar.SetText(addOnName + "  Versión Confirmada", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            oUserTable = Nothing
            System.GC.Collect()
        End Try
    End Sub

    ' XML


    ''' <summary>
    ''' Levanta un formulario desde un archivo XML ubicado en la carpeta de formularios del Add-On.
    ''' </summary>
    ''' <param name="FileName">Nombre del archivo (sin la extensión .srf) del formulario.</param>
    ''' <param name="cerrarSiExiste">Si el formulario se encuentra levantado, lo cierra.</param>
    ''' <remarks></remarks>
    Public Sub cargaFormXML(ByVal FileName As String, Optional ByVal cerrarSiExiste As Boolean = False)
        Try
            If cerrarSiExiste Then
                fPappl.Forms.Item(FileName.ToString).Close()
            Else
                Dim oXmlDoc As Xml.XmlDocument
                Dim sXmlFileName As String
                oXmlDoc = New Xml.XmlDocument
                sXmlFileName = System.Windows.Forms.Application.StartupPath & "\" & carpetaFormularios & "\" & FileName & ".srf"
                oXmlDoc.Load(sXmlFileName)
                fPappl.LoadBatchActions(CStr(oXmlDoc.InnerXml))
            End If

        Catch ex As Exception
            If mostrarMensajesError Then fPappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            fPappl.Forms.Item(FileName).Close()
        End Try
    End Sub

    ''' <summary>
    ''' Importa a B1 un layout desde un archivo XML.
    ''' </summary>
    ''' <param name="Report">Nombre del archivo (sin la extensión .xml)</param>
    ''' <param name="igualarQuery">Indica si se debe sobreescribir el query del layout por el que se encuentra en el UserQuery del mismo nombre</param>
    ''' <remarks></remarks>
    Public Sub cargaReportXML(ByVal Report As String, Optional ByVal igualarQuery As Boolean = False)
        Try

            Dim ooFuncionesB1Srv As SAPbobsCOM.CompanyService
            Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService
            Dim oReportLayoutParam As SAPbobsCOM.ReportLayoutParams
            Dim oReportLayout As SAPbobsCOM.ReportLayout
            Dim sXmlFileName As String
            sXmlFileName = System.Windows.Forms.Application.StartupPath & "\" & carpetaReportes & "\" & Report & ".xml"

            ooFuncionesB1Srv = fCompany.GetCompanyService
            oReportLayoutService = ooFuncionesB1Srv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
            Dim oLocalX As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalX.DoQuery("select DocName as total from rdoc where DocName = '" & Report & "'")

            If oLocalX.EoF Then
                oReportLayout = oReportLayoutService.GetDataInterfaceFromXMLFile(sXmlFileName)
                oReportLayoutParam = oReportLayoutService.AddReportLayout(oReportLayout)

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oReportLayoutParam)
                oReportLayoutParam = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oReportLayout)
                oReportLayout = Nothing
            End If
            oLocalX = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ooFuncionesB1Srv)
            ooFuncionesB1Srv = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oReportLayoutService)
            oReportLayoutService = Nothing
            GC.Collect()

            If igualarQuery Then
                Dim myQ As String = "UPDATE rdoc SET RDOC.QString = OUQR.QString " &
                                    "FROM RDOC INNER JOIN OUQR ON RDOC.DocName = OUQR.QName " &
                                    "WHERE DocName = '" & Report & "'"
                Dim rsQ As SAPbobsCOM.Recordset = getRecordSet(myQ)
                Release(rsQ)
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Exporta un layout de B1 a un archivo XML.
    ''' </summary>
    ''' <param name="Report">Nombre del archivo (sin la extensión .xml)</param>
    ''' <remarks></remarks>
    Public Sub exportReportXML(ByVal Report As String)
        Try
            Dim ooFuncionesB1Srv As SAPbobsCOM.CompanyService
            Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService
            Dim oReportLayoutParam As SAPbobsCOM.ReportLayoutParams
            Dim oReportLayout As SAPbobsCOM.ReportLayout
            'get company service
            ooFuncionesB1Srv = fCompany.GetCompanyService
            'get report layout service
            oReportLayoutService = ooFuncionesB1Srv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
            'get Report Layout Param
            oReportLayoutParam = oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams)
            'set the report layout code
            oReportLayoutParam.LayoutCode = Report
            'get the report layout using layout code
            oReportLayout = oReportLayoutService.GetReportLayout(oReportLayoutParam)
            ' ,  , , , ,, 02, 03, 17
            Dim strSQLx As String = ""
            strSQLx = System.Windows.Forms.Application.StartupPath & "\" & carpetaReportes & "\" & Report & ".xml"
            oReportLayout.ToXMLFile(strSQLx)

        Catch ex As Exception
        End Try
    End Sub


    ' CREACIONES

    ''' <summary>
    ''' Crea una tabla de usuario (UDT) en B1.
    ''' </summary>
    ''' <param name="NbTabla">Código de la tabla (max 8 caracteres)</param>
    ''' <param name="DescTabla">Descripción de la tabla (30 caracteres)</param>
    ''' <param name="TablaTipo">Tipo de tabla</param>
    ''' <remarks></remarks>
    Public Sub creaTablaMD(ByVal NbTabla As String, ByVal DescTabla As String, ByVal TablaTipo As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try

            Dim iVer As Integer = 0
            oUserTablesMD = Nothing
            oUserTablesMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            If Not oUserTablesMD.GetByKey(NbTabla) Then

                Dim tablaACrear As SAPbobsCOM.UserTablesMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                tablaACrear.TableName = Format(NbTabla)
                tablaACrear.TableDescription = Format(DescTabla)
                tablaACrear.TableType = TablaTipo

                Dim retX As Integer = 0
                Dim strSQLx As String = ""
                retX = tablaACrear.Add
                If Not retX = 0 Then
                    iVer = iVer + 1
                    fCompany.GetLastError(retX, strSQLx)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas, Nombre Tabla: " + NbTabla + "- Descripcion Tabla: " + DescTabla + "-Error: " + strSQLx, "FuncionesB1")
                    'If mantenerLogErrores Then System.IO.File.AppendAllText("C:\LogCreaTabla_" & Replace(fPappl.Company.ServerDate, "/", "-") & ".txt", Trim$(strSQLx) & vbCrLf)
                Else
                    If mostrarMensajesExito Then fPappl.StatusBar.SetText("Tabla " & NbTabla & " creada con éxito", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
                Release(tablaACrear)

            End If
            Release(oUserTablesMD)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas_Catch, " + NbTabla + "-" + DescTabla + "-Error: " + ex.Message.ToString(), "FuncionesB1")


            oUserTablesMD = Nothing

        Finally
            GC.Collect()
        End Try
    End Sub


    ''' <summary>
    ''' Crea un campo de usuario (UDF) en B1.
    ''' </summary>
    ''' <param name="NbTabla">Nombre de la tabla en la que se creará el campo (sin arroba)</param>
    ''' <param name="NbCampo">Código del campo a crear (8 caracteres)</param>
    ''' <param name="DescCampo">Descripción del campo (30 caracteres)</param>
    ''' <param name="TipoDato">Establece el tipo de dato que almacenará el campo</param>
    ''' <param name="subtipo">Sub-Tipo de campo</param>
    ''' <param name="Tamaño">Tamaño del campo</param>
    ''' <param name="Obligatorio">Establece si el campo admite o no valores nulos (requiere que se establezca un valor por defecto)</param>
    ''' <param name="validValues">Arreglo de valores string que contiene los valores válidos para el campo</param>
    ''' <param name="validDescription">Arreglo de valores string que contiene descripciones para los valores válidos para el campo</param>
    ''' <param name="valorPorDef">El valor que tomará el campo por defecto (debe ser un miembro de la lista de valores válidos)</param>
    ''' <param name="tablaVinculada">Nombre de la tabla de usuario de la cual se obtendrán los valores para el campo (sin arroba)</param>
    ''' <remarks>DI API Type - SubType, db_Alpha - st_None, db_Alpha - st_Address, db_Alpha - st_Phone, db_Memo - st_None, db_Numeric - st_None, db_Date st_None, db_Date st_Time, db_Float	st_Rate, db_Float st_Sum, db_Float st_Price, db_Float st_Quantity, db_Float st_Percentage, db_Float st_Measurement, db_Memo st_Link, db_Alpha st_Image </remarks>


    Public Sub creaCampoMD(ByVal NbTabla As String, ByVal NbCampo As String, ByVal DescCampo As String, ByVal TipoDato As SAPbobsCOM.BoFieldTypes, Optional ByVal subtipo As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal Tamaño As Integer = 10, Optional ByVal Obligatorio As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal validValues As String() = Nothing, Optional ByVal validDescription As String() = Nothing, Optional ByVal valorPorDef As String = "", Optional ByVal tablaVinculada As String = "", Optional ByVal tablaUDO As String = "", Optional ByVal ObjetoVinculado As String = "")
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Try
            'Dim NbTablaA As String = "@" & NbTabla
            'Dim NbCampoU As String = "U_" & NbCampo

            If checkCampoBD(NbTabla, NbCampo) = False Then
                oUserFieldsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldsMD.TableName = NbTabla
                oUserFieldsMD.Name = NbCampo
                oUserFieldsMD.Description = DescCampo
                oUserFieldsMD.Type = TipoDato

                If TipoDato <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    oUserFieldsMD.EditSize = Tamaño
                End If

                If TipoDato = SAPbobsCOM.BoFieldTypes.db_Float Then
                    oUserFieldsMD.SubType = subtipo
                End If

                If TipoDato = SAPbobsCOM.BoFieldTypes.db_Date Then
                    oUserFieldsMD.SubType = subtipo
                End If

                If TipoDato = SAPbobsCOM.BoFieldTypes.db_Alpha Then
                    oUserFieldsMD.SubType = subtipo
                End If

                If tablaVinculada <> "" Then
                    oUserFieldsMD.LinkedTable = tablaVinculada


                    'ElseIf ObjetoVinculado <> "" Then
                    '    If CInt(fCompany.Version.ToString.Substring(0, 2)) = 10 Then
                    '        oUserFieldsMD.LinkedSystemObject = ObjetoVinculado
                    '    End If

                ElseIf tablaUDO <> "" Then
                    oUserFieldsMD.LinkedUDO = tablaUDO

                Else

                    If Not validValues Is Nothing Then
                        For i As Integer = 0 To validValues.Length - 1
                            If validDescription Is Nothing Then
                                oUserFieldsMD.ValidValues.Description = validValues(i)
                            Else
                                oUserFieldsMD.ValidValues.Description = validDescription(i)
                            End If
                            oUserFieldsMD.ValidValues.Value = validValues(i)
                            oUserFieldsMD.ValidValues.Add()
                        Next
                    End If

                    If valorPorDef <> "" Then
                        oUserFieldsMD.DefaultValue = valorPorDef
                    End If

                    If Obligatorio = SAPbobsCOM.BoYesNoEnum.tYES Then
                        oUserFieldsMD.Mandatory = Obligatorio
                    End If
                End If

                Dim retX As Integer = 0
                Dim strSQLx As String = ""
                retX = oUserFieldsMD.Add()

                If retX <> 0 Then
                    fCompany.GetLastError(retX, strSQLx)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos, Nombre Tabla: " + NbTabla + "- Nombre Campo: " + NbCampo + "- Descripcion Campo: " + DescCampo + "-Error: " + strSQLx, "FuncionesB1")
                    'If mostrarMensajesError Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": " & strSQLx, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
                Else
                    If mostrarMensajesExito Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": Creado con éxito", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                oUserFieldsMD = Nothing
                GC.Collect()

                Exit Sub
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_Catch, Nombre Tabla: " + NbTabla + "- Nombre Campo: " + NbCampo + "- Descripcion Campo: " + DescCampo + "-Error: " + ex.Message.ToString(), "FuncionesB1")
        End Try

    End Sub

    'Public Sub creaCampoMD(ByVal NbTabla As String, ByVal NbCampo As String, ByVal DescCampo As String, ByVal TipoDato As SAPbobsCOM.BoFieldTypes, Optional ByVal subtipo As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal Tamaño As Integer = 10, Optional ByVal Obligatorio As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal validValues As String() = Nothing, Optional ByVal validDescription As String() = Nothing, Optional ByVal valorPorDef As String = "", Optional ByVal tablaVinculada As String = "")
    '    Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
    '    Try
    '        'Dim NbTablaA As String = "@" & NbTabla
    '        'Dim NbCampoU As String = "U_" & NbCampo

    '        If checkCampoBD(NbTabla, NbCampo) = False Then
    '            oUserFieldsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

    '            oUserFieldsMD.TableName = NbTabla
    '            oUserFieldsMD.Name = NbCampo
    '            oUserFieldsMD.Description = DescCampo
    '            oUserFieldsMD.Type = TipoDato

    '            If TipoDato <> SAPbobsCOM.BoFieldTypes.db_Date Then
    '                oUserFieldsMD.EditSize = Tamaño
    '            End If

    '            If TipoDato = SAPbobsCOM.BoFieldTypes.db_Float Then
    '                oUserFieldsMD.SubType = subtipo
    '            End If

    '            If TipoDato = SAPbobsCOM.BoFieldTypes.db_Date Then
    '                oUserFieldsMD.SubType = subtipo
    '            End If

    '            If TipoDato = SAPbobsCOM.BoFieldTypes.db_Alpha Then
    '                oUserFieldsMD.SubType = subtipo
    '            End If

    '            If tablaVinculada <> "" Then
    '                oUserFieldsMD.LinkedTable = tablaVinculada
    '            Else
    '                If Not validValues Is Nothing Then
    '                    For i As Integer = 0 To validValues.Length - 1
    '                        If validDescription Is Nothing Then
    '                            oUserFieldsMD.ValidValues.Description = validValues(i)
    '                        Else
    '                            oUserFieldsMD.ValidValues.Description = validDescription(i)
    '                        End If
    '                        oUserFieldsMD.ValidValues.Value = validValues(i)
    '                        oUserFieldsMD.ValidValues.Add()
    '                    Next
    '                End If

    '                If valorPorDef <> "" Then
    '                    oUserFieldsMD.DefaultValue = valorPorDef
    '                End If

    '                If Obligatorio = SAPbobsCOM.BoYesNoEnum.tYES Then
    '                    oUserFieldsMD.Mandatory = Obligatorio
    '                End If
    '            End If

    '            Dim retX As Integer = 0
    '            Dim strSQLx As String = ""
    '            retX = oUserFieldsMD.Add()

    '            If retX <> 0 Then
    '                fCompany.GetLastError(retX, strSQLx)
    '                Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos, Nombre Tabla: " + NbTabla + "- Nombre Campo: " + NbCampo + "- Descripcion Campo: " + DescCampo + "-Error: " + strSQLx, "FuncionesB1")
    '                'If mostrarMensajesError Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": " & strSQLx, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
    '            Else
    '                If mostrarMensajesExito Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": Creado con éxito", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '            End If
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
    '            oUserFieldsMD = Nothing
    '            GC.Collect()

    '            Exit Sub
    '        End If

    '    Catch ex As Exception
    '        Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_Catch, Nombre Tabla: " + NbTabla + "- Nombre Campo: " + NbCampo + "- Descripcion Campo: " + DescCampo + "-Error: " + ex.Message.ToString(), "FuncionesB1")
    '    End Try

    'End Sub

    ''' <summary>
    ''' Verifica si un campo existe en la Base de datos
    ''' </summary>
    ''' <param name="Tabla">Nombre de la tabla (incluyendo arroba)</param>
    ''' <param name="Campo">Nombre del campo (incluyendo prefijo U_)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function checkCampoBD(ByVal Tabla As String, ByVal Campo As String) As Boolean
        Dim retorno As Boolean = False
        Dim strSQLBD As String = ""
        Dim oLocalBD As SAPbobsCOM.Recordset = Nothing

        Try
            oLocalBD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If fCompany.DbServerType = "9" Then
                strSQLBD = " SELECT ""TableID""  FROM """ & fCompany.CompanyDB & """.""CUFD""  WHERE ""TableID"" ='" & Tabla & "' AND ""AliasID"" = '" & Campo & "'"
            Else
                strSQLBD = "SELECT TableID  FROM CUFD  WHERE TableID ='" & Tabla & "' AND AliasID = '" & Campo & "'"
            End If
            'strSQLBD = "SELECT column_name "
            'strSQLBD &= "FROM [" & fCompany.CompanyDB & "].INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = '" & Campo & "' AND Table_Name ='" & Tabla & "'"
            oLocalBD.DoQuery(strSQLBD)
            If oLocalBD.EoF = False Then
                retorno = True
            End If
            Release(oLocalBD)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_checkCampoBD_Catch, Query: " + strSQLBD, "FuncionesB1")
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_checkCampoBD_Catch, Nombre Tabla: " + Tabla + "- Nombre Campo: " + Campo + "-Error: " + ex.Message.ToString(), "FuncionesB1")
        End Try
        Return retorno
    End Function

    ' ACTUALIZA CAMPO

    Public Function checkCampoBD_ID(ByVal Tabla As String, ByVal Campo As String) As String
        Dim retorno As String = ""
        Dim strSQLBD As String = ""
        Dim oLocalBD As SAPbobsCOM.Recordset = Nothing
        'Dim oFCs As SAPbobsCOM.UserFieldsMD
        Try
            'oLocalBD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ''strSQLBD = "SELECT TableID, FieldID, AliasID, Descr   FROM cufd  WHERE TableID ='" & Tabla & "' AND AliasID = '" & Campo & "'  "
            ' ''strSQLBD = "SELECT column_name "
            ' ''strSQLBD &= "FROM [" & fCompany.CompanyDB & "].INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = '" & Campo & "' AND Table_Name ='" & Tabla & "'"

            'oFCs = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            'oLocalBD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oLocalBD.DoQuery("SELECT * FROM cufd  WHERE TableID ='" & Tabla & "' AND AliasID = '" & Campo & "'  ")
            'oFCs.Browser.Recordset = oLocalBD
            'oFCs.Browser.MoveFirst()
            'While oFCs.Browser.EoF = False
            '    retorno = oFCs.FieldID
            '    oFCs.Browser.MoveNext()
            'End While
            ''oLocalBD.DoQuery(strSQLBD)
            ''If oLocalBD.EoF = False Then
            ''    retorno = oLocalBD.Fields.Item(1).FieldID
            ''Else
            ''    retorno = ""
            ''End If
            'Release(oLocalBD)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oFCs)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oLocalBD)
            'oFCs = Nothing
            'oLocalBD = Nothing
            'GC.Collect()
            'GC.WaitForPendingFinalizers()

            oLocalBD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If fCompany.DbServerType = "9" Then
                strSQLBD = " SELECT ""FieldID""  FROM """ & fCompany.CompanyDB & """.""CUFD""  WHERE ""TableID"" ='" & Tabla & "' AND ""AliasID"" = '" & Campo & "'"
            Else
                strSQLBD = "SELECT FieldID  FROM CUFD  WHERE TableID ='" & Tabla & "' AND AliasID = '" & Campo & "'"
            End If

            oLocalBD.DoQuery(strSQLBD)

            While Not oLocalBD.EoF
                retorno = oLocalBD.Fields.Item("FieldID").Value
                oLocalBD.MoveNext()
            End While

            Release(oLocalBD)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oLocalBD)
            oLocalBD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception

        End Try
        Return retorno
    End Function

    Public Sub ActualizaCampos(ByVal NbTabla As String, ByVal NbCampo As String, ByVal DescCampo As String, ByVal TipoDato As SAPbobsCOM.BoFieldTypes, Optional ByVal subtipo As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal Tamaño As Integer = 10, Optional ByVal Obligatorio As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal validValues As String() = Nothing, Optional ByVal validDescription As String() = Nothing, Optional ByVal valorPorDef As String = "", Optional ByVal tablaVinculada As String = "")

        Dim oDatos As SAPbobsCOM.UserFieldsMD = Nothing
        Dim Err As Integer
        Dim errms As String = ""
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
        Try

            Dim m As String = checkCampoBD_ID(NbTabla, NbCampo)
            If m <> "" Then
                oUserFieldsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUserFieldsMD.GetByKey(NbTabla, CInt(m))
                oUserFieldsMD.TableName = NbTabla
                oUserFieldsMD.Name = NbCampo
                oUserFieldsMD.Description = DescCampo
                oUserFieldsMD.Type = TipoDato

                If TipoDato <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    oUserFieldsMD.EditSize = Tamaño
                End If

                If TipoDato = SAPbobsCOM.BoFieldTypes.db_Float Then
                    oUserFieldsMD.SubType = subtipo
                End If

                If TipoDato = SAPbobsCOM.BoFieldTypes.db_Date Then
                    oUserFieldsMD.SubType = subtipo
                End If

                If TipoDato = SAPbobsCOM.BoFieldTypes.db_Alpha Then
                    oUserFieldsMD.SubType = subtipo
                End If

                If tablaVinculada <> "" Then
                    oUserFieldsMD.LinkedTable = tablaVinculada
                Else
                    If Not validValues Is Nothing Then
                        oUserFieldsMD.ValidValues.Delete()
                        For i As Integer = 0 To validValues.Length - 1
                            If validDescription Is Nothing Then
                                oUserFieldsMD.ValidValues.Description = validValues(i)
                            Else
                                oUserFieldsMD.ValidValues.Description = validDescription(i)
                            End If
                            oUserFieldsMD.ValidValues.Value = validValues(i)
                            oUserFieldsMD.ValidValues.Add()
                        Next
                    End If

                    If valorPorDef <> "" Then
                        oUserFieldsMD.DefaultValue = valorPorDef
                    End If

                    If Obligatorio = SAPbobsCOM.BoYesNoEnum.tYES Then
                        oUserFieldsMD.Mandatory = Obligatorio
                    End If
                End If
                Err = oUserFieldsMD.Update()
            End If

            If Err <> 0 Then
                fCompany.GetLastError(Err, errms)
                If mostrarMensajesError Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": " & errms, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            Else
                If mostrarMensajesExito Then fPappl.StatusBar.SetText("Campo " & NbCampo & ": Creado con éxito", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            oUserFieldsMD = Nothing
            GC.Collect()

        Catch ex As Exception

        Finally
            Release(oDatos)
            Release(oUserFieldsMD)

        End Try

    End Sub


    ' END ACTUALIZA CAMPO


    ''' <summary>
    ''' Crea un índice (UserKey) en una tabla específica. Un índice permite validar a nivel de metadatos la no-duplicidad de un dato o combinación de estos, además de acelerar los procesos de búsqueda.
    ''' </summary>
    ''' <param name="nombreDelIndice">Código de 8 caracteres máx que distingue al índice</param>
    ''' <param name="tablaSinArroba">Nombre de la tabla (sin arroba)</param>
    ''' <param name="camposSinU">Nombre de los campos a indexar (sin prefijo U_)</param>
    ''' <param name="esUnique">Establece si el índice permite o no valores duplicados</param>
    ''' <remarks></remarks>
    Public Sub creaIndice(ByVal nombreDelIndice As String, ByVal tablaSinArroba As String, ByVal camposSinU() As String, Optional ByVal esUnique As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES)
        Try
            Dim oInd As SAPbobsCOM.UserKeysMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)
            Try
                Dim resI As Integer = 0
                Dim strErr As String = ""
                oInd.KeyName = nombreDelIndice
                oInd.TableName = tablaSinArroba
                oInd.Unique = esUnique
                For resI = 0 To camposSinU.Length - 1
                    oInd.Elements.ColumnAlias = camposSinU(resI)
                    If resI < camposSinU.Length - 1 Then oInd.Elements.Add()
                Next
                resI = oInd.Add()
                If resI <> 0 Then
                    strErr = fCompany.GetLastErrorDescription()
                End If
            Catch exxx As Exception
            Finally
                Release(oInd)
            End Try
        Catch
        End Try
    End Sub



#Disable Warning BC42307 ' El parámetro de comentario XML 'FindColumn2' no coincide con un parámetro de la instrucción 'function' correspondiente.
#Disable Warning BC42307 ' El parámetro de comentario XML 'FindColumn1' no coincide con un parámetro de la instrucción 'function' correspondiente.
    ''' <summary>
    ''' Crea un objeto definido por el usuario (UDO) en B1.
    ''' </summary>
    ''' <param name="Code">Código del UDO</param>
    ''' <param name="Name">Nombre del UDO</param>
    ''' <param name="TableName">Tabla principal del UDO (sin arroba)</param>
    ''' <param name="FindColumn1">Campo para búsqueda (sin prefijo U_)</param>
    ''' <param name="FindColumn2">Campo para búsqueda (sin prefijo U_)</param>
    ''' <param name="Cancel">Permitir Cancelar</param>
    ''' <param name="Close">Permitir Cerrar</param>
    ''' <param name="Deleted">Permitir Eliminar</param>
    ''' <param name="DefaultForm">Generar formulario por defecto</param>
    ''' <param name="Find">Permitir Buscar</param>
    ''' <param name="Log">Llevar Log</param>
    ''' <param name="objectType">Tipo de objeto correspondiente al UDO</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaUDO(ByVal Code As String, ByVal Name As String, ByVal TableName As String, Optional ByVal FindColumn As String() = Nothing, Optional ByVal Cancel As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Close As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO,
Optional ByVal Deleted As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal DefaultForm As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Find As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Log As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal objectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_MasterData) As Boolean
#Enable Warning BC42307 ' El parámetro de comentario XML 'FindColumn1' no coincide con un parámetro de la instrucción 'function' correspondiente.
#Enable Warning BC42307 ' El parámetro de comentario XML 'FindColumn2' no coincide con un parámetro de la instrucción 'function' correspondiente.
        Try
            '
            Dim oUserDataOMD As SAPbobsCOM.UserObjectsMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            oUserDataOMD.Code = Code
            oUserDataOMD.Name = Name
            oUserDataOMD.ObjectType = objectType
            oUserDataOMD.TableName = TableName
            '
            oUserDataOMD.CanCancel = Cancel
            oUserDataOMD.CanClose = Close
            oUserDataOMD.CanDelete = Deleted
            oUserDataOMD.CanCreateDefaultForm = DefaultForm
            oUserDataOMD.CanFind = Find
            oUserDataOMD.CanLog = Log
            '
            If Not FindColumn Is Nothing Then
                For FCi As Integer = 0 To FindColumn.Length - 1
                    oUserDataOMD.FindColumns.ColumnAlias = FindColumn(FCi)
                    oUserDataOMD.FindColumns.Add()
                Next
            End If

            Dim ret As Integer = 0
            Dim strSQL As String = ""
            ret = oUserDataOMD.Add
            If ret <> 0 Then
                fCompany.GetLastError(ret, strSQL)
                If mantenerLogErrores Then System.IO.File.AppendAllText("C:\LogCreaUDO_" & Replace(Date.Today, "/", "-") & ".txt", Trim$(strSQL) & vbCrLf)
            End If

            Release(oUserDataOMD)

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Crea un objeto definido por el usuario (UDO) en B1.
    ''' </summary>
    ''' <param name="Code">Código del UDO</param>
    ''' <param name="Name">Nombre del UDO</param>
    ''' <param name="TableName">Tabla principal del UDO (sin arroba)</param>
    ''' <param name="FindColumn">Arreglo de valores String que indica las columnas (sin U_) para búsqueda</param>
    ''' <param name="ChildTables">Arreglo de valores String que indica las tablas hijo (sin arroba)</param>
    ''' <param name="Cancel">Permitir Cancelar</param>
    ''' <param name="Close">Permitir Cerrar</param>
    ''' <param name="Deleted">Permitir Eliminar</param>
    ''' <param name="DefaultForm">Generar formulario por defecto</param>
    ''' <param name="Find">Permitir Buscar</param>
    ''' <param name="Log">Llevar Log</param>
    ''' <param name="objectType">Tipo de objeto correspondiente al UDO</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaUDOC(ByVal Code As String, ByVal Name As String, ByVal TableName As String, Optional ByVal FindColumn As String() = Nothing, Optional ByVal ChildTables() As String = Nothing, Optional ByVal Cancel As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Close As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO,
Optional ByVal Deleted As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal DefaultForm As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Find As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal Log As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO,
Optional ByVal objectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_MasterData) As Boolean

        Dim oUserDataOMD As SAPbobsCOM.UserObjectsMD = Nothing
        Try
            '
            oUserDataOMD = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserDataOMD.GetByKey(Code) = False Then
                oUserDataOMD.Code = Code
                oUserDataOMD.Name = Name
                oUserDataOMD.ObjectType = objectType
                oUserDataOMD.TableName = TableName
                '
                oUserDataOMD.CanCancel = Cancel
                oUserDataOMD.CanClose = Close
                oUserDataOMD.CanDelete = Deleted
                oUserDataOMD.CanCreateDefaultForm = DefaultForm
                oUserDataOMD.CanFind = Find
                oUserDataOMD.CanLog = Log

                '
                If Not FindColumn Is Nothing Then
                    For FCi As Integer = 0 To FindColumn.Length - 1
                        oUserDataOMD.FindColumns.ColumnAlias = FindColumn(FCi)
                        oUserDataOMD.FindColumns.Add()
                    Next
                End If

                If Not ChildTables Is Nothing Then
                    For CTi As Integer = 0 To ChildTables.Length - 1
                        oUserDataOMD.ChildTables.TableName = ChildTables(CTi)
                        oUserDataOMD.ChildTables.Add()
                    Next
                End If

                Dim ret As Integer = 0
                Dim strSQL As String = ""
                ret = oUserDataOMD.Add
                If ret <> 0 Then
                    fCompany.GetLastError(ret, strSQL)
                    ' If mantenerLogErrores Then addLogTxt(Trim(strSQL), "creaUDOC")
                    Utilitario.Util_Log.Escribir_Log("creaUDOC Nombre:" + Name + ", TableName: " + TableName + " Error: " + strSQL.ToString(), "FuncionesB1")
                    fPappl.StatusBar.SetText("UDO " & strSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    fPappl.StatusBar.SetText("UDO " & Code & " Se creo correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'If mantenerLogErrores Then System.IO.File.AppendAllText(System.Windows.Forms.Application.StartupPath & "\" & Replace(Date.Today, "/", "-") & ".txt", Trim$(strSQL) & vbCrLf)
                End If
            End If

            Return True
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("creaUDOC_Catch Nombre:" + Name + ", TableName: " + TableName + " Error: " + ex.Message.ToString(), "FuncionesB1")
            Return False
        Finally
            Release(oUserDataOMD)
            GC.Collect()
        End Try

    End Function


    ''' <summary>
    ''' Crea una Consulta (UserQuery) en B1.
    ''' </summary>
    ''' <param name="Nombre">Nombre con el que se identificará el Query</param>
    ''' <param name="Query">Sentencia SQL de la consulta</param>
    ''' <param name="QryCat">Nombre de la categoría en la que se registrará el query</param>
    ''' <param name="creaCat">Indica si se debe crear la categoría que contiene al query</param>
    ''' <remarks></remarks>
    Public Sub creaQuery(ByVal Nombre As String, ByVal Query As String, ByVal QryCat As String, Optional ByVal creaCat As Boolean = False)
        Try
            If creaCat Then
                creaQueryCat(QryCat)
            End If
            Dim strSQLQ As String = ""
            Dim oUserQuery As SAPbobsCOM.UserQueries = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
            oUserQuery.Query = Query
            oUserQuery.QueryCategory = getIdQueryCat(QryCat)
            oUserQuery.QueryDescription = Nombre
            Dim ret As Integer = oUserQuery.Add
            If ret <> 0 Then
                fCompany.GetLastError(ret, strSQLQ)
            End If
            Release(oUserQuery)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Crea una Consulta (UserQuery) en B1. como Company
    ''' </summary>
    ''' <param name="Nombre">Nombre con el que se identificará el Query</param>
    ''' <param name="Query">Sentencia SQL de la consulta</param>
    ''' <param name="QryCat">Nombre de la categoría en la que se registrará el query</param>
    ''' <param name="creaCat">Indica si se debe crear la categoría que contiene al query</param>
    ''' <remarks></remarks>
    ''' 
    Public Sub creaQuery(ByVal Nombre As String, ByVal Query As String, ByVal QryCat As String, ByVal mCompany As SAPbobsCOM.Company, Optional ByVal creaCat As Boolean = False)
        Try
            If creaCat Then
                creaQueryCat(QryCat, mCompany)
            End If
            Dim strSQLQ As String = ""
            Dim oUserQuery As SAPbobsCOM.UserQueries = mCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
            oUserQuery.Query = Query
            oUserQuery.QueryCategory = getIdQueryCat(QryCat, mCompany)
            oUserQuery.QueryDescription = Nombre
            Dim ret As Integer = oUserQuery.Add
            If ret <> 0 Then
                mCompany.GetLastError(ret, strSQLQ)
            End If
            Release(oUserQuery)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Crea una categoría de Querys
    ''' </summary>
    ''' <param name="grupoQuery">Nombre de la categoría</param>
    ''' <param name="permisos">Permisos por grupo para la categoría de querys</param>
    ''' <remarks></remarks>
    Public Sub creaQueryCat(ByVal grupoQuery As String, Optional ByVal permisos As String = "YYYYYYYYYYYYYYYYYYYY")
        Try
            If getIdQueryCat(grupoQuery) = -1 Then
                ' la categoría no existe. la creo
                Dim gQ As SAPbobsCOM.QueryCategories = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories)
                gQ.Name = grupoQuery
                gQ.Permissions = permisos
                gQ.Add()
                Release(gQ)
            End If
        Catch ex01 As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Crea una categoría de Querys con company
    ''' </summary>
    ''' <param name="grupoQuery">Nombre de la categoría</param>
    ''' <param name="permisos">Permisos por grupo para la categoría de querys</param>
    ''' <remarks></remarks>
    Public Sub creaQueryCat(ByVal grupoQuery As String, ByVal mCompany As SAPbobsCOM.Company, Optional ByVal permisos As String = "YYYYYYYYYYYYYYYYYYYY")
        Try
            If getIdQueryCat(grupoQuery, mCompany) = -1 Then
                ' la categoría no existe. la creo
                Dim gQ As SAPbobsCOM.QueryCategories = mCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories)
                gQ.Name = grupoQuery
                gQ.Permissions = permisos
                gQ.Add()
                Release(gQ)
            End If
        Catch ex01 As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Crea una Búsqueda Formateada
    ''' </summary>
    ''' <param name="queryName">Nombre del User Query</param>
    ''' <param name="query">Consulta SQL</param>
    ''' <param name="formID">Type del formulario</param>
    ''' <param name="itemUID">UID del item al que está vinculado la BF</param>
    ''' <param name="colUID">Columna a la que está vinculada la BF</param>
    ''' <param name="autoRefresh">Actualizar el valor automáticamente</param>
    ''' <param name="autoRefreshField">Campo que desencadena la BF</param>
    ''' <param name="borrarSiExiste">Borrar y volver a crear la BF cada vez que se inicie el Add-On</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaBusquedaF(ByVal queryName As String, ByVal query As String, ByVal formID As String, ByVal itemUID As String, Optional ByVal colUID As String = "-1", Optional ByVal autoRefresh As Boolean = False, Optional ByVal autoRefreshField As String = "", Optional ByVal borrarSiExiste As Boolean = True) As Boolean
        Dim fR As Boolean = False
        Try
            ' creo el query
            creaQuery(queryName, query, grupoQueryBusqF, True)

            ' eliminación de la BF
            Dim ret As Integer = 0
            Dim fUserBusFor2 As SAPbobsCOM.FormattedSearches
            fUserBusFor2 = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
            Dim existe As Boolean = fUserBusFor2.GetByKey(getIdBusquedaF(formID, itemUID, colUID))
            If existe And borrarSiExiste Then
                ret = fUserBusFor2.Remove()
                existe = False
            End If
            Release(fUserBusFor2)

            ' creación de la BF
            If Not existe Then
                Dim fUserBusFor As SAPbobsCOM.FormattedSearches
                fUserBusFor = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                fUserBusFor.FormID = formID
                fUserBusFor.ItemID = itemUID
                If colUID <> "-1" Then fUserBusFor.ColumnID = colUID
                fUserBusFor.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery

                fUserBusFor.QueryID = getIdQuery(queryName, getIdQueryCat(grupoQueryBusqF))
                If autoRefresh And autoRefreshField <> "" Then
                    fUserBusFor.Refresh = SAPbobsCOM.BoYesNoEnum.tYES
                    If colUID = "-1" Then
                        fUserBusFor.ByField = SAPbobsCOM.BoYesNoEnum.tYES
                    Else
                        fUserBusFor.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                    End If
                    fUserBusFor.FieldID = autoRefreshField
                    fUserBusFor.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    fUserBusFor.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                ret = fUserBusFor.Add
                If Not ret = 0 Then
                    fCompany.GetLastError(ret, query)
                End If
                Release(fUserBusFor)
            End If

            fR = True

        Catch ex As Exception
        End Try
        Return fR
    End Function

    ''' <summary>
    ''' Crea una Alerta con un query 
    ''' </summary>
    ''' <param name="NameAlert">Nombre de la Alerta</param>
    ''' <param name="NumQuery">Numero del Query</param>
    ''' <param name="UserCode">Numero de Usuario Ejemplo "manager=1"</param>
    ''' <remarks></remarks>
    Public Sub AddAlertManagement(ByVal NameAlert As String, ByVal NumQuery As Integer, Optional ByVal UserCode As Integer = 1)
        'get alert
        Dim oAlertManagement As SAPbobsCOM.AlertManagement
        Dim oAlertManagementParams As SAPbobsCOM.AlertManagementParams
        Dim oAlertManagementRecipients As SAPbobsCOM.AlertManagementRecipients
        Dim oAlertRecipient As SAPbobsCOM.AlertManagementRecipient
        Dim oAlertManagementService As SAPbobsCOM.AlertManagementService = New SAPbobsCOM.AlertManagementService
        'Assuming that oAlertManagementService is already defined!

        'Get alert
        oAlertManagement = oAlertManagementService.GetDataInterface(SAPbobsCOM.AlertManagementServiceDataInterfaces.atsdiAlertManagement)

        'set alert name
        oAlertManagement.Name = NameAlert

        'set query
        oAlertManagement.QueryID = NumQuery

        'activate the alert
        oAlertManagement.Active = SAPbobsCOM.BoYesNoEnum.tYES

        'set priority
        oAlertManagement.Priority = SAPbobsCOM.AlertManagementPriorityEnum.atp_Normal

        'Set the Frequency
        oAlertManagement.FrequencyInterval = 1

        ' set the Frequency type to hours
        oAlertManagement.FrequencyType = SAPbobsCOM.AlertManagementFrequencyType.atfi_Minutes

        'get Recipients collection
        oAlertManagementRecipients = oAlertManagement.AlertManagementDocuments

        'add recipient
        oAlertRecipient = oAlertManagementRecipients.Add()

        'set recipient code(manager=1)
        oAlertRecipient.UserCode = UserCode

        'set internal message
        oAlertRecipient.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES

        'add alert
        oAlertManagementParams = oAlertManagementService.AddAlertManagement(oAlertManagement)

    End Sub

    ' GETS

    ''' <summary>
    ''' Devuelve un objeto UserFieldsMD lleno con los datos solicitados.
    ''' </summary>
    ''' <param name="tabla">Nombre de la tabla</param>
    ''' <param name="nombreCampo">Nombre del campo (sin prefijo U_)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCampo(ByVal tabla As String, ByVal nombreCampo As String) As SAPbobsCOM.UserFieldsMD
        Try
            Dim ufRs As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ufRs.DoQuery("select FieldID from cufd where TableID = '" & tabla & "' and AliasID = '" & nombreCampo & "'")
            Dim k As Integer = 0
            k = ufRs.Fields.Item(0).Value.ToString
            Dim UFretorno As SAPbobsCOM.UserFieldsMD
            UFretorno = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            UFretorno.GetByKey(tabla, k)
            If UFretorno.Name = nombreCampo Then
                Return UFretorno
                Exit Function
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un UserQuery. Si no existe, devuelve -1.
    ''' </summary>
    ''' <param name="nombreQuery">Nombre del query</param>
    ''' <param name="idCat">ID interno de la categoría del query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdQuery(ByVal nombreQuery As String, ByVal idCat As Integer) As Integer
        Try
            Dim queryId As Integer = -1
            Dim oLocalQ As SAPbobsCOM.Recordset
            oLocalQ = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalQ.DoQuery("SELECT IntrnalKey as 'Id' FROM OUQR WHERE QName = '" & nombreQuery & "'  AND QCategory = " & idCat)
            If oLocalQ.EoF = False Then queryId = oLocalQ.Fields.Item("Id").Value
            Return queryId
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un UserQuery. Si no existe, devuelve -1. Con Company
    ''' </summary>
    ''' <param name="nombreQuery">Nombre del query</param>
    ''' <param name="idCat">ID interno de la categoría del query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdQuery(ByVal nombreQuery As String, ByVal idCat As Integer, ByVal mCompany As SAPbobsCOM.Company) As Integer
        Try
            Dim queryId As Integer = -1
            Dim oLocalQ As SAPbobsCOM.Recordset
            oLocalQ = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalQ.DoQuery("SELECT IntrnalKey as 'Id' FROM OUQR WHERE QName = '" & nombreQuery & "'  AND QCategory = " & idCat)
            If oLocalQ.EoF = False Then queryId = oLocalQ.Fields.Item("Id").Value
            Return queryId
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Devuleve el ID interno de una categoría del Query Manager. Si no existe, devuelve -1.
    ''' </summary>
    ''' <param name="nombreCat">Nombre de la categoría</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdQueryCat(ByVal nombreCat As String) As Integer
        Try
            Dim queryId As Integer = -1
            Dim oLocalQ As SAPbobsCOM.Recordset
            oLocalQ = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalQ.DoQuery("SELECT CategoryId as 'Id' FROM OQCN WHERE CatName = '" & nombreCat & "'")
            If oLocalQ.EoF = False Then queryId = oLocalQ.Fields.Item("Id").Value
            Return queryId
        Catch ex As Exception
            Return -1
        End Try
    End Function
    ''' <summary>
    ''' Devuleve el ID interno de una categoría del Query Manager. Si no existe, devuelve -1. Con company
    ''' </summary>
    ''' <param name="nombreCat">Nombre de la categoría</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdQueryCat(ByVal nombreCat As String, ByVal mCompany As SAPbobsCOM.Company) As Integer
        Try
            Dim queryId As Integer = -1
            Dim oLocalQ As SAPbobsCOM.Recordset
            oLocalQ = mCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalQ.DoQuery("SELECT CategoryId as 'Id' FROM OQCN WHERE CatName = '" & nombreCat & "'")
            If oLocalQ.EoF = False Then queryId = oLocalQ.Fields.Item("Id").Value
            Return queryId
        Catch ex As Exception
            Return -1
        End Try
    End Function


    ''' <summary>
    ''' Devuelve el ID interno de una búsqueda formateada para ser usado en un getByKey. Si no existe, devuelve -1.
    ''' </summary>
    ''' <param name="FormID_o_TYPE">Type del formulario</param>
    ''' <param name="ItemID">UID del item al cual se encuentra ligado la BF</param>
    ''' <param name="ColID">Columna a la cual se encuentra asociada la BF (si el item es una matriz)</param>
    ''' <returns>Código interno de la búsqueda formateada</returns>
    ''' <remarks></remarks>
    Public Function getIdBusquedaF(ByVal FormID_o_TYPE As String, ByVal ItemID As String, Optional ByVal ColID As String = "-1") As Integer
        Try
            Dim oLocalBF As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strSQLBF As String = "SELECT IndexID FROM CSHS WHERE FormID='" & FormID_o_TYPE & "' and ItemID='" & ItemID & "' and ColID='" & ColID & "'"
            oLocalBF.DoQuery(strSQLBF)
            If oLocalBF.EoF = False Then
                Return oLocalBF.Fields.Item(0).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Devuelve el DocEntry de un Documento de Marketing. En caso de error, devuelve -1.
    ''' </summary>
    ''' <param name="DocNum"></param>
    ''' <param name="TipoDoc"></param>
    ''' <param name="SubType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDocEntry(ByVal DocNum As String, ByVal TipoDoc As T_Doc, Optional ByVal SubType As SAPbobsCOM.BoObjectTypes = SAPbobsCOM.BoObjectTypes.oCreditNotes) As Integer
        '
        Dim oDoc As Integer = -1
        Try
            Dim oLocalDoc As SAPbobsCOM.Recordset
            Dim strSQLDoc As String = ""
            oLocalDoc = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Select Case TipoDoc
                Case T_Doc.FacturaInmediata
                    strSQLDoc = "select docentry from oinv where docsubtype='--' and docnum='" & DocNum & "' and IsIns='N'"
                Case T_Doc.FacturaReserva
                    strSQLDoc = "select docentry from oinv where docsubtype='--' and docnum='" & DocNum & "' and IsIns='Y'"
                Case T_Doc.NotaDebito
                    strSQLDoc = "select docentry from oinv where docsubtype='DN' and docnum='" & DocNum & "'"
                Case T_Doc.Cotizacion
                    strSQLDoc = "select docentry from OQUT where docnum='" & DocNum & "'"
                Case T_Doc.NotaCredito
                    strSQLDoc = "select docentry from ORIN where docnum='" & DocNum & "'"
                Case T_Doc.DocumentoPreliminar
                    strSQLDoc = "select docentry from ODRF where docnum='" & DocNum & "' and ObjType='" & SubType & "'"
            End Select

            oLocalDoc.DoQuery(strSQLDoc)
            If oLocalDoc.EoF = False Then
                oDoc = oLocalDoc.Fields.Item("docentry").Value
            End If
        Catch ex As Exception
            oDoc = -1
            'Pappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return oDoc
        ''
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un Banco. En caso de error, devuelve -1.
    ''' </summary>
    ''' <param name="BankCode">Código del Banco</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdBanco(ByVal BankCode As String) As Integer
        Dim oLocalODSC As SAPbobsCOM.Recordset
        Try
            Dim strSQL As String = "SELECT ABSENTRY FROM ODSC WHERE BANKCODE='" & BankCode & "'"
            oLocalODSC = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalODSC.DoQuery(strSQL)
            If oLocalODSC.EoF = False Then
                Return oLocalODSC.Fields.Item(0).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
        Release(oLocalODSC)
    End Function

    ''' <summary>
    ''' Devuelve el ID interno de un Campo de Usuario (UDF). En caso de error, devuelve -1.
    ''' </summary>
    ''' <param name="Tabla">Código de la tabla (con arroba)</param>
    ''' <param name="Campo">Nombre del campo (sin prefijo U_)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdUserField(ByVal Tabla As String, ByVal Campo As String) As Integer

        Dim oLocalUF As SAPbobsCOM.Recordset
        Try
            Dim strSQL As String = "SELECT FIELDID FROM CUFD WHERE TABLEID ='" & Tabla & "' AND ALIASID = '" & Campo & "'"
            oLocalUF = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLocalUF.DoQuery(strSQL)
            If oLocalUF.EoF = False Then
                Return oLocalUF.Fields.Item(0).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
        Release(oLocalUF)

    End Function

    ''' <summary>
    ''' Devuelve el formulario del Item Event
    ''' </summary>
    ''' <param name="pVal">ItemEvent</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getForm(ByVal pVal As SAPbouiCOM.ItemEvent) As SAPbouiCOM.Form
        Try
            Return fPappl.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Devuelve el formulario del Data Event
    ''' </summary>
    ''' <param name="BusinessObjectInfo">Parámetro del data event</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getForm(ByVal BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo) As SAPbouiCOM.Form
        Try
            Return fPappl.Forms.Item(BusinessObjectInfo.FormUID)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Devuelve el formulario actual
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getForm() As SAPbouiCOM.Form
        Try
            Return fPappl.Forms.ActiveForm
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuleve el Item al que hace referencia el pVal
    ''' </summary>
    ''' <param name="pVal">Parámetro del ItemEvent</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getItem(ByVal pVal As SAPbouiCOM.ItemEvent) As SAPbouiCOM.Item
        Try
            Return getForm(pVal).Items.Item(pVal.ItemUID)
        Catch ex As Exception
            Return Nothing
        End Try
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
            Utilitario.Util_Log.Escribir_Log("getRecordSet " + ex.Message.ToString, "FuncionesB1")
        End Try
        Return fRS
    End Function

    ''' <summary>
    ''' Devuelve un Recordset a partir de un query con company
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Function getRecordSet(ByVal query As String, ByVal mCompany As SAPbobsCOM.Company) As SAPbobsCOM.Recordset
        Dim fRS As SAPbobsCOM.Recordset = mCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            fRS.DoQuery(query)
        Catch ex As Exception
        End Try
        Return fRS
    End Function

    ''' <summary>
    ''' Devuelve el valor de un campo de una consulta en formato String con company
    ''' </summary>
    ''' <param name="query">Consulta SQL a ejecutar</param>
    ''' <param name="columnaRet">Columna de la consulta a retornar</param>
    ''' <param name="valorNulo">Valor a retornar en caso de error/nulo</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRSvalue(ByVal query As String, ByVal columnaRet As String, ByVal mCompany As SAPbobsCOM.Company, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query, mCompany)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRSvalue Catch:" + ex.Message().ToString() + "-QUERY: " + query, "FuncionesB1")
        End Try
        Return ret
    End Function

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
            Utilitario.Util_Log.Escribir_Log("getRSvalue-QUERY: " + query, "FuncionesB1")
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            Utilitario.Util_Log.Escribir_Log("getRSvalue-QUERY: " + query, "FuncionesB1")
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("getRSvalue Catch:" + ex.Message().ToString() + "-QUERY: " + query, "FuncionesB1")
        End Try
        Return ret
    End Function
    Public Function getRSvalue(ByVal query As String, Optional ByVal columnaRet As Integer = 0, Optional ByVal valorNulo As String = "") As String
        Dim ret As String = valorNulo
        Try
            Utilitario.Util_Log.Escribir_Log("getRSvalue-QUERY: " + query, "FuncionesB1")
            Dim r As SAPbobsCOM.Recordset = getRecordSet(query)
            ret = nzString(r.Fields.Item(columnaRet).Value, , valorNulo)
            Release(r)
        Catch ex As Exception
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' Devuelve el valor seleccionado en un combo
    ''' </summary>
    ''' <param name="combo">ComboBox instanciado</param>
    ''' <param name="returnValue">Verdadero para retornar Value, Falso para description</param>
    ''' <param name="valorSiNulo">Valor a retornar si no hya nada seleccionado</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getComboSelected(ByVal combo As SAPbouiCOM.ComboBox, Optional ByVal returnValue As Boolean = True, Optional ByVal valorSiNulo As String = "") As String
        Dim r As String = valorSiNulo
        Try
            If returnValue Then
                r = combo.Selected.Value
            Else
                r = combo.Selected.Description
            End If
        Catch ex As Exception
        End Try
        Return r
    End Function

    ''' <summary>
    ''' Captura el valor seleccionado en un CFL y lo escribe en el EditText desde el que fue invocado.
    ''' </summary>
    ''' <param name="pVal">ItemEvent</param>
    ''' <param name="columna">Columna del CFL a devolver. Si se deja en blanco devuelve la primera</param>
    ''' <remarks></remarks>
    Public Sub getCFLvalue(ByVal pVal As SAPbouiCOM.ItemEvent, Optional ByVal columna As String = "")
        Try
            Dim fcflEvent As SAPbouiCOM.ChooseFromListEvent = pVal
            If columna = "" Then
                getForm(pVal).Items.Item(pVal.ItemUID).Specific.String = fcflEvent.SelectedObjects.GetValue(0, 0)
            Else
                getForm(pVal).Items.Item(pVal.ItemUID).Specific.String = fcflEvent.SelectedObjects.GetValue(columna, 0)
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Retorna datos de cuentas contables en cualquiera de 4 formatos (Código, Nombre, FormatCode y Segmentada)
    ''' </summary>
    ''' <param name="valor">Valor de la cuenta</param>
    ''' <param name="formatoOriginal">Formato del valor que se provee</param>
    ''' <param name="formatoDestino">Formato al que se desea convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAccount(ByVal valor As String, ByVal formatoOriginal As fCodigosDeCuenta, ByVal formatoDestino As fCodigosDeCuenta) As String
        Dim r As String = ""
        Try
            Dim fieldRetorno As String = ""
            Dim fieldWhere As String = ""

            If formatoOriginal = fCodigosDeCuenta.AcctCode Then fieldWhere = "AcctCode"
            If formatoOriginal = fCodigosDeCuenta.AcctName Then fieldWhere = "AcctName"
            If formatoOriginal = fCodigosDeCuenta.FormatCode Then fieldWhere = "FormatCode"
            If formatoOriginal = fCodigosDeCuenta.SegmentedCode Then
                fieldWhere = "FormatCode"
                valor = valor.Replace("-", "")
            End If

            If formatoDestino = fCodigosDeCuenta.AcctCode Then fieldRetorno = "AcctCode"
            If formatoDestino = fCodigosDeCuenta.AcctName Then fieldRetorno = "AcctName"
            If formatoDestino = fCodigosDeCuenta.FormatCode Then fieldRetorno = "FormatCode"
            If formatoDestino = fCodigosDeCuenta.SegmentedCode Then fieldRetorno = "Cuenta"

            Dim query As String = ""
            query += "SELECT AcctCode, AcctName, FormatCode, segment_0 + "
            query += "case when not segment_1 is null then '-' + segment_1 else '' end + "
            query += "case when not segment_2 is null then '-' + segment_2 else '' end + "
            query += "case when not segment_3 is null then '-' + segment_3 else '' end + "
            query += "case when not segment_4 is null then '-' + segment_4 else '' end + "
            query += "case when not segment_5 is null then '-' + segment_5 else '' end + "
            query += "case when not segment_6 is null then '-' + segment_6 else '' end + "
            query += "case when not segment_7 is null then '-' + segment_7 else '' end + "
            query += "case when not segment_8 is null then '-' + segment_8 else '' end + "
            query += "case when not segment_9 is null then '-' + segment_9 else '' end "
            query += "as Cuenta FROM OACT WHERE " & fieldWhere & " = '" & valor & "'"

            r = getRSvalue(query, fieldRetorno)

        Catch ex As Exception

        End Try
        Return r
    End Function

    ' USER INTERFACE

    ''' <summary>
    ''' Dibuja un item en un formulario
    ''' </summary>
    ''' <param name="oFormItem">Objeto formulario en el que se dibujará el item</param>
    ''' <param name="TipoItem">Tipo de item a dibujar</param>
    ''' <param name="ItemID">UID del item</param>
    ''' <param name="ItemDesc">Descripción del item</param>
    ''' <param name="Left">Posición en píxeles desde la izquierda del formulario</param>
    ''' <param name="Top">Posición en píxeles desde el tope del formulario</param>
    ''' <param name="Width">Ancho del item</param>
    ''' <param name="Height">Altura del item</param>
    ''' <param name="Npanel">Panel del item</param>
    ''' <param name="DisplayDesc">Indica si la descripción debe mostrarse o no</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 

    Public Function creaControl(ByVal oFormItem As SAPbouiCOM.Form, ByVal TipoItem As SAPbouiCOM.BoFormItemTypes, ByVal ItemID As String, ByVal ItemDesc As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal Npanel As Integer = 0, Optional ByVal DisplayDesc As Boolean = False, Optional AcentoEtiquetas As Boolean = False, Optional esInactivo As Boolean = False) As Boolean
        '
        Try
            Dim oItemDLL As SAPbouiCOM.Item
            oItemDLL = oFormItem.Items.Add(ItemID, TipoItem)
            With oItemDLL
                If Not Left = 0 Then .Left = Left
                If Not Width = 0 Then .Width = Width
                If Not Top = 0 Then .Top = Top
                If Not Height = 0 Then .Height = Height


                '
                ' Validación de Panels
                If Npanel <> 0 Then
                    .FromPane = Npanel
                    .ToPane = Npanel
                End If
                '
                'Se crean los rectángulos
                If Not ItemDesc = vbNullString Then
                    If TipoItem = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                        If AcentoEtiquetas Then
                            .TextStyle = 1
                        End If

                        Dim oLabelDLL As SAPbouiCOM.StaticText
                        oLabelDLL = .Specific
                        oLabelDLL.Caption = ItemDesc

                    ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
                        Dim oButtonDLL As SAPbouiCOM.Button
                        oButtonDLL = .Specific
                        oButtonDLL.Caption = ItemDesc
                    ElseIf TipoItem = (SAPbouiCOM.BoFormItemTypes.it_EDIT Or SAPbouiCOM.BoFormItemTypes.it_EXTEDIT) Then
                        If esInactivo Then
                            .Enabled = False
                        End If

                        Dim oEditDLL As SAPbouiCOM.EditText
                        oEditDLL = .Specific
                        oEditDLL.Caption = ItemDesc

                    ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
                        Dim ofol As SAPbouiCOM.Folder
                        ofol = .Specific
                        ofol.Caption = ItemDesc
                        .AffectsFormMode = False
                    Else
                        Return False
                    End If

                ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX Then

                    If DisplayDesc = True Then
                        .DisplayDesc = True

                    End If


                End If

                Return True
            End With
            '
        Catch ex As Exception
            'Pappl.MessageBox(ex.Message, 1, "Aceptar")
            Return False
            ' Pappl.MessageBox("Ya existe un objeto con ese nombre!!! ", 1, "Aceptar")
        End Try
        ''
    End Function

    Public Function creaControl_origin(ByVal oFormItem As SAPbouiCOM.Form, ByVal TipoItem As SAPbouiCOM.BoFormItemTypes, ByVal ItemID As String, ByVal ItemDesc As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal Npanel As Integer = 0, Optional ByVal DisplayDesc As Boolean = False) As Boolean
        '
        Try
            Dim oItemDLL As SAPbouiCOM.Item
            oItemDLL = oFormItem.Items.Add(ItemID, TipoItem)
            With oItemDLL
                If Not Left = 0 Then .Left = Left
                If Not Width = 0 Then .Width = Width
                If Not Top = 0 Then .Top = Top
                If Not Height = 0 Then .Height = Height
                '
                ' Validación de Panels
                If Npanel <> 0 Then
                    .FromPane = Npanel
                    .ToPane = Npanel
                End If
                '
                'Se crean los rectángulos
                If Not ItemDesc = vbNullString Then
                    If TipoItem = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                        Dim oLabelDLL As SAPbouiCOM.StaticText
                        oLabelDLL = .Specific
                        oLabelDLL.Caption = ItemDesc
                    ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
                        Dim oButtonDLL As SAPbouiCOM.Button
                        oButtonDLL = .Specific
                        oButtonDLL.Caption = ItemDesc
                    ElseIf TipoItem = (SAPbouiCOM.BoFormItemTypes.it_EDIT Or SAPbouiCOM.BoFormItemTypes.it_EXTEDIT Or SAPbouiCOM.BoFormItemTypes.it_FOLDER) Then
                        Dim oEditDLL As SAPbouiCOM.EditText
                        oEditDLL = .Specific
                        oEditDLL.Caption = ItemDesc
                    Else
                        Return False
                    End If
                ElseIf TipoItem = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX Then
                    If DisplayDesc = True Then
                        oFormItem.Items.Item(ItemID).DisplayDesc = True
                    End If
                End If
                Return True
            End With
            '
        Catch ex As Exception
            'Pappl.MessageBox(ex.Message, 1, "Aceptar")
            Return False
            ' Pappl.MessageBox("Ya existe un objeto con ese nombre!!! ", 1, "Aceptar")
        End Try
        ''
    End Function
    ''' <summary>
    ''' Dibuja un nuevo item en un formulario
    ''' </summary>
    ''' <param name="formulario">Formulario en el que se dibujará el item</param>
    ''' <param name="tipo">Tipo de item a dibujar</param>
    ''' <param name="itemUID">UID del nuevo item</param>
    ''' <param name="hOffSet">posición horizontal o desface con respecto a otro item</param>
    ''' <param name="vOffSet">posición vertical o desface con respecto a otro item</param>
    ''' <param name="respectoAItem">item en referencia al cual se dibujará</param>
    ''' <param name="ancho">ancho del nuevo item</param>
    ''' <param name="alto">alto del nuevo item</param>
    ''' <param name="copiarTamano">indica si se desea copiar el tamaño del item de referencia</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function creaControl(ByVal formulario As SAPbouiCOM.Form, ByVal tipo As SAPbouiCOM.BoFormItemTypes, ByVal itemUID As String, ByVal hOffSet As Integer, ByVal vOffSet As Integer, Optional ByVal respectoAItem As String = "", Optional ByVal ancho As Integer = -1, Optional ByVal alto As Integer = -1, Optional ByVal copiarTamano As Boolean = False, Optional ByVal visible As Boolean = True) As SAPbouiCOM.Item
        Try
            formulario.Freeze(True)
            formulario.Items.Add(itemUID, tipo)
            formulario.Items.Item(itemUID).Visible = visible
            If copiarTamano Then
                If respectoAItem <> "" Then
                    formulario.Items.Item(itemUID).Width = formulario.Items.Item(respectoAItem).Width
                    formulario.Items.Item(itemUID).Height = formulario.Items.Item(respectoAItem).Height
                End If
            Else
                If ancho >= 0 Then
                    formulario.Items.Item(itemUID).Width = ancho
                End If
                If alto >= 0 Then
                    formulario.Items.Item(itemUID).Height = alto
                End If
            End If
            If respectoAItem = "" Then
                formulario.Items.Item(itemUID).Top = vOffSet
                formulario.Items.Item(itemUID).Left = hOffSet
            Else

                If vOffSet > 0 Then
                    formulario.Items.Item(itemUID).Top = formulario.Items.Item(respectoAItem).Top + formulario.Items.Item(respectoAItem).Height + vOffSet
                ElseIf vOffSet = 0 Then
                    formulario.Items.Item(itemUID).Top = formulario.Items.Item(respectoAItem).Top
                Else
                    formulario.Items.Item(itemUID).Top = formulario.Items.Item(respectoAItem).Top - formulario.Items.Item(itemUID).Height + vOffSet
                End If

                If hOffSet > 0 Then
                    formulario.Items.Item(itemUID).Left = formulario.Items.Item(respectoAItem).Left + formulario.Items.Item(respectoAItem).Width + hOffSet
                ElseIf hOffSet = 0 Then
                    formulario.Items.Item(itemUID).Left = formulario.Items.Item(respectoAItem).Left
                Else
                    formulario.Items.Item(itemUID).Left = formulario.Items.Item(respectoAItem).Left - formulario.Items.Item(itemUID).Width + hOffSet
                End If
                formulario.Items.Item(itemUID).LinkTo = respectoAItem
            End If
            formulario.Refresh()
        Catch ex As Exception
        Finally
            formulario.Freeze(False)
        End Try
        Return formulario.Items.Item(itemUID)
    End Function

    ''' <summary>
    ''' Elimina un dato de un Grid
    ''' </summary>
    ''' <param name="Grid">UID del Grid</param>
    ''' <param name="col">Columna a evaluar</param>
    ''' <param name="cond">Valor a eliminar</param>
    ''' <param name="Formu">Formulario del grid</param>
    ''' <remarks></remarks>
    Public Sub borraGrid(ByVal Grid As String, ByVal col As Integer, ByVal cond As String, ByVal Formu As String)
        Dim kz As Integer = 0
        Try
            Dim fForm As SAPbouiCOM.Form = fPappl.Forms.Item(Formu)
            Dim oGridDLL As SAPbouiCOM.Grid = fForm.Items.Item(Grid).Specific

            For kz = oGridDLL.Rows.Count To 1 Step -1
                If Trim(oGridDLL.DataTable.GetValue(col, kz - 1)) = cond Then
                    oGridDLL.DataTable.Rows.Remove(kz - 1)
                End If
            Next

        Catch ex As Exception
            If mostrarMensajesError Then fPappl.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Cierra todas las instancias de un formulario.
    ''' </summary>
    ''' <param name="FormID">Identificador del formulario</param>
    ''' <param name="TipoForm">Tipo de formulario (Usuario o Estándar)</param>
    ''' <remarks>Puye patrocinado por Gabriel Mendes, muy feo pero se debe validar que siempre tengan la ventana y los campos correctos</remarks>
    Public Sub cierraFormularios(ByVal FormID As String, ByVal TipoForm As T_Form)
        Dim oFormM As SAPbouiCOM.Form
        Dim Count As Integer = fPappl.Forms.Count
        Dim Cantidad As New ArrayList
        '
        For il As Integer = 0 To Count - 1
            If TipoForm = T_Form.Standard Then
                If fPappl.Forms.Item(il).TypeEx = FormID Then
                    Cantidad.Add(fPappl.Forms.Item(il).TypeCount)
                End If
            Else
                If fPappl.Forms.Item(il).UniqueID = FormID Then
                    fPappl.Forms.Item(il).Close()
                End If
            End If
        Next
        If TipoForm = T_Form.Standard Then
            For il2 As Integer = 0 To Cantidad.Count - 1
                oFormM = fPappl.Forms.GetFormByTypeAndCount(FormID, Cantidad.Item(il2))
                If oFormM.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then oFormM.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                oFormM.Close()
            Next
        End If
        ''
    End Sub

    ''' <summary>
    ''' Indica si un formulario se encuentra abierto.
    ''' </summary>
    ''' <param name="FormID">Identificador del formulario</param>
    ''' <param name="Tipo">Tipo de formulario (Estándar o Usuario)</param>
    ''' <param name="TypeCount">Número de instancias del formulario</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isFormLoaded(ByVal FormID As String, ByVal Tipo As T_Form, Optional ByVal TypeCount As Integer = 1) As Boolean
        '
        'Dim oFormA As SAPbouiCOM.Form
        Dim Count As Integer = fPappl.Forms.Count
        'Dim Cantidad As New ArrayList
        '
        For ia As Integer = 0 To Count - 1  'Suponiendo que pueden manejar 20 formularios iguales
            '
            If Tipo = T_Form.Standard Then
                If fPappl.Forms.Item(ia).TypeEx = FormID Then
                    If fPappl.Forms.Item(ia).TypeCount = TypeCount Then
                        Return True
                    End If
                End If
            Else
                If fPappl.Forms.Item(ia).UniqueID = FormID Then
                    Return True
                End If
            End If
        Next
        Return False
        ''
    End Function

    ''' <summary>
    ''' Devuleve un número que indica la cantidad de veces que se encuentra abierto un formulario.
    ''' </summary>
    ''' <param name="FormId">Identificador del Formulario</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFormCount(ByVal FormId As String) As Integer
        Dim retorno As Integer = 1
        Try
            For retorno = 1 To 200
                fPappl.Forms.GetFormByTypeAndCount(FormId, retorno)
            Next
        Catch ex As Exception
        End Try
        Return retorno - 1
    End Function

    ''' <summary>
    ''' Funcion para renombrar los nombres de los campos en los formularios
    ''' </summary>
    ''' <param name="FormID">El FormTypeEx para los Estandar y el FormUID para Forms desarrollados</param>
    ''' <param name="ItemID">El nombre del Item dibujado en el formulario</param>
    ''' <param name="Descripcion">Nueva descripcion a asignar al Item</param>
    ''' <param name="ColumnID">ColumnID en caso de ser una Columna, Por default es -1</param>
    ''' <param name="IsBold">Si aplica Negritas</param>
    ''' <param name="IsItalic">Si aplica Cursiva</param>
    ''' <returns>Booleano, si es True se realizo el cambio, si es False hubo un error</returns>
    ''' <remarks></remarks>
    Public Function DynamicSystemStrings(ByVal FormID As String, ByVal ItemID As String, ByVal Descripcion As String, Optional ByVal ColumnID As String = "-1",
    Optional ByVal IsBold As Boolean = False, Optional ByVal IsItalic As Boolean = False) As Boolean
        Dim DS As SAPbobsCOM.DynamicSystemStrings
        Dim ret As Integer = 0
        Try
            DS = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDynamicSystemStrings)
            If DS.GetByKey(FormID, ItemID, ColumnID) = False Then
                DS.FormID = FormID
                DS.ItemID = ItemID
                DS.ColumnID = ColumnID
                DS.ItemString = Descripcion
                DS.IsBold = IIf(IsBold = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                DS.IsItalics = IIf(IsItalic = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                ret = DS.Add
                If ret <> 0 Then : Return False
                Else : Return True
                End If
            Else
                DS.ItemString = Descripcion
                DS.IsBold = IIf(IsBold = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                DS.IsItalics = IIf(IsItalic = False, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES)
                ret = DS.Update()
                If ret <> 0 Then : Return False
                Else : Return True
                End If
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Limpia cualquier mensaje del statusbar
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearStatusBar()
        Try
            fPappl.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        Catch ex As Exception
        End Try
    End Sub


    ' MENU

    ''' <summary>
    ''' Añade un item tipo POP_UP al menú de usuario de B1.
    ''' </summary>
    ''' <param name="nombreCarpeta">Nombre con el que aparecerá identificada la carpeta</param>
    ''' <param name="idCarpeta">Código identificador interno del item de menú</param>
    ''' <param name="menu1">Carpeta de nivel 1 que contiene este item</param>
    ''' <param name="menu2">Carpeta de nivel 2 que contiene este item (contenida por la carpeta de nivel 1)</param>
    ''' <param name="menu3">Carpeta de nivel 3 que contiene este item (contenida por la carpeta de nivel 2)</param>
    ''' <param name="imagen">Imágen que se mostrará a la izquierda del item (solo nivel 1)</param>
    ''' <remarks></remarks>
    Public Sub addMenuFolder(ByVal nombreCarpeta As String, ByVal idCarpeta As String, Optional ByVal menu1 As String = "", Optional ByVal menu2 As String = "", Optional ByVal menu3 As String = "", Optional ByVal imagen As String = "", Optional ByVal posicion As Integer = -1)
        Try
            If imagen <> "" Then
                imagen = System.Windows.Forms.Application.StartupPath & "\" & carpetaImagenes & "\" & imagen
            End If
            If menu3 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
            ElseIf menu2 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
            ElseIf menu1 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item(menu1).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
            Else
                If fPappl.Menus.Item("43520").SubMenus.Exists(idCarpeta) Then fPappl.Menus.Item("43520").SubMenus.Remove(fPappl.Menus.Item("43520").SubMenus.Item(idCarpeta))
                If posicion = -1 Then posicion = fPappl.Menus.Item("43520").SubMenus.Count
                fPappl.Menus.Item("43520").SubMenus.Add(idCarpeta, nombreCarpeta, SAPbouiCOM.BoMenuType.mt_POPUP, posicion)
                fPappl.Menus.Item("43520").SubMenus.Item(idCarpeta).Image = imagen
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Añade un item tipo STRING al menú de usuario de B1.
    ''' </summary>
    ''' <param name="idForm">Código identificador interno del item</param>
    ''' <param name="nombreForm">Nombre que se presentará al usuario</param>
    ''' <param name="menu1">Carpeta de nivel 1 que contiene este item</param>
    ''' <param name="menu2">Carpeta de nivel 2 que contiene este item (contenida por la carpeta de nivel 1)</param>
    ''' <param name="menu3">Carpeta de nivel 3 que contiene este item (contenida por la carpeta de nivel 2)</param>
    ''' <param name="menu4">Carpeta de nivel 4 que contiene este item (contenida por la carpeta de nivel 3)</param>
    ''' <remarks></remarks>
    Public Sub addMenuItem(ByVal idForm As String, ByVal nombreForm As String, ByVal menu1 As String, Optional ByVal menu2 As String = "", Optional ByVal menu3 As String = "", Optional ByVal menu4 As String = "", Optional ByVal posicion As Integer = -1)
        Try
            If menu4 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(menu4).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            ElseIf menu3 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(menu3).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            ElseIf menu2 <> "" Then
                If fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Item(menu2).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            Else
                If fPappl.Menus.Item(menu1).SubMenus.Exists(idForm) Then fPappl.Menus.Item(menu1).SubMenus.Remove(fPappl.Menus.Item(menu1).SubMenus.Item(idForm))
                If posicion = -1 Then posicion = fPappl.Menus.Item(menu1).SubMenus.Count
                fPappl.Menus.Item(menu1).SubMenus.Add(idForm, nombreForm, SAPbouiCOM.BoMenuType.mt_STRING, posicion)
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Elimina un item de menú
    ''' </summary>
    ''' <param name="menuId">Identificador del Menú</param>
    ''' <remarks></remarks>
    Public Sub removeMenuItem(ByVal menuId As String)
        Try
            fPappl.Menus.Remove(fPappl.Menus.Item(menuId))
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Abre el formulario por defecto de una UDT
    ''' </summary>
    ''' <param name="tabla">Nombre (Código) de la Tabla de usuario a abrir</param>
    ''' <remarks></remarks>
    Public Sub abrirTablaUsuario(ByVal tabla As String)
        Try
            For i As Integer = 51201 To 51999
                If fPappl.Menus.Item(i.ToString).String.ToString.StartsWith(tabla) Then
                    fPappl.Menus.Item(i.ToString).Activate()
                    Exit Sub
                End If
            Next
            fPappl.MessageBox("No se pudo encontrar la tabla de usuario " & tabla)
        Catch ex As Exception

        End Try
    End Sub


    ' CARGAS AUTOMÁTICAS

    ''' <summary>
    ''' Carga los datos de un query en un combo.
    ''' </summary>
    ''' <param name="strFormUID">UID del formulario en el que se encuentra el combo</param>
    ''' <param name="strItemUID">UID del Combo o de la Matriz que lo contiene</param>
    ''' <param name="strQRY">String con la consulta de selección de los datos para el combo. Tomará la primera columna como VALUE y la segunda (opcional) como DESCRIPTION.</param>
    ''' <param name="incluirValorCero">Si se coloca en true, incluye un valor en blanco en el combo</param>
    ''' <remarks></remarks>
    Public Sub cargaCombo(ByVal strFormUID As String, ByVal strItemUID As String, ByVal strQRY As String, Optional ByVal incluirValorCero As Boolean = False)
        Try
            Dim Fila As Integer = 0
            Dim oComboX As SAPbouiCOM.ComboBox
            Dim strCMBdesc As String = vbNullString
            Dim xForm As SAPbouiCOM.Form = fPappl.Forms.Item(strFormUID)

            oComboX = xForm.Items.Item(strItemUID).Specific

            cargaCombo(oComboX, strQRY, incluirValorCero)

        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Carga los datos de un query en un combo que se encuentra en una Matriz.
    ''' </summary>
    ''' <param name="strFormUID">UID del formulario en el que se encuentra el combo</param>
    ''' <param name="strMatrixUID">UID de la Matriz que lo contiene</param>
    ''' <param name="colUID">UID de la columna tipo combo</param>
    ''' <param name="Fila">Fila de la matriz a actualizar</param>
    ''' <param name="strQRY">String con la consulta de selección de los datos para el combo. Tomará la primera columna como VALUE y la segunda (opcional) como DESCRIPTION.</param>
    ''' <param name="incluirValorCero">Si se coloca en true, incluye un valor en blanco</param>
    ''' <remarks></remarks>
    Public Sub cargaCombo(ByVal strFormUID As String, ByVal strMatrixUID As String, ByVal colUID As String, ByVal Fila As Integer, ByVal strQRY As String, Optional ByVal incluirValorCero As Boolean = False)
        Try
            Dim oComboX As SAPbouiCOM.ComboBox
            Dim strCMBdesc As String = vbNullString
            Dim xForm As SAPbouiCOM.Form = fPappl.Forms.Item(strFormUID)
            Dim xMatrix As SAPbouiCOM.Matrix
            Dim xColumn As SAPbouiCOM.Column
            xMatrix = xForm.Items.Item(strMatrixUID).Specific
            xColumn = xMatrix.Columns.Item(colUID)
            oComboX = xColumn.Cells.Item(Fila).Specific

            cargaCombo(oComboX, strQRY, incluirValorCero)

        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Carga los datos de un query en un combo.
    ''' </summary>
    ''' <param name="oComboX">ComboBox a llenar</param>
    ''' <param name="strQRY">String con la consulta de selección de los datos para el combo. Tomará la primera columna como VALUE y la segunda (opcional) como DESCRIPTION.</param>
    ''' <param name="incluirValorCero">Si se coloca en true, incluye un valor en blanco en el combo</param>
    ''' <remarks></remarks>
    Public Sub cargaCombo(ByVal oComboX As SAPbouiCOM.ComboBox, ByVal strQRY As String, Optional ByVal incluirValorCero As Boolean = False)
        Try
            Dim strCMBdesc As String = vbNullString

            If oComboX.ValidValues.Count > 0 Then
                For Fila As Integer = 0 To oComboX.ValidValues.Count - 1
                    oComboX.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Next Fila
            End If
            Dim xAuxLocal As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            xAuxLocal.DoQuery(strQRY)
            If incluirValorCero = True Then oComboX.ValidValues.Add("", "")
            If xAuxLocal.RecordCount > 0 Then
                xAuxLocal.MoveFirst()
                strCMBdesc = xAuxLocal.Fields.Item(0).Value
                While Not xAuxLocal.EoF
                    If xAuxLocal.Fields.Count = 1 Then
                        oComboX.ValidValues.Add(xAuxLocal.Fields.Item(0).Value, xAuxLocal.Fields.Item(0).Value)
                    Else
                        oComboX.ValidValues.Add(xAuxLocal.Fields.Item(0).Value, xAuxLocal.Fields.Item(1).Value)
                    End If
                    xAuxLocal.MoveNext()
                End While
            Else
                strCMBdesc = "Definir nuevo"
            End If
            Release(xAuxLocal)
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Añade una fila en una matriz, validando que el campo de la primera columna de la última fila esté lleno.
    ''' </summary>
    ''' <param name="formulario">UID del formulario en el que se encuentra la matriz</param>
    ''' <param name="matriz">UID de la matriz</param>
    ''' <param name="celdaEsCombo">Indica si la celda a evaluar es un combo o un editText</param>
    ''' <remarks></remarks>
    Public Sub addRow(ByVal formulario As String, ByVal matriz As String, Optional ByVal celdaEsCombo As Boolean = False)
        Try
            Dim fMatrix As SAPbouiCOM.Matrix = fPappl.Forms.Item(formulario).Items.Item(matriz).Specific
            If celdaEsCombo = True Then
                Dim fCombo As SAPbouiCOM.ComboBox = fMatrix.Columns.Item(1).Cells.Item(fMatrix.RowCount).Specific
                If fMatrix.RowCount = 0 Or fCombo.Selected.Value > 0 Then
                    fMatrix.AddRow()
                    fPappl.Forms.Item(formulario).Update()
                    fPappl.Forms.Item(formulario).Refresh()
                    fCombo = fMatrix.Columns.Item(1).Cells.Item(fMatrix.RowCount).Specific
                    fCombo.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
            Else
                Dim fEdit As SAPbouiCOM.EditText = fMatrix.Columns.Item(1).Cells.Item(fMatrix.RowCount).Specific
                If fMatrix.RowCount = 0 Or fEdit.Value > 0 Then
                    fMatrix.AddRow()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Devuelve el siguiente valor numérico para un campo. Si falla, devuelve cero.
    ''' </summary>
    ''' <param name="CampoMax">Campo del correlativo</param>
    ''' <param name="Tabla">Tabla del campo correlativo</param>
    ''' <param name="condicion">Cláusula WHERE... para la consulta</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCorrelativo(ByVal CampoMax As String, ByVal Tabla As String, Optional ByVal condicion As String = "", Optional ByVal primerCorrelativo As Integer = 1) As String
        Dim oMax As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Srt As String = primerCorrelativo.ToString
        Try
            'getCorrelativo("Code", "[@SS_SETUP]", , 1) 
            If TipoServer = "9" Then
                Srt = "SELECT IFNULL(MAX(CAST(""" & CampoMax & """ AS Integer)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM """ & BD & """.""" & Tabla & """"
                If condicion <> "" Then
                    Srt = "SELECT IFNULL(MAX(CAST(""" & CampoMax & """ AS Integer)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM (SELECT * FROM """ & BD & """.""OWHS"" WHERE """ & condicion & """) AS X WHERE """ & condicion & """"
                End If
            Else
                Srt = "SELECT ISNULL(MAX(CAST(" & CampoMax & " AS numeric)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM  """ & Tabla & """"
                If condicion <> "" Then
                    Srt = "SELECT ISNULL(MAX(CAST(" & CampoMax & " AS numeric)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM (SELECT * FROM OWHS WHERE " & condicion & ") AS X WHERE " & condicion
                End If
            End If

            oMax.DoQuery(Srt)
            Srt = IIf(oMax.EoF = True, primerCorrelativo.ToString, oMax.Fields.Item("Numero").Value)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("GetCorrelativo_Catch, Error: " + ex.Message.ToString(), "FuncionesB1")
            Utilitario.Util_Log.Escribir_Log("GetCorrelativo_Catch, Query: " + Srt, "FuncionesB1")
            Srt = "0"
        Finally
            Release(oMax)
        End Try
        Return Srt
    End Function

    ''' <summary>
    ''' Devuelve el siguiente valor numérico para un campo. Si falla, devuelve cero. con company
    ''' </summary>
    ''' <param name="CampoMax">Campo del correlativo</param>
    ''' <param name="Tabla">Tabla del campo correlativo</param>
    ''' <param name="condicion">Cláusula WHERE... para la consulta</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCorrelativo(ByVal CampoMax As String, ByVal Tabla As String, ByVal mCompany As SAPbobsCOM.Company, Optional ByVal condicion As String = "", Optional ByVal primerCorrelativo As Integer = 1) As String
        Dim oMax As SAPbobsCOM.Recordset = mCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Srt As String = primerCorrelativo.ToString
        Try
            If TipoServer = "9" Then
                Srt = "SELECT IFNULL(MAX(CAST(""" & CampoMax & """ AS Integer)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM """ & BD & """.""" & Tabla & """"
                If condicion <> "" Then
                    Srt = "SELECT IFNULL(MAX(CAST(""" & CampoMax & """ AS Integer)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM (SELECT * FROM """ & BD & """.""OWHS"" WHERE """ & condicion & """) AS X WHERE """ & condicion & """"
                End If
            Else
                Srt = "SELECT ISNULL(MAX(CAST(" & CampoMax & " AS numeric)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM " & Tabla
                If condicion <> "" Then
                    Srt = "SELECT ISNULL(MAX(CAST(" & CampoMax & " AS numeric)), " & primerCorrelativo - 1 & ") + 1 AS Numero FROM (SELECT * FROM OWHS WHERE " & condicion & ") AS X WHERE " & condicion
                End If
            End If
            oMax.DoQuery(Srt)


            Srt = IIf(oMax.EoF = True, primerCorrelativo.ToString, oMax.Fields.Item("Numero").Value)

        Catch ex As Exception
            Srt = "0"
        Finally
            Release(oMax)
        End Try
        Return Srt
    End Function


    ' MANEJO DE VALORES NULOS Y CONVERSIÓN

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
                'If unString = "0" Then
                '    unString = ""
                'End If
                valorSiNulo = unString
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("nzString Catch:" + ex.Message().ToString(), "FuncionesB1")
        End Try
        Return valorSiNulo
    End Function

    ''' <summary>
    ''' Devuelve un valor double, validando si es nulo o infinito, y si lo es devuelve un valor establecido.
    ''' </summary>
    ''' <param name="a">Valor a evaluar</param>
    ''' <param name="valorSiEsNulo">Valor que será devuelto si es nulo o infinito</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzDouble(ByVal a As Double, Optional ByVal valorSiEsNulo As Double = 0) As Double
        Try
            If Not IsDBNull(a) And Not Double.IsInfinity(a) Then
                Return a
            Else
                Return valorSiEsNulo
            End If
        Catch ex As Exception
            Return valorSiEsNulo
        End Try
    End Function

    ''' <summary>
    ''' Devuelve una fecha en formato de Business One
    ''' </summary>
    ''' <param name="fecha">Fecha a convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzDate(ByVal fecha As String, Optional ByVal valorSiNulo As String = "") As String
        Dim retorno As String = valorSiNulo
        Try
            If Not IsDBNull(fecha) Then

                Dim f As Date = CDate(fecha)
                retorno += f.Year.ToString
                If f.Month < 10 Then retorno += "0"
                retorno += f.Month.ToString
                If f.Day < 10 Then retorno += "0"
                retorno += f.Day.ToString

            End If
        Catch ex As Exception
            retorno = valorSiNulo
        End Try
        Return retorno
    End Function

    ''' <summary>
    ''' Devuelve fecha y hora en formato Date
    ''' </summary>
    ''' <param name="Fecha">Fecha en formato String</param>
    ''' <param name="Hora">Hora y minutos como un número</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFechaHoraB1(ByVal Fecha As String, ByVal Hora As Long) As Date
        Dim FechaA As Date
        Try
            Dim Minutos As Long
            If Hora <> 0 Then
                'Explicacion
                'Hora = CInt(Hora.ToString.Remove(Hora.ToString.Length - 2, 2))
                'Minutos = CInt(Hora.ToString.Substring(Hora.ToString.Length - 2, 2))
                'Minutos = (Hora * 60) + Minutos

                Minutos = (CInt(Hora.ToString.Remove(Hora.ToString.Length - 2, 2)) * 60) + CInt(Hora.ToString.Substring(Hora.ToString.Length - 2, 2))

            Else
                Minutos = CInt(Hora.ToString.Substring(Hora.ToString.Length - 2, 2))
            End If

            FechaA = CDate(Fecha & " 00:00").AddMinutes(Minutos)

        Catch ex As Exception

        End Try

        Return FechaA

    End Function

    ''' <summary>
    ''' Devuelve Now en formato YYYYMMDD
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFechaActual() As String
        Dim mynum As String = ""
        Try
            mynum = Now.Year.ToString
            If Now.Month < 10 Then mynum += "0"
            mynum += Now.Month.ToString
            If Now.Day < 10 Then mynum += "0"
            mynum += Now.Day.ToString
            Return mynum
        Catch ex As Exception
            Return mynum
        End Try
    End Function

    ''' <summary>
    ''' Redondea un número bajo los criterios especificados.
    ''' </summary>
    ''' <param name="valor">Valor a redondear</param>
    ''' <param name="posicionDecimal">Cantidad de decimales para redondeo. Ejemplo: 12,34... (-1) = 10 ... (0) = 12 ... (1) = 12,3</param>
    ''' <param name="siempreHaciaArriba">Indica si se desea que siempre redondee al siguiente número si existen decimales. Ejemplo: 12,34 ... (true) = 13 ... (false) = 12</param>
    ''' <param name="aCeroOCinco">Indica si se redondea en base 5 en lugar de base 10. Ejemplo: 12,34 ... (true) = 12,5 ... (false) = 12,0</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function nzRedondear(ByVal valor As Double, Optional ByVal posicionDecimal As Integer = 0, Optional ByVal siempreHaciaArriba As Boolean = False, Optional ByVal aCeroOCinco As Boolean = False) As Double
        Dim retorno As Double = valor
        Dim sensibilidad As Double = 0.01
        Dim Rnumero As Double = 0
        Dim valorASumar As Double = 0
        Try
            If aCeroOCinco = True Then posicionDecimal += 1

            valor = valor / Math.Pow(10, posicionDecimal)
            If siempreHaciaArriba = True And aCeroOCinco = False Then valor += 0.5 - sensibilidad

            If aCeroOCinco = True And Not valor = Math.Round(valor) Then
                Dim Rvst As String = valor.ToString
                If Rvst.Contains(",") Then
                    Rvst = Rvst.Substring(Rvst.IndexOf(",") + 1, Rvst.Length - Rvst.IndexOf(",") - 1)
                    Rvst = CDbl(Rvst) / Math.Pow(10, Rvst.Length - 1)
                ElseIf Rvst.Contains(".") Then
                    Rvst = Rvst.Substring(Rvst.IndexOf(".") + 1, Rvst.Length - Rvst.IndexOf(".") - 1)
                    Rvst = CDbl(Rvst) / Math.Pow(10, Rvst.Length - 1)
                Else
                    Rvst = Rvst.Substring(Rvst.Length - 1, 1)
                End If
                Rnumero = CDbl(Rvst)

                If Rnumero = 0 Then
                    valorASumar = -10
                ElseIf Rnumero > 0 And Rnumero < 5 Then
                    valorASumar = 5
                ElseIf Rnumero = 5 Then
                    valorASumar = 5
                ElseIf Rnumero > 5 Then
                    valorASumar = 0
                End If

            End If

            retorno = Math.Round(valor, 0)
            If aCeroOCinco = True Then
                retorno = (retorno * 10) + valorASumar
                posicionDecimal -= 1
            End If
            retorno = retorno * Math.Pow(10, posicionDecimal)

        Catch ex As Exception
        End Try
        Return retorno

    End Function



    ' CONEXIÓN 

    ' A B1 desde afuera

    ''' <summary>
    ''' Connexion a un Sap busines one desde afuera 
    ''' </summary>
    ''' <param name="Server">Nombre del servidor B1</param>
    ''' <param name="CompanyDB">Nombre de la base de datos de la compañia</param>
    ''' <param name="UserName">Nombre del Usuario del B1 Ejem "Manager"</param>
    ''' <param name="Password">Clave del usuario del B1</param>
    ''' <param name="language">En que lenguaje esta la base de datos</param>
    ''' <param name="DbUserName">Nombre del Usuario del servidor de SQL Server Ejem "SA"</param>
    ''' <param name="DbPassword">Clave del Usuario del servidor de SQL Server</param>
    ''' <param name="DbServerType">Tipo del servidor de la base de datos Ejem "MServer2008"</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ConnectToCompany(ByVal Server As String, ByVal CompanyDB As String, ByVal UserName As String, ByVal Password As String, ByVal language As SAPbobsCOM.BoSuppLangs, ByVal DbUserName As String, ByVal DbPassword As String, ByVal DbServerType As SAPbobsCOM.BoDataServerTypes, ByVal Lcsr As String)
        Try

            Dim sErrMsg As String = vbNullString
            Dim lErrCode As Long


            mCompany.Server = Server
            mCompany.CompanyDB = CompanyDB
            mCompany.UserName = UserName
            mCompany.Password = Password
            mCompany.language = language
            mCompany.DbUserName = DbUserName
            mCompany.DbPassword = DbPassword
            mCompany.DbServerType = DbServerType
            mCompany.LicenseServer = Lcsr
            mCompany.Connect()

            ' Check for errors during connect
            mCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                Return sErrMsg
            Else
                Return mCompany
            End If
        Catch ex As Exception
            Return False
        End Try


    End Function

    Public Sub ExitToCompany(ByVal cCompany As SAPbobsCOM.Company)

        cCompany.Disconnect()
        cCompany = Nothing
        GC.Collect()

    End Sub

    'A EXCEL

    ''' <summary>
    ''' Retorna una conexión ADO a un archivo Excel.
    ''' </summary>
    ''' <param name="NombreArch">Nombre del archivo incluyendo extensión (Ej: "Prueba.xls")</param>
    ''' <param name="Ruta">Ubicación del archivo (Ej: "C:\Excel\"). Si no se especifica toma la carpeta de archivos excel de las parametrizaciones generales.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getExcelConnection(ByVal NombreArch As String, Optional ByVal Ruta As String = "") As ADODB.Connection
        Try
            Dim ConexionExcel As New ADODB.Connection
            ConexionExcel.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            ConexionExcel.CommandTimeout = 1200

            Dim strSQL As String = ""
            If Ruta = "" Then Ruta = fCompany.ExcelDocsPath
            If Ruta = vbNullString Then
                Ruta = "C:\"
            End If

            Dim v2007 As Boolean = NombreArch.ToUpper.EndsWith(".XLSX")

            If v2007 Then
                strSQL = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Ruta & NombreArch & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""

            Else
                strSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Ruta & NombreArch & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""

            End If

            ConexionExcel.ConnectionString = strSQL
            ConexionExcel.Open()

            Return ConexionExcel

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Devuelve un Recordset sobre un archivo Excel.
    ''' </summary>
    ''' <param name="Query">Consulta a ejecutar sobre el archivo Excel</param>
    ''' <param name="NombreArch">Nombre del archivo incluyendo extensión (Ej: "Prueba.xls")</param>
    ''' <param name="Ruta">Ubicación del archivo (Ej: "C:\Excel\"). Si no se especifica toma la carpeta de archivos excel de las parametrizaciones generales.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getExcelRecordset(ByVal Query As String, ByVal NombreArch As String, Optional ByVal Ruta As String = "") As ADODB.Recordset
        Try
            Dim ConexionExcel As ADODB.Connection = getExcelConnection(NombreArch, Ruta)
            Dim rsLocal As New ADODB.Recordset
            rsLocal.Open(Query, ConexionExcel, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
            Return rsLocal
        Catch ex As Exception
            Return Nothing
        End Try
    End Function



    ' OTRAS

    ''' <summary>
    ''' Si el idioma de Business One es cualquier variante de Español, devuelve true. En caso contrario, false.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isInSpanish() As Boolean
        Dim miBool As Boolean = False
        Try
            If fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish Or fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish_Ar Or fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish_La Or fPappl.Language = SAPbouiCOM.BoLanguages.ln_Spanish_Pa Then miBool = True
        Catch ex As Exception
        End Try
        Return miBool
    End Function

    ''' <summary>
    ''' Agrega una línea al archivo txt del log.
    ''' </summary>
    ''' <param name="Contenido">Contenido de la línea de texto</param>
    ''' <param name="FileName">Nombre del archivo an el que se registra el log (sin extensión .txt)</param>
    ''' <param name="oRuta">Ruta en la que se guardará el archivo (Ejemplo: C:\Logs)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function addLogTxt(ByVal Contenido As String, ByVal FileName As String, Optional ByVal oRuta As String = "Default") As Boolean
        Try
            If oRuta = "Default" Then
                oRuta = System.IO.Directory.GetCurrentDirectory & "\Logs"
            End If
            If System.IO.Directory.Exists(oRuta) = False Then
                System.IO.Directory.CreateDirectory(oRuta)
            End If
            System.IO.File.AppendAllText(oRuta & "\" & FileName & ".txt", Date.Now & ", " & Contenido)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Muestra un MessageBox con las opciones Sí / No (en el idioma correcto)
    ''' </summary>
    ''' <param name="mensaje">Mensaje a mostrar</param>
    ''' <param name="defaultButton">Botón por defecto</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function msgboxYN(ByVal mensaje As String, Optional ByVal defaultButton As Integer = 1) As Integer
        If isInSpanish() Then
            Return fPappl.MessageBox(mensaje, defaultButton, "Sí", "No")
        Else
            Return fPappl.MessageBox(mensaje, defaultButton, "Yes", "No")
        End If
    End Function

    ' NUEVO
    Public Function FormularioActivo(Optional ByRef idForm As String = "", Optional ByRef count As Integer = 0) As String

        For index As Integer = 0 To fPappl.Forms.Count - 1
            If fPappl.Forms.Item(index).Selected Then
                idForm = fPappl.Forms.Item(index).UniqueID
                count = fPappl.Forms.Item(index).TypeCount
                Return fPappl.Forms.Item(index).TypeEx
            End If
        Next
        Return String.Empty

    End Function
    Public Sub InsertaConfiguracion()
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim oUserTable As SAPbobsCOM.UserTable
        Try
            '// set the object with the requested table
            oUserTable = fCompany.UserTables.Item("GS_Configuracion")
            If Not oUserTable.GetByKey(99) Then

                '// set the two default fields 
                oUserTable.Code = 99
                oUserTable.Name = "(GS)Configuracion Servicios"
                '// set the user fields  
                'If oUserTable.UserFields.Fields.Item("U_ws_FC").Value.ToString() <> "" Then
                '    oUserTable.UserFields.Fields.Item("U_ws_FC").Value = "http://www.gurusoft-lab.com/ws_edoc/WSEDOC_FACTURAS.svc"
                'End If
                oUserTable.UserFields.Fields.Item("U_ws_FC").Value = "http://gurusoft-lab.com/ws_edoc/WSEDOC_FACTURAS.svc"
                oUserTable.UserFields.Fields.Item("U_ws_ND").Value = "http://gurusoft-lab.com/ws_edoc/WSEDOC_NOTAS_DEBITO.svc"
                oUserTable.UserFields.Fields.Item("U_ws_NC").Value = "http://gurusoft-lab.com/ws_edoc/WSEDOC_NOTAS_CREDITO.svc"
                oUserTable.UserFields.Fields.Item("U_ws_GR").Value = "http://gurusoft-lab.com/ws_edoc/WSEDOC_GUIAS_REMISION.svc"
                oUserTable.UserFields.Fields.Item("U_ws_RE").Value = "http://gurusoft-lab.com/ws_edoc/WSEDOC_RETENCIONES.svc"
                oUserTable.UserFields.Fields.Item("U_ws_CE").Value = "http://www.gurusoft-lab.com/ws_edoc/WSEDOC_CONSULTA.svc"
                oUserTable.UserFields.Fields.Item("U_ws_RM").Value = "http://192.168.0.103/EDOCWS_ENVIARMAIL/"
                oUserTable.UserFields.Fields.Item("U_ws_CL").Value = "gsedoc"

                oUserTable.UserFields.Fields.Item("U_ws_C").Value = "http://www.gurusoft-lab.com/WSRAD_KEY_RECEPCION/WSRAD_KEY_CONSULTA.svc"
                oUserTable.UserFields.Fields.Item("U_ws_CES").Value = "http://www.gurusoft-lab.com/WSRAD_KEY_RECEPCION/WSRAD_KEY_CAMBIARESTADO.svc"
                oUserTable.UserFields.Fields.Item("U_ws_CC").Value = "XcjWY6C4qz75Yw/mShynPVx7mC2p97Qj"
                oUserTable.UserFields.Fields.Item("U_ws_OC").Value = "OC"
                oUserTable.UserFields.Fields.Item("U_ws_RA").Value = "http://www.gurusoft-lab.com/WSRAD_KEY_RECEPCION/WSRAD_KEY_ARCHIVO.svc"

                'oUserTable.UserFields.Fields.Item("U_ws_FC").Value = "http://www.edocnube.com/EDOCWS_NUBE/WSEDOCNUBE_FACTURAS.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_ND").Value = "http://www.edocnube.com/EDOCWS_NUBE/WSEDOCNUBE_NOTAS_DEBITO.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_NC").Value = "http://www.edocnube.com/EDOCWS_NUBE/WSEDOCNUBE_NOTAS_CREDITO.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_GR").Value = "http://www.edocnube.com/EDOCWS_NUBE/WSEDOCNUBE_GUIAS_REMISION.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_RE").Value = "http://www.edocnube.com/EDOCWS_NUBE/WSEDOCNUBE_RETENCIONES.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_CE").Value = "http://www.edocnube.com/EDOCWS_NUBE2/WSEDOCNUBE_CONSULTA.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_RM").Value = "http://www.edocnube.com/EDOCWS_ENVIARMAIL/WSEDOC_ENVIARMAIL.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_CL").Value = "2MULhn7O/LNM0Xs5w2eb6ZaP1ao3otIQ8+Rnz9KRvoN3RZN6bfLS+8Vv2fLNXBNk3m+9Y/zWJRDkaBWcKomI7qFbVqqe6VVYseRY0tIaXME="

                'oUserTable.UserFields.Fields.Item("U_ws_C").Value = "http://www.edocnube.com/EDOCWSRAD_KEYRECEPCION/WSRAD_KEY_CONSULTA.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_CES").Value = "http://www.edocnube.com/EDOCWSRAD_KEYRECEPCION/WSRAD_KEY_CAMBIARESTADO.svc"
                'oUserTable.UserFields.Fields.Item("U_ws_CC").Value = "DOLEWSEDOCRAD14"
                'oUserTable.UserFields.Fields.Item("U_ws_OC").Value = "OC"
                'oUserTable.UserFields.Fields.Item("U_ws_RA").Value = "http://www.edocnube.com/EDOCWSRAD_KEYRECEPCION/WSRAD_KEY_ARCHIVO.svc"

                oUserTable.Add()
                '// Check for errors
                fCompany.GetLastError(lErrCode, sErrMsg)
                If lErrCode <> 0 Then
                    fPappl.StatusBar.SetText(NombreAddon + " - Error al ingresar configuracion previa: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    fPappl.StatusBar.SetText(NombreAddon + " - Configuracion previa Ingresada", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            oUserTable = Nothing
            System.GC.Collect()
        End Try
    End Sub


    Public Sub InsertaConfiguracionUDO()

#Disable Warning BC42024 ' Variable local sin usar: 'oGeneralService'.
        Dim oGeneralService As SAPbobsCOM.GeneralService
#Enable Warning BC42024 ' Variable local sin usar: 'oGeneralService'.
#Disable Warning BC42024 ' Variable local sin usar: 'oGeneralData'.
        Dim oGeneralData As SAPbobsCOM.GeneralData
#Enable Warning BC42024 ' Variable local sin usar: 'oGeneralData'.
#Disable Warning BC42024 ' Variable local sin usar: 'oChild'.
        Dim oChild As SAPbobsCOM.GeneralData
#Enable Warning BC42024 ' Variable local sin usar: 'oChild'.
#Disable Warning BC42024 ' Variable local sin usar: 'oChildren'.
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
#Enable Warning BC42024 ' Variable local sin usar: 'oChildren'.
#Disable Warning BC42024 ' Variable local sin usar: 'oGeneralParams'.
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
#Enable Warning BC42024 ' Variable local sin usar: 'oGeneralParams'.
#Disable Warning BC42024 ' Variable local sin usar: 'oCompanyService'.
        Dim oCompanyService As SAPbobsCOM.CompanyService
#Enable Warning BC42024 ' Variable local sin usar: 'oCompanyService'.
        Try
            'Dim query As String
            'Dim CodeExist As String = "0"
            'If fCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            '    query = "Select ""DocEntry"" From """ & fCompany.CompanyDB & """.""@GS_LOG"" Where ""U_Clave"" = '" + DocEntry_Clave + "' AND ""U_ObjType"" = '" + ObjType + "'"
            'Else
            '    query = "Select DocEntry From [@GS_LOG] Where U_Clave = '" + DocEntry_Clave + "' AND U_ObjType = '" + ObjType + "'"
            'End If
            'CodeExist = getRSvalue(query, "DocEntry")

        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function FechaSql(ByVal fecha As DateTime) As String

        Dim anio As String = fecha.Year
        Dim mes As String = fecha.Month
        Dim dia As String = fecha.Day

        If anio.Length = 2 Then
            anio = "20" & anio
        End If

        Return "{d'" & anio & "-" & mes.PadLeft(2, "0") & "-" & dia.PadLeft(2, "0") & "'}"

    End Function

    Public Function BobStringToDate(ByVal str As String, ByRef fecha As Date) As Boolean

        Try
            Dim objBob As SAPbobsCOM.SBObob
            Dim objRSet As SAPbobsCOM.Recordset

            objBob = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objRSet = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            fecha = Convert.ToDateTime(objBob.Format_StringToDate(str).Fields.Item(0).Value)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objBob)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRSet)

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function BobDateToString(ByVal fecha As Date, ByRef str As String) As Boolean

        Try
            Dim objBob As SAPbobsCOM.SBObob
            Dim objRSet As SAPbobsCOM.Recordset

            objBob = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objRSet = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            str = objBob.Format_DateToString(fecha).Fields.Item(0).Value

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objBob)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRSet)

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function CargaChooseFromList(ByRef oEvent As SAPbouiCOM.ItemEvent, ByVal oForm As SAPbouiCOM.Form) As SAPbouiCOM.DataTable

        Dim oChooseFromListEvent As SAPbouiCOM.ChooseFromListEvent
        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
        Dim ItemUID As String

        oChooseFromListEvent = oEvent
        ItemUID = oChooseFromListEvent.ChooseFromListUID
        oChooseFromList = oForm.ChooseFromLists.Item(ItemUID)

        Return oChooseFromListEvent.SelectedObjects

    End Function

    Public Shared Function _FechaSql(ByVal fecha As DateTime) As String

        Dim anio As String = fecha.Year
        Dim mes As String = fecha.Month
        Dim dia As String = fecha.Day

        If anio.Length = 2 Then
            anio = "20" & anio
        End If

        'Return "'" & anio & "-" & mes.PadLeft(2, "0") & "-" & dia.PadLeft(2, "0") & "'"
        Return "'" & mes & "/" & dia & "/" & anio & "'"
    End Function
    Public Function getCorrelativoCount(ByVal Tabla As String) As String
        Dim oMax As SAPbobsCOM.Recordset = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Srt As String = ""
        Try
            'getCorrelativo("Code", "[@SS_SETUP]", , 1) 
            If TipoServer = "9" Then

                Srt = "SELECT  COUNT(*) as  Numero FROM """ & BD & """.""" & Tabla & """"

            Else
                Srt = "SELECT  COUNT(*) as  Numero FROM  """ & Tabla & """"

            End If

            oMax.DoQuery(Srt)

            If oMax.RecordCount > 0 Then
                Srt = oMax.Fields.Item("Numero").Value.ToString
            Else
                Srt = "0"

            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("GetCorrelativo_Catch, Error: " + ex.Message.ToString(), "FuncionesB1")
            Utilitario.Util_Log.Escribir_Log("GetCorrelativo_Catch, Query: " + Srt, "FuncionesB1")
            Srt = "0"
        Finally
            Release(oMax)
        End Try
        Return Srt
    End Function

    Public Function ObtenerUIDMenu(ByVal nombreBuscado As String, ByVal MenuPadre As String) As String

        Try

            'el codigo 30338 es Informes electrónicos
            For i As Integer = 0 To fPappl.Menus.Item(MenuPadre).SubMenus.Count - 1


                If fPappl.Menus.Item(MenuPadre).SubMenus.Item(i).String.ToLower.Contains(nombreBuscado.ToLower) Then Return fPappl.Menus.Item(MenuPadre).SubMenus.Item(i).UID


            Next

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error al ejecutar funcion ObtenerUIDMenu " & ex.Message, "EventosLE")
        End Try

        Return ""

    End Function


    Public Function ParseFechaDesdeExcel(fechaCruda As Object, banco As String) As Date
        If TypeOf fechaCruda Is Date Then
            Return CType(fechaCruda, Date) ' Directo
        End If

        Dim textoFecha As String = Convert.ToString(fechaCruda).Trim()

        ' Eliminar hora si viene
        If textoFecha.Contains(" ") Then
            textoFecha = textoFecha.Split(" "c)(0)
        End If

        ' PRODUBANCO: viene como MMddyyyy
        If banco.ToUpperInvariant().Contains("PRODUBANCO") Then
            If textoFecha.Length >= 8 Then
                Dim mm = textoFecha.Substring(0, 2)
                Dim dd = textoFecha.Substring(3, 2)
                Dim yyyy = textoFecha.Substring(6, 4)
                Return New Date(CInt(yyyy), CInt(mm), CInt(dd))
            End If

            ' PICHINCHA: viene como ddMMyyyy
        ElseIf banco.ToUpperInvariant().Contains("PICHINCHA") Then
            If textoFecha.Length >= 8 Then
                Dim dd = textoFecha.Substring(0, 2)
                Dim mm = textoFecha.Substring(2, 2)
                Dim yyyy = textoFecha.Substring(4, 4)
                Return New Date(CInt(yyyy), CInt(mm), CInt(dd))
            End If
        End If

        ' Último intento genérico
        Return Date.ParseExact(textoFecha, "dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture)
    End Function


    Public Function SeleccionarArchivoExcel() As String
        Dim rutaArchivo As String = Nothing
        Dim hilo As New Thread(Sub()
                                   Dim dialog As New OpenFileDialog()
                                   dialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx"
                                   dialog.Title = "Seleccionar archivo Excel"
                                   If dialog.ShowDialog() = DialogResult.OK Then
                                       rutaArchivo = dialog.FileName
                                   End If
                               End Sub)
        hilo.SetApartmentState(Threading.ApartmentState.STA)
        hilo.Start()
        hilo.Join()
        Return rutaArchivo
    End Function


End Class

Imports System.Data.SqlClient
Imports System.Net

Public Class DatabaseQueryManager
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    Private _tipoManejo As String
    Private oFuncionesAddon As Functions.FuncionesAddon
    Public CONEXION As Odbc.OdbcConnection

    Public Sub New(company As SAPbobsCOM.Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
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

    Public Function EjecutarSP(SP As String, docentry As Integer) As DataSet
        Dim ds As New DataSet

        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            Utilitario.Util_Log.Escribir_Log("Query Consulta : " & SP, "ManejoDeDocumentos")
            ds = ObtenerColeccion(SP, False)
        Else
            Try
                Utilitario.Util_Log.Escribir_Log("Query Consulta : " & SP, "ManejoDeDocumentos")

                Using Cn As SqlConnection = GetSqlConnectionBase()
                    Using cm As New SqlCommand(SP, Cn)
                        Cn.Open()
                        cm.CommandType = CommandType.Text
                        Dim da As New SqlDataAdapter
                        da.SelectCommand = cm
                        da.Fill(ds)
                    End Using
                End Using
            Catch ex As Exception
                rsboApp.SetStatusBarMessage("Ejecutar SP: " & ex.Message().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Utilitario.Util_Log.Escribir_Log("Catch Query Consulta : " & ex.Message().ToString(), "ManejoDeDocumentos")
                Return Nothing
            End Try
        End If

        Return ds
    End Function

    Public Function GetSqlConnectionBase() As SqlConnection
        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        Dim cnBaseSAP As New SqlConnection
        Try
            If _tipoManejo <> "A" Then
                BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")
            Else
                BD_User = Functions.VariablesGlobales._vgUserBD
            End If

            If BD_User = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            If _tipoManejo <> "A" Then
                BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")
            Else
                BD_Pass = Functions.VariablesGlobales._vgPassBD
            End If

            If BD_Pass = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            Dim cadena As New SqlConnectionStringBuilder

            If _tipoManejo = "A" Then
                If Not String.IsNullOrEmpty(Functions.VariablesGlobales._vgServerNode) Then
                    cadena.DataSource = Functions.VariablesGlobales._vgServerNode
                    cadena.InitialCatalog = rCompany.CompanyDB
                    cadena.UserID = Functions.VariablesGlobales._vgUserBD
                    cadena.Password = Functions.VariablesGlobales._vgPassBD
                Else
                    cadena.DataSource = rCompany.Server
                    cadena.InitialCatalog = rCompany.CompanyDB
                    cadena.UserID = Functions.VariablesGlobales._vgUserBD
                    cadena.Password = Functions.VariablesGlobales._vgPassBD
                End If
            Else
                cadena.DataSource = rCompany.Server
                cadena.InitialCatalog = rCompany.CompanyDB
                cadena.UserID = BD_User
                cadena.Password = BD_Pass
                Utilitario.Util_Log.Escribir_Log("datos conexion sql User: " + BD_User + " Pass: " + BD_Pass + " tipo: " + _tipoManejo, "ManejoDeDocumentos")
            End If

            cnBaseSAP.ConnectionString = cadena.ConnectionString
            Return cnBaseSAP

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("error GetSqlConnectionBase: " + ex.Message.ToString + " User: " + BD_User + " Pass: " + BD_Pass + " tipo: " + _tipoManejo, "ManejoDeDocumentos")
            Return Nothing
        End Try
    End Function

    Private Function ObtenerColeccion(ByVal Consulta As String, Optional ByVal KeepOpen As Boolean = False) As DataSet
        Dim ds As New DataSet
        Try
            If Consulta = String.Empty Then Return Nothing

            ConectaHANA()

            If CONEXION.State = ConnectionState.Closed Then
                CONEXION.Open()
            End If

            Dim DapTable As New Odbc.OdbcDataAdapter(Consulta, CONEXION)
            DapTable.SelectCommand.CommandTimeout = 0
            DapTable.Fill(ds)

            If Not KeepOpen Then
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If
            End If
            Return ds

        Catch ex As Odbc.OdbcException
            Utilitario.Util_Log.Escribir_Log("ObtenerColeccion: " + ex.Message + " QUERY: " + Consulta.ToString(), "ManejoDeDocumentos")
            Return Nothing
        End Try
    End Function


    Public Function ConectaHANA(Optional ByRef mensaje As String = "") As Boolean
        Dim ConexionHana As String = String.Empty

        Dim BD_User As String = ""
        Dim BD_Pass As String = ""
        Dim _ServerNode As String = ""
        If _tipoManejo = "S" Then
            _ServerNode = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "ServerNode")
            If String.IsNullOrEmpty(_ServerNode) Then
                _ServerNode = rCompany.Server
            End If
            BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")
            BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")
            Utilitario.Util_Log.Escribir_Log("_ServerNode: " + _ServerNode.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("BD_User: " + BD_User.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("BD_Pass: " + BD_Pass.ToString(), "ManejoDeDocumentos")
        End If


        Try


            If _tipoManejo <> "A" Then


            Else
                BD_User = Functions.VariablesGlobales._vgUserBD
            End If
            'BD_User = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_User")

            If BD_User = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion del Usuario Base De Datos, BD_User. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If


            If _tipoManejo <> "A" Then


            Else
                BD_Pass = Functions.VariablesGlobales._vgPassBD
            End If
            'BD_Pass = ConsultaParametro("SAED", "PARAMETROS", "CONFIGURACION", "BD_Pass")

            If BD_Pass = "" Then
                rsboApp.SetStatusBarMessage("GS - No existe configuracion de la Clave del Usuario Base De Datos, BD_Pass. Contacte a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Exit Function
            End If

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then

                If (IntPtr.Size = 8) Then
                    ConexionHana = String.Concat(ConexionHana, "Driver={HDBODBC};")
                Else
                    ConexionHana = String.Concat(ConexionHana, "Driver={HDBODBC32};")
                End If
                If _tipoManejo = "A" Then
                    If Not String.IsNullOrEmpty(Functions.VariablesGlobales._vgServerNode) Then
                        ConexionHana = String.Concat(ConexionHana, "ServerNode=", Functions.VariablesGlobales._vgServerNode & ";")
                        ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._vgUserBD, ";")
                        ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._vgPassBD, ";")
                    Else
                        ConexionHana = String.Concat(ConexionHana, "ServerNode=", rCompany.Server & ";")
                        ConexionHana = String.Concat(ConexionHana, "UID=", Functions.VariablesGlobales._vgUserBD, ";")
                        ConexionHana = String.Concat(ConexionHana, "PWD=", Functions.VariablesGlobales._vgPassBD, ";")
                    End If
                Else

                    ConexionHana = String.Concat(ConexionHana, "ServerNode=", _ServerNode & ";")
                    ConexionHana = String.Concat(ConexionHana, "UID=", BD_User, ";")
                    ConexionHana = String.Concat(ConexionHana, "PWD=", BD_Pass, ";")


                End If


                'pswBD_HANA

                CONEXION = New Odbc.OdbcConnection(ConexionHana)

                If CONEXION.State = ConnectionState.Closed Then
                    CONEXION.Open()
                End If
                If CONEXION.State = ConnectionState.Open Then
                    CONEXION.Close()
                End If

                Return True

                'Else
                '    'CONEXION = New Odbc.OdbcConnection("DRIVER={SQL Server Native Client 10.0}; Server= " & serv & "; Database=" & bd & "; Uid=" & userdb & "; Pwd=" & passdb)
                '    CONEXION = New Odbc.OdbcConnection("DRIVER={" + _driversql + "}; Server= " & serv & "; Database=" & bd & "; Uid=" & userdb & "; Pwd=" & passdb)
                '    'CONEXION = New Odbc.OdbcConnection(GetSqlConnectionBaseString())
                '    If CONEXION.State = ConnectionState.Closed Then
                '        CONEXION.Open()
                '    End If
                '    If CONEXION.State = ConnectionState.Open Then
                '        CONEXION.Close()
                '    End If

                '    Return True

            End If
            Return False

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ConexionHana: " + ConexionHana.ToString(), "ManejoDeDocumentos")
            Utilitario.Util_Log.Escribir_Log("Conecta_HANA: " + ex.Message, "ManejoDeDocumentos")
            Return False

        End Try

#Disable Warning BC42353 ' La función 'ConectaHANA' no devuelve un valor en todas las rutas de acceso de código. ¿Falta alguna instrucción 'Return'?
    End Function

    Public Function ValidarCamposNulos(dataset As DataSet, numTabla As String, ByRef _camponulo As String) As Boolean

        Try
            Dim nombretabla = "Table" & numTabla
            'Dim DescripcionConcepto As String = ""
            'Dim ListaInforAdicional As New List(Of String)
            Dim concepto As String = Nothing
            Dim descripcion As String = Nothing
            For Each table As DataTable In dataset.Tables

                If table.TableName.ToString() = nombretabla Then

                    For rowIndex As Integer = 0 To table.Rows.Count - 1
                        Dim row As DataRow = table.Rows(rowIndex)
                        For columnIndex As Integer = 0 To table.Columns.Count - 1
                            Dim currentColumn As DataColumn = table.Columns(columnIndex)
                            'InforAdicional(rowIndex, columnIndex) = row(currentColumn)
                            If currentColumn.ColumnName = "Descripcion" Then
                                'descripcion = row(currentColumn)
                                If IsDBNull(row(currentColumn)) Then
                                    descripcion = "Nulo"
                                End If
                            End If
                            If currentColumn.ColumnName = "Concepto" Then
                                concepto = row(currentColumn)
                            End If



                        Next
                        'ListaInforAdicional.Add(descripcion & "|" & concepto)
                        'DescripcionConcepto = ""
                        If descripcion = "Nulo" Then
                            _camponulo = concepto.ToString & " se encuentra en nulo, por favor validar"
                            rsboApp.SetStatusBarMessage(table.TableName.ToString() & " Comcepto: " & concepto.ToString & " se encuentra en nulo, por favor validar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Return False
                        End If

                    Next

                    'For Each lista In ListaInforAdicional
                    '    If String.IsNullOrEmpty(lista.Split("|")(0)) Then
                    '        rsboApp.SetStatusBarMessage(table.TableName.ToString() & " Comcepto: " & lista.Split("|")(1).ToString & " se encuentra en nulo, por favor validar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    '        Return False
                    '    End If
                    'Next

                Else
                    For Each row As DataRow In table.Rows
                        For Each column As DataColumn In table.Columns
                            If IsDBNull(row(column)) Then
                                _camponulo = column.ColumnName.ToString & " se encuentra en nulo, por favor validar"
                                rsboApp.SetStatusBarMessage(table.TableName.ToString() & " Columna: " & column.ColumnName & " se encuentra en nulo, por favor validar", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Return False
                            End If

                        Next
                        'Console.WriteLine()
                    Next

                End If



            Next
        Catch ex As Exception
            rsboApp.SetStatusBarMessage("Error en funcion Validar Campos Nulos " & ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try

        Return True
    End Function
End Class

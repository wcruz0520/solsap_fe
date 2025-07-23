Imports System.Data.SqlClient
Imports System.Net

Public Class DatabaseQueryManager
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    Private _tipoManejo As String
    Private oFuncionesAddon As Functions.FuncionesAddon

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

    Private Function ObtenerColeccion(SP As String, flag As Boolean) As DataSet
        ' Esta función es placeholder para mantener compatibilidad
        Return Nothing
    End Function

End Class

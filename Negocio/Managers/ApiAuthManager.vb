Imports Newtonsoft.Json.Linq
Imports System.Net

Public Class ApiAuthManager
    Private _tipoManejo As String
    Private _sboApp As SAPbouiCOM.Application
    Public ReadOnly rCompany As SAPbobsCOM.Company

    Public Sub New(tipoManejo As String, sboApp As SAPbouiCOM.Application, company As SAPbobsCOM.Company)
        _tipoManejo = tipoManejo
        _sboApp = sboApp
        rCompany = company
    End Sub

    Public Sub ActivarTLS()
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or 768 Or 3072
    End Sub

    Public Function ObtenerTokenAutenticacion() As String
        Try
            Dim usuario As String = Functions.VariablesGlobales._ApiAutUser
            Dim password As String = Functions.VariablesGlobales._ApiAutPw
            Dim endpoint As String = Functions.VariablesGlobales._ApiAutSS

            If String.IsNullOrEmpty(usuario) OrElse String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(endpoint) Then
                If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Faltan datos de autenticación (usuario, clave o endpoint)", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return Nothing
            End If

            Dim jsonBody As String = $"{{""usuario"":""{usuario}"", ""password"":""{password}""}}"
            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonBody)
            End Using

            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Using reader As New StreamReader(response.GetResponseStream())
                Dim result As String = reader.ReadToEnd()
                Dim json As JObject = JObject.Parse(result)
                Dim token As String = json("token")?.ToString()

                If Not String.IsNullOrEmpty(token) Then
                    If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Autenticación exitosa", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Return token
                Else
                    If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("No se recibió token de autenticación", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return Nothing
                End If
            End Using

        Catch ex As Exception
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error al autenticar: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try
    End Function

End Class

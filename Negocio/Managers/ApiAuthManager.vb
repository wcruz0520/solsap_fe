Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Net
Imports System.Text
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



    'Public Function ObtenerTokenAutenticacion() As String
    '    Try

    '        'Reutilizar token existente si no ha expirado
    '        If Not String.IsNullOrEmpty(Functions.VariablesGlobales._ApiAuthToken) AndAlso Functions.VariablesGlobales._ApiAuthTokenExpiration > DateTime.UtcNow Then
    '            Return Functions.VariablesGlobales._ApiAuthToken
    '        End If

    '        Dim usuario As String = Functions.VariablesGlobales._ApiAutUser
    '        Dim password As String = Functions.VariablesGlobales._ApiAutPw
    '        Dim endpoint As String = Functions.VariablesGlobales._ApiAutSS

    '        If String.IsNullOrEmpty(usuario) OrElse String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(endpoint) Then
    '            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Faltan datos de autenticación (usuario, clave o endpoint)", SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '            Return Nothing
    '        End If

    '        Dim jsonBody As String = $"{{""usuario"":""{usuario}"", ""password"":""{password}""}}"
    '        Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
    '        request.Method = "POST"
    '        request.ContentType = "application/json"

    '        Using streamWriter As New StreamWriter(request.GetRequestStream())
    '            streamWriter.Write(jsonBody)
    '        End Using

    '        Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
    '        Using reader As New StreamReader(response.GetResponseStream())
    '            Dim result As String = reader.ReadToEnd()
    '            Dim json As JObject = JObject.Parse(result)
    '            Dim token As String = json("token")?.ToString()

    '            If Not String.IsNullOrEmpty(token) Then
    '                If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Token obtenido, pasando a envío", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    '                Return token
    '            Else
    '                If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("No se recibió token de autenticación", SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '                Return Nothing
    '            End If
    '        End Using

    '    Catch ex As Exception
    '        If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error al autenticar: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '        Return Nothing
    '    End Try
    'End Function

    Public Function ObtenerTokenAutenticacion() As String
        Try
            ' Reutilizar token existente si no ha expirado
            If Not String.IsNullOrEmpty(Functions.VariablesGlobales._ApiAuthToken) _
               AndAlso Functions.VariablesGlobales._ApiAuthTokenExpiration > DateTime.UtcNow Then
                Return Functions.VariablesGlobales._ApiAuthToken
            End If

            Dim usuario As String = Functions.VariablesGlobales._ApiAutUser
            Dim password As String = Functions.VariablesGlobales._ApiAutPw
            Dim endpoint As String = Functions.VariablesGlobales._ApiAutSS

            If String.IsNullOrEmpty(usuario) OrElse String.IsNullOrEmpty(password) OrElse String.IsNullOrEmpty(endpoint) Then
                If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Faltan datos de autenticación (usuario, clave o endpoint)", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return Nothing
            End If

            ' Asegurar TLS 1.2 en Framework 4.8 (por si el endpoint lo exige)
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls

            Dim jsonBody As String = $"{{""usuario"":""{usuario}"", ""password"":""{password}""}}"
            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Timeout = 1000 * 60 ' 60s

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonBody)
            End Using

            Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(response.GetResponseStream())
                    Dim result As String = reader.ReadToEnd()
                    Dim json As JObject = JObject.Parse(result)
                    Dim token As String = json("token")?.ToString()

                    If String.IsNullOrEmpty(token) Then
                        If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("No se recibió token de autenticación", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Return Nothing
                    End If

                    ' Decodificar el JWT para obtener exp (sin usar FromUnixTimeSeconds)
                    Try
                        Dim parts() As String = token.Split("."c)
                        If parts.Length >= 2 Then
                            Dim payloadJson As String = Base64UrlToString(parts(1))
                            Dim payloadObj As JObject = JObject.Parse(payloadJson)

                            Dim expStr As String = payloadObj("exp")?.ToString()
                            Dim expLong As Long
                            If Not String.IsNullOrEmpty(expStr) AndAlso Long.TryParse(expStr, expLong) Then
                                Functions.VariablesGlobales._ApiAuthTokenExpiration = UnixSecondsToUtc(expLong)
                            ElseIf json("expires_in") IsNot Nothing AndAlso Long.TryParse(json("expires_in").ToString(), expLong) Then
                                ' Fallback: si la API devuelve expires_in en segundos
                                Functions.VariablesGlobales._ApiAuthTokenExpiration = DateTime.UtcNow.AddSeconds(expLong - 30) ' pequeño margen
                            Else
                                ' Si no hay exp, forzar renovación pronto
                                Functions.VariablesGlobales._ApiAuthTokenExpiration = DateTime.UtcNow.AddMinutes(5)
                            End If
                        Else
                            ' Token inesperado, renovar pronto
                            Functions.VariablesGlobales._ApiAuthTokenExpiration = DateTime.UtcNow.AddMinutes(5)
                        End If
                    Catch
                        ' Si falló la decodificación, no usar el token por mucho tiempo
                        Functions.VariablesGlobales._ApiAuthTokenExpiration = DateTime.UtcNow
                    End Try

                    Functions.VariablesGlobales._ApiAuthToken = token
                    If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Token obtenido, pasando a envío", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Return token
                End Using
            End Using

        Catch ex As WebException
            Dim msg As String = ex.Message
            If ex.Response IsNot Nothing Then
                Try
                    Using sr As New StreamReader(ex.Response.GetResponseStream())
                        Dim body = sr.ReadToEnd()
                        If Not String.IsNullOrEmpty(body) Then msg &= " | " & body
                    End Using
                Catch
                End Try
            End If
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error al autenticar: " & msg, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        Catch ex As Exception
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error al autenticar: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try
    End Function

    ' === Helpers .NET Framework 4.8 ===

    Private Shared ReadOnly UnixEpochUtc As DateTime = New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)

    Private Function UnixSecondsToUtc(seconds As Long) As DateTime
        ' Reemplaza a DateTimeOffset.FromUnixTimeSeconds en .NET Fx 4.8
        Return UnixEpochUtc.AddSeconds(seconds)
    End Function

    Private Function Base64UrlToString(base64Url As String) As String
        Dim s As String = base64Url.Replace("-"c, "+"c).Replace("_"c, "/"c)
        Select Case (s.Length Mod 4)
            Case 2 : s &= "=="
            Case 3 : s &= "="
            Case 1 : s &= "==="
        End Select
        Dim bytes = Convert.FromBase64String(s)
        Return Encoding.UTF8.GetString(bytes)
    End Function
End Class

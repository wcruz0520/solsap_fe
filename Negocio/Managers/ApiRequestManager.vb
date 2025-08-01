Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Net

Public Class ApiRequestManager
    Private _tipoManejo As String
    Private _sboApp As SAPbouiCOM.Application
    Private _auth As ApiAuthManager

    Public Sub New(tipoManejo As String, sboApp As SAPbouiCOM.Application, auth As ApiAuthManager)
        _tipoManejo = tipoManejo
        _sboApp = sboApp
        _auth = auth
    End Sub

    Public Function EnviarFacturaSolsap(factura As Entidades.RequestFactura) As Entidades.ResponseDocuments
        Try
            _auth.ActivarTLS()
            Dim token As String = _auth.ObtenerTokenAutenticacion()
            If String.IsNullOrEmpty(token) Then Return Nothing

            Dim endpoint As String = Functions.VariablesGlobales._ApiFactEmiSS
            If String.IsNullOrEmpty(endpoint) Then Return Nothing

            Dim settings As New JsonSerializerSettings()
            settings.NullValueHandling = NullValueHandling.Ignore
            Dim jsonBody As String = JsonConvert.SerializeObject(factura, settings)

            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Headers.Add("Authorization", $"Bearer {token}")

            Using sw As New StreamWriter(request.GetRequestStream())
                sw.Write(jsonBody)
            End Using

            Using resp As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(resp.GetResponseStream())
                    Dim result As String = reader.ReadToEnd()
                    Dim json As JObject = JObject.Parse(result)
                    Dim estado As String = json.SelectToken("data.result.codigo")?.ToString()
                    Dim mensaje As String = json.SelectToken("data.result.mensaje")?.ToString()
                    Dim clave As String = json.SelectToken("data.result.claveAcceso")?.ToString()
                    Dim identificador As String = json.SelectToken("data.result.identificador")?.ToString()

                    Dim respuesta As New Entidades.ResponseDocuments()
                    respuesta.codigo = estado
                    respuesta.mensaje = mensaje
                    respuesta.claveAcceso = clave
                    respuesta.identificador = identificador
                    Return respuesta
                End Using
            End Using

        Catch webEx As WebException
            Dim mensajeError As String = webEx.Message
            Dim respErr As HttpWebResponse = TryCast(webEx.Response, HttpWebResponse)
            If respErr IsNot Nothing Then
                Using reader As New StreamReader(respErr.GetResponseStream())
                    mensajeError &= " - " & reader.ReadToEnd()
                End Using
            End If
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error enviando factura: " & mensajeError, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing

        Catch ex As Exception
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error enviando factura: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try
    End Function

    Public Function EnviarNCSolsap(NCredito As Entidades.RequestNotaCredito) As Entidades.ResponseDocuments
        Try
            _auth.ActivarTLS()
            Dim token As String = _auth.ObtenerTokenAutenticacion()
            If String.IsNullOrEmpty(token) Then Return Nothing

            Dim endpoint As String = Functions.VariablesGlobales._ApiNCEmiSS
            If String.IsNullOrEmpty(endpoint) Then Return Nothing

            Dim settings As New JsonSerializerSettings()
            settings.NullValueHandling = NullValueHandling.Ignore
            Dim jsonBody As String = JsonConvert.SerializeObject(NCredito, settings)

            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Headers.Add("Authorization", $"Bearer {token}")

            Using sw As New StreamWriter(request.GetRequestStream())
                sw.Write(jsonBody)
            End Using

            Using resp As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(resp.GetResponseStream())
                    Dim result As String = reader.ReadToEnd()
                    Dim json As JObject = JObject.Parse(result)
                    Dim estado As String = json.SelectToken("data.result.codigo")?.ToString()
                    Dim mensaje As String = json.SelectToken("data.result.mensaje")?.ToString()
                    Dim clave As String = json.SelectToken("data.result.claveAcceso")?.ToString()
                    Dim identificador As String = json.SelectToken("data.result.identificador")?.ToString()

                    Dim respuesta As New Entidades.ResponseDocuments()
                    respuesta.codigo = estado
                    respuesta.mensaje = mensaje
                    respuesta.claveAcceso = clave
                    respuesta.identificador = identificador
                    Return respuesta
                End Using
            End Using

        Catch webEx As WebException
            Dim mensajeError As String = webEx.Message
            Dim respErr As HttpWebResponse = TryCast(webEx.Response, HttpWebResponse)
            If respErr IsNot Nothing Then
                Using reader As New StreamReader(respErr.GetResponseStream())
                    mensajeError &= " - " & reader.ReadToEnd()
                End Using
            End If
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error enviando nota de crédito: " & mensajeError, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing

        Catch ex As Exception
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error enviando nota de crédito: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try
    End Function

    Public Function ConsultarDoc(reqEstado As Entidades.RequestConsulta) As Entidades.ResponseConsulta
        Try
            _auth.ActivarTLS()
            Dim token As String = _auth.ObtenerTokenAutenticacion()
            If String.IsNullOrEmpty(token) Then Return Nothing

            Dim endpoint As String = Functions.VariablesGlobales._ApiConEstSS
            If String.IsNullOrEmpty(endpoint) Then Return Nothing

            Dim settings As New JsonSerializerSettings()
            settings.NullValueHandling = NullValueHandling.Ignore
            Dim jsonBody As String = JsonConvert.SerializeObject(reqEstado, settings)

            Dim request As HttpWebRequest = CType(WebRequest.Create(endpoint), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Headers.Add("Authorization", $"Bearer {token}")

            Using sw As New StreamWriter(request.GetRequestStream())
                sw.Write(jsonBody)
            End Using

            Using resp As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(resp.GetResponseStream())
                    Dim result As String = reader.ReadToEnd()
                    Dim json As JObject = JObject.Parse(result)
                    'Dim estado As String = json.SelectToken("data.result.codigo")?.ToString()
                    Dim mensaje As String = json.SelectToken("data.result.mensaje")?.ToString()
                    Dim clave As String = json.SelectToken("data.result.claveAcceso")?.ToString()
                    Dim tipoComprobante As String = json.SelectToken("data.result.tipoComprobante")
                    Dim fechaAutorizacion As String = json.SelectToken("data.result.fechaAutorizacion")?.ToString()
                    Dim codigo As String = json.SelectToken("data.result.codigo")?.ToString()

                    Dim respuesta As New Entidades.ResponseConsulta()

                    respuesta.mensaje = mensaje
                    respuesta.claveAcceso = clave
                    respuesta.tipoComprobante = tipoComprobante
                    respuesta.fechaAutorizacion = fechaAutorizacion
                    respuesta.codigo = codigo
                    Return respuesta
                End Using
            End Using

        Catch webEx As WebException
            Dim mensajeError As String = webEx.Message
            Dim respErr As HttpWebResponse = TryCast(webEx.Response, HttpWebResponse)
            If respErr IsNot Nothing Then
                Using reader As New StreamReader(respErr.GetResponseStream())
                    mensajeError &= " - " & reader.ReadToEnd()
                End Using
            End If
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error consultar documento SRI: " & mensajeError, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing

        Catch ex As Exception
            If _tipoManejo = "A" Then _sboApp.SetStatusBarMessage("Error consultar documento SRI: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return Nothing
        End Try

    End Function

    Public Function ConsultaPDF(sClaveAcceso As String, tipoDoc As String) As Boolean
        Try
            Dim taxIdNum As String = ""
            Dim TbDoc As String = ""
            Dim oRs As SAPbobsCOM.Recordset = CType(_auth.rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRs.DoQuery("SELECT ""TaxIdNum"" FROM ""OADM""")
            If Not oRs.EoF Then
                taxIdNum = oRs.Fields.Item("TaxIdNum").Value.ToString().Trim()
            End If

            Dim urlReal As String = ""
            Select Case tipoDoc
                Case "OINV"
                    TbDoc = "facturas"
                Case "ORIN"
                    TbDoc = "notacredito"
                Case Else
                    _sboApp.SetStatusBarMessage("Tipo de documento no soportado.", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
            End Select

            Try
                urlReal = $"{Functions.VariablesGlobales._ApiPDFSS}/{TbDoc}/{taxIdNum}/{sClaveAcceso}"
            Catch ex As Exception
                _sboApp.SetStatusBarMessage("Error al descargar o abrir el PDF: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

            Dim rutaTempPDF As String = Path.Combine(Path.GetTempPath(), $"{sClaveAcceso}.pdf")
            Using wc As New WebClient()
                wc.DownloadFile(urlReal, rutaTempPDF)
            End Using

            Dim proc As New Process()
            proc.StartInfo.FileName = rutaTempPDF
            proc.Start()
            proc.Dispose()

            Return True

        Catch ex As Exception
            _sboApp.SetStatusBarMessage("Error al descargar o abrir el PDF: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return False
        End Try
    End Function

    Public Function ConsultaXML(sClaveAcceso As String, tipoDoc As String) As Boolean
        Try
            Dim taxIdNum As String = ""
            Dim TbDoc As String = ""
            Dim oRs As SAPbobsCOM.Recordset = CType(_auth.rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRs.DoQuery("SELECT ""TaxIdNum"" FROM ""OADM""")
            If Not oRs.EoF Then
                taxIdNum = oRs.Fields.Item("TaxIdNum").Value.ToString().Trim()
            End If

            Dim urlReal As String = ""
            Select Case tipoDoc
                Case "OINV"
                    TbDoc = "facturas"
                Case "ORIN"
                    TbDoc = "notacredito"
                Case Else
                    _sboApp.SetStatusBarMessage("Tipo de documento no soportado.", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
            End Select

            Try
                urlReal = $"{Functions.VariablesGlobales._ApiXMLSS}/{TbDoc}/{taxIdNum}/{sClaveAcceso}"
            Catch ex As Exception
                _sboApp.SetStatusBarMessage("Error al descargar o abrir el XML: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

            Dim rutaTempPDF As String = Path.Combine(Path.GetTempPath(), $"{sClaveAcceso}.xml")
            Using wc As New WebClient()
                wc.DownloadFile(urlReal, rutaTempPDF)
            End Using

            Dim proc As New Process()
            proc.StartInfo.FileName = rutaTempPDF
            proc.Start()
            proc.Dispose()

            Return True

        Catch ex As Exception
            _sboApp.SetStatusBarMessage("Error al descargar o abrir el XML: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return False
        End Try
    End Function
End Class

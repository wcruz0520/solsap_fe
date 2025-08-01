Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports Negocio.Documentos

Public Class DesencriptarQuery

    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application
    Private _tipoManejo As String
    Private oFuncionesAddon As Functions.FuncionesAddon
    Public Const sKey As String = "S01s7p1" ' CLAVE DE ENCRIPTACION LICENCIA
    Dim QueryDesencriptado As String = ""

    Public Sub New(company As SAPbobsCOM.Company, sboApp As SAPbouiCOM.Application, tipoManejo As String, funciones As Functions.FuncionesAddon)
        rCompany = company
        rsboApp = sboApp
        _tipoManejo = tipoManejo
        oFuncionesAddon = funciones
    End Sub

    Public Function GetQueryConsulta(ByVal tipodoc As tipoDocumento, ByVal docentry As Integer, Optional ByVal Seccion As String = "") As String

        Try
            Dim partesQuerys As New List(Of String)
            Select Case Seccion

                Case "EXPORT"
                    partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_CompleExportacion.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                Case "REEMBOLSO"
                    partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_CompleReembolso.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                Case "DOCSENV"
                    partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_DocumentosEnviados.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))


                Case Else

                    If tipodoc = tipoDocumento.Factura Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.FacturaAnticipo Then

                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaAnticipoSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_FacturaAnticipoSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.NotaCredito Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaCreditoSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaCreditoSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.NotaDebito Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaDebitoSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_NotaDebitoSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf (tipodoc = tipoDocumento.GuiaRemisionEntrega) Or (tipodoc = tipoDocumento.GuiaRemisionTraslado) Or (tipodoc = tipoDocumento.GuiaRemisionSolicitudTraslado) Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiaRemisionSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiaRemisionSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf (tipodoc = tipoDocumento.Retencion) Or (tipodoc = tipoDocumento.RetencionNotaDebito) Or (tipodoc = tipoDocumento.RetencionAnticipo) Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_RetencionSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_RetencionSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))


                    ElseIf tipodoc = tipoDocumento.Liquidacion Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_LiquidacionSeccion01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_LiquidacionSeccion02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    ElseIf tipodoc = tipoDocumento.GuiaRemisionDesatendida Then
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiasDesatendidas01.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))
                        partesQuerys.Add(Utilitario.Util_Encriptador.Desencriptar(Functions.VariablesGlobales._Query_GuiasDesatendidas02.ToString().Replace("{", "").Replace("}", "").ToString(), sKey))

                    End If
            End Select

            'Procesamiento de Querys
            QueryDesencriptado = ""

            For Each querysession As String In partesQuerys
                QueryDesencriptado = QueryDesencriptado + querysession
            Next

            'remplazamos 3 parametros
            Dim midata As String = QueryDesencriptado

            Select Case tipodoc
                Case tipoDocumento.Factura
                   ' midata = midata.Replace("A.""Docentry""=@DocKey", "A.""Docentry""=@DocKey AND A.""DocSubType"" <> 'DN'")
                Case tipoDocumento.NotaCredito
                   ' midata = midata.Replace("INV", "RIN")
                Case tipoDocumento.NotaDebito
                   ' midata = midata.Replace("A.""Docentry""=@DocKey", "A.""Docentry""=@DocKey AND A.""DocSubType"" = 'DN'")
                Case tipoDocumento.Retencion
                   ' midata = midata.Replace("INV", "PCH")
                Case tipoDocumento.FacturaAnticipo
                    midata = midata.Replace("INV", "DPI")
                Case tipoDocumento.RetencionAnticipo
                    midata = midata.Replace("PCH", "DPO")
                Case tipoDocumento.RetencionNotaDebito
                   ' midata = midata.Replace("PCH", "DPO")
                Case tipoDocumento.GuiaRemisionEntrega
                    'midata = midata.Replace("INV", "DLN")
                Case tipoDocumento.GuiaRemisionTraslado
                    midata = midata.Replace("DLN", "WTR")
                Case tipoDocumento.GuiaRemisionSolicitudTraslado
                    midata = midata.Replace("DLN", "WTQ")
            End Select

            'EL DOcentry
            midata = midata.Replace("@DocKey", docentry.ToString)
            midata = midata.Replace("@TipoDoc", "'" + tipodoc.ToString() + "'")
            midata = midata.Replace("@GS_SS_NAMEDB", rCompany.CompanyDB)

            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                midata = HanaTablesSapReplace(midata)
            End If

            Return midata

        Catch ex As Exception
            Return "GSCODEEXCEPCION :" & ex.Message
        End Try
    End Function

    Public Sub SetProtocolosdeSeguridad()
        'PARA TLS 1.2
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
        ServicePointManager.DefaultConnectionLimit = 9999

        'PARA HTTPS
        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf customCertValidation)
    End Sub

    Shared Function customCertValidation(ByVal sender As Object,
                                             ByVal cert As X509Certificate,
                                             ByVal chain As X509Chain,
                                             ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function

    Public Function HanaTablesSapReplace(s As String) As String

        Dim MYSP As String = ""

        MYSP = s.Replace("""OINV""", rCompany.CompanyDB + ".""OINV""")
        MYSP = MYSP.Replace("""OCRD""", rCompany.CompanyDB + ".""OCRD""")
        MYSP = MYSP.Replace("""OITM""", rCompany.CompanyDB + ".""OITM""")
        MYSP = MYSP.Replace("""OPLN""", rCompany.CompanyDB + ".""OPLN""")
        MYSP = MYSP.Replace("""OUSR""", rCompany.CompanyDB + ".""OUSR""")
        MYSP = MYSP.Replace("""OEXD""", rCompany.CompanyDB + ".""OEXD""")
        MYSP = MYSP.Replace("""OSLP""", rCompany.CompanyDB + ".""OSLP""")
        MYSP = MYSP.Replace("""OCTG""", rCompany.CompanyDB + ".""OCTG""")
        MYSP = MYSP.Replace("""OCRN""", rCompany.CompanyDB + ".""OCRN""")
        MYSP = MYSP.Replace("""ORTT""", rCompany.CompanyDB + ".""ORTT""")
        MYSP = MYSP.Replace("""OIBT""", rCompany.CompanyDB + ".""OIBT""")
        MYSP = MYSP.Replace("""OITB""", rCompany.CompanyDB + ".""OITB""")
        MYSP = MYSP.Replace("""ORIN""", rCompany.CompanyDB + ".""ORIN""")
        MYSP = MYSP.Replace("""ODLN""", rCompany.CompanyDB + ".""ODLN""")
        MYSP = MYSP.Replace("""OWHT""", rCompany.CompanyDB + ".""OWHT""")
        MYSP = MYSP.Replace("""OWHS""", rCompany.CompanyDB + ".""OWHS""")
        MYSP = MYSP.Replace("""OCRG""", rCompany.CompanyDB + ".""OCRG""")
        MYSP = MYSP.Replace("""OPCH""", rCompany.CompanyDB + ".""OPCH""")
        MYSP = MYSP.Replace("""CUFD""", rCompany.CompanyDB + ".""CUFD""")
        MYSP = MYSP.Replace("""OSTA""", rCompany.CompanyDB + ".""OSTA""")
        MYSP = MYSP.Replace("""OWTR""", rCompany.CompanyDB + ".""OWTR""")
        MYSP = MYSP.Replace("""OITW""", rCompany.CompanyDB + ".""OITW""")
        MYSP = MYSP.Replace("""OCRY""", rCompany.CompanyDB + ".""OCRY""")
        MYSP = MYSP.Replace("""NNM1""", rCompany.CompanyDB + ".""NNM1""")
        MYSP = MYSP.Replace("""ODPI""", rCompany.CompanyDB + ".""ODPI""")
        MYSP = MYSP.Replace("""OADM""", rCompany.CompanyDB + ".""OADM""")
        MYSP = MYSP.Replace("""ODPO""", rCompany.CompanyDB + ".""ODPO""")
        MYSP = MYSP.Replace("""OCPR""", rCompany.CompanyDB + ".""OCPR""")
        MYSP = MYSP.Replace("""OBTN""", rCompany.CompanyDB + ".""OBTN""")
        MYSP = MYSP.Replace("""OITL""", rCompany.CompanyDB + ".""OITL""")

        'Logica para que dependiendo de una Opcion del Addon replace tablas que no se encuentren
        If Functions.VariablesGlobales._TablasNativasReplace <> "" Then

            Dim rtablas = Functions.VariablesGlobales._TablasNativasReplace.Split(";")

            For Each t In rtablas

                MYSP = MYSP.Replace($"""{t}""", rCompany.CompanyDB + $".""{t}""")

            Next

        End If

        'Remplazando sub tablas
        For i As Integer = 1 To 12

            MYSP = MYSP.Replace("""INV" & i.ToString & """", rCompany.CompanyDB + ".""INV" & i.ToString & """")
            MYSP = MYSP.Replace("""CRD" & i.ToString & """", rCompany.CompanyDB + ".""CRD" & i.ToString & """")
            MYSP = MYSP.Replace("""ITM" & i.ToString & """", rCompany.CompanyDB + ".""ITM" & i.ToString & """")
            MYSP = MYSP.Replace("""PLN" & i.ToString & """", rCompany.CompanyDB + ".""PLN" & i.ToString & """")
            MYSP = MYSP.Replace("""USR" & i.ToString & """", rCompany.CompanyDB + ".""USR" & i.ToString & """")
            MYSP = MYSP.Replace("""EXD" & i.ToString & """", rCompany.CompanyDB + ".""EXD" & i.ToString & """")
            MYSP = MYSP.Replace("""SLP" & i.ToString & """", rCompany.CompanyDB + ".""SLP" & i.ToString & """")
            MYSP = MYSP.Replace("""CTG" & i.ToString & """", rCompany.CompanyDB + ".""CTG" & i.ToString & """")
            MYSP = MYSP.Replace("""CRN" & i.ToString & """", rCompany.CompanyDB + ".""CRN" & i.ToString & """")
            MYSP = MYSP.Replace("""RTT" & i.ToString & """", rCompany.CompanyDB + ".""RTT" & i.ToString & """")
            MYSP = MYSP.Replace("""IBT" & i.ToString & """", rCompany.CompanyDB + ".""IBT" & i.ToString & """")
            MYSP = MYSP.Replace("""ITB" & i.ToString & """", rCompany.CompanyDB + ".""ITB" & i.ToString & """")
            MYSP = MYSP.Replace("""RIN" & i.ToString & """", rCompany.CompanyDB + ".""RIN" & i.ToString & """")
            MYSP = MYSP.Replace("""DLN" & i.ToString & """", rCompany.CompanyDB + ".""DLN" & i.ToString & """")
            MYSP = MYSP.Replace("""WHT" & i.ToString & """", rCompany.CompanyDB + ".""WHT" & i.ToString & """")
            MYSP = MYSP.Replace("""WHS" & i.ToString & """", rCompany.CompanyDB + ".""WHS" & i.ToString & """")
            MYSP = MYSP.Replace("""CRG" & i.ToString & """", rCompany.CompanyDB + ".""CRG" & i.ToString & """")
            MYSP = MYSP.Replace("""PCH" & i.ToString & """", rCompany.CompanyDB + ".""PCH" & i.ToString & """")
            MYSP = MYSP.Replace("""WTR" & i.ToString & """", rCompany.CompanyDB + ".""WTR" & i.ToString & """")
            MYSP = MYSP.Replace("""DPI" & i.ToString & """", rCompany.CompanyDB + ".""DPI" & i.ToString & """")
            MYSP = MYSP.Replace("""ADM" & i.ToString & """", rCompany.CompanyDB + ".""ADM" & i.ToString & """")
            MYSP = MYSP.Replace("""DPO" & i.ToString & """", rCompany.CompanyDB + ".""DPO" & i.ToString & """")
            MYSP = MYSP.Replace("""CPR" & i.ToString & """", rCompany.CompanyDB + ".""CPR" & i.ToString & """")
            MYSP = MYSP.Replace("""BTN" & i.ToString & """", rCompany.CompanyDB + ".""BTN" & i.ToString & """")
            MYSP = MYSP.Replace("""ITL" & i.ToString & """", rCompany.CompanyDB + ".""ITL" & i.ToString & """")


            'Logica para que dependiendo de una Opcion del Addon replace tablas que no se encuentren
            If Functions.VariablesGlobales._TablasNativasReplace <> "" Then

                Dim rtablas = Functions.VariablesGlobales._TablasNativasReplace.Split(";")

                For Each t In rtablas

                    MYSP = MYSP.Replace($"""{t.Substring(1)}" & i.ToString & """", rCompany.CompanyDB + $".""{t.Substring(1)}" & i.ToString & """")

                Next

            End If

        Next

        'tablas de usuario
        MYSP = MYSP.Replace("""@", String.Format("""{0}"".""@", rCompany.CompanyDB))


        Return MYSP

    End Function
End Class

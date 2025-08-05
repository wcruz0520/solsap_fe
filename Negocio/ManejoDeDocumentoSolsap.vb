'Imports Entidades
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
'https
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Functions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SAPbobsCOM
Imports System.Globalization
Imports Spire.Pdf
Imports Spire.Pdf.AutomaticFields
Imports Spire.Pdf.Graphics

Public Class ManejoDeDocumentoSolsap
    Private rCompany As SAPbobsCOM.Company
    Private rsboApp As SAPbouiCOM.Application

    ''' OBSERVACION, ESTE USUARIO Y CLAVE LO DEBE TOMAR DESDE LA TABLA DE CONFIGURACION QUE DEBE SER UN UDO,
    ''' NO HABRÍA PROBLEMA YA QUE LA CONSULTA LA HARÍA POR LA DIAPI,
    ''' YA QUE ESTE USUARIO Y CLAVE LO USA SOLO PARA EJECUTAR LOS QUERY
    ''' OJO CON EL SERVICIO.

    Private _EstadoAutorizacion As String = ""
    Private _ClaveAcceso As String = ""
    Private _Observacion As String = ""
    Private _CampoNulo As String = ""
    Private _NumAutorizacion As String = ""
    Private _FechaAutorizacion As Date
    Private _EstadoSAP As String = ""
    Private _Error As String = ""
    Dim mensaje As String = ""
    Dim oObjeto As Object
    'Dim ObjetoRespuesta As Object = Nothing
    Dim oFuncionesAddon As FuncionesAddon

    Dim oFuncionesB1 As FuncionesB1

    Dim _tipoManejo As String
    Dim _errorMensajeWSEnvío As String
    Public _Nombre_Proveedor_SAP_BO As String = ""

    Dim proxyobject As System.Net.WebProxy
    Dim cred As System.Net.NetworkCredential

    Dim oDocumento As SAPbobsCOM.Documents

    Dim _GuardarLog As String = "N"

    Private _NumeroDeDocumentoSRI As String = ""

    Dim mensajeDocAut As String = ""

    Dim ApiAuthManager_ As ApiAuthManager
    Dim ApiRequestManager_ As ApiRequestManager
    Dim DesencriptarQuery_ As DesencriptarQuery
    Dim DatabaseQueryManager_ As DatabaseQueryManager
    Dim FacturaManager_ As FacturaManager
    Dim ProcesoEmision_ As ProcesoEmision

    ''' <summary>
    ''' Tipo Manejo, A - Addon, S -  Servicio
    ''' </summary>
    ''' <param name="Company"></param>
    ''' <param name="sboApp"></param>
    ''' <param name="tipoManejo"></param>
    ''' <remarks>Tipo Manejo, A - Addon, S -  Servicio</remarks>
    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application, tipoManejo As String, ByVal ProveedorSAPBO As String)
        'Utilitario.Util_Log.Escribir_Log("SubNew Inicio", "ManejoDeDocumentos")
        rCompany = Company
        _tipoManejo = tipoManejo
        _Nombre_Proveedor_SAP_BO = ProveedorSAPBO
        If tipoManejo = "A" Then
            rsboApp = sboApp
            oFuncionesAddon = New Functions.FuncionesAddon(rCompany, rsboApp, True, False)
            oFuncionesB1 = New Functions.FuncionesB1(rCompany, rsboApp, True, False)
        Else
            ' SI ES SERVICIO INSTANCIO ESTA CLASE, YA QUE NO USA LA UIAPI
            oFuncionesAddon = New Functions.FuncionesAddon(rCompany, rsboApp, True, False)
        End If

        ApiAuthManager_ = New Negocio.ApiAuthManager(_tipoManejo, rsboApp, rCompany)
        ApiRequestManager_ = New Negocio.ApiRequestManager(_tipoManejo, rsboApp, ApiAuthManager_)
        DesencriptarQuery_ = New Negocio.DesencriptarQuery(rCompany, rsboApp, _tipoManejo, oFuncionesAddon)
        DatabaseQueryManager_ = New Negocio.DatabaseQueryManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon)
        FacturaManager_ = New Negocio.FacturaManager(rCompany, rsboApp, _tipoManejo, oFuncionesAddon, DatabaseQueryManager_, DesencriptarQuery_)
        ProcesoEmision_ = New Negocio.ProcesoEmision(rCompany, rsboApp, tipoManejo, oFuncionesAddon)
    End Sub

    Public Function ProcesaEnvioDocumento(DocEntry As Integer, TipoDocumento As String, Optional ByVal sincronizado As Boolean = False) As String
        Return ProcesoEmision_.ProcesaEnvioDocumento(DocEntry, TipoDocumento, _Error, _Error, _EstadoAutorizacion, _ClaveAcceso, _NumAutorizacion, _Observacion, _FechaAutorizacion, _NumeroDeDocumentoSRI, mensaje, _errorMensajeWSEnvío, oObjeto, _CampoNulo, _Nombre_Proveedor_SAP_BO, sincronizado)
    End Function

    Public Function ConsultaPDF(sClaveAcceso As String, tipoDoc As String, Optional docsubtype As String = "") As Boolean
        Return ApiRequestManager_.ConsultaPDF(sClaveAcceso, tipoDoc, docsubtype)
    End Function

    Public Function ConsultaXML(sClaveAcceso As String, tipoDoc As String, Optional docsubtype As String = "") As Boolean
        Return ApiRequestManager_.ConsultaXML(sClaveAcceso, tipoDoc, docsubtype)
    End Function

    Private Function FormatearNumero(valor As Object) As String
        If valor Is Nothing Then Return Nothing

        Dim numero As Decimal
        If Decimal.TryParse(valor.ToString(), numero) Then
            Return numero.ToString(CultureInfo.InvariantCulture)
        End If

        Return valor.ToString()
    End Function

    Public Function AbrirEnlaceExterno(enlace As String) As Boolean
        Try
            If Not String.IsNullOrEmpty(enlace) Then
                Dim rn As New System.Diagnostics.Process
                rn.StartInfo.FileName = enlace
                rn.Start()
                rn.Dispose()
                Return True
            End If
        Catch ex As Exception
        End Try
        Return False
    End Function
End Class

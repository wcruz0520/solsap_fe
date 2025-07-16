Imports ZXing

Public Class ClsGeneradorQR


    Public Shared Function SaveStringToQR(ByVal Datos As String, NombreImg As String, alto As Integer, ancho As Integer, Optional ByRef msgg As String = "") As String

        Try

            Dim cadenaQR As String = ""

            cadenaQR = Datos

            'Dim rutaTemporal = System.IO.Path.GetTempPath() & Guid.NewGuid().ToString("N") & ".Png"

            Dim rutaTemporal = System.IO.Path.GetTempPath() & NombreImg & ".Png"

            If Not System.IO.File.Exists(rutaTemporal) Then

                Dim escritor As New BarcodeWriter

                escritor.Format = BarcodeFormat.QR_CODE
                '236 px = 2cm tamano minimos establecido por la Dian
                escritor.Options.Height = alto
                escritor.Options.Width = ancho

                Dim lienso As System.Drawing.Bitmap

                lienso = escritor.Write(cadenaQR)

                lienso.Save(rutaTemporal, System.Drawing.Imaging.ImageFormat.Png)

            End If

            msgg = "ok"

            Return rutaTemporal
        Catch ex As Exception

            msgg = "Ocurrio un error al generar el QR : " & ex.Message

        End Try

        Return String.Empty

    End Function


End Class

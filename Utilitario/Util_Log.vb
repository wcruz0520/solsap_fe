Imports System.IO
Imports System.Text

Public Class Util_Log


    Public Shared Sub Escribir_Log(ByVal sMsg As String, ByVal Transaccion As String)
        Try
            'Dim sRutaCarpeta As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase & "LOG\"
            Dim sRutaCarpeta As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOG_SAED\"
            Dim sRuta As String = sRutaCarpeta & Transaccion + ".txt"

            'If Not System.IO.Directory.Exists(sRutaCarpeta) Then
            '    System.IO.Directory.CreateDirectory(sRutaCarpeta)
            'End If
            'If Not Directory.Exists(sRutaCarpeta) Then

            '    Directory.CreateDirectory(sRutaCarpeta)

            'End If
            ' SI EXISTE UNA CARPETA LLAMADA LOG, ESCRIBIRÀ
            If System.IO.Directory.Exists(sRutaCarpeta) Then
                If Not File.Exists(sRuta) Then
                    Dim strStreamW As Stream = Nothing
                    Dim strStreamWriter As StreamWriter = Nothing

                    strStreamW = File.Create(sRuta) ' lo creamos
                    strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default) '
                    strStreamWriter.Close() ' cerramos
                End If

                Dim sTexto As New StringBuilder

                sTexto.AppendLine("FECHA: " & Now)
                sTexto.AppendLine("----------------------------------------------------------")
                sTexto.AppendLine(sMsg.ToString())

                Try
                    Dim oTextWriter As TextWriter = New StreamWriter(sRuta, True)
                    oTextWriter.WriteLine(sTexto.ToString)
                    oTextWriter.Flush()
                    oTextWriter.Close()
                    oTextWriter = Nothing

                Catch ex As Exception
                    ' EventLog.WriteEntry("MyWindowsService", "Error: " & ex.Message.ToString)
                End Try

            End If

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' Guarda el Log Emisión en una carpeta "EMISION"
    ''' Transaccion = "Creacion", "Reenvío", ","Consulta PDF", "Envío Mail"
    ''' </summary>
    ''' <param name="sPathName"></param>
    ''' <param name="sErrMsg"></param>
    ''' <param name="Transaccion"></param>
    ''' <param name="TipoDocumento"></param>
    ''' <param name="DocNum"></param>
    ''' <remarks></remarks>
    Public Shared Sub LogEmisión(ByVal sPathName As String, ByVal sErrMsg As String, ByVal Transaccion As String, TipoDocumento As String, Optional ByVal DocNum As String = "0")
        Try
            Dim sLogFormat As String = ""
            Dim sErrorTime As String = ""

            Dim sYear As String = DateTime.Now.Year.ToString()
            Dim sMonth As String = DateTime.Now.Month.ToString()
            Dim sDay As String = DateTime.Now.Day.ToString()
            sErrorTime = sYear + sMonth + sDay

            sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToLongTimeString().ToString() + " ==> "

            sPathName += "\Emision\" + Transaccion + "\" + DocNum

            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(sPathName, True)
            sw.WriteLine(sLogFormat + sErrMsg)
            sw.Flush()
            sw.Close()

        Catch ex As Exception
            Dim fi As System.IO.FileInfo = New System.IO.FileInfo(sPathName)
            System.IO.Directory.CreateDirectory(fi.DirectoryName)
            Escribir_Log(sErrMsg, "Util_Log")
        End Try

    End Sub

    Structure Transacciones
        Const Creacion = "Creacion"
        Const Reenvío = "Reenvío"
        Const Consulta_PDF = "Consulta_PDF"
        Const Envío_Mail = "Envío_Mail"
    End Structure

End Class

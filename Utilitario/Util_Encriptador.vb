Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class Util_Encriptador
    Sub New()

    End Sub

    Public Shared Function Encriptar(Texto As [String], Valor As [String]) As [String]
        Dim bitTexto As Byte()

        Dim objCtStream As CryptoStream
        Dim strTextoEncryptado As [String] = [String].Empty
        Dim objStream As New MemoryStream()

        Try
            bitTexto = TraerBinario(Texto, Codificacion.UTF8)
            objCtStream = New CryptoStream(objStream, ObtenerPassWord(Valor, "E"), CryptoStreamMode.Write)
            objCtStream.Write(bitTexto, 0, bitTexto.Length)
            objCtStream.FlushFinalBlock()
            strTextoEncryptado = Convert.ToBase64String(objStream.ToArray())
            objStream.Close()
            objCtStream.Close()
        Catch ex As Exception
            Return ex.Message
        End Try

        Return strTextoEncryptado
    End Function

    Public Shared Function Desencriptar(Texto As [String], Valor As [String]) As [String]
        Dim bitTextoBase64 As [Byte]()
        Dim bitTexto As [Byte]()

        Dim objCtStream As CryptoStream
        Dim strTextoOriginal As [String] = [String].Empty

        Try
            bitTextoBase64 = Convert.FromBase64String(Texto)
            Dim objStream As New MemoryStream(bitTextoBase64)
            objCtStream = New CryptoStream(objStream, ObtenerPassWord(Valor, "D"), CryptoStreamMode.Read)
            bitTexto = New [Byte](bitTextoBase64.Length - 1) {}
            strTextoOriginal = Encoding.UTF8.GetString(bitTexto, 0, Convert.ToInt32(objCtStream.Read(bitTexto, 0, bitTexto.Length)))
            objStream.Close()
            objCtStream.Close()
        Catch ex As Exception
            Return ex.Message
        End Try

        Return strTextoOriginal
    End Function

    Public Enum Codificacion
        ASCII = 1
        UTF8 = 2
        UNICODE = 3
    End Enum

    Private Shared Function ObtenerPassWord(Valor As [String], TipoCrifrado As [String]) As ICryptoTransform
        Dim bitOriginal As [Byte]()
        Dim objPassBytes As PasswordDeriveBytes
        bitOriginal = TraerBinario("MiEmpresaEsLaMejor", Util_Encriptador.Codificacion.ASCII)
        Dim objSymmetricClave As New RijndaelManaged()
        objSymmetricClave.Mode = CipherMode.CBC
        objPassBytes = New PasswordDeriveBytes(TraerBinario(GenerarCodigoHash(Util_Encriptador.ObtenerGUID(GetType(Util_Encriptador).Assembly).Value.ToString()), Codificacion.ASCII), TraerBinario(Valor, Codificacion.ASCII), "SHA1", 1)

        If TipoCrifrado = "E" Then
#Disable Warning BC40000 ' 'Public Overrides Function GetBytes(cb As Integer) As Byte()' está obsoleto: 'Rfc2898DeriveBytes replaces PasswordDeriveBytes for deriving key material from a password and is preferred in new applications.'.
            Return objSymmetricClave.CreateEncryptor(objPassBytes.GetBytes(32), bitOriginal)
#Enable Warning BC40000 ' 'Public Overrides Function GetBytes(cb As Integer) As Byte()' está obsoleto: 'Rfc2898DeriveBytes replaces PasswordDeriveBytes for deriving key material from a password and is preferred in new applications.'.
        ElseIf TipoCrifrado = "D" Then
#Disable Warning BC40000 ' 'Public Overrides Function GetBytes(cb As Integer) As Byte()' está obsoleto: 'Rfc2898DeriveBytes replaces PasswordDeriveBytes for deriving key material from a password and is preferred in new applications.'.
            Return objSymmetricClave.CreateDecryptor(objPassBytes.GetBytes(32), bitOriginal)
#Enable Warning BC40000 ' 'Public Overrides Function GetBytes(cb As Integer) As Byte()' está obsoleto: 'Rfc2898DeriveBytes replaces PasswordDeriveBytes for deriving key material from a password and is preferred in new applications.'.
        Else
            Return Nothing
        End If
    End Function

    Private Shared Function TraerBinario(Texto As [String], Tipo As Util_Encriptador.Codificacion) As [Byte]()
        If Tipo = Util_Encriptador.Codificacion.ASCII Then
            Return Encoding.ASCII.GetBytes(Texto)
        ElseIf Tipo = Util_Encriptador.Codificacion.UTF8 Then
            Return Encoding.UTF8.GetBytes(Texto)
        ElseIf Tipo = Util_Encriptador.Codificacion.UNICODE Then
            Return Encoding.Unicode.GetBytes(Texto)
        Else
            Return Nothing
        End If
    End Function

    Private Shared Function ObtenerGUID(Ensamblado As Assembly) As GuidAttribute
        Dim objGuid As [Object] = Nothing
        Dim objAtributos As Attribute() = Attribute.GetCustomAttributes(Ensamblado)
        For intContador As Integer = 0 To objAtributos.Length - 1
            objGuid = objAtributos(intContador)
            If objGuid.[GetType]().FullName.Equals("System.Runtime.InteropServices.GuidAttribute") Then
                Return DirectCast(objGuid, GuidAttribute)
            End If
        Next

        Return Nothing
    End Function

    Private Shared Function GenerarCodigoHash(Cadena As [String]) As String
        Dim bitResultado As [Byte]()
        Dim bitUnicodeTexto As [Byte]()
        Dim objEncoder As Encoder
        Dim objSB As StringBuilder
        Dim objCrypto As MD5 = New MD5CryptoServiceProvider()
        objEncoder = System.Text.Encoding.Unicode.GetEncoder()
        objCrypto = New MD5CryptoServiceProvider()
        bitUnicodeTexto = New [Byte](Cadena.Length * 2 - 1) {}
        objEncoder.GetBytes(Cadena.ToCharArray(), 0, Cadena.Length, bitUnicodeTexto, 0, True)
        bitResultado = objCrypto.ComputeHash(bitUnicodeTexto)
        objSB = New StringBuilder()
        For i As Integer = 0 To bitResultado.Length - 1
            objSB.Append(bitResultado(i).ToString("X2"))
        Next

        Return objSB.ToString()
    End Function

End Class

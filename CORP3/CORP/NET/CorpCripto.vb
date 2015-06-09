Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text

Namespace CORP3.NET
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure CorpCripto
        Private aCorpCrip As String
        Private Shared des As TripleDESCryptoServiceProvider
        Private Shared k As Byte()
        Private Shared v As Byte()
        Shared Sub New()
            CorpCripto.des = New TripleDESCryptoServiceProvider
            CorpCripto.k = Encoding.Unicode.GetBytes("etujwxrr")
            CorpCripto.v = Encoding.Unicode.GetBytes("26zzgg4t")
        End Sub

        Public Shared Function EncryptString(ByVal encryptValue As String) As String
            Dim bytes As Byte() = Encoding.Unicode.GetBytes(encryptValue)
            Dim transform As ICryptoTransform = CorpCripto.des.CreateEncryptor(CorpCripto.k, CorpCripto.v)
            Dim stream2 As New MemoryStream
            Dim stream As New CryptoStream(stream2, transform, CryptoStreamMode.Write)
            stream.Write(bytes, 0, bytes.Length)
            stream.FlushFinalBlock()
            Dim inArray As Byte() = stream2.ToArray
            stream.Close()
            Return CorpCripto.ToHex(Convert.ToBase64String(inArray))
        End Function

        Public Shared Function DecryptString(ByVal encryptedValue As String) As String
            Dim buffer2 As Byte() = Convert.FromBase64String(CorpCripto.HexToString(encryptedValue))
            Dim transform As ICryptoTransform = CorpCripto.des.CreateDecryptor(CorpCripto.k, CorpCripto.v)
            Dim stream2 As New MemoryStream
            Dim stream As New CryptoStream(stream2, transform, CryptoStreamMode.Write)
            stream.Write(buffer2, 0, buffer2.Length)
            stream.FlushFinalBlock()
            Dim bytes As Byte() = stream2.ToArray
            stream.Close()
            Return Encoding.Unicode.GetString(bytes)
        End Function

        Public Shared Function ToHex(ByVal byteArray As String) As String
            Dim str As String = ""
            Dim num As Byte
            For Each num In Encoding.ASCII.GetBytes(byteArray)
                str = (str & Conversion.Hex(num))
            Next
            Return str
        End Function

        Public Shared Function HexToString(ByVal hexString As String) As String
            Dim buffer As Byte() = New Byte((CInt(Math.Round(CDbl((CDbl(hexString.Length) / 2)))) + 1) - 1) {}
            Dim str2 As String = ""
            Dim length As Integer = hexString.Length
            Dim i As Integer = 1
            Do While (i <= length)
                str2 = (str2 & Conversions.ToString(Strings.Chr(Convert.ToByte(Strings.Mid(hexString, i, 2), &H10))))
                i = (i + 2)
            Loop
            Return str2
        End Function
    End Structure
End Namespace


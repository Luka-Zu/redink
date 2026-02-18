Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class SecureTokenStorage
    Private Const Entropy As String = "RedInkAuthEntropy"
    Private ReadOnly TokenFilePath As String

    Public Sub New()
        ' Save to the user's local application data folder
        Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim redInkFolder = Path.Combine(appData, "RedInk")
        If Not Directory.Exists(redInkFolder) Then Directory.CreateDirectory(redInkFolder)
        TokenFilePath = Path.Combine(redInkFolder, "rt.dat")
    End Sub

    Public Sub SaveRefreshToken(token As String)
        If String.IsNullOrEmpty(token) Then Return
        Dim tokenBytes = Encoding.UTF8.GetBytes(token)
        Dim entropyBytes = Encoding.UTF8.GetBytes(Entropy)
        Dim encryptedBytes = ProtectedData.Protect(tokenBytes, entropyBytes, DataProtectionScope.CurrentUser)
        File.WriteAllBytes(TokenFilePath, encryptedBytes)
    End Sub

    Public Function LoadRefreshToken() As String
        If Not File.Exists(TokenFilePath) Then Return Nothing
        Try
            Dim encryptedBytes = File.ReadAllBytes(TokenFilePath)
            Dim entropyBytes = Encoding.UTF8.GetBytes(Entropy)
            Dim decryptedBytes = ProtectedData.Unprotect(encryptedBytes, entropyBytes, DataProtectionScope.CurrentUser)
            Return Encoding.UTF8.GetString(decryptedBytes)
        Catch ex As Exception
            ' Decryption failed (e.g., different user or corrupted file)
            Return Nothing
        End Try
    End Function

    Public Sub ClearToken()
        If File.Exists(TokenFilePath) Then File.Delete(TokenFilePath)
    End Sub
End Class
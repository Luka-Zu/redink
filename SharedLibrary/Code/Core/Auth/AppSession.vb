Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports System.Threading.Tasks

Namespace SharedLibrary

    Public Module AppSession
        ' Replace with your actual API base URL
        Public ReadOnly Auth As New AuthManager("https://dev-backend.logiks.app/")

        ' The Gatekeeper method
        Public Async Function EnsureAuthenticatedAsync() As Task(Of Boolean)
            ' If we already have an access token in memory, allow them through
            If Auth.IsAuthenticated Then Return True

            ' Otherwise, prompt the user with the modal login popup
            Using loginDialog As New LoginForm()
                ' ShowDialog blocks Word until the user logs in or closes the window
                Dim result = loginDialog.ShowDialog()
                Return (result = DialogResult.OK)
            End Using
        End Function
    End Module

End Namespace
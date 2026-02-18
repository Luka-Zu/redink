Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json ' Using Newtonsoft.Json as listed in your dependencies

Public Class AuthManager
    Private ReadOnly _client As HttpClient
    Private ReadOnly _handler As HttpClientHandler
    Private _cookieContainer As CookieContainer
    Private ReadOnly _storage As SecureTokenStorage

    Private _accessToken As String
    Private ReadOnly _baseUri As Uri

    Public Sub New(baseUrl As String)
        _baseUri = New Uri(baseUrl)
        _storage = New SecureTokenStorage()
        _cookieContainer = New CookieContainer()

        ' Set up handler with CookieContainer to capture HTTP-Only cookies automatically
        _handler = New HttpClientHandler() With {
            .CookieContainer = _cookieContainer,
            .UseCookies = True
        }

        _client = New HttpClient(_handler)
        _client.BaseAddress = _baseUri

        ' Attempt to load a saved refresh cookie into the container on startup
        RestoreRefreshCookie()
    End Sub

    Public ReadOnly Property IsAuthenticated As Boolean
        Get
            Return Not String.IsNullOrEmpty(_accessToken)
        End Get
    End Property

    ' --- CORE AUTHENTICATION METHODS ---

    Public Async Function LoginAsync(email As String, password As String) As Task(Of String)
        Dim payload = New With {.email = email, .password = password}
        Dim content = New StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json")

        Dim response = Await _client.PostAsync("/api/login", content)
        Dim responseString = Await response.Content.ReadAsStringAsync()

        If Not response.IsSuccessStatusCode Then
            Throw New Exception($"Login failed: {response.StatusCode}")
        End If

        ' Returns the JSON string containing requires_2fa, method, and temp_token
        Return responseString
    End Function

    Public Async Function Verify2FaAsync(tempToken As String, code As String) As Task(Of Boolean)
        Dim payload = New With {.temp_token = tempToken, .code = code}
        Dim content = New StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json")

        Dim response = Await _client.PostAsync("/api/login/verify-2fa", content)
        If response.IsSuccessStatusCode Then
            Dim responseString = Await response.Content.ReadAsStringAsync()
            Dim result = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(responseString)

            _accessToken = result("access_token").ToString()
            SaveRefreshCookie() ' Extract and save the newly acquired refresh cookie
            Return True
        End If
        Return False
    End Function

    Public Async Function LogoutAsync() As Task
        If IsAuthenticated Then
            _client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _accessToken)
            Await _client.PostAsync("/api/logout", Nothing)
        End If

        _accessToken = Nothing
        _storage.ClearToken()
        _cookieContainer = New CookieContainer()
        _handler.CookieContainer = _cookieContainer
    End Function

    ' --- 401 RETRY LOGIC & REFRESH ---

    Public Async Function SendAuthenticatedRequestAsync(request As HttpRequestMessage) As Task(Of HttpResponseMessage)
        ' Ensure we attach the current access token
        If Not String.IsNullOrEmpty(_accessToken) Then
            request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", _accessToken)
        End If

        Dim response = Await _client.SendAsync(request)

        ' If unauthorized, try to refresh the token and retry once
        If response.StatusCode = HttpStatusCode.Unauthorized Then
            Dim refreshSuccess = Await RefreshTokenAsync()

            If refreshSuccess Then
                ' Clone the request because HttpRequestMessage cannot be sent twice
                Dim clonedRequest = Await CloneRequestAsync(request)
                clonedRequest.Headers.Authorization = New AuthenticationHeaderValue("Bearer", _accessToken)
                response = Await _client.SendAsync(clonedRequest)
            Else
                ' Refresh failed (cookie expired or invalid). Force logout.
                Await LogoutAsync()
                Throw New UnauthorizedAccessException("Session expired. Please log in again.")
            End If
        End If

        Return response
    End Function

    Public Sub SetDirectLoginToken(accessToken As String)
        _accessToken = accessToken
        SaveRefreshCookie()
    End Sub

    Private Async Function RefreshTokenAsync() As Task(Of Boolean)
        Dim response = Await _client.PostAsync("/api/refresh", Nothing) ' CookieContainer sends the refresh cookie automatically

        If response.IsSuccessStatusCode Then
            Dim responseString = Await response.Content.ReadAsStringAsync()
            Dim result = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(responseString)

            _accessToken = result("access_token").ToString()
            SaveRefreshCookie() ' Save the rotated refresh cookie
            Return True
        End If
        Return False
    End Function

    ' --- COOKIE MANAGEMENT HELPERS ---

    Private Sub SaveRefreshCookie()
        ' Extract the specific refresh cookie from the container
        Dim cookies = _cookieContainer.GetCookies(_baseUri)
        Dim refreshTokenCookie = cookies("refresh_token")

        If refreshTokenCookie IsNot Nothing Then
            _storage.SaveRefreshToken(refreshTokenCookie.Value)
        End If
    End Sub

    Private Sub RestoreRefreshCookie()
        Dim savedToken = _storage.LoadRefreshToken()
        If Not String.IsNullOrEmpty(savedToken) Then
            ' Inject it back into the container so /api/refresh can use it
            Dim cookie = New Cookie("refresh_token", savedToken) With {
                .Domain = _baseUri.Host,
                .HttpOnly = True,
                .Secure = True
            }
            _cookieContainer.Add(cookie)
        End If
    End Sub

    Private Async Function CloneRequestAsync(req As HttpRequestMessage) As Task(Of HttpRequestMessage)
        Dim clone = New HttpRequestMessage(req.Method, req.RequestUri)
        If req.Content IsNot Nothing Then
            Dim ms = New IO.MemoryStream()
            Await req.Content.CopyToAsync(ms)
            ms.Position = 0
            clone.Content = New StreamContent(ms)
            For Each header In req.Content.Headers
                clone.Content.Headers.Add(header.Key, header.Value)
            Next
        End If
        clone.Version = req.Version
        For Each prop In req.Properties
            clone.Properties.Add(prop)
        Next
        For Each header In req.Headers
            clone.Headers.TryAddWithoutValidation(header.Key, header.Value)
        Next
        Return clone
    End Function
End Class
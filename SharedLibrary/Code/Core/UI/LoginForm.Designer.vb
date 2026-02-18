' =============================================================================
' File: LoginForm.vb
' Purpose: Implements a modal Windows Forms login dialog built programmatically.
'          Handles initial Email/Password login and subsequent 2FA verification.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports System.Drawing
Imports Newtonsoft.Json
Imports System.Threading.Tasks

Namespace SharedLibrary
    Public Class LoginForm
        Inherits Form

        ' --- UI Elements ---
        Private WithEvents lblHeader As Label
        Private WithEvents lblEmail As Label
        Private WithEvents txtEmail As TextBox
        Private WithEvents lblPassword As Label
        Private WithEvents txtPassword As TextBox
        Private WithEvents lbl2FA As Label
        Private WithEvents txt2FA As TextBox
        Private WithEvents lblStatus As Label
        Private WithEvents btnLogin As Button

        ' --- State Variables ---
        Private _requires2FA As Boolean = False
        Private _tempToken As String = ""

        Public Sub New()
            ' --- Auto-scale for DPI and font ---
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = AutoScaleMode.Font

            ' --- Form properties ---
            Me.ClientSize = New Size(350, 320)
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Text = "Sign In to Red Ink"

            Dim standardFont As New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Dim headerFont As New Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)

            ' --- Header label ---
            lblHeader = New Label() With {
                .Text = "Welcome to Red Ink",
                .AutoSize = True,
                .Font = headerFont,
                .Location = New Point(20, 20)
            }
            Me.Controls.Add(lblHeader)

            ' --- Email ---
            lblEmail = New Label() With {
                .Text = "Email:",
                .AutoSize = True,
                .Font = standardFont,
                .Location = New Point(20, 60)
            }
            Me.Controls.Add(lblEmail)

            txtEmail = New TextBox() With {
                .Font = standardFont,
                .Location = New Point(20, 80),
                .Width = Me.ClientSize.Width - 40
            }
            Me.Controls.Add(txtEmail)

            ' --- Password ---
            lblPassword = New Label() With {
                .Text = "Password:",
                .AutoSize = True,
                .Font = standardFont,
                .Location = New Point(20, 115)
            }
            Me.Controls.Add(lblPassword)

            txtPassword = New TextBox() With {
                .Font = standardFont,
                .Location = New Point(20, 135),
                .Width = Me.ClientSize.Width - 40,
                .UseSystemPasswordChar = True
            }
            Me.Controls.Add(txtPassword)

            ' --- 2FA (Hidden Initially) ---
            lbl2FA = New Label() With {
                .Text = "2FA Code:",
                .AutoSize = True,
                .Font = standardFont,
                .Location = New Point(20, 170),
                .Visible = False
            }
            Me.Controls.Add(lbl2FA)

            txt2FA = New TextBox() With {
                .Font = standardFont,
                .Location = New Point(20, 190),
                .Width = Me.ClientSize.Width - 40,
                .Visible = False
            }
            Me.Controls.Add(txt2FA)

            ' --- Status Label ---
            lblStatus = New Label() With {
                .Text = "Please enter your credentials.",
                .AutoSize = False,
                .Font = standardFont,
                .ForeColor = Color.DarkGray,
                .Location = New Point(20, 230),
                .Size = New Size(Me.ClientSize.Width - 40, 40)
            }
            Me.Controls.Add(lblStatus)

            ' --- Login Button ---
            btnLogin = New Button() With {
                .Text = "Sign In",
                .Font = standardFont,
                .Location = New Point(Me.ClientSize.Width - 100, 275),
                .Size = New Size(80, 30)
            }
            AddHandler btnLogin.Click, AddressOf btnLogin_Click
            Me.Controls.Add(btnLogin)

        End Sub

        ''' <summary>
        ''' Handles the login and 2FA process when the button is clicked.
        ''' </summary>
        Private Async Sub btnLogin_Click(sender As Object, e As EventArgs)
            lblStatus.ForeColor = Color.Black
            lblStatus.Text = "Authenticating..."
            btnLogin.Enabled = False

            Try
                If Not _requires2FA Then
                    ' --- STEP 1: INITIAL LOGIN ---
                    Dim responseStr = Await AppSession.Auth.LoginAsync(txtEmail.Text, txtPassword.Text)
                    Dim result = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(responseStr)

                    If result.ContainsKey("requires_2fa") AndAlso System.Convert.ToBoolean(result("requires_2fa")) Then
                        _requires2FA = True
                        _tempToken = result("temp_token").ToString()

                        ' Morph UI for 2FA
                        lblStatus.Text = $"Please enter the code sent via {result("method")}."
                        txtEmail.Enabled = False
                        txtPassword.Enabled = False
                        lbl2FA.Visible = True
                        txt2FA.Visible = True
                        txt2FA.Focus()
                        btnLogin.Text = "Verify Code"
                    Else
                        If result.ContainsKey("access_token") Then
                            AppSession.Auth.SetDirectLoginToken(result("access_token").ToString())
                            Me.DialogResult = DialogResult.OK
                            Me.Close()
                        Else
                            lblStatus.ForeColor = Color.Red
                            lblStatus.Text = "Error: Server did not return an access token."
                        End If
                    End If

                Else
                    ' --- STEP 2: VERIFY 2FA ---
                    Dim success = Await AppSession.Auth.Verify2FaAsync(_tempToken, txt2FA.Text)

                    If success Then
                        Me.DialogResult = DialogResult.OK
                        Me.Close()
                    Else
                        lblStatus.ForeColor = Color.Red
                        lblStatus.Text = "Invalid code. Please try again."
                    End If
                End If

            Catch ex As Exception
                lblStatus.ForeColor = Color.Red
                lblStatus.Text = "Error: " & ex.Message

                ' Reset to initial state on error
                _requires2FA = False
                txtEmail.Enabled = True
                txtPassword.Enabled = True
                lbl2FA.Visible = False
                txt2FA.Visible = False
                btnLogin.Text = "Sign In"
            Finally
                btnLogin.Enabled = True
            End Try
        End Sub
    End Class

End Namespace
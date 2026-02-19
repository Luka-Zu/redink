' =============================================================================
' File: LoginForm.vb
' Purpose: Implements a modal Windows Forms login dialog built programmatically.
'          Handles initial Email/Password login and subsequent 2FA verification.
'          Styling aligned with ProgressForm.vb using TableLayoutPanel.
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
        Private WithEvents layout As TableLayoutPanel
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
            Dim standardFont As New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Dim headerFont As New Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)

            Me.AutoScaleMode = AutoScaleMode.Font
            Me.Font = standardFont
            Me.AutoSize = True
            Me.AutoSizeMode = AutoSizeMode.GrowAndShrink
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False

            ' Use SharedMethods for consistent branding
            Me.Text = "Sign In - " & SharedMethods.AN

            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
                ' Fallback if icon fails to load
            End Try

            ' --- TableLayoutPanel Setup ---
            layout = New TableLayoutPanel() With {
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Dock = DockStyle.Fill,
                .Padding = New Padding(20),
                .ColumnCount = 1,
                .RowCount = 9
            }

            ' Configure rows to auto-size
            For i As Integer = 0 To 8
                layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Next

            ' --- 1. Header label ---
            lblHeader = New Label() With {
                .Text = "Welcome to Red Ink",
                .AutoSize = True,
                .Font = headerFont,
                .Margin = New Padding(0, 0, 0, 15) ' Bottom margin for spacing
            }

            ' --- 2. Email ---
            lblEmail = New Label() With {
                .Text = "Email:",
                .AutoSize = True,
                .Margin = New Padding(0, 5, 0, 2)
            }
            txtEmail = New TextBox() With {
                .MinimumSize = New Size(300, 0),
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 10)
            }

            ' --- 3. Password ---
            lblPassword = New Label() With {
                .Text = "Password:",
                .AutoSize = True,
                .Margin = New Padding(0, 5, 0, 2)
            }
            txtPassword = New TextBox() With {
                .Dock = DockStyle.Fill,
                .UseSystemPasswordChar = True,
                .Margin = New Padding(0, 0, 0, 10)
            }

            ' --- 4. 2FA (Hidden Initially) ---
            lbl2FA = New Label() With {
                .Text = "2FA Code:",
                .AutoSize = True,
                .Margin = New Padding(0, 5, 0, 2),
                .Visible = False
            }
            txt2FA = New TextBox() With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 10),
                .Visible = False
            }

            ' --- 5. Status Label ---
            lblStatus = New Label() With {
                .Text = "Please enter your credentials.",
                .AutoSize = True,
                .ForeColor = Color.DarkGray,
                .Margin = New Padding(0, 10, 0, 15),
                .MaximumSize = New Size(300, 0) ' Allow text wrapping if long
            }

            ' --- 6. Login Button ---
            btnLogin = New Button() With {
                .Text = "Sign In",
                .AutoSize = True,
                .MinimumSize = New Size(100, 30),
                .Anchor = AnchorStyles.Right ' Align button to the right
            }
            AddHandler btnLogin.Click, AddressOf btnLogin_Click

            ' --- Add Controls to Layout ---
            layout.Controls.Add(lblHeader, 0, 0)
            layout.Controls.Add(lblEmail, 0, 1)
            layout.Controls.Add(txtEmail, 0, 2)
            layout.Controls.Add(lblPassword, 0, 3)
            layout.Controls.Add(txtPassword, 0, 4)
            layout.Controls.Add(lbl2FA, 0, 5)
            layout.Controls.Add(txt2FA, 0, 6)
            layout.Controls.Add(lblStatus, 0, 7)
            layout.Controls.Add(btnLogin, 0, 8)

            Me.Controls.Add(layout)
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

                        ' Disable previous inputs
                        txtEmail.Enabled = False
                        txtPassword.Enabled = False

                        ' Show 2FA inputs (Form will auto-resize to fit these)
                        lbl2FA.Visible = True
                        txt2FA.Visible = True
                        txt2FA.Focus()

                        btnLogin.Text = "Verify Code"
                    Else
                        ' Login successful without 2FA
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
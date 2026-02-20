' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Win32
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary

Public Class Ribbon1

    Private Enum OfficeTheme
        Unknown
        Light
        Dark
    End Enum

    Private Sub ApplyThemeAwareMenuIcon()
        Try
            Dim theme = DetectOfficeTheme()
            Select Case theme
                Case OfficeTheme.Dark
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
                Case Else
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Medium)
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            Menu1.ShowImage = True
        End Try
    End Sub

    Private Sub Ribbon1_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ApplyThemeAwareMenuIcon()
    End Sub

    Private Function DetectOfficeTheme() As OfficeTheme
        Const registryPath As String = "Software\Microsoft\Office\16.0\Common"
        Const valueName As String = "UI Theme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(registryPath)
                If key Is Nothing Then Return OfficeTheme.Unknown

                Dim raw = key.GetValue(valueName)
                If raw Is Nothing Then Return OfficeTheme.Unknown

                Dim value As Integer
                If Integer.TryParse(raw.ToString(), value) Then
                    Select Case value
                        Case 0 ' Colorful
                            Return OfficeTheme.Light
                        Case 1, 2 ' Dark Gray, Black
                            Return OfficeTheme.Dark
                        Case 3 ' White
                            Return OfficeTheme.Light
                        Case 4 ' Use system setting -> resolve via Windows app theme
                            Return If(IsWindowsAppsLightTheme(), OfficeTheme.Light, OfficeTheme.Dark)
                    End Select
                End If
            End Using
        Catch
        End Try
        Return OfficeTheme.Unknown
    End Function

    Private Function IsWindowsAppsLightTheme() As Boolean
        Const personalizePath As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const appsUseLightTheme As String = "AppsUseLightTheme"
        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizePath)
                If key Is Nothing Then Return True
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0
                End If
            End Using
        Catch
        End Try
        Return True
    End Function

    ' --- AUTH GATEKEEPER FOR RIBBON 1 ---
    Private Async Sub ExecuteAuthenticated(action As System.Action)
        Try
            Dim isAuthenticated = Await AppSession.EnsureAuthenticatedAsync()
            If Not isAuthenticated Then Return
            Me.RI_AuthToggle.Label = "Sign Out"
            action.Invoke()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error checking authentication: " & ex.Message, "Red Ink Auth", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub RI_AuthToggle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AuthToggle.Click
        Try
            If AppSession.Auth.IsAuthenticated Then
                Dim result = System.Windows.Forms.MessageBox.Show("Are you sure you want to sign out?", "Red Ink", System.Windows.Forms.MessageBoxButtons.YesNo)
                If result = System.Windows.Forms.DialogResult.Yes Then
                    Await AppSession.Auth.LogoutAsync()
                    RI_AuthToggle.Label = "Sign In"
                    System.Windows.Forms.MessageBox.Show("You have been signed out.")
                End If
            Else
                Dim success = Await AppSession.EnsureAuthenticatedAsync()
                If success Then RI_AuthToggle.Label = "Sign Out"
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Correct")
                             End Sub)
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct2_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Correct")
                             End Sub)
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Summarize_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Summarize")
                             End Sub)
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Shorten")
                             End Sub)
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
                                 Globals.ThisAddIn.MainMenu("PrimLang")
                             End Sub)
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
                                 Globals.ThisAddIn.MainMenu("PrimLang")
                             End Sub)
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Improve")
                             End Sub)
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Freestyle_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Freestyle")
                             End Sub)
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Answers_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Answers")
                             End Sub)
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Translate")
                             End Sub)
    End Sub

    Private Sub RI_QuickTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_QuickTranslate.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_QuickTranslate_Click invoked")
                                 Globals.ThisAddIn.ShowQuickTranslate()
                             End Sub)
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
                                 Globals.ThisAddIn.ShowSettings()
                             End Sub)
    End Sub

    Private Sub RI_Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Sumup")
                             End Sub)
    End Sub

    Private Sub RI_Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup2.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Sumup2_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Sumup")
                             End Sub)
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_NoFillers_Click invoked")
                                 Globals.ThisAddIn.MainMenu("NoFillers")
                             End Sub)
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Friendly_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Friendly")
                             End Sub)
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Convincing_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Convincing")
                             End Sub)
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Clipboard_Click invoked")
                                 Globals.ThisAddIn.MainMenu("InsertClipboard")
                             End Sub)
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyMyStyle_Click invoked")
                                 Globals.ThisAddIn.MainMenu("ApplyMyStyle")
                             End Sub)
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DefineMyStyle_Click_1 invoked")
                                 Globals.ThisAddIn.DefineMyStyle()
                             End Sub)
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
                                 Globals.ThisAddIn.HelpMeInky()
                             End Sub)
    End Sub

    Private Sub RI_CompareSelected_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CompareSelected.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CompareSelected_Click invoked")
                                 Globals.ThisAddIn.CompareSelectedTextRangesOutlook()
                             End Sub)
    End Sub

End Class

Public Class Ribbon2

    Private Enum OfficeTheme
        Unknown
        Light
        Dark
    End Enum

    Public Sub ApplyThemeAwareMenuIcon()
        Try
            Dim theme = DetectOfficeTheme()
            Select Case theme
                Case OfficeTheme.Light
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Medium)
                Case Else
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            Menu1.ShowImage = True
        End Try
    End Sub

    Private Sub Ribbon2_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ApplyThemeAwareMenuIcon()
    End Sub

    Private Function DetectOfficeTheme() As OfficeTheme
        Const registryPath As String = "Software\Microsoft\Office\16.0\Common"
        Const valueName As String = "UI Theme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(registryPath)
                If key Is Nothing Then Return OfficeTheme.Unknown

                Dim raw = key.GetValue(valueName)
                If raw Is Nothing Then Return OfficeTheme.Unknown

                Dim value As Integer
                If Integer.TryParse(raw.ToString(), value) Then
                    Select Case value
                        Case 0 ' Colorful
                            Return OfficeTheme.Light
                        Case 1, 2 ' Dark Gray, Black
                            Return OfficeTheme.Dark
                        Case 3 ' White
                            Return OfficeTheme.Light
                        Case 4 ' Use system setting -> resolve via Windows app theme
                            Return If(IsWindowsAppsLightTheme(), OfficeTheme.Light, OfficeTheme.Dark)
                    End Select
                End If
            End Using
        Catch
        End Try
        Return OfficeTheme.Unknown
    End Function

    Private Function IsWindowsAppsLightTheme() As Boolean
        Const personalizePath As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const appsUseLightTheme As String = "AppsUseLightTheme"
        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizePath)
                If key Is Nothing Then Return True
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then Return v <> 0
            End Using
        Catch
        End Try
        Return True
    End Function

    ' --- AUTH GATEKEEPER FOR RIBBON 2 ---
    Private Async Sub ExecuteAuthenticated(action As System.Action)
        Try
            Dim isAuthenticated = Await AppSession.EnsureAuthenticatedAsync()
            If Not isAuthenticated Then Return
            Me.RI_AuthToggle.Label = "Sign Out"
            action.Invoke()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error checking authentication: " & ex.Message, "Red Ink Auth", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub RI_AuthToggle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AuthToggle.Click
        Try
            If AppSession.Auth.IsAuthenticated Then
                Dim result = System.Windows.Forms.MessageBox.Show("Are you sure you want to sign out?", "Red Ink", System.Windows.Forms.MessageBoxButtons.YesNo)
                If result = System.Windows.Forms.DialogResult.Yes Then
                    Await AppSession.Auth.LogoutAsync()
                    RI_AuthToggle.Label = "Sign In"
                    System.Windows.Forms.MessageBox.Show("You have been signed out.")
                End If
            Else
                Dim success = Await AppSession.EnsureAuthenticatedAsync()
                If success Then RI_AuthToggle.Label = "Sign Out"
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Correct")
                             End Sub)
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct2_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Correct")
                             End Sub)
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Summarize_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Summarize")
                             End Sub)
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Shorten")
                             End Sub)
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
                                 Globals.ThisAddIn.MainMenu("PrimLang")
                             End Sub)
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
                                 Globals.ThisAddIn.MainMenu("PrimLang")
                             End Sub)
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Improve")
                             End Sub)
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Freestyle_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Freestyle")
                             End Sub)
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Answers_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Answers")
                             End Sub)
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Translate")
                             End Sub)
    End Sub

    Private Sub RI_QuickTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_QuickTranslate.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_QuickTranslate_Click invoked")
                                 Globals.ThisAddIn.ShowQuickTranslate()
                             End Sub)
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
                                 Globals.ThisAddIn.ShowSettings()
                             End Sub)
    End Sub

    Private Sub RI_Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Sumup")
                             End Sub)
    End Sub

    Private Sub RI_Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup2.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Sumup2_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Sumup")
                             End Sub)
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_NoFillers_Click invoked")
                                 Globals.ThisAddIn.MainMenu("NoFillers")
                             End Sub)
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Friendly_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Friendly")
                             End Sub)
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Convincing_Click invoked")
                                 Globals.ThisAddIn.MainMenu("Convincing")
                             End Sub)
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Clipboard_Click invoked")
                                 Globals.ThisAddIn.MainMenu("InsertClipboard")
                             End Sub)
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyMyStyle_Click invoked")
                                 Globals.ThisAddIn.MainMenu("ApplyMyStyle")
                             End Sub)
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DefineMyStyle_Click_1 invoked")
                                 Globals.ThisAddIn.DefineMyStyle()
                             End Sub)
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
                                 Globals.ThisAddIn.HelpMeInky()
                             End Sub)
    End Sub

    Private Sub RI_CompareSelected_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CompareSelected.Click
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CompareSelected_Click invoked")
                                 Globals.ThisAddIn.CompareSelectedTextRangesOutlook()
                             End Sub)
    End Sub

End Class
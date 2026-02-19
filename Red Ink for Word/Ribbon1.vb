' Part of "Red Ink for Word"
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
            ' fall through
        End Try

        Return OfficeTheme.Unknown
    End Function

    Private Function IsWindowsAppsLightTheme() As Boolean
        Const personalizePath As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const appsUseLightTheme As String = "AppsUseLightTheme"
        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizePath)
                If key Is Nothing Then Return True ' default to light if unknown
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0 ' 1=Light, 0=Dark
                End If
            End Using
        Catch
            ' default to light on error
        End Try
        Return True
    End Function

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
                                 Globals.ThisAddIn.Correct()
                             End Sub)
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click2 invoked")
                                 Globals.ThisAddIn.Correct()
                             End Sub)
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Summarize_Click invoked")
                                 Globals.ThisAddIn.Summarize()
                             End Sub)
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
                                 Globals.ThisAddIn.Shorten()
                             End Sub)
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
                                 Globals.ThisAddIn.InLanguage1()
                             End Sub)
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
                                 Globals.ThisAddIn.InLanguage1()
                             End Sub)
    End Sub

    Public Sub RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SecLang_Click invoked")
                                 Globals.ThisAddIn.InLanguage2()
                             End Sub)
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
                                 Globals.ThisAddIn.Improve()
                             End Sub)
    End Sub

    Public Sub RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleNM_Click invoked")
                                 Globals.ThisAddIn.FreeStyleNM()
                             End Sub)
    End Sub

    Public Sub RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Anonymize_Click invoked")
                                 Globals.ThisAddIn.Anonymize()
                             End Sub)
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Chat_Click invoked")
                                 Globals.ThisAddIn.ShowChatForm()
                             End Sub)
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Chat2_Click invoked")
                                 Globals.ThisAddIn.ShowChatForm()
                             End Sub)
    End Sub

    Public Sub RI_TimeSpan_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_TimeSpan_Click invoked")
                                 Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
                             End Sub)
    End Sub

    Public Sub RI_AcceptFormat_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_AcceptFormat_Click invoked")
                                 Globals.ThisAddIn.AcceptFormatting()
                             End Sub)
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
                                 Globals.ThisAddIn.InOther()
                             End Sub)
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
                                 Globals.ThisAddIn.ShowSettings()
                             End Sub)
    End Sub

    Private Sub RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleAM_Click invoked")
                                 Globals.ThisAddIn.FreeStyleAM()
                             End Sub)
    End Sub

    Private Sub RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SwitchParty_Click invoked")
                                 Globals.ThisAddIn.SwitchParty()
                             End Sub)
    End Sub

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Regex_Click invoked")
                                 Globals.ThisAddIn.RegexSearchReplace()
                             End Sub)
    End Sub

    Private Sub RI_Import_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Import_Click invoked")
                                 Globals.ThisAddIn.ImportTextFile()
                             End Sub)
    End Sub

    Private Sub RI_Halves_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Halves_Click invoked")
                                 Globals.ThisAddIn.CompareSelectionHalves()
                             End Sub)
    End Sub

    Private Sub RI_Search_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Search_Click invoked")
                                 Globals.ThisAddIn.ContextSearch()
                             End Sub)
    End Sub

    Private Sub Easteregg_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Easteregg_Click invoked")
                                 Globals.ThisAddIn.EasterEgg()
                             End Sub)
    End Sub

    Private Sub RI_Transcriptor_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Transcriptor_Click invoked")
                                 Globals.ThisAddIn.Transcriptor()
                             End Sub)
    End Sub

    Private Sub RI_Explain_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Explain_Click invoked")
                                 Globals.ThisAddIn.Explain()
                             End Sub)
    End Sub

    Private Sub RI_SuggestTitles_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SuggestTitles_Click invoked")
                                 Globals.ThisAddIn.SuggestTitles()
                             End Sub)
    End Sub

    Private Sub RI_CreatePodcast_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CreatePodcast_Click invoked")
                                 Globals.ThisAddIn.CreatePodcast()
                             End Sub)
    End Sub

    Private Sub RI_CreateAudio_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CreateAudio_Click invoked")
                                 Globals.ThisAddIn.CreateAudio()
                             End Sub)
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_NoFillers_Click invoked")
                                 Globals.ThisAddIn.NoFillers()
                             End Sub)
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Friendly_Click invoked")
                                 Globals.ThisAddIn.Friendly()
                             End Sub)
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Convincing_Click invoked")
                                 Globals.ThisAddIn.Convincing()
                             End Sub)
    End Sub

    Private Sub RI_SpecialModel_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SpecialModel_Click invoked")
                                 Globals.ThisAddIn.SpecialModel()
                             End Sub)
    End Sub

    Private Sub RI_Anonymization_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Anonymization_Click invoked")
                                 Globals.ThisAddIn.AnonymizeSelection()
                             End Sub)
    End Sub

    Private Sub RI_InsertClipboard_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_InsertClipboard_Click invoked")
                                 Globals.ThisAddIn.InsertClipboard()
                             End Sub)
    End Sub

    Private Sub RI_BallooMergePart_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BallooMergePart_Click invoked")
                                 Globals.ThisAddIn.BalloonMerge(False, True)
                             End Sub)
    End Sub

    Private Sub RI_BalloonMergeFull_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BalloonMergeFull_Click invoked")
                                 Globals.ThisAddIn.BalloonMerge(True, True)
                             End Sub)
    End Sub

    Private Sub RI_BalloonMergePartPrompt_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BalloonMergePartPrompt_Click invoked")
                                 Globals.ThisAddIn.BalloonMerge(False, False)
                             End Sub)
    End Sub

    Private Sub RI_BalloonMergeFullPrompt_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BalloonMergeFullPrompt_Click invoked")
                                 Globals.ThisAddIn.BalloonMerge(True, False)
                             End Sub)
    End Sub

    Private Sub RI_FreestyleRepeat_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleRepeat_Click invoked")
                                 Globals.ThisAddIn.FreeStyleRepeat()
                             End Sub)
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyMyStyle_Click invoked")
                                 Globals.ThisAddIn.ApplyMyStyle()
                             End Sub)
    End Sub

    Private Sub RI_DefineMyStyle_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DefineMyStyle_Click invoked")
                                 Globals.ThisAddIn.DefineMyStyle()
                             End Sub)
    End Sub

    Private Sub RI_DocCheck_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DocCheck_Click invoked")
                                 Globals.ThisAddIn.RunDocCheck()
                             End Sub)
    End Sub

    Private Sub RI_FindClause_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FindClause_Click invoked")
                                 Globals.ThisAddIn.FindClause()
                             End Sub)
    End Sub

    Private Sub RI_AddClause_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_AddClause_Click invoked")
                                 Globals.ThisAddIn.AddClause()
                             End Sub)
    End Sub

    Private Sub RI_WebAgent_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_WebAgent_Click invoked")
                                 Globals.ThisAddIn.WebAgent()
                             End Sub)
    End Sub

    Private Sub RI_EditWebAgent_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_EditWebAgent_Click invoked")
                                 Globals.ThisAddIn.CreateModifyWebAgentScript()
                             End Sub)
    End Sub

    Private Sub RI_Markdown_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Markdown_Click invoked")
                                 Globals.ThisAddIn.ConvertMarkdownToWord()
                             End Sub)
    End Sub

    Private Sub RI_FindHidden_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FindHidden_Click invoked")
                                 Globals.ThisAddIn.FindHiddenPrompts()
                             End Sub)
    End Sub

    Private Sub RI_ContentControls_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ContentControls_Click invoked")
                                 Globals.ThisAddIn.RemoveContentControlsRespectSelection()
                             End Sub)
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
                                 Globals.ThisAddIn.HelpMeInky()
                             End Sub)
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrepareRedactions_Click invoked")
                                 Globals.ThisAddIn.PrepareRedactedPDF()
                             End Sub)
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FinalizeRedactions_Click invoked")
                                 Globals.ThisAddIn.FlattenRedactedPDF()
                             End Sub)
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CheckDocumentsII_Click invoked")
                                 Globals.ThisAddIn.CheckDocumentII()
                             End Sub)
    End Sub

    Private Sub RI_EditRedact_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_EditRedact_Click invoked")
                                 Globals.ThisAddIn.EditRedactionInstructions()
                             End Sub)
    End Sub

    Private Sub RI_Filibuster_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Filibuster_Click invoked")
                                 Globals.ThisAddIn.Filibuster()
                             End Sub)
    End Sub

    Private Sub RI_ArgueAgainst_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ArgueAgainst_Click invoked")
                                 Globals.ThisAddIn.ArgueAgainst()
                             End Sub)
    End Sub

    Private Sub RI_LiveCompare_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_LiveCompare_Click invoked")
                                 Globals.ThisAddIn.CompareActiveDocWithOtherOpenDoc()
                             End Sub)
    End Sub

    Private Sub RI_RevisionSummary_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_RevisionSummary_Click invoked")
                                 Globals.ThisAddIn.SummarizeDocumentChanges()
                             End Sub)
    End Sub

    Private Sub RI_DiscussInky_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DiscussInky_Click invoked")
                                 Globals.ThisAddIn.DiscussInky()
                             End Sub)
    End Sub

    Private Sub RI_LearnDocStyle_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_LearnDocStyle_Click invoked")
                                 Globals.ThisAddIn.ExtractParagraphStylesToJson()
                             End Sub)
    End Sub

    Private Sub RI_ApplyDocStyle_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyDocStyle_Click invoked")
                                 Globals.ThisAddIn.ApplyStyleTemplate()
                             End Sub)
    End Sub

    Private Sub RI_ConvertDocToTxt_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ConvertDocToTxt_Click invoked")
                                 Globals.ThisAddIn.ExportFileContentToText()
                             End Sub)
    End Sub

    Private Sub RI_FlattenPDF_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FlattenPDF_Click invoked")
                                 Globals.ThisAddIn.FlattenPdfToImages()
                             End Sub)
    End Sub

    Private Sub RI_Charting_Click(sender As Object, e As RibbonControlEventArgs)
        ExecuteAuthenticated(Sub()
                                 SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Charting_Click invoked")
                                 Globals.ThisAddIn.OpenExistingDrawioFileForEditing()
                             End Sub)
    End Sub

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
                If success Then
                    RI_AuthToggle.Label = "Sign Out"
                End If
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub
End Class
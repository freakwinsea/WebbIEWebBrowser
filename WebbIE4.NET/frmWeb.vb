Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms

﻿''' <summary>
''' Hosts the WebView2 control and handles detection of page navigation.
''' </summary>
Public Class frmWeb

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ' The CoreWebView2InitializationCompleted event handler will be added in the Load event
        ' to ensure the control handle has been created.
    End Sub

    Private Sub frmWeb_Load(sender As Object, e As EventArgs) Handles Me.Load
        AddHandler webMain.CoreWebView2InitializationCompleted, AddressOf webMain_CoreWebView2InitializationCompleted
        ' The initialization is triggered implicitly by setting the Source or by accessing the CoreWebView2 property.
        ' We can also explicitly trigger it. Let's ensure it starts.
        Call webMain.EnsureCoreWebView2Async(Nothing)
    End Sub

    Private Async Sub webMain_CoreWebView2InitializationCompleted(sender As Object, e As CoreWebView2InitializationCompletedEventArgs)
        If e.IsSuccess Then
            ' CoreWebView2 is ready. We can now add event handlers to it.
            AddHandler webMain.CoreWebView2.NavigationStarting, AddressOf CoreWebView2_NavigationStarting
            AddHandler webMain.CoreWebView2.NavigationCompleted, AddressOf CoreWebView2_NavigationCompleted
            AddHandler webMain.CoreWebView2.SourceChanged, AddressOf CoreWebView2_SourceChanged
            AddHandler webMain.CoreWebView2.DocumentTitleChanged, AddressOf CoreWebView2_DocumentTitleChanged
            AddHandler webMain.CoreWebView2.NewWindowRequested, AddressOf CoreWebView2_NewWindowRequested

            ' Inject the DOM parser script.
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Using scriptStream As System.IO.Stream = assembly.GetManifestResourceStream("WebbIE4.Scripts.dom-parser.js")
                Using reader As New System.IO.StreamReader(scriptStream)
                    Dim script As String = reader.ReadToEnd()
                    Await webMain.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(script)
                End Using
            End Using

            ' Inject the Transformers.js library from CDN
            Dim cdnScript = "
                const script = document.createElement('script');
                script.src = 'https://cdn.jsdelivr.net/npm/@xenova/transformers/dist/transformers.min.js';
                document.head.appendChild(script);
            "
            Await webMain.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(cdnScript)

            ' Inject the LLM handler script
            Using scriptStream As System.IO.Stream = assembly.GetManifestResourceStream("WebbIE4.Scripts.llm-handler.js")
                Using reader As New System.IO.StreamReader(scriptStream)
                    Dim script As String = reader.ReadToEnd()
                    Await webMain.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(script)
                End Using
            End Using

            ' Expose the .NET ToolHost object to JavaScript
            If modGlobals.gFrmMain IsNot Nothing AndAlso modGlobals.gFrmMain.ToolHost IsNot Nothing Then
                webMain.CoreWebView2.AddHostObjectToScript("webbieTools", modGlobals.gFrmMain.ToolHost)
            End If

        Else
            ' Handle initialization failure
            MessageBox.Show("WebView2 creation failed: " & e.InitializationException.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub CoreWebView2_NavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
        If modGlobals.gClosing Then
            Return
        End If

        Try
            ' Now reset the "we're only showing forms" mode
            gShowFormsOnly = False

            ' Update UI to show we are busy
            frmMain.btnStop.Enabled = True
            frmMain.tmrBusyAnimation.Enabled = True
            frmMain.staMain.Items.Item(0).Text = modI18N.GetText("Downloading")

            ' Update history/bookmarks
            frmMain.mnuFavoritesAdd.Enabled = True
            frmMain.AddURLToRecent()
            frmMain.mnuLinksViewlinks.Enabled = False

            gForceNavigation = False ' Reset one-time "Must navigate" flag.
            Debug.Print("NavigationStarting: " & e.Uri)
        Catch ex As Exception
            Debug.Print("Exception in NavigationStarting: " & ex.Message)
        End Try
    End Sub

    Private Sub CoreWebView2_NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If modGlobals.gClosing Then
            Return
        End If

        Debug.Print("NavigationCompleted. IsSuccess: " & e.IsSuccess)

        ' Regardless of success or failure, stop the busy animation.
        frmMain.btnStop.Enabled = False
        frmMain.tmrBusyAnimation.Enabled = False

        If e.IsSuccess Then
            ' Process the page content
            Call frmMain.ProcessAfterLoad()
        Else
            ' Handle navigation failure
            Dim errorStatus As String = e.WebErrorStatus.ToString()
            frmMain.staMain.Items.Item(0).Text = "Failed to load: " & errorStatus
            ' We might still want to show what we got
            Call frmMain.ProcessAfterLoad()
        End If

        ' Update Back/Forward buttons
        frmMain.btnBack.Enabled = webMain.CoreWebView2.CanGoBack
        frmMain.mnuNavigateBack.Enabled = webMain.CoreWebView2.CanGoBack
        frmMain.mnuNavigateForward.Enabled = webMain.CoreWebView2.CanGoForward
    End Sub

    Private Sub CoreWebView2_SourceChanged(sender As Object, e As CoreWebView2SourceChangedEventArgs)
        If modGlobals.gClosing Then
            Return
        End If
        ' Update the address bar
        If webMain IsNot Nothing AndAlso webMain.CoreWebView2 IsNot Nothing Then
            frmMain.cboAddress.Text = webMain.CoreWebView2.Source
        End If
    End Sub

    Private Sub CoreWebView2_DocumentTitleChanged(sender As Object, e As Object)
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Dim title As String = webMain.CoreWebView2.DocumentTitle
        title = "WebbIE - " & title.Replace(vbNewLine, " ")
        frmMain.Text = title
        Me.Text = title
    End Sub

    Private Sub CoreWebView2_NewWindowRequested(sender As Object, e As CoreWebView2NewWindowRequestedEventArgs)
        If Not My.Settings.AllowPopupWindows Then
            e.Handled = True ' This prevents the new window from opening
            PlayErrorSound()
            frmMain.staMain.Items.Item(1).Text = modI18N.GetText("Pop-up window blocked")
        End If
    End Sub

    Private Sub frmWeb_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            If e.CloseReason = CloseReason.UserClosing Then
                e.Cancel = True
                SwitchBackToMain()
            End If
        Catch
        End Try
    End Sub

    Private Sub SwitchBackToMain()
        Try
            Call Me.Hide()
            Call frmMain.DoDelayedRefresh()
            frmMain.Visible = True
            modGlobals.gShowingWebpage = False
        Catch
        End Try
    End Sub

    Private Sub tmrCheckForEscape_Tick(sender As Object, e As EventArgs) Handles tmrCheckForEscape.Tick
        Try
            If modGlobals.gClosing Then
                tmrCheckForEscape.Enabled = False
            ElseIf Me.Visible Then
                If NativeMethods.GetKeyState(NativeMethods.VK_Escape) < 0 Then
                    SwitchBackToMain()
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrCheckForClosing_Tick(sender As Object, e As EventArgs) Handles tmrCheckForClosing.Tick
        If modGlobals.gClosing Then
            tmrCheckForClosing.Enabled = False
            Me.Close()
        End If
    End Sub

    Private Sub tmrCheckForNavigating_Tick(sender As Object, e As EventArgs) Handles tmrCheckForNavigating.Tick
        If modGlobals.gDesiredURL <> "" Then
            If webMain IsNot Nothing AndAlso webMain.CoreWebView2 IsNot Nothing Then
                webMain.CoreWebView2.Navigate(modGlobals.gDesiredURL)
                modGlobals.gDesiredURL = ""
            End If
        End If
    End Sub

    Private Sub tmrCheckForNavigationComplete_Tick(sender As Object, e As EventArgs) Handles tmrCheckForNavigationComplete.Tick
        ' This timer is now obsolete. The logic has been moved to NavigationCompleted.
        tmrCheckForNavigationComplete.Enabled = False
    End Sub

    Public Function URL() As String
        If webMain IsNot Nothing AndAlso webMain.CoreWebView2 IsNot Nothing Then
            Return webMain.CoreWebView2.Source
        Else
            Return ""
        End If
    End Function

End Class
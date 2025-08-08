Imports Newtonsoft.Json

﻿Option Strict On
Option Explicit On

'Compiler options:
'   MONITOR_FOCUS   If you add tmrFocus (or something) to the form and enable this then
'       you'll track what has focus, in case of the MSAA bug.

Public Class InteractableElement
    Public Property id As String
    Public Property type As String
    Public Property text As String
    Public Property href As String
End Class

Public Class frmMain
    Public WithEvents ToolHost As ToolHost

    Public Async Sub ExecuteClick(elementId As String)
        Dim clickScript = $"document.querySelector('[data-webbie-id=""{elementId}""]').click();"
        Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(clickScript)
        DoDelayedRefresh()
    End Sub

    Public Async Sub ExecuteType(elementId As String, textToType As String)
        Await SetElementValueAsync(elementId, textToType)
        DoDelayedRefresh()
    End Sub
    'WebbIE

    '.Net version of WebbIE

    'LICENCE
    'This program is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General Public License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.
    '
    '    This program is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General Public License for more details.
    '
    '    You should have received a copy of the GNU General Public License
    '    along with this program.  If not, see <http://www.gnu.or5g/licenses/>.

    '"WebbIE" is copyright 2002 Paul Blenkhorn and Gareth Evans, and is released by them under the GNU Public
    'Licence Version 3. A copy is available at http://www.gnu.org/licenses/gpl.txt
    'WebbIE includes code copyright 2002, 2005 Alasdair King, and where this is the case it is also licenced
    'under the GNU Public Licence Version 3.
    'Copyright 2007 Alasdair King.

    'Name: Alasdair King alasdair@webbie.org.uk

    '4.0 Conversion to .Net
    '       Removed:
    '           Ability to report errors.
    '           Bespoke font form
    '           Help forms.
    '           Table display.
    '           Bookmarks
    '           External browser
    '           Ability to start a dial-up connection automatically and monitor it.
    '           Display of online status
    '           Any kind of Ajax tracking.
    '           Readability
    '           Fetching PDF files from Google.
    '4.0.10 Jan 2013    Shipped as Beta
    '4.0.11 Jan 2013    Bugfixes (registry error preventing startup, -reinstall not working)
    '4.0.12 Jan 2013    Bugfixes to list of links and crop functions.
    '4.1.0 Feb 2013    Added back a way to see the web page.
    '                   Removed image save.
    '                   Reverted to just showing the browser home page.
    '4.2.0 Feb 2013     Fixed crash if homepage is "http://"
    '                   Added Favorites window on B (or Ctrl+B) that shows all your favourites in a 
    '                   treeview window. There are now three ways to get at your favourites: the menu,
    '                   the home page, and the popup window. I like the last best, but there's no reason
    '                   to drop the others as yet. I should get rid of the home page, but someone wrote
    '                   to say they love it, so I don't want to!
    '                   Made Print print the web page, not the text area.
    '                   Added a keyhook so that Escape reliably closes the frmWeb view.
    '                   Made the frmWeb view fullscreen.
    '                   Now a choice for home page: the IE home page, or the WebbIE home page (default)
    '                   Lots of work to suppress navigation errors.
    '                   Should now more accurately stop navigation sounds when page finishes loading.
    '4.3.0  26 Feb 2013 Fixed uncrop bug - crash when failing to place caret.
    '                   Made crop use MAIN element, if found.
    '                   Added input type="email" support (HTML5)
    '                   Added input type="range" support (HTML5)
    '                   Added progress element support (HTML5)
    '                   Improved labelling of controls: many more controls will have working labels,
    '                   and there will be less duplication of label text.
    '                   Fixed bug with toolbar option not hiding/showing toolbar.
    '                   Added back in the ability to number links in WebbIE (off by default)
    '4.3.1  27 Mar 2013 Made loading smoother when loading Favorites menu - better feel.
    '                   Made loading much faster.
    '                   Fixed toolbar display bug. 
    '4.3.3  7 May 2013  Put in comprehensive error handling. This means that users should never
    '                   see any error messages unless there is a comprehensive Windows system
    '                   crash. Instead, things will just not work or work oddly. This is to
    '                   increase confidence for users: ideally they get an error message which
    '                   would help them to fix it, or report it to me to get fixed. However,
    '                   this generally doesn't work as a process. 
    '                   Made "show IE homepage" the default again because of negative feedback.
    '                   Made "localhost" work as URL.
    '                   Added support for aria-label and aria-labelledby (found on Facebook)
    '4.3.4  7 June 2013 Made about:blank load faster.
    '                   Fixed textarea input so it works(!)
    '                   Made textarea not label with Name, because that's code and sucks.
    '                   Made text notruntogether on Facebook.
    '4.3.5 14 June 2013 Settings now update from one version to the next.
    '                   Defaults to Maximimised window state.
    '                   Won't check for updates more than once a day.
    '4.3.6 23 Feb 2014  Fixed a link text bug, should label links better.
    '                   Fixed a focus problem with youtube.com.
    '                   Fixed a "finished navigation" problem with youtube.com.
    '                   No longer lets you change disabled text inputs etc.
    '                   Made manufacturer "Accessible and WebbIE"
    '                   Added online activation
    '                   Updated updater DLL
    '4.4.0  2 Mar 2014  Fixed a bug with making WebbIE default browser: should now work.
    '                   Fixed font and appearance for text view, select forms.
    '                   Added colour dialog option to set foreground/background colours.
    '                   Now uses IE11 rendering.
    '                   Made most forms maximised.
    '4.4.1  6 June 2014 Fixed many input buttons not working.
    '4.5.0  22 Dec 2014 Fixed case-sensitive urls not working when typed into address bar.
    '                   Added ability to download/open VIDEO and AUDIO HTML5 elements directly in your
    '                   default media player: just hit Open. Doesn't work with embedded data URI elements.
    '                   Shortcut keys for media play: Ctrl+P play, Ctrl+O open, Space stop.
    '                   If you hold Control while clicking Refresh, or Shift while selecting Ctrl+R, 
    '                   then you get a full page refresh in the web browser. (Previously you couldn't
    '                   get a full browser refresh, just WebbIE re-displaying)
    '                   Can now open saved MHT files from File > Open.
    '                   Added TeamViewer download link.
    '4.5.1              Fixed focus going to webbrowser, not text area (e.g. if you go to google.com)
    '                   Took out display of frames that are hard to get at (security exceptions) which
    '                   will break some pages but fixes lots of navigation complexity and loads faster. 
    '                   ... And they are probably all ads anyway. 
    '                   Fixed activation DLL counting every use as an activation.
    '4.5.2              Fixed focus getting lost on Google search form when Next Page Of Results used. May 2016
    '                       (At least, I think I have fixed this, I can't test it because I'm in a hospital with 
    '                       no wi-fi.) TODO Test focus on Google Search!
    '4.5.3              Updated frame handling a bit, will now detect and refresh when an IFRAME loads.
    '                   Links which are also headings now displayed correctly.
    '4.5.4, 12 Mar 2017
    '                   Fixed a bug With the password input Not applying the password If you pressed Return instead
    '                   of clicking. 
    '                   Added a check for navigation not terminating and forcing page rendering.
    '                   Removed WebbIE update call if WebbIEUpdater.dll not found - for Windows Store.
    '                   Took out activation code, since I've never used it. 
    '5.0.0 22 Dec 2018  
    '                   Fixed web search. 


    'TODO
    '   Highlighting for headings.
    '   Spellcheck on text areas.
    '   ARIA hidden attribute (see Facebook front page)
    '   Tab through links.
    '   Escape/stop/back interrupt navigation.
    '   ARIA role main support
    '   RSS link when we have finished RSS Reader
    '   Hide addressbar and statusbar options
    '   Back button shouldn't try to go back to pages that autoforwarded (e.g. Google results)
    '
    'Little stuff:
    '   RSS button should only enable when RSS
    '

    ''' <summary>
    ''' Indicates that there should be a blank line at this point in the output. 
    ''' </summary>
    Private Const BLANK_LINE_MARKER As String = "jweofijweoifj"

#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
    'For capturing focus changes that break MSAA.
    Private focusHandler As System.Windows.Automation.AutomationFocusChangedEventHandler = Nothing
    Private mControlName As String ' The result of UIA telling me the name of the current element.
    Private mControlType As String ' the result of UIA telling me the type of the current element, e.g. "menu item"
#End If
    Public Enum IEVisible
        Toggle
        MakeIEVisible
        MakeTextVisible
    End Enum

    ''' <summary>
    ''' Indicates that the page has had links etc. stripped from it when set.
    ''' </summary>
    ''' <remarks></remarks>
    Private mCropped As Boolean

    'Parsing State Variables
    Private mOutput As System.Text.StringBuilder
    Private Const NUMBER_CHARS_AFTER_LINK_PERMITTED As Integer = 500 ' the number of chars permitted

    Private mForceFollowAddress As Boolean ' If set, you must follow the href for a link, not do a click.
    Private mForceDownloadLink As Boolean ' if set, you must download the file this link points at, not do a click
    Private ReadOnly mProcessNextDocumentComplete As Boolean ' exception to the general "process on _onload" rule from 3.11.
    '   If set, we have an internal navigation, so we should start processing on _DocumentComplete instead.
    Private mJustDidLink As Boolean ' indicates that we've just done a link, so can add a very
    'limited amount of text after a link before a newline is applied.
    Private mNewline As String ' whether we've just done a newline in the page
    Private mInArticle As Integer ' Whether we're in an article section. Not Boolean because they nest.

    Private mGoingBack As Boolean ' used to indicate we're going back
    Private mobjNavigationRecord As System.Collections.Generic.Dictionary(Of String, Integer) ' stores line numbers of where we've been before
    Private mblnErrorPage As Boolean ' whether there has been an error detected by modGlobals.gWebHost.webMain
    Private mHeadingLevel As String ' the tags for the heading level found
    'on the page: we check first that we can find a heading, and if
    'we can, set the heading level to that.
    Private mJustDidHeading As Boolean ' indicates that we've just found the heading for the page.
    'If the following text is a link, we have to handle it a bit differently.
    Private mSeekingInternalTarget As Boolean ' indicates that we are looking for
    'an internal target on this page parse called...
    Private mInternalTarget As String
    Private mControlKeyPressed As Boolean
    Private mShiftKeyPressed As Boolean
    Private mElementWithFocus As mshtml.IHTMLElement ' For Ajax handling. Tracks the element that has suddenly
    '   got the focus in IE and should therefore now get the focus in WebbIE.
    Private mSeekingFocusElement As Boolean ' indicates we are looking for an internal target node
    '   on this page: mElementWithFocus in fact.
    Private mInternalLinkNavigationStart As Integer ' The line number we started on when we did an internal
    '   link navigation.
    Private ReadOnly mBBCIPlayer As mshtml.IHTMLElement ' the embedded BBC iPlayer Flash object.
    Private mTerminateParsing As Boolean ' if set, stop parsing and present contents of the page. This isn't
    '   very neat, but we do it because of sites like http://www.epoznan.pl/, which fails to stop parsing.
    '   It keeps adding new nodes.
    Private mInForm As Boolean ' indicates that we are in a form.

    'Refresh levels for web browser
    Private Const REFRESH_NORMAL As Integer = 0 ' Perform a lightweight refresh that does not include sending the HTTP "pragma:nocache" header to the server.
    Private Const REFRESH_IFEXPIRED As Integer = 1 ' Perform a lightweight refresh if the page has expired.
    Private Const REFRESH_COMPLETELY As Integer = 3 ' Perform a full refresh that includes sending a "pragma:nocache" header to the server (HTTP URLs only).

    Private mFormClosing As Boolean = False

    Public Sub DoBack()
        On Error GoTo cantGoBack
        Dim lineNumber As Integer
        If btnBack.Enabled Then
            If mInternalLinkNavigationStart > -1 Then
                'aha, we're trying to do back on an internal link
                Call modAPIFunctions.SetCurrentLineIndex(txtText, mInternalLinkNavigationStart)
                mInternalLinkNavigationStart = -1
            Else
                'go back
                mGoingBack = True
                Call modAccesskeys.ClearAccessKeys()
                'Get the current line before we navigate - used
                'for back and forward. Used to be in BeforeNavigate, but
                'gets too confusing with all the frames and adverts triggering
                'BeforeNavigate
                'record the line number
                lineNumber = GetCurrentLineIndex(txtText)
                'we often get several  events for a page while it
                'loads (e.g. adverts) and we only want to record the line number when
                'there's actually something on the page, e.g. lineNumber > 1
                If lineNumber > 1 Then
                    If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.URL) Then
                        mobjNavigationRecord.Remove(modGlobals.gWebHost.URL)
                    End If
                    mobjNavigationRecord.Add(modGlobals.gWebHost.URL, lineNumber)
                End If
                'Start navigating timer.
                Me.tmrBusyAnimation.Enabled = True
                If modGlobals.gWebHost.webMain IsNot Nothing AndAlso modGlobals.gWebHost.webMain.CoreWebView2 IsNot Nothing Then
                    modGlobals.gWebHost.webMain.CoreWebView2.GoBack()
                End If
            End If
        Else
            Call PlayErrorSound()
        End If
        Exit Sub
cantGoBack:
        'can't go backwards, don't do anything
        Call StopBrowsers()
    End Sub

    Public Sub DoForward()
        On Error GoTo cantGoForward
        Dim lineNumber As Integer
        'Call PlayNavigationStartSound
        If mnuNavigateForward.Enabled Then
            mGoingBack = True
            'get the current line before we navigate - used
            'for back and forward. Used to be in BeforeNavigate, but
            'gets too confusing with all the frames and adverts triggering
            'BeforeNavigate
            'record the line number
            lineNumber = GetCurrentLineIndex(txtText)
            'we often get several  events for a page while it
            'loads (e.g. adverts) and we only want to record the line number when
            'there's actually something on the page, e.g. lineNumber > 1
            If lineNumber > 1 Then
                mobjNavigationRecord.Add(modGlobals.gWebHost.URL, lineNumber)
            End If

            If modGlobals.gWebHost.webMain IsNot Nothing AndAlso modGlobals.gWebHost.webMain.CoreWebView2 IsNot Nothing Then
                modGlobals.gWebHost.webMain.CoreWebView2.GoForward()
            End If
        Else
            Call PlayErrorSound()
        End If
        Exit Sub
cantGoForward:
        'can't go forwards, don't do anything
        Call StopBrowsers()
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        On Error Resume Next
        Call DoBack()
    End Sub

    Private Sub btnHome_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHome.Click
        On Error Resume Next
        Call DoHome()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        On Error Resume Next
        If modGlobals.gWebHost.webMain IsNot Nothing AndAlso modGlobals.gWebHost.webMain.CoreWebView2 IsNot Nothing Then
            modGlobals.gWebHost.webMain.CoreWebView2.Reload()
        End If
    End Sub

    Private Sub btnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStop.Click
        On Error Resume Next
        Call StopBrowsers()
    End Sub

    Private Sub frmMain_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        On Error Resume Next
    End Sub

    Public Sub mnuLinksDownloadlink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksDownloadlinktarget.Click
        On Error Resume Next
        mForceDownloadLink = True
        Call UserPressedReturn()
        mForceDownloadLink = False
    End Sub


    Private Sub frmMain_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            'Detect CAPSLOCK is down, and if so don't do any key handling, since it's the standard screenreader control key.
            If modKeys.IsCapslockPressed Then
                Exit Sub
            End If

            Dim KeyCode As Integer = eventArgs.KeyCode
            'process user keypresses by calling the appropriate GUI component
            Dim ctrlDown As Boolean 'control is pressed
            Dim quickKeysOn As Boolean
            Dim handled As Boolean = True
            If Me.ActiveControl Is Nothing Then
                quickKeysOn = False
            ElseIf eventArgs.Alt Then
                ' If user is holding down Alt then don't do single-key action. Allows access keys to operate.
                quickKeysOn = False
            ElseIf Me.ActiveControl.Name.Contains("mnu") Then
                ' In a menu item, don't do single-key action.
                quickKeysOn = False
            Else
                quickKeysOn = (My.Settings.QuickKeys And (Me.ActiveControl.Name = Me.txtText.Name))
            End If
            ctrlDown = eventArgs.Control Or quickKeysOn
            Select Case KeyCode
                Case Is = System.Windows.Forms.Keys.Escape ' Stop loading
                    If btnStop.Enabled Then
                        Call StopBrowsers()
                    End If
                Case Is = System.Windows.Forms.Keys.F7
                    'goto headline
                    Call GotoHeadline(1)
                Case Is = System.Windows.Forms.Keys.D ' address bar
                    If ctrlDown Then
                        Call cboAddress.Focus()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.F ' Find
                    If ctrlDown Then
                        Call FindText()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.Back ' Go back
                    'check it isn't in the address field
                    If Me.ActiveControl Is Nothing Then
                        Call btnBack_Click(btnBack, New System.EventArgs())
                    ElseIf Me.ActiveControl.Name <> cboAddress.Name Then
                        Call btnBack_Click(btnBack, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.Right ' Go forward (with ALT)
                    If eventArgs.Alt Then
                        Call DoForward()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.L ' List links
                    If ctrlDown Then
                        Call ListLinks()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.Left ' Go back (with ALT)
                    If eventArgs.Alt Then
                        btnBack_Click(btnBack, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.F5 ' Refresh
                    btnRefresh_Click(btnRefresh, New System.EventArgs())
                    handled = False
                Case Is = System.Windows.Forms.Keys.Home ' Home (with Alt)
                    If eventArgs.Alt And ctrlDown And Not My.Settings.QuickKeys Then
                        Call SetHomepage()
                    ElseIf eventArgs.Alt Then
                        btnHome_Click(btnHome, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.H ' next headline
                    If ctrlDown Or quickKeysOn Then
                        KeyCode = 0
                        eventArgs.Handled = True
                        eventArgs.SuppressKeyPress = True
                        Call GotoHeadline(CInt(IIf(eventArgs.Shift, -1, 1)))
                    Else
                        handled = False
                    End If

                    'DEV: can't hard-code access keys - doesn't allow for I18N
                    '        Case Is = vbKeyD ' in case the access key doesn't work
                    '            If altDown Then
                    '                Call cboAddress.SetFocus
                    '                KeyCode = 0
                    '            End If
                    '        Case Is = vbKeyT 'in case the access key doesn't work
                    '            If altDown Then
                    '                Call txtText.SetFocus
                    '                KeyCode = 0
                    '            End If
                Case Is = System.Windows.Forms.Keys.W ' for Google search
                    If ctrlDown Then
                        KeyCode = 0
                        Call DoWebSearch()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.B ' For Favourites
                    If ctrlDown Then
                        KeyCode = 0
                        Call ShowFavoritesWindow()
                    Else
                        handled = False
                    End If
                    'We now have a bunch of shortcuts specifically for when are in QuickLinks
                    'mode, that is, when users can hit a single character to operate a function.
                Case Is = System.Windows.Forms.Keys.S 'for skip down
                    If quickKeysOn Then
                        Call DoSkip(SKIP_DOWN)
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.A ' select all
                    If quickKeysOn Then
                        Call mnuEditSelectall_Click(mnuEditSelectall, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.C ' copy
                    If quickKeysOn Then
                        Call mnuEditCopy_Click(mnuEditCopy, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.X ' cut
                    If quickKeysOn Then
                        Call EditCut()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.V ' paste
                    If quickKeysOn Then
                        Call Paste()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.U ' for skip up
                    If quickKeysOn Then
                        Call DoSkip(SKIP_UP)
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.K ' crop
                    If quickKeysOn Then
                        Call CropPage()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.R ' refresh
                    If quickKeysOn Or ctrlDown Then ' Need also ctrlDown to handle Ctrl+Shift+R
                        Call mnuNavigateRefresh_Click(mnuNavigateRefresh, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.E ' rss
                    If quickKeysOn Then
                        KeyCode = 0
                        Call RSS()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.F6
                    If eventArgs.Shift Then
                        Call SkipToFormElement(-1)
                    Else
                        handled = False
                    End If
                Case Else
                    handled = False
            End Select
            If handled Then
                eventArgs.Handled = True
                eventArgs.SuppressKeyPress = True
            End If
        Catch
        End Try
    End Sub

    Private Sub Paste()
        Try
            'process use clicking paste button in menu
            Dim currentControlName As String
            If Me.ActiveControl Is Nothing Then
                'Assume it's the web browser
                currentControlName = modGlobals.gWebHost.webMain.Name
            Else
                currentControlName = Me.ActiveControl.Name
            End If
            Select Case currentControlName
                Case modGlobals.gWebHost.webMain.Name
                    'Dim wb As SHDocVw.WebBrowser
                    'wb = CType(workBrowser.ActiveXInstance, SHDocVw.WebBrowser)
                    'Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_PASTE, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT)
                    Call modSendKeys.SendPaste()
                    'Call SendKeys.Send("^V")
                Case cboAddress.Name
                    cboAddress.SelectedText = My.Computer.Clipboard.GetText
                Case txtText.Name
                    Call Beep()
            End Select
        Catch
        End Try
    End Sub

    Public Sub mnuLinksFollowlinkaddress_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksFollowlinkaddress.Click
        On Error Resume Next
        mForceFollowAddress = True
        Call UserPressedReturn()
        mForceFollowAddress = False
    End Sub

    Public Sub ResizeToolbar()
        Try
            If My.Settings.ShowToolbar Then
                Me.MainToolStrip.Visible = True
                Dim ds As System.Windows.Forms.ToolStripItemDisplayStyle = CType(IIf(My.Settings.ToolbarCaptions, ToolStripItemDisplayStyle.ImageAndText, ToolStripItemDisplayStyle.Image), ToolStripItemDisplayStyle)
                For Each tsi As ToolStripItem In Me.MainToolStrip.Items
                    tsi.DisplayStyle = ds
                Next tsi

                If My.Settings.ToolbarCaptions Then
                    Me.MainToolStrip.Height = 90
                    picBusy.Image = My.Resources.timer_done_big
                Else
                    Me.MainToolStrip.Height = 74
                    picBusy.Image = My.Resources.timer_done
                End If
            Else
                ' No toolbar at all!
                Me.MainToolStrip.Visible = False
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuViewLinkinformation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error Resume Next
        Call txtText_KeyDown(txtText, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.F8))
    End Sub

    Public Sub mnuViewRSSNewsFeed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewRSSNewsFeed.Click
        On Error Resume Next
        Call RSS()
    End Sub

    Public Sub RSS()
        On Error Resume Next
        Call frmRSS.ShowDialog(Me)
    End Sub

    ''' <summary>
    ''' Page navigation is complete (enough) so call this to render the page. 
    ''' </summary>
    Public Sub ProcessAfterLoad()
        On Error Resume Next
        Dim lineNumber As Integer
        Call ProcessPage()
        'Right, we've just finished loading and displaying a page. Now, where
        'do we put the caret? Two factors:
        ' - there is an internal target to which we should move the caret
        ' - if we've gone forwards or backwards, put the caret where it was
        'You could go f/b to a page with an internal target, but you
        'should go to where the caret was, not where the target indicates.
        'So check f/b first.
        If mobjNavigationRecord Is Nothing Then
            mobjNavigationRecord = New System.Collections.Generic.Dictionary(Of String, Integer)
        End If
        If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.URL) Then
            'found an entry for this page.
            lineNumber = mobjNavigationRecord.Item(modGlobals.gWebHost.URL)
        End If
        If lineNumber = 0 Then
            If mSeekingInternalTarget Then
                Call GotoInternalTarget()
            End If
        Else
            'okay, we have somewhere to go.
            'Debug.Print "Decided to put caret at line:" & lineNumber
            Dim newSelectionStart As Integer = GetCharacterIndexOfLine(txtText, lineNumber)
            If newSelectionStart > -1 Then
                txtText.SelectionStart = newSelectionStart
            End If
            txtText.SelectionLength = 0
            Call txtText.ScrollToCaret()
            'This call sets the focus to the text area when the WebBrowser object steals it. You can 
            'observe this by navigating to google.com. 
            'I tried many approaches to this, but this works (testing Windows 10 Feb 2016). It didn't
            'work before but I was getting the API call wrong. Use the Analyze functions!
            Call NativeMethods.SetFocus(txtText.Handle)
        End If
    End Sub

    'Private Sub tmrProcessAfterLoad_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrProcessAfterLoad.Tick
    '    If modGlobals.gClosing Then
    '        Me.tmrProcessAfterLoad.Enabled = False
    '        Exit Sub
    '    End If

    '    Static busy As Boolean
    '    Try
    '        'Fires after the Load event fires on the main webpage.

    '        Dim lineNumber As Integer
    '        tmrProcessAfterLoad.Enabled = False

    '        'Call MessageBox.Show("tmrProcessAfterLoad")

    '        If busy Then
    '        Else
    '            busy = True
    '            Call ProcessPage()

    '            'Right, we've just finished loading and displaying a page. Now, where
    '            'do we put the caret? Two factors:
    '            ' - there is an internal target to which we should move the caret
    '            ' - if we've gone forwards or backwards, put the caret where it was
    '            'You could go f/b to a page with an internal target, but you
    '            'should go to where the caret was, not where the target indicates.
    '            'So check f/b first.
    '            If mobjNavigationRecord Is Nothing Then
    '                mobjNavigationRecord = New System.Collections.Generic.Dictionary(Of String, Integer)
    '            End If
    '            If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.webMain.Url.ToString()) Then
    '                'found an entry for this page.
    '                lineNumber = mobjNavigationRecord.Item(modGlobals.gWebHost.webMain.Url.ToString())
    '            End If
    '            If lineNumber = 0 Then
    '                If mSeekingInternalTarget Then
    '                    Call GotoInternalTarget()
    '                End If
    '            Else
    '                'okay, we have somewhere to go.
    '                'Debug.Print "Decided to put caret at line:" & lineNumber
    '                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
    '                txtText.SelectionLength = 0
    '                Call txtText.ScrollToCaret()
    '                'This call sets the focus to the text area when the WebBrowser object steals it. You can 
    '                'observe this by navigating to google.com. 
    '                'I tried many approaches to this, but this works (testing Windows 10 Feb 2016). It didn't
    '                'work before but I was getting the API call wrong. Use the Analyze functions!
    '                Call NativeMethods.SetFocus(txtText.Handle)
    '            End If
    '            busy = False
    '        End If
    '    Catch
    '        busy = False
    '    End Try
    'End Sub

    Private Sub GotoInternalTarget()
        Try
            Dim selStart As Integer = txtText.Text.IndexOf(TARGET_MARKER)
            If selStart > -1 Then
                'In theory I should be able to select the TARGET_MARKET text and then
                'delete it (txtText.SelectedText = "") but since this magically does
                'not work for no reason I can see, do with strings.
                Dim contents As String = txtText.Text.Substring(0, selStart) + txtText.Text.Substring(selStart + TARGET_MARKER.Length, txtText.Text.Length - selStart - TARGET_MARKER.Length)
                Call SetText(contents)
                txtText.SelectionStart = selStart
            End If
            mSeekingInternalTarget = False
        Catch
        End Try
    End Sub

    Private Sub tmrRefreshIfNotChange_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrRefreshIfNotChange.Tick
        If modGlobals.gClosing Then
            Me.tmrRefreshIfNotChange.Enabled = False
            Exit Sub
        End If
        Try
            tmrRefreshIfNotChange.Enabled = False
            If gExiting Then
            ElseIf gNoPageActionSoRefresh Then
                gNoPageActionSoRefresh = False
                Debug.Print("Triggered refresh!")
                Call RefreshCurrentPage()
                Call PlayDoneSound()
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrSetFocus_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrSetFocus.Tick
        On Error Resume Next
        tmrSetFocus.Enabled = False
        If modGlobals.gClosing Then
            Exit Sub
        End If
        'log "tmrSetFocus_Timer"
        If modGlobals.gFrmWebHasFocusAndItShouldNot Then
            modGlobals.gFrmWebHasFocusAndItShouldNot = False
            Call Me.Focus()
            tmrSetFocus.Enabled = True ' let's let it come back round...
        Else
            Call txtText.Focus()
        End If
    End Sub

    Private Sub mnuHelpWebbiehome_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpWebbIEOrg.Click
        On Error Resume Next
        Call GotoWebbIEDotOrgDotUK()
    End Sub

    Public Sub GotoWebbIEDotOrgDotUK()
        On Error Resume Next
        'go to home for webbie - website
        Call StartNavigating(modI18N.GetText("http://www.webbie.org.uk"))
    End Sub

    ''' <summary>
    ''' Move cursor to the next headline
    ''' </summary>
    ''' <param name="direction"></param>
    ''' <remarks></remarks>
    Private Sub GotoHeadline(ByVal direction As Integer)
        On Error Resume Next
        If GotoLineStarting(SECTION_MARKER_COMMON, direction, False) Then
            'Don't select - stops screenreaders reading line, they read the document instead.
            'Call SelectCurrentParagraph
            'okay, this will have selected the paragraph starting "Page Headline:". But
            'what if the headline is the next line as well? Happens with links.
            'Well, then, select the next line
            '        If Len(txtText.SelText) < Len(SECTION_MARKER_H1) + 10 Then
            '            nextLine = txtText.selStart + txtText.SelLength + Len(vbNewLine)
            '            nextLine = InStr(nextLine, txtText.Text, vbNewLine)
            '            If nextLine > 0 Then
            '                'found new end-of-line
            '                txtText.SelLength = nextLine - txtText.selStart
            '            End If
            '        End If
            Call txtText.ScrollToCaret()
        End If
    End Sub

    ''' <summary>
    ''' Moves cursor to line beginning with startsWith. Returns true if found.
    ''' </summary>
    ''' <param name="startsWith"></param>
    ''' <param name="direction"></param>
    ''' <param name="wrap"></param>
    ''' <returns>true if line starting with startsWith found</returns>
    ''' <remarks></remarks>
    Private Function GotoLineStarting(ByRef startsWith As String, Optional ByRef direction As Integer = 1, Optional ByRef wrap As Boolean = True) As Boolean
        Try
            Dim numberLines As Integer
            Dim i As Integer
            Dim lineContent As String
            Dim found As Boolean
            Dim currentLine As Integer
            Dim LinkLength As Integer = ID_LINK.Length

            numberLines = GetNumberOfLines(txtText)
            currentLine = GetCurrentLineIndex(txtText)
            Debug.Print("Start: " & currentLine)
            i = currentLine + direction
            'Debug.Print "Line: " & i
            While i >= 0 And i < numberLines And (Not found)
                lineContent = GetLine(txtText, i)
                If InStr(1, lineContent, startsWith, CompareMethod.Text) = 1 Or
                   InStr(1, lineContent, startsWith, CompareMethod.Text) = 1 + LinkLength + 2 Then
                    'found it
                    Debug.Print("Found")
                    found = True
                    txtText.SelectionStart = GetCharacterIndexOfLine(txtText, i)
                    txtText.SelectionLength = 0
                    'Call ScrollToCursor(txtText)
                    'Call PlayDoneSound ' Don't play the done sound, it's really intrusive. Aug 2010.
                End If
                i += direction
                'Debug.Print "Line: " & i
            End While
            If Not found And wrap Then
                If direction > 0 Then
                    i = 0
                Else
                    i = numberLines - 1
                End If
                While i >= 0 And i <= currentLine And Not found
                    lineContent = GetLine(txtText, i)
                    If InStr(1, lineContent, startsWith, CompareMethod.Text) = 1 Then
                        'found it
                        found = True
                        txtText.SelectionStart = GetCharacterIndexOfLine(txtText, i)
                        txtText.SelectionLength = 0
                        'Call ScrollToCursor(txtText)
                        'Call PlayDoneSound ' Don't play the done sound, it's really intrusive. Aug 2010.
                    End If
                    i += direction
                End While
            End If
            If Not found Then Call PlayErrorSound()
            GotoLineStarting = found
        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Sub mnuLinksSkipup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksSkipup.Click
        On Error Resume Next
        Call DoSkip(SKIP_UP)
    End Sub

    Public Sub mnuNavigateGotoform_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateGotoform.Click
        On Error Resume Next
        'Causes the cursor to jump to the next input element over any links
        Call SkipToFormElement(1)
    End Sub

    Private Sub SkipToFormElement(ByRef direction As Integer)
        Try
            Dim lineNumber As Integer
            Dim found As Boolean
            Dim line As String
            Dim numberLines As Integer

            lineNumber = GetCurrentLineIndex(txtText) + direction
            numberLines = GetNumberOfLines(txtText)
            found = False
            While lineNumber < numberLines And lineNumber >= 0 And Not found
                line = GetLine(txtText, lineNumber)
                If LineIsForm(line) Then
                    'found form element
                    found = True
                Else
                    'keep going
                    lineNumber += direction
                End If
            End While
            If found Then
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                txtText.SelectionLength = 0
                'Call ScrollToCursor(txtText)
                'Don't select - stops screenreaders reading line, they read the document instead.
                'Call SelectCurrentParagraph
            Else
                Call PlayErrorSound()
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuLinksPreviousLink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksPreviouslink.Click
        Try
            'go up to previous link

            Dim lineNumber As Integer
            Dim linkFound As Boolean

            lineNumber = GetCurrentLineIndex(txtText) - 1
            linkFound = False
            While lineNumber >= 0 And Not linkFound
                If txtText.Lines(lineNumber).Contains(ID_LINK) Then
                    'found a link
                    linkFound = True
                Else
                    lineNumber -= 1
                End If
            End While
            If linkFound Then
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                txtText.SelectionLength = 0
                'txtText.SelLength = Len(Trim(modAPIFunctions.GetCurrentLine(txtText)))
                'Call ScrollToCursor(txtText)
            Else
                Call PlayErrorSound()
            End If
        Catch
        End Try
    End Sub



    Public Sub mnuLinksSkiplinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksSkipDown.Click
        On Error Resume Next
        Call DoSkip(SKIP_DOWN)
    End Sub

    ''' <summary>
    ''' returns true if the line is a form element, false if not
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function LineIsForm(ByRef line As String) As Boolean
        On Error Resume Next

        LineIsForm = True
        If InStr(1, line, ID_SELECT, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_BUTTON, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_CHECKBOX, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_RADIO, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_TEXTBOX, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_PASSWORD, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_SUBMIT, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_FILE, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_RESET, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_TEXTAREA, CompareMethod.Binary) = 1 Then
        Else
            'nope, this isn't a form element
            LineIsForm = False
        End If
    End Function

    ''' <summary>
    ''' returns true if the line is text, false if it is a link or form element or similar: used by mnuViewCroppage
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Sub CropPage()
        Try
            Static originalContents As String
            Static preCropPosition As Integer

            If mCropped Then
                ' Uncrop the page
                SetText(originalContents)
                If preCropPosition > -1 AndAlso preCropPosition < txtText.TextLength Then
                    txtText.SelectionStart = preCropPosition
                End If
                mCropped = False
                mnuViewCroppage.Text = modI18N.GetText("&Crop page")
            Else
                ' Crop the page
                originalContents = txtText.Text
                preCropPosition = txtText.SelectionStart

                ' Prioritize <main>, then <article>
                Dim mainElements = gInteractableElements.FindAll(Function(el) el.type = "main")
                Dim articleElements = gInteractableElements.FindAll(Function(el) el.type = "article")

                Dim elementsToKeep As List(Of InteractableElement)

                If mainElements.Count > 0 Then
                    elementsToKeep = mainElements
                ElseIf articleElements.Count > 0 Then
                    elementsToKeep = articleElements
                Else
                    ' Fallback: Keep headings, paragraphs, and general text blocks
                    elementsToKeep = gInteractableElements.FindAll(Function(el)
                                                                       Return el.type.StartsWith("h") OrElse
                                                                              el.type = "p" OrElse
                                                                              el.type = "div" OrElse
                                                                              el.type = "span" OrElse
                                                                              el.type = "text"
                                                                   End Function)
                End If

                ' Build the new cropped text
                Dim croppedOutput As New System.Text.StringBuilder()
                For Each element As InteractableElement In elementsToKeep
                    ' We can reuse the same formatting logic from ParseDocument for consistency if needed
                    ' For now, just append the text.
                    If Not String.IsNullOrWhiteSpace(element.text) Then
                        croppedOutput.AppendLine(element.text)
                    End If
                Next

                If croppedOutput.Length > 0 Then
                    SetText(croppedOutput.ToString())
                    mCropped = True
                    mnuViewCroppage.Text = modI18N.GetText("Un&crop page")
                Else
                    ' Cropping resulted in an empty page, so do nothing.
                    PlayErrorSound()
                End If
            End If
        Catch ex As Exception
            Debug.Print("Error in CropPage: " & ex.Message)
        End Try
    End Sub

    Public Sub mnuViewCroppage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewCroppage.Click
        On Error Resume Next
        Call CropPage()
    End Sub

    Private Sub SetText(text As String)
        Try
            Call StopBusyAnimation()
            Call txtText.Clear()
            txtText.SelectionIndent = 10
            txtText.SelectionRightIndent = 10
            txtText.Text = text
        Catch ex As ObjectDisposedException
            'We get this if you close WebbIE when the web browser is still navigating,
            'so the txtText object has been disposed, but the web browser events are
            'being cleaned up. So ignore, we're exiting.
        End Try
    End Sub

    Private Function GetUserFriendlyURL(ByRef url As String) As String
        Try
            Dim found As Integer
            Dim friendlyURL As String

            friendlyURL = url
            'take off ? queries
            found = InStr(1, friendlyURL, "?")
            If found > 0 Then
                friendlyURL = Microsoft.VisualBasic.Left(friendlyURL, Len(friendlyURL) - found - 1)
            End If
            'take off //
            found = InStr(1, friendlyURL, "//")
            If found > 0 Then
                friendlyURL = Microsoft.VisualBasic.Right(friendlyURL, Len(friendlyURL) - found - 1)
            End If
            'take off trailing /
            If Microsoft.VisualBasic.Right(friendlyURL, 1) = "/" Then
                friendlyURL = Microsoft.VisualBasic.Left(friendlyURL, Len(friendlyURL) - 1)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object GetUserFriendlyURL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetUserFriendlyURL = friendlyURL
        Catch
            Return url
        End Try
    End Function

    Public Async Sub mnuViewSource_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewSource.Click
        Try
            Dim path As String
            Dim htmlString As String = Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync("document.documentElement.outerHTML")
            ' The result from ExecuteScriptAsync is a JSON-encoded string. We need to un-escape it.
            htmlString = System.Text.RegularExpressions.Regex.Unescape(htmlString)
            ' It's also wrapped in quotes, so remove them.
            If htmlString.StartsWith("""") AndAlso htmlString.EndsWith("""") Then
                htmlString = htmlString.Substring(1, htmlString.Length - 2)
            End If

            frmTextView.Text = modI18N.GetText("Page HTML Source")
            path = System.IO.Path.GetTempFileName
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(path)
            Call sw.Write(htmlString)
            Call sw.Close()
            Call Shell("notepad.exe """ & path & """", AppWinStyle.NormalFocus)
        Catch
        End Try
    End Sub

    Private Sub cboaddress_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboAddress.KeyPress
        Try
            If Microsoft.VisualBasic.AscW(eventArgs.KeyChar) = Keys.Return Then 'if return is hit
                eventArgs.Handled = True
                If cboAddress.SelectedIndex > -1 Then
                    cboAddress.Text = CStr(cboAddress.Items(cboAddress.SelectedIndex))
                End If
                Call UserEnteredURL(cboAddress.Text.Trim) ' Don't change case! Breaks Unix servers.
            End If
        Catch
        End Try
    End Sub

    Public Sub UserEnteredURL(ByVal targetURL As String)
        Try
            If targetURL = "" Then Exit Sub ' nowhere to go if the address bar is empty
            'okay, so there is something to get in the address box.
            'If it looks like a url or local file, go get it.
            'Otherwise, start a Google search.
            If targetURL = "home" Then
                'go to home
                Call btnHome_Click(btnHome, New System.EventArgs())
                'This "check if it's a string" is nice in theory, but there
                'are too many dumb things people can type or paste into the address
                'box for it to work - odd characters, spaces, local paths.
                'ElseIf System.Uri.IsWellFormedUriString(targetURL, UriKind.Absolute) Then
                '    'it's definitely a url or local file or config line
                '    Call StartNavigating(targetURL) 'navigate to specified address
                'ElseIf System.Uri.IsWellFormedUriString("http://" & targetURL, UriKind.Absolute) Then
                '    'Yep, again, looks like a URL.
                '    Call StartNavigating(targetURL)
                'Else
            ElseIf targetURL.ToLowerInvariant().Trim = "localhost" Then
                'Local web server!
                Call StartNavigating("localhost")
            ElseIf targetURL.Contains(".") Then
                'Looks ENOUGH like a web address!
                Call StartNavigating(targetURL)
            ElseIf targetURL.ToLowerInvariant().Trim = "about:blank" Then
                'Blank page.
                Call StartNavigating("about:blank")
            Else
                'No idea what this is, do a google search.
                Call frmGoogle.StartSearch(targetURL)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' UpgradeSettings. This function migrates your application's settings from the previous
    '''    /// version, if any, to this one. This is because Properties.Settings are saved to a 
    '''    /// different user folder with every version, so unless you explicitly call this function
    '''    /// then user settings will be lost with every upgrade.
    '''    /// You must create a String setting called "LastVersionRun"
    '''    /// Alasdair 11 June 2013
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpgradeSettings()
        Try
            If My.Settings.LastVersionRun <> Application.ProductVersion Then
                My.Settings.Upgrade()
                My.Settings.Reload()
                My.Settings.LastVersionRun = Application.ProductVersion
                My.Settings.Save()
            End If
        Catch
            'MessageBox.Show("Error in UpgradeSettings. Have you created a property called \"LastVersionRun\"?");
        End Try
    End Sub
    ' End of UpgradeSettings()

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            ' Set up global reference and tool host
            modGlobals.gFrmMain = Me
            Me.ToolHost = New ToolHost(Me)

            'load settings, displays, bookmarks when the program starts up.
            Dim commandLine As String
            Dim helpIndex As String
            Dim startedNavigating As Boolean
            Dim filePath As String = ""

            'Upgrade settings
            Call UpgradeSettings()

            'Check for updates. Only do this is WebbIEUpdater.dll is found. This is so that we can do a 
            'build without the updater for the Windows 10 Store. 
            Dim exeFolder As String = System.IO.Path.GetDirectoryName(Application.ExecutablePath)
            If System.IO.File.Exists(exeFolder & "\WebbIEUpdater.dll") Then
                Call WebbIE.Updater.CheckForUpdates("https://www.webbie.org.uk/webbrowser/updates.xml")
            End If

            'Do Windows XP style.
            Call Application.EnableVisualStyles()

            'Stop Javascript errors showing up in our web browsers.
            Call DisableScriptErrors()

            'Instantiate the WebBrowser object we'll use.
            modGlobals.gWebHost = New frmWeb

            While modGlobals.gWebHost Is Nothing
                Call Threading.Thread.Sleep(100)
                Application.DoEvents()
            End While
            Call modGlobals.gWebHost.Show()

            txtText.ReadOnly = True 'make txtText non-editable.
            txtText.DetectUrls = False ' Leave webbie to do this!
            txtText.ScrollBars = RichTextBoxScrollBars.Vertical

            'set the tabs correctly.
            Call SetupTabs()

            'language settings
            'get the language settings for the default character set for the user's system
            Call modCharacterSupport.InitSystemLocale()
            Call modCharacterSupport.InitCharsetMappings()

            'Load data structures
            Call LoadDataStructures()

            Call modGlobals.Initialise()

            staMain.Items.Item(0).Text = modI18N.GetText("Idle")

            txtText.Font = My.Settings.DisplayFont
            cboAddress.Font = My.Settings.DisplayFont

            'Do colours
            Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))

            modGlobals.gWebHost.webMain.ScriptErrorsSuppressed = Not My.Settings.AllowMessages

            Call ResizeToolbar()

            'query the registry for the current IE homepage
            gCurrentHomepage = RetrieveStartPage()
            'If we don't use the IE homepage, disable the ability to change it.
            mnuOptionsSetHomepage.Enabled = My.Settings.UseIEHomepage

            Call modI18N.DoForm(Me)

            ' Make full-screen or normal.
            Me.WindowState = My.Settings.WindowState

            Call cboAddress.Items.Clear()

            'check for commandline request to go somewhere.
            commandLine = Microsoft.VisualBasic.Command().ToLowerInvariant.Trim
            If commandLine = "" Then
                'OK, we're just starting up.
            ElseIf commandLine = "-hide" Then
                'Hide myself from the user. This is to comply with Microsoft's IE anti-trust suit. 
                If MessageBox.Show("To hide access to this program, you need to uninstall it by using Add/Remove Programs in Control Panel.\n\nWould you like to start Add/Remove Programs?", Application.ProductName, MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                    Dim proc As System.Diagnostics.Process = New System.Diagnostics.Process()
                    Dim psi As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo("appwiz.cpl")
                    psi.UseShellExecute = True
                    proc.StartInfo = psi
                    Call proc.Start()
                End If
                Call Application.Exit()
            ElseIf commandLine = "-reinstall" Then
                'Make myself the default web browser:
                Try
                    Dim regKey As Microsoft.Win32.RegistryKey
                    regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\.htm", True)
                    Call regKey.SetValue("", "WebbIE.HTM.4") ' Windows XP?
                    Call regKey.Close()
                    regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\.html", True)
                    Call regKey.SetValue("", "WebbIE.HTM.4") ' Windows XP?
                    Call regKey.Close()
                    regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Clients\StartMenuInternet", True)
                    Call regKey.SetValue("", "WEBBIE4.EXE") ' Default Browser.
                    Call regKey.Close()
                    'Below is something I think only applies to MIME types. See http://msdn.microsoft.com/en-us/library/windows/desktop/cc144154(v=vs.85).aspx
                    'regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\MIME\Database\Content Type\text/html", True)
                    'Call regKey.SetValue("CLSID", "{D664177E-25AE-4B14-ADD0-FFD930C9C210}")
                    'Dim binValue As Byte() = {&H8, &H0, &H0, &H0}
                    'Call regKey.SetValue("Encoding", binValue, Microsoft.Win32.RegistryValueKind.Binary)
                    'Call regKey.SetValue("Extension", ".htm")
                    'Call regKey.Close()
                Catch ex As Exception
                    'Failed to set registry keys
                    MessageBox.Show(modI18N.GetText("WebbIE failed to set itself as the default web browser:") & vbNewLine & ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End Try

            ElseIf commandLine = "-show" Then
                'Show myself to the user, but don't make default. This is to comply with Microsoft's 
                'IE anti-trust suit. I can ignore it because Hide is done by removing the program,
                'so this never makes any sense to call. 
                Call Application.Exit()
            ElseIf commandLine <> "" Then
                'goody, we have an address already
                'strip off any "" first
                commandLine = commandLine.Replace("""", "").Replace("""", "")
                'is this an absolute path (e.g. "D:\page.htm") or a relative one (e.g. "page.htm")
                Try
                    If System.IO.File.Exists(commandLine) Then
                        filePath = commandLine
                    ElseIf System.IO.File.Exists(My.Application.Info.DirectoryPath & "\" & commandLine) Then
                        filePath = My.Application.Info.DirectoryPath & "\" & commandLine
                    ElseIf System.IO.File.Exists(My.Application.Info.DirectoryPath & commandLine) Then
                        filePath = My.Application.Info.DirectoryPath & commandLine
                    End If
                Catch ex As Exception
                    filePath = "" ' Probably an invalid path.
                End Try
                'Does file indicated by commandline exist?
                If Len(filePath) > 0 Then
                    'yes!
                    startedNavigating = True
                    If System.IO.Path.GetExtension(filePath).ToLowerInvariant = ".chm" And filePath.Contains("\") Then
                        'local help file
                        Call Me.Show()
                        frmParseHTMLHelp = New frmParseHTMLHelp
                        helpIndex = frmParseHTMLHelp.ConvertHTMLHelp(filePath)
                        If Len(helpIndex) > 0 Then
                            cboAddress.Text = helpIndex
                            startedNavigating = True
                        End If
                    Else
                        'just an HTML file
                        cboAddress.Text = filePath
                        startedNavigating = True
                    End If
                Else
                    'nope!
                    'web page
                    cboAddress.Text = commandLine 'set the address to the command line argument
                    startedNavigating = True
                End If

            End If

            gDisplayingYoutube = False


            If startedNavigating Then
                'okay, we've started going somewhere from the command prompt
                Call StartNavigating(cboAddress.Text)
            Else
                Call DoHome()
            End If

            'Load bookmarks after loading and setting off browsing, because it takes ages.
            tmrDelayLoadBookmarks.Enabled = True


#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
        'Subscribe to UI Automation events so we can spot when modGlobals.gWebHost.webMain steals focus and bring it back.
        focusHandler = New System.Windows.Automation.AutomationFocusChangedEventHandler(AddressOf OnFocusChanged)
        'This throws a System.ArgumentException in debug mode, but I THINK it's an internal thing you won't see when running.

        Call System.Windows.Automation.Automation.AddAutomationFocusChangedEventHandler(focusHandler)
#End If
        Catch ex As Exception
            Call Debug.Print("Exception in frmMain_Load: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Go to homepage
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DoHome()
        Try
            If My.Settings.UseIEHomepage Then
                cboAddress.Text = gCurrentHomepage
                Call StartNavigating(cboAddress.Text)
            Else
                Call GotoWebbIEHomePage()
            End If
        Catch
        End Try
    End Sub

    Private Sub LoadDataStructures()
        Try
            Dim i As Integer

            For i = 0 To MAX_NUMBER_BUTTON_INPUTS_SUPPORTED - 1
                Call selects(i).Initialize()
            Next i
        Catch
        End Try
    End Sub

    Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'indicate we should stop operations (parsing web pages)
            gExiting = True

#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
        'Unsubscribe from UI Automation events
        If (focusHandler IsNot Nothing) Then
            Call System.Windows.Automation.Automation.RemoveAutomationFocusChangedEventHandler(focusHandler)
        End If
#End If
            'save user settings 
            If Me.WindowState = FormWindowState.Normal Then
                My.Settings.WindowState = FormWindowState.Normal
            ElseIf Me.WindowState = FormWindowState.Maximized Then
                My.Settings.WindowState = FormWindowState.Maximized
            End If
            My.Settings.DisplayFont = Me.txtText.Font
            Call My.Settings.Save()
            'Restore script debugging settings
            Call RestoreScriptErrors()
        Catch
        End Try
    End Sub

    Public Sub SetHomepage()
        Try
            'Update the home page to the current url
            Call SetStartPage(cboAddress.Text)
            'update our global variable: gCurrentHomepage
            gCurrentHomepage = cboAddress.Text
            'update tooltip
            'Tell the user what we've done
            MsgBox(modI18N.GetText("Home page changed. WebbIE will come here when it starts or you select Home from the Navigate menu."), MsgBoxStyle.OkOnly)
        Catch
        End Try
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpAbout.Click
        On Error Resume Next
        Call ShowAbout()
    End Sub

    Public Sub ShowAbout()
        On Error Resume Next
        Call MsgBox(Application.ProductName & vbTab & Application.ProductVersion, MsgBoxStyle.Information, Application.ProductName)
    End Sub

    ''' <summary>
    ''' User selects back option from menu
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Public Sub mnuNavigateBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateBack.Click
        On Error Resume Next
        Call btnBack_Click(btnBack, New System.EventArgs())
    End Sub

    ''' <summary>
    ''' Convert invalid filenames into valid DOS names.
    ''' </summary>
    ''' <param name="inputFilename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetValidFilename(ByRef inputFilename As String) As String
        On Error Resume Next
        For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
            inputFilename = inputFilename.Replace(c, "_")
        Next c
        Return inputFilename.Substring(0, 255)
    End Function


    Public Sub mnuOptionsChangeFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOptionsFont.Click
        'user asks to change the font - show appropriate dialogue
        On Error GoTo errorHandler
        Me.fdMain.ShowApply = False
        Me.fdMain.AllowSimulations = False
        Me.fdMain.AllowVerticalFonts = False
        Me.fdMain.ShowEffects = False
        Me.fdMain.ShowHelp = False
        Me.fdMain.Font = CType(Me.txtText.Font.Clone(), System.Drawing.Font)
        If Me.fdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Me.txtText.Font = fdMain.Font
        End If
        Exit Sub
errorHandler:
        Exit Sub
    End Sub

    Public Sub mnuEditCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditCopy.Click
        'process user menu command copy
        Try
            If Not Me.ActiveControl Is Nothing Then
                Select Case Me.ActiveControl.Name
                    Case "txtText"
                        If txtText.SelectedText <> "" Then ' only copy if we have something selected
                            Call My.Computer.Clipboard.Clear()
                            Call My.Computer.Clipboard.SetText(txtText.SelectedText, System.Windows.Forms.TextDataFormat.UnicodeText)
                        End If
                    Case "cboAddress"
                        If cboAddress.SelectedText <> "" Then ' only copy if we have something selected
                            Call My.Computer.Clipboard.Clear()
                            Call My.Computer.Clipboard.SetText(cboAddress.SelectedText)
                        End If
                    Case Else
                        Dim wb As SHDocVw.WebBrowser
                        wb = CType(modGlobals.gWebHost.webMain.ActiveXInstance, SHDocVw.WebBrowser)
                        Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_COPY, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT)
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub mnuEditPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditPaste.Click
        On Error Resume Next
        Call Paste()
    End Sub

    Public Sub mnuEditSelectall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditSelectall.Click
        Try
            'select everything on the current object.
            Dim currentObjectName As String

            If Me.ActiveControl Is Nothing Then
                currentObjectName = modGlobals.gWebHost.webMain.Name
            Else
                currentObjectName = Me.ActiveControl.Name
            End If
            If currentObjectName = modGlobals.gWebHost.webMain.Name Then
                Call modSendKeys.SendSelectAll()
            ElseIf currentObjectName = cboAddress.Name Then
                Call cboAddress.SelectAll()
            ElseIf currentObjectName = txtText.Name Then
                Call txtText.SelectAll()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Find the next instance of a word searched for by the user and select it. Display a message if not found. If no word to search for, call FindText()
    ''' </summary>
    Private Sub FindNext()
        Try
            '

            Dim where As Integer
            Dim originalPos As Integer
            Dim searchIn As String

            'If mIEVisible Then
            '    Dim tr As mshtml.IHTMLTxtRange
            '    Dim doc As mshtml.IHTMLDocument2
            '    doc = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.IHTMLDocument2)
            '    tr = CType(doc.selection.createRange, mshtml.IHTMLTxtRange)
            '    If tr Is Nothing Then
            '    ElseIf tr.text = "" Then
            '        'Start a new Find
            '        Dim wb As SHDocVw.WebBrowser
            '        wb = CType(modGlobals.gWebHost.webMain.ActiveXInstance, SHDocVw.WebBrowser)
            '        Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_FIND, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT, CObj(gfindText))
            '    Else
            '        'Find the next instance of the selected text.
            '        gfindText = tr.text
            '        Call tr.collapse(False)
            '        If tr.findText(gfindText) Then
            '            Call tr.select()
            '            Call tr.scrollIntoView()
            '        Else
            '            Call Beep()
            '        End If
            '    End If

            'Else
            If gfindText <> "" Then
                originalPos = txtText.SelectionStart 'go to start of the found word
                'make lower case and search onwards from the existing word
                searchIn = Mid(txtText.Text, originalPos + 2, Len(txtText.Text) - txtText.SelectionStart - 2)
                where = InStr(1, searchIn, gfindText, CompareMethod.Text)
                If (where > 0) Then 'if found
                    where += originalPos
                    'startLine = GetCurrentLineIndex(txtText)
                    txtText.SelectionStart = where
                    txtText.SelectionLength = Len(gfindText) 'highlight the word
                    'endLine = GetCurrentLineIndex(txtText)
                    'Call Scroll(txtText, endLine - startLine)
                    'Call ScrollToCursor(txtText)
                Else
                    'if unfound, display a warning
                    MsgBox(modI18N.GetText("No further occurrences found"), MsgBoxStyle.Information, Application.ProductName)
                End If
            Else 'if there is no text to search
                Call FindText()
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuEditFindnext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFindnext.Click
        On Error Resume Next
        Call FindNext()
    End Sub

    Private Sub FindText()
        Try
            'Find a word looked for in the main text box

            Dim where As Integer

            'in text view
            gfindText = InputBox(modI18N.GetText("Find what:"), modI18N.GetText("Find"), gfindText)
            'if OK is clicked
            If gfindText <> "" Then
                ' Find string in text
                where = InStr(txtText.SelectionStart + 2, txtText.Text, gfindText, CompareMethod.Text)
                If where > 0 Then ' If found..
                    Call txtText.Focus()
                    'startLine = GetCurrentLineIndex(txtText)

                    txtText.SelectionStart = where - 1 ' set selection start and
                    txtText.SelectionLength = Len(gfindText) ' set selection length
                    Call txtText.ScrollToCaret()
                Else
                    'try from start
                    where = InStr(1, txtText.Text, gfindText, CompareMethod.Text)
                    If where > 0 Then
                        'found it: is it where we started?
                        If where = txtText.SelectionStart + 1 Then
                            'whoops, already there
                            'Debug.Print "Already there"
                            MsgBox(modI18N.GetText("No more") & " " & gfindText, MsgBoxStyle.Information, modI18N.GetText("Word not found"))
                        Else
                            'nope, go there
                            Call txtText.Focus()
                            txtText.SelectionStart = where - 1
                            txtText.SelectionLength = Len(gfindText)
                        End If
                    Else
                        'not found at all
                        MsgBox(modI18N.GetText("Cannot find") & " " & gfindText, MsgBoxStyle.Information, modI18N.GetText("Word not found"))
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuEditFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFindtext.Click
        On Error Resume Next
        Call FindText()
    End Sub

    Public Sub mnuNavigateForward_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateForward.Click
        On Error Resume Next
        Call DoForward()
    End Sub

    Public Sub mnuNavigateHome_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateHome.Click
        On Error Resume Next
        btnHome_Click(btnHome, New System.EventArgs())
    End Sub

    Public Sub mnuLinksNextLink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksNextlink.Click
        Try
            Dim lineNumber As Integer
            Dim linkFound As Boolean

            If Me.txtText.Text <> "" Then
                lineNumber = GetCurrentLineIndex(txtText) + 1
                linkFound = False
                While lineNumber < GetNumberOfLines(txtText) And Not linkFound
                    If txtText.Lines(lineNumber).Contains(ID_LINK) Then
                        'found a link
                        linkFound = True
                    Else
                        lineNumber += 1
                    End If
                End While
                If linkFound Then
                    txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                    'Don't set length of link.
                    'txtText.SelLength = Len(Trim(GetCurrentLine(txtText)))
                    txtText.SelectionLength = 0
                    'Call ScrollToCursor(txtText)
                Else
                    Call PlayErrorSound()
                End If
            End If
        Catch
        End Try
    End Sub

#Region "Printing"
    'http://msdn.microsoft.com/en-us/library/system.drawing.printing.printdocument.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-2

    Private ReadOnly _printFont As Font
    Private ReadOnly _streamToPrint As System.IO.StreamReader

    Public Sub mnuFilePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePrint.Click
        'print the webpage
        Try
            'So this code prints directly from the WebBrowser control on frmWeb.
            Dim wb As SHDocVw.WebBrowser
            wb = CType(gWebHost.webMain.ActiveXInstance, SHDocVw.WebBrowser)
            Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_PRINT, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT)
        Catch ex As Exception
            'Ah, this isn't written right: won't do anything.
            'Call System.Diagnostics.EventLog.WriteEntry("mnuFilePrint_Click", "Failed to print using OLECMDID_PRINT: " + ex.Message)
        End Try

        'Whereas THIS code prints the actual text in the text area. I've had a request to print the actual
        'web page, which of course is useful for tickets and bank statements and the like, so this makes sense.
        'If a user wants to print the whole page of actual text they can cut/paste it into Word or similar. They
        'can't get at the web page.
        'Dim pd As New System.Drawing.Printing.PrintDocument()
        'Dim printDialog As Windows.Forms.PrintDialog = New Windows.Forms.PrintDialog()
        'printDialog.Document = pd
        'If (printDialog.ShowDialog() = DialogResult.OK) Then

        '    Dim path As String = System.IO.Path.GetTempFileName
        '    Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(path, False, System.Text.Encoding.UTF8)
        '    Call sw.Write(txtText.Text)
        '    Call sw.Close()
        '    Call Application.DoEvents()
        '    Try
        '        _streamToPrint = New System.IO.StreamReader(path, System.Text.Encoding.UTF8)
        '        Try
        '            _printFont = CType(txtText.Font.Clone, Drawing.Font) 'New Font("Arial", 10)
        '            AddHandler pd.PrintPage, AddressOf Me.pd_PrintPage
        '            pd.Print()
        '        Finally
        '            _streamToPrint.Close()
        '        End Try
        '    Catch ex As Exception
        '        MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End If
    End Sub

    ' The PrintPage event is raised for each page to be printed. 
    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As System.Drawing.Printing.PrintPageEventArgs)
        Try
            Dim linesPerPage As Single = 0
            Dim yPos As Single = 0
            Dim count As Integer = 0
            Dim leftMargin As Single = ev.MarginBounds.Left
            Dim topMargin As Single = ev.MarginBounds.Top
            Dim line As String = Nothing

            ' Calculate the number of lines per page.
            linesPerPage = ev.MarginBounds.Height / _printFont.GetHeight(ev.Graphics)

            ' Print each line of the file. 
            While count < linesPerPage
                line = _streamToPrint.ReadLine()
                If line Is Nothing Then
                    Exit While
                End If
                yPos = topMargin + count * _printFont.GetHeight(ev.Graphics)
                ev.Graphics.DrawString(line, _printFont, Brushes.Black, leftMargin, yPos, New StringFormat())
                count += 1
            End While

            ' If more lines exist, print another page. 
            If (line IsNot Nothing) Then
                ev.HasMorePages = True
            Else
                ev.HasMorePages = False
            End If
        Catch
        End Try
    End Sub

#End Region

    Public Sub mnuNavigateRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateRefresh.Click
        On Error Resume Next
        If modGlobals.gWebHost.webMain IsNot Nothing AndAlso modGlobals.gWebHost.webMain.CoreWebView2 IsNot Nothing Then
            modGlobals.gWebHost.webMain.CoreWebView2.Reload()
        End If
    End Sub

    Public Sub mnuNavigateStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateStop.Click
        On Error Resume Next
        Call btnStop_Click(btnStop, New System.EventArgs())
    End Sub

    ''' <summary>
    ''' Prepare the links from the current page and show the lstLinks form to let
    ''' the user select one.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ListLinks()
        Try
            ' Filter for link elements from the main list
            Dim linksOnly As List(Of InteractableElement) = gInteractableElements.FindAll(Function(el) el.type = "link")

            ' Sort the links alphabetically by their text
            linksOnly.Sort(Function(el1, el2) String.Compare(el1.text, el2.text, StringComparison.CurrentCulture))

            ' Pass the sorted list to the form and show it
            Call frmLinks.PopulateList(linksOnly) ' New signature for PopulateList
            frmLinks.cmdGo.Enabled = False
            Call frmLinks.ShowDialog(Me)
        Catch ex As Exception
            Debug.Print("Error in ListLinks: " & ex.Message)
        End Try
    End Sub

    Public Sub mnuLinksViewLinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksViewlinks.Click
        On Error Resume Next
        Call ListLinks()
    End Sub

    Private Sub mnuEditWebsearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditWebsearch.Click
        On Error Resume Next
        Call DoWebSearch()
    End Sub

    ''' <summary>
    ''' Pop up a box to prompt user for a search phrase for google
    ''' </summary>
    Public Sub DoWebSearch()
        On Error Resume Next
        If Not frmGoogle.Visible Then
            Call frmGoogle.ShowDialog(Me)
        End If
    End Sub

    Private Sub StopBusyAnimation()
        Try
            tmrBusyAnimation.Enabled = False
            If My.Settings.ToolbarCaptions Then
                picBusy.Image = My.Resources.timer_done_big
            Else
                picBusy.Image = My.Resources.timer_done
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrBusyAnimation_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrBusyAnimation.Tick
        If modGlobals.gClosing Then
            Me.tmrBusyAnimation.Enabled = False
            Exit Sub
        End If
        'Not busy? Don't play. This is because we get a whole lot of clicking when we shouldn't. Fails to stop this timer
        'probably. 
        'But then we don't get lots of ticking when we should! 

        Try
            'every 0.3 second, move the animation. Play working sound 
            Static counter As Integer 'static retains value on each call

            'Debug.Print("tmrBusy! " & New System.Random().Next())
            Select Case counter
                Case 0
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer1_big
                    Else
                        picBusy.Image = My.Resources.timer1
                    End If
                    counter = 1
                    'Call Debug.Print("Working sound in image graphic")
                    Call PlayWorkingSound()
                Case 1
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer2_big
                    Else
                        picBusy.Image = My.Resources.timer2
                    End If
                    counter = 2
                Case 2
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer3_big
                    Else
                        picBusy.Image = My.Resources.timer3
                    End If
                    counter = 3
                Case 3
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer4_big
                    Else
                        picBusy.Image = My.Resources.timer4
                    End If
                    counter = 0
            End Select
        Catch
        End Try
    End Sub

    Private Sub ClearPageData()
        On Error Resume Next
        'Clears the page arrays and other per-page data.
        gInteractableElements.Clear()
        gRSSFeedURL = ""
        'These flags can be determined from the new parsed data if needed
        gPageHasAnArticle = False
        gPageHasMain = False
    End Sub

    Private Sub DisplayOutput()
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            Dim newPageContent As New System.Text.StringBuilder()

            If modGlobals.gWebHost.webMain IsNot Nothing AndAlso modGlobals.gWebHost.webMain.CoreWebView2 IsNot Nothing Then
                If Not String.IsNullOrEmpty(modGlobals.gWebHost.webMain.CoreWebView2.DocumentTitle) Then
                    Call newPageContent.AppendLine(modI18N.GetText("WEBPAGE:") & " " & modGlobals.gWebHost.webMain.CoreWebView2.DocumentTitle)
                    Call newPageContent.AppendLine()
                End If
            End If

            ' The new mOutput is already cleaned up by the parser logic.
            ' We can add more advanced blank line removal here if needed, but for now, just append.
            newPageContent.Append(mOutput.ToString())

            Call SetText(newPageContent.ToString())
        Catch ex As Exception
            Debug.Print("Exception in DisplayOutput: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Process the loading HTML object model and display it as text. Call this after refresh, page navigation, or any time
    ''' you need to update this view to reflect a change to the DOM.
    ''' </summary>
    Private _testCompletionSource As TaskCompletionSource(Of Boolean)

    Private Class GoldenTestCase
        Public Property name As String
        Public Property user_command As String
        Public Property mock_dom_content As String
        Public Property expected_tool As String
        Public Property expected_args As List(Of String)
    End Class

    Private Async Sub btnRunTests_Click(sender As Object, e As EventArgs) Handles btnRunTests.Click
        Dim overallResults As New System.Text.StringBuilder()
        Dim totalPassed As Integer = 0
        Dim totalFailed As Integer = 0

        ' ********** Layer 1: DOM Parser Tests **********
        overallResults.AppendLine("--- LAYER 1: DOM Parser Tests ---")
        Dim layer1Passed As Integer = 0
        Dim layer1Failed As Integer = 0
        Try
            Dim testDirectory As String = System.IO.Path.Combine(Application.StartupPath, "Tests", "mock_pages")
            If Not System.IO.Directory.Exists(testDirectory) Then
                overallResults.AppendLine("FAIL: Test directory not found: " & testDirectory)
                layer1Failed += 1
            Else
                Dim testFiles = System.IO.Directory.GetFiles(testDirectory, "*.html")
                For Each testFile As String In testFiles
                    _testCompletionSource = New TaskCompletionSource(Of Boolean)()
                    Dim fileUri As New Uri(testFile)
                    modGlobals.gWebHost.webMain.CoreWebView2.Navigate(fileUri.AbsoluteUri)
                    Await _testCompletionSource.Task

                    Dim actualJson As String = modGlobals.gLastParserResult
                    Dim expectedFile As String = testFile.Replace(".html", ".expected.json")
                    Dim expectedJson As String = If(System.IO.File.Exists(expectedFile), System.IO.File.ReadAllText(expectedFile), "{}")

                    Dim normalizedActual = JsonConvert.SerializeObject(JsonConvert.DeserializeObject(actualJson))
                    Dim normalizedExpected = JsonConvert.SerializeObject(JsonConvert.DeserializeObject(expectedJson))

                    If normalizedActual = normalizedExpected Then
                        layer1Passed += 1
                        overallResults.AppendLine($"PASS: {System.IO.Path.GetFileName(testFile)}")
                    Else
                        layer1Failed += 1
                        overallResults.AppendLine($"FAIL: {System.IO.Path.GetFileName(testFile)}")
                    End If
                Next
            End If
        Catch ex As Exception
            layer1Failed += 1
            overallResults.AppendLine("FATAL ERROR in Layer 1: " & ex.Message)
        Finally
            _testCompletionSource = Nothing
        End Try
        totalPassed += layer1Passed
        totalFailed += layer1Failed
        overallResults.AppendLine($"Layer 1 Complete. Passed: {layer1Passed}, Failed: {layer1Failed}")
        overallResults.AppendLine("------------------------------------")
        overallResults.AppendLine()


        ' ********** Layer 2: LLM Logic Tests **********
        overallResults.AppendLine("--- LAYER 2: LLM Logic Tests ---")
        Dim layer2Passed As Integer = 0
        Dim layer2Failed As Integer = 0
        Try
            Me.ToolHost.IsInTestMode = True
            Dim goldenDatasetFile As String = System.IO.Path.Combine(Application.StartupPath, "Tests", "golden_dataset.json")
            If Not System.IO.File.Exists(goldenDatasetFile) Then
                layer2Failed += 1
                overallResults.AppendLine("FAIL: golden_dataset.json not found.")
            Else
                Dim datasetJson = System.IO.File.ReadAllText(goldenDatasetFile)
                Dim testCases = JsonConvert.DeserializeObject(Of List(Of GoldenTestCase))(datasetJson)

                For Each testCase As GoldenTestCase In testCases
                    Me.ToolHost.LastToolCalled = ""
                    Me.ToolHost.LastArgsCalled = Nothing

                    Dim script = $"window.aiProcess(`{testCase.mock_dom_content.Replace("`", "\`")}`, `{testCase.user_command.Replace("`", "\`")}`);"
                    Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)

                    ' Compare results
                    If Me.ToolHost.LastToolCalled.ToLower() = testCase.expected_tool.ToLower() AndAlso
                       Me.ToolHost.LastArgsCalled.SequenceEqual(testCase.expected_args) Then
                        layer2Passed += 1
                        overallResults.AppendLine($"PASS: {testCase.name}")
                    Else
                        layer2Failed += 1
                        overallResults.AppendLine($"FAIL: {testCase.name}")
                        overallResults.AppendLine($"  Expected: {testCase.expected_tool}({String.Join(", ", testCase.expected_args)})")
                        overallResults.AppendLine($"  Actual:   {Me.ToolHost.LastToolCalled}({String.Join(", ", Me.ToolHost.LastArgsCalled)})")
                    End If
                Next
            End If
        Catch ex As Exception
            layer2Failed += 1
            overallResults.AppendLine("FATAL ERROR in Layer 2: " & ex.Message)
        Finally
            Me.ToolHost.IsInTestMode = False
        End Try
        totalPassed += layer2Passed
        totalFailed += layer2Failed
        overallResults.AppendLine($"Layer 2 Complete. Passed: {layer2Passed}, Failed: {layer2Failed}")
        overallResults.AppendLine("------------------------------------")


        ' ********** Final Summary **********
        Dim summary As String = $"Test run complete. Total Passed: {totalPassed}, Total Failed: {totalFailed}" & vbCrLf & vbCrLf & overallResults.ToString()
        MessageBox.Show(summary, "Test Results", MessageBoxButtons.OK, If(totalFailed > 0, MessageBoxIcon.Error, MessageBoxIcon.Information))
    End Sub

    ''' <summary>
    ''' Process the loading HTML object model and display it as text. Call this after refresh, page navigation, or any time
    ''' you need to update this view to reflect a change to the DOM.
    ''' </summary>
    Private Async Sub ParseDocument()
        If modGlobals.gClosing OrElse mFormClosing Then
            Exit Sub
        End If

        Try
            ' Tell user we're working, unless we're in a test run
            If _testCompletionSource Is Nothing OrElse _testCompletionSource.Task.IsCompleted Then
                staMain.Items.Item(0).Text = modI18N.GetText("Examining")
                Call Application.DoEvents()
            End If

            ' Clear old data
            Call ClearPageData()

            ' Read the DOM parser script from resources
            Dim parserScript As String = My.Resources.dom_parser

            ' Execute the script to get interactable elements as JSON
            Dim jsonResult As String = Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(parserScript)
            modGlobals.gLastParserResult = jsonResult ' Store for testing

            ' The result is a JSON string literal. We need to un-escape and parse it.
            If String.IsNullOrWhiteSpace(jsonResult) OrElse jsonResult = "null" OrElse jsonResult.Length < 3 Then
                ' Script failed or returned nothing
                staMain.Items.Item(0).Text = modI18N.GetText("Done") ' Page might be empty
                Call SetText("") ' Clear the text view
                ' If in a test run, signal completion
                _testCompletionSource?.TrySetResult(True)
                Return
            End If

            ' Deserialize the JSON string into our list of elements
            ' ExecuteScriptAsync returns a JSON-encoded *string*. So we must first deserialize the string itself,
            ' then deserialize the content of that string.
            Dim jsonContent As String = JsonConvert.DeserializeObject(Of String)(jsonResult)
            gInteractableElements = JsonConvert.DeserializeObject(Of List(Of InteractableElement))(jsonContent)

            ' Reset flags
            mCropped = False

            ' Build the output for the text view
            mOutput = New System.Text.StringBuilder(32000)

            ' TODO: Check for RSS feed link separately if needed
            ' gRSSFeedURL = Await CheckForRssAsync() ' This would need a new helper

            ' Process the list of elements to build the text display
            Dim i As Integer = 0
            For Each element As InteractableElement In gInteractableElements
                Dim line As String = ""
                Dim prefix As String = ""
                Dim numberString As String = If(My.Settings.NumberLinks, $" [{i}]", "")

                Select Case element.type
                    Case "link"
                        prefix = ID_LINK
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "button", "submit", "reset"
                        prefix = If(element.type = "submit", ID_SUBMIT, If(element.type = "reset", ID_RESET, ID_BUTTON))
                        line = $"{prefix}{numberString}: ({element.text})"
                    Case "text", "search", "email", "url", "tel"
                        prefix = ID_TEXTBOX
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "password"
                        prefix = ID_PASSWORD
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "textarea"
                        prefix = ID_TEXTAREA
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "checkbox"
                        prefix = ID_CHECKBOX
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "radio"
                        prefix = ID_RADIO
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "select-one", "select-multiple"
                        prefix = ID_SELECT
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "file"
                        prefix = ID_FILE
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "h1", "h2", "h3", "h4", "h5", "h6"
                        line = $"Heading {element.type.Substring(1)}: {element.text}"
                    Case "p", "div", "span", "text", "main", "article"
                        ' For simple text blocks, just show the text. We might want to filter short/empty ones.
                        If Not String.IsNullOrWhiteSpace(element.text) Then
                            line = element.text
                        End If
                    ' Add other cases as needed for VIDEO, AUDIO, etc.
                End Select

                If Not String.IsNullOrWhiteSpace(line) Then
                    Call mOutput.AppendLine(line)
                End If
                i += 1
            Next

            ' Display the result
            Call DisplayOutput()

            ' Update UI state, unless in a test run
            If _testCompletionSource Is Nothing OrElse _testCompletionSource.Task.IsCompleted Then
                btnStop.Enabled = False
                btnRefresh.Enabled = True
                staMain.Items.Item(0).Text = modI18N.GetText("Done")
                mnuLinksViewlinks.Enabled = (gInteractableElements.Count > 0)
                cboAddress.SelectionLength = 0
                cboAddress.Refresh()
                Call txtText.Focus()
                Application.DoEvents()
                txtText.SelectionLength = 0
            End If

        Catch ex As Exception
            Debug.Print("Exception in ParseDocument: " & ex.Message)
            staMain.Items.Item(0).Text = modI18N.GetText("Error")
        Finally
            ' If in a test run, signal completion
            _testCompletionSource?.TrySetResult(True)
            ' Only play sounds if not in a test run
            If _testCompletionSource Is Nothing OrElse _testCompletionSource.Task.IsCompleted Then
                Call PlayDoneSound()
                Call StopBusyAnimation()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' stops all browser activity and ticking
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StopBrowsers()
        Call StopBusyAnimation()
        If modGlobals.gWebHost.webMain IsNot Nothing AndAlso modGlobals.gWebHost.webMain.CoreWebView2 IsNot Nothing Then
            modGlobals.gWebHost.webMain.CoreWebView2.Stop()
        End If
    End Sub

    Public Sub RefreshCurrentPage()
        Try
            'Call this when you want to redisplay the current page without reloading it from the server -
            'for example when you've amended the DOM and want to show the differences

            Dim caretPosition As Integer ' position of caret

            Debug.Print("REFRESH CURRENT PAGE")
            'store the caret position and line number
            caretPosition = txtText.SelectionStart
            'clear text box
            txtText.Clear()
            'Parse page
            Call ParseDocument()
            'simple refresh, restore the caret position
            txtText.SelectionStart = caretPosition
            txtText.SelectionLength = 0
            Call txtText.ScrollToCaret()
            Call UpdateMenus()
            Call Me.txtText.Focus()
        Catch
        End Try
    End Sub

    Private Sub SkipToForm(ByRef direction As Integer)
        Try
            If GotoLineStarting(GetText("Form"), direction) Then
                'Don't select - stops screenreaders reading line, they read the document instead.
                'Call SelectCurrentParagraph
            End If
        Catch
        End Try
    End Sub

    Private Sub SkipLinks(ByRef direction As Integer, Optional ByRef reverse As Boolean = False)
        Try
            'Causes the cursor to jump to the start of the next paragraph over any links
            'If reverse is set then goes to the first line that IS a link or whatever.
            Dim lineNumber As Integer
            Dim found As Boolean
            Dim line As String
            Dim numberLines As Integer
            Dim startPreviousLine As Integer
            Dim endThisLine As Integer
            Dim foundNewline As Integer

            lineNumber = GetCurrentLineIndex(txtText) + direction
            numberLines = GetNumberOfLines(txtText)
            found = False
            While lineNumber < numberLines And lineNumber >= 0 And Not found
                'Debug.Print "Linenumber:" & lineNumber
                line = GetLine(txtText, lineNumber).Trim
                'Debug.Print "Line:[" & line & "]"
                If Not line.Contains(ID_LINK) And InStr(1, line, ID_SELECT, CompareMethod.Text) <> 1 And InStr(1, line, ID_BUTTON, CompareMethod.Text) <> 1 And InStr(1, line, ID_CHECKBOX, CompareMethod.Text) <> 1 And InStr(1, line, ID_RADIO, CompareMethod.Text) <> 1 And InStr(1, line, ID_TEXTBOX, CompareMethod.Text) <> 1 And InStr(1, line, ID_RANGEINPUT, CompareMethod.Binary) <> 1 And InStr(1, line, ID_EMAILINPUT, CompareMethod.Text) <> 1 And InStr(1, line, ID_PASSWORD, CompareMethod.Text) <> 1 And InStr(1, line, ID_SUBMIT, CompareMethod.Text) <> 1 And InStr(1, line, ID_FILE, CompareMethod.Text) <> 1 And InStr(1, line, ID_RESET, CompareMethod.Text) <> 1 And InStr(1, line, ID_TEXTAREA, CompareMethod.Text) <> 1 Then
                    'ASSERTION: The current line is none of the "focusable" items on the page, like a link.
                    'Debug.Print "Structural"
                    If reverse Then
                        'Not any of the types we're looking for, definitely not found.
                        found = False
                    Else
                        'hurrah, it's not a WebbIE structural item!  Now check it has some content
                        If UBound(Split(Trim(line), " ")) > 1 Then
                            'And some content: finally, the question is whether this line is
                            'the start of a paragraph or just some text.
                            'Debug.Print "content:" & line
                            If lineNumber > 0 Then
                                startPreviousLine = modAPIFunctions.GetCharacterIndexOfLine(txtText, lineNumber - 1)
                                If startPreviousLine = 0 Then startPreviousLine = 1
                                endThisLine = modAPIFunctions.GetCharacterIndexOfLine(txtText, lineNumber) + Len(line) - 2
                                foundNewline = InStr(startPreviousLine, txtText.Text, vbNewLine)
                                If foundNewline = 0 Then foundNewline = InStr(startPreviousLine, txtText.Text, vbCr)
                                If foundNewline = 0 Then foundNewline = InStr(startPreviousLine, txtText.Text, vbLf)
                                If (foundNewline > 0) And (foundNewline < endThisLine) Then
                                    'okay, this line is directly preceded by a newline
                                    found = True
                                Else
                                    'nope, no newline: this is part of a paragraph, or end of the page, so skip
                                    lineNumber = lineNumber + direction
                                End If
                            Else
                                'okay, we're at the start of the text, so call it found
                                found = True
                            End If
                        End If
                    End If
                Else
                    If reverse Then
                        'Found a link etc., good!
                        found = True
                    Else
                        'Nope, keep looking.
                        lineNumber = lineNumber + direction
                    End If
                End If
                If Not found Then
                    lineNumber = lineNumber + direction
                End If
            End While
            If found Then
                txtText.SelectionStart = 1
                txtText.SelectionLength = 0
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                txtText.SelectionLength = 0
                'We don't call SelectCurrentParagraph because the selection stops screenreaders (Thunder,
                'Narrator) from reading the line. Instead they read the whole document.
                'Call SelectCurrentParagraph
                Call txtText.ScrollToCaret()
            Else
                Call PlayErrorSound()
            End If
            Call txtText.Focus()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' update the current text area
    ''' </summary>
    ''' <param name="elementText"></param>
    ''' <param name="node"></param>
    ''' <remarks></remarks>

    ''' <summary>
    ''' Called when WebbIE needs to go somewhere new
    ''' </summary>
    ''' <param name="url"></param>
    ''' <remarks></remarks>
    Public Sub StartNavigating(ByRef url As String)
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            If url.ToLowerInvariant().Trim = "about:blank" Then
                Call ClearPageData()
                Call txtText.Clear()
                Call txtText.Focus()
                Call StopBusyAnimation()
            Else
                tmrBusyAnimation.Enabled = True
                'Now reset the "we're only showing forms" mode
                gShowFormsOnly = False
                'Might be XML or HTML
                'finally send WebBrowser on its way, it's a normal file
                'Not going to just do this, going to count stuff.

                Call AddURLToRecent(url)
                cboAddress.Text = url ' update display
                Call PlayStartSound()
                modGlobals.gDesiredURL = url
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Call when page is loaded, including all frames
    ''' </summary>
    ''' <remarks></remarks>
    Friend Async Sub ProcessPage()
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try

            'Note that a page action HAS been triggered.
            gNoPageActionSoRefresh = False

            'update location bar with actual URL (inc. http://etc.)
            If modGlobals.gWebHost.URL = "" Then
            Else
                cboAddress.Text = modGlobals.gWebHost.URL
            End If

            'okay, now go through all the loaded docs and display
            Call ParseDocument()
            'Check for YouTube, and if found then autostart it.
            Call StartYoutubeIfFound()
            'Change display
            Call UpdateMenus()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Updates (enables or disables) menu items and buttons to reflect the state of the application.
    ''' Called (for example) by RefreshCurrentPage after it has run.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateMenus()
        Try
            mnuFilePrint.Enabled = True
            mnuFileSave.Enabled = True
            mnuViewCroppage.Text = modI18N.GetText("&Crop page")
            btnSkiplinks.Enabled = True
            btnHeading.Enabled = True
        Catch
        End Try
    End Sub

    Private Sub StartYoutubeIfFound()
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            If modGlobals.gWebHost.URL.ToUpperInvariant.StartsWith("HTTP://WWW.YOUTUBE.COM") Or modGlobals.gWebHost.URL.ToUpperInvariant().StartsWith("HTTPS://WWW.YOUTUBE.COM") Then
                'Ah, got a YouTube link. Have to click on the video to play it.
                gDisplayingYoutube = True
                Call Application.DoEvents()
                modGlobals.gWebHost.Visible = True
                For i As Integer = 0 To 20
                    Call System.Threading.Thread.Sleep(5)
                    Call Application.DoEvents()
                Next i
                Dim pageUIA As System.Windows.Automation.AutomationElement = System.Windows.Automation.AutomationElement.FromHandle(modGlobals.gWebHost.webMain.Handle)
                Dim pc As System.Windows.Automation.PropertyCondition = New System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ClassNameProperty, "MacromediaFlashPlayerActiveX", System.Windows.Automation.PropertyConditionFlags.IgnoreCase)
                Dim playerUIA As System.Windows.Automation.AutomationElement = pageUIA.FindFirst(System.Windows.Automation.TreeScope.Descendants, pc)
                '        pageUIA.FindAll(TreeScope.Descendants,New PropertyCondition(AutomationElement.
                If playerUIA Is Nothing Then
                Else
                    Dim rect As System.Windows.Rect = playerUIA.Current.BoundingRectangle
                    Call NativeMethods.SetCursorPos(CInt(rect.Left) + 3, CInt(rect.Top) + 3)
                    Call NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, New IntPtr(0))
                    Call NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTUP, 0, 0, 0, New IntPtr(0))
                End If
                For i As Integer = 0 To 20
                    Call System.Threading.Thread.Sleep(5)
                    Call Application.DoEvents()
                Next i
                modGlobals.gWebHost.Visible = False
                gDisplayingYoutube = False
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Selects the current paragraph: call this after SkipLinks or similar
    ''' to give a visual representation of where we're up to.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SelectCurrentParagraph()
        Try
            Dim foundEnd As Integer
            Dim foundStart As Integer
            Dim startFrom As Integer

            'Get the end point
            startFrom = txtText.SelectionStart
            If startFrom = 0 Then startFrom = 1
            foundEnd = InStr(startFrom, txtText.Text, vbNewLine, CompareMethod.Binary)
            If foundEnd = 0 Then
                foundEnd = Len(txtText.Text)
            End If
            foundStart = InStrRev(txtText.Text, vbNewLine, startFrom, CompareMethod.Binary)
            If foundStart > 0 Then
                foundStart = foundStart + 1 ' Dev: otherwise we select the newline.
                'You'd think it would be +len(vbnewline), but that goes too far!
            End If
            txtText.SelectionStart = foundStart
            txtText.SelectionLength = foundEnd - foundStart
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' corrects the tabstops on the main form. Call during Form_Load
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetupTabs()
        Try
            'correct the tabs
            staMain.TabIndex = 0
            cboAddress.TabIndex = 1
            txtText.TabIndex = 2 '       alternate tabstop property
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub GetPDFFromGoogle(ByRef url As String)

    '    cboAddress.Text = "http://www.google.com/search?q=cache:" & Replace(url, "http://", "")
    '    'UPGRADE_WARNING: Navigate2 was upgraded to Navigate and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '    Call modGlobals.gWebHost.webMain.Navigate(New System.Uri(cboAddress.Text))
    'End Sub

    Public Sub AddURLToRecent(Optional url As String = "")
        Try
            Static recent As String
            If url = "" Then url = cboAddress.Text
            If Len(url) > 0 Then
                url = Replace(url, "http://", "", , , CompareMethod.Text)
                url = Replace(url, "https://", "", , , CompareMethod.Text)
                If InStr(1, url, "?") > 0 Then
                    url = Microsoft.VisualBasic.Left(url, InStr(1, url, "?", CompareMethod.Binary) - 1)
                End If
                If InStr(1, recent, url, CompareMethod.Text) > 0 Then
                    'already in list
                Else
                    'add to list
                    Try
                        Call cboAddress.Items.Add(url)
                        recent = recent & "*" & url
                    Catch ex As System.ArgumentNullException
                        ' End up here if url is null: don't add to combobox.
                    End Try
                End If
            End If
        Catch
        End Try
    End Sub

    Private Function ContainsInputElement(ByRef e As mshtml.IHTMLElement) As Boolean
        Try
            If FindInputElement(e) Is Nothing Then
                Return False
            Else
                Return True
            End If
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns the element that is a child of e (or e itself) that is an input/form element, 
    ''' or Nothing if not found.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindInputElement(ByRef e As mshtml.IHTMLElement) As mshtml.IHTMLElement
        Try
            Dim childE As mshtml.IHTMLElement

            FindInputElement = Nothing
            Select Case UCase(e.tagName)
                Case "INPUT"
                    FindInputElement = e
                Case "DIR"
                    FindInputElement = e
                Case "BUTTON"
                    FindInputElement = e
                Case "ISINDEX"
                    FindInputElement = e
                Case "MENU"
                    FindInputElement = e
                Case "SELECT"
                    FindInputElement = e
                Case "TEXTAREA"
                    FindInputElement = e
                Case Else
                    Dim ci As mshtml.IHTMLElementCollection
                    ci = CType(e.children, mshtml.IHTMLElementCollection)
                    For Each childE In ci
                        FindInputElement = FindInputElement(childE)
                        If Not (FindInputElement Is Nothing) Then Exit For
                    Next childE
            End Select
        Catch
            Return Nothing
        End Try
    End Function

#If DEBUGGING Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUGGING did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Public Sub DummyFunction()
	Debug.Print DummyFunctionPutInToStopYouCompilingAndReleasingWithDebuggingTurnedOn
	End Sub
#End If

    Private Function GetCurrentElementIndex() As Integer
        Try
            Dim currentLine As String = GetCurrentLine(txtText)
            If String.IsNullOrWhiteSpace(currentLine) Then Return -1

            Dim startIndex As Integer = currentLine.IndexOf("[")
            Dim endIndex As Integer = currentLine.IndexOf("]")

            If startIndex = -1 Or endIndex = -1 Or endIndex <= startIndex + 1 Then
                Return -1
            End If

            Dim indexString As String = currentLine.Substring(startIndex + 1, endIndex - startIndex - 1)
            Dim index As Integer
            If Integer.TryParse(indexString, index) Then
                Return index
            Else
                Return -1
            End If
        Catch
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Allow users to follow a particular link by number
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtText_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtText.KeyDown
        Try
            Dim KeyCode As Integer = e.KeyCode

            'Detect CAPSLOCK is down, and if so don't do any key handling, since it's the standard screenreader control key.
            If modKeys.IsCapslockPressed Then
                Exit Sub
            End If

            Dim handled As Boolean = False
            Static busy As Boolean

            If busy Then
                Exit Sub
            End If
            busy = True
            mControlKeyPressed = e.Control
            mShiftKeyPressed = e.Shift

            If KeyCode = System.Windows.Forms.Keys.Tab And e.Control Then
                'Don't be tempted to put any SendKeys in here to encourage screenreaders to read the new line.
                'This will be keypress-handled and since Ctrl is down you'll get SkipLinks happening.
                If e.Shift Then
                    Call SkipLinks(SKIP_UP, True)
                Else
                    Call SkipLinks(SKIP_DOWN, True)
                End If
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D1 And e.Control Then
                'Jump to next/previous Header 1
                Call GotoLineStarting("Heading 1:", CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D2 And e.Control Then
                'Jump to next/previous Header 2
                Call GotoLineStarting("Heading 2:", CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D3 And e.Control Then
                'Jump to next/previous Header 3
                Call GotoLineStarting("Heading 3:", CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D4 And e.Control Then
                'Jump to next/previous Header 4
                Call GotoLineStarting("Heading 4:", CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D5 And e.Control Then
                'Jump to next/previous Header 5
                Call GotoLineStarting("Heading 5:", CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D6 And e.Control Then
                'Jump to next/previous Header 6
                Call GotoLineStarting("Heading 6:", CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.Return Then
                Call UserPressedReturn()
                handled = True
            End If
            If handled Then
                e.Handled = True
                e.SuppressKeyPress = True
            End If
            busy = False
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' This handles user interaction with the txtText area - in other words, when return is pressed.
    ''' </summary>
    ''' <remarks></remarks>
    Private Async Sub UserPressedReturn()
        Try
            Dim index As Integer = GetCurrentElementIndex()
            If index = -1 OrElse index >= gInteractableElements.Count Then
                ' Not on a recognized element, or index is out of bounds
                Call PlayErrorSound()
                Return
            End If

            Dim element As InteractableElement = gInteractableElements(index)

            ' If shift is pressed, we might want to do a special action, e.g., download
            If mShiftKeyPressed AndAlso element.type = "link" Then
                ' Handle download logic (new helper needed)
                ' Call DownloadLink(element.href)
                Return
            End If

            ' The core action is to execute a script to interact with the element.
            Dim script As String = $"document.querySelector('[data-webbie-id=""{element.id}""]').click();"

            ' For text/password/textarea, we should pop an input box first.
            Select Case element.type
                Case "text", "search", "email", "url", "tel"
                    Dim promptText As String = If(String.IsNullOrWhiteSpace(element.text), "Enter text", element.text)
                    Dim currentValue As String = Await GetElementValueAsync(element.id)
                    Dim newValue As String = InputBox(promptText, "Text Input", currentValue)
                    If newValue <> currentValue Then ' Check if user cancelled or entered same value
                        Await SetElementValueAsync(element.id, newValue)
                        Call DoDelayedRefresh()
                    End If

                Case "password"
                    Dim promptText As String = If(String.IsNullOrWhiteSpace(element.text), "Enter password", element.text)
                    Dim pw As New frmPassword
                    pw.SetLabel(promptText)
                    If pw.ShowDialog(Me) = DialogResult.OK Then
                        Await SetElementValueAsync(element.id, pw.GetPassword())
                        Call DoDelayedRefresh()
                    End If
                    pw.Close()

                Case "textarea"
                    ' Show the textarea input form
                    frmTextareaInput.gAreaLabel = element.text
                    ' We need a way to pass the element id to frmTextareaInput so it can update the value
                    ' Let's assume we add a public property to it.
                    frmTextareaInput.TargetElementId = element.id
                    Dim currentValue As String = Await GetElementValueAsync(element.id)
                    frmTextareaInput.PopulateWithText(currentValue) ' New method needed
                    Call frmTextareaInput.ShowDialog(Me)
                    ' The form will handle the update and refresh.

                Case "select-one"
                    ' Show the select form
                    Call frmSelect.Populate(element.id)
                    Call frmSelect.ShowDialog(Me)
                    ' The frmSelect form now handles the refresh logic internally.


                Case Else
                    ' Default action is to click the element, which includes links, buttons, checkboxes, radios
                    gForceNavigation = True ' To let the webview events know a navigation is expected
                    Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)
                    Call DoDelayedRefresh()
            End Select

        Catch ex As Exception
            Debug.Print("Error in UserPressedReturn: " & ex.Message)
            Call PlayErrorSound()
        End Try
    End Sub

    Public Async Function GetElementValueAsync(elementId As String) As Task(Of String)
        Dim script = $"document.querySelector('[data-webbie-id=""{elementId}""]').value;"
        Dim jsonResult = Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)
        If String.IsNullOrWhiteSpace(jsonResult) OrElse jsonResult = "null" Then Return ""
        Return JsonConvert.DeserializeObject(Of String)(jsonResult)
    End Function

    Public Async Function SetElementValueAsync(elementId As String, value As String) As Task
        ' Properly escape the value for JavaScript
        Dim escapedValue = JsonConvert.SerializeObject(value)
        Dim script = $"document.querySelector('[data-webbie-id=""{elementId}""]').value = {escapedValue};"
        Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)
    End Function

    ''' <summary>
    ''' Call this to turn on a delayed refresh of the display: that is, it'll use a time to wait a bit then
    ''' do a screen refresh UNLESS a page navigation has commenced in the meantime. 
    ''' </summary>
    Public Sub DoDelayedRefresh()
        tmrRefreshIfNotChange.Enabled = True
        gNoPageActionSoRefresh = True
    End Sub


    ''' <summary>
    ''' Takes a proposed email address and does some simple checks that it is valid, returning true/false
    ''' and an informative error message. Does NOT comply with the appropriate RFC. 
    ''' </summary>
    ''' <param name="emailAddress"></param>
    ''' <param name="failureMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValidEmail(emailAddress As String, ByRef failureMessage As String) As Boolean
        Try
            If emailAddress.Length = 0 Then
                failureMessage = ""
                Return False
            ElseIf emailAddress.IndexOf("@") = -1 Then
                failureMessage = modI18N.GetText("Invalid email address: you must have an @ somewhere in the address!")
                Return False
            Else
                Dim parts() As String
                Dim atSymbol As Char = "@'".ToCharArray()(0)
                parts = emailAddress.Split(atSymbol)
                If parts.Length > 2 Then
                    failureMessage = modI18N.GetText("Invalid email address: you must have only one @ in the address!")
                    Return False
                ElseIf parts(1).Length = 0 Then
                    failureMessage = modI18N.GetText("Invalid email address: there must be something after the @ symbol!")
                    Return False
                ElseIf Not parts(1).Contains(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the part after the @ symbol must have a full stop in it!")
                    Return False
                ElseIf parts(1).EndsWith(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the part after the @ symbol must not end with a full stop!")
                    Return False
                ElseIf parts(1).StartsWith(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the part after the @ symbol must not start with a full stop!")
                    Return False
                ElseIf parts(0).Length = 0 Then
                    failureMessage = modI18N.GetText("Invalid email address: there must be something before the @ symbol!")
                    Return False
                ElseIf parts(0).StartsWith(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the email address must not start with a full stop!")
                    Return False
                Else
                    ' I guess that's okay!
                    failureMessage = ""
                    Return True
                End If
            End If
        Catch
            Return True
        End Try
    End Function

    Private Sub txtText_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtText.KeyUp
    End Sub

    Private Function SanitiseLinkForDisplay(ByVal linkText As String) As String
        Try
            Dim outputText As String
            'TODO Get more stuff to use this!
            If linkText.Length > 0 Then
                outputText = Replace(linkText, "_", " ")
                outputText = Replace(outputText, "-", " ")

                'I think I'm doing this to create something useful when you get links
                'which can only be labelled with the url, like:
                'amazon.com/ref.php?kjkljklj=sdfsdf&sdljkjk=fkljoji...
                'But I can't see any benefit of this. So comment it out for now!
                'If outputText.Contains(".") Then
                'If outputText.LastIndexOf(".") > outputText.Length - 6 Then
                'outputText = outputText.Substring(0, outputText.LastIndexOf("."))
                'End If
                'End If
                Return outputText
            Else
                Return ""
            End If
        Catch
            Return linkText
        End Try
    End Function

#If MONITOR_FOCUS Then
    Private Sub tmrFocusMonitor_Tick(sender As System.Object, e As System.EventArgs) Handles tmrFocusMonitor.Tick
        Static busy As Boolean

        Debug.WriteLine("tmrFocusMonitor running")
        If busy Then
        Else
            busy = True
            If frmGoogle.Visible Or Not Me.Visible Then
                Debug.WriteLine("tmrFocusMonitor inactive")
            Else
                'If mIEVisible Then
                '    'Fine, web browser probably has the focus.
                'Else
                Debug.WriteLine("tmrFocusMonitor valid to act")
                'So, if the mControlName is "", then MSAA has broken, probably because the modGlobals.gWebHost.webMain
                'control has seized focus. We can fix this now:
#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
                If mControlName = "" And mControlType = "ControlType.Document" Then
                    Debug.WriteLine("tmrFocusMonitor ACTING")
                    Call Me.Focus()
                    Call System.Windows.Forms.Application.DoEvents()
                    Call txtText.Focus()
                End If
#End If
                'If Me.ActiveControl Is Nothing Then
                '    'Whoops! Nothing has the focus. Head to txtText.
                '    Call Me.Focus()
                '    Call System.Windows.Forms.Application.DoEvents()
                '    Call txtText.Focus()
                'ElseIf Me.ActiveControl.Name = "cboAddress" Then
                '    'OK, we're fine with being in cboAddress.
                '    'Call Me.Focus()
                '    'Call System.Windows.Forms.Application.DoEvents()
                '    'Call cboAddress.Focus()
                'Else
                '    'Otherwise we should be in txtText.
                '    Call Me.Focus()
                '    Call System.Windows.Forms.Application.DoEvents()
                '    Call txtText.Focus()
                'End If
                'End If
            End If
            busy = False
        End If
    End Sub
#End If

#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
    Private Sub OnFocusChanged(src As Object, e As System.Windows.Automation.AutomationFocusChangedEventArgs)
        On Error Resume Next
        'Cast the object to an AutomationElement.
        Dim elementFocused As System.Windows.Automation.AutomationElement
        elementFocused = CType(src, System.Windows.Automation.AutomationElement)
        mControlName = ""
        If elementFocused Is Nothing Then
            mControlName = "Unknown"
        ElseIf elementFocused.Current Is Nothing Then
            mControlName = "Unknown"
        Else
            If elementFocused.Current Is Nothing Then
                mControlName = "Unknown"
            Else
                mControlName = elementFocused.Current.ClassName
                mControlType = elementFocused.Current.ControlType.ProgrammaticName
                Debug.WriteLine("mControlType:" & mControlType)
            End If
        End If
    End Sub
#End If
    Private Sub mnuFavoritesGotofavorites_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFavoritesGotofavorites.Click
        On Error Resume Next
        Call GotoWebbIEHomePage()
    End Sub

    Private Sub GotoWebbIEHomePage()
        Try
            Dim path As String = System.IO.Path.GetTempFileName
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(path)
            Call sw.Write(modFavorites.GenerateHomepageHTML)
            Call sw.Close()
            Call StartNavigating(path)
        Catch
        End Try
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
        Try
            'open an html file on disk or internet if a url

            Dim Path As String

            cdlgOpen.Filter = modI18N.GetText("HTML file") & " (*.htm, *.html, *.mht)|*.htm;*.xhtm;*.html;*.xhtml;*.mht|" & modI18N.GetText("HTML Help files") & " (*.chm)|*.chm|" & modI18N.GetText("PDF Files") & " (*.pdf)|*.pdf|" & modI18N.GetText("All files") & " (*.*)|*.*;"
            cdlgOpen.ShowReadOnly = False
            cdlgOpen.CheckFileExists = False
            cdlgOpen.CheckPathExists = False
            Call cdlgOpen.ShowDialog()
            'if we have a valid filename, try to load it...
            Dim cmdline As String
            Dim outputFile As String
            If cdlgOpen.FileName <> "" Then
                'okay, is this a valid file?
                If System.IO.File.Exists(cdlgOpen.FileName) Then
                    'Compiled HTML help file?
                    If StrComp(Microsoft.VisualBasic.Right(cdlgOpen.FileName, 3), "chm", CompareMethod.Text) = 0 Then
                        'compiled html help
                        frmParseHTMLHelp = New frmParseHTMLHelp
                        Path = frmParseHTMLHelp.ConvertHTMLHelp((cdlgOpen.FileName))
                        If Len(Path) > 0 Then
                            Call StartNavigating(Path)
                        End If
                    ElseIf LCase(Microsoft.VisualBasic.Right(cdlgOpen.FileName, 4)) = ".pdf" Then
                        'PDF file!
                        outputFile = Replace(cdlgOpen.FileName, ".pdf", "", , , CompareMethod.Text)
                        cmdline = """" & My.Application.Info.DirectoryPath & "\pdftohtml.exe"""
                        Dim argumentLine As String = """" & cdlgOpen.FileName & """ """ & outputFile & """ -c -p -noframes"
                        Dim si As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(cmdline, argumentLine)
                        si.CreateNoWindow = True
                        Dim makeProcess As System.Diagnostics.Process = System.Diagnostics.Process.Start(si)
                        Call makeProcess.WaitForExit()
                        outputFile = outputFile & ".html"
                        Call StartNavigating(outputFile)
                    Else
                        'html
                        Call StartNavigating((cdlgOpen.FileName))
                    End If
                Else
                    'nope: is it a url?
                    Dim newURL As System.Uri = New System.Uri(cdlgOpen.FileName.Replace(".html", "").Replace(".htm", ""))
                    If newURL.IsWellFormedOriginalString And newURL.IsAbsoluteUri Then
                        Call StartNavigating(newURL.ToString)
                    Else
                        Call PlayErrorSound()
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        On Error Resume Next
        'save the current webpage  
        Call SaveAs()
    End Sub

    Private Sub EditCut()
        Try
            'process user menu command Cut

            Dim selStart As Integer ' position of the cursor in the text to be cut from
            'only works in address bar
            If Me.ActiveControl.ToString() = cboAddress.Text And cboAddress.SelectedText <> "" Then
                Call My.Computer.Clipboard.Clear()
                Call My.Computer.Clipboard.SetText(cboAddress.SelectedText)
                selStart = cboAddress.SelectionStart
                cboAddress.Text = Microsoft.VisualBasic.Left(cboAddress.Text, cboAddress.SelectionStart) & Microsoft.VisualBasic.Right(cboAddress.Text, Len(cboAddress.Text) - cboAddress.SelectionStart - cboAddress.SelectionLength)
                cboAddress.SelectionStart = selStart
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuEditCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Call EditCut()
    End Sub

    Private Sub mnuLinksSkiplinksDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLinksSkipDown.Click
        On Error Resume Next
        Call DoSkip(SKIP_DOWN)
    End Sub

    Private Sub DoSkip(SkipDirection As Integer)
        Try
            If gShowFormsOnly Then
                Call SkipToForm(SkipDirection)
            Else
                Call SkipLinks(SkipDirection)
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuNavigateGotoheadline_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNavigateGotoheadline.Click
        On Error Resume Next
        Call GotoHeadline(SKIP_DOWN)
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptionsOptions.Click
        Try
            Dim frmOptions As frmOptionsForm = New frmOptionsForm
            If frmOptions.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                'TODO is this right? Script errors = messages?
                modGlobals.gWebHost.webMain.ScriptErrorsSuppressed = Not My.Settings.AllowMessages
                'If we don't use the IE homepage, disable the ability to change it.
                mnuOptionsSetHomepage.Enabled = My.Settings.UseIEHomepage
                Call ResizeToolbar()
                Call RefreshCurrentPage()
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuFavoritesAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFavoritesAdd.Click
        Try
            'adds a bookmark to the IE favorites list
            Dim name As String = modGlobals.gWebHost.webMain.DocumentTitle
            Dim url As String = modGlobals.gWebHost.URL
            Dim c As Char

            For Each c In System.IO.Path.GetInvalidFileNameChars
                name = name.Replace(c, "")
            Next c

            name = name.Trim

            If name = "" Then name = GetText("Untitled Webpage")
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Favorites) & "\" & name & ".url")
            Call sw.WriteLine("[DEFAULT]")
            Call sw.WriteLine("BASEURL=" & url)
            Call sw.WriteLine("[InternetShortcut]")
            Call sw.WriteLine("URL=" & url)
            Call sw.Close()
            'Add to bookmarks menu.
            Dim tsi As ToolStripItem = mnuFavorites.DropDownItems.Add(name, Nothing, AddressOf FavoriteClickHandler)
            tsi.Tag = url
        Catch
        End Try
    End Sub

    Private Sub mnuFavoritesOrganise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFavoritesOrganise.Click
        Try
            Dim folderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Favorites)
            Call Shell("explorer """ & folderPath & """", AppWinStyle.NormalFocus)
            'Ideally I could do something like track the folder and see if it changes and then reload bookmarks. 
        Catch
        End Try
    End Sub

    Private Sub FavoriteClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim tsi As ToolStripItem = CType(sender, System.Windows.Forms.ToolStripItem)
            Call StartNavigating(CStr(tsi.Tag))
        Catch
        End Try
    End Sub

    Public Sub LoadBookmarks(ByVal currentMenu As ToolStripMenuItem, ByVal dir As System.IO.DirectoryInfo)
        Try
            'Iterate through dir's files, adding to currentMenu.
            Dim f As System.IO.FileInfo
            Dim childDir As System.IO.DirectoryInfo
            Dim url As String
            Dim name As String
            Dim tsi As System.Windows.Forms.ToolStripItem
            Dim tsmi As System.Windows.Forms.ToolStripMenuItem
            'Do subfolders. 
            For Each childDir In dir.EnumerateDirectories
                name = ""
                name = childDir.Name
                tsmi = New System.Windows.Forms.ToolStripMenuItem(name)
                Call currentMenu.DropDownItems.Add(tsmi)
                Call LoadBookmarks(tsmi, childDir)
            Next childDir
            'Do favorites. It seems more standard to put folders first, I think? Though
            'going to your favorites doesn't think that way: it does links first. TODO.
            For Each f In dir.EnumerateFiles()
                name = ""
                name = f.Name
                url = modIniFile.GetString("InternetShortcut", "URL", "", f.FullName)
                If name <> "" And url <> "" Then
                    name = Replace(name, ".url", "", , , CompareMethod.Text)
                    tsi = currentMenu.DropDownItems.Add(name, Nothing, AddressOf FavoriteClickHandler)
                    tsi.Tag = url
                End If
                System.Windows.Forms.Application.DoEvents()
            Next f
        Catch
        End Try
    End Sub

    Private Sub SaveAs()
        On Error Resume Next
        'saves the current webpage as an html file
        Call modGlobals.gWebHost.webMain.ShowSaveAsDialog()
    End Sub


    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        Try
            'respond to user exiting WebbIE

            '    Dim EndCall As VbMsgBoxResult
            '    'if the user is dialled up on program exit
            '    If (ActiveConnection() = True And gblnModem = True) Then
            '        'prompt for disconnection
            '        EndCall = MsgBox("Would you like to disconnect your Internet connection as well as leave WebbIE?", vbYesNo, "Active Internet connection")
            '    End If
            '    'follow request as above
            '    If EndCall = vbYes Then
            '        Call HangUp
            '    End If
            'unload form
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub btnHeading_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHeading.Click
        Try
            'TODO check for pressing Shift, which should do skip up.
            Call GotoHeadline(SKIP_DOWN)
            Call txtText.Focus()
        Catch
        End Try
    End Sub

    Private Sub btnSkiplinks_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSkiplinks.Click
        Try
            Call DoSkip(SKIP_DOWN)
            'If mIEVisible Then
            'Else
            Call txtText.Focus()
            'End If
        Catch
        End Try
    End Sub

    Private Sub mnuOptionsSetHomepage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptionsSetHomepage.Click
        On Error Resume Next
        Call SetHomepage()
    End Sub

    Private Sub btnRSS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRSS.Click
        On Error Resume Next
        Call RSS()
    End Sub

    Private Sub mnuHelpManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpManual.Click
        On Error Resume Next
        Call modI18N.ShowHelp()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        On Error Resume Next
        Call DoWebSearch()
    End Sub

    Public Sub DoGotoAddressbar()
        Try
            cboAddress.Text = String.Empty
            Call cboAddress.Focus()
        Catch
        End Try
    End Sub

    Private Sub btnCrop_Click(sender As System.Object, e As System.EventArgs) Handles btnCrop.Click
        On Error Resume Next
        Call CropPage()
    End Sub

    Private Sub WebpageToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuViewIE.Click
        Try
            modGlobals.gShowingWebpage = True
            Call Me.Hide()
            Call gWebHost.Show()
        Catch
        End Try
    End Sub

    Private Sub ShowFavoritesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuFavoritesShowfavorites.Click
        On Error Resume Next
        Call ShowFavoritesWindow()
    End Sub

    Private Sub ShowFavoritesWindow()
        On Error Resume Next
        Call frmFavorites.ShowDialog(Me)
    End Sub

    Private Sub mnuFavorites_Click(sender As System.Object, e As System.EventArgs) Handles mnuFavorites.Click

    End Sub

    Private Sub frmMain_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        On Error Resume Next
        modGlobals.gClosing = True
        mFormClosing = True
    End Sub

    ''' <summary>
    ''' Starts off the "load the bookmarks and populate the bookmarks menu" process, which is really slow
    ''' because of all the menu work. So start the timer after loading so the program starts faster and
    ''' THEN loads in all the bookmark menu.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub tmrDelayLoadBookmarks_Tick(sender As System.Object, e As System.EventArgs) Handles tmrDelayLoadBookmarks.Tick
        If modGlobals.gClosing Then
            Me.tmrDelayLoadBookmarks.Enabled = False
            Exit Sub
        End If
        Try
            tmrDelayLoadBookmarks.Enabled = False
            Call LoadBookmarks(mnuFavorites, New System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Favorites)))
        Catch
        End Try
    End Sub

    Private Sub mnuOptionsColour_Click(sender As Object, e As EventArgs) Handles mnuOptionsColour.Click
        Dim fcs As frmColourSelect = New frmColourSelect
        Call fcs.SetCurrentSelection(My.Settings.ColourScheme)
        If fcs.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            ' Update global colours
            My.Settings.ColourScheme = fcs.SelectedColourScheme
            Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
        End If
    End Sub

    ''' <summary>
    ''' Used in mnuHelpTeamviewer_Click to store where TeamViewerQS.exe will be saved to (so when
    ''' it has downloaded it can be run.)
    ''' </summary>
    ''' <remarks></remarks>
    Private mTeamViewerPath As String = ""
    Private WithEvents mWebClient As System.Net.WebClient

    Private Sub mnuHelpTeamviewer_Click(sender As Object, e As EventArgs) Handles mnuHelpTeamviewer.Click
        mWebClient = New System.Net.WebClient()
        mTeamViewerPath = System.IO.Path.GetTempFileName() + ".exe"
        Call mWebClient.DownloadFileAsync(New System.Uri("http://www.webbie.org.uk/download/TeamViewerQS.exe"), mTeamViewerPath)
        '    wc.DownloadProgressChanged += wc_DownloadProgressChangedTV;
        
    End Sub

    Private Sub mWebClient_DownloadFileCompleted(sender As Object, e As System.ComponentModel.AsyncCompletedEventArgs) Handles mWebClient.DownloadFileCompleted
        Try
            staMain.Items.Item(0).Text = modI18N.GetText("Idle")
            System.Diagnostics.Process.Start(mTeamViewerPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub mWebClient_DownloadProgressChanged(sender As Object, e As Net.DownloadProgressChangedEventArgs) Handles mWebClient.DownloadProgressChanged
        Try
            staMain.Items.Item(0).Text = "TeamViewer " & e.ProgressPercentage & "%"
        Catch ex As Exception

        End Try
    End Sub

    Private Class AIResponse
        Public Property tool As String
        Public Property args As List(Of String)
    End Class

    Private Async Sub btnAskAI_Click(sender As Object, e As EventArgs) Handles btnAskAI.Click
        Try
            ' 1. Get user query
            Dim userQuery As String = InputBox("What would you like to do?", "Ask AI")
            If String.IsNullOrWhiteSpace(userQuery) Then Return

            ' 2. Get page content
            Dim pageContent As String = txtText.Text

            ' 3. Call the AI process. The JS side will now directly call back to the ToolHost object.
            ' The returned JSON is now just for logging/debugging.
            Dim script = $"window.aiProcess(`{pageContent.Replace("`", "\`")}`, `{userQuery.Replace("`", "\`")}`);"
            Dim jsonResult As String = Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)

            ' Optional: Log the result for debugging
            Debug.Print("AI Execution Result: " & jsonResult)

        Catch ex As Exception
            MessageBox.Show("An error occurred while communicating with the AI assistant: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#Region "Test Harness"
    Private _testCompletionSource As TaskCompletionSource(Of Boolean)

    Private Async Sub btnRunTests_Click(sender As Object, e As EventArgs) Handles btnRunTests.Click
        Dim testResults As New System.Text.StringBuilder()
        Dim testsPassed As Integer = 0
        Dim testsFailed As Integer = 0

        Try
            Dim testDirectory As String = System.IO.Path.Combine(Application.StartupPath, "Tests", "mock_pages")
            If Not System.IO.Directory.Exists(testDirectory) Then
                MessageBox.Show("Test directory not found: " & testDirectory, "Test Runner Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim testFiles = System.IO.Directory.GetFiles(testDirectory, "*.html")

            For Each testFile As String In testFiles
                _testCompletionSource = New TaskCompletionSource(Of Boolean)()

                Dim fileUri As New Uri(testFile)
                modGlobals.gWebHost.webMain.CoreWebView2.Navigate(fileUri.AbsoluteUri)

                Await _testCompletionSource.Task

                Dim actualJson As String = modGlobals.gLastParserResult
                Dim expectedFile As String = testFile.Replace(".html", ".expected.json")
                Dim expectedJson As String = ""

                If System.IO.File.Exists(expectedFile) Then
                    expectedJson = System.IO.File.ReadAllText(expectedFile)
                End If

                ' Normalize JSON for comparison
                Dim normalizedActual = JsonConvert.SerializeObject(JsonConvert.DeserializeObject(actualJson))
                Dim normalizedExpected = JsonConvert.SerializeObject(JsonConvert.DeserializeObject(expectedJson))

                If normalizedActual = normalizedExpected Then
                    testsPassed += 1
                    testResults.AppendLine($"PASS: {System.IO.Path.GetFileName(testFile)}")
                Else
                    testsFailed += 1
                    testResults.AppendLine($"FAIL: {System.IO.Path.GetFileName(testFile)}")
                    testResults.AppendLine("--- EXPECTED ---")
                    testResults.AppendLine(normalizedExpected)
                    testResults.AppendLine("--- ACTUAL ---")
                    testResults.AppendLine(normalizedActual)
                    testResults.AppendLine("----------------")
                End If
            Next

            Dim summary As String = $"Test run complete. Passed: {testsPassed}, Failed: {testsFailed}" & vbCrLf & vbCrLf & testResults.ToString()
            MessageBox.Show(summary, "Test Results", MessageBoxButtons.OK, If(testsFailed > 0, MessageBoxIcon.Error, MessageBoxIcon.Information))

        Catch ex As Exception
            MessageBox.Show("An error occurred during the test run: " & ex.Message, "Test Runner Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Process the loading HTML object model and display it as text. Call this after refresh, page navigation, or any time
    ''' you need to update this view to reflect a change to the DOM.
    ''' </summary>
    Private Async Sub ParseDocument()
        If modGlobals.gClosing OrElse mFormClosing Then
            Exit Sub
        End If

        Try
            ' Tell user we're working, unless we're in a test run
            If _testCompletionSource Is Nothing OrElse _testCompletionSource.Task.IsCompleted Then
                staMain.Items.Item(0).Text = modI18N.GetText("Examining")
                Call Application.DoEvents()
            End If

            ' Clear old data
            Call ClearPageData()

            ' Read the DOM parser script from resources
            Dim parserScript As String = My.Resources.dom_parser

            ' Execute the script to get interactable elements as JSON
            Dim jsonResult As String = Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(parserScript)
            modGlobals.gLastParserResult = jsonResult ' Store for testing

            ' The result is a JSON string literal. We need to un-escape and parse it.
            If String.IsNullOrWhiteSpace(jsonResult) OrElse jsonResult = "null" OrElse jsonResult.Length < 3 Then
                ' Script failed or returned nothing
                staMain.Items.Item(0).Text = modI18N.GetText("Done") ' Page might be empty
                Call SetText("") ' Clear the text view
                ' If in a test run, signal completion
                _testCompletionSource?.TrySetResult(True)
                Return
            End If

            ' Deserialize the JSON string into our list of elements
            ' ExecuteScriptAsync returns a JSON-encoded *string*. So we must first deserialize the string itself,
            ' then deserialize the content of that string.
            Dim jsonContent As String = JsonConvert.DeserializeObject(Of String)(jsonResult)
            gInteractableElements = JsonConvert.DeserializeObject(Of List(Of InteractableElement))(jsonContent)

            ' Reset flags
            mCropped = False

            ' Build the output for the text view
            mOutput = New System.Text.StringBuilder(32000)

            ' TODO: Check for RSS feed link separately if needed
            ' gRSSFeedURL = Await CheckForRssAsync() ' This would need a new helper

            ' Process the list of elements to build the text display
            Dim i As Integer = 0
            For Each element As InteractableElement In gInteractableElements
                Dim line As String = ""
                Dim prefix As String = ""
                Dim numberString As String = If(My.Settings.NumberLinks, $" [{i}]", "")

                Select Case element.type
                    Case "link"
                        prefix = ID_LINK
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "button", "submit", "reset"
                        prefix = If(element.type = "submit", ID_SUBMIT, If(element.type = "reset", ID_RESET, ID_BUTTON))
                        line = $"{prefix}{numberString}: ({element.text})"
                    Case "text", "search", "email", "url", "tel"
                        prefix = ID_TEXTBOX
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "password"
                        prefix = ID_PASSWORD
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "textarea"
                        prefix = ID_TEXTAREA
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "checkbox"
                        prefix = ID_CHECKBOX
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "radio"
                        prefix = ID_RADIO
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "select-one", "select-multiple"
                        prefix = ID_SELECT
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "file"
                        prefix = ID_FILE
                        line = $"{prefix}{numberString}: {element.text}"
                    Case "h1", "h2", "h3", "h4", "h5", "h6"
                        line = $"Heading {element.type.Substring(1)}: {element.text}"
                    Case "p", "div", "span", "text", "main", "article"
                        ' For simple text blocks, just show the text. We might want to filter short/empty ones.
                        If Not String.IsNullOrWhiteSpace(element.text) Then
                            line = element.text
                        End If
                    ' Add other cases as needed for VIDEO, AUDIO, etc.
                End Select

                If Not String.IsNullOrWhiteSpace(line) Then
                    Call mOutput.AppendLine(line)
                End If
                i += 1
            Next

            ' Display the result
            Call DisplayOutput()

            ' Update UI state, unless in a test run
            If _testCompletionSource Is Nothing OrElse _testCompletionSource.Task.IsCompleted Then
                btnStop.Enabled = False
                btnRefresh.Enabled = True
                staMain.Items.Item(0).Text = modI18N.GetText("Done")
                mnuLinksViewlinks.Enabled = (gInteractableElements.Count > 0)
                cboAddress.SelectionLength = 0
                cboAddress.Refresh()
                Call txtText.Focus()
                Application.DoEvents()
                txtText.SelectionLength = 0
            End If

        Catch ex As Exception
            Debug.Print("Exception in ParseDocument: " & ex.Message)
            staMain.Items.Item(0).Text = modI18N.GetText("Error")
        Finally
            ' If in a test run, signal completion
            _testCompletionSource?.TrySetResult(True)
            ' Only play sounds if not in a test run
            If _testCompletionSource Is Nothing OrElse _testCompletionSource.Task.IsCompleted Then
                Call PlayDoneSound()
                Call StopBusyAnimation()
            End If
        End Try
    End Sub
#End Region
End Class
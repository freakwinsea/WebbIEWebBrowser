Option Strict On
Option Explicit On

Imports Newtonsoft.Json

Public Class SelectData
    Public Property options As List(Of OptionData)
    Public Property selectedIndex As Integer
End Class

Public Class OptionData
    Public Property text As String
    Public Property value As String
End Class

Friend Class frmSelect
	Inherits System.Windows.Forms.Form
	'   This file is part of WebbIE.
	'
	'    WebbIE is free software: you can redistribute it and/or modify
	'    it under the terms of the GNU General Public License as published by
	'    the Free Software Foundation, either version 3 of the License, or
	'    (at your option) any later version.
	'
	'    WebbIE is distributed in the hope that it will be useful,
	'    but WITHOUT ANY WARRANTY; without even the implied warranty of
	'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	'    GNU General Public License for more details.
	'
	'    You should have received a copy of the GNU General Public License
	'    along with WebbIE.  If not, see <http://www.gnu.org/licenses/>.
	
	'This module puts up the selection menu for webforms that appears when the user clicks on a drop-down
	'box in a webpage.   Alasdair 20 September 2002
	
	
    Private _elementId As String
    Public gTargetForm As frmMain

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error Resume Next
        Call Me.Hide()
    End Sub

    Private Async Sub cmdOKfrmSelect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOKfrmSelect.Click
        Try
            Dim selectedIndex As Integer = lstSelect.SelectedIndex
            If selectedIndex = -1 Then Return

            Dim script As String = $"document.querySelector('[data-webbie-id=""{_elementId}""]').selectedIndex = {selectedIndex};"
            Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)

            ' Fire the 'change' event to trigger any JavaScript listeners
            Dim eventScript As String = $"var event = new Event('change', {{ 'bubbles': true, 'cancelable': true }}); document.querySelector('[data-webbie-id=""{_elementId}""]').dispatchEvent(event);"
            Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(eventScript)

            Call frmMain.DoDelayedRefresh()
            Call Me.Hide()
        Catch ex As Exception
            Debug.Print("Error in cmdOKfrmSelect_Click: " & ex.Message)
        End Try
    End Sub

    Private Sub frmSelect_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error Resume Next
        lstSelect.Font = frmMain.txtText.Font
    End Sub

    Private Sub frmSelect_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        On Error Resume Next
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000

        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Hide()
    End Sub

    Public Async Sub Populate(ByVal elementId As String)
        Try
            _elementId = elementId

            ' Clear the listbox
            lstSelect.Items.Clear()

            ' Create the script to get options from the select element
            Dim script As String = $"(function() {{
                const select = document.querySelector('[data-webbie-id=""{elementId}""]');
                if (!select) return null;
                const options = Array.from(select.options).map(opt => ({{ text: opt.text, value: opt.value }}));
                return JSON.stringify({{ options: options, selectedIndex: select.selectedIndex }});
            }})();"

            ' Execute the script and get the JSON result
            Dim jsonResult As String = Await modGlobals.gWebHost.webMain.CoreWebView2.ExecuteScriptAsync(script)
            If String.IsNullOrWhiteSpace(jsonResult) OrElse jsonResult = "null" Then Return

            ' Deserialize the JSON string, and then the data within it
            Dim jsonContent As String = JsonConvert.DeserializeObject(Of String)(jsonResult)
            Dim data As SelectData = JsonConvert.DeserializeObject(Of SelectData)(jsonContent)

            If data IsNot Nothing AndAlso data.options IsNot Nothing Then
                ' Populate the list
                For Each opt As OptionData In data.options
                    lstSelect.Items.Add(opt.text)
                Next
                ' Set the selected item
                If data.selectedIndex >= 0 AndAlso data.selectedIndex < lstSelect.Items.Count Then
                    lstSelect.SelectedIndex = data.selectedIndex
                End If
            End If

        Catch ex As Exception
            Debug.Print("Error in frmSelect.Populate: " & ex.Message)
        End Try
    End Sub
	
	Private Sub lstSelect_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lstSelect.KeyDown
        Try
            Dim KeyCode As Integer = eventArgs.KeyCode
            Dim Shift As Integer = eventArgs.KeyData \ &H10000
            'respond to user actions


            If KeyCode = System.Windows.Forms.Keys.Escape Then
                'exit form
                Call Me.Hide()
                If Me.gTargetForm IsNot Nothing Then
                    Call Me.gTargetForm.Show()
                End If
            End If
            If KeyCode = System.Windows.Forms.Keys.Return Or KeyCode = System.Windows.Forms.Keys.Space Then
                'user has chosen an item
                If lstSelect.SelectedIndex < 0 Then
                    'don't do anything - need user to select an item
                Else
                    'user has selected an item
                    Call cmdOKfrmSelect_Click(cmdOKfrmSelect, New System.EventArgs())
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub frmSelect_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub
End Class
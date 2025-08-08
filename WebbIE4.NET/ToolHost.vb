Imports System.Runtime.InteropServices
Imports Microsoft.Web.WebView2.Core

<ComVisible(True)>
Public Class ToolHost
    Private _frmMain As frmMain

    Public Sub New(mainForm As frmMain)
        _frmMain = mainForm
    End Sub

    Public Sub Click(elementId As String)
        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String)(AddressOf _frmMain.ExecuteClick), elementId)
        Else
            _frmMain.ExecuteClick(elementId)
        End If
    End Sub

    Public Sub Type(elementId As String, text As String)
        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String, String)(AddressOf _frmMain.ExecuteType), elementId, text)
        Else
            _frmMain.ExecuteType(elementId, text)
        End If
    End Sub

    Public Sub GoTo(url As String)
        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String)(AddressOf _frmMain.StartNavigating), url)
        Else
            _frmMain.StartNavigating(url)
        End If
    End Sub

    Public Sub Answer(text As String)
        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String)(Sub(t) MessageBox.Show(t, "AI Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information)), text)
        Else
            MessageBox.Show(text, "AI Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class

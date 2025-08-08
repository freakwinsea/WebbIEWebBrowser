Imports System.Runtime.InteropServices
Imports Microsoft.Web.WebView2.Core

<ComVisible(True)>
Public Class ToolHost
    Private _frmMain As frmMain

    ' Test mode properties
    Public IsInTestMode As Boolean = False
    Public LastToolCalled As String = ""
    Public LastArgsCalled As List(Of String) = Nothing

    Public Sub New(mainForm As frmMain)
        _frmMain = mainForm
    End Sub

    Public Sub Click(elementId As String)
        If IsInTestMode Then
            LastToolCalled = "click"
            LastArgsCalled = New List(Of String) From {elementId}
            Return
        End If

        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String)(AddressOf _frmMain.ExecuteClick), elementId)
        Else
            _frmMain.ExecuteClick(elementId)
        End If
    End Sub

    Public Sub Type(elementId As String, text As String)
        If IsInTestMode Then
            LastToolCalled = "type"
            LastArgsCalled = New List(Of String) From {elementId, text}
            Return
        End If

        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String, String)(AddressOf _frmMain.ExecuteType), elementId, text)
        Else
            _frmMain.ExecuteType(elementId, text)
        End If
    End Sub

    Public Sub GoTo(url As String)
        If IsInTestMode Then
            LastToolCalled = "goto"
            LastArgsCalled = New List(Of String) From {url}
            Return
        End If

        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String)(AddressOf _frmMain.StartNavigating), url)
        Else
            _frmMain.StartNavigating(url)
        End If
    End Sub

    Public Sub Answer(text As String)
        If IsInTestMode Then
            LastToolCalled = "answer"
            LastArgsCalled = New List(Of String) From {text}
            Return
        End If

        If _frmMain.InvokeRequired Then
            _frmMain.Invoke(New Action(Of String)(Sub(t) MessageBox.Show(t, "AI Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information)), text)
        Else
            MessageBox.Show(text, "AI Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class

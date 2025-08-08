Option Strict On
Option Explicit On
Friend Class frmLinks
    Inherits System.Windows.Forms.Form

    Private _links As List(Of InteractableElement)

    Private Sub cmdCloseLinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error Resume Next
        cmdGo.Enabled = False
        Call Me.Hide()
    End Sub

    Private Sub cmdGo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGo.Click
        On Error Resume Next
        Call DoGo()
    End Sub

    Private Sub DoGo()
        On Error Resume Next
        Dim selection As Integer = lstLinks.SelectedIndex
        If selection = -1 Then
            'nothing selected - don't do anything
        Else
            Dim selectedLink As InteractableElement = _links(selection)
            If selectedLink IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selectedLink.href) Then
                Call frmMain.StartNavigating(selectedLink.href)
                Call Me.Hide()
            End If
        End If
    End Sub


    Private Sub frmLinks_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error Resume Next
        lstLinks.Font = frmMain.txtText.Font
        If lstLinks.SelectedIndex > -1 Then
            cmdGo.Enabled = True
        End If
        Call lstLinks.Focus()
    End Sub


    Private Sub lstLinks_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstLinks.SelectedIndexChanged
        On Error Resume Next
        cmdGo.Enabled = (lstLinks.SelectedIndex > -1)
    End Sub

    Private Sub lstLinks_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstLinks.DoubleClick
        Call DoGo()
    End Sub

    ''' <summary>
    ''' Populates the list of links from a list of InteractableElement objects.
    ''' </summary>
    ''' <param name="links"></param>
    ''' <remarks></remarks>
    Public Sub PopulateList(ByVal links As List(Of InteractableElement))
        Try
            _links = links
            Call lstLinks.Items.Clear()

            If _links Is Nothing Then Return

            For Each link As InteractableElement In _links
                Call lstLinks.Items.Add(link.text)
            Next

            If (lstLinks.Items.Count > 0) Then
                lstLinks.SelectedIndex = 0
            End If
            Call lstLinks.Focus()
        Catch ex As Exception
            Debug.Print("Error in frmLinks.PopulateList: " & ex.Message)
        End Try
    End Sub

    Private Sub frmLinks_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
        Call lstLinks.Focus()
    End Sub

End Class
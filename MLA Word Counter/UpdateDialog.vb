Imports System.Windows.Forms

Public Class UpdateDialog

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.No
        Me.Close()
    End Sub
    Dim _lv As String
    Dim _chngelog As String
    Overloads Function ShowDialog(LatestVersion As String, Changlog As String) As DialogResult
        _lv = LatestVersion
        _chngelog = Changlog
        Return ShowDialog()
    End Function
    Private Sub UpdateDialog_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = "A new update for the counting engine is available." & vbCrLf & "Version: " & _lv & vbCrLf & "Changelog: " & _chngelog
    End Sub
End Class

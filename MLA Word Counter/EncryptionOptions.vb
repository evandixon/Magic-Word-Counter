Imports System.Windows.Forms
Imports CryptoLibrary
Public Class EncryptionOptions

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub EncryptionOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboAlgorithm.Items.AddRange(System.Enum.GetNames(GetType(SymmetricEncryption.EncryptionAlgorithm)))
        cboAlgorithm.SelectedIndex = 0
    End Sub
    Public Property Password As String
        Get
            Return txtPassword.Text
        End Get
        Set(ByVal value As String)
            txtPassword.Text = value
        End Set
    End Property
    Public Property PaddingCharacter As String
        Get
            Return txtPadChar.Text
        End Get
        Set(ByVal value As String)
            txtPadChar.Text = value
        End Set
    End Property
    Public Property Salt As String
        Get
            Return txtSalt.Text
        End Get
        Set(ByVal value As String)
            txtSalt.Text = value
        End Set
    End Property
    Public Property Algorythm As SymmetricEncryption.EncryptionAlgorithm
        Get
            Return System.Enum.Parse(GetType(SymmetricEncryption.EncryptionAlgorithm), cboAlgorithm.SelectedItem.ToString)
        End Get
        Set(ByVal value As SymmetricEncryption.EncryptionAlgorithm)
            cboAlgorithm.SelectedText = value.ToString
        End Set
    End Property
End Class

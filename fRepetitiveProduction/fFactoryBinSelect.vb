Public Class fFactoryBinSelect
    Private frm As frmRepetitive
    Private bin As String
    Public Event SetFactoryBin(sender As Object, e As EventArgs)

    Public Sub New()

        InitializeComponent()
        frm = frmRepetitive

    End Sub

    Private Sub btnDone_Click(sender As System.Object, e As System.EventArgs) Handles btnDone.Click
        SendFactoryBinToForm()
    End Sub
    Private Sub SendFactoryBinToForm()
        frm.cboDefaultFactoryBin.Text = cboDefaultFactoryBin.Text
        RaiseEvent SetFactoryBin(Me, Nothing)
        Me.Close()
    End Sub

    Private Sub cboDefaultFactoryBin_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles cboDefaultFactoryBin.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendFactoryBinToForm()
        End If
    End Sub

End Class
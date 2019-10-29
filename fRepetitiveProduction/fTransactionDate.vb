Imports System.Text.RegularExpressions

Public Class fTransactionDate
    Private frm As frmRepetitive
    Private transactionDate As Date
    Private currentPeriod As String
    Private currPeriod As Integer
    Private currYear As Integer
    Private tranDateMonth As Integer
    Private tranDateYear As Integer

    Public Event SetTransactionDate(sender As Object, e As EventArgs)

    Public Sub New()

        InitializeComponent()
        frm = frmRepetitive

    End Sub

    Private Sub fTransactionDate_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        currPeriod = frm.systemperiod.current_prd
        currYear = frm.systemperiod.current_year
        lblCurrentPeriod.Text = CurrentPeriod
        TransactionDateDateTimePicker.Focus()
    End Sub

    Private Sub btnDone_Click(sender As System.Object, e As System.EventArgs) Handles btnDone.Click
        SendDateToForm()
    End Sub
    Private Sub SendDateToForm()
        transactionDate = TransactionDateDateTimePicker.Value

        If Not cBusObj.ValidateTransactionDate(transactionDate, currPeriod, currYear) Then
            MsgBox("The month and year must be in the period " & currPeriod.ToString & "/" & currYear.ToString & "." & vbCrLf & vbCrLf & _
                   "Enter a Transaction Effective Date that is within the period " & currPeriod.ToString & "/" & currYear.ToString & ".", vbOKOnly, "Set Period Date")
            Exit Sub
        End If

        frm.TransactionEffectiveDate = transactionDate.ToShortDateString
        RaiseEvent SetTransactionDate(Me, Nothing)
        Me.Close()
    End Sub

    Private Function ValidateTransactionDate(trandate As Date)
       
        If currYear = trandate.Year AndAlso trandate.Month = currPeriod Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub TransactionDateDateTimePicker_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TransactionDateDateTimePicker.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendDateToForm()
        End If
    End Sub

End Class
Imports System.Windows.Forms

Public Class frmZonas
    Private Sub frmZonas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Administrar Zonas"
    End Sub

    Private Sub frmZonas_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        FrmZ.Dispose()
        FrmZ = Nothing
    End Sub
End Class
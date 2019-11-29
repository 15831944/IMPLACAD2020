Public Class frmDescargar
    Public prefijo As String = "Descargando "
    Private Sub pb1_Invalidated(sender As Object, e As EventArgs) Handles pb1.Invalidated
        Me.lblInfo.Text = prefijo & pb1.Value & " de " & pb1.Maximum
    End Sub

    Private Sub frmDescargar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "Actualizando... " & My.Application.Info.ProductName & " - v." & My.Application.Info.Version.ToString
    End Sub
End Class